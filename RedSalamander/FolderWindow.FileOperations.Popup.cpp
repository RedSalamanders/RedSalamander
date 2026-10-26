#include "FolderWindow.FileOperations.Popup.h"
#include "WindowActivation.h"

#include "DxUiThemePalette.h"
#include "FileOperationsFooterLayout.h"
#include "FileSystemPathIdentity.h"
#include "FluentIcons.h"
#include "FolderWindow.FileOperations.PromptText.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "HostServices.h"
#include "LocalizationManager.h"
#include "NavigationLocation.h"
#include "StringConversion.h"
#include "ThroughputParsing.h"
#include "UnicodeClipboard.h"
#include "Win32CallbackHelpers.h"
#include "WindowMaximizeBehavior.h"
#include "WindowSizing.h"
#ifdef ENABLE_TESTS
#include "Common.Tests.Runner.h"
#endif
#include <DxUi/DxUi.h>
#include <DxUi/Typography.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <d2d1.h>
#include <dwrite.h>
#include <initializer_list>
#include <limits>
#include <shobjidl.h>
#include <unordered_map>
#include <unordered_set>
#include <windowsx.h>

namespace
{
// File-operation row semantics: preserve the exact provider path separately from
// the localized label and abbreviated prepared text. Shared DxUi owns focus,
// tooltip layout and chrome; this adapter owns which immutable value is copied.
// The row is a focusable action that opens the path detail, so it is a Button to UI
// Automation (popup-a#2); its own paint, pointer and keyboard routes replace the
// Button's chrome, press and hover behavior.
class FileOperationsPathRegion final : public DxUi::Button
{
public:
    FileOperationsPathRegion()
    {
        SetFocusable(true);
    }

    FileOperationsPathRegion(const FileOperationsPathRegion&)            = delete;
    FileOperationsPathRegion(FileOperationsPathRegion&&)                 = delete;
    FileOperationsPathRegion& operator=(const FileOperationsPathRegion&) = delete;
    FileOperationsPathRegion& operator=(FileOperationsPathRegion&&)      = delete;

    void SetOnOpenDetail(std::function<void()> openDetail)
    {
        _openDetail = std::move(openDetail);
    }

    void SetPath(DxUi::WindowHost& host, std::wstring_view path)
    {
        const auto bounds  = GetBounds();
        const auto tooltip = GetTooltipText();
        // Other active streams can change the help while the first path and its
        // geometry stay fixed. A focused tooltip must follow that content too.
        if (_path == path && _tooltipText == tooltip && bounds.left == _tooltipBounds.left && bounds.top == _tooltipBounds.top &&
            bounds.right == _tooltipBounds.right && bounds.bottom == _tooltipBounds.bottom)
            return;
        _path          = path;
        _tooltipText   = tooltip;
        _tooltipBounds = bounds;
        if (HasFocus())
            ShowPath(host);
    }

    void Paint(DxUi::WindowHost& host) const override
    {
        if (! HasFocus())
            return;
        DxUi::ButtonChromeDrawSpec spec{};
        spec.bounds = GetBounds();
        DxUi::ButtonChromeCustomStyle style{};
        style.showFill   = false;
        style.showFocus  = true;
        style.focus      = host.GetTheme().focusStrokeOuter;
        style.focusRing  = DxUi::ButtonChromeFocusRing::Standard;
        spec.customStyle = style;
        DxUi::DrawButtonChrome(host.GetDeviceContext(), host.GetSolidBrush(style.focus), nullptr, nullptr, host.GetTheme(), spec);
    }

    bool OnCopy(DxUi::WindowHost& host) override
    {
        return Common::Clipboard::TrySetUnicodeText(host.GetHwnd(), _path, Common::Clipboard::EmptyUnicodeTextPolicy::Reject);
    }

    bool OnKeyDown(DxUi::WindowHost& host, UINT key, UINT modifiers) override
    {
        if (DxUi::ModifiersContainCtrl(modifiers) && ! DxUi::ModifiersContainAlt(modifiers) && key == 'C')
            return OnCopy(host);
        if (! DxUi::ModifiersContainCtrl(modifiers) && ! DxUi::ModifiersContainAlt(modifiers) && (key == VK_RETURN || key == VK_SPACE))
        {
            if (_openDetail)
                _openDetail();
            return true;
        }
        if (key == VK_ESCAPE && host.HasTooltip())
        {
            static_cast<void>(host.ClearTooltip());
            return true;
        }
        return DxUi::Control::OnKeyDown(host, key, modifiers);
    }

    bool OnMouseDown(DxUi::WindowHost& host, D2D1_POINT_2F, bool rightButton, UINT) override
    {
        if (rightButton)
            return false;
        host.SetFocusControl(this);
        if (_openDetail)
            _openDetail();
        return true;
    }

    bool OnMouseMove(DxUi::WindowHost& host, D2D1_POINT_2F point, UINT modifiers) override
    {
#ifdef ENABLE_TESTS
        NoteDetachedInput();
#endif
        return DxUi::Control::OnMouseMove(host, point, modifiers);
    }

    bool OnMouseUp(DxUi::WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override
    {
#ifdef ENABLE_TESTS
        NoteDetachedInput();
#endif
        return DxUi::Control::OnMouseUp(host, point, rightButton, modifiers);
    }

    bool OnMouseLeave(DxUi::WindowHost& host) override
    {
        return DxUi::Control::OnMouseLeave(host);
    }

    bool Tick(DxUi::WindowHost& host, uint64_t nowTickMs) override
    {
        return DxUi::Control::Tick(host, nowTickMs);
    }

#ifdef ENABLE_TESTS
    // Self-test witness: a region a reconcile removed is kept alive (and counted here) only while a
    // self-test retains retired controls, so pointer input a panel still routes to it is observable
    // instead of a use-after-free.
    void DebugMarkDetached(uint64_t* detachedInputCount) noexcept
    {
        _debugDetachedInputCount = detachedInputCount;
    }

    // Self-test only: the region's own hover flag stays as it is while a panel still records it as
    // the hovered child, so a test can show the panel's pointer is dropped without trusting the flag.
    void DebugFreezeHoverFlag() noexcept
    {
        _debugHoverFlagFrozen = true;
    }
#endif

protected:
    void OnFocusChanged(DxUi::WindowHost& host, bool focused) override
    {
        DxUi::Control::OnFocusChanged(host, focused);
        if (focused)
            ShowPath(host);
        else
            static_cast<void>(host.ClearTooltip());
    }

    void OnHoverChanged(DxUi::WindowHost& host, bool hovered) override
    {
#ifdef ENABLE_TESTS
        NoteDetachedInput();
        if (_debugHoverFlagFrozen)
            return;
#endif
        DxUi::Control::OnHoverChanged(host, hovered);
    }

    void OnCaptureLost(DxUi::WindowHost& host) override
    {
        DxUi::Control::OnCaptureLost(host);
    }

private:
#ifdef ENABLE_TESTS
    void NoteDetachedInput() const noexcept
    {
        if (_debugDetachedInputCount)
            ++*_debugDetachedInputCount;
    }

    uint64_t* _debugDetachedInputCount = nullptr;
    bool _debugHoverFlagFrozen         = false;
#endif

    void ShowPath(DxUi::WindowHost& host) const
    {
        // The tooltip follows every frame; the accessible name is published on a bounded cadence.
        const auto bounds = GetBounds();
        static_cast<void>(host.SetTooltip(std::wstring(GetTooltipText()), D2D1::Point2F(bounds.left, bounds.bottom)));
    }

    std::wstring _path;
    std::wstring _tooltipText;
    D2D1_RECT_F _tooltipBounds{};
    std::function<void()> _openDetail;
};

// Consumer-specific prepared chrome shares the retained controls' presentation. The
// owning popup detaches its root before destroying these referenced resource slots.
class FileOperationsPopupRoot final : public DxUi::Panel
{
public:
    FileOperationsPopupRoot(const wil::com_ptr<ID2D1DeviceContext>& context, const wil::com_ptr<ID2D1CommandList>& background) noexcept
        : _context(context),
          _background(background)
    {
    }

    FileOperationsPopupRoot(const FileOperationsPopupRoot&)            = delete;
    FileOperationsPopupRoot(FileOperationsPopupRoot&&)                 = delete;
    FileOperationsPopupRoot& operator=(const FileOperationsPopupRoot&) = delete;
    FileOperationsPopupRoot& operator=(FileOperationsPopupRoot&&)      = delete;

    void Paint(DxUi::WindowHost& host) const override
    {
        if (auto* context = host.GetDeviceContext(); context && context == _context.get() && _background)
        {
            D2D1_MATRIX_3X2_F previousTransform{};
            context->GetTransform(&previousTransform);
            // Existing popup drawing uses physical pixels; hosted controls use DIPs.
            const float scale = 96.0f / std::max(1.0f, host.GetDpi());
            context->SetTransform(D2D1::Matrix3x2F::Scale(scale, scale) * previousTransform);
            context->DrawImage(_background.get());
            context->SetTransform(previousTransform);
        }
        DxUi::Panel::Paint(host);
    }

    // popup-b#1: reconciling the hosted tree moves a retired child out of its slot; drop the empty slots.
    void CompactChildren() noexcept
    {
        std::erase(AccessChildren(), nullptr);
    }

private:
    const wil::com_ptr<ID2D1DeviceContext>& _context;
    const wil::com_ptr<ID2D1CommandList>& _background;
};

// The hosted body and conflict comparison panels. Reconciling the hosted tree (popup-b#1) moves a
// retired child out of its slot; the panel and its other children live on.
class FileOperationsHostedScrollPanel final : public DxUi::ScrollPanel
{
public:
    FileOperationsHostedScrollPanel()                                                  = default;
    FileOperationsHostedScrollPanel(const FileOperationsHostedScrollPanel&)            = delete;
    FileOperationsHostedScrollPanel(FileOperationsHostedScrollPanel&&)                 = delete;
    FileOperationsHostedScrollPanel& operator=(const FileOperationsHostedScrollPanel&) = delete;
    FileOperationsHostedScrollPanel& operator=(FileOperationsHostedScrollPanel&&)      = delete;

    void CompactChildren() noexcept
    {
        std::erase(AccessChildren(), nullptr);
    }
};

// I35 R10.2: the stall band of one throughput graph, painted right after the graph over the same
// bounds. It hatches every sample spent stalled, mapping sample indexes the way
// DxUi::ThroughputGraph does (x = left + index * width / (count - 1)). It has no semantics of its own:
// the stall is named in the rate label text, and pointer input goes to the graph underneath.
class FileOperationsGraphStallBand final : public DxUi::Control
{
public:
    FileOperationsGraphStallBand()                                               = default;
    FileOperationsGraphStallBand(const FileOperationsGraphStallBand&)            = delete;
    FileOperationsGraphStallBand(FileOperationsGraphStallBand&&)                 = delete;
    FileOperationsGraphStallBand& operator=(const FileOperationsGraphStallBand&) = delete;
    FileOperationsGraphStallBand& operator=(FileOperationsGraphStallBand&&)      = delete;

    void SetStalledSamples(std::span<const uint8_t> stalled)
    {
        _stalled.assign(stalled.begin(), stalled.end());
    }

    void Paint(DxUi::WindowHost& host) const override
    {
        const D2D1_RECT_F bounds = GetBounds();
        const size_t count       = _stalled.size();
        if (count < 2u || bounds.right <= bounds.left || bounds.bottom <= bounds.top)
            return;
        const auto& palette = host.GetTheme();
        D2D1_COLOR_F color  = palette.disabledText;
        color.a             = palette.highContrast ? 1.0f : 0.45f;
        auto* const context = host.GetDeviceContext();
        auto* const brush   = host.GetSolidBrush(color);
        if (! context || ! brush)
            return;
        constexpr float kHatchStepDip = 5.0f;
        const float step              = (bounds.right - bounds.left) / static_cast<float>(count - 1u);
        const float rise              = bounds.bottom - bounds.top;
        for (size_t first = 0u; first < count;)
        {
            if (_stalled[first] == 0u)
            {
                ++first;
                continue;
            }
            size_t last = first;
            while (last + 1u < count && _stalled[last + 1u] != 0u)
                ++last;
            // A sample plots the bucket that ended at it, so a run starts at the sample before it.
            const D2D1_RECT_F band = D2D1::RectF(bounds.left + step * static_cast<float>(first > 0u ? first - 1u : 0u),
                                                 bounds.top,
                                                 bounds.left + step * static_cast<float>(last),
                                                 bounds.bottom);
            if (band.right > band.left)
            {
                context->PushAxisAlignedClip(band, D2D1_ANTIALIAS_MODE_ALIASED);
                for (float x = band.left - rise; x < band.right; x += kHatchStepDip)
                    context->DrawLine(D2D1::Point2F(x, band.bottom), D2D1::Point2F(x + rise, band.top), brush, 1.0f);
                context->PopAxisAlignedClip();
            }
            first = last + 1u;
        }
    }

    [[nodiscard]] DxUi::Control* HitTest(D2D1_POINT_2F /*point*/) override
    {
        return nullptr;
    }

    [[nodiscard]] const DxUi::Control* HitTest(D2D1_POINT_2F /*point*/) const override
    {
        return nullptr;
    }

private:
    std::vector<uint8_t> _stalled;
};

// I35 (owner report 2026-10-13): the verification half of a Verify On bar, painted right after the hosted bar over
// the same bounds. DxUi::ProgressBar draws its divider in the track color and leaves the part not verified as plain
// track, so a finished copy whose verification did not cover every byte read as an unfinished copy beside its 100%.
// The marks draw the divider in the verification color at any share and, once the task has finished, a hatched
// outline of the part not verified (a running task's pending verification stays plain track behind the divider).
// No semantics of their own: the bar's accessible text states both halves, and pointer input goes to the bar.
class FileOperationsVerificationHalfMarks final : public DxUi::Control
{
public:
    FileOperationsVerificationHalfMarks()                                                      = default;
    FileOperationsVerificationHalfMarks(const FileOperationsVerificationHalfMarks&)            = delete;
    FileOperationsVerificationHalfMarks(FileOperationsVerificationHalfMarks&&)                 = delete;
    FileOperationsVerificationHalfMarks& operator=(const FileOperationsVerificationHalfMarks&) = delete;
    FileOperationsVerificationHalfMarks& operator=(FileOperationsVerificationHalfMarks&&)      = delete;

    void SetMarks(bool shown, double verifiedPercent, D2D1_COLOR_F color, bool outlinePending) noexcept
    {
        _shown           = shown;
        _verifiedPercent = std::clamp(verifiedPercent, 0.0, 100.0);
        _color           = color;
        _outlinePending  = outlinePending;
    }

    [[nodiscard]] bool OutlinesPending() const noexcept
    {
        return _outlinePending;
    }

    [[nodiscard]] bool MarksShown() const noexcept
    {
        return _shown;
    }

    [[nodiscard]] double VerifiedPercent() const noexcept
    {
        return _verifiedPercent;
    }

    void Paint(DxUi::WindowHost& host) const override
    {
        const D2D1_RECT_F bounds = GetBounds();
        if (! _shown || bounds.right <= bounds.left || bounds.bottom <= bounds.top)
            return;
        auto* const context = host.GetDeviceContext();
        D2D1_COLOR_F faint  = _color;
        faint.a             = host.GetTheme().highContrast ? 1.0f : 0.6f;
        auto* const brush   = host.GetSolidBrush(_color);
        auto* const pending = host.GetSolidBrush(faint);
        if (! context || ! brush || ! pending)
            return;
        // The popup sets the bar's track height to its bounds, so the track is the bounds; its halves meet midway.
        const float height     = bounds.bottom - bounds.top;
        const float radius     = height * 0.5f;
        const float midpoint   = bounds.left + (bounds.right - bounds.left) * 0.5f;
        const float verifiedTo = midpoint + static_cast<float>(_verifiedPercent / 100.0) * (bounds.right - midpoint);
        if (_outlinePending && bounds.right - verifiedTo > 0.5f)
        {
            // The outline of the part not verified (its top and bottom edges and the track's rounded end), hatched
            // inside up to the rounded end.
            context->PushAxisAlignedClip(D2D1::RectF(verifiedTo, bounds.top, bounds.right, bounds.bottom), D2D1_ANTIALIAS_MODE_ALIASED);
            const float inset = 0.5f;
            context->DrawRoundedRectangle(
                D2D1::RoundedRect(D2D1::RectF(bounds.left + inset, bounds.top + inset, bounds.right - inset, bounds.bottom - inset), radius, radius),
                pending,
                1.0f);
            context->PopAxisAlignedClip();
            const float hatchRight = (std::max)(verifiedTo, bounds.right - radius);
            if (hatchRight - verifiedTo > 0.5f)
            {
                constexpr float kHatchStepDip = 4.0f;
                context->PushAxisAlignedClip(D2D1::RectF(verifiedTo, bounds.top, hatchRight, bounds.bottom), D2D1_ANTIALIAS_MODE_ALIASED);
                for (float x = verifiedTo - height; x < hatchRight; x += kHatchStepDip)
                    context->DrawLine(D2D1::Point2F(x, bounds.bottom), D2D1::Point2F(x + height, bounds.top), pending, 1.0f);
                context->PopAxisAlignedClip();
            }
        }
        context->DrawLine(D2D1::Point2F(midpoint, bounds.top), D2D1::Point2F(midpoint, bounds.bottom), brush, 1.5f);
    }

    [[nodiscard]] DxUi::Control* HitTest(D2D1_POINT_2F /*point*/) override
    {
        return nullptr;
    }

    [[nodiscard]] const DxUi::Control* HitTest(D2D1_POINT_2F /*point*/) const override
    {
        return nullptr;
    }

private:
    bool _shown             = false;
    double _verifiedPercent = 0.0;
    D2D1_COLOR_F _color{};
    bool _outlinePending = false;
};

constexpr wchar_t kFileOperationsPopupClassName[] = L"RedSalamander.FileOperationsPopup";

constexpr UINT_PTR kFileOperationsPopupTimerId                      = 1;
constexpr UINT kFileOperationsPopupTimerIntervalMs                  = 100;
constexpr UINT kFileOperationsPopupUiMotionTimerIntervalMs          = 16;
constexpr UINT kFileOperationsPopupWaitingTimerIntervalMs           = 1000;
constexpr ULONGLONG kFileOperationsPopupUiMotionDurationMs          = 260ull;
constexpr ULONGLONG kFileOperationsPopupUiMotionTimerWindowMs       = 520ull;
constexpr ULONGLONG kFileOperationsPopupFramePrepareFailureWindowMs = 2000ull;
constexpr ULONGLONG kFileOperationsPopupFramePrepareRetryBaseMs     = 100ull;
constexpr ULONGLONG kFileOperationsPopupFramePrepareRetryMaxMs      = 1600ull;
constexpr ULONGLONG kFileOperationsPopupHostReattachRetryBaseMs     = 1000ull;
constexpr ULONGLONG kFileOperationsPopupHostReattachRetryMaxMs      = 60000ull;
constexpr uint32_t kFileOperationsPopupFramePrepareFailureLimit     = 5u;
constexpr ULONGLONG kFileOperationsPopupAccessibleTextIntervalMs    = 1000ull;
constexpr ULONGLONG kRateSampleBucketMs                             = 100ull;
constexpr UINT kFileOperationsPopupDeferredSpeedLimitPromptMessage  = WM_APP + 0x71;
constexpr wchar_t kFileOperationsSpeedLimitPromptClassName[]        = L"RedSalamander.FileOperations.SpeedLimitPrompt";
constexpr UINT kFileOperationsSpeedLimitPromptDeferredActionMessage = WM_APP + 0x74;
constexpr std::wstring_view kEllipsisText                           = L"\u2026";
constexpr float kFileOperationsExpandedCardHeightDip                = 280.0f;
constexpr float kFileOperationsExpandedTransferCardHeightDip        = 244.0f;
// Header, three ordinary text rows, a separated progress bar and the action row.
// Deletion has an item rate; it does not reserve a byte-throughput plot.
constexpr float kFileOperationsExpandedDeleteCardHeightDip     = 152.0f;
constexpr int kFileOperationsPopupMinClientHeightDip           = 320;
constexpr int kFileOperationsPopupFooterOnlyMinClientHeightDip = 96;
constexpr uint32_t kCompletedOverflowActionShowLog             = 1u;
constexpr uint32_t kCompletedOverflowActionExportIssues        = 2u;
constexpr uint32_t kCompletedOverflowActionFailedItems         = 3u;
constexpr uint32_t kCompletedOverflowActionOpenDestination     = 4u;
constexpr uint32_t kCompletedOverflowActionRevealDestination   = 5u;
constexpr uint32_t kCompletedOverflowActionOpenSource          = 6u;
constexpr uint32_t kCompletedOverflowActionSelectRetained      = 7u;
constexpr uint32_t kCompletedOverflowActionCutRetainedAgain    = 8u;
constexpr uint32_t kFooterPauseResumeAllPauseAction            = 1u;
constexpr uint32_t kFooterPauseResumeAllResumeAction           = 2u;
constexpr uint32_t kFooterQueueModeQueueAction                 = 1u;
constexpr uint32_t kFooterQueueModeParallelAction              = 2u;
constexpr ULONGLONG kTaskbarListRetryDelayMs                   = 1000ull;

[[nodiscard]] bool OpensButtonFlyout(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
{
    using Kind = FileOperationsPopupInternal::PopupHitTest::Kind;
    return kind == Kind::FooterQueueMode || kind == Kind::TaskSpeedLimit || kind == Kind::TaskCompletedMore || kind == Kind::TaskConflictMore;
}

[[nodiscard]] bool UsesSelectorButtonChrome(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
{
    using Kind = FileOperationsPopupInternal::PopupHitTest::Kind;
    return kind == Kind::FooterQueueMode || kind == Kind::TaskSpeedLimit;
}

[[nodiscard]] DxUi::ButtonVariant HostedButtonVariant(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
{
    using Kind = FileOperationsPopupInternal::PopupHitTest::Kind;
    if (kind == Kind::FooterToggleDetails || kind == Kind::CompletedGroupToggle || kind == Kind::TaskToggleCollapse || kind == Kind::TaskToggleThroughput)
    {
        return DxUi::ButtonVariant::Disclosure;
    }
    if (UsesSelectorButtonChrome(kind))
    {
        return DxUi::ButtonVariant::Selector;
    }
    if (OpensButtonFlyout(kind))
    {
        return DxUi::ButtonVariant::DropDown;
    }
    if (kind == FileOperationsPopupInternal::PopupHitTest::Kind::FooterOptions)
    {
        return DxUi::ButtonVariant::IconOnly;
    }
    return DxUi::ButtonVariant::Standard;
}

[[nodiscard]] std::wstring FormatSpeedLimitSelectorText(uint64_t bytesPerSecond)
{
    return bytesPerSecond == 0 ? LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_BUTTON_UNLIMITED)
                               : FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_BUTTON_BYTES, FormatBytesCompact(bytesPerSecond));
}

[[nodiscard]] std::wstring FormatDiscoveryProgressText(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    if (task.operation == FILESYSTEM_COPY && task.copiedFileCountKnown && task.copiedFolderCountKnown)
        return FormatStringResource(
            nullptr, IDS_FMT_FILEOPS_COPIED_FILES_FOLDERS_WRITTEN, task.copiedFileCount, task.copiedFolderCount, FormatBytesCompact(task.completedBytes));
    if (task.operation == FILESYSTEM_COPY && task.copiedFileCountKnown)
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_COPIED_FILES_SO_FAR, task.copiedFileCount, FormatBytesCompact(task.completedBytes));
    if (task.operation == FILESYSTEM_COPY)
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_WRITTEN_BYTES, FormatBytesCompact(task.completedBytes));
    if (task.operation == FILESYSTEM_DELETE && task.completedWorkCountsKnown)
        return FormatStringResource(nullptr,
                                    IDS_FMT_FILEOPS_DELETED_FILES_FOLDERS_BYTES,
                                    task.completedWorkFileCount,
                                    task.completedWorkFolderCount,
                                    FormatBytesCompact(task.completedWorkBytes));
    if (task.operation == FILESYSTEM_DELETE)
        return LoadStringResource(nullptr, IDS_FILEOPS_DELETING_WITHOUT_COUNTS);
    return FormatStringResource(nullptr,
                                task.operation == FILESYSTEM_DELETE ? IDS_FMT_FILEOPS_DELETION_PROGRESS : IDS_FMT_FILEOPS_DISCOVERY_PROGRESS,
                                FormatBytesCompact(task.completedBytes),
                                static_cast<unsigned long>(task.completedItems));
}

[[nodiscard]] bool TaskUsesSkippedCountingWorkSummary(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return ! task.finished && task.discoverySkipped && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_DELETE);
}

[[nodiscard]] std::wstring FormatDiscoveredSoFarText(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    if (task.operation == FILESYSTEM_COPY)
    {
        if (task.discoveredBytesKnown || task.discoveredTotalBytes > 0u)
            return FormatStringResource(nullptr,
                                        IDS_FMT_FILEOPS_FOUND_FILE_COUNTS_SIZE_SO_FAR,
                                        task.discoveredFileCount,
                                        task.discoveredDirectoryCount,
                                        FormatBytesCompact(task.discoveredTotalBytes));
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOUND_FILE_COUNTS_SO_FAR, task.discoveredFileCount, task.discoveredDirectoryCount);
    }
    const uint64_t discoveredBytes = std::max(task.discoveredTotalBytes, task.completedBytes);
    const uint64_t discoveredItems =
        std::max<uint64_t>(static_cast<uint64_t>(task.discoveredFileCount) + static_cast<uint64_t>(task.discoveredDirectoryCount), task.completedItems);
    const auto itemCount = static_cast<unsigned long>((std::min<uint64_t>)(discoveredItems, ULONG_MAX));
    // R6.9 / P5: never a literal "?" and never "0 B" for a size that is unknown (a sizeless provider)
    // or not yet counted; without a byte total the row names only the items found so far.
    if (discoveredItems == 0u)
        return LoadStringResource(nullptr, IDS_FILEOPS_DISCOVERED_NOTHING_YET);
    if (discoveredBytes == 0u)
        return FormatStringResource(
            nullptr, task.discoverySkipped ? IDS_FMT_FILEOPS_DISCOVERED_AS_NEEDED_ITEMS_ONLY : IDS_FMT_FILEOPS_DISCOVERED_ITEMS_ONLY, itemCount);
    return FormatStringResource(nullptr,
                                task.discoverySkipped ? IDS_FMT_FILEOPS_DISCOVERED_AS_NEEDED : IDS_FMT_FILEOPS_DISCOVERED_SO_FAR,
                                FormatBytesCompact(discoveredBytes),
                                itemCount);
}

[[nodiscard]] bool FileProgressIsVisible(const FileOperationsPopupInternal::TaskSnapshot::InFlightFileSnapshot& entry, ULONGLONG nowTick) noexcept
{
    if (entry.explicitLifecycle)
        return ! entry.terminal || (entry.lastUpdateTick != 0u && nowTick >= entry.lastUpdateTick && nowTick - entry.lastUpdateTick <= 300u);
    return entry.totalBytes == 0u || entry.completedBytes < entry.totalBytes ||
           (entry.lastUpdateTick != 0u && nowTick >= entry.lastUpdateTick && nowTick - entry.lastUpdateTick <= 300u);
}

enum class PopupDisplayRowKind : uint8_t
{
    Task,
    CompletedGroup,
};

struct PopupDisplayRow
{
    PopupDisplayRowKind kind = PopupDisplayRowKind::Task;
    size_t taskIndex         = 0u;
};

#ifdef ENABLE_TESTS
constexpr UINT kFileOperationsSpeedLimitPromptDebugMessage = WM_APP + 0x73;
std::atomic<unsigned int> g_fileOperationsTaskbarListForcedFailures{0};
std::atomic<unsigned int> g_fileOperationsDxUiHostAttachForcedFailures{0};
std::atomic<unsigned int> g_fileOperationsD2DTargetForcedFailures{0};
std::atomic<unsigned int> g_fileOperationsLayoutMeasurementForcedFailures{0};
std::atomic<unsigned int> g_fileOperationsFailureSurfaceChildCreationOrdinal{0};
std::atomic<unsigned int> g_fileOperationsFailureSurfaceButtonSetupOrdinal{0};

[[nodiscard]] bool ConsumeForcedFailure(std::atomic<unsigned int>& remaining) noexcept
{
    unsigned int value = remaining.load(std::memory_order_acquire);
    while (value > 0u)
    {
        if (remaining.compare_exchange_weak(value, value - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
            return true;
    }
    return false;
}

[[nodiscard]] bool ShouldFailFailureSurfaceChildCreation(unsigned int ordinal) noexcept
{
    unsigned int expected = ordinal;
    return g_fileOperationsFailureSurfaceChildCreationOrdinal.compare_exchange_strong(expected, 0u, std::memory_order_acq_rel, std::memory_order_acquire);
}

[[nodiscard]] bool ShouldFailFailureSurfaceButtonSetup(unsigned int ordinal) noexcept
{
    unsigned int expected = ordinal;
    return g_fileOperationsFailureSurfaceButtonSetupOrdinal.compare_exchange_strong(expected, 0u, std::memory_order_acq_rel, std::memory_order_acquire);
}

[[nodiscard]] bool IsActuallyVisibleChildWindow(HWND hwnd) noexcept
{
    return hwnd && IsWindow(hwnd) != FALSE && IsWindowVisible(hwnd) != FALSE && (GetWindowLongPtrW(hwnd, GWL_STYLE) & WS_CHILD) != 0;
}
#endif

[[nodiscard]] uint64_t PerfNowUs() noexcept
{
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now().time_since_epoch()).count());
}

[[nodiscard]] uint64_t PerfElapsedUs(uint64_t startUs) noexcept
{
    const uint64_t nowUs = PerfNowUs();
    return (nowUs >= startUs) ? (nowUs - startUs) : 0u;
}

[[nodiscard]] POINT ResolveOwnerCenterScreenPoint(HWND hwnd) noexcept
{
    POINT point{};
    RECT client{};
    if (hwnd && GetClientRect(hwnd, &client) != FALSE)
    {
        point.x = client.left + ((client.right - client.left) / 2);
        point.y = client.top + ((client.bottom - client.top) / 2);
        static_cast<void>(ClientToScreen(hwnd, &point));
    }
    return point;
}

float DipsToPixels(float dip, UINT dpi) noexcept
{
    return dip * (static_cast<float>(dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI));
}

int DipsToPixels(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

UINT FileOperationsTaskbarButtonCreatedMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"TaskbarButtonCreated");
    return message;
}

float PixelsToDips(float px, UINT dpi) noexcept
{
    if (dpi == 0)
    {
        return px;
    }
    return px * (static_cast<float>(USER_DEFAULT_SCREEN_DPI) / static_cast<float>(dpi));
}

[[nodiscard]] float EaseFileOperationsUiMotionFraction(ULONGLONG elapsedMs, ULONGLONG durationMs) noexcept
{
    if (durationMs == 0 || elapsedMs >= durationMs)
    {
        return 1.0f;
    }

    const float t = std::clamp(static_cast<float>(elapsedMs) / static_cast<float>(durationMs), 0.0f, 1.0f);
    return t * t * (3.0f - 2.0f * t);
}

[[nodiscard]] bool RectsNearEqual(const RECT& lhs, const RECT& rhs, int tolerancePx = 1) noexcept
{
    return std::abs(lhs.left - rhs.left) <= tolerancePx && std::abs(lhs.top - rhs.top) <= tolerancePx && std::abs(lhs.right - rhs.right) <= tolerancePx &&
           std::abs(lhs.bottom - rhs.bottom) <= tolerancePx;
}

[[nodiscard]] LONG LerpLong(LONG from, LONG to, float fraction) noexcept
{
    const float value = static_cast<float>(from) + (static_cast<float>(to - from) * fraction);
    return static_cast<LONG>(std::lround(value));
}

[[nodiscard]] RECT LerpRect(const RECT& from, const RECT& to, float fraction) noexcept
{
    return RECT{LerpLong(from.left, to.left, fraction),
                LerpLong(from.top, to.top, fraction),
                LerpLong(from.right, to.right, fraction),
                LerpLong(from.bottom, to.bottom, fraction)};
}

D2D1_RECT_F RectPixelsToDips(const RECT& rc, UINT dpi) noexcept
{
    return D2D1::RectF(PixelsToDips(static_cast<float>(rc.left), dpi),
                       PixelsToDips(static_cast<float>(rc.top), dpi),
                       PixelsToDips(static_cast<float>(rc.right), dpi),
                       PixelsToDips(static_cast<float>(rc.bottom), dpi));
}

[[nodiscard]] bool DirectWriteFormatHasGlyph(IDWriteFactory* factory, IDWriteTextFormat* format, wchar_t glyph) noexcept
{
    if (! factory || ! format || glyph == 0)
    {
        return false;
    }

    const UINT32 familyLength = format->GetFontFamilyNameLength();
    if (familyLength == 0)
    {
        return false;
    }

    std::wstring familyName(familyLength + 1u, L'\0');
    if (FAILED(format->GetFontFamilyName(familyName.data(), familyLength + 1u)))
    {
        return false;
    }
    familyName.resize(familyLength);

    wil::com_ptr<IDWriteFontCollection> collection;
    if (FAILED(factory->GetSystemFontCollection(collection.addressof(), FALSE)) || ! collection)
    {
        return false;
    }

    UINT32 familyIndex = 0;
    BOOL familyExists  = FALSE;
    if (FAILED(collection->FindFamilyName(familyName.c_str(), &familyIndex, &familyExists)) || familyExists == FALSE)
    {
        return false;
    }

    wil::com_ptr<IDWriteFontFamily> fontFamily;
    if (FAILED(collection->GetFontFamily(familyIndex, fontFamily.addressof())) || ! fontFamily)
    {
        return false;
    }

    wil::com_ptr<IDWriteFont> font;
    if (FAILED(fontFamily->GetFirstMatchingFont(format->GetFontWeight(), format->GetFontStretch(), format->GetFontStyle(), font.addressof())) || ! font)
    {
        return false;
    }

    BOOL hasGlyph = FALSE;
    if (FAILED(font->HasCharacter(static_cast<UINT32>(glyph), &hasGlyph)))
    {
        return false;
    }
    return hasGlyph != FALSE;
}

bool IsRectFullyVisible(const RECT& rect) noexcept
{
    if (rect.right <= rect.left || rect.bottom <= rect.top)
    {
        return false;
    }

    HMONITOR monitor = MonitorFromRect(&rect, MONITOR_DEFAULTTONULL);
    if (! monitor)
    {
        return false;
    }

    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (! GetMonitorInfoW(monitor, &mi))
    {
        return false;
    }

    const RECT& work = mi.rcWork;
    return rect.left >= work.left && rect.top >= work.top && rect.right <= work.right && rect.bottom <= work.bottom;
}

float Clamp01(float v) noexcept
{
    return std::clamp(v, 0.0f, 1.0f);
}

// I35 (owner report 2026-10-13): the share of a finished Verify On task's bytes its verification covered, 0..100. A
// task with no bytes counts as verified when its record says so.
[[nodiscard]] double FinishedVerificationPercent(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    const uint64_t total = task.verificationTotalBytes > 0u ? task.verificationTotalBytes : task.totalBytes;
    if (total == 0u)
        return task.verificationState == static_cast<uint8_t>(FileOperations::VerificationState::Verified) ? 100.0 : 0.0;
    return std::clamp(static_cast<double>((std::min)(task.verificationCompletedBytes, total)) / static_cast<double>(total), 0.0, 1.0) * 100.0;
}

// A finished, successful Verify On task whose verification did not cover every byte: its card says so beside the bar.
[[nodiscard]] bool FinishedTaskVerificationIncomplete(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.finished && SUCCEEDED(task.resultHr) && task.verificationRequested && task.operation != FILESYSTEM_DELETE &&
           FinishedVerificationPercent(task) < 100.0;
}

float ComputeFileOperationsTaskCompleteFractionForDisplay(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (task.finished && SUCCEEDED(task.resultHr))
    {
        // Verify On: the transfer half is full and the verification half shows the share verified, so a finished task
        // reads 100% only when every byte was verified (UI_FileOperationsPopup.md, the whole-task bar).
        if (task.verificationRequested && task.operation != FILESYSTEM_DELETE)
            return Clamp01(0.5f + 0.5f * static_cast<float>(FinishedVerificationPercent(task) / 100.0));
        return 1.0f;
    }

    if (task.operation == FILESYSTEM_DELETE)
    {
        if (task.discoveryClosed && task.totalBytes > 0 && task.completedBytes > 0)
        {
            return Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
        }
        if (task.discoveryClosed && task.totalItems > 0)
        {
            return Clamp01(static_cast<float>(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems)));
        }
        return 0.0f;
    }

    if (task.discoveryClosed && task.totalBytes > 0 && (task.operation != FILESYSTEM_COPY || task.finished || task.discoveredBytesKnown))
    {
        const float transferFraction = Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
        if (task.verificationRequested)
        {
            const uint64_t verificationTotal = task.verificationTotalBytes > 0 ? task.verificationTotalBytes : task.totalBytes;
            const float verificationFraction =
                verificationTotal > 0
                    ? Clamp01(static_cast<float>(static_cast<double>(task.verificationCompletedBytes) / static_cast<double>(verificationTotal)))
                    : 0.0f;
            return 0.5f * transferFraction + 0.5f * verificationFraction;
        }
        return transferFraction;
    }
    if (task.operation == FILESYSTEM_COPY && ! task.finished)
    {
        if (task.discoveryClosed && task.discoveredBytesKnown && task.totalBytes == 0u && ! task.verificationRequested && task.copiedFileCountKnown &&
            task.discoveredFileCount > 0u)
            return Clamp01(static_cast<float>(static_cast<double>(task.copiedFileCount) / static_cast<double>(task.discoveredFileCount)));
        return 0.0f;
    }
    if (task.discoveryClosed && task.totalItems > 0)
    {
        return Clamp01(static_cast<float>(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems)));
    }

    return 0.0f;
}

void NormalizeCompletedTaskSnapshotForDisplay(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    // Succeeded-with-warnings tasks keep their REAL counters: forcing "N of N"/100% would hide
    // that children were skipped. Only a clean success rounds the display up. Typed skipped roots
    // keep their counters too (I35 D-5.2: a skip is no longer a counted warning).
    if (! task.finished || FAILED(task.resultHr) || task.warningCount > 0 || task.errorCount > 0 || task.skippedItemCount > 0u)
    {
        return;
    }

    if (task.totalItems > 0)
    {
        task.completedItems = task.totalItems;
    }
    if (task.totalBytes > 0)
    {
        task.completedBytes = task.totalBytes;
    }
    if (task.itemTotalBytes > 0)
    {
        task.itemCompletedBytes = task.itemTotalBytes;
    }
    // Verify On: a clean success that verified every item shows the verification half full too, so the
    // bar agrees with the 100% beside it. A verification problem keeps the real verification counters.
    if (task.verificationRequested && task.verificationState == static_cast<uint8_t>(FileOperations::VerificationState::Verified))
    {
        const uint64_t verificationTotal = task.verificationTotalBytes > 0 ? task.verificationTotalBytes : task.totalBytes;
        task.verificationTotalBytes      = verificationTotal;
        task.verificationCompletedBytes  = verificationTotal;
    }
}

[[nodiscard]] bool CompletedTaskCanUseDestinationActions(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.finished && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) && task.destinationPane.has_value() &&
           ! task.destinationPluginId.empty() && ! task.destinationPluginShortId.empty() && ! task.destinationFolder.empty();
}

struct CompletedTaskRevealLocation
{
    std::filesystem::path folder;
    std::wstring leaf;
};

[[nodiscard]] std::optional<CompletedTaskRevealLocation> ResolveCompletedTaskRevealLocation(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (! CompletedTaskCanUseDestinationActions(task) || task.currentDestinationPath.empty())
    {
        return std::nullopt;
    }

    const std::wstring_view destinationPath(task.currentDestinationPath);
    const size_t separator = destinationPath.find_last_of(L"/\\");
    if (separator == std::wstring_view::npos || separator + 1u >= destinationPath.size())
    {
        return std::nullopt;
    }

    const bool isRootSeparator = separator == 0u || (separator == 2u && destinationPath.size() > 2u && destinationPath[1] == L':' &&
                                                     (destinationPath[2] == L'\\' || destinationPath[2] == L'/'));
    const size_t folderLength  = separator + (isRootSeparator ? 1u : 0u);
    std::wstring folder(destinationPath.substr(0, folderLength));
    std::wstring leaf(destinationPath.substr(separator + 1u));

    if (NavigationLocation::EqualsNoCase(destinationPath, task.destinationFolder.native()))
    {
        return std::nullopt;
    }

    return CompletedTaskRevealLocation{std::filesystem::path(std::move(folder)), std::move(leaf)};
}

[[nodiscard]] bool CompletedTaskHasOverflowActions(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.warningCount > 0 || task.errorCount > 0 || task.hasRetainedDiagnostics || CompletedTaskCanUseDestinationActions(task) ||
           ResolveCompletedTaskRevealLocation(task).has_value() || task.hasSourceActionPaths || task.exactRetainedSourceActionAvailable;
}

[[nodiscard]] std::wstring FormatFileTimeLocalCompact(__int64 fileTime) noexcept
{
    if (fileTime <= 0)
    {
        return {};
    }

    ULARGE_INTEGER uli{};
    uli.QuadPart = static_cast<ULONGLONG>(fileTime);

    FILETIME ft{};
    ft.dwLowDateTime  = uli.LowPart;
    ft.dwHighDateTime = uli.HighPart;

    SYSTEMTIME utc{};
    SYSTEMTIME local{};
    if (FileTimeToSystemTime(&ft, &utc) == FALSE || SystemTimeToTzSpecificLocalTime(nullptr, &utc, &local) == FALSE)
    {
        return {};
    }

    const int dateLength = GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, DATE_SHORTDATE, &local, nullptr, nullptr, 0, nullptr);
    const int timeLength = GetTimeFormatEx(LOCALE_NAME_USER_DEFAULT, TIME_NOSECONDS, &local, nullptr, nullptr, 0);
    if (dateLength <= 1 || timeLength <= 1)
    {
        return {};
    }

    std::wstring dateText(static_cast<size_t>(dateLength), L'\0');
    std::wstring timeText(static_cast<size_t>(timeLength), L'\0');
    if (GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, DATE_SHORTDATE, &local, nullptr, dateText.data(), dateLength, nullptr) == 0 ||
        GetTimeFormatEx(LOCALE_NAME_USER_DEFAULT, TIME_NOSECONDS, &local, nullptr, timeText.data(), timeLength) == 0)
    {
        return {};
    }
    dateText.resize(static_cast<size_t>(dateLength - 1));
    timeText.resize(static_cast<size_t>(timeLength - 1));
    return FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFLICT_DATE_TIME, dateText, timeText);
}

[[nodiscard]] std::wstring FormatConflictMetadataText(const FileOperationsPopupInternal::TaskSnapshot::ConflictPromptSnapshot::ItemMetadata& metadata)
{
    if (! metadata.available)
    {
        return {};
    }

    std::wstring result;
    if (metadata.isDirectory)
    {
        result = LoadStringResource(nullptr, IDS_FOLDERVIEW_TYPE_FOLDER);
    }
    else if (metadata.sizeKnown)
    {
        result = FormatBytesCompact(metadata.sizeBytes);
    }

    const std::wstring timeText = FormatFileTimeLocalCompact(metadata.lastWriteTime);
    if (! timeText.empty())
    {
        if (! result.empty())
        {
            result.append(L" \u2022 ");
        }
        result.append(timeText);
    }

    return result;
}

using ConflictAction = FolderWindow::FileOperationState::Task::ConflictAction;
using ConflictBucket = FolderWindow::FileOperationState::Task::ConflictBucket;

float MeasureWrappedTextHeight(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text, float maxWidth, float fallbackLineHeight) noexcept;

// I35 R8.3: the card and the decision mirror share one prompt text (FolderWindow.FileOperations.PromptText).
[[nodiscard]] FileOperations::ConflictPromptTextFacts ConflictPromptTextFactsFor(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    const auto& conflict = task.conflict;
    FileOperations::ConflictPromptTextFacts facts{};
    facts.operation               = task.operation;
    facts.bucket                  = conflict.bucket;
    facts.status                  = conflict.status;
    facts.repeatUnavailableReason = conflict.repeatUnavailableReason;
    facts.sourceIsDirectory       = conflict.sourceIsDirectory;
    const auto offeredActions     = std::span(conflict.actions.data(), std::min(conflict.actionCount, conflict.actions.size()));
    facts.readOnlyDeleteOffered   = std::ranges::find(offeredActions, static_cast<uint8_t>(ConflictAction::DeleteReadOnly)) != offeredActions.end();
    facts.skipAllOffered          = std::ranges::find(offeredActions, static_cast<uint8_t>(ConflictAction::SkipAll)) != offeredActions.end();
    facts.attemptCount            = conflict.attemptCount;
    facts.overlapProblem          = conflict.overlapProblem;
    facts.overlapTaskId           = conflict.overlapTaskId;
    facts.sourcePath              = conflict.sourcePath;
    facts.destinationPath         = conflict.destinationPath;
    facts.sourceName              = conflict.sourceName;
    facts.sourceParentPath        = conflict.sourceParentPath;
    facts.destinationName         = conflict.destinationName;
    facts.destinationParentPath   = conflict.destinationParentPath;
    facts.consentDetail           = conflict.consentDetail;
    facts.consentNote             = conflict.consentNote;
    facts.factItemCountKnown      = conflict.factItemCountKnown;
    facts.factItemCount           = conflict.factItemCount;
    facts.factBytesKnown          = conflict.factBytesKnown;
    facts.factBytes               = conflict.factBytes;
    return facts;
}

[[nodiscard]] std::wstring_view CommonConflictName(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return FileOperations::CommonConflictPromptName(ConflictPromptTextFactsFor(task));
}

[[nodiscard]] std::wstring BuildConflictQuestion(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    return FileOperations::BuildConflictPromptQuestion(ConflictPromptTextFactsFor(task));
}

[[nodiscard]] std::wstring BuildConflictFactText(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    return FileOperations::BuildConflictPromptFactText(ConflictPromptTextFactsFor(task));
}

struct ConflictPromptContent final
{
    std::wstring state;
    std::wstring commonName;
    std::wstring question;
    std::wstring facts;
    std::wstring sourceLabel;
    std::wstring sourcePath;
    std::wstring sourceMetadata;
    std::wstring sourceName;
    std::wstring sourceParentPath;
    std::wstring destinationName;
    std::wstring destinationParentPath;
    std::wstring destinationLabel;
    std::wstring destinationPath;
    std::wstring destinationMetadata;
    std::wstring lastNote;
    std::wstring inspectionResult;
};

[[nodiscard]] ConflictPromptContent BuildConflictPromptContent(const FileOperationsPopupInternal::TaskSnapshot& task, std::wstring inspectionResult = {})
{
    ConflictPromptContent content{};
    content.inspectionResult = std::move(inspectionResult);
    if (! task.conflict.active)
    {
        return content;
    }

    content.state = LoadStringResource(nullptr, task.conflict.metadataLoading ? IDS_FILEOPS_CONFLICT_READING_DETAILS : IDS_FILEOPS_STATUS_WAITING_FOR_DECISION);
    if (! task.conflict.metadataLoading)
    {
        content.question = BuildConflictQuestion(task);
    }
    content.commonName            = CommonConflictName(task);
    content.facts                 = BuildConflictFactText(task);
    content.sourceLabel           = LoadStringResource(nullptr, task.operation == FILESYSTEM_DELETE ? IDS_FILEOPS_LABEL_DELETING : IDS_FILEOPS_LABEL_FROM);
    content.sourcePath            = task.conflict.sourcePath;
    content.sourceMetadata        = FormatConflictMetadataText(task.conflict.sourceMetadata);
    content.sourceName            = task.conflict.sourceName;
    content.sourceParentPath      = task.conflict.sourceParentPath;
    content.destinationName       = task.conflict.destinationName;
    content.destinationParentPath = task.conflict.destinationParentPath;
    if (task.operation != FILESYSTEM_DELETE)
    {
        content.destinationLabel    = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO);
        content.destinationPath     = task.conflict.destinationPath;
        content.destinationMetadata = FormatConflictMetadataText(task.conflict.destinationMetadata);
    }
    if (! task.lastDiagnosticMessage.empty())
    {
        content.lastNote = FormatStringResource(nullptr, IDS_FMT_FILEOPS_LAST_NOTE, task.lastDiagnosticMessage);
    }
    return content;
}

[[nodiscard]] float MeasureConflictPromptContentHeight(IDWriteFactory* factory,
                                                       IDWriteTextFormat* headingFormat,
                                                       IDWriteTextFormat* bodyFormat,
                                                       IDWriteTextFormat* smallFormat,
                                                       const ConflictPromptContent& content,
                                                       float maxWidth,
                                                       float lineHeight,
                                                       float rowGap,
                                                       float inspectionButtonHeight,
                                                       bool comparisonOnly = false) noexcept
{
    float height          = 0.0f;
    const auto addWrapped = [&](std::wstring_view text, IDWriteTextFormat* format) noexcept
    {
        if (text.empty())
        {
            return;
        }
        if (height > 0.0f)
        {
            height += rowGap;
        }
        height += MeasureWrappedTextHeight(factory, format, text, maxWidth, lineHeight);
    };
    const auto addContext =
        [&](std::wstring_view label, std::wstring_view path, std::wstring_view name, std::wstring_view parent, std::wstring_view metadata) noexcept
    {
        if (label.empty() || path.empty())
        {
            return;
        }
        if (height > 0.0f)
        {
            height += rowGap;
        }
        height += lineHeight;
        if (! name.empty() && ! parent.empty())
        {
            if (content.commonName.empty())
                height += MeasureWrappedTextHeight(factory, headingFormat, name, maxWidth, lineHeight) + rowGap;
            height += MeasureWrappedTextHeight(factory, smallFormat, parent, maxWidth, lineHeight);
        }
        else
            height += MeasureWrappedTextHeight(factory, smallFormat, path, maxWidth, lineHeight);
        if (! metadata.empty())
            height += rowGap + MeasureWrappedTextHeight(factory, smallFormat, metadata, maxWidth, lineHeight);
        height += rowGap + inspectionButtonHeight;
    };

    if (! comparisonOnly)
    {
        addWrapped(content.state, bodyFormat);
        addWrapped(content.commonName, headingFormat);
        addWrapped(content.question, bodyFormat);
        addWrapped(content.facts, smallFormat);
    }
    addContext(content.sourceLabel, content.sourcePath, content.sourceName, content.sourceParentPath, content.sourceMetadata);
    addContext(content.destinationLabel, content.destinationPath, content.destinationName, content.destinationParentPath, content.destinationMetadata);
    addWrapped(content.lastNote, smallFormat);
    if (! comparisonOnly)
        addWrapped(content.inspectionResult, smallFormat);
    return height;
}

[[nodiscard]] ConflictAction ConflictActionFromRaw(uint8_t rawAction) noexcept
{
    return static_cast<ConflictAction>(rawAction);
}

[[nodiscard]] uint8_t RawConflictAction(ConflictAction action) noexcept
{
    return static_cast<uint8_t>(action);
}

void CopyPublishedConflictPolicy(const FolderWindow::FileOperationState::Task::ConflictPromptState& source,
                                 FileOperationsPopupInternal::TaskSnapshot::ConflictPromptSnapshot& destination) noexcept
{
    const auto copyActions = [](const auto& sourceActions, size_t sourceCount, auto& destinationActions, size_t& destinationCount) noexcept
    {
        destinationCount = std::min(sourceCount, destinationActions.size());
        for (size_t index = 0u; index < destinationCount; ++index)
        {
            destinationActions[index] = RawConflictAction(sourceActions[index]);
        }
    };

    destination.generation = source.generation;
    copyActions(source.actions, source.actionCount, destination.actions, destination.actionCount);
    copyActions(source.primaryActions, source.primaryActionCount, destination.primaryActions, destination.primaryActionCount);
    copyActions(source.overflowActions, source.overflowActionCount, destination.overflowActions, destination.overflowActionCount);
    destination.defaultAction           = RawConflictAction(source.defaultAction);
    destination.escapeAction            = RawConflictAction(source.escapeAction);
    destination.applyToAllEligible      = source.applyToAllEligible;
    destination.skipAllEligible         = source.skipAllEligible;
    destination.repeatUnavailableReason = static_cast<uint8_t>(source.repeatUnavailableReason);
    destination.sourceIsDirectory  = source.sourceMetadata.objectKindKnown ? source.sourceMetadata.isDirectory
                                                                           : source.decisionScope.sourceKind == FileOperations::ConflictItemKind::Directory;
    destination.buttonsPublishable = source.buttonsPublishable;
}

[[nodiscard]] std::wstring PopupActionAutomationId(const FileOperationsPopupInternal::PopupHitTest& hit)
{
    auto id = std::format(L"FileOperations.Action.{}.{}.{}", static_cast<uint32_t>(hit.kind), hit.taskId, hit.data);
    if (hit.promptGeneration != 0u)
        id += std::format(L".{}", hit.promptGeneration);
    return id;
}
[[nodiscard]] std::wstring ConflictActionText(ConflictAction action,
                                              ConflictBucket bucket         = ConflictBucket::Unknown,
                                              FileSystemOperation operation = FILESYSTEM_COPY) noexcept
{
    return FileOperations::FormatConflictActionLabel(action, bucket, operation);
}

D2D1_RECT_F ComputeIndeterminateBarFill(const D2D1_RECT_F& bar, ULONGLONG tick, bool reducedMotion) noexcept
{
    const float width = bar.right - bar.left;
    if (width <= 0.0f)
    {
        return bar;
    }

    constexpr ULONGLONG kPeriodMs = 1200ull;
    const float segmentW          = width * 0.28f;

    if (reducedMotion)
    {
        const float left = bar.left + (width - segmentW) * 0.5f;
        return D2D1::RectF(left, bar.top, left + segmentW, bar.bottom);
    }

    const ULONGLONG phaseMs = tick % kPeriodMs;
    const float t           = static_cast<float>(phaseMs) / static_cast<float>(kPeriodMs);

    const float travel = width + segmentW;
    const float x      = bar.left + travel * t - segmentW;

    const float left  = std::clamp(x, bar.left, bar.right);
    const float right = std::clamp(x + segmentW, bar.left, bar.right);
    return D2D1::RectF(left, bar.top, right, bar.bottom);
}

float ClampCornerRadius(const D2D1_RECT_F& rc, float desired) noexcept
{
    const float w         = std::max(0.0f, rc.right - rc.left);
    const float h         = std::max(0.0f, rc.bottom - rc.top);
    const float maxRadius = std::min(w, h) * 0.5f;
    return std::clamp(desired, 0.0f, maxRadius);
}

std::wstring FormatDurationHms(uint64_t seconds)
{
    const uint64_t hours64   = seconds / 3600u;
    const uint64_t minutes64 = (seconds % 3600u) / 60u;
    const uint64_t seconds64 = seconds % 60u;

    const unsigned long long hours = static_cast<unsigned long long>(hours64);
    const unsigned int minutes     = static_cast<unsigned int>(minutes64);
    const unsigned int secs        = static_cast<unsigned int>(seconds64);

    if (hours > 0ull)
    {
        return std::format(L"{}:{:02d}:{:02d}", hours, minutes, secs);
    }

    return std::format(L"{:02d}:{:02d}", minutes, secs);
}

using TaskStatusKind        = FileOperationsPopupInternal::TaskSnapshot::StatusKind;
using PopupStatusVisualTone = FileOperationsPopupInternal::PopupStatusVisualTone;

[[nodiscard]] double ClampFiniteNonNegative(double value) noexcept;

[[nodiscard]] uint64_t SaturatingRoundNonNegativeToUint64(double value) noexcept
{
    value = ClampFiniteNonNegative(value);
    if (value <= 0.0)
    {
        return 0;
    }

    const double maxValue = static_cast<double>(std::numeric_limits<uint64_t>::max());
    if (value >= maxValue)
    {
        return std::numeric_limits<uint64_t>::max();
    }

    return static_cast<uint64_t>(value + 0.5);
}

[[nodiscard]] uint64_t SaturatingCeilNonNegativeToUint64(double value) noexcept
{
    value = ClampFiniteNonNegative(value);
    if (value <= 0.0)
    {
        return 0;
    }

    const double maxValue = static_cast<double>(std::numeric_limits<uint64_t>::max());
    if (value >= maxValue)
    {
        return std::numeric_limits<uint64_t>::max();
    }

    return static_cast<uint64_t>(std::ceil(value));
}

[[nodiscard]] bool IsByteRateUsableForEta(double bytesPerSecond) noexcept
{
    constexpr double kMinimumEtaRateBytesPerSecond = 1.0;
    return ClampFiniteNonNegative(bytesPerSecond) >= kMinimumEtaRateBytesPerSecond;
}

struct GlobalFileOperationsStatusSummary
{
    uint32_t running              = 0;
    uint32_t stopping             = 0;
    uint32_t waiting              = 0;
    uint32_t needAttention        = 0;
    uint32_t activeRunning        = 0;
    uint32_t paused               = 0;
    uint32_t pauseEligibleRunning = 0;
    uint32_t pauseEligiblePaused  = 0;
    uint64_t completedBytes       = 0;
    uint64_t totalBytes           = 0;
    uint64_t completedItems       = 0;
    uint64_t totalItems           = 0;
    // Progress weights are separate from transferred-byte counters. A verified copy
    // contributes its equally weighted transfer/proof fraction, using its byte total
    // as the cohort weight. Double accumulation avoids uint64 saturation bias here.
    double completedByteWork           = 0.0;
    double totalByteWork               = 0.0;
    bool includesVerificationWork      = false;
    bool hasUnknownActiveProgress      = false;
    bool hasUnknownActiveTransferBytes = false;
    bool hasActiveTaskbarProgress      = false;
    uint32_t openDiscoveryTasks        = 0u;
    uint32_t visibleCountingTasks      = 0u;
    // A closed task with no total at all: nothing is known about it and it will not close later,
    // so it keeps the aggregate indeterminate (unlike an open task, which only adds to the count).
    bool hasClosedUnknownTotals = false;
    double displayedBytesPerSec = 0.0;
    bool hasAggregateEta        = false;
    double aggregateEtaSeconds  = 0.0;
};

struct GlobalTaskbarProgressModel
{
    uint32_t state     = static_cast<uint32_t>(TBPF_NOPROGRESS);
    uint64_t completed = 0;
    uint64_t total     = 0;
};

[[nodiscard]] bool TaskHasKnownCompactProgress(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (task.operation == FILESYSTEM_COPY && ! task.finished)
        return task.discoveryClosed && task.discoveredBytesKnown &&
               (task.totalBytes > 0u || (! task.verificationRequested && task.copiedFileCountKnown && task.discoveredFileCount > 0u));
    return task.discoveryClosed && (task.totalBytes > 0 || task.totalItems > 0);
}

struct WholeTaskProgressPresentation final
{
    double fraction               = 0.0;
    bool determinate              = false;
    bool provisional              = false;
    bool transferActive           = false;
    bool showWorkFacts            = false;
    bool showDiscoveryActivity    = false;
    bool showTransferProgress     = false;
    bool showTransferMarquee      = false;
    bool showThroughputGraph      = false;
    bool includeInAggregateCohort = false;
};

[[nodiscard]] WholeTaskProgressPresentation ResolveWholeTaskProgressPresentation(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    WholeTaskProgressPresentation result{};
    // operationStartTick is published when ExecuteOperation is underway. It makes the startup
    // marquee visible before the first transfer callback without mistaking discovery-only work
    // for an active transfer.
    result.transferActive          = task.started && task.operationStartTick != 0u;
    result.provisional             = ! task.discoveryClosed;
    const bool waitingForAdmission = task.queuePaused || task.waitingInQueue || task.waitingForOthers;
    // A waiting task has no current work to describe, even if its scan/stream
    // snapshot still contains facts from before a queue pause.
    result.showWorkFacts            = ! task.finished && ! waitingForAdmission && ! task.conflict.active;
    result.includeInAggregateCohort = task.started && ! task.finished && (task.cancellationRequested || (! waitingForAdmission && ! task.conflict.active));
    result.showDiscoveryActivity =
        result.includeInAggregateCohort && ! task.cancellationRequested && ! task.discoveryClosed && task.discoveryAheadActive && ! task.discoverySkipped;
    result.showTransferProgress = result.includeInAggregateCohort;
    result.showThroughputGraph =
        result.includeInAggregateCohort && result.transferActive && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE);

    // D2-A08: no growing-denominator fraction. Until the total closes the bar is indeterminate (a
    // marquee once transfer is active) and the card shows exact counters instead.
    if (task.discoveryClosed)
    {
        if (task.totalBytes > 0 && task.completedBytes <= task.totalBytes && (task.operation != FILESYSTEM_COPY || task.finished || task.discoveredBytesKnown))
        {
            result.fraction    = task.verificationRequested && task.operation != FILESYSTEM_DELETE
                                     ? static_cast<double>(ComputeFileOperationsTaskCompleteFractionForDisplay(task))
                                     : std::clamp(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes), 0.0, 1.0);
            result.determinate = true;
            return result;
        }
        if (task.operation == FILESYSTEM_COPY && ! task.finished)
        {
            if (task.discoveredBytesKnown && task.totalBytes == 0u && ! task.verificationRequested && task.copiedFileCountKnown &&
                task.discoveredFileCount > 0u && task.copiedFileCount <= task.discoveredFileCount)
            {
                result.fraction    = static_cast<double>(task.copiedFileCount) / static_cast<double>(task.discoveredFileCount);
                result.determinate = true;
            }
        }
        else if (task.totalItems > 0 && task.completedItems <= task.totalItems)
        {
            result.fraction    = std::clamp(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems), 0.0, 1.0);
            result.determinate = true;
        }
    }
    result.showTransferMarquee = result.showTransferProgress && ! result.determinate && result.transferActive && ! task.cancellationRequested;
    return result;
}

[[nodiscard]] double VerificationTransferProgressPercent(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.totalBytes > 0
               ? std::clamp(static_cast<double>((std::min)(task.completedBytes, task.totalBytes)) / static_cast<double>(task.totalBytes), 0.0, 1.0) * 100.0
               : 0.0;
}

[[nodiscard]] double VerificationProofProgressPercent(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    const uint64_t total = task.verificationTotalBytes > 0 ? task.verificationTotalBytes : task.totalBytes;
    return total > 0 ? std::clamp(static_cast<double>((std::min)(task.verificationCompletedBytes, total)) / static_cast<double>(total), 0.0, 1.0) * 100.0 : 0.0;
}

// The two halves of a Verify On bar. A finished, successful task has finished its transfer; its verification half
// shows the share verified (I35, owner report 2026-10-13).
[[nodiscard]] double TransferSegmentPercent(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.finished && SUCCEEDED(task.resultHr) ? 100.0 : VerificationTransferProgressPercent(task);
}

[[nodiscard]] double VerificationSegmentPercent(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.finished && SUCCEEDED(task.resultHr) ? FinishedVerificationPercent(task) : VerificationProofProgressPercent(task);
}

// The compact meter's label of a finished Verify On task that did not verify every byte, in place of its percent:
// "Not verified", or the share verified ("Verified 40%"). Empty otherwise (the percent is shown).
[[nodiscard]] std::wstring FinishedVerificationLabel(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    if (! FinishedTaskVerificationIncomplete(task))
        return {};
    const auto percent = static_cast<unsigned int>(std::floor(FinishedVerificationPercent(task)));
    return percent == 0u ? LoadStringResource(nullptr, IDS_FILEOPS_COMPACT_NOT_VERIFIED)
                         : FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_VERIFIED_PERCENT, percent);
}

void PublishPlannedItemTotalAfterDiscovery(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (task.discoveryClosed && task.totalItems == 0 && task.operation != FILESYSTEM_DELETE)
    {
        task.totalItems = task.plannedItems;
    }
}

[[nodiscard]] bool TaskShowsPreparingStatus(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (task.kind != FileOperationsPopupInternal::TaskSnapshot::Kind::FileOperation)
    {
        return false;
    }
    const auto phase = static_cast<FileOperations::TaskLifecyclePhase>(task.lifecyclePhase);
    return phase == FileOperations::TaskLifecyclePhase::Preparing || phase == FileOperations::TaskLifecyclePhase::AwaitingAcceptance;
}

[[nodiscard]] bool TaskHasSkippedOutcome(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.resultHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) || task.skippedItemCount > 0u;
}

// A Partial with no skipped root and no counted problem is a Move that kept sources (for example a
// disclosed Copy-only Move): its caption states that fact, never "with warnings" next to zero warnings.
[[nodiscard]] bool TaskPartialIsKeptSourcesOnly(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept;

[[nodiscard]] bool TaskHasPartialOutcome(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    const unsigned long downgradedRetained = task.retainedSourceCount - std::min(task.canceledRetainedSourceCount, task.retainedSourceCount);
    return TaskHasSkippedOutcome(task) || (task.operation == FILESYSTEM_MOVE && downgradedRetained > 0u);
}

[[nodiscard]] bool TaskPartialIsKeptSourcesOnly(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return ! TaskHasSkippedOutcome(task) && task.warningCount == 0u && task.errorCount == 0u && TaskHasPartialOutcome(task);
}

[[nodiscard]] TaskStatusKind ResolveTaskStatusKind(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    const HRESULT partialHr   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);

    if (task.finished)
    {
        if (task.resultHr == cancelledHr || task.resultHr == E_ABORT)
        {
            // A canceled item may keep an unknown publication or source axis: that cancel is not a
            // clean Canceled, and its chip, badge and caption agree with "Outcome unknown".
            return task.hasIndeterminateResult ? TaskStatusKind::Indeterminate : TaskStatusKind::Canceled;
        }
        if (task.hasIndeterminateResult || task.resultHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
        {
            return TaskStatusKind::Indeterminate;
        }
        if (task.resultHr == partialHr || (SUCCEEDED(task.resultHr) && (task.warningCount > 0 || task.errorCount > 0 || TaskHasPartialOutcome(task))))
        {
            return TaskStatusKind::Partial;
        }
        if (FAILED(task.resultHr))
        {
            return TaskStatusKind::Failed;
        }
        return TaskStatusKind::Done;
    }

    if (task.cancellationRequested)
    {
        return TaskStatusKind::Stopping;
    }
    if (task.conflict.active)
    {
        return TaskStatusKind::Conflict;
    }
    if (task.queuePaused || task.waitingInQueue || task.waitingForOthers)
    {
        return TaskStatusKind::Waiting;
    }
    if (task.paused)
    {
        return TaskStatusKind::Paused;
    }
    if (task.verificationActive)
    {
        return TaskStatusKind::Verifying;
    }
    // 5F early admission: discovery now runs concurrently with the transfer. The blocking
    // "Discovering" status is only truthful before the transfer has actually started; once
    // ExecuteOperation is underway (operationStartTick set) show the live transfer status, whose
    // ETA already reads "estimating" until discovery totals settle.
    if (task.discoveryAheadActive && task.operationStartTick == 0)
    {
        return TaskStatusKind::Discovering;
    }
    if (TaskShowsPreparingStatus(task))
    {
        return TaskStatusKind::Preparing;
    }
    return TaskStatusKind::Running;
}

// Only task decisions are withdrawn on cancel; path inspection, disclosure and
// diagnostic/history affordances remain available.
[[nodiscard]] bool IsTaskDecisionHit(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
{
    using Kind = FileOperationsPopupInternal::PopupHitTest::Kind;
    switch (kind)
    {
        case Kind::TaskPause:
        case Kind::TaskCancel:
        case Kind::TaskStartNow:
        case Kind::TaskQueueMoveUp:
        case Kind::TaskQueueMoveDown:
        case Kind::TaskSkip:
        case Kind::TaskSpeedLimit:
        case Kind::TaskConflictToggleApplyToAll:
        case Kind::TaskConflictMore:
        case Kind::TaskConflictAction: return true;
        case Kind::None:
        case Kind::FooterCancelAll:
        case Kind::FooterPauseResumeAll:
        case Kind::FooterAutoDismiss:
        case Kind::FooterQueueMode:
        case Kind::FooterDensity:
        case Kind::FooterToggleDetails:
        case Kind::CompletedGroupToggle:
        case Kind::CompletedGroupClear:
        case Kind::TaskToggleCollapse:
        case Kind::TaskShowLog:
        case Kind::TaskExportIssues:
        case Kind::TaskCompletedMore:
        case Kind::TaskDismiss:
        case Kind::FooterOptions:
        case Kind::TaskToggleThroughput:
        case Kind::TaskConflictInspectSource:
        case Kind::TaskConflictInspectDestination:
        case Kind::TaskIssues: return false;
    }
    return false;
}

// Popup-only projection: retain the worker's facts and history, but withdraw
// pending decisions as soon as durable cancellation authority is observed.
void ProjectTaskCancellationForDisplay(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (task.cancellationRequested && ! task.finished)
    {
        task.conflict         = {};
        task.canMoveQueueUp   = false;
        task.canMoveQueueDown = false;
    }
    task.statusKind = ResolveTaskStatusKind(task);
}

void EnsureFinishedTaskDiagnosticAffordance(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (! task.finished || task.interruptedMoveNotice || task.warningCount > 0 || task.errorCount > 0)
    {
        return;
    }

    const HRESULT partialHr   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    if (task.resultHr == partialHr)
    {
        // I35 D-5.2: a partial the user's skips explain keeps its Partial status and note without a
        // synthesized warning.
        if (task.partialExplainedByUserSkips)
        {
            return;
        }
        task.warningCount = 1;
        if (task.lastDiagnosticMessage.empty())
        {
            task.lastDiagnosticMessage = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_SKIPPED_ITEMS);
        }
        return;
    }

    if (FAILED(task.resultHr) && task.resultHr != cancelledHr && task.resultHr != E_ABORT)
    {
        task.errorCount = 1;
        if (task.lastDiagnosticMessage.empty())
        {
            task.lastDiagnosticMessage = FileOperations::FormatConflictProblemText(FileOperations::ConflictClass::Unknown, task.resultHr);
        }
    }
}

[[nodiscard]] bool StatusNeedsAttention(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Conflict || status == TaskStatusKind::Partial || status == TaskStatusKind::Indeterminate ||
           status == TaskStatusKind::Failed;
}

[[nodiscard]] bool StatusCountsAsRunning(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Discovering || status == TaskStatusKind::Preparing || status == TaskStatusKind::Running ||
           status == TaskStatusKind::Verifying || status == TaskStatusKind::Paused || status == TaskStatusKind::Stopping;
}

[[nodiscard]] bool StatusCountsAsWaiting(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Waiting;
}

[[nodiscard]] bool StatusIsWarning(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Conflict || status == TaskStatusKind::Partial || status == TaskStatusKind::Indeterminate;
}

[[nodiscard]] bool StatusIsError(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Failed;
}

[[nodiscard]] bool StatusIsOk(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Done;
}

[[nodiscard]] PopupStatusVisualTone StatusVisualToneForTaskStatus(TaskStatusKind status) noexcept
{
    switch (status)
    {
        case TaskStatusKind::Running:
        case TaskStatusKind::Verifying:
        case TaskStatusKind::Discovering:
        case TaskStatusKind::Preparing: return PopupStatusVisualTone::Accent;
        case TaskStatusKind::Waiting:
        case TaskStatusKind::Stopping:
        case TaskStatusKind::Canceled: return PopupStatusVisualTone::Neutral;
        case TaskStatusKind::Paused: return PopupStatusVisualTone::Muted;
        case TaskStatusKind::Conflict:
        case TaskStatusKind::Partial:
        case TaskStatusKind::Indeterminate: return PopupStatusVisualTone::Warning;
        case TaskStatusKind::Failed: return PopupStatusVisualTone::Error;
        case TaskStatusKind::Done: return PopupStatusVisualTone::Ok;
        case TaskStatusKind::None:
        default: break;
    }

    return PopupStatusVisualTone::None;
}

[[nodiscard]] D2D1::ColorF StatusVisualColorForTone(const AppTheme& theme, PopupStatusVisualTone tone) noexcept
{
    switch (tone)
    {
        case PopupStatusVisualTone::Accent: return theme.accent;
        case PopupStatusVisualTone::Muted:
        case PopupStatusVisualTone::Neutral: return ColorFromCOLORREF(theme.menu.disabledText);
        case PopupStatusVisualTone::Warning: return theme.folderView.warningText;
        case PopupStatusVisualTone::Error: return theme.folderView.errorText;
        case PopupStatusVisualTone::Ok: return theme.fileOperations.successText;
        case PopupStatusVisualTone::None:
        default: break;
    }

    return ColorFromCOLORREF(theme.menu.text);
}

[[nodiscard]] std::wstring StatusTextForTask(const FileOperationsPopupInternal::TaskSnapshot& task, TaskStatusKind status, ULONGLONG nowTick)
{
    if (task.finished && ! task.resultSummary.empty() &&
        (status == TaskStatusKind::Done || status == TaskStatusKind::Partial || status == TaskStatusKind::Indeterminate || status == TaskStatusKind::Failed))
    {
        return task.resultSummary;
    }
    switch (status)
    {
        case TaskStatusKind::Waiting: return LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_WAITING);
        case TaskStatusKind::Paused: return LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_PAUSED);
        case TaskStatusKind::Conflict:
            return LoadStringResource(nullptr, task.conflict.metadataLoading ? IDS_FILEOPS_CONFLICT_READING_DETAILS : IDS_FILEOPS_STATUS_WAITING_FOR_DECISION);
        case TaskStatusKind::Discovering:
        {
            const uint64_t elapsedSec = task.discoveryElapsedMs / 1000;
            return elapsedSec > 0 ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_DISCOVERING_TIME, FormatDurationHms(elapsedSec))
                                  : LoadStringResource(nullptr, IDS_FILEOPS_DISCOVERING);
        }
        case TaskStatusKind::Preparing:
        {
            const ULONGLONG opStartTick = task.operationStartTick;
            const uint64_t elapsedSec   = (opStartTick > 0 && nowTick >= opStartTick) ? static_cast<uint64_t>((nowTick - opStartTick) / 1000ull) : 0ull;
            return elapsedSec > 0 ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_PREPARING_TIME, FormatDurationHms(elapsedSec))
                                  : LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
        }
        case TaskStatusKind::Verifying: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_VERIFYING);
        case TaskStatusKind::Stopping: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_STOPPING);
        case TaskStatusKind::Done: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
        case TaskStatusKind::Partial:
            return LoadStringResource(nullptr,
                                      TaskHasSkippedOutcome(task)          ? static_cast<UINT>(IDS_FILEOPS_STATUS_PARTIAL)
                                      : TaskPartialIsKeptSourcesOnly(task) ? static_cast<UINT>(IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT)
                                                                           : static_cast<UINT>(IDS_FILEOPS_STATUS_COMPLETED_WITH_WARNINGS));
        case TaskStatusKind::Indeterminate: return LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
        // P5: the failure's HRESULT stays in Technical details, the log and the export.
        case TaskStatusKind::Failed: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_FAILED_SHORT);
        case TaskStatusKind::Canceled: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_CANCELED);
        case TaskStatusKind::Running:
        case TaskStatusKind::None:
        default: break;
    }

    return {};
}

[[nodiscard]] std::wstring StatusChipTextForTask(const FileOperationsPopupInternal::TaskSnapshot& task, TaskStatusKind status, ULONGLONG nowTick)
{
    switch (status)
    {
        case TaskStatusKind::Running: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_RUNNING);
        case TaskStatusKind::Verifying: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_VERIFYING);
        case TaskStatusKind::Waiting: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_WAITING);
        case TaskStatusKind::Done: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
        case TaskStatusKind::Partial:
            return LoadStringResource(nullptr,
                                      TaskHasSkippedOutcome(task) || TaskPartialIsKeptSourcesOnly(task) ? IDS_FILEOPS_STATUS_PARTIAL_SHORT
                                                                                                        : IDS_FILEOPS_STATUS_WARNINGS_SHORT);
        case TaskStatusKind::Indeterminate: return LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
        case TaskStatusKind::Failed: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_FAILED_SHORT);
        // P5: the chip's short label; the card's status line says "Counting items…".
        case TaskStatusKind::Discovering: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COUNTING_SHORT);
        case TaskStatusKind::Preparing: return LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
        case TaskStatusKind::Paused:
        case TaskStatusKind::Stopping:
        case TaskStatusKind::Canceled: return StatusTextForTask(task, status, nowTick);
        case TaskStatusKind::Conflict: return {};
        case TaskStatusKind::None:
        default: break;
    }

    return {};
}

struct TerminalAxisBadge final
{
    std::wstring text;
    PopupStatusVisualTone tone = PopupStatusVisualTone::None;
    std::wstring_view automationSuffix;
};

[[nodiscard]] bool TaskHasVerificationResultBadge(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (! task.finished)
    {
        return false;
    }

    switch (static_cast<FileOperations::VerificationState>(task.verificationState))
    {
        case FileOperations::VerificationState::Verified:
        case FileOperations::VerificationState::Failed:
        case FileOperations::VerificationState::Unavailable:
        case FileOperations::VerificationState::Canceled: return true;
        case FileOperations::VerificationState::NotRequested:
        case FileOperations::VerificationState::NotApplicable:
        default: break;
    }
    return false;
}

// Move sources kept as a downgrade: a source a cancel or stop kept before publishing was never copied.
[[nodiscard]] unsigned long DowngradedRetainedSourceCount(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.retainedSourceCount - std::min(task.canceledRetainedSourceCount, task.retainedSourceCount);
}

[[nodiscard]] bool TaskHasTerminalAxisBadges(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return TaskHasVerificationResultBadge(task) ||
           (task.finished && task.operation == FILESYSTEM_MOVE && (DowngradedRetainedSourceCount(task) > 0u || task.unknownSourceCount > 0u));
}

[[nodiscard]] size_t BuildTerminalAxisBadges(const FileOperationsPopupInternal::TaskSnapshot& task, std::array<TerminalAxisBadge, 3u>& badges)
{
    if (! task.finished)
    {
        return 0u;
    }

    size_t count = 0u;
    switch (static_cast<FileOperations::VerificationState>(task.verificationState))
    {
        case FileOperations::VerificationState::Verified:
            badges[count++] = TerminalAxisBadge{LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFIED), PopupStatusVisualTone::Ok, L"Verification"};
            break;
        case FileOperations::VerificationState::Failed:
            badges[count++] =
                TerminalAxisBadge{LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFICATION_FAILED), PopupStatusVisualTone::Error, L"Verification"};
            break;
        case FileOperations::VerificationState::Unavailable:
            badges[count++] =
                TerminalAxisBadge{LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFICATION_UNAVAILABLE), PopupStatusVisualTone::Warning, L"Verification"};
            break;
        case FileOperations::VerificationState::Canceled:
            badges[count++] =
                TerminalAxisBadge{LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFICATION_CANCELED), PopupStatusVisualTone::Neutral, L"Verification"};
            break;
        case FileOperations::VerificationState::NotRequested:
        case FileOperations::VerificationState::NotApplicable:
        default: break;
    }

    if (task.operation == FILESYSTEM_MOVE && task.unknownSourceCount > 0u)
    {
        badges[count++] = TerminalAxisBadge{LoadStringResource(nullptr, IDS_FILEOPS_RESULT_SOURCE_UNKNOWN), PopupStatusVisualTone::Error, L"SourceUnknown"};
    }
    const unsigned long downgradedRetainedCount = DowngradedRetainedSourceCount(task);
    if (task.operation == FILESYSTEM_MOVE && downgradedRetainedCount > 0u)
    {
        const UINT wording =
            task.retainedSourceNativeCount == downgradedRetainedCount ? IDS_FILEOPS_RESULT_MOVED_SOURCE_FOLDER_KEPT : IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT;
        badges[count++] = TerminalAxisBadge{LoadStringResource(nullptr, wording), PopupStatusVisualTone::Warning, L"SourceRetained"};
    }
    return count;
}

[[nodiscard]] std::array<unsigned long, 10> TerminalFactCounts(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return {{task.publishedItemCount,
             task.unknownPublicationItemCount,
             task.removedSourceCount,
             task.retainedSourceCount,
             task.unknownSourceCount,
             task.missingSourceCount,
             task.verifiedItemCount,
             task.verificationFailedItemCount,
             task.verificationUnavailableItemCount,
             task.verificationCanceledItemCount}};
}

// I35 P5/R5.8/R9.3: a finished card's counters line exists only for a nonzero counter.
[[nodiscard]] bool TaskShowsCountersLine(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.warningCount > 0u || task.errorCount > 0u || (task.showInformationalNotes && task.noteCount > 0u);
}

// I35 R5.3: the card's Last note is the last Warning/Error problem; with informational notes shown it
// falls back to the last Information note when no problem exists.
[[nodiscard]] const std::wstring& EffectiveLastNote(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.lastDiagnosticMessage.empty() && task.showInformationalNotes ? task.lastInfoNoteMessage : task.lastDiagnosticMessage;
}

[[nodiscard]] std::wstring FormatTaskCountersText(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    std::wstring text;
    const auto append = [&text](std::wstring part)
    {
        if (! text.empty())
            text += LoadStringResource(nullptr, IDS_FILEOPS_GLOBAL_STATUS_SEPARATOR);
        text += std::move(part);
    };
    if (task.warningCount > 0u)
        append(FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_WARNINGS, task.warningCount));
    if (task.errorCount > 0u)
        append(FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ERRORS, task.errorCount));
    if (task.showInformationalNotes && task.noteCount > 0u)
        append(FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_NOTES, task.noteCount));
    return text;
}

// I35 string inventory (P5): the card states facts the user can act on. A Copy keeps its sources by
// definition, so "left in place" is not a fact for it; a Move whose every published item also
// removed its source, with nothing retained or unknown, says only how many items moved. Unknown,
// retained and failed axes are never hidden.
[[nodiscard]] bool MoveRemovedEveryPublishedSource(const std::array<unsigned long, 10>& counts) noexcept
{
    return counts[2] == counts[0] && counts[1] == 0u && counts[3] == 0u && counts[4] == 0u && counts[5] == 0u;
}

[[nodiscard]] std::array<unsigned long, 10> VisibleTerminalFactCounts(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    auto counts = TerminalFactCounts(task);
    if (task.operation == FILESYSTEM_COPY)
        counts[3] = 0u;
    if (task.operation == FILESYSTEM_MOVE && MoveRemovedEveryPublishedSource(counts))
        counts[2] = 0u;
    return counts;
}

// I35 R9.4 (P3): a task canceled before any selected root crossed the mutation gate states one fact,
// its result line "Not started; nothing was changed", instead of its untouched sources.
[[nodiscard]] bool TaskCanceledBeforeMutation(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.finished && ! task.startedMutation && ResolveTaskStatusKind(task) == TaskStatusKind::Canceled;
}

[[nodiscard]] bool TaskHasTerminalFacts(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (! task.finished || task.interruptedMoveNotice || TaskCanceledBeforeMutation(task))
        return false;
    const auto counts = VisibleTerminalFactCounts(task);
    return std::ranges::any_of(counts, [](unsigned long count) noexcept { return count > 0u; }) || TaskHasTerminalAxisBadges(task);
}

[[nodiscard]] std::wstring BuildTerminalFactsText(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    const auto counts = VisibleTerminalFactCounts(task);
    // A Move says "moved" only when every published item also removed its source; a Move that
    // left sources in place or has an unknown outcome states the neutral "completed" count next
    // to its Source facts, so the card never claims a move that did not happen.
    const bool everyItemMoved    = task.operation == FILESYSTEM_MOVE && MoveRemovedEveryPublishedSource(TerminalFactCounts(task));
    const UINT completedClauseId = task.operation == FILESYSTEM_COPY ? static_cast<UINT>(IDS_FMT_FILEOPS_TERMINAL_COPIED)
                                   : everyItemMoved                  ? static_cast<UINT>(IDS_FMT_FILEOPS_TERMINAL_MOVED)
                                                                     : static_cast<UINT>(IDS_FMT_FILEOPS_TERMINAL_PUBLISHED);
    const std::array<UINT, 10> ids{{completedClauseId,
                                    IDS_FMT_FILEOPS_TERMINAL_PUBLICATION_UNKNOWN,
                                    IDS_FMT_FILEOPS_TERMINAL_SOURCE_REMOVED,
                                    IDS_FMT_FILEOPS_TERMINAL_SOURCE_RETAINED,
                                    IDS_FMT_FILEOPS_TERMINAL_SOURCE_UNKNOWN,
                                    IDS_FMT_FILEOPS_TERMINAL_SOURCE_MISSING,
                                    IDS_FMT_FILEOPS_TERMINAL_VERIFIED,
                                    IDS_FMT_FILEOPS_TERMINAL_VERIFY_FAILED,
                                    IDS_FMT_FILEOPS_TERMINAL_VERIFY_UNAVAILABLE,
                                    IDS_FMT_FILEOPS_TERMINAL_VERIFY_CANCELED}};
    constexpr std::array<size_t, 4> offsets{{0u, 2u, 6u, 10u}};
    constexpr std::array<UINT, 3> labels{
        {IDS_FILEOPS_TERMINAL_AXIS_RESULT, IDS_ITEM_PROPERTIES_ARTIFACT_FIELD_SOURCE, IDS_ITEM_PROPERTIES_ARTIFACT_FIELD_VERIFICATION}};
    std::wstring result;
    for (size_t axis = 0u; axis < labels.size(); ++axis)
    {
        std::wstring clauses;
        for (size_t index = offsets[axis]; index < offsets[axis + 1u]; ++index)
        {
            if (counts[index] == 0u)
                continue;
            const std::wstring clause = FormatStringResource(nullptr, ids[index], counts[index]);
            clauses                   = clauses.empty() ? clause : FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_FACT_JOIN, clauses, clause);
        }
        if (axis == 2u && clauses.empty() && TaskHasVerificationResultBadge(task))
        {
            // Historical/synthetic snapshots may retain a categorical verification
            // result without a numeric receipt. Preserve that fact without inventing a count.
            std::array<TerminalAxisBadge, 3u> legacy{};
            if (BuildTerminalAxisBadges(task, legacy) > 0u)
                clauses = legacy.front().text;
        }
        if (clauses.empty())
            continue;
        if (! result.empty())
            result += L"\n";
        const std::wstring label = LoadStringResource(nullptr, labels[axis]);
        result += FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_FACT_AXIS, label, clauses);
    }
    return result;
}

[[nodiscard]] std::wstring BuildTaskHeaderText(const FileOperationsPopupInternal::TaskSnapshot& task, std::wstring_view operationText, ULONGLONG nowTick)
{
    if (task.selection.itemCount > 0u)
    {
        const auto& selected       = task.selection;
        const std::wstring subject = selected.itemCount == 1u && ! selected.singleName.empty() ? selected.singleName
                                     : selected.kindsKnown
                                         ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_SELECTION_KINDS, selected.fileCount, selected.folderCount)
                                         : FormatStringResource(nullptr, IDS_FMT_FILEOPS_SELECTION_ITEMS, selected.itemCount);
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_OPERATION_SUBJECT, std::wstring(operationText), subject);
    }
    // The status chip names the terminal state; the expanded body carries the
    // exact result. Keep the operation heading concise instead of clipping that result.
    if (task.finished && ! task.resultSummary.empty())
        return std::wstring(operationText);
    // After Skip counting, only the completed-work summary in the body carries counts.
    // Selected roots are not a denominator for the files and folders inside them.
    if (TaskUsesSkippedCountingWorkSummary(task))
        return std::wstring(operationText);
    const TaskStatusKind status = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
    if (status == TaskStatusKind::Running || status == TaskStatusKind::Verifying || status == TaskStatusKind::Stopping || status == TaskStatusKind::None)
    {
        if (task.discoveryClosed && task.totalItems > 0)
        {
            return FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_OP_COUNTS, std::wstring(operationText), task.completedItems, task.totalItems);
        }
        return FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_OP_COUNTS_UNKNOWN_TOTAL, std::wstring(operationText), task.completedItems);
    }

    return FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, std::wstring(operationText), StatusTextForTask(task, status, nowTick));
}

[[nodiscard]] std::wstring GraphOverlayTextForStatus(const FileOperationsPopupInternal::TaskSnapshot& task, TaskStatusKind status, bool& showAnimation)
{
    showAnimation = false;
    switch (status)
    {
        case TaskStatusKind::Waiting: return LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_WAITING);
        case TaskStatusKind::Paused: return LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_PAUSED);
        case TaskStatusKind::Discovering: return {};
        case TaskStatusKind::Preparing: showAnimation = true; return LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
        case TaskStatusKind::Verifying: showAnimation = true; return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_VERIFYING);
        case TaskStatusKind::Stopping: return {};
        case TaskStatusKind::Conflict:
            return LoadStringResource(nullptr, task.conflict.metadataLoading ? IDS_FILEOPS_CONFLICT_READING_DETAILS : IDS_FILEOPS_STATUS_WAITING_FOR_DECISION);
        case TaskStatusKind::Running:
        case TaskStatusKind::Done:
        case TaskStatusKind::Partial:
        case TaskStatusKind::Indeterminate:
        case TaskStatusKind::Failed:
        case TaskStatusKind::Canceled:
        case TaskStatusKind::None:
        default: break;
    }

    static_cast<void>(task);
    return {};
}

[[nodiscard]] GlobalFileOperationsStatusSummary BuildGlobalStatusSummary(
    const std::vector<FileOperationsPopupInternal::TaskSnapshot>& snapshot,
    const std::unordered_map<uint64_t, FileOperationsPopupInternal::RateHistory>* rates = nullptr) noexcept
{
    GlobalFileOperationsStatusSummary summary{};
    double aggregateBytesPerSec                = 0.0;
    double aggregateTransferBytesPerSec        = 0.0;
    double aggregateVerificationBytesPerSec    = 0.0;
    double aggregateRemainingTransferBytes     = 0.0;
    double aggregateRemainingVerificationBytes = 0.0;
    bool etaInputsComplete                     = true;
    for (const auto& task : snapshot)
    {
        if (task.kind != FileOperationsPopupInternal::TaskSnapshot::Kind::FileOperation)
        {
            continue;
        }

        const TaskStatusKind status =
            task.finished || task.cancellationRequested || task.statusKind == TaskStatusKind::None ? ResolveTaskStatusKind(task) : task.statusKind;
        const WholeTaskProgressPresentation presentation = ResolveWholeTaskProgressPresentation(task);
        if (task.finished)
        {
            continue;
        }

        if (StatusNeedsAttention(status))
        {
            ++summary.needAttention;
        }
        else if (StatusCountsAsWaiting(status))
        {
            ++summary.waiting;
        }
        else if (StatusCountsAsRunning(status))
        {
            ++summary.running;
            if (status == TaskStatusKind::Stopping)
                ++summary.stopping;
            if (status == TaskStatusKind::Paused)
            {
                ++summary.paused;
                if (task.started)
                {
                    ++summary.pauseEligiblePaused;
                }
            }
            else
            {
                ++summary.activeRunning;
                if (task.started && status != TaskStatusKind::Stopping)
                {
                    ++summary.pauseEligibleRunning;
                }
            }
        }

        if (presentation.includeInAggregateCohort)
        {
            summary.hasActiveTaskbarProgress = true;
        }

        const auto accumulateSaturated = [](uint64_t value, uint64_t& aggregate) noexcept
        {
            const uint64_t remaining = (std::numeric_limits<uint64_t>::max)() - aggregate;
            aggregate                = value > remaining ? (std::numeric_limits<uint64_t>::max)() : aggregate + value;
        };

        const bool copySizeUnavailable = task.operation == FILESYSTEM_COPY && ! task.finished && ! task.discoveredBytesKnown;
        if (presentation.includeInAggregateCohort && task.discoveryClosed && task.totalBytes > 0 && ! copySizeUnavailable)
        {
            accumulateSaturated(task.totalBytes, summary.totalBytes);
            accumulateSaturated((std::min)(task.completedBytes, task.totalBytes), summary.completedBytes);
            const bool includesVerification = task.verificationRequested && task.operation != FILESYSTEM_DELETE;
            const double weight             = static_cast<double>(task.totalBytes);
            const double fraction           = includesVerification ? static_cast<double>(ComputeFileOperationsTaskCompleteFractionForDisplay(task))
                                                                   : static_cast<double>((std::min)(task.completedBytes, task.totalBytes)) / weight;
            summary.totalByteWork += weight;
            summary.completedByteWork += weight * fraction;
            summary.includesVerificationWork = summary.includesVerificationWork || includesVerification;
        }
        if (presentation.includeInAggregateCohort && task.discoveryClosed && task.operation == FILESYSTEM_COPY && ! task.finished)
        {
            if (task.discoveredBytesKnown && task.totalBytes == 0u && ! task.verificationRequested && task.copiedFileCountKnown &&
                task.discoveredFileCount > 0u)
            {
                accumulateSaturated(task.discoveredFileCount, summary.totalItems);
                accumulateSaturated((std::min)(task.copiedFileCount, static_cast<uint64_t>(task.discoveredFileCount)), summary.completedItems);
            }
        }
        else if (presentation.includeInAggregateCohort && task.discoveryClosed && task.totalItems > 0)
        {
            accumulateSaturated(task.totalItems, summary.totalItems);
            accumulateSaturated((std::min)(task.completedItems, task.totalItems), summary.completedItems);
        }
        if (presentation.includeInAggregateCohort && ! presentation.determinate)
        {
            summary.hasUnknownActiveProgress = true;
        }
        if (presentation.includeInAggregateCohort && ! task.discoveryClosed && status != TaskStatusKind::Stopping)
        {
            ++summary.openDiscoveryTasks;
            if (! task.discoverySkipped)
                ++summary.visibleCountingTasks;
        }
        if (presentation.includeInAggregateCohort && task.discoveryClosed && ! presentation.determinate)
        {
            summary.hasClosedUnknownTotals = true;
        }
        if (presentation.includeInAggregateCohort && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) &&
            (! task.discoveryClosed || task.totalBytes == 0 || copySizeUnavailable))
        {
            summary.hasUnknownActiveTransferBytes = true;
        }

        if (! rates || ! presentation.includeInAggregateCohort || (task.operation != FILESYSTEM_COPY && task.operation != FILESYSTEM_MOVE))
        {
            continue;
        }

        const auto rateIt = rates->find(task.taskId);
        if (rateIt == rates->end())
        {
            if (task.discoveryClosed && task.totalBytes > task.completedBytes)
            {
                etaInputsComplete = false;
            }
            continue;
        }

        const double transferBytesPerSec     = ClampFiniteNonNegative(rateIt->second.displayedBytesPerSec);
        const double verificationBytesPerSec = ClampFiniteNonNegative(rateIt->second.displayedVerificationBytesPerSec);
        const double bytesPerSec             = transferBytesPerSec + verificationBytesPerSec;
        aggregateBytesPerSec += bytesPerSec;
        if (task.discoveryClosed && task.totalBytes > 0 && task.completedBytes < task.totalBytes)
        {
            if (! IsByteRateUsableForEta(transferBytesPerSec))
            {
                etaInputsComplete = false;
                continue;
            }
            aggregateRemainingTransferBytes += static_cast<double>(task.totalBytes - task.completedBytes);
            aggregateTransferBytesPerSec += transferBytesPerSec;
        }
        if (task.verificationRequested && task.verificationTotalBytes > task.verificationCompletedBytes)
        {
            if (! IsByteRateUsableForEta(verificationBytesPerSec))
            {
                etaInputsComplete = false;
                continue;
            }
            aggregateRemainingVerificationBytes += static_cast<double>(task.verificationTotalBytes - task.verificationCompletedBytes);
            aggregateVerificationBytesPerSec += verificationBytesPerSec;
        }
    }

    summary.completedBytes = (std::min)(summary.completedBytes, summary.totalBytes);
    summary.completedItems = (std::min)(summary.completedItems, summary.totalItems);

    if (rates)
    {
        summary.displayedBytesPerSec        = aggregateBytesPerSec;
        const bool hasTransferRemainder     = aggregateRemainingTransferBytes > 0.0;
        const bool hasVerificationRemainder = aggregateRemainingVerificationBytes > 0.0;
        const bool transferEtaAvailable     = ! hasTransferRemainder || IsByteRateUsableForEta(aggregateTransferBytesPerSec);
        const bool verificationEtaAvailable = ! hasVerificationRemainder || IsByteRateUsableForEta(aggregateVerificationBytesPerSec);
        if (summary.stopping == 0u && ! summary.hasUnknownActiveTransferBytes && etaInputsComplete && transferEtaAvailable && verificationEtaAvailable &&
            (hasTransferRemainder || hasVerificationRemainder))
        {
            summary.hasAggregateEta     = true;
            summary.aggregateEtaSeconds = (hasTransferRemainder ? aggregateRemainingTransferBytes / aggregateTransferBytesPerSec : 0.0) +
                                          (hasVerificationRemainder ? aggregateRemainingVerificationBytes / aggregateVerificationBytesPerSec : 0.0);
        }
    }

    return summary;
}

[[nodiscard]] bool HasGlobalAggregateProgress(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    return summary.hasActiveTaskbarProgress;
}

// FO-DISCOVERY-01: the closed-total cohort is Known work and stays determinate while other tasks
// still discover; the taskbar model keeps its own indeterminate rule for open totals.
[[nodiscard]] bool HasDeterminateGlobalAggregateProgress(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    return ! summary.hasClosedUnknownTotals && (summary.totalBytes > 0 || summary.totalItems > 0);
}

[[nodiscard]] bool HasFooterPauseResumeAllControl(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    return summary.pauseEligibleRunning > 0 || summary.pauseEligiblePaused > 0;
}

[[nodiscard]] bool FooterPauseResumeAllShouldPause(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    return summary.pauseEligibleRunning > 0;
}

// An interrupted-Move restart notice stays a standalone card: it is never hidden behind a
// group badge, and only its own Dismiss may acknowledge its durable breadcrumb.
[[nodiscard]] bool IsCompletedGroupTask(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.kind == FileOperationsPopupInternal::TaskSnapshot::Kind::FileOperation && task.finished && ! task.interruptedMoveNotice;
}

struct CompletedGroupResultSummary final
{
    uint32_t total         = 0u;
    uint32_t completed     = 0u;
    uint32_t partial       = 0u;
    uint32_t indeterminate = 0u;
    uint32_t failed        = 0u;
    uint32_t canceled      = 0u;
};

[[nodiscard]] CompletedGroupResultSummary BuildCompletedGroupResultSummary(const std::vector<FileOperationsPopupInternal::TaskSnapshot>& snapshot) noexcept
{
    CompletedGroupResultSummary summary{};
    for (const auto& task : snapshot)
    {
        if (! IsCompletedGroupTask(task))
        {
            continue;
        }

        ++summary.total;
        const TaskStatusKind status = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
        switch (status)
        {
            case TaskStatusKind::Done: ++summary.completed; break;
            case TaskStatusKind::Partial: ++summary.partial; break;
            case TaskStatusKind::Indeterminate: ++summary.indeterminate; break;
            case TaskStatusKind::Failed: ++summary.failed; break;
            case TaskStatusKind::Canceled: ++summary.canceled; break;
            case TaskStatusKind::None:
            case TaskStatusKind::Waiting:
            case TaskStatusKind::Discovering:
            case TaskStatusKind::Preparing:
            case TaskStatusKind::Running:
            case TaskStatusKind::Verifying:
            case TaskStatusKind::Paused:
            case TaskStatusKind::Stopping:
            case TaskStatusKind::Conflict: break;
        }
    }
    return summary;
}

[[nodiscard]] bool ShouldShowCompletedGroup(uint32_t completedCount) noexcept
{
    return completedCount >= 2u;
}

[[nodiscard]] float GlobalAggregateProgressFraction(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    if (summary.includesVerificationWork && summary.totalByteWork > 0.0)
    {
        return Clamp01(static_cast<float>(summary.completedByteWork / summary.totalByteWork));
    }
    if (summary.totalBytes > 0)
    {
        return Clamp01(static_cast<float>(static_cast<double>(summary.completedBytes) / static_cast<double>(summary.totalBytes)));
    }
    if (summary.totalItems > 0)
    {
        return Clamp01(static_cast<float>(static_cast<double>(summary.completedItems) / static_cast<double>(summary.totalItems)));
    }
    return 0.0f;
}

[[nodiscard]] std::wstring FormatGlobalStatusSummaryText(const GlobalFileOperationsStatusSummary& summary)
{
    std::vector<std::wstring> statusParts;
    statusParts.reserve(7u);
    // x-ux#7: paused and stopping tasks have their own clauses; Running counts only the rest,
    // so the footer agrees with the taskbar's paused state.
    if (summary.running > summary.stopping + summary.paused)
    {
        statusParts.push_back(FormatStringResource(
            nullptr, IDS_FMT_FILEOPS_GLOBAL_RUNNING_COUNT, static_cast<unsigned long>(summary.running - summary.stopping - summary.paused)));
    }
    if (summary.paused > 0u)
    {
        statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_PAUSED_COUNT, static_cast<unsigned long>(summary.paused)));
    }
    if (summary.stopping > 0u)
    {
        statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_STOPPING_COUNT, static_cast<unsigned long>(summary.stopping)));
    }
    if (summary.waiting > 0u)
    {
        statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_WAITING_COUNT, static_cast<unsigned long>(summary.waiting)));
    }
    if (summary.needAttention > 0u)
    {
        statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_ATTENTION_COUNT, static_cast<unsigned long>(summary.needAttention)));
    }
    if (summary.visibleCountingTasks > 0u)
    {
        if (HasDeterminateGlobalAggregateProgress(summary))
        {
            const unsigned long knownWorkPercent =
                static_cast<unsigned long>(std::clamp(std::lround(static_cast<double>(GlobalAggregateProgressFraction(summary)) * 100.0), 0L, 100L));
            statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_KNOWN_WORK, knownWorkPercent));
        }
        statusParts.push_back(
            FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_DISCOVERING_COUNT, static_cast<unsigned long>(summary.visibleCountingTasks)));
    }

    const std::wstring separator = LoadStringResource(nullptr, IDS_FILEOPS_GLOBAL_STATUS_SEPARATOR);
    std::wstring statusText;
    for (const auto& part : statusParts)
    {
        if (! statusText.empty())
        {
            statusText.append(separator);
        }
        statusText.append(part);
    }
    if (summary.displayedBytesPerSec <= 0.0)
    {
        return statusText;
    }

    const uint64_t roundedBytesPerSecond = SaturatingRoundNonNegativeToUint64(summary.displayedBytesPerSec);
    const std::wstring speedText         = FormatBytesCompact(roundedBytesPerSecond);
    if (! summary.hasAggregateEta || summary.aggregateEtaSeconds <= 0.0)
    {
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_STATUS_WITH_SPEED, statusText, speedText);
    }

    const uint64_t seconds     = SaturatingCeilNonNegativeToUint64(summary.aggregateEtaSeconds);
    const std::wstring etaText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA, FormatDurationHms(seconds));
    return FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_STATUS_WITH_ETA_SPEED, statusText, etaText, speedText);
}

[[nodiscard]] GlobalTaskbarProgressModel BuildGlobalTaskbarProgressModel(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    GlobalTaskbarProgressModel model{};
    if (summary.needAttention > 0)
    {
        // Needs-attention is a taskbar state, not progress evidence. Publish no
        // value while keeping the blocked decision visible outside the popup.
        model.state = static_cast<uint32_t>(TBPF_ERROR);
        return model;
    }

    if (! summary.hasActiveTaskbarProgress)
    {
        return model;
    }

    if (! summary.hasUnknownActiveProgress && summary.totalBytes > 0)
    {
        if (summary.includesVerificationWork)
        {
            // ITaskbarList3 accepts work units. Do not overwrite the real byte
            // counters or claim that verification transfers additional bytes.
            model.total     = 1'000'000u;
            model.completed = static_cast<uint64_t>(std::clamp(summary.completedByteWork / summary.totalByteWork, 0.0, 1.0) * static_cast<double>(model.total));
        }
        else
        {
            model.completed = summary.completedBytes;
            model.total     = summary.totalBytes;
        }
    }
    else if (! summary.hasUnknownActiveProgress && summary.totalItems > 0)
    {
        model.completed = summary.completedItems;
        model.total     = summary.totalItems;
    }

    if (summary.activeRunning == 0 && (summary.waiting > 0 || summary.paused > 0))
    {
        model.state = static_cast<uint32_t>(TBPF_PAUSED);
        return model;
    }

    model.state = static_cast<uint32_t>(model.total > 0 ? TBPF_NORMAL : TBPF_INDETERMINATE);
    return model;
}

#ifdef ENABLE_TESTS
[[nodiscard]] size_t SurfacedTaskStatusCount(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::None ? 0u : 1u;
}
#endif

class FileOperationsSpeedLimitPromptWindow final
{
public:
    FileOperationsSpeedLimitPromptWindow(HWND ownerWindow, const AppTheme& theme, uint64_t initialLimitBytesPerSecond) noexcept
        : _ownerWindow(ownerWindow && IsWindow(ownerWindow) != FALSE ? ownerWindow : nullptr),
          _theme(theme),
          _initialLimitBytesPerSecond(initialLimitBytesPerSecond)
    {
    }

    FileOperationsSpeedLimitPromptWindow(const FileOperationsSpeedLimitPromptWindow&)            = delete;
    FileOperationsSpeedLimitPromptWindow& operator=(const FileOperationsSpeedLimitPromptWindow&) = delete;
    FileOperationsSpeedLimitPromptWindow(FileOperationsSpeedLimitPromptWindow&&)                 = delete;
    FileOperationsSpeedLimitPromptWindow& operator=(FileOperationsSpeedLimitPromptWindow&&)      = delete;

    [[nodiscard]] std::optional<uint64_t> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, DipsToPixels(440, dpi), DipsToPixels(212, dpi)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        // Creation and initial sizing must use the same monitor. Creating at CW_USEDEFAULT
        // can attach the host at another DPI before owner centering scales it a second time.
        int initialX = CW_USEDEFAULT;
        int initialY = CW_USEDEFAULT;
        MONITORINFO ownerMonitor{};
        ownerMonitor.cbSize = sizeof(ownerMonitor);
        if (_ownerWindow && GetMonitorInfoW(MonitorFromWindow(_ownerWindow, MONITOR_DEFAULTTONEAREST), &ownerMonitor))
        {
            initialX = ownerMonitor.rcWork.left;
            initialY = ownerMonitor.rcWork.top;
        }

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFileOperationsSpeedLimitPromptClassName,
                                          LoadStringResource(nullptr, IDS_CAPTION_FILEOP_SPEED_LIMIT_CUSTOM).c_str(),
                                          style,
                                          initialX,
                                          initialY,
                                          bounds.right - bounds.left,
                                          bounds.bottom - bounds.top,
                                          _ownerWindow,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          this);
        if (! hwnd)
        {
            return std::nullopt;
        }

        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwnerWorkArea(_hWnd.get(), _ownerWindow));
        }

        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        Common::RequestForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done && ! HostIsPromptShutdown())
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                _result.reset();
                _done = true;
                break;
            }
            if (getMessageResult == 0)
            {
                _result.reset();
                _done = true;
                PostQuitMessage(static_cast<int>(msg.wParam));
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
            if (HostIsPromptShutdown())
            {
                _result.reset();
                _done = true;
            }
        }

        if (HostIsPromptShutdown())
        {
            _result.reset();
        }

        return _result;
    }

    // Settings reload already posts WM_THEMECHANGED to the owning popup. Forward
    // that owned notification without timers, rebuilding controls or storing a borrowed theme.
    static void RefreshOwnedPrompts(HWND owner, const AppTheme& theme) noexcept
    {
        struct RefreshContext
        {
            HWND owner;
            const AppTheme* theme;
        } context{owner, &theme};
        EnumThreadWindows(GetCurrentThreadId(),
                          [](HWND candidate, LPARAM value) noexcept -> BOOL
        {
            const auto& refresh = *reinterpret_cast<const RefreshContext*>(value);
            if (GetWindow(candidate, GW_OWNER) != refresh.owner)
                return TRUE;
            std::array<wchar_t, 128u> name{};
            if (GetClassNameW(candidate, name.data(), static_cast<int>(name.size())) == 0 ||
                std::wstring_view(name.data()) != kFileOperationsSpeedLimitPromptClassName)
                return TRUE;
            auto* prompt = reinterpret_cast<FileOperationsSpeedLimitPromptWindow*>(GetWindowLongPtrW(candidate, GWLP_USERDATA));
            if (prompt && prompt->_hWnd.get() == candidate && ! prompt->_done)
            {
                prompt->_theme = *refresh.theme;
                prompt->RefreshPresentation();
            }
            return TRUE;
        },
                          reinterpret_cast<LPARAM>(&context));
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        auto* self = reinterpret_cast<FileOperationsSpeedLimitPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (message == WM_NCCREATE)
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            self               = create ? reinterpret_cast<FileOperationsSpeedLimitPromptWindow*>(create->lpCreateParams) : nullptr;
            if (! self)
            {
                return FALSE;
            }

            self->_hWnd.reset(hwnd);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }

        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        return self->WindowProc(hwnd, message, wParam, lParam);
    }

#ifdef ENABLE_TESTS
public:
    enum class DebugCommand : uintptr_t
    {
        GetSnapshot,
        SetText,
        Confirm,
        Cancel,
        SetClientHeightLimit,
        SetSelection,
    };
#endif

private:
    // Application-specific help navigation uses the shared retained Label and
    // button focus chrome; no separate scrolling or paint ownership is introduced.
    class HelpLabel final : public DxUi::Label
    {
    public:
        using DxUi::Label::Label;
        HelpLabel(const HelpLabel&)            = delete;
        HelpLabel(HelpLabel&&)                 = delete;
        HelpLabel& operator=(const HelpLabel&) = delete;
        HelpLabel& operator=(HelpLabel&&)      = delete;

        void Paint(DxUi::WindowHost& host) const override
        {
            DxUi::Label::Paint(host);
            if (! HasFocus())
                return;
            DxUi::ButtonChromeDrawSpec spec{};
            spec.bounds = GetBounds();
            DxUi::ButtonChromeCustomStyle style{};
            style.showFill   = false;
            style.showFocus  = true;
            style.focus      = host.GetTheme().focusStrokeOuter;
            style.focusRing  = DxUi::ButtonChromeFocusRing::Standard;
            spec.customStyle = style;
            DxUi::DrawButtonChrome(host.GetDeviceContext(), host.GetSolidBrush(style.focus), nullptr, nullptr, host.GetTheme(), spec);
        }
    };

    enum class DeferredAction : WPARAM
    {
        Confirm = 1,
        Cancel  = 2,
    };

    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = FileOperationsSpeedLimitPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFileOperationsSpeedLimitPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        if (atom != 0)
        {
            return S_OK;
        }

        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            atom = 1;
            return S_OK;
        }

        return HRESULT_FROM_WIN32(lastError);
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        ApplyTheme();
        Layout();
        _dxHost.SetOnFocusChanged([this](DxUi::Control* control) { EnsureBodyControlVisible(control); });
        _dxHost.SetFocusControl(_field);
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (_root != nullptr)
        {
            return;
        }

        using namespace DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _body = _root->AddChild<ScrollPanel>();
        _body->SetScrollStepDip(32.0f);
        _body->SetOnScrollChanged([this](float)
        {
            if (! _layingOut)
            {
                _dxHost.SyncTextInput(_field);
                _dxHost.RefreshAccessibilitySnapshot();
            }
        });

        _label = _body->AddChild<Label>(LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_LABEL));
        _label->SetMultiline(true);
        _label->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _field = _body->AddChild<TextField>(_initialLimitBytesPerSecond == 0 ? L"0" : FormatBytesCompact(_initialLimitBytesPerSecond));
        _field->SetOnTextChanged([this](std::wstring_view) { ClearValidation(); });

        _hintLabel = _body->AddChild<HelpLabel>(LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_HINT));
        _hintLabel->SetMultiline(true);
        _hintLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _hintLabel->SetFontRole(FontRole::Small);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { RequestDeferredAction(DeferredAction::Confirm); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { RequestDeferredAction(DeferredAction::Cancel); });

        _field->SetAccessibleName(std::wstring(_label->GetText()));
        _field->SetAccessibleAutomationId(L"FileOps.SpeedLimit.Edit");
        _hintLabel->SetAccessibleAutomationId(L"FileOps.SpeedLimit.Help");
        _languageRevision = Localization::GetLanguagePreferenceRevision();
        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        UpdateHintVisuals();
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
        }
    }

    void RefreshPresentation() noexcept
    {
        const uint64_t revision = Localization::GetLanguagePreferenceRevision();
        if (_languageRevision != revision && _root)
        {
            // The host routes WM_SETTEXT to its focused native edit session. This
            // update belongs to the window chrome and must not replace the speed text.
            const auto caption = LoadStringResource(nullptr, IDS_CAPTION_FILEOP_SPEED_LIMIT_CUSTOM);
            DefWindowProcW(_hWnd.get(), WM_SETTEXT, 0, reinterpret_cast<LPARAM>(caption.c_str()));
            _label->SetText(LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_LABEL));
            _field->SetAccessibleName(std::wstring(_label->GetText()));
            _okButton->SetText(LoadStringResource(nullptr, IDS_BTN_OK));
            _cancelButton->SetText(LoadStringResource(nullptr, IDS_BTN_CANCEL));
            if (_validationResourceId != 0u)
                _validationText = LoadStringResource(nullptr, _validationResourceId);
            _languageRevision = revision;
        }
        ApplyTheme();
        Layout();
        _dxHost.Invalidate();
    }

    void EnsureBodyControlVisible(DxUi::Control* control) noexcept
    {
        if (! _body || (control != _field && control != _hintLabel))
            return;
        const auto viewport = _body->GetViewportRect();
        const auto bounds   = control->GetBounds();
        const float offset  = _body->GetScrollOffset();
        const float top     = bounds.top - offset - 2.0f;
        const float bottom  = bounds.bottom - offset + 2.0f;
        // A long help paragraph starts at its top and is then keyboard/wheel scrollable.
        const float delta = top < viewport.top || bounds.bottom - bounds.top + 4.0f > viewport.bottom - viewport.top
                                ? top - viewport.top
                                : (bottom > viewport.bottom ? bottom - viewport.bottom : 0.0f);
        _body->SetScrollOffset(offset + delta);
    }

    [[nodiscard]] bool OnBodyKeyDown(UINT key) noexcept
    {
        // Help is an ordinary semantic Label; only its focused short-form view owns
        // these navigation keys. Editing, default Enter and Escape keep their host routes.
        if (! _body || ! _body->NeedsScrollbar() || _dxHost.GetFocusControl() != _hintLabel || (GetKeyState(VK_CONTROL) & 0x8000) != 0 ||
            (GetKeyState(VK_MENU) & 0x8000) != 0)
            return false;
        const auto viewport = _body->GetViewportRect();
        const float page    = std::max(24.0f, viewport.bottom - viewport.top - 24.0f);
        float offset        = _body->GetScrollOffset();
        switch (key)
        {
            case VK_UP: offset -= 24.0f; break;
            case VK_DOWN: offset += 24.0f; break;
            case VK_PRIOR: offset -= page; break;
            case VK_NEXT: offset += page; break;
            case VK_HOME: offset = 0.0f; break;
            case VK_END: offset = _body->GetContentHeight(); break;
            default: return false;
        }
        _body->SetScrollOffset(offset);
        return true;
    }

    void Layout() noexcept
    {
        if (! _root || IsIconic(_hWnd.get()))
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        if (client.right <= client.left || client.bottom <= client.top)
            return;
        _root->SetBounds(client);

        constexpr float kMarginDip       = 16.0f;
        constexpr float kGapDip          = 10.0f;
        constexpr float kFieldHeightDip  = 32.0f;
        constexpr float kButtonHeightDip = 34.0f;
        constexpr float kButtonWidthDip  = 96.0f;

        const float left      = client.left + kMarginDip;
        const float right     = client.right - kMarginDip;
        const float textWidth = std::max(1.0f, right - left);
        float labelHeight =
            MeasureWrappedTextHeight(_dxHost.GetWriteFactory(), _dxHost.GetTextFormat(DxUi::FontRole::Body), _label->GetText(), textWidth, 24.0f);
        float hintHeight =
            MeasureWrappedTextHeight(_dxHost.GetWriteFactory(), _dxHost.GetTextFormat(DxUi::FontRole::Small), _hintLabel->GetText(), textWidth, 20.0f);
        const float requiredHeight = kMarginDip * 2.0f + labelHeight + 4.0f + kFieldHeightDip + kGapDip * 2.0f + hintHeight + kButtonHeightDip;
        if (! _resizingForLayout && requiredHeight > client.bottom - client.top + 0.5f)
        {
            const UINT dpi = GetDpiForWindow(_hWnd.get());
            RECT needed{0, 0, 0, static_cast<LONG>(std::ceil(DipsToPixels(requiredHeight, dpi)))};
            RECT window{};
            if (GetWindowRect(_hWnd.get(), &window) && AdjustWindowRectExForDpi(&needed,
                                                                                static_cast<DWORD>(GetWindowLongPtrW(_hWnd.get(), GWL_STYLE)),
                                                                                FALSE,
                                                                                static_cast<DWORD>(GetWindowLongPtrW(_hWnd.get(), GWL_EXSTYLE)),
                                                                                dpi))
            {
                LONG targetHeight = needed.bottom - needed.top;
                LONG targetTop    = window.top;
                MONITORINFO monitor{};
                monitor.cbSize = sizeof(monitor);
                if (GetMonitorInfoW(MonitorFromWindow(_hWnd.get(), MONITOR_DEFAULTTONEAREST), &monitor))
                {
                    targetHeight = std::min(targetHeight, monitor.rcWork.bottom - monitor.rcWork.top);
                    targetTop    = std::clamp(targetTop, monitor.rcWork.top, monitor.rcWork.bottom - targetHeight);
                }
#ifdef ENABLE_TESTS
                if (_debugClientHeightLimitDip.has_value())
                {
                    const LONG nonclient = window.bottom - window.top - static_cast<LONG>(std::ceil(DipsToPixels(client.bottom - client.top, dpi)));
                    targetHeight = std::min(targetHeight, nonclient + static_cast<LONG>(std::ceil(DipsToPixels(_debugClientHeightLimitDip.value(), dpi))));
                }
#endif
                if (targetHeight > window.bottom - window.top)
                {
                    _resizingForLayout     = true;
                    const auto resetResize = wil::scope_exit([&] { _resizingForLayout = false; });
                    if (SetWindowPos(_hWnd.get(), nullptr, window.left, targetTop, window.right - window.left, targetHeight, SWP_NOZORDER | SWP_NOACTIVATE))
                        return; // Synchronous WM_SIZE arranged the new client bounds.
                }
            }
        }
        _layingOut              = true;
        const auto finishLayout = wil::scope_exit([&]
        {
            _layingOut = false;
            EnsureBodyControlVisible(_dxHost.GetFocusControl());
            _dxHost.SyncTextInput(_field);
            _dxHost.RefreshAccessibilitySnapshot();
        });
        const float buttonsTop  = std::max(client.top, client.bottom - kMarginDip - kButtonHeightDip);
        const float bodyTop     = client.top + kMarginDip;
        const float bodyBottom  = std::max(bodyTop, buttonsTop - kGapDip);
        _body->SetBounds(D2D1::RectF(left, bodyTop, right, bodyBottom));
        float contentRight = right;
        float bodyHeight   = labelHeight + 4.0f + kFieldHeightDip + kGapDip + hintHeight;
        if (bodyHeight > bodyBottom - bodyTop)
        {
            contentRight = std::max(left + 1.0f, right - _body->GetScrollbarThickness());
            labelHeight =
                MeasureWrappedTextHeight(_dxHost.GetWriteFactory(), _dxHost.GetTextFormat(DxUi::FontRole::Body), _label->GetText(), contentRight - left, 24.0f);
            hintHeight = MeasureWrappedTextHeight(
                _dxHost.GetWriteFactory(), _dxHost.GetTextFormat(DxUi::FontRole::Small), _hintLabel->GetText(), contentRight - left, 20.0f);
            bodyHeight = labelHeight + 4.0f + kFieldHeightDip + kGapDip + hintHeight;
        }
        _body->SetContentHeight(bodyHeight);
        _hintLabel->SetFocusable(_body->NeedsScrollbar());
        if (! _body->NeedsScrollbar() && _dxHost.GetFocusControl() == _hintLabel)
            _dxHost.SetFocusControl(_field, false);
        float y = bodyTop;

        _label->SetBounds(D2D1::RectF(left, y, contentRight, y + labelHeight));
        y += labelHeight + 4.0f;

        _field->SetBounds(D2D1::RectF(left, y, contentRight, y + kFieldHeightDip));
        y += kFieldHeightDip + kGapDip;

        _hintLabel->SetBounds(D2D1::RectF(left, y, contentRight, y + hintHeight));

        const float cancelLeft = right - kButtonWidthDip;
        const float okLeft     = cancelLeft - 8.0f - kButtonWidthDip;
        _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
    }

    void ClearValidation() noexcept
    {
        if (! _showingValidationError)
        {
            return;
        }

        _showingValidationError = false;
        _validationResourceId   = 0u;
        _validationText.clear();
        UpdateHintVisuals();
        Layout();
        _dxHost.Invalidate();
    }

    void ShowValidation(UINT messageId) noexcept
    {
        _showingValidationError = true;
        _validationResourceId   = messageId;
        _validationText         = LoadStringResource(nullptr, messageId);
        UpdateHintVisuals();
        Layout();
        _dxHost.Invalidate();
        MessageBeep(MB_ICONWARNING);
    }

    void UpdateHintVisuals() noexcept
    {
        if (! _hintLabel)
        {
            return;
        }

        _hintLabel->SetText(_showingValidationError ? _validationText : LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_HINT));
        _hintLabel->SetTextColor(_showingValidationError ? _palette.errorText : _palette.disabledText);
    }

    void RequestDeferredAction(DeferredAction action) noexcept
    {
        if (_done)
        {
            return;
        }

        if (_hWnd && IsWindow(_hWnd.get()) != FALSE &&
            PostMessageW(_hWnd.get(), kFileOperationsSpeedLimitPromptDeferredActionMessage, static_cast<WPARAM>(action), 0) != FALSE)
        {
            return;
        }

        RunDeferredAction(action);
    }

    void RunDeferredAction(DeferredAction action) noexcept
    {
        switch (action)
        {
            case DeferredAction::Confirm: Confirm(); break;
            case DeferredAction::Cancel: Cancel(); break;
        }
    }

    void Confirm() noexcept
    {
        ClearValidation();

        const std::wstring text = _field ? std::wstring(_field->GetText()) : std::wstring{};
        uint64_t parsed         = 0;
        if (! Common::Parsing::TryParseBinaryThroughputText(text, Common::Parsing::ThroughputBoundaryWhitespacePolicy::AsciiWhitespace, parsed))
        {
            ShowValidation(IDS_MSG_FILEOP_SPEED_LIMIT_INVALID);
            _dxHost.SetFocusControl(_field);
            return;
        }

        _result = parsed;
        _done   = true;
        ClosePromptWindow();
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        ClosePromptWindow();
    }

    void ClosePromptWindow() noexcept
    {
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _dxHost.SetFocusControl(nullptr);
            _hWnd.reset();
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(DebugCommand command, LPARAM payload) noexcept
    {
        switch (command)
        {
            case DebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<FileOperationsSpeedLimitPromptDebugSnapshot*>(payload);
                if (! snapshot)
                {
                    return FALSE;
                }

                snapshot->bodyBounds           = _body->GetViewportRect();
                snapshot->bodyScrollOffsetDip  = _body->GetScrollOffset();
                snapshot->bodyContentHeightDip = _body->GetContentHeight();
                snapshot->bodyScrollable       = _body->NeedsScrollbar();
                snapshot->fieldFocused         = _dxHost.GetFocusControl() == _field;
                snapshot->helpFocused          = _dxHost.GetFocusControl() == _hintLabel;
                snapshot->fieldIdentity        = reinterpret_cast<uintptr_t>(_field);
                snapshot->labelText            = _label->GetText();
                snapshot->cancelText           = _cancelButton->GetText();
                DxUi::TextInputState editState{};
                snapshot->hasEditState            = _dxHost.TryReadTextInputState(_field, editState);
                snapshot->caretIndex              = editState.caretIndex;
                snapshot->selectionAnchorIndex    = editState.selectionAnchorIndex;
                snapshot->usesDxUiHost            = _dxHost.GetRoot() == _root;
                snapshot->visibleChildWindowCount = 0u;
                if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
                {
                    EnumChildWindows(_hWnd.get(),
                                     [](HWND child, LPARAM cookie) noexcept -> BOOL
                    {
                        if (! IsActuallyVisibleChildWindow(child))
                        {
                            return TRUE;
                        }

                        auto* count = reinterpret_cast<size_t*>(cookie);
                        if (count)
                        {
                            *count += 1u;
                        }
                        return TRUE;
                    },
                                     reinterpret_cast<LPARAM>(&snapshot->visibleChildWindowCount));
                }
                snapshot->initialLimitBytesPerSecond = _initialLimitBytesPerSecond;
                snapshot->text                       = _field ? std::wstring(_field->GetText()) : std::wstring{};
                snapshot->hintText               = _showingValidationError ? std::wstring{} : LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_HINT);
                snapshot->validationText         = _validationText;
                snapshot->clientBounds           = _dxHost.GetClientBoundsDip();
                snapshot->labelBounds            = _label->GetBounds();
                snapshot->fieldBounds            = _field->GetBounds();
                snapshot->hintBounds             = _hintLabel->GetBounds();
                snapshot->okBounds               = _okButton->GetBounds();
                snapshot->cancelBounds           = _cancelButton->GetBounds();
                snapshot->requiredLabelHeightDip = MeasureWrappedTextHeight(_dxHost.GetWriteFactory(),
                                                                            _dxHost.GetTextFormat(DxUi::FontRole::Body),
                                                                            _label->GetText(),
                                                                            snapshot->labelBounds.right - snapshot->labelBounds.left,
                                                                            24.0f);
                snapshot->requiredHintHeightDip  = MeasureWrappedTextHeight(_dxHost.GetWriteFactory(),
                                                                            _dxHost.GetTextFormat(DxUi::FontRole::Small),
                                                                            _hintLabel->GetText(),
                                                                            snapshot->hintBounds.right - snapshot->hintBounds.left,
                                                                            20.0f);
                return TRUE;
            }
            case DebugCommand::SetText:
            {
                auto* text = reinterpret_cast<const std::wstring*>(payload);
                if (! text || ! _field)
                {
                    return FALSE;
                }

                _field->SetTextAndNotify(*text);
                ClearValidation();
                _dxHost.SetFocusControl(_field);
                _dxHost.SyncTextInput(_field);
                _dxHost.Invalidate();
                return TRUE;
            }
            case DebugCommand::SetClientHeightLimit:
            {
                const auto* limit = reinterpret_cast<const std::optional<float>*>(payload);
                if (! limit || (limit->has_value() && (! std::isfinite(limit->value()) || limit->value() < 100.0f)))
                    return FALSE;
                _debugClientHeightLimitDip = *limit;
                Layout();
                return TRUE;
            }
            case DebugCommand::SetSelection:
            {
                const auto* selection = reinterpret_cast<const std::pair<size_t, size_t>*>(payload);
                if (! selection || ! _field)
                    return FALSE;
                _field->SetSelectionRange(selection->first, selection->second);
                _dxHost.SyncTextInput(_field);
                return TRUE;
            }
            case DebugCommand::Confirm: RequestDeferredAction(DeferredAction::Confirm); return TRUE;
            case DebugCommand::Cancel: RequestDeferredAction(DeferredAction::Cancel); return TRUE;
        }

        return FALSE;
    }
#endif

    LRESULT WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_KEYDOWN && OnBodyKeyDown(static_cast<UINT>(wParam)))
            return 0;
        if (message == WM_THEMECHANGED || message == WM_SETTINGCHANGE)
        {
            RefreshPresentation();
            return 0;
        }
        bool dxHandled   = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = _dxHost.HandleMessage(hwnd, message, wParam, lParam, dxHandled);
        }
        if (dxHandled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                Layout();
            }
            if (message == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                _dxHost.Detach();
                if (_hWnd.get() == hwnd)
                {
                    static_cast<void>(_hWnd.release());
                }
                if (! _done)
                {
                    _result.reset();
                    _done = true;
                }
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: Layout(); return 0;
            case kFileOperationsSpeedLimitPromptDeferredActionMessage: RunDeferredAction(static_cast<DeferredAction>(wParam)); return 0;
            case WM_DPICHANGED:
            {
                if (const auto* suggested = reinterpret_cast<const RECT*>(lParam))
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                ApplyTheme();
                Layout();
                return 0;
            }
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, _theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_ERASEBKGND: return 1;
            case WM_CLOSE: RequestDeferredAction(DeferredAction::Cancel); return 0;
#ifdef ENABLE_TESTS
            case kFileOperationsSpeedLimitPromptDebugMessage: return OnDebugCommand(static_cast<DebugCommand>(wParam), lParam);
#endif
            case WM_NCDESTROY:
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                _dxHost.Detach();
                if (_hWnd.get() == hwnd)
                {
                    static_cast<void>(_hWnd.release());
                }
                if (! _done)
                {
                    _result.reset();
                    _done = true;
                }
                return 0;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    HWND _ownerWindow = nullptr;
    AppTheme _theme{};
    uint64_t _initialLimitBytesPerSecond = 0;
    wil::unique_hwnd _hWnd;
    DxUi::WindowHost _dxHost;
    std::unique_ptr<DxUi::Panel> _rootStorage;
    DxUi::Panel* _root = nullptr;
    DxUi::ThemePalette _palette{};
    DxUi::ScrollPanel* _body    = nullptr;
    DxUi::Label* _label         = nullptr;
    DxUi::TextField* _field     = nullptr;
    DxUi::Label* _hintLabel     = nullptr;
    DxUi::Button* _okButton     = nullptr;
    DxUi::Button* _cancelButton = nullptr;
    std::wstring _validationText;
    UINT _validationResourceId = 0u;
    uint64_t _languageRevision = 0u;
    bool _layingOut            = false;
#ifdef ENABLE_TESTS
    std::optional<float> _debugClientHeightLimitDip;
#endif
    bool _resizingForLayout      = false;
    bool _showingValidationError = false;
    bool _done                   = false;
    std::optional<uint64_t> _result;
};

[[nodiscard]] std::optional<uint64_t> ShowCustomSpeedLimitPrompt(HWND ownerWindow, const AppTheme& theme, uint64_t initialLimitBytesPerSecond) noexcept
{
    FileOperationsSpeedLimitPromptWindow window(ownerWindow, theme, initialLimitBytesPerSecond);
    return window.ShowModal();
}

bool PointInRectF(const D2D1_RECT_F& rc, float x, float y) noexcept
{
    return rc.left <= x && x <= rc.right && rc.top <= y && y <= rc.bottom;
}

float MeasureTextWidth(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text, float maxWidth, float height) noexcept
{
    if (! factory || ! format || text.empty())
    {
        return 0.0f;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT hr = factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, maxWidth, height, layout.addressof());
    if (FAILED(hr) || ! layout)
    {
        return 0.0f;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return 0.0f;
    }

    return metrics.width;
}

float MeasureWrappedTextHeight(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text, float maxWidth, float fallbackLineHeight) noexcept
{
    if (! factory || ! format || text.empty() || maxWidth <= 0.0f)
    {
        return 0.0f;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    const float maxHeight = std::max(fallbackLineHeight, 4096.0f);
    if (FAILED(factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, maxWidth, maxHeight, layout.addressof())) || ! layout ||
        FAILED(layout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP)))
    {
        return fallbackLineHeight;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return fallbackLineHeight;
    }
    return std::max(fallbackLineHeight, std::ceil(metrics.height));
}

float DrawWrappedText(IDWriteFactory* factory,
                      ID2D1RenderTarget* target,
                      IDWriteTextFormat* format,
                      ID2D1Brush* brush,
                      std::wstring_view text,
                      const D2D1_RECT_F& bounds,
                      float fallbackLineHeight) noexcept
{
    const float maxWidth = std::max(0.0f, bounds.right - bounds.left);
    const float height   = MeasureWrappedTextHeight(factory, format, text, maxWidth, fallbackLineHeight);
    if (! factory || ! target || ! format || ! brush || text.empty() || maxWidth <= 0.0f || height <= 0.0f)
    {
        return height;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, maxWidth, height, layout.addressof())) || ! layout ||
        FAILED(layout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP)))
    {
        target->DrawTextW(text.data(),
                          static_cast<UINT32>(text.size()),
                          format,
                          D2D1::RectF(bounds.left, bounds.top, bounds.right, bounds.top + fallbackLineHeight),
                          brush,
                          D2D1_DRAW_TEXT_OPTIONS_CLIP);
        return fallbackLineHeight;
    }

    target->DrawTextLayout(D2D1::Point2F(bounds.left, bounds.top), layout.get(), brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
    return height;
}

std::wstring TruncateTextMiddleToWidth(IDWriteFactory* factory,
                                       IDWriteTextFormat* format,
                                       std::wstring_view text,
                                       float maxWidth,
                                       float height,
                                       std::wstring_view ellipsisText,
                                       size_t fixedPrefixChars) noexcept
{
    if (! factory || ! format || text.empty() || ! std::isfinite(maxWidth) || ! std::isfinite(height) || maxWidth <= 0.0f || height <= 0.0f)
        return {};

    wil::com_ptr<IDWriteTextLayout> fullLayout;
    DWRITE_TEXT_METRICS fullMetrics{};
    if (FAILED(factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, maxWidth, height, fullLayout.addressof())) ||
        FAILED(fullLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP)) || FAILED(fullLayout->GetMetrics(&fullMetrics)))
        return {};
    if (fullMetrics.widthIncludingTrailingWhitespace <= maxWidth)
        return std::wstring(text);

    const float dotsWidth = MeasureTextWidth(factory, format, ellipsisText, maxWidth, height);
    if (dotsWidth <= 0.0f || dotsWidth > maxWidth)
        return {};

    // Use shaping-cluster boundaries so abbreviation cannot split a surrogate pair, combining
    // sequence or shaped emoji. These temporary arrays exist only for an overflowing current row.
    UINT32 clusterCount = 0u;
    static_cast<void>(fullLayout->GetClusterMetrics(nullptr, 0u, &clusterCount));
    if (clusterCount == 0u)
        return std::wstring(ellipsisText);
    std::vector<DWRITE_CLUSTER_METRICS> clusters(clusterCount);
    if (FAILED(fullLayout->GetClusterMetrics(clusters.data(), clusterCount, &clusterCount)))
        return std::wstring(ellipsisText);
    std::vector<size_t> boundaries;
    boundaries.reserve(static_cast<size_t>(clusterCount) + 1u);
    boundaries.push_back(0u);
    for (UINT32 index = 0u; index < clusterCount; ++index)
        boundaries.push_back(boundaries.back() + clusters[index].length);

    fixedPrefixChars               = std::min(fixedPrefixChars, text.size());
    const auto afterPrefix         = std::upper_bound(boundaries.begin(), boundaries.end(), fixedPrefixChars);
    fixedPrefixChars               = *std::prev(afterPrefix);
    const std::wstring_view prefix = text.substr(0u, fixedPrefixChars);
    if (MeasureTextWidth(factory, format, std::wstring(prefix) + std::wstring(ellipsisText), maxWidth, height) >= maxWidth)
        return std::wstring(ellipsisText);

    const auto candidateAt = [&](size_t boundary)
    {
        std::wstring candidate(prefix);
        candidate.append(ellipsisText);
        candidate.append(text.substr(boundaries[boundary]));
        return candidate;
    };
    // An entire filename is a preference, not a minimum width: a long leaf must also abbreviate.
    // Searching all valid suffixes retains the extension whenever it fits and always fits the row.
    size_t low  = static_cast<size_t>(std::lower_bound(boundaries.begin(), boundaries.end(), fixedPrefixChars) - boundaries.begin());
    size_t high = boundaries.size() - 1u;
    while (low < high)
    {
        const size_t mid = low + (high - low) / 2u;
        if (MeasureTextWidth(factory, format, candidateAt(mid), maxWidth, height) <= maxWidth)
            high = mid;
        else
            low = mid + 1u;
    }
    return candidateAt(low);
}

size_t ComputePathFixedPrefixChars(std::wstring_view path) noexcept
{
    if (path.size() >= 3 && path[1] == L':' && (path[2] == L'\\' || path[2] == L'/'))
    {
        return 3u;
    }

    if (! path.empty() && (path.front() == L'\\' || path.front() == L'/'))
    {
        return 1u;
    }

    return 0u;
}

float RateSampleHue(std::wstring_view sourcePath) noexcept
{
    if (sourcePath.empty())
    {
        return -1.0f;
    }

    const uint32_t pathHash = StableHash32(sourcePath);
    return static_cast<float>(pathHash % 360u);
}

[[nodiscard]] float AssignedRateStreamHue(const FileOperationsPopupInternal::RateHistory* history,
                                          const FileOperationsPopupInternal::TaskSnapshot::InFlightFileSnapshot* stream) noexcept
{
    if (! history || ! stream)
    {
        return -1.0f;
    }

    for (size_t i = 0; i < history->streamProgressCount && i < history->streamProgress.size(); ++i)
    {
        const auto& candidate = history->streamProgress[i];
        if (candidate.cookieKey == stream->cookieKey && candidate.progressStreamId == stream->progressStreamId && candidate.sourcePath == stream->sourcePath)
        {
            return candidate.assignedHue;
        }
    }
    return -1.0f;
}

[[nodiscard]] D2D1_COLOR_F RainbowProgressColor(const AppTheme& theme,
                                                std::wstring_view seed,
                                                const FileOperationsPopupInternal::RateHistory* history                       = nullptr,
                                                const FileOperationsPopupInternal::TaskSnapshot::InFlightFileSnapshot* stream = nullptr) noexcept
{
    float hue = AssignedRateStreamHue(history, stream);
    if (hue < 0.0f)
    {
        hue = RateSampleHue(seed);
    }
    if (hue >= 0.0f)
    {
        return DxUi::ThroughputGraphColorFromHue(hue, theme.dark);
    }
    const D2D1::ColorF accent = theme.navigationView.accent;
    return D2D1_COLOR_F{accent.r, accent.g, accent.b, accent.a};
}

bool IsRateSamplingBlocked(const FileOperationsPopupInternal::RateSnapshot& task) noexcept
{
    // I35 R10.2: a question waiting for the user moves no bytes by design. That is neither
    // throughput nor a stall, so it is measured like a pause.
    return task.paused || task.queuePaused || task.waitingInQueue || task.waitingForDecision;
}

void AddHueWeight(std::array<FileOperationsPopupInternal::RateHistory::HueWeight, FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample>& weights,
                  size_t& weightCount,
                  uint8_t colorSlot,
                  float hue,
                  double weight) noexcept
{
    if (colorSlot >= FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample || hue < 0.0f || ! std::isfinite(weight) || weight <= 0.0)
    {
        return;
    }

    for (size_t i = 0; i < weightCount && i < weights.size(); ++i)
    {
        if (weights[i].colorSlot == colorSlot)
        {
            weights[i].weight += weight;
            return;
        }
    }

    if (weightCount < weights.size())
    {
        weights[weightCount] = {.hue = hue, .weight = weight, .colorSlot = colorSlot};
        ++weightCount;
    }
}

void AddPendingHueWeight(FileOperationsPopupInternal::RateHistory& history, uint8_t colorSlot, float hue, double weight) noexcept
{
    AddHueWeight(history.pendingHueWeights, history.pendingHueWeightCount, colorSlot, hue, weight);
}

[[nodiscard]] float DominantHue(
    const std::array<FileOperationsPopupInternal::RateHistory::HueWeight, FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample>& weights,
    size_t weightCount,
    float fallbackHue) noexcept
{
    float dominantHue     = fallbackHue;
    double dominantWeight = 0.0;
    for (size_t i = 0; i < weightCount && i < weights.size(); ++i)
    {
        if (weights[i].weight > dominantWeight)
        {
            dominantWeight = weights[i].weight;
            dominantHue    = weights[i].hue;
        }
    }
    return dominantHue;
}

[[nodiscard]] size_t MaximumConcurrentRateHistoryColorSlots(const FileOperationsPopupInternal::RateHistory& history) noexcept
{
    size_t maximum = 0u;
    for (size_t sampleOffset = 0u; sampleOffset < history.count; ++sampleOffset)
    {
        const size_t sampleIndex = (history.writeIndex + FileOperationsPopupInternal::RateHistory::kMaxSamples - history.count + sampleOffset) %
                                   FileOperationsPopupInternal::RateHistory::kMaxSamples;
        std::array<bool, FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample> present{};
        size_t count             = 0u;
        const size_t weightCount = std::min<size_t>(history.hueWeightCounts[sampleIndex], history.hueWeights[sampleIndex].size());
        for (size_t weightIndex = 0u; weightIndex < weightCount; ++weightIndex)
        {
            const auto& weight = history.hueWeights[sampleIndex][weightIndex];
            if (weight.weight <= 0.0 || weight.colorSlot >= present.size() || present[weight.colorSlot])
            {
                continue;
            }
            present[weight.colorSlot] = true;
            ++count;
        }
        maximum = std::max(maximum, count);
    }
    return maximum;
}

[[nodiscard]] bool RateHistoryGraphBandsAreActive(const FileOperationsPopupInternal::RateHistory* history, const AppTheme& theme) noexcept
{
    const size_t maximumConcurrentColorSlots = history ? MaximumConcurrentRateHistoryColorSlots(*history) : 0u;
    return DxUi::ShouldRenderThroughputGraphBands(theme.menu.rainbowMode, true, theme.highContrast, maximumConcurrentColorSlots);
}

void ResetRateStreamProgress(FileOperationsPopupInternal::RateHistory& history) noexcept
{
    history.streamProgressCount = 0u;
}

void SyncRateStreamBaselines(FileOperationsPopupInternal::RateHistory& history, const FileOperationsPopupInternal::RateSnapshot& task) noexcept
{
    history.streamProgressCount = 0u;
    for (size_t i = 0; i < task.inFlightFileCount && i < task.inFlightFiles.size() && history.streamProgressCount < history.streamProgress.size(); ++i)
    {
        const auto& stream      = task.inFlightFiles[i];
        auto& entry             = history.streamProgress[history.streamProgressCount++];
        entry.cookieKey         = stream.cookieKey;
        entry.progressStreamId  = stream.progressStreamId;
        entry.sourcePath        = stream.sourcePath;
        entry.completedBytes    = stream.completedBytes;
        entry.lastUpdateTick    = stream.lastUpdateTick;
        entry.assignedHue       = -1.0f;
        entry.assignedColorSlot = DxUi::ThroughputGraphSample::kInvalidColorSlot;
    }
}

[[nodiscard]] uint8_t AssignFirstFreeStreamColorSlot(std::array<bool, FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample>& occupied) noexcept
{
    const auto available = std::ranges::find(occupied, false);
    return available == occupied.end() ? DxUi::ThroughputGraphSample::kInvalidColorSlot : static_cast<uint8_t>(std::distance(occupied.begin(), available));
}

// The fixed slot palette keeps concurrent colors well separated while bounding
// historical identity and renderer geometry cardinality.
[[nodiscard]] float StreamHueFromColorSlot(uint8_t colorSlot) noexcept
{
    constexpr float kGoldenAngleDegrees = 137.5083f;
    return colorSlot < FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample
               ? std::fmod(static_cast<float>(colorSlot) * kGoldenAngleDegrees, 360.0f)
               : -1.0f;
}

void AccumulateStreamHueWeights(FileOperationsPopupInternal::RateHistory& history,
                                const FileOperationsPopupInternal::RateSnapshot& task,
                                uint64_t aggregateDeltaBytes,
                                float fallbackHue) noexcept
{
    using StreamProgress = FileOperationsPopupInternal::RateHistory::StreamProgress;

    ++history.debugAccumulateCalls;
    history.debugMaxStreamsSeen = std::max(history.debugMaxStreamsSeen, static_cast<uint32_t>(task.inFlightFileCount));

    std::array<StreamProgress, FileOperationsPopupInternal::TaskSnapshot::kMaxInFlightFiles> previous{};
    const size_t previousCount = std::min(history.streamProgressCount, history.streamProgress.size());
    for (size_t i = 0; i < previousCount; ++i)
    {
        previous[i] = history.streamProgress[i];
    }

    history.streamProgressCount = 0u;

    std::array<bool, FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample> occupiedSlots{};
    for (size_t i = 0; i < task.inFlightFileCount && i < task.inFlightFiles.size(); ++i)
    {
        const auto& stream             = task.inFlightFiles[i];
        const StreamProgress* baseline = nullptr;
        for (size_t j = 0; j < previousCount; ++j)
        {
            if (previous[j].cookieKey == stream.cookieKey && previous[j].progressStreamId == stream.progressStreamId)
            {
                baseline = &previous[j];
                break;
            }
        }

        const bool sameItem = baseline && baseline->sourcePath == stream.sourcePath;
        if (sameItem && baseline->assignedColorSlot < occupiedSlots.size())
        {
            occupiedSlots[baseline->assignedColorSlot] = true;
        }
    }

    // Weights are the active streams' CUMULATIVE byte shares, sampled every tick. Per-tick byte
    // deltas would under-represent streams whose copy callbacks arrive less often than the
    // display bucket (CopyFileEx reports per ~1MB chunk, so a throttled stream can be silent for
    // whole seconds); cumulative shares stay correct regardless of callback cadence and keep
    // equal streams at visually equal bands at every instant.
    for (size_t i = 0; i < task.inFlightFileCount && i < task.inFlightFiles.size(); ++i)
    {
        const auto& stream             = task.inFlightFiles[i];
        const StreamProgress* baseline = nullptr;
        for (size_t j = 0; j < previousCount; ++j)
        {
            if (previous[j].cookieKey == stream.cookieKey && previous[j].progressStreamId == stream.progressStreamId)
            {
                baseline = &previous[j];
                break;
            }
        }

        const bool sameItem     = baseline && baseline->sourcePath == stream.sourcePath;
        uint8_t streamColorSlot = DxUi::ThroughputGraphSample::kInvalidColorSlot;
        if (sameItem && baseline->assignedColorSlot < FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample)
        {
            streamColorSlot = baseline->assignedColorSlot;
        }
        else
        {
            streamColorSlot = AssignFirstFreeStreamColorSlot(occupiedSlots);
            if (streamColorSlot < occupiedSlots.size())
            {
                occupiedSlots[streamColorSlot] = true;
            }
        }
        const float streamHue = StreamHueFromColorSlot(streamColorSlot);

        if (stream.completedBytes > 0)
        {
            AddPendingHueWeight(history, streamColorSlot, streamHue, static_cast<double>(stream.completedBytes));
        }

        if (history.streamProgressCount < history.streamProgress.size())
        {
            auto& entry             = history.streamProgress[history.streamProgressCount++];
            entry.cookieKey         = stream.cookieKey;
            entry.progressStreamId  = stream.progressStreamId;
            entry.sourcePath        = stream.sourcePath;
            entry.completedBytes    = stream.completedBytes;
            entry.lastUpdateTick    = stream.lastUpdateTick;
            entry.assignedHue       = streamHue;
            entry.assignedColorSlot = streamColorSlot;
        }
    }

    // When throttling delays the first byte-progress callback, concurrent streams are still real
    // live graph bands. Preserve their distinct hues instead of collapsing the bucket to the
    // aggregate fallback color.
    if (history.pendingHueWeightCount == 0u && history.streamProgressCount > 1u)
    {
        for (size_t i = 0; i < history.streamProgressCount && i < history.streamProgress.size(); ++i)
        {
            AddPendingHueWeight(history, history.streamProgress[i].assignedColorSlot, history.streamProgress[i].assignedHue, 1.0);
        }
    }

    // No stream data at all (e.g. providers that never report streams): fall back to the
    // aggregate so the graph still gets a color. AppendRateSample carries the previous bucket's
    // distribution when even that is absent.
    if (history.pendingHueWeightCount == 0u && aggregateDeltaBytes > 0)
    {
        // Aggregate-only samples intentionally have no admitted stream slot and
        // therefore stay on the base accent rather than expanding the palette.
        static_cast<void>(fallbackHue);
    }

    history.debugLastPendingCount = static_cast<uint32_t>(history.pendingHueWeightCount);
}

[[nodiscard]] double ClampFiniteNonNegative(double value) noexcept
{
    return std::isfinite(value) && value > 0.0 ? value : 0.0;
}

[[nodiscard]] double SmoothFileOperationValue(double previousValue, double sampleValue, ULONGLONG elapsedMs, double tauMs, double maxAlpha) noexcept
{
    previousValue = ClampFiniteNonNegative(previousValue);
    sampleValue   = ClampFiniteNonNegative(sampleValue);
    if (previousValue <= 0.0 || elapsedMs == 0)
    {
        return sampleValue;
    }

    const double elapsed = static_cast<double>(elapsedMs);
    const double alpha   = std::clamp(1.0 - std::exp(-elapsed / tauMs), 0.02, maxAlpha);
    return previousValue + (sampleValue - previousValue) * alpha;
}

[[nodiscard]] double SmoothEtaSecondsForDisplay(double previousEtaSeconds, double sampleEtaSeconds, ULONGLONG elapsedMs) noexcept
{
    constexpr double kEtaSmoothingTauMs = 1500.0;
    constexpr double kEtaMaxAlpha       = 0.45;
    return SmoothFileOperationValue(previousEtaSeconds, sampleEtaSeconds, elapsedMs, kEtaSmoothingTauMs, kEtaMaxAlpha);
}

// I35 R10.2: the displayed speed is what the task delivered over the last kRateWindowMs of wall
// clock. A real stall lowers it honestly, nothing is held or guessed, and a silence followed by a
// burst within the window is averaged instead of drawn as a trough.
constexpr ULONGLONG kRateWindowMs = 5'000ull;
// Below this much measured time a window says nothing reliable: the speed stays unknown.
constexpr ULONGLONG kRateWindowMinimumSpanMs = 1'000ull;
// Window samples closer than this replace the newest one, so a fast UI-motion timer cannot push the
// window's base out of the ring.
constexpr ULONGLONG kRateWindowSampleSpacingMs = 50ull;
// A task that delivered nothing for longer than this is stalled: the popup names what it waits on. A
// task whose own report cadence is coarser waits 1.5 x its median gap between byte reports instead.
constexpr ULONGLONG kRateStallThresholdMs = 600ull;
// Before the task's cadence is known (two byte reports since measuring started), a silence is a stall
// only after this long: the first bytes of a slow device or a share legitimately take a moment.
constexpr ULONGLONG kRateStallStartupThresholdMs = 2'000ull;
// A transfer whose window speed is at least this share of its speed limit is held by the limit.
constexpr double kRateHeldAtLimitShare = 0.8;
static_assert(FileOperationsPopupInternal::RateHistory::kMaxWindowSamples * kRateWindowSampleSpacingMs > kRateWindowMs,
              "the window ring must hold a base sample older than the window");

using RateWindowSample = FileOperationsPopupInternal::RateHistory::WindowSample;

[[nodiscard]] RateWindowSample& RateWindowAt(FileOperationsPopupInternal::RateHistory& history, size_t offset) noexcept
{
    return history.window[(history.windowFirst + offset) % FileOperationsPopupInternal::RateHistory::kMaxWindowSamples];
}

[[nodiscard]] const RateWindowSample& RateWindowAt(const FileOperationsPopupInternal::RateHistory& history, size_t offset) noexcept
{
    return history.window[(history.windowFirst + offset) % FileOperationsPopupInternal::RateHistory::kMaxWindowSamples];
}

// Measuring restarts at floorTick: the window never reaches back before it.
void ResetRateWindow(FileOperationsPopupInternal::RateHistory& history, ULONGLONG floorTick) noexcept
{
    history.windowFirst     = 0u;
    history.windowCount     = 0u;
    history.windowFloorTick = floorTick;
}

[[nodiscard]] ULONGLONG RateWindowStartTick(const FileOperationsPopupInternal::RateHistory& history, ULONGLONG nowTick) noexcept
{
    const ULONGLONG start = nowTick > kRateWindowMs ? nowTick - kRateWindowMs : 0ull;
    return std::max(start, history.windowFloorTick);
}

// Records the task's cumulative values at this tick and drops the samples the window no longer
// needs, keeping one sample at or before the window start as the interpolation base.
void AppendRateWindowSample(FileOperationsPopupInternal::RateHistory& history,
                            const FileOperationsPopupInternal::RateSnapshot& task,
                            ULONGLONG nowTick) noexcept
{
    using FileOperationsPopupInternal::RateHistory;
    const RateWindowSample sample{.tick = nowTick, .bytes = task.completedBytes, .verificationBytes = task.verificationReadBytes, .items = task.completedItems};
    if (history.windowCount > 0u)
    {
        const RateWindowSample& newest = RateWindowAt(history, history.windowCount - 1u);
        if (sample.bytes < newest.bytes || sample.verificationBytes < newest.verificationBytes || sample.items < newest.items)
        {
            // A cumulative value went down (a rolled-back stage): restart rather than average a
            // negative delta.
            ResetRateWindow(history, nowTick);
        }
        else if (nowTick <= newest.tick || (history.windowCount >= 2u && nowTick - newest.tick < kRateWindowSampleSpacingMs))
        {
            RateWindowAt(history, history.windowCount - 1u) = RateWindowSample{
                .tick = std::max(nowTick, newest.tick), .bytes = sample.bytes, .verificationBytes = sample.verificationBytes, .items = sample.items};
            return;
        }
    }
    if (history.windowCount == RateHistory::kMaxWindowSamples)
    {
        history.windowFirst = (history.windowFirst + 1u) % RateHistory::kMaxWindowSamples;
        --history.windowCount;
    }
    RateWindowAt(history, history.windowCount) = sample;
    ++history.windowCount;
    const ULONGLONG windowStart = RateWindowStartTick(history, nowTick);
    while (history.windowCount >= 2u && RateWindowAt(history, 1u).tick <= windowStart)
    {
        history.windowFirst = (history.windowFirst + 1u) % RateHistory::kMaxWindowSamples;
        --history.windowCount;
    }
}

struct WindowedRates final
{
    double bytesPerSec             = 0.0;
    double verificationBytesPerSec = 0.0;
    double itemsPerSec             = 0.0;
    bool available                 = false;
};

// The rates over [window start, newest sample]; the cumulative values at the window start are
// interpolated between the two samples around it.
[[nodiscard]] WindowedRates ComputeWindowedRates(const FileOperationsPopupInternal::RateHistory& history, ULONGLONG nowTick) noexcept
{
    WindowedRates rates{};
    if (history.windowCount < 2u)
        return rates;
    const RateWindowSample& oldest = RateWindowAt(history, 0u);
    const RateWindowSample& newest = RateWindowAt(history, history.windowCount - 1u);
    const ULONGLONG start          = std::max(RateWindowStartTick(history, nowTick), oldest.tick);
    if (newest.tick <= start || newest.tick - start < kRateWindowMinimumSpanMs)
        return rates;

    size_t baseIndex = 0u;
    while (baseIndex + 2u < history.windowCount && RateWindowAt(history, baseIndex + 1u).tick <= start)
        ++baseIndex;
    const RateWindowSample& before = RateWindowAt(history, baseIndex);
    const RateWindowSample& after  = RateWindowAt(history, baseIndex + 1u);
    const double fraction =
        after.tick > before.tick ? std::clamp(static_cast<double>(start - before.tick) / static_cast<double>(after.tick - before.tick), 0.0, 1.0) : 0.0;
    const auto deltaSince = [fraction](uint64_t from, uint64_t to, uint64_t now) noexcept
    {
        const double base = static_cast<double>(from) + (static_cast<double>(to) - static_cast<double>(from)) * fraction;
        return static_cast<double>(now) - base;
    };
    const double spanSeconds = static_cast<double>(newest.tick - start) / 1000.0;
    rates.bytesPerSec        = ClampFiniteNonNegative(deltaSince(before.bytes, after.bytes, newest.bytes) / spanSeconds);
    rates.verificationBytesPerSec =
        ClampFiniteNonNegative(deltaSince(before.verificationBytes, after.verificationBytes, newest.verificationBytes) / spanSeconds);
    rates.itemsPerSec = ClampFiniteNonNegative(deltaSince(before.items, after.items, newest.items) / spanSeconds);
    rates.available   = true;
    return rates;
}

#ifdef ENABLE_TESTS
[[nodiscard]] float EaseGraphLatestPointYForDisplay(float previousY, float targetY, ULONGLONG elapsedMs) noexcept
{
    constexpr ULONGLONG kLineEaseMs = 260ull;
    if (elapsedMs >= kLineEaseMs)
    {
        return targetY;
    }

    const float easedT = EaseFileOperationsUiMotionFraction(elapsedMs, kLineEaseMs);
    return previousY + (targetY - previousY) * easedT;
}
#endif

[[nodiscard]] double CurrentBandwidthForGraphMarker(const FileOperationsPopupInternal::RateHistory& history) noexcept
{
    // I35 R10.2: the marker shows the published window rate only. It never falls back to an older
    // plotted sample, which would hold a rate through a pause or a stall.
    return ClampFiniteNonNegative(history.displayedBytesPerSec) + ClampFiniteNonNegative(history.displayedVerificationBytesPerSec);
}

void AppendRateSample(FileOperationsPopupInternal::RateHistory& history, double sample, double verificationSample, float hue, bool stalled = false) noexcept
{
    const double clampedSample = ClampFiniteNonNegative(sample);
    const size_t slot          = history.writeIndex;

    history.samples[slot]             = static_cast<float>(std::min<double>(clampedSample, std::numeric_limits<float>::max()));
    history.verificationSamples[slot] = static_cast<float>(std::min<double>(ClampFiniteNonNegative(verificationSample), std::numeric_limits<float>::max()));
    history.stalledSamples[slot]      = stalled ? 1u : 0u;

    const size_t pendingCount = std::min(history.pendingHueWeightCount, history.pendingHueWeights.size());
    if (pendingCount > 0u)
    {
        const size_t storedCount      = std::min(pendingCount, history.hueWeights[slot].size());
        history.hueWeightCounts[slot] = static_cast<uint8_t>(storedCount);
        for (size_t i = 0; i < storedCount; ++i)
        {
            history.hueWeights[slot][i] = history.pendingHueWeights[i];
        }
        history.hues[slot] = DominantHue(history.hueWeights[slot], storedCount, hue);
    }
    else
    {
        // No fresh per-stream weights for this bucket (timer jitter, multi-bucket flush): carry
        // the previous bucket's distribution forward instead of recoloring the column with a
        // single hue — the latest callback must never repaint a whole sample by itself.
        const size_t previousSlot = (slot + FileOperationsPopupInternal::RateHistory::kMaxSamples - 1u) % FileOperationsPopupInternal::RateHistory::kMaxSamples;
        if (history.count > 0u && history.hueWeightCounts[previousSlot] > 0u)
        {
            history.hueWeightCounts[slot] = history.hueWeightCounts[previousSlot];
            history.hueWeights[slot]      = history.hueWeights[previousSlot];
            history.hues[slot]            = history.hues[previousSlot];
        }
        else
        {
            history.hueWeightCounts[slot] = hue >= 0.0f ? 1u : 0u;
            if (hue >= 0.0f)
            {
                history.hueWeights[slot][0] = {
                    .hue       = hue,
                    .weight    = 1.0,
                    .colorSlot = DxUi::ThroughputGraphSample::kInvalidColorSlot,
                };
            }
            history.hues[slot] = hue;
        }
    }

    history.writeIndex = (history.writeIndex + 1u) % FileOperationsPopupInternal::RateHistory::kMaxSamples;
    history.count      = std::min(FileOperationsPopupInternal::RateHistory::kMaxSamples, history.count + 1u);
}

void ResetPendingRateSample(FileOperationsPopupInternal::RateHistory& history) noexcept
{
    history.pendingBucketMs                     = 0;
    history.pendingWeightedSampleMs             = 0.0;
    history.pendingWeightedVerificationSampleMs = 0.0;
    history.pendingStalledMs                    = 0;
    history.pendingHue                          = -1.0f;
    history.pendingHueWeightCount               = 0u;
}

void AppendResampledRateSamples(
    FileOperationsPopupInternal::RateHistory& history, ULONGLONG elapsedMs, double sample, double verificationSample, float hue, bool stalled = false) noexcept
{
    const ULONGLONG maxResampleMs = static_cast<ULONGLONG>(FileOperationsPopupInternal::RateHistory::kMaxSamples) * kRateSampleBucketMs;
    ULONGLONG remainingMs         = (std::min)(elapsedMs, maxResampleMs);
    while (remainingMs > 0)
    {
        const ULONGLONG bucketRemainingMs = kRateSampleBucketMs - history.pendingBucketMs;
        const ULONGLONG sliceMs           = std::min(remainingMs, bucketRemainingMs);

        history.pendingWeightedSampleMs += static_cast<double>(sample) * static_cast<double>(sliceMs);
        history.pendingWeightedVerificationSampleMs += verificationSample * static_cast<double>(sliceMs);
        history.pendingBucketMs += sliceMs;
        if (stalled)
        {
            history.pendingStalledMs += sliceMs;
        }
        if (hue >= 0.0f)
        {
            history.pendingHue = hue;
        }

        remainingMs -= sliceMs;
        if (history.pendingBucketMs < kRateSampleBucketMs)
        {
            continue;
        }

        const float bucketHue                 = history.pendingHue >= 0.0f ? history.pendingHue : hue;
        const double bucketSample             = history.pendingWeightedSampleMs / static_cast<double>(kRateSampleBucketMs);
        const double bucketVerificationSample = history.pendingWeightedVerificationSampleMs / static_cast<double>(kRateSampleBucketMs);
        // A bucket belongs to the stall band when the stall covers at least half of it.
        const bool bucketStalled = history.pendingStalledMs * 2u >= kRateSampleBucketMs;
        AppendRateSample(history, bucketSample, bucketVerificationSample, bucketHue, bucketStalled);
        ResetPendingRateSample(history);
    }
}

// I35 R10.2: a stall is recognized only after kRateStallThresholdMs without bytes, but its band
// starts where the bytes stopped: mark the last spanMs of plotted time, the pending bucket first.
void MarkRecentRateSamplesStalled(FileOperationsPopupInternal::RateHistory& history, ULONGLONG spanMs) noexcept
{
    using FileOperationsPopupInternal::RateHistory;
    const ULONGLONG pendingStalledMs = std::min(spanMs, history.pendingBucketMs);
    history.pendingStalledMs         = std::max(history.pendingStalledMs, pendingStalledMs);
    ULONGLONG remainingMs            = spanMs - pendingStalledMs;
    for (size_t offset = 1u; remainingMs > 0u && offset <= history.count; ++offset)
    {
        const size_t slot = (history.writeIndex + RateHistory::kMaxSamples - offset) % RateHistory::kMaxSamples;
        if (remainingMs * 2u >= kRateSampleBucketMs)
        {
            history.stalledSamples[slot] = 1u;
        }
        remainingMs -= std::min(remainingMs, kRateSampleBucketMs);
    }
}

// I35 R10.2: records the gap since the task's previous byte-advancing report (its report cadence).
void NoteByteReportCadence(FileOperationsPopupInternal::RateHistory& history, ULONGLONG lastByteProgressTick) noexcept
{
    using FileOperationsPopupInternal::RateHistory;
    if (lastByteProgressTick == 0u || lastByteProgressTick <= history.lastByteReportTick)
        return;
    if (history.lastByteReportTick != 0u)
    {
        history.byteReportGapsMs[history.byteReportGapNext] = lastByteProgressTick - history.lastByteReportTick;
        history.byteReportGapNext                           = (history.byteReportGapNext + 1u) % RateHistory::kMaxByteReportGaps;
        history.byteReportGapCount                          = std::min(history.byteReportGapCount + 1u, RateHistory::kMaxByteReportGaps);
    }
    history.lastByteReportTick = lastByteProgressTick;
}

// I35 R10.2: how long a running transfer may deliver nothing before it is stalled: 1.5 x the median
// gap between its recent byte reports, never less than 600 ms, so the steady cadence of a throttled or
// slow destination (a 1 MiB chunk a second) is not a stall; 2 s while that cadence is not known yet.
[[nodiscard]] ULONGLONG RateStallThresholdMs(const FileOperationsPopupInternal::RateHistory& history) noexcept
{
    using FileOperationsPopupInternal::RateHistory;
    if (history.byteReportGapCount == 0u)
        return kRateStallStartupThresholdMs;
    std::array<ULONGLONG, RateHistory::kMaxByteReportGaps> gaps{};
    const auto end = std::copy_n(history.byteReportGapsMs.begin(), history.byteReportGapCount, gaps.begin());
    const auto mid = gaps.begin() + static_cast<std::ptrdiff_t>(history.byteReportGapCount / 2u);
    std::nth_element(gaps.begin(), mid, end);
    return std::max(kRateStallThresholdMs, *mid + *mid / 2u);
}

// I35 R10.2: the name a stall would get now. Streams that were all committing or publishing their
// file around the last report make it "Finalizing"; otherwise the transfer waits on the slower side.
[[nodiscard]] FileOperationsPopupInternal::RateStallKind RateStallKindFor(FolderWindow::FileOperationState::ProgressPhase phase,
                                                                          FolderWindow::FileOperationState::ProgressStallSide side) noexcept
{
    using Phase = FolderWindow::FileOperationState::ProgressPhase;
    using Side  = FolderWindow::FileOperationState::ProgressStallSide;
    using FileOperationsPopupInternal::RateStallKind;
    if (phase == Phase::Finalizing || phase == Phase::Publishing)
    {
        return RateStallKind::Finalizing;
    }
    switch (side)
    {
        case Side::Destination: return RateStallKind::WaitingForDestination;
        case Side::Source: return RateStallKind::WaitingForSource;
        case Side::Storage: break;
    }
    return RateStallKind::WaitingForStorage;
}

// What a stalled transfer is waiting on, named next to its speed (I35 R10.2).
[[nodiscard]] UINT RateStallTextId(FileOperationsPopupInternal::RateStallKind stall) noexcept
{
    using FileOperationsPopupInternal::RateStallKind;
    switch (stall)
    {
        case RateStallKind::Finalizing: return IDS_FILEOPS_STALL_FINALIZING;
        case RateStallKind::WaitingForDestination: return IDS_FILEOPS_STALL_WAITING_DESTINATION;
        case RateStallKind::WaitingForSource: return IDS_FILEOPS_STALL_WAITING_SOURCE;
        case RateStallKind::WaitingForStorage: return IDS_FILEOPS_STALL_WAITING_STORAGE;
        case RateStallKind::None: break;
    }
    return 0u;
}

// The rate label of a task card: the window speed (or "Rate: —"), and while the task is stalled the
// name of what it waits on next to it.
[[nodiscard]] std::wstring FormatRateLabelText(double bytesPerSecond, FileOperationsPopupInternal::RateStallKind stall)
{
    std::wstring rateText =
        bytesPerSecond > 0.0 ? FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_BYTES, FormatBytesCompact(SaturatingRoundNonNegativeToUint64(bytesPerSecond)))
                             : LoadStringResource(nullptr, IDS_FILEOPS_RATE_UNAVAILABLE);
    const UINT stallTextId = RateStallTextId(stall);
    if (stallTextId == 0u)
    {
        return rateText;
    }
    return FormatStringResource(nullptr, IDS_FMT_FILEOPS_RATE_WITH_STALL, rateText, LoadStringResource(nullptr, stallTextId));
}

struct RateUpdateTotals final
{
    uint64_t maxCallbackSilenceMs = 0;
    uint64_t maxDisplayGapMs      = 0;
    uint64_t silentTaskCount      = 0;
    uint64_t updatedTaskCount     = 0;
};

// One telemetry tick of one task's rate history (I35 R10.2). UpdateRates drives it with the live
// snapshot and the self-tests with synthetic ticks, so the gates exercise the production model.
void AdvanceRateHistory(FileOperationsPopupInternal::RateHistory& history,
                        const FileOperationsPopupInternal::RateSnapshot& task,
                        ULONGLONG nowTick,
                        RateUpdateTotals& totals) noexcept
{
    using FileOperationsPopupInternal::RateStallKind;
    const bool blocked  = IsRateSamplingBlocked(task);
    const bool itemRate = task.operation == FILESYSTEM_DELETE;

    if (! history.initialized)
    {
        history.initialized              = true;
        history.lastProgressCallbackTick = task.lastProgressCallbackTick;
        history.lastDisplaySampleTick    = nowTick;
        history.lastStateChangeTick      = task.progressStateChangeTick;
        history.lastByteAdvanceTick      = nowTick;
        history.lastByteReportTick       = task.lastByteProgressTick;
        ResetRateWindow(history, nowTick);
        AppendRateWindowSample(history, task, nowTick);
        SyncRateStreamBaselines(history, task);
        return;
    }

    if (task.progressStateChangeTick > history.lastStateChangeTick || blocked)
    {
        // Pause, queue waits, a question for the user and the resume after them restart the
        // measurement: a resumed task's speed is measured from the resume, never averaged with the
        // pause, and nothing is displayed while it moves nothing by design. The report cadence goes on
        // from the resume: the gap spanning the pause is not part of it.
        history.lastStateChangeTick      = std::max(history.lastStateChangeTick, task.progressStateChangeTick);
        history.lastProgressCallbackTick = task.lastProgressCallbackTick;
        history.lastDisplaySampleTick    = nowTick;
        history.lastByteAdvanceTick      = nowTick;
        history.lastByteReportTick       = std::max(task.lastByteProgressTick, nowTick);
        history.stallKind                = RateStallKind::None;
        ResetPendingRateSample(history);
        ResetRateStreamProgress(history);
        SyncRateStreamBaselines(history, task);
        ResetRateWindow(history, nowTick);
        AppendRateWindowSample(history, task, nowTick);
        if (blocked)
        {
            history.displayedBytesPerSec             = 0.0;
            history.displayedVerificationBytesPerSec = 0.0;
            history.displayedItemsPerSec             = 0.0;
            history.hasSmoothedEta                   = false;
            history.hasProvisionalEta                = false;
            return;
        }
    }

    const RateWindowSample previous = history.windowCount > 0u ? RateWindowAt(history, history.windowCount - 1u) : RateWindowSample{.tick = nowTick};
    const bool advanced = history.windowCount > 0u && (task.completedBytes > previous.bytes || task.verificationReadBytes > previous.verificationBytes ||
                                                       task.completedItems > previous.items);
    if (advanced)
    {
        // A stall longer than the window ends here. Measuring restarts at its end, as it does after a
        // resume, so the speed after it neither averages the empty window nor jumps back to a rate
        // from before it.
        if (nowTick > history.lastByteAdvanceTick && nowTick - history.lastByteAdvanceTick > kRateWindowMs)
        {
            ResetRateWindow(history, previous.tick);
            RateWindowAt(history, 0u) = previous;
            history.windowCount       = 1u;
        }
        history.lastByteAdvanceTick = nowTick;
        ++totals.updatedTaskCount;
    }
    AppendRateWindowSample(history, task, nowTick);

    if (task.lastProgressCallbackTick != 0 && task.lastProgressCallbackTick > history.lastProgressCallbackTick)
    {
        history.lastProgressCallbackTick = task.lastProgressCallbackTick;
    }
    else if (! task.finished && task.lastProgressCallbackTick != 0 && nowTick >= task.lastProgressCallbackTick)
    {
        const ULONGLONG silenceMs   = nowTick - task.lastProgressCallbackTick;
        totals.maxCallbackSilenceMs = std::max<uint64_t>(totals.maxCallbackSilenceMs, silenceMs);
        if (silenceMs > 0)
        {
            ++totals.silentTaskCount;
        }
    }

    const WindowedRates rates                = ComputeWindowedRates(history, nowTick);
    history.displayedItemsPerSec             = itemRate ? rates.itemsPerSec : 0.0;
    history.displayedBytesPerSec             = itemRate ? 0.0 : rates.bytesPerSec;
    history.displayedVerificationBytesPerSec = itemRate ? 0.0 : rates.verificationBytesPerSec;

    // Stalled: a running transfer delivered nothing for longer than its stall threshold (600 ms, or 1.5 x
    // its own report cadence when that is coarser). A task that never reported progress is still
    // starting, and a deletion has no transfer to wait on. A transfer held at its speed limit (its window
    // speed at least 80 % of the limit, or not measured yet) waits for the limit, and a stopping task
    // waits for its own abort: neither is named a stall.
    NoteByteReportCadence(history, task.lastByteProgressTick);
    const ULONGLONG stallReference = std::max({task.lastByteProgressTick, history.lastByteAdvanceTick, history.windowFloorTick});
    const bool heldAtLimit = task.speedLimitBytesPerSecond != 0u &&
                             (! rates.available || history.displayedBytesPerSec >= kRateHeldAtLimitShare * static_cast<double>(task.speedLimitBytesPerSecond));
    const bool stalled     = ! itemRate && ! task.stopping && ! heldAtLimit && task.started && ! task.finished && task.lastProgressCallbackTick != 0 &&
                             nowTick > stallReference && nowTick - stallReference > RateStallThresholdMs(history);
    const RateStallKind stall = ! stalled ? RateStallKind::None : (task.stallKind != RateStallKind::None ? task.stallKind : RateStallKind::WaitingForStorage);
    if (stall != RateStallKind::None && history.stallKind == RateStallKind::None && history.lastDisplaySampleTick > stallReference)
    {
        MarkRecentRateSamplesStalled(history, history.lastDisplaySampleTick - stallReference);
    }
    history.stallKind = stall;

    const ULONGLONG smoothingElapsedMs = std::max<ULONGLONG>(1ull, nowTick > previous.tick ? nowTick - previous.tick : 0ull);
    if (! itemRate && task.discoveryClosed && task.totalBytes > 0 && task.completedBytes <= task.totalBytes)
    {
        const uint64_t remainingTransferBytes = task.totalBytes - task.completedBytes;
        const uint64_t remainingVerificationBytes =
            task.verificationCompletedBytes < task.verificationTotalBytes ? task.verificationTotalBytes - task.verificationCompletedBytes : 0u;
        const bool transferEtaAvailable     = remainingTransferBytes == 0u || IsByteRateUsableForEta(history.displayedBytesPerSec);
        const bool verificationEtaAvailable = remainingVerificationBytes == 0u || IsByteRateUsableForEta(history.displayedVerificationBytesPerSec);
        if (transferEtaAvailable && verificationEtaAvailable && (remainingTransferBytes > 0u || remainingVerificationBytes > 0u))
        {
            const double rawEtaSeconds =
                (remainingTransferBytes > 0u ? static_cast<double>(remainingTransferBytes) / history.displayedBytesPerSec : 0.0) +
                (remainingVerificationBytes > 0u ? static_cast<double>(remainingVerificationBytes) / history.displayedVerificationBytesPerSec : 0.0);
            history.smoothedEtaSeconds =
                history.hasSmoothedEta ? SmoothEtaSecondsForDisplay(history.smoothedEtaSeconds, rawEtaSeconds, smoothingElapsedMs) : rawEtaSeconds;
            history.hasSmoothedEta = true;
        }
        else
        {
            history.hasSmoothedEta = false;
        }
    }
    else if (! itemRate)
    {
        history.hasSmoothedEta = false;
    }

    // D2-A08 provisional ETA: only while discovery is open, over the workload discovered so far,
    // from the same usable window rate. It may rise or fall, and it disappears without a usable rate
    // (paused, a stall that emptied the window, no callbacks) or once the total closes.
    if (! itemRate && ! task.discoveryClosed && ! task.paused && task.discoveredTotalBytes > task.completedBytes &&
        IsByteRateUsableForEta(history.displayedBytesPerSec))
    {
        const double rawProvisionalSeconds = static_cast<double>(task.discoveredTotalBytes - task.completedBytes) / history.displayedBytesPerSec;
        history.provisionalEtaSeconds      = history.hasProvisionalEta
                                                 ? SmoothEtaSecondsForDisplay(history.provisionalEtaSeconds, rawProvisionalSeconds, smoothingElapsedMs)
                                                 : rawProvisionalSeconds;
        history.hasProvisionalEta          = true;
    }
    else
    {
        history.hasProvisionalEta = false;
    }

    if (task.finished)
    {
        return;
    }

    // Per-stream hue attribution samples the in-flight cumulative shares every tick, independent of
    // how often the published aggregate or the plugin callbacks advance.
    const float hue = RateSampleHue(task.currentSourcePath);
    if (! itemRate)
    {
        AccumulateStreamHueWeights(history, task, 0u, hue);
    }

    if (history.lastDisplaySampleTick == 0 || history.lastDisplaySampleTick > nowTick)
    {
        history.lastDisplaySampleTick = nowTick;
        return;
    }
    const ULONGLONG displayElapsedMs = nowTick - history.lastDisplaySampleTick;
    if (displayElapsedMs == 0)
    {
        return;
    }
    totals.maxDisplayGapMs                 = std::max<uint64_t>(totals.maxDisplayGapMs, displayElapsedMs);
    const double displayVerificationSample = itemRate ? 0.0 : history.displayedVerificationBytesPerSec;
    const double displaySample             = itemRate ? history.displayedItemsPerSec : history.displayedBytesPerSec;
    // The plot shows the published window rate. While the window is still too short to measure,
    // nothing is plotted, so the plot never starts with an invented zero.
    if (rates.available && (displaySample > 0.0 || history.count > 0 || history.pendingHueWeightCount > 0u))
    {
        AppendResampledRateSamples(history, displayElapsedMs, displaySample, displayVerificationSample, hue, stall != RateStallKind::None);
    }
    history.lastDisplaySampleTick = nowTick;
}

std::wstring TruncatePathMiddleToWidth(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view path, float maxWidth, float height) noexcept
{
    const size_t prefixChars = ComputePathFixedPrefixChars(path);

    return TruncateTextMiddleToWidth(factory, format, path, maxWidth, height, kEllipsisText, prefixChars);
}

// The named "Destination unchanged" notice is one clipped line with the blocking name at its end.
// Abbreviate only that name (middle, keeping its extension) so the sentence and the useful part both
// stay visible; the exact text remains on the tooltip and UIA region.

ATOM RegisterFileOperationsPopupWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = FileOperationsPopupInternal::FileOperationsPopupState::WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    wc.hIconSm       = LoadIconW(instance, MAKEINTRESOURCEW(IDI_SMALL));
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kFileOperationsPopupClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

} // namespace

using FileOperationsPopupInternal::PopupButton;
using FileOperationsPopupInternal::PopupHitTest;
using FileOperationsPopupInternal::RateHistory;
using FileOperationsPopupInternal::RateSnapshot;
using FileOperationsPopupInternal::TaskSnapshot;

#ifdef ENABLE_TESTS
void PopulateGraphHueDebugSummary(const RateHistory& history, FileOperationsPopupInternal::PopupLayoutDebugSnapshot& result) noexcept
{
    result.graphMultiHueBucketCount  = 0u;
    result.graphSingleHueBucketCount = 0u;
    result.graphDistinctHueCount     = 0u;
    result.graphMinHueShare          = 0.0;
    result.graphMaxHueShare          = 0.0;
    result.graphDebugAccumulateCalls = history.debugAccumulateCalls;
    result.graphDebugLastPending     = history.debugLastPendingCount;
    result.graphDebugMaxStreams      = history.debugMaxStreamsSeen;

    std::array<float, 32> distinctHues{};
    std::array<double, 32> hueTotals{};
    size_t distinctCount = 0;
    double totalWeight   = 0.0;

    for (size_t i = 0; i < history.count; ++i)
    {
        const size_t index       = (history.writeIndex + RateHistory::kMaxSamples - history.count + i) % RateHistory::kMaxSamples;
        const size_t weightCount = std::min<size_t>(history.hueWeightCounts[index], history.hueWeights[index].size());
        if (weightCount == 0u)
        {
            continue;
        }

        if (weightCount >= 2u)
        {
            ++result.graphMultiHueBucketCount;
        }
        else
        {
            ++result.graphSingleHueBucketCount;
        }
    }

    // Distinct hues and per-hue shares come from a recent window of multi-hue buckets.
    // The warm-up where streams have not all reported yet must not dilute steady-state fairness.
    constexpr size_t kFairnessWindowBuckets = 30u;
    size_t windowBuckets                    = 0u;
    for (size_t back = 0; back < history.count && windowBuckets < kFairnessWindowBuckets; ++back)
    {
        const size_t index       = (history.writeIndex + RateHistory::kMaxSamples - 1u - back) % RateHistory::kMaxSamples;
        const size_t weightCount = std::min<size_t>(history.hueWeightCounts[index], history.hueWeights[index].size());
        if (weightCount < 2u)
        {
            continue;
        }
        ++windowBuckets;

        for (size_t band = 0; band < weightCount; ++band)
        {
            const auto& weight = history.hueWeights[index][band];
            if (weight.hue < 0.0f || weight.weight <= 0.0)
            {
                continue;
            }

            size_t hueSlot = distinctCount;
            for (size_t k = 0; k < distinctCount; ++k)
            {
                if (distinctHues[k] == weight.hue)
                {
                    hueSlot = k;
                    break;
                }
            }
            if (hueSlot == distinctCount && distinctCount < distinctHues.size())
            {
                distinctHues[distinctCount++] = weight.hue;
            }
            if (hueSlot < hueTotals.size())
            {
                hueTotals[hueSlot] += weight.weight;
                totalWeight += weight.weight;
            }
        }
    }

    result.graphDistinctHueCount = static_cast<uint32_t>(distinctCount);
    if (distinctCount > 0 && totalWeight > 0.0)
    {
        double minShare = 1.0;
        double maxShare = 0.0;
        for (size_t k = 0; k < distinctCount; ++k)
        {
            const double share = hueTotals[k] / totalWeight;
            minShare           = std::min(minShare, share);
            maxShare           = std::max(maxShare, share);
        }
        result.graphMinHueShare = minShare;
        result.graphMaxHueShare = maxShare;
    }
}
#endif

void FileOperationsPopupInternal::FileOperationsPopupState::ApplyScrollBarTheme(HWND hwnd) const noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    if (! folderWindow)
    {
        return;
    }

    const AppTheme& theme = folderWindow->GetTheme();
    if (theme.highContrast)
    {
        SetWindowTheme(hwnd, L"", nullptr);
        return;
    }

    if (theme.dark)
    {
        SetWindowTheme(hwnd, L"DarkMode_Explorer", nullptr);
        return;
    }

    SetWindowTheme(hwnd, L"Explorer", nullptr);
}

#ifdef ENABLE_TESTS
std::wstring FileOperationsPopupInternal::FileOperationsPopupState::DebugHostedPaintedTextForSelfTest(std::wstring_view automationId, bool& fits) const
{
    fits = false;
    for (const HostedTooltipDescriptor& descriptor : _hostedTooltipDescriptors)
    {
        if (descriptor.automationId != automationId)
            continue;
        const float width = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), descriptor.paintedPath, 65536.0f, 128.0f);
        fits              = ! descriptor.paintedPath.empty() && width > 0.0f && width <= descriptor.bounds.right - descriptor.bounds.left + 0.5f;
        return descriptor.paintedPath;
    }
    return {};
}

bool FileOperationsPopupInternal::FileOperationsPopupState::DebugFreezeHostedRegionHoverFlagForSelfTest(std::wstring_view automationId) noexcept
{
    for (DxUi::Control* region : _hostedTooltipRegions)
    {
        if (region && region->GetAccessibleAutomationId() == automationId)
        {
            auto* const pathRegion = dynamic_cast<FileOperationsPathRegion*>(region);
            if (! pathRegion)
                return false;
            pathRegion->DebugFreezeHoverFlag();
            return true;
        }
    }
    return false;
}
#endif

bool FileOperationsPopupInternal::FileOperationsPopupState::IsReducedMotionEnabled() const noexcept
{
    if (folderWindow && hostLifetime.lock())
    {
        const std::optional<bool>& reducedMotionOverride = folderWindow->GetTheme().reducedMotionOverride;
        if (reducedMotionOverride.has_value())
        {
            return reducedMotionOverride.value();
        }
    }

    return _reducedMotion;
}

void FileOperationsPopupInternal::FileOperationsPopupState::RefreshLocalizedFooterText()
{
    _footerNewTasksText       = LoadStringResource(nullptr, IDS_FILEOPS_MODE_NEW_TASKS);
    _footerQueueText          = LoadStringResource(nullptr, IDS_FILEOPS_BTN_MODE_QUEUE);
    _footerParallelText       = LoadStringResource(nullptr, IDS_FILEOPS_BTN_MODE_PARALLEL);
    _footerAutoDismissOnText  = LoadStringResource(nullptr, IDS_FILEOPS_CHECK_AUTODISMISS_ON);
    _footerAutoDismissOffText = LoadStringResource(nullptr, IDS_FILEOPS_CHECK_AUTODISMISS_OFF);
}

bool FileOperationsPopupInternal::FileOperationsPopupState::IsTaskCollapsed(uint64_t taskId) const noexcept
{
    const auto it = _collapsedTasks.find(taskId);
    if (it == _collapsedTasks.end())
    {
        return false;
    }

    return it->second;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::IsTaskCollapsedForDisplay(uint64_t taskId, bool compactDensity) const noexcept
{
    const auto it = _collapsedTasks.find(taskId);
    if (it != _collapsedTasks.end())
    {
        return it->second;
    }

    return compactDensity && _compactExpandedTasks.find(taskId) == _compactExpandedTasks.end();
}

bool FileOperationsPopupInternal::FileOperationsPopupState::IsTaskThroughputExpanded(uint64_t taskId) const noexcept
{
    const auto found = _throughputExpandedTasks.find(taskId);
    return found != _throughputExpandedTasks.end() ? found->second : ! fileOps || fileOps->GetPopupThroughputExpanded();
}

void FileOperationsPopupInternal::FileOperationsPopupState::ToggleTaskCollapsed(uint64_t taskId, bool compactDensity) noexcept
{
    if (compactDensity && _collapsedTasks.find(taskId) == _collapsedTasks.end())
    {
        const auto [it, inserted] = _compactExpandedTasks.insert(taskId);
        if (! inserted)
        {
            _compactExpandedTasks.erase(it);
        }
        return;
    }

    const bool next         = ! IsTaskCollapsedForDisplay(taskId, compactDensity);
    _collapsedTasks[taskId] = next;
    _compactExpandedTasks.erase(taskId);
}

void FileOperationsPopupInternal::FileOperationsPopupState::AutoCollapseCompletedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept
{
    bool addedCollapse = false;
    std::unordered_set<uint64_t> liveTaskIds;
    liveTaskIds.reserve(snapshot.size());
    for (const TaskSnapshot& task : snapshot)
    {
        liveTaskIds.insert(task.taskId);
    }
    std::erase_if(_announcedTaskStatuses, [&](const auto& entry) noexcept { return ! liveTaskIds.contains(entry.first); });

    for (const TaskSnapshot& task : snapshot)
    {
        if (task.finished && _collapsedTasks.find(task.taskId) == _collapsedTasks.end())
        {
            _collapsedTasks[task.taskId] = true;
            _compactExpandedTasks.erase(task.taskId);
            // Retire the expanded height on this terminal paint: no later timer or paint is guaranteed.
#ifdef ENABLE_TESTS
            if (debugAutoCollapseSnapEnabled)
#endif
            {
                _taskHeightAnimations.erase(task.taskId);
                _taskDisclosureAnimations.erase(task.taskId);
            }
            addedCollapse = true;
        }
    }

    if (addedCollapse)
    {
        _maxAutoSizedWindowHeight   = 0;
        _lastAutoSizedContentHeight = -1.0f;
    }
}

// I35 R9.5: once every task the last Cancel all requested has finished (or left), the completed group
// collapses one time when each of its members is a clean success or a clean cancel, so its header's
// badges and Clear completed are in view. This is not a stored preference: the user's toggle wins
// afterwards, and a group with any problem stays as the user left it.
void FileOperationsPopupInternal::FileOperationsPopupState::ResolveBulkCancelGroupCollapse(HWND hwnd,
                                                                                           const std::vector<TaskSnapshot>& snapshot,
                                                                                           bool showCompletedGroup) noexcept
{
    if (_bulkCancelCohort.empty())
        return;
    for (const uint64_t taskId : _bulkCancelCohort)
    {
        const auto member = std::ranges::find(snapshot, taskId, &TaskSnapshot::taskId);
        if (member != snapshot.end() && (! member->finished || (member->kind == TaskSnapshot::Kind::FileOperation && ! member->completedSummaryProjected)))
            return;
        // A member the snapshot does not show yet (its reveal delay) is still running, not gone.
        if (member == snapshot.end() && fileOps && fileOps->FindTask(taskId) != nullptr)
            return;
    }
    _bulkCancelCohort.clear();
    if (! showCompletedGroup || ! _completedGroupExpanded)
        return;
    const bool everyMemberClean = std::ranges::all_of(snapshot,
                                                      [](const TaskSnapshot& task) noexcept
    {
        if (! IsCompletedGroupTask(task))
            return true;
        const TaskStatusKind status = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
        return (status == TaskStatusKind::Done || status == TaskStatusKind::Canceled) && task.warningCount == 0u && task.errorCount == 0u;
    });
    if (! everyMemberClean)
        return;
    _completedGroupExpanded     = false;
    _maxAutoSizedWindowHeight   = 0;
    _lastAutoSizedContentHeight = -1.0f;
    if (hwnd)
        ArmUiMotionTimer(hwnd);
}

void FileOperationsPopupInternal::FileOperationsPopupState::CleanupCollapsedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept
{
    std::unordered_map<uint64_t, bool> seen;
    seen.reserve(snapshot.size());
    for (const TaskSnapshot& task : snapshot)
    {
        seen[task.taskId] = true;
        // Snapshot the remembered default once per card, including temporarily hidden/queued cards.
        // A settings reload changes future cards without moving existing cards or queuing a save.
        _throughputExpandedTasks.try_emplace(task.taskId, ! fileOps || fileOps->GetPopupThroughputExpanded());
    }
    std::erase_if(_throughputExpandedTasks, [&](const auto& entry) { return ! seen.contains(entry.first); });

    for (auto it = _collapsedTasks.begin(); it != _collapsedTasks.end();)
    {
        if (seen.find(it->first) == seen.end())
        {
            it = _collapsedTasks.erase(it);
            continue;
        }
        ++it;
    }

    for (auto it = _compactExpandedTasks.begin(); it != _compactExpandedTasks.end();)
    {
        if (seen.find(*it) == seen.end())
        {
            it = _compactExpandedTasks.erase(it);
            continue;
        }
        ++it;
    }

    for (auto it = _taskHeightAnimations.begin(); it != _taskHeightAnimations.end();)
    {
        if (! seen.contains(it->first))
        {
            it = _taskHeightAnimations.erase(it);
            continue;
        }
        ++it;
    }

    for (auto it = _taskDisclosureAnimations.begin(); it != _taskDisclosureAnimations.end();)
    {
        if (! seen.contains(it->first))
        {
            it = _taskDisclosureAnimations.erase(it);
            continue;
        }
        ++it;
    }
}

float FileOperationsPopupInternal::FileOperationsPopupState::ResolveAnimatedTaskHeight(
    uint64_t taskId, float targetHeight, ULONGLONG tick, bool reducedMotion, bool expanded) noexcept
{
    constexpr ULONGLONG kDurationMs    = kFileOperationsPopupUiMotionDurationMs;
    TaskHeightAnimation& taskAnimation = _taskHeightAnimations[taskId];
    ScalarAnimation& animation         = taskAnimation.height;
    // Expanded content changes immediately, including expansion and active stream rows.
    // Reserve its complete height on that frame. Collapse can animate because its body is hidden.
    reducedMotion          = reducedMotion || (expanded && (! taskAnimation.expanded || animation.target != targetHeight));
    taskAnimation.expanded = expanded;
    if (! animation.initialized || reducedMotion)
    {
        animation.displayed   = targetHeight;
        animation.start       = targetHeight;
        animation.target      = targetHeight;
        animation.startTick   = tick;
        animation.initialized = true;
        return targetHeight;
    }

    const auto advance = [&]() noexcept
    {
        const ULONGLONG elapsed = tick >= animation.startTick ? tick - animation.startTick : 0u;
        const double eased      = static_cast<double>(EaseFileOperationsUiMotionFraction(elapsed, kDurationMs));
        animation.displayed     = animation.start + (animation.target - animation.start) * eased;
    };
    advance();
    if (animation.target != static_cast<double>(targetHeight))
    {
        animation.start     = animation.displayed;
        animation.target    = targetHeight;
        animation.startTick = tick;
    }
    advance();
    return static_cast<float>(animation.displayed);
}

float FileOperationsPopupInternal::FileOperationsPopupState::ResolveCompletedGroupVisibility(ULONGLONG tick, bool reducedMotion) noexcept
{
    constexpr ULONGLONG kDurationMs = kFileOperationsPopupUiMotionDurationMs;
    const double target             = _completedGroupExpanded ? 1.0 : 0.0;
    ScalarAnimation& animation      = _completedGroupVisibilityAnimation;
    if (! animation.initialized || reducedMotion)
    {
        animation.displayed   = target;
        animation.start       = target;
        animation.target      = target;
        animation.startTick   = tick;
        animation.initialized = true;
        return static_cast<float>(target);
    }

    const auto advance = [&]() noexcept
    {
        const ULONGLONG elapsed = tick >= animation.startTick ? tick - animation.startTick : 0u;
        const double eased      = static_cast<double>(EaseFileOperationsUiMotionFraction(elapsed, kDurationMs));
        animation.displayed     = animation.start + (animation.target - animation.start) * eased;
    };
    advance();
    if (animation.target != target)
    {
        animation.start     = animation.displayed;
        animation.target    = target;
        animation.startTick = tick;
    }
    advance();
    return static_cast<float>(animation.displayed);
}

float FileOperationsPopupInternal::FileOperationsPopupState::ResolveDisclosureProgress(ScalarAnimation& animation,
                                                                                       bool expanded,
                                                                                       ULONGLONG tick,
                                                                                       bool reducedMotion) noexcept
{
    constexpr ULONGLONG kDurationMs = kFileOperationsPopupUiMotionDurationMs;
    const double target             = expanded ? 1.0 : 0.0;
    if (! animation.initialized || reducedMotion)
    {
        animation.displayed   = target;
        animation.start       = target;
        animation.target      = target;
        animation.startTick   = tick;
        animation.initialized = true;
        return static_cast<float>(target);
    }

    const auto advance = [&]() noexcept
    {
        const ULONGLONG elapsed = tick >= animation.startTick ? tick - animation.startTick : 0u;
        const float linear      = std::clamp(static_cast<float>(elapsed) / static_cast<float>(kDurationMs), 0.0f, 1.0f);
        const float eased       = DxUi::EvaluateEasing(DxUi::EasingCurve::PointToPoint, linear);
        animation.displayed     = std::lerp(animation.start, animation.target, static_cast<double>(eased));
    };
    advance();
    if (animation.target != target)
    {
        animation.start     = animation.displayed;
        animation.target    = target;
        animation.startTick = tick;
    }
    advance();
    return static_cast<float>(animation.displayed);
}

double FileOperationsPopupInternal::FileOperationsPopupState::ResolveHostedProgressValue(size_t index,
                                                                                         double targetValue,
                                                                                         ULONGLONG tick,
                                                                                         bool reducedMotion) noexcept
{
    constexpr ULONGLONG kDurationMs = 160u;
    if (_hostedProgressAnimations.size() <= index)
    {
        _hostedProgressAnimations.resize(index + 1u);
    }

    ScalarAnimation& animation = _hostedProgressAnimations[index];
    targetValue                = std::clamp(targetValue, 0.0, 100.0);
    if (! animation.initialized || reducedMotion)
    {
        animation.displayed   = targetValue;
        animation.start       = targetValue;
        animation.target      = targetValue;
        animation.startTick   = tick;
        animation.initialized = true;
        return targetValue;
    }

    const auto advance = [&]() noexcept
    {
        const ULONGLONG elapsed = tick >= animation.startTick ? tick - animation.startTick : 0u;
        const double linear     = std::clamp(static_cast<double>(elapsed) / static_cast<double>(kDurationMs), 0.0, 1.0);
        const double eased      = 1.0 - std::pow(1.0 - linear, 3.0);
        animation.displayed     = animation.start + (animation.target - animation.start) * eased;
    };
    advance();
    if (animation.target != targetValue)
    {
        animation.start     = animation.displayed;
        animation.target    = targetValue;
        animation.startTick = tick;
    }
    advance();
    return animation.displayed;
}

void FileOperationsPopupInternal::FileOperationsPopupState::DiscardDeviceResources() noexcept
{
    _preparedBackground.reset();
    _target.reset();
    _captionGlyphTarget.reset();

    _bgBrush.reset();
    _textBrush.reset();
    _subTextBrush.reset();
    _borderBrush.reset();
    _progressBgBrush.reset();
    _progressGlobalBrush.reset();
    _progressItemBrush.reset();
    _progressVerifyBrush.reset();
    _checkboxFillBrush.reset();
    _checkboxCheckBrush.reset();
    _statusOkBrush.reset();
    _statusWarningBrush.reset();
    _statusErrorBrush.reset();
    _graphDynamicBrush.reset();
    _buttonBgBrush.reset();
    _buttonChromeBrush.reset();
    _captionGlyphBrush.reset();
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureFactories() noexcept
{
    if (! _d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, _d2dFactory.addressof());
        if (FAILED(hr))
        {
            _d2dFactory.reset();
        }
    }

    if (! _dwriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.addressof()));
        if (FAILED(hr))
        {
            _dwriteFactory.reset();
        }
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureTextFormats() noexcept
{
    if (! _dwriteFactory)
    {
        return;
    }

    if (_headerFormat && _bodyFormat && _smallFormat && _graphTrailingLabelFormat && _buttonFormat && _buttonSmallFormat && _graphOverlayFormat &&
        _statusIconFallbackFormat)
    {
        return;
    }

    if (! _headerFormat)
    {
        static_cast<void>(DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerFormat.put(), L""));
    }

    if (! _bodyFormat)
    {
        static_cast<void>(
            DxUi::Typography::CreateTextFormat(_dwriteFactory.get(), DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi)), _bodyFormat.put(), L""));
    }

    if (! _smallFormat)
    {
        static_cast<void>(
            DxUi::Typography::CreateTextFormat(_dwriteFactory.get(), DxUi::Typography::MakeUiTextSpec(DipsToPixels(11.0f, _dpi)), _smallFormat.put(), L""));
    }

    if (! _graphTrailingLabelFormat)
    {
        static_cast<void>(DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), DxUi::Typography::MakeUiTextSpec(DipsToPixels(11.0f, _dpi)), _graphTrailingLabelFormat.put(), L""));
    }

    if (! _buttonFormat)
    {
        static_cast<void>(
            DxUi::Typography::CreateTextFormat(_dwriteFactory.get(), DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi)), _buttonFormat.put(), L""));
    }

    if (! _buttonSmallFormat)
    {
        static_cast<void>(DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), DxUi::Typography::MakeUiTextSpec(DipsToPixels(11.0f, _dpi)), _buttonSmallFormat.put(), L""));
    }

    if (! _graphOverlayFormat)
    {
        static_cast<void>(DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), DxUi::Typography::MakeUiTextSpec(DipsToPixels(14.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD), _graphOverlayFormat.put(), L""));
    }

    if (! _statusIconFallbackFormat)
    {
        static_cast<void>(DxUi::Typography::CreateTextFormat(_dwriteFactory.get(),
                                                             DxUi::Typography::MakeUiTextSpec(DipsToPixels(14.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD),
                                                             _statusIconFallbackFormat.put(),
                                                             L""));
    }

    if (! _statusIconFormat)
    {
        // Optional: Segoe Fluent Icons. If missing, fallback format draws standard Unicode glyphs.
        const HRESULT hr =
            DxUi::Typography::CreateTextFormat(_dwriteFactory.get(), DxUi::Typography::MakeUiIconSpec(DipsToPixels(14.0f, _dpi)), _statusIconFormat.put(), L"");
        if (FAILED(hr))
        {
            _statusIconFormat.reset();
        }
    }

    auto configureLineFormat = [](IDWriteTextFormat* format) noexcept
    {
        if (! format)
        {
            return;
        }
        format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    };

    auto configureButtonFormat = [](IDWriteTextFormat* format) noexcept
    {
        if (! format)
        {
            return;
        }
        format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    };

    configureLineFormat(_headerFormat.get());
    configureLineFormat(_bodyFormat.get());
    configureLineFormat(_smallFormat.get());
    configureLineFormat(_graphTrailingLabelFormat.get());
    if (_graphTrailingLabelFormat)
    {
        _graphTrailingLabelFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
    }
    configureButtonFormat(_buttonFormat.get());
    configureButtonFormat(_buttonSmallFormat.get());
    configureButtonFormat(_graphOverlayFormat.get());

    configureButtonFormat(_statusIconFormat.get());
    configureButtonFormat(_statusIconFallbackFormat.get());
}

float FileOperationsPopupInternal::FileOperationsPopupState::MeasurePathLabelWidth(size_t slot, std::wstring_view text) noexcept
{
    if (slot >= _pathLabelWidthsPx.size() || ! _smallFormat || ! _dwriteFactory || text.empty())
        return 0.0f;
    if (_pathLabelsMeasuredFormat.get() != _smallFormat.get())
    {
        _pathLabelsMeasuredFormat = _smallFormat;
        _pathLabelWidthsPx.fill(0.0f);
    }
    if (_pathLabelWidthsPx[slot] > 0.0f && _pathLabelsMeasuredText[slot] == text)
        return _pathLabelWidthsPx[slot];

    const uint64_t startedUs = Debug::Perf::IsCaptureEnabled() ? PerfNowUs() : 0u;
    const float measured     = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), text, DipsToPixels(4096.0f, _dpi), DipsToPixels(128.0f, _dpi));
    if (measured > 0.0f)
    {
        _pathLabelsMeasuredText[slot] = text;
        _pathLabelWidthsPx[slot]      = std::ceil(measured) + 1.0f;
    }
    if (Debug::Perf::IsCaptureEnabled())
        Debug::Perf::Emit(L"FileOps.Popup.PathLabels.MeasureUs", L"path label cache miss", PerfElapsedUs(startedUs), slot, _dpi, S_OK);
    return measured > 0.0f ? _pathLabelWidthsPx[slot] : 0.0f;
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureCaptionGlyphTextFormats(UINT dpi) noexcept
{
    if (! _dwriteFactory)
    {
        return;
    }

    if (_captionGlyphDpi != dpi)
    {
        _captionGlyphDpi = dpi;
        _captionGlyphFormat.reset();
        _captionGlyphFallbackFormat.reset();
    }

    if (! _captionGlyphFallbackFormat)
    {
        const HRESULT hr = DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), DxUi::Typography::MakeUiTextSpec(20.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _captionGlyphFallbackFormat.put(), L"");
        if (FAILED(hr))
        {
            _captionGlyphFallbackFormat.reset();
        }
    }

    if (! _captionGlyphFormat)
    {
        const HRESULT hr = DxUi::Typography::CreateTextFormat(_dwriteFactory.get(), DxUi::Typography::MakeUiIconSpec(20.0f), _captionGlyphFormat.put(), L"");
        if (FAILED(hr))
        {
            _captionGlyphFormat.reset();
        }
    }

    auto configure = [](IDWriteTextFormat* format) noexcept
    {
        if (! format)
        {
            return;
        }
        format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
        format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    };

    configure(_captionGlyphFormat.get());
    configure(_captionGlyphFallbackFormat.get());
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureTarget(HWND hwnd) noexcept
{
    EnsureFactories();
    bool primed = _controlsHost.PrimeForShow();
#ifdef ENABLE_TESTS
    // Stands in for a transient host failure, such as ResizeBuffers after a device reset.
    if (primed && ConsumeForcedFailure(g_fileOperationsD2DTargetForcedFailures))
        primed = false;
#endif
    if (! primed)
    {
        DiscardDeviceResources();
        return;
    }

    ID2D1DeviceContext* context = _controlsHost.GetDeviceContext();
    if (_target.get() == context)
    {
        return;
    }
    DiscardDeviceResources();

    RECT rc{};
    GetClientRect(hwnd, &rc);
    _clientSize.cx = std::max(0L, rc.right - rc.left);
    _clientSize.cy = std::max(0L, rc.bottom - rc.top);

    _dpi = GetDpiForWindow(hwnd);

    // All client pixels share the host's device and presentation. Keep the context alive
    // while preparing a frame; a host device change discards these dependent resources.
    _target = context;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::EnsureCaptionGlyphTarget(UINT dpi) noexcept
{
    EnsureFactories();
    EnsureCaptionGlyphTextFormats(dpi);
    if (! _d2dFactory)
    {
        return false;
    }

    if (! _captionGlyphTarget)
    {
        const D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        const HRESULT hr                          = _d2dFactory->CreateDCRenderTarget(&props, _captionGlyphTarget.addressof());
        if (FAILED(hr) || ! _captionGlyphTarget)
        {
            _captionGlyphTarget.reset();
            return false;
        }

        _captionGlyphTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }

    _captionGlyphTarget->SetDpi(static_cast<float>(dpi), static_cast<float>(dpi));

    if (! _captionGlyphBrush)
    {
        const HRESULT hr = _captionGlyphTarget->CreateSolidColorBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f), _captionGlyphBrush.addressof());
        if (FAILED(hr))
        {
            _captionGlyphBrush.reset();
            return false;
        }
    }

    return _captionGlyphTarget && _captionGlyphBrush;
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureBrushes() noexcept
{
    if (! _target)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    if (! folderWindow)
    {
        return;
    }

    const AppTheme& theme     = folderWindow->GetTheme();
    const D2D1::ColorF bg     = ColorFromCOLORREF(theme.windowBackground);
    const D2D1::ColorF fg     = ColorFromCOLORREF(theme.menu.text);
    const D2D1::ColorF sub    = ColorFromCOLORREF(theme.menu.disabledText);
    const D2D1::ColorF border = ColorFromCOLORREF(theme.menu.border);

    const D2D1::ColorF progressBg     = theme.fileOperations.progressBackground;
    const D2D1::ColorF progressGlobal = theme.fileOperations.progressTotal;
    const D2D1::ColorF progressItem   = theme.fileOperations.progressItem;
    const D2D1::ColorF progressVerify = theme.fileOperations.progressVerify;
    _progressItemBaseColor            = progressItem;

    const D2D1::ColorF okAccent    = theme.fileOperations.successText;
    const D2D1::ColorF warningText = theme.folderView.warningText;
    const D2D1::ColorF errorText   = theme.folderView.errorText;

    const D2D1::ColorF graphLine = theme.fileOperations.graphLine;

    if (! _bgBrush)
    {
        _target->CreateSolidColorBrush(bg, _bgBrush.addressof());
    }
    else
    {
        _bgBrush->SetColor(bg);
    }

    if (! _textBrush)
    {
        _target->CreateSolidColorBrush(fg, _textBrush.addressof());
    }
    else
    {
        _textBrush->SetColor(fg);
    }

    if (! _subTextBrush)
    {
        _target->CreateSolidColorBrush(sub, _subTextBrush.addressof());
    }
    else
    {
        _subTextBrush->SetColor(sub);
    }

    if (! _borderBrush)
    {
        _target->CreateSolidColorBrush(border, _borderBrush.addressof());
    }
    else
    {
        _borderBrush->SetColor(border);
    }

    if (! _progressBgBrush)
    {
        _target->CreateSolidColorBrush(progressBg, _progressBgBrush.addressof());
    }
    else
    {
        _progressBgBrush->SetColor(progressBg);
    }

    if (! _progressGlobalBrush)
    {
        _target->CreateSolidColorBrush(progressGlobal, _progressGlobalBrush.addressof());
    }
    else
    {
        _progressGlobalBrush->SetColor(progressGlobal);
    }

    if (! _progressItemBrush)
    {
        _target->CreateSolidColorBrush(progressItem, _progressItemBrush.addressof());
    }
    else
    {
        _progressItemBrush->SetColor(progressItem);
    }

    if (! _progressVerifyBrush)
    {
        _target->CreateSolidColorBrush(progressVerify, _progressVerifyBrush.addressof());
    }
    else
    {
        _progressVerifyBrush->SetColor(progressVerify);
    }

    if (! _statusOkBrush)
    {
        _target->CreateSolidColorBrush(okAccent, _statusOkBrush.addressof());
    }
    else
    {
        _statusOkBrush->SetColor(okAccent);
    }

    if (! _statusWarningBrush)
    {
        _target->CreateSolidColorBrush(warningText, _statusWarningBrush.addressof());
    }
    else
    {
        _statusWarningBrush->SetColor(warningText);
    }

    if (! _statusErrorBrush)
    {
        _target->CreateSolidColorBrush(errorText, _statusErrorBrush.addressof());
    }
    else
    {
        _statusErrorBrush->SetColor(errorText);
    }

    // The hosted ThroughputGraph paints itself; this brush colors the per-file mini bars only.
    const float graphFillAlpha   = theme.dark ? 0.22f : 0.18f;
    const D2D1::ColorF graphFill = D2D1::ColorF(graphLine.r, graphLine.g, graphLine.b, graphFillAlpha);
    if (! _graphDynamicBrush)
    {
        _target->CreateSolidColorBrush(graphFill, _graphDynamicBrush.addressof());
    }

    const D2D1::ColorF btnBg = ColorFromCOLORREF(theme.menu.background);

    if (! _buttonBgBrush)
    {
        _target->CreateSolidColorBrush(btnBg, _buttonBgBrush.addressof());
    }
    else
    {
        _buttonBgBrush->SetColor(btnBg);
    }

    if (! _buttonChromeBrush)
    {
        _target->CreateSolidColorBrush(btnBg, _buttonChromeBrush.addressof());
    }

    const D2D1::ColorF checkboxFill = ColorFromCOLORREF(theme.menu.selectionBg);
    if (! _checkboxFillBrush)
    {
        _target->CreateSolidColorBrush(checkboxFill, _checkboxFillBrush.addressof());
    }
    else
    {
        _checkboxFillBrush->SetColor(checkboxFill);
    }

    const D2D1::ColorF checkMark = ColorFromCOLORREF(theme.menu.selectionText);
    if (! _checkboxCheckBrush)
    {
        _target->CreateSolidColorBrush(checkMark, _checkboxCheckBrush.addressof());
    }
    else
    {
        _checkboxCheckBrush->SetColor(checkMark);
    }
}

std::vector<TaskSnapshot> FileOperationsPopupInternal::FileOperationsPopupState::BuildSnapshot() const
{
#ifdef ENABLE_TESTS
    ++_taskSnapshotBuildCount;
#endif
#ifdef ENABLE_TESTS
    if (debugVisualScenario.has_value())
    {
        return debugVisualScenario.value();
    }
#endif
    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    std::vector<FolderWindow::InformationalTaskUpdate> informationalTasks;
    std::vector<uint64_t> issueTaskIds;
    // x-perf#3: one shared, immutable projection per issues revision; no per-tick summary copies.
    std::shared_ptr<const std::vector<FolderWindow::FileOperationState::CompletedTaskPresentation>> completedPresentation;
    bool showInformationalNotes = false;
    if (fileOps && hostLifetime.lock())
    {
        fileOps->CollectTasks(tasks);
        fileOps->CollectInformationalTasks(informationalTasks);
        fileOps->CollectProjectableIssueTaskIds(issueTaskIds);
        completedPresentation  = fileOps->CollectCompletedTaskPresentations();
        showInformationalNotes = fileOps->GetShowInformationalNotes();
    }
    const std::vector<FolderWindow::FileOperationState::CompletedTaskPresentation> noCompletedTasks;
    const auto& completedTasks = completedPresentation ? *completedPresentation : noCompletedTasks;

    std::vector<TaskSnapshot> result;
    result.reserve(tasks.size() + completedTasks.size() + informationalTasks.size());
    std::unordered_map<uint64_t, bool> activeTaskIds;
    activeTaskIds.reserve(tasks.size());
    std::unordered_map<uint64_t, const FolderWindow::FileOperationState::CompletedTaskPresentation*> completedTaskById;
    completedTaskById.reserve(completedTasks.size());
    for (const auto& completed : completedTasks)
    {
        completedTaskById.emplace(completed.facts.taskId, &completed);
    }

    const auto projectCompletedTerminalCounts = [](TaskSnapshot& snap, const FolderWindow::FileOperationState::CompletedTaskFacts& completed) noexcept
    {
        snap.publishedItemCount               = completed.publishedItemCount;
        snap.unknownPublicationItemCount      = completed.unknownPublicationItemCount;
        snap.removedSourceCount               = completed.removedSourceCount;
        snap.retainedSourceCount              = completed.retainedSourceCount;
        snap.retainedSourceNativeCount        = completed.retainedSourceNativeCount;
        snap.canceledRetainedSourceCount      = completed.canceledRetainedSourceCount;
        snap.unknownSourceCount               = completed.unknownSourceCount;
        snap.missingSourceCount               = completed.missingSourceCount;
        snap.verifiedItemCount                = completed.verifiedItemCount;
        snap.verificationProblemItemCount     = completed.verificationProblemItemCount;
        snap.verificationFailedItemCount      = completed.verificationFailedItemCount;
        snap.verificationUnavailableItemCount = completed.verificationUnavailableItemCount;
        snap.verificationCanceledItemCount    = completed.verificationCanceledItemCount;
        snap.noteCount                        = completed.noteCount;
        snap.lastInfoNoteMessage              = completed.lastInfoNoteMessage;
        snap.skippedItemCount                 = completed.skippedItemCount;
        snap.partialExplainedByUserSkips      = completed.partialExplainedByUserSkips;
        snap.startedMutation                  = completed.startedMutation;
        snap.completedSummaryProjected        = true;
    };

    std::vector<FolderWindow::FileOperationState::Task*> waitingTasks;
    waitingTasks.reserve(tasks.size());
    for (auto* task : tasks)
    {
        if (task && ! task->HasStarted() && (task->IsWaitingForOthers() || task->IsWaitingInQueue()))
        {
            waitingTasks.push_back(task);
        }
    }
    std::ranges::sort(waitingTasks, {}, &FolderWindow::FileOperationState::Task::GetQueueOrderKey);
    std::unordered_map<uint64_t, size_t> waitingTaskIndex;
    waitingTaskIndex.reserve(waitingTasks.size());
    for (size_t index = 0u; index < waitingTasks.size(); ++index)
    {
        waitingTaskIndex.emplace(waitingTasks[index]->GetId(), index);
    }
    std::ranges::stable_sort(tasks,
                             [](const auto* left, const auto* right)
    {
        if (! left || ! right)
        {
            return right != nullptr;
        }
        const bool leftWaiting  = ! left->HasStarted() && (left->IsWaitingForOthers() || left->IsWaitingInQueue());
        const bool rightWaiting = ! right->HasStarted() && (right->IsWaitingForOthers() || right->IsWaitingInQueue());
        const uint64_t leftKey  = leftWaiting ? left->GetQueueOrderKey() : left->GetId();
        const uint64_t rightKey = rightWaiting ? right->GetQueueOrderKey() : right->GetId();
        return leftKey < rightKey;
    });

    for (const auto& info : informationalTasks)
    {
        if (info.taskId == 0)
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.kind                  = TaskSnapshot::Kind::Informational;
        snap.taskId                = info.taskId;
        activeTaskIds[snap.taskId] = true;
        snap.informational         = info;
        snap.started               = true;
        snap.finished              = info.finished;
        snap.resultHr              = info.resultHr;
        snap.statusKind            = ResolveTaskStatusKind(snap);

        result.push_back(std::move(snap));
    }

    for (auto* task : tasks)
    {
        if (! task || task->IsPresentationHidden())
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.taskId                = task->GetId();
        activeTaskIds[snap.taskId] = true;
        snap.operation             = task->GetOperation();
        snap.lifecyclePhase        = static_cast<uint8_t>(task->GetLifecyclePhase());

        snap.totalItems                     = task->_publishedProgressTotalItems.load(std::memory_order_acquire);
        snap.completedItems                 = task->_publishedProgressCompletedItems.load(std::memory_order_acquire);
        snap.totalBytes                     = task->_publishedProgressTotalBytes.load(std::memory_order_acquire);
        snap.completedBytes                 = task->_publishedProgressCompletedBytes.load(std::memory_order_acquire);
        snap.itemTotalBytes                 = task->_publishedProgressItemTotalBytes.load(std::memory_order_acquire);
        snap.itemCompletedBytes             = task->_publishedProgressItemCompletedBytes.load(std::memory_order_acquire);
        snap.verificationRequested          = task->_verificationRequested.load(std::memory_order_acquire);
        snap.verificationActive             = task->_verificationActive.load(std::memory_order_acquire);
        snap.verificationTotalBytes         = task->_verificationTotalBytes.load(std::memory_order_acquire);
        snap.verificationCompletedBytes     = task->_verificationCompletedBytes.load(std::memory_order_acquire);
        snap.verificationItemTotalBytes     = task->_verificationItemTotalBytes.load(std::memory_order_acquire);
        snap.verificationItemCompletedBytes = task->_verificationItemCompletedBytes.load(std::memory_order_acquire);
        snap.completedFiles                 = task->_publishedCompletedTopLevelFiles.load(std::memory_order_acquire);
        snap.completedFolders               = task->_publishedCompletedTopLevelFolders.load(std::memory_order_acquire);

        {
            std::scoped_lock lock(task->_progressPathMutex);
            snap.currentSourcePath = ! task->_lastProgressCallbackSourcePath.empty() ? task->_lastProgressCallbackSourcePath : task->_progressSourcePath;
            snap.currentDestinationPath =
                ! task->_lastProgressCallbackDestinationPath.empty() ? task->_lastProgressCallbackDestinationPath : task->_progressDestinationPath;
            snap.hasProgressCallbacks     = task->_lastProgressCallbackTick != 0;
            snap.lastProgressCallbackTick = task->_lastProgressCallbackTick;
        }

        {
            std::scoped_lock lock(task->_inFlightFilesMutex);
            snap.discoveryClosed          = task->_discoveryClosed.load(std::memory_order_acquire);
            snap.discoveredFileCount      = task->_discoveredFileCount.load(std::memory_order_acquire);
            snap.discoveredDirectoryCount = task->_discoveredDirectoryCount.load(std::memory_order_acquire);
            snap.discoveredTotalBytes     = task->_discoveredTotalBytes.load(std::memory_order_acquire);
            snap.copiedFileCount          = task->_copiedFileCount.load(std::memory_order_acquire);
            snap.copiedFileCountKnown     = task->_copiedFileCountKnown.load(std::memory_order_acquire);
            snap.copiedFolderCount        = task->_copiedFolderCount.load(std::memory_order_acquire);
            snap.copiedFolderCountKnown   = task->_copiedFolderCountKnown.load(std::memory_order_acquire);
            snap.completedWorkFileCount   = task->_completedWorkFileCount.load(std::memory_order_acquire);
            snap.completedWorkFolderCount = task->_completedWorkFolderCount.load(std::memory_order_acquire);
            snap.completedWorkBytes       = task->_completedWorkBytes.load(std::memory_order_acquire);
            snap.completedWorkCountsKnown = task->_completedWorkCountsKnown.load(std::memory_order_acquire);
            if (task->_operation == FILESYSTEM_DELETE && snap.completedWorkCountsKnown)
            {
                snap.completedBytes = std::max(snap.completedBytes, snap.completedWorkBytes);
                const uint64_t workItems =
                    std::min<uint64_t>(snap.completedWorkFileCount, ULONG_MAX) + std::min<uint64_t>(snap.completedWorkFolderCount, ULONG_MAX);
                snap.completedItems = std::max(snap.completedItems, static_cast<unsigned long>(std::min<uint64_t>(workItems, ULONG_MAX)));
            }
            snap.discoveredBytesKnown = task->_discoveredBytesKnown.load(std::memory_order_acquire);
            snap.exactDiscoveredBytes = task->_fileWorkBytesKnownScopes == task->_fileWorkScopes.size() && ! task->_fileWorkScopes.empty()
                                            ? snap.discoveredTotalBytes
                                            : task->_exactDiscoveredBytes.load(std::memory_order_acquire);
            snap.inFlightFileCount    = std::min(task->_inFlightFileCount, snap.inFlightFiles.size());
            for (size_t i = 0; i < snap.inFlightFileCount; ++i)
            {
                snap.inFlightFiles[i].cookieKey         = task->_inFlightFiles[i].cookieKey;
                snap.inFlightFiles[i].progressStreamId  = task->_inFlightFiles[i].progressStreamId;
                snap.inFlightFiles[i].sourcePath        = task->_inFlightFiles[i].sourcePath;
                snap.inFlightFiles[i].totalBytes        = task->_inFlightFiles[i].totalBytes;
                snap.inFlightFiles[i].completedBytes    = task->_inFlightFiles[i].completedBytes;
                snap.inFlightFiles[i].logicalLeafId     = task->_inFlightFiles[i].logicalLeafId;
                snap.inFlightFiles[i].phase             = task->_inFlightFiles[i].phase;
                snap.inFlightFiles[i].explicitLifecycle = task->_inFlightFiles[i].explicitLifecycle;
                snap.inFlightFiles[i].sizeKnown         = task->_inFlightFiles[i].sizeKnown;
                snap.inFlightFiles[i].terminal          = task->_inFlightFiles[i].terminal;

                // Defensive: for display purposes, avoid showing a misleading "100%" when a plugin reports
                // currentItemCompletedBytes > currentItemTotalBytes (can happen with out-of-order updates or bugs).
                if (snap.inFlightFiles[i].totalBytes > 0 && snap.inFlightFiles[i].completedBytes > snap.inFlightFiles[i].totalBytes)
                {
                    constexpr uint64_t kClampThresholdBytes = 64ull * 1024ull;
                    const uint64_t delta                    = snap.inFlightFiles[i].completedBytes - snap.inFlightFiles[i].totalBytes;
                    if (delta <= kClampThresholdBytes)
                    {
                        snap.inFlightFiles[i].completedBytes = snap.inFlightFiles[i].totalBytes;
                    }
                    else
                    {
                        // Unknown/invalid totals: render as indeterminate.
                        snap.inFlightFiles[i].totalBytes     = 0;
                        snap.inFlightFiles[i].completedBytes = 0;
                    }
                }
                snap.inFlightFiles[i].lastUpdateTick = task->_inFlightFiles[i].lastUpdateTick;
            }
        }

        {
            std::scoped_lock lock(task->_conflictArbiter.mutex);
            snap.conflict.active                            = task->_conflictArbiter.prompt.active;
            snap.conflict.inspection                        = task->_conflictArbiter.prompt.pathDisplay.inspection;
            snap.conflict.metadataLoading                   = task->_conflictArbiter.prompt.metadataLoading;
            snap.conflict.deferredConsent                   = task->_conflictArbiter.prompt.deferredConsent;
            snap.conflict.bucket                            = static_cast<uint8_t>(task->_conflictArbiter.prompt.bucket);
            snap.conflict.status                            = task->_conflictArbiter.prompt.status;
            snap.conflict.sourcePath                        = task->_conflictArbiter.prompt.sourcePath;
            snap.conflict.destinationPath                   = task->_conflictArbiter.prompt.destinationPath;
            snap.conflict.sourceName                        = task->_conflictArbiter.prompt.pathDisplay.sourceName;
            snap.conflict.sourceParentPath                  = task->_conflictArbiter.prompt.pathDisplay.sourceParentPath;
            snap.conflict.destinationName                   = task->_conflictArbiter.prompt.pathDisplay.destinationName;
            snap.conflict.destinationParentPath             = task->_conflictArbiter.prompt.pathDisplay.destinationParentPath;
            snap.conflict.sourceMetadata.available          = task->_conflictArbiter.prompt.sourceMetadata.available;
            snap.conflict.sourceMetadata.isDirectory        = task->_conflictArbiter.prompt.sourceMetadata.isDirectory;
            snap.conflict.sourceMetadata.isLink             = task->_conflictArbiter.prompt.sourceMetadata.isLink;
            snap.conflict.sourceMetadata.sizeKnown          = task->_conflictArbiter.prompt.sourceMetadata.sizeKnown;
            snap.conflict.sourceMetadata.sizeBytes          = task->_conflictArbiter.prompt.sourceMetadata.sizeBytes;
            snap.conflict.sourceMetadata.lastWriteTime      = task->_conflictArbiter.prompt.sourceMetadata.lastWriteTime;
            snap.conflict.sourceMetadata.attributes         = task->_conflictArbiter.prompt.sourceMetadata.attributes;
            snap.conflict.destinationMetadata.available     = task->_conflictArbiter.prompt.destinationMetadata.available;
            snap.conflict.destinationMetadata.isDirectory   = task->_conflictArbiter.prompt.destinationMetadata.isDirectory;
            snap.conflict.destinationMetadata.isLink        = task->_conflictArbiter.prompt.destinationMetadata.isLink;
            snap.conflict.destinationMetadata.sizeKnown     = task->_conflictArbiter.prompt.destinationMetadata.sizeKnown;
            snap.conflict.destinationMetadata.sizeBytes     = task->_conflictArbiter.prompt.destinationMetadata.sizeBytes;
            snap.conflict.destinationMetadata.lastWriteTime = task->_conflictArbiter.prompt.destinationMetadata.lastWriteTime;
            snap.conflict.destinationMetadata.attributes    = task->_conflictArbiter.prompt.destinationMetadata.attributes;
            CopyPublishedConflictPolicy(task->_conflictArbiter.prompt, snap.conflict);
            snap.conflict.applyToAllChecked  = task->_conflictArbiter.prompt.applyToAllChecked;
            snap.conflict.retryFailed        = task->_conflictArbiter.prompt.retryFailed;
            snap.conflict.attemptCount       = task->_conflictArbiter.prompt.attemptCount;
            snap.conflict.factItemCountKnown = task->_conflictArbiter.prompt.factItemCountKnown;
            snap.conflict.factItemCount      = task->_conflictArbiter.prompt.factItemCount;
            snap.conflict.consentDetail      = task->_conflictArbiter.prompt.consentDetail;
            snap.conflict.consentNote        = task->_conflictArbiter.prompt.consentNote;
            snap.conflict.factBytesKnown     = task->_conflictArbiter.prompt.factBytesKnown;
            snap.conflict.factBytes          = task->_conflictArbiter.prompt.factBytes;
            snap.conflict.overlapProblem     = task->_conflictArbiter.prompt.overlapProblem;
            snap.conflict.overlapTaskId      = task->_conflictArbiter.prompt.overlapTaskId;
        }

        snap.started          = task->HasStarted();
        snap.paused           = task->IsPaused();
        snap.waitingForOthers = task->IsWaitingForOthers();
        snap.waitingInQueue   = task->IsWaitingInQueue();
        snap.overlapQueued    = task->IsOverlapQueued();
        snap.queuePaused      = task->IsQueuePaused();
        snap.queueOrderKey    = task->GetQueueOrderKey();
        if (const auto waitingIt = waitingTaskIndex.find(snap.taskId); waitingIt != waitingTaskIndex.end())
        {
            snap.canMoveQueueUp   = waitingIt->second > 0u;
            snap.canMoveQueueDown = waitingIt->second + 1u < waitingTasks.size();
        }
        const auto selectedAdmission    = task->LoadSelectedAdmission();
        snap.plannedItems               = task->GetPlannedItemCount();
        snap.selection                  = task->_selection;
        snap.destinationFolder          = selectedAdmission->destination->folder;
        snap.destinationPane            = selectedAdmission->destination->pane;
        snap.destinationPluginId        = selectedAdmission->destination->pluginId;
        snap.destinationPluginShortId   = selectedAdmission->destination->pluginShortId;
        snap.destinationInstanceContext = selectedAdmission->destination->instanceContext;
        snap.operationStartTick         = task->_operationStartTick.load(std::memory_order_acquire);
        snap.clipboardMoveAdmission     = task->_clipboardMoveAdmission;
        snap.clipboardMoveConsumed      = task->_clipboardMoveConsumed.load(std::memory_order_acquire);

        // A task whose thread already completed renders its final status immediately instead of
        // flashing "Running" until the completed-summary row replaces this live row.
        if (task->_taskFinished.load(std::memory_order_acquire))
        {
            snap.finished          = true;
            snap.resultHr          = task->_resultHr.load(std::memory_order_acquire);
            const auto completedIt = completedTaskById.find(snap.taskId);
            if (completedIt != completedTaskById.end() && completedIt->second)
            {
                const auto& presentation      = *completedIt->second;
                const auto& completed         = presentation.facts;
                snap.warningCount             = completed.warningCount;
                snap.errorCount               = completed.errorCount;
                snap.lastDiagnosticMessage    = completed.lastDiagnosticMessage;
                snap.resultSummary            = completed.resultSummary;
                snap.copiedFileCount          = completed.copiedFileCount;
                snap.copiedFileCountKnown     = completed.copiedFileCountKnown;
                snap.copiedFolderCount        = completed.copiedFolderCount;
                snap.copiedFolderCountKnown   = completed.copiedFolderCountKnown;
                snap.completedWorkFileCount   = completed.completedWorkFileCount;
                snap.completedWorkFolderCount = completed.completedWorkFolderCount;
                snap.completedWorkBytes       = completed.completedWorkBytes;
                snap.completedWorkCountsKnown = completed.completedWorkCountsKnown;
                projectCompletedTerminalCounts(snap, completed);
                snap.verificationState = static_cast<uint8_t>(completed.verificationFailedItemCount > 0u        ? FileOperations::VerificationState::Failed
                                                              : completed.verificationCanceledItemCount > 0u    ? FileOperations::VerificationState::Canceled
                                                              : completed.verificationUnavailableItemCount > 0u ? FileOperations::VerificationState::Unavailable
                                                              : completed.verifiedItemCount > 0u ? FileOperations::VerificationState::Verified
                                                                                                 : FileOperations::VerificationState::NotRequested);
                snap.hasIndeterminateResult             = completed.indeterminateItemCount > 0u;
                snap.interruptedMoveNotice              = completed.interruptedMoveNotice;
                snap.interruptedOperationId             = completed.interruptedOperationId;
                snap.clipboardMoveAdmission             = completed.clipboardMoveAdmission;
                snap.clipboardMoveConsumed              = completed.clipboardMoveConsumed;
                snap.hasSourceActionPaths               = presentation.hasSourceActionPaths;
                snap.hasRetainedDiagnostics             = presentation.hasRetainedDiagnostics;
                snap.exactRetainedSourceActionAvailable = completed.clipboardMoveAdmission &&
                                                          NavigationLocation::IsFilePluginShortId(completed.sourcePluginShortId) &&
                                                          completed.retainedSourceCount > 0u && completed.unknownSourceCount == 0u &&
                                                          presentation.exactRetainedSourceItemCount == completed.retainedSourceCount;
            }
            else
            {
                fileOps->CollectTaskDiagnosticSnapshot(
                    snap.taskId, snap.warningCount, snap.errorCount, snap.lastDiagnosticMessage, &snap.noteCount, &snap.lastInfoNoteMessage);
                snap.hasRetainedDiagnostics = snap.warningCount > 0u || snap.errorCount > 0u || snap.noteCount > 0u;
            }
        }
        if (! snap.finished && snap.clipboardMoveConsumed && snap.lastDiagnosticMessage.empty())
        {
            snap.lastDiagnosticMessage = LoadStringResource(nullptr, IDS_FILEOPS_CLIPBOARD_MOVE_QUEUED);
        }

        snap.desiredSpeedLimitBytesPerSecond       = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        snap.effectiveSpeedLimitBytesPerSecond     = task->_effectiveSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        snap.autoConcurrencyUsed                   = task->_autoConcurrencyUsed.load(std::memory_order_acquire);
        snap.autoConcurrencyStorageKind            = task->_autoConcurrencyStorageKind.load(std::memory_order_acquire);
        snap.autoConcurrencyDestinationStorageKind = task->_autoConcurrencyDestinationStorageKind.load(std::memory_order_acquire);
        snap.autoTunedConcurrency                  = task->_autoTunedConcurrency.load(std::memory_order_acquire);
        snap.effectiveConcurrencyBudget            = task->_effectiveConcurrencyBudget.load(std::memory_order_acquire);

        // Discovery-ahead state
        snap.discoveryAheadActive               = task->_discoveryAheadActive.load(std::memory_order_acquire);
        snap.discoverySkipped                   = task->_discoverySkipped.load(std::memory_order_acquire);
        snap.firstMutationBeforeDiscoveryClosed = task->_firstMutationBeforeDiscoveryClosed.load(std::memory_order_acquire);

        const ULONGLONG startTick = task->_discoveryStartTick.load(std::memory_order_acquire);
        if (! snap.discoveryClosed && startTick > 0)
        {
            const ULONGLONG nowTick = GetTickCount64();
            snap.discoveryElapsedMs = (nowTick >= startTick) ? (nowTick - startTick) : 0;
        }

        if (snap.discoveredBytesKnown)
        {
            snap.discoveredTotalBytes = snap.exactDiscoveredBytes;
            snap.totalBytes           = snap.exactDiscoveredBytes;
        }

        PublishPlannedItemTotalAfterDiscovery(snap);

        if (snap.totalItems > 0)
        {
            snap.completedItems = std::min(snap.completedItems, snap.totalItems);
        }
        if (snap.totalBytes > 0)
        {
            snap.completedBytes = std::min(snap.completedBytes, snap.totalBytes);
        }
        if (snap.itemTotalBytes > 0)
        {
            snap.itemCompletedBytes = std::min(snap.itemCompletedBytes, snap.itemTotalBytes);
        }
        snap.showInformationalNotes = showInformationalNotes;
        snap.hasProjectableIssues = snap.finished && ! snap.interruptedMoveNotice && std::binary_search(issueTaskIds.begin(), issueTaskIds.end(), snap.taskId);
        EnsureFinishedTaskDiagnosticAffordance(snap);
        NormalizeCompletedTaskSnapshotForDisplay(snap);
        snap.cancellationRequested = task->CancelRequestedDurable();
        ProjectTaskCancellationForDisplay(snap);

        result.push_back(std::move(snap));
    }

    for (const auto& presentation : completedTasks)
    {
        const auto& completed = presentation.facts;
        if (activeTaskIds.find(completed.taskId) != activeTaskIds.end())
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.taskId                                = completed.taskId;
        snap.selection                             = completed.selection;
        snap.operation                             = completed.operation;
        snap.totalItems                            = completed.totalItems;
        snap.completedItems                        = completed.completedItems;
        snap.totalBytes                            = completed.totalBytes;
        snap.completedBytes                        = completed.completedBytes;
        snap.completedFiles                        = completed.completedFiles;
        snap.completedFolders                      = completed.completedFolders;
        snap.copiedFileCount                       = completed.copiedFileCount;
        snap.copiedFileCountKnown                  = completed.copiedFileCountKnown;
        snap.copiedFolderCount                     = completed.copiedFolderCount;
        snap.copiedFolderCountKnown                = completed.copiedFolderCountKnown;
        snap.completedWorkFileCount                = completed.completedWorkFileCount;
        snap.completedWorkFolderCount              = completed.completedWorkFolderCount;
        snap.completedWorkBytes                    = completed.completedWorkBytes;
        snap.completedWorkCountsKnown              = completed.completedWorkCountsKnown;
        snap.currentSourcePath                     = completed.sourcePath;
        snap.currentDestinationPath                = completed.destinationPath;
        snap.destinationFolder                     = completed.destinationFolder;
        snap.destinationPane                       = completed.destinationPane;
        snap.destinationPluginId                   = completed.destinationPluginId;
        snap.destinationPluginShortId              = completed.destinationPluginShortId;
        snap.destinationInstanceContext            = completed.destinationInstanceContext;
        snap.started                               = true;
        snap.finished                              = true;
        snap.autoConcurrencyUsed                   = completed.autoConcurrencyUsed;
        snap.autoConcurrencyStorageKind            = completed.autoConcurrencyStorageKind;
        snap.autoConcurrencyDestinationStorageKind = completed.autoConcurrencyDestinationStorageKind;
        snap.autoTunedConcurrency                  = completed.autoTunedConcurrency;
        snap.effectiveConcurrencyBudget            = completed.effectiveConcurrencyBudget;
        snap.resultHr                              = completed.resultHr;
        snap.warningCount                          = completed.warningCount;
        snap.errorCount                            = completed.errorCount;
        snap.lastDiagnosticMessage                 = completed.lastDiagnosticMessage;
        snap.resultSummary                         = completed.resultSummary;
        snap.verificationRequested                 = completed.verifiedItemCount > 0u || completed.verificationProblemItemCount > 0u;
        snap.verificationTotalBytes                = completed.verificationTotalBytes;
        snap.verificationCompletedBytes            = completed.verificationCompletedBytes;
        projectCompletedTerminalCounts(snap, completed);
        snap.verificationState      = static_cast<uint8_t>(completed.verificationFailedItemCount > 0u        ? FileOperations::VerificationState::Failed
                                                           : completed.verificationCanceledItemCount > 0u    ? FileOperations::VerificationState::Canceled
                                                           : completed.verificationUnavailableItemCount > 0u ? FileOperations::VerificationState::Unavailable
                                                           : completed.verifiedItemCount > 0u                ? FileOperations::VerificationState::Verified
                                                                                                             : FileOperations::VerificationState::NotRequested);
        snap.hasIndeterminateResult = completed.indeterminateItemCount > 0u;
        snap.interruptedMoveNotice  = completed.interruptedMoveNotice;
        snap.interruptedOperationId = completed.interruptedOperationId;
        snap.clipboardMoveAdmission = completed.clipboardMoveAdmission;
        snap.clipboardMoveConsumed  = completed.clipboardMoveConsumed;
        snap.hasSourceActionPaths   = presentation.hasSourceActionPaths;
        snap.hasRetainedDiagnostics = presentation.hasRetainedDiagnostics;
        snap.exactRetainedSourceActionAvailable = completed.clipboardMoveAdmission && NavigationLocation::IsFilePluginShortId(completed.sourcePluginShortId) &&
                                                  completed.retainedSourceCount > 0u && completed.unknownSourceCount == 0u &&
                                                  presentation.exactRetainedSourceItemCount == completed.retainedSourceCount;
        snap.discoverySkipped                   = completed.discoverySkipped;
        snap.discoveryClosed                    = completed.discoveryClosed;
        snap.hasProgressCallbacks               = completed.lastProgressCallbackTick != 0;
        snap.lastProgressCallbackTick           = completed.lastProgressCallbackTick;

        if (snap.totalItems > 0)
        {
            snap.completedItems = std::min(snap.completedItems, snap.totalItems);
        }
        if (snap.totalBytes > 0)
        {
            snap.completedBytes = std::min(snap.completedBytes, snap.totalBytes);
        }
        snap.showInformationalNotes = showInformationalNotes;
        snap.hasProjectableIssues = snap.finished && ! snap.interruptedMoveNotice && std::binary_search(issueTaskIds.begin(), issueTaskIds.end(), snap.taskId);
        EnsureFinishedTaskDiagnosticAffordance(snap);
        NormalizeCompletedTaskSnapshotForDisplay(snap);
        snap.statusKind = ResolveTaskStatusKind(snap);

        result.push_back(std::move(snap));
    }

    return result;
}

#ifdef ENABLE_TESTS
void FileOperationsPopupInternal::FileOperationsPopupState::DebugSetVisualScenario(std::vector<TaskSnapshot> tasks, bool collapsed, bool completedGroupExpanded)
{
    _collapsedTasks.clear();
    _compactExpandedTasks.clear();
    _taskHeightAnimations.clear();
    _rates.clear();
    _completedGroupExpanded     = completedGroupExpanded;
    _maxAutoSizedWindowHeight   = 0;
    _lastAutoSizedContentHeight = -1.0f;
    for (auto& task : tasks)
    {
        ProjectTaskCancellationForDisplay(task);
        _collapsedTasks[task.taskId] = collapsed;
        if (task.conflict.active && ! task.conflict.metadataLoading)
        {
            FolderWindow::FileOperationState::Task::ConflictPromptState prompt{};
            prompt.bucket                         = static_cast<FileOperations::ConflictClass>(task.conflict.bucket);
            prompt.generation                     = task.conflict.generation;
            prompt.repeatUnavailableReason        = static_cast<FileOperations::ConflictRepeatUnavailableReason>(task.conflict.repeatUnavailableReason);
            prompt.sourceMetadata.objectKindKnown = true;
            prompt.sourceMetadata.isDirectory     = task.conflict.sourceIsDirectory || task.conflict.sourceMetadata.isDirectory;
            prompt.decisionScope.sourceKind =
                prompt.sourceMetadata.isDirectory ? FileOperations::ConflictItemKind::Directory : FileOperations::ConflictItemKind::RegularFile;
            const bool readOnlyDelete = task.operation == FILESYSTEM_DELETE && prompt.bucket == FileOperations::ConflictClass::ReadOnlyRegularFile;
            DebugBuildFileOpsVisualConflictPolicyForSelfTest(prompt, task.conflict.destinationMetadata.available, readOnlyDelete, task.operation);
            CopyPublishedConflictPolicy(prompt, task.conflict);
        }
        auto& history                            = _rates[task.taskId];
        history.initialized                      = true;
        history.count                            = 120u;
        history.writeIndex                       = 120u;
        history.displayedBytesPerSec             = task.paused || task.conflict.active || task.finished || ! task.started || task.verificationActive ||
                                                           task.operation == FILESYSTEM_DELETE || task.totalBytes == 0u
                                                       ? 0.0
                                                       : 64.0 * 1024.0 * 1024.0;
        history.displayedItemsPerSec             = task.operation == FILESYSTEM_DELETE ? 24.0 : 0.0;
        history.displayedVerificationBytesPerSec = task.verificationActive ? 128.0 * 1024.0 * 1024.0 : 0.0;
        history.smoothedEtaSeconds               = task.verificationActive ? 32.0 : 80.0;
        history.hasSmoothedEta                   = task.started && ! task.finished && ! task.paused && task.discoveryClosed;
        history.provisionalEtaSeconds            = 80.0;
        history.hasProvisionalEta                = task.started && ! task.discoveryClosed && ! task.paused;
        history.streamProgressCount              = task.inFlightFileCount;
        for (size_t i = 0u; i < task.inFlightFileCount; ++i)
        {
            auto& stream             = history.streamProgress[i];
            stream.cookieKey         = task.inFlightFiles[i].cookieKey;
            stream.progressStreamId  = task.inFlightFiles[i].progressStreamId;
            stream.sourcePath        = task.inFlightFiles[i].sourcePath;
            stream.assignedColorSlot = static_cast<uint8_t>(i);
            stream.assignedHue       = static_cast<float>(i) * 137.5f;
        }
        for (size_t i = 0u; i < history.count; ++i)
        {
            history.samples[i] =
                task.operation == FILESYSTEM_DELETE || task.totalBytes == 0u || ! task.started ? 0.0f : static_cast<float>((48u + i % 23u) * 1024u * 1024u);
            history.verificationSamples[i] = task.verificationActive ? history.samples[i] * 2.0f : 0.0f;
            history.hues[i]                = 210.0f;
            history.hueWeightCounts[i]     = static_cast<uint8_t>(task.inFlightFileCount);
            for (size_t j = 0u; j < task.inFlightFileCount; ++j)
            {
                history.hueWeights[i][j] = {history.streamProgress[j].assignedHue, 1.0, static_cast<uint8_t>(j)};
            }
            // I35 R10.2: every running transfer fixture shows one short past stall band; the task is
            // not stalled now, so its rate label carries no stall text.
            history.stalledSamples[i] = history.displayedBytesPerSec > 0.0 && i >= 96u && i < 104u ? 1u : 0u;
        }
    }
    // A replaced fixture is a new state, not live text: publish it at the next synchronization.
    _hostedAccessibleTextFlush = true;
    debugVisualScenario        = std::move(tasks);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DebugRunIdleTimerTickForSelfTest(HWND hwnd) noexcept
{
    _hostedAccessibleTextPending = false;
    _uiMotionTimerDueTick        = 0u;
    static_cast<void>(OnTimer(hwnd, kFileOperationsPopupTimerId));
}
#endif

std::vector<RateSnapshot> FileOperationsPopupInternal::FileOperationsPopupState::BuildRateSnapshot() const
{
#ifdef ENABLE_TESTS
    ++_rateSnapshotBuildCount;
#endif
    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    std::shared_ptr<const std::vector<FolderWindow::FileOperationState::CompletedTaskPresentation>> completedPresentation;
    if (fileOps && hostLifetime.lock())
    {
        fileOps->CollectTasks(tasks);
        completedPresentation = fileOps->CollectCompletedTaskPresentations();
    }
    const std::vector<FolderWindow::FileOperationState::CompletedTaskPresentation> noCompletedTasks;
    const auto& completedTasks = completedPresentation ? *completedPresentation : noCompletedTasks;

    std::unordered_map<uint64_t, bool> activeTaskIds;
    activeTaskIds.reserve(tasks.size());

    std::vector<RateSnapshot> result;
    result.reserve(tasks.size() + completedTasks.size());

    for (auto* task : tasks)
    {
        if (! task || task->IsPresentationHidden())
        {
            continue;
        }

        RateSnapshot snap{};
        snap.taskId                = task->GetId();
        snap.operation             = task->GetOperation();
        activeTaskIds[snap.taskId] = true;

        snap.completedItems             = task->_publishedProgressCompletedItems.load(std::memory_order_acquire);
        snap.totalBytes                 = task->_publishedProgressTotalBytes.load(std::memory_order_acquire);
        snap.completedBytes             = task->_publishedProgressCompletedBytes.load(std::memory_order_acquire);
        snap.verificationTotalBytes     = task->_verificationTotalBytes.load(std::memory_order_acquire);
        snap.verificationCompletedBytes = task->_verificationCompletedBytes.load(std::memory_order_acquire);
        snap.verificationReadBytes      = task->_perf.verificationReadBytes.load(std::memory_order_acquire);

        {
            std::scoped_lock lock(task->_progressPathMutex);
            snap.currentSourcePath        = ! task->_lastProgressCallbackSourcePath.empty() ? task->_lastProgressCallbackSourcePath : task->_progressSourcePath;
            snap.lastProgressCallbackTick = task->_lastProgressCallbackTick;
        }

        {
            std::scoped_lock lock(task->_inFlightFilesMutex);
            snap.discoveryClosed      = task->_discoveryClosed.load(std::memory_order_acquire);
            snap.discoveredTotalBytes = task->_discoveredTotalBytes.load(std::memory_order_acquire);
            snap.discoveredBytesKnown = task->_discoveredBytesKnown.load(std::memory_order_acquire);
            if (snap.discoveredBytesKnown)
            {
                if (task->_fileWorkBytesKnownScopes != task->_fileWorkScopes.size())
                    snap.discoveredTotalBytes = task->_exactDiscoveredBytes.load(std::memory_order_acquire);
                snap.totalBytes = snap.discoveredTotalBytes;
            }
            else if (task->_operation == FILESYSTEM_COPY)
                snap.totalBytes = 0u;
            snap.inFlightFileCount = std::min(task->_inFlightFileCount, snap.inFlightFiles.size());
            for (size_t i = 0; i < snap.inFlightFileCount; ++i)
            {
                snap.inFlightFiles[i].cookieKey        = task->_inFlightFiles[i].cookieKey;
                snap.inFlightFiles[i].progressStreamId = task->_inFlightFiles[i].progressStreamId;
                snap.inFlightFiles[i].sourcePath       = task->_inFlightFiles[i].sourcePath;
                snap.inFlightFiles[i].totalBytes       = task->_inFlightFiles[i].totalBytes;
                snap.inFlightFiles[i].completedBytes   = task->_inFlightFiles[i].completedBytes;
                snap.inFlightFiles[i].lastUpdateTick   = task->_inFlightFiles[i].lastUpdateTick;
            }
        }

        snap.progressStateChangeTick = task->_rateSamplingStateChangeTick.load(std::memory_order_acquire);
        snap.started                 = task->HasStarted();
        snap.paused                  = task->IsPaused();
        snap.waitingForOthers        = task->IsWaitingForOthers();
        snap.waitingInQueue          = task->IsWaitingInQueue();
        snap.queuePaused             = task->IsQueuePaused();
        snap.lastByteProgressTick    = task->_lastByteProgressTick.load(std::memory_order_acquire);
        snap.stallKind = RateStallKindFor(task->_progressStallPhase.load(std::memory_order_acquire), task->_progressStallSide.load(std::memory_order_acquire));
        snap.speedLimitBytesPerSecond = task->_effectiveSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        snap.stopping                 = task->CancelRequestedDurable();
        {
            std::scoped_lock lock(task->_conflictArbiter.mutex);
            snap.waitingForDecision = task->_conflictArbiter.prompt.active;
        }

        result.push_back(snap);
    }

    for (const auto& presentation : completedTasks)
    {
        const auto& completed = presentation.facts;
        if (activeTaskIds.find(completed.taskId) != activeTaskIds.end())
        {
            continue;
        }

        RateSnapshot snap{};
        snap.taskId                   = completed.taskId;
        snap.operation                = completed.operation;
        snap.completedItems           = completed.completedItems;
        snap.totalBytes               = completed.totalBytes;
        snap.completedBytes           = completed.completedBytes;
        snap.currentSourcePath        = completed.sourcePath;
        snap.lastProgressCallbackTick = completed.lastProgressCallbackTick;
        snap.started                  = true;
        snap.discoveryClosed          = completed.discoveryClosed;
        snap.finished                 = true;

        result.push_back(std::move(snap));
    }

    return result;
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateRates() noexcept
{
#ifdef ENABLE_TESTS
    ++_rateUpdateCount;
    if (debugVisualScenario.has_value())
    {
        return;
    }
#endif
    const bool capturePerf                   = Debug::Perf::IsCaptureEnabled();
    const uint64_t perfStartUs               = capturePerf ? PerfNowUs() : 0u;
    const ULONGLONG nowTick                  = GetTickCount64();
    const std::vector<RateSnapshot> snapshot = BuildRateSnapshot();

    std::unordered_map<uint64_t, bool> seen;
    seen.reserve(snapshot.size());

    RateUpdateTotals totals{};
    for (const RateSnapshot& task : snapshot)
    {
        seen[task.taskId] = true;
        AdvanceRateHistory(_rates[task.taskId], task, nowTick, totals);
    }

    for (auto it = _rates.begin(); it != _rates.end();)
    {
        const auto found = seen.find(it->first);
        if (found == seen.end())
        {
            it = _rates.erase(it);
            continue;
        }
        ++it;
    }

    if (capturePerf)
    {
        const std::wstring detail =
            std::format(L"tasks={} updated={} silent={} rates={}", snapshot.size(), totals.updatedTaskCount, totals.silentTaskCount, _rates.size());
        Debug::Perf::Emit(L"FileOps.Popup.Rate.UpdateUs", detail, PerfElapsedUs(perfStartUs), snapshot.size(), _rates.size(), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Rate.MaxCallbackSilenceMs", detail, totals.maxCallbackSilenceMs, totals.silentTaskCount, snapshot.size(), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Rate.MaxDisplayGapMs", detail, totals.maxDisplayGapMs, snapshot.size(), _rates.size(), S_OK);
    }
}

float FileOperationsPopupInternal::FileOperationsPopupState::FooterHeightPixels() const noexcept
{
    return DipsToPixels(_footerHeightDip, _dpi);
}

float FileOperationsPopupInternal::FileOperationsPopupState::MeasureSpeedLimitWidthDip(size_t slot, std::wstring_view caption) noexcept
{
    constexpr float minimumWidthDip = 140.0f;
    auto* format                    = _controlsHost.GetTextFormat(DxUi::FontRole::Body);
    if (! format || ! _dwriteFactory || slot >= _speedLimitWidthMeasurements.size())
        return minimumWidthDip;
    if (_speedLimitMeasuredFormat.get() != format)
    {
        _speedLimitMeasuredFormat = format;
        for (auto& cached : _speedLimitWidthMeasurements)
            cached.widthDip = 0.0f;
    }
    auto& cached = _speedLimitWidthMeasurements[slot];
    if (cached.widthDip > 0.0f && cached.caption == caption)
        return cached.widthDip;
    const float labelWidthDip = MeasureTextWidth(_dwriteFactory.get(), format, caption, 4096.0f, 128.0f);
    if (labelWidthDip <= 0.0f)
        return minimumWidthDip; // Retry a temporarily unavailable measurement on the next paint.
    // Selector chrome reserves matching leading/trailing lanes, not just the
    // visible chevron. Add breathing room beyond the complete measured caption.
    const float chromeWidthDip = 2.0f * DxUi::GetButtonChromeDropDownSegmentWidthDip(DxUi::ButtonVariant::Selector);
    cached.caption             = caption;
    cached.widthDip            = std::max(minimumWidthDip, std::ceil(labelWidthDip) + chromeWidthDip + 8.0f);
    return cached.widthDip;
}

HRESULT FileOperationsPopupInternal::FileOperationsPopupState::PrepareTerminalSummary(size_t slot, const TaskSnapshot& task, float widthPx) noexcept
{
    if (slot >= _terminalSummaries.size() || ! _dwriteFactory || ! _bodyFormat || ! std::isfinite(widthPx) || widthPx <= 0.0f)
        return E_INVALIDARG;
    auto& cached = _terminalSummaries[slot];
    if (cached.layout && cached.taskId == task.taskId && cached.text == task.resultSummary && cached.format.get() == _bodyFormat.get() &&
        cached.widthPx == widthPx)
        return S_OK;
    cached                   = {};
    const uint64_t startedUs = PerfNowUs();
    wil::com_ptr<IDWriteTextLayout> layout;
    HRESULT hr = _dwriteFactory->CreateTextLayout(
        task.resultSummary.data(), static_cast<UINT32>(task.resultSummary.size()), _bodyFormat.get(), widthPx, 4096.0f, layout.put());
    if (SUCCEEDED(hr))
        hr = layout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
    DWRITE_TEXT_METRICS metrics{};
    if (SUCCEEDED(hr))
        hr = layout->GetMetrics(&metrics);
    const float heightPx = std::ceil(metrics.height);
    if (SUCCEEDED(hr) && (! std::isfinite(heightPx) || heightPx <= 0.0f || metrics.width > widthPx + 0.5f))
        hr = E_UNEXPECTED;
    if (SUCCEEDED(hr))
        hr = layout->SetMaxHeight(heightPx);
    if (SUCCEEDED(hr))
    {
        cached.taskId   = task.taskId;
        cached.text     = task.resultSummary;
        cached.format   = _bodyFormat;
        cached.layout   = std::move(layout);
        cached.widthPx  = widthPx;
        cached.heightPx = heightPx;
    }
    Debug::Perf::Emit(L"FileOps.Popup.TerminalSummary.MeasureUs",
                      L"current terminal summary cache miss",
                      PerfElapsedUs(startedUs),
                      task.resultSummary.size(),
                      metrics.lineCount,
                      hr);
    return hr;
}

HRESULT FileOperationsPopupInternal::FileOperationsPopupState::PrepareTerminalFacts(size_t slot,
                                                                                    const TaskSnapshot& task,
                                                                                    float widthPx,
                                                                                    uint64_t languageRevision) noexcept
{
    if (slot >= _terminalFacts.size() || ! _dwriteFactory || ! _bodyFormat || ! std::isfinite(widthPx) || widthPx <= 0.0f)
        return E_INVALIDARG;
    auto& cached      = _terminalFacts[slot];
    const auto counts = TerminalFactCounts(task);
    auto& prepared    = cached.prepared;
    if (prepared.layout && prepared.taskId == task.taskId && cached.counts == counts && cached.verificationState == task.verificationState &&
        cached.startedMutation == task.startedMutation && cached.languageRevision == languageRevision && prepared.format.get() == _bodyFormat.get() &&
        prepared.widthPx == widthPx)
        return S_OK;
    cached                   = {};
    const uint64_t startedUs = PerfNowUs();
    std::wstring text        = BuildTerminalFactsText(task);
    if (text.empty())
        return E_UNEXPECTED;
    wil::com_ptr<IDWriteTextLayout> layout;
    HRESULT hr = _dwriteFactory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), _bodyFormat.get(), widthPx, 4096.0f, layout.put());
    if (SUCCEEDED(hr))
        hr = layout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
    DWRITE_TEXT_METRICS metrics{};
    if (SUCCEEDED(hr))
        hr = layout->GetMetrics(&metrics);
    const float heightPx = std::ceil(metrics.height);
    if (SUCCEEDED(hr) && (! std::isfinite(heightPx) || heightPx <= 0.0f || metrics.width > widthPx + 0.5f))
        hr = E_UNEXPECTED;
    if (SUCCEEDED(hr))
        hr = layout->SetMaxHeight(heightPx);
    if (SUCCEEDED(hr))
    {
        prepared.taskId          = task.taskId;
        prepared.text            = std::move(text);
        prepared.format          = _bodyFormat;
        prepared.layout          = std::move(layout);
        prepared.widthPx         = widthPx;
        prepared.heightPx        = heightPx;
        cached.counts            = counts;
        cached.verificationState = task.verificationState;
        cached.startedMutation   = task.startedMutation;
        cached.languageRevision  = languageRevision;
    }
    Debug::Perf::Emit(
        L"FileOps.Popup.TerminalFacts.MeasureUs", L"current terminal facts cache miss", PerfElapsedUs(startedUs), prepared.text.size(), metrics.lineCount, hr);
    return hr;
}

HRESULT FileOperationsPopupInternal::FileOperationsPopupState::PrepareTerminalActions(size_t slot,
                                                                                      const TaskSnapshot& task,
                                                                                      float widthPx,
                                                                                      bool rightToLeft) noexcept
{
    if (slot >= _terminalActions.size() || ! std::isfinite(widthPx) || widthPx <= 0.0f)
        return E_INVALIDARG;
    auto* format         = _controlsHost.GetTextFormat(DxUi::FontRole::Body,
                                                       DWRITE_TEXT_ALIGNMENT_CENTER,
                                                       DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                                                       true,
                                                       rightToLeft ? DWRITE_READING_DIRECTION_RIGHT_TO_LEFT : DWRITE_READING_DIRECTION_LEFT_TO_RIGHT);
    const float widthDip = _controlsHost.PixelsToDip(widthPx);
    std::array<std::wstring, 3> labels{{LoadStringResource(nullptr, IDS_FILEOP_BTN_DISMISS),
                                        task.hasProjectableIssues ? LoadStringResource(nullptr, IDS_FILEOPS_BTN_ISSUES) : std::wstring{},
                                        CompletedTaskHasOverflowActions(task) ? std::wstring(L"\u2026") : std::wstring{}}};
    auto& cached = _terminalActions[slot];
    if (cached.format.get() == format && cached.widthDip == widthDip && cached.dpi == _dpi && cached.rightToLeft == rightToLeft && cached.labels == labels)
        return S_OK;
    const uint64_t startedUs = PerfNowUs();
    std::array<D2D1_SIZE_F, 3> sizes{};
    std::array<D2D1_RECT_F, 3> bounds{};
    DxUi::MeasuredActionLayout layout{};
    const HRESULT hr = [&]() noexcept -> HRESULT
    {
        if (! format || ! _dwriteFactory)
            return E_PENDING;
        if (widthDip <= 52.0f)
            return E_INVALIDARG;
#ifdef ENABLE_TESTS
        if (ConsumeForcedFailure(g_fileOperationsLayoutMeasurementForcedFailures))
            return E_OUTOFMEMORY;
#endif
        for (size_t index = 0u; index < labels.size(); ++index)
        {
            if (labels[index].empty())
                continue;
            const float padding = index == 2u ? 52.0f : 24.0f;
            wil::com_ptr<IDWriteTextLayout> text;
            RETURN_IF_FAILED(_dwriteFactory->CreateTextLayout(
                labels[index].data(), static_cast<UINT32>(labels[index].size()), format, widthDip - padding, 4096.0f, text.put()));
            DWRITE_TEXT_METRICS metrics{};
            RETURN_IF_FAILED(text->GetMetrics(&metrics));
            sizes[index] = D2D1::SizeF(std::min(widthDip, std::max(80.0f, std::ceil(metrics.widthIncludingTrailingWhitespace) + padding)),
                                       std::max(32.0f, std::ceil(metrics.height) + 16.0f));
        }
        return DxUi::ArrangeMeasuredActions(
            sizes, widthDip, D2D1::SizeF(8.0f, 6.0f), bounds, layout, rightToLeft ? DxUi::FlowDirection::RightToLeft : DxUi::FlowDirection::LeftToRight);
    }();
    if (FAILED(hr))
    {
        // x-arch#2: Dismiss, Issues and More stay reachable on full-width 32-DIP rows. The
        // empty format key measures again on the next paint.
        cached    = {};
        float top = 0.0f;
        for (size_t index = 0u; index < labels.size(); ++index)
        {
            if (labels[index].empty())
                continue;
            cached.bounds[index] = D2D1::RectF(0.0f, top, std::max(1.0f, widthDip), top + 32.0f);
            cached.heightDip     = top + 32.0f;
            top += 32.0f + 6.0f;
        }
        cached.labels      = std::move(labels);
        cached.widthDip    = widthDip;
        cached.dpi         = _dpi;
        cached.rightToLeft = rightToLeft;
        return hr;
    }
    cached.labels      = std::move(labels);
    cached.bounds      = bounds;
    cached.format      = format;
    cached.widthDip    = widthDip;
    cached.heightDip   = layout.heightDip;
    cached.dpi         = _dpi;
    cached.rightToLeft = rightToLeft;
    Debug::Perf::Emit(L"FileOps.Popup.TerminalActions.MeasureUs", L"current terminal actions cache miss", PerfElapsedUs(startedUs), 0u, layout.rowCount, S_OK);
    return S_OK;
}

HRESULT FileOperationsPopupInternal::FileOperationsPopupState::PrepareCompletedGroup(float widthPx,
                                                                                     uint32_t total,
                                                                                     const std::array<uint32_t, 5>& counts) noexcept
{
    auto* actionFormat = _controlsHost.GetTextFormat(DxUi::FontRole::Body);
    if (! _dwriteFactory || ! actionFormat || ! _headerFormat || ! _smallFormat || widthPx <= 0.0f)
        return E_PENDING;
    const std::wstring clearText   = LoadStringResource(nullptr, IDS_FILEOPS_BTN_CLEAR_COMPLETED);
    const std::wstring headingText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPLETED_GROUP, static_cast<unsigned long>(total));
    auto& cached                   = _completedGroupMeasurement;
    if (cached.widthPx == widthPx && cached.dpi == _dpi && cached.expanded == _completedGroupExpanded && cached.clearText == clearText &&
        cached.headingText == headingText && cached.counts == counts && cached.actionFormat.get() == actionFormat &&
        cached.headingFormat.get() == _headerFormat.get() && cached.countFormat.get() == _smallFormat.get())
        return S_OK;
    cached                   = {};
    const float availableDip = _controlsHost.PixelsToDip(widthPx);
    wil::com_ptr<IDWriteTextLayout> headingLayout;
    FileOperationsCompletedGroupLayout::Layout arranged{};
    const HRESULT hr = [&]() noexcept -> HRESULT
    {
#ifdef ENABLE_TESTS
        if (ConsumeForcedFailure(g_fileOperationsLayoutMeasurementForcedFailures))
            return E_OUTOFMEMORY;
#endif
        const float clearWidthDip  = MeasureTextWidth(_dwriteFactory.get(), actionFormat, clearText, 4096.0f, 128.0f);
        const float headingWidthPx = MeasureTextWidth(_dwriteFactory.get(), _headerFormat.get(), headingText, 4096.0f, DipsToPixels(128.0f, _dpi));
        if (clearWidthDip <= 0.0f || headingWidthPx <= 0.0f)
            return E_FAIL;
        std::array<D2D1_SIZE_F, 5> badgeSizes{};
        size_t badgeCount = 0u;
        if (! _completedGroupExpanded)
        {
            for (const uint32_t count : counts)
            {
                if (count == 0u)
                    continue;
                const float countWidthPx =
                    MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), std::to_wstring(count), 4096.0f, DipsToPixels(128.0f, _dpi));
                if (countWidthPx <= 0.0f)
                    return E_FAIL;
                // 4-DIP leading inset, 18-DIP icon and 4-DIP trailing inset.
                badgeSizes[badgeCount++] = D2D1::SizeF(std::max(48.0f, std::ceil(_controlsHost.PixelsToDip(countWidthPx)) + 26.0f), 22.0f);
            }
        }
        D2D1_SIZE_F headingSize = D2D1::SizeF(1.0f, 1.0f);
        if (_completedGroupExpanded)
        {
            const float headingWidth = std::min(widthPx, std::ceil(headingWidthPx));
            RETURN_IF_FAILED(_dwriteFactory->CreateTextLayout(
                headingText.data(), static_cast<UINT32>(headingText.size()), _headerFormat.get(), headingWidth, 4096.0f, headingLayout.put()));
            RETURN_IF_FAILED(headingLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP));
            DWRITE_TEXT_METRICS headingMetrics{};
            RETURN_IF_FAILED(headingLayout->GetMetrics(&headingMetrics));
            if (! std::isfinite(headingMetrics.height) || headingMetrics.height <= 0.0f || headingMetrics.width > headingWidth + 0.5f)
                return E_UNEXPECTED;
            RETURN_IF_FAILED(headingLayout->SetMaxHeight(std::ceil(headingMetrics.height)));
            headingSize = D2D1::SizeF(_controlsHost.PixelsToDip(headingWidth), _controlsHost.PixelsToDip(std::ceil(headingMetrics.height)));
        }
        return FileOperationsCompletedGroupLayout::Arrange(availableDip,
                                                           D2D1::SizeF(std::ceil(clearWidthDip) + 16.0f, 28.0f),
                                                           headingSize,
                                                           std::span(badgeSizes).first(badgeCount),
                                                           _completedGroupExpanded,
                                                           arranged);
    }();
    if (FAILED(hr))
    {
        // x-arch#2 (for example ERROR_INSUFFICIENT_BUFFER from a fit check): keep the disclosure
        // and Clear completed on full-width rows and skip the optional heading and badges. The
        // empty format keys measure again on the next paint.
        using namespace FileOperationsCompletedGroupLayout;
        const float rowWidthDip  = std::max(availableDip, kDisclosureSizeDip);
        const float clearTop     = kDisclosureSizeDip + kActionGapDip;
        cached.layout.disclosure = D2D1::RectF(rowWidthDip - kDisclosureSizeDip, 0.0f, rowWidthDip, kDisclosureSizeDip);
        cached.layout.clear      = D2D1::RectF(0.0f, clearTop, rowWidthDip, clearTop + 28.0f);
        cached.layout.rowCount   = 2u;
        cached.layout.heightDip  = clearTop + 28.0f;
        cached.clearText         = clearText;
        cached.headingText       = headingText;
        cached.counts            = counts;
        cached.widthPx           = widthPx;
        cached.dpi               = _dpi;
        cached.expanded          = _completedGroupExpanded;
        return hr;
    }
    cached.clearText     = clearText;
    cached.headingText   = headingText;
    cached.counts        = counts;
    cached.actionFormat  = actionFormat;
    cached.headingFormat = _headerFormat;
    cached.countFormat   = _smallFormat;
    cached.headingLayout = std::move(headingLayout);
    cached.layout        = arranged;
    cached.widthPx       = widthPx;
    cached.dpi           = _dpi;
    cached.expanded      = _completedGroupExpanded;
    return S_OK;
}

void FileOperationsPopupInternal::FileOperationsPopupState::MeasureFooterActionWidths() noexcept
{
    IDWriteTextFormat* const actionFormat = _controlsHost.GetTextFormat(DxUi::FontRole::Body);
    if (actionFormat)
    {
        std::array labels{LoadStringResource(nullptr, IDS_FILEOPS_BTN_PAUSE_ALL),
                          LoadStringResource(nullptr, IDS_FILEOPS_BTN_RESUME_ALL),
                          LoadStringResource(nullptr, IDS_FILEOPS_BTN_CANCEL_ALL)};
        if (_footerMeasuredFormat.get() != actionFormat || _footerMeasuredLabels != labels)
        {
            const uint64_t measureStartUs = PerfNowUs();
            std::array<float, 2> widths{112.0f, 112.0f};
            bool measured = true;
            for (size_t i = 0u; i < labels.size(); ++i)
            {
                const float labelWidth = MeasureTextWidth(_dwriteFactory.get(), actionFormat, labels[i], 4096.0f, 128.0f);
                measured               = measured && labelWidth > 0.0f;
                const size_t slot      = i == 2u ? 1u : 0u;
                widths[slot]           = std::max(widths[slot], std::ceil(labelWidth + 24.0f));
            }
            if (measured)
            {
                _footerActionWidthsDip = widths;
                _footerMeasuredLabels  = std::move(labels);
                _footerMeasuredFormat  = actionFormat;
            }
            Debug::Perf::Emit(L"FileOps.Popup.FooterActions.MeasureUs", L"footer label cache miss", PerfElapsedUs(measureStartUs), labels.size(), 0u, S_OK);
        }
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::PrepareFooterSummary(float widthPx, std::wstring_view text) noexcept
{
    if (_footerSummaryLayout && _footerSummaryFormat.get() == _smallFormat.get() && _footerSummaryWidthPx == widthPx && _footerSummaryText == text)
        return;

    _footerSummaryLayout.reset();
    _footerSummaryFormat.reset();
    _footerSummaryText.clear();
    _footerSummaryHeightPx        = DipsToPixels(FileOperationsFooterLayout::kSummaryHeightDip, _dpi);
    _footerMinimumSummaryHeightPx = _footerSummaryHeightPx;
    if (! _dwriteFactory || ! _smallFormat || widthPx <= 0.0f || text.empty())
        return;

    const uint64_t startedUs = PerfNowUs();
    wil::com_ptr<IDWriteTextLayout> layout;
    HRESULT hr = _dwriteFactory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), _smallFormat.get(), widthPx, 4096.0f, layout.put());
    if (SUCCEEDED(hr))
        hr = layout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
    DWRITE_TEXT_METRICS metrics{};
    if (SUCCEEDED(hr))
        hr = layout->GetMetrics(&metrics);
    // Reserve the summary at minimum client width too. Reuse this layout; the
    // WM_GETMINMAXINFO path must not allocate text resources while dragging.
    DWRITE_TEXT_METRICS minimumMetrics = metrics;
    const float minimumWidthPx =
        DipsToPixels(static_cast<float>(FileOperationsFooterLayout::kMinimumClientWidthDip) - 2.0f * FileOperationsFooterLayout::kMarginDip, _dpi);
    if (SUCCEEDED(hr) && widthPx != minimumWidthPx)
    {
        hr = layout->SetMaxWidth(minimumWidthPx);
        if (SUCCEEDED(hr))
            hr = layout->GetMetrics(&minimumMetrics);
        if (SUCCEEDED(hr))
            hr = layout->SetMaxWidth(widthPx);
    }
    const float heightPx = std::max(_footerSummaryHeightPx, std::ceil(metrics.height));
    if (SUCCEEDED(hr))
        hr = layout->SetMaxHeight(heightPx);
    if (SUCCEEDED(hr))
    {
        _footerSummaryLayout          = std::move(layout);
        _footerSummaryFormat          = _smallFormat;
        _footerSummaryText            = text;
        _footerSummaryWidthPx         = widthPx;
        _footerSummaryHeightPx        = heightPx;
        _footerMinimumSummaryHeightPx = std::max(DipsToPixels(FileOperationsFooterLayout::kSummaryHeightDip, _dpi), std::ceil(minimumMetrics.height));
    }
    Debug::Perf::Emit(L"FileOps.Popup.FooterSummary.MeasureUs", L"current summary cache miss", PerfElapsedUs(startedUs), text.size(), metrics.lineCount, hr);
}

void FileOperationsPopupInternal::FileOperationsPopupState::LayoutChrome(
    float width, float height, bool showPauseResumeAll, bool showCancelAll, std::wstring_view summaryText) noexcept
{
    FileOperationsFooterLayout::Layout actions{};
    const float footerMargin = DipsToPixels(FileOperationsFooterLayout::kMarginDip, _dpi);
    PrepareFooterSummary(width - 2.0f * footerMargin, summaryText);
    const float availableWidthDip = _controlsHost.PixelsToDip(width) - 2.0f * FileOperationsFooterLayout::kMarginDip;
    const HRESULT layoutHr        = FileOperationsFooterLayout::Arrange(availableWidthDip,
                                                                        showPauseResumeAll ? _footerActionWidthsDip[0] : 0.0f,
                                                                        showCancelAll ? _footerActionWidthsDip[1] : 0.0f,
                                                                        actions,
                                                                        _controlsHost.PixelsToDip(_footerSummaryHeightPx));
    _footerHeightDip              = actions.heightDip;
    const float footerTop         = std::max(0.0f, height - FooterHeightPixels());
    _listViewportRect             = D2D1::RectF(0.0f, 0.0f, width, footerTop);
    const float rightEdge         = std::max(footerMargin, width - footerMargin);
    const float progressY         = footerTop + DipsToPixels(8.0f, _dpi);
    _footerAggregateProgressRect  = D2D1::RectF(footerMargin, progressY, rightEdge, std::min(height, progressY + DipsToPixels(6.0f, _dpi)));
    const float summaryY          = footerTop + DipsToPixels(18.0f, _dpi);
    _footerSummaryRect            = D2D1::RectF(footerMargin, summaryY, rightEdge, summaryY + _footerSummaryHeightPx);
    const float commandY          = footerTop + DipsToPixels(actions.commandTopDip, _dpi);
    const auto place              = [&](const D2D1_RECT_F& bounds) noexcept
    {
        // A transient undersized/zero-width host must not publish inverted hit
        // targets. Supported minimum-width layouts always admit the full cluster.
        if (FAILED(layoutHr) || bounds.right <= bounds.left || bounds.bottom <= bounds.top)
            return D2D1_RECT_F{};
        return D2D1::RectF(footerMargin + DipsToPixels(bounds.left, _dpi),
                           commandY + DipsToPixels(bounds.top, _dpi),
                           footerMargin + DipsToPixels(bounds.right, _dpi),
                           commandY + DipsToPixels(bounds.bottom, _dpi));
    };
    _footerPauseResumeAllRect = place(actions.pauseResume);
    _footerCancelAllRect      = place(actions.cancel);
    _footerQueueModeRect      = place(actions.queue);
    _footerOptionsRect        = place(actions.options);
    _footerDetailsToggleRect  = place(actions.details);
    _footerAutoDismissRect    = {};
    _footerDensityRect        = {};
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateScrollBar(HWND hwnd, float viewH, float contentH) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const int viewHeight      = std::max(0, static_cast<int>(std::ceil(viewH)));
    const int contentHeightPx = std::max(0, static_cast<int>(std::ceil(contentH)));

    if (! _scrollBarVisible)
    {
        _scrollPos = 0;
    }

    SCROLLINFO si{};
    si.cbSize      = sizeof(si);
    si.fMask       = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin        = 0;
    si.nMax        = std::max(0, contentHeightPx - 1);
    const int page = std::clamp(viewHeight, 1, std::numeric_limits<int>::max());
    si.nPage       = static_cast<UINT>(page);

    const int maxPos = std::max(0, si.nMax - page + 1);
    _scrollPos       = std::clamp(_scrollPos, 0, maxPos);
    si.nPos          = _scrollPos;

    SCROLLINFO current{};
    current.cbSize = sizeof(current);
    current.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    if (GetScrollInfo(hwnd, SB_VERT, &current) && current.nMin == si.nMin && current.nMax == si.nMax && current.nPage == si.nPage && current.nPos == si.nPos)
        return; // Avoid synchronous native non-client painting on unchanged frames.
    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
}

void FileOperationsPopupInternal::FileOperationsPopupState::AutoResizeWindow(
    HWND hwnd, float desiredContentHeight, size_t taskCount, bool footerOnly, bool reducedMotion) noexcept
{
#ifdef ENABLE_TESTS
    if (debugFixedGeometry && debugVisualScenario.has_value())
        return;
#endif
    if (! hwnd || _inSizeMove)
    {
        return;
    }
    if (! _pathDetailId.empty())
    {
        // x-lifetime#3: path detail suspends auto-resize. Drop a pending or running resize, as
        // entering size-move does, so it cannot hold the 16 ms motion cadence while detail stays
        // open, and measure again from scratch once it closes.
        _autoResizePending          = false;
        _autoResizeAnimating        = false;
        _lastAutoSizedContentHeight = -1.0f;
        return;
    }

    const ULONGLONG nowTick = GetTickCount64();
    if (footerOnly)
    {
        _maxAutoSizedWindowHeight = 0;
        _autoResizePending        = false;
        _autoResizeAnimating      = false;
    }

    // Only auto-resize if task count or content height changed
    const bool taskCountChanged     = taskCount != _lastTaskCount;
    const bool contentHeightChanged = std::abs(desiredContentHeight - _lastAutoSizedContentHeight) > 1.0f;
    const bool footerHeightChanged  = std::abs(_footerHeightDip - _lastAutoSizedFooterHeight) > 0.5f;

    if (! taskCountChanged && ! contentHeightChanged && ! footerHeightChanged && ! _autoResizePending && ! _autoResizeAnimating && ! _footerOnlyRestorePending)
    {
        return;
    }

    if (taskCountChanged || contentHeightChanged || footerHeightChanged)
    {
        _lastTaskCount              = taskCount;
        _lastAutoSizedContentHeight = desiredContentHeight;
        _lastAutoSizedFooterHeight  = _footerHeightDip;
    }

    // Get current window rect
    RECT windowRc{};
    GetWindowRect(hwnd, &windowRc);

    // Get screen work area (excludes taskbar)
    HMONITOR hMonitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (! GetMonitorInfoW(hMonitor, &mi))
    {
        return;
    }
    const RECT& workArea      = mi.rcWork;
    const int maxScreenHeight = workArea.bottom - workArea.top;

    // Calculate the footer and chrome heights
    const float footerH             = FooterHeightPixels();
    const float desiredClientHeight = desiredContentHeight + footerH;

    // Get window style for AdjustWindowRectExForDpi
    const DWORD style   = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const DWORD exStyle = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE));

    // Calculate desired window height from client height
    RECT clientRc{0, 0, windowRc.right - windowRc.left, static_cast<LONG>(std::ceil(desiredClientHeight))};
    AdjustWindowRectExForDpi(&clientRc, style, FALSE, exStyle, _dpi);

    int desiredWindowHeight = clientRc.bottom - clientRc.top;

    // Apply minimum height constraint
    const int minClientHeightDip = footerOnly ? kFileOperationsPopupFooterOnlyMinClientHeightDip : kFileOperationsPopupMinClientHeightDip;
    const int minClientH = std::max(DipsToPixels(minClientHeightDip, _dpi), static_cast<int>(std::ceil(FooterHeightPixels() + DipsToPixels(8.0f, _dpi))));
    RECT minRc{0, 0, 0, minClientH};
    AdjustWindowRectExForDpi(&minRc, style, FALSE, exStyle, _dpi);
    const int minWindowHeight = minRc.bottom - minRc.top;
    desiredWindowHeight       = std::max(desiredWindowHeight, minWindowHeight);

    // Clamp to screen height
    desiredWindowHeight = std::min(desiredWindowHeight, maxScreenHeight);

    // Prevent resize "dancing": once the window grows to fit more lines/tasks, don't auto-shrink it again.
    if (! footerOnly && _maxAutoSizedWindowHeight > 0)
    {
        desiredWindowHeight = std::max(desiredWindowHeight, _maxAutoSizedWindowHeight);
        desiredWindowHeight = std::min(desiredWindowHeight, maxScreenHeight);
    }

    // Calculate new position - keep top position, adjust bottom
    int newTop    = windowRc.top;
    int newBottom = newTop + desiredWindowHeight;

    // If window would extend below work area, move it up
    if (newBottom > workArea.bottom)
    {
        newBottom = workArea.bottom;
        newTop    = newBottom - desiredWindowHeight;
        // But don't go above work area
        if (newTop < workArea.top)
        {
            newTop    = workArea.top;
            newBottom = newTop + std::min(desiredWindowHeight, maxScreenHeight);
        }
    }

    const RECT targetRect{windowRc.left, newTop, windowRc.right, newBottom};
    if (RectsNearEqual(windowRc, targetRect))
    {
        _autoResizePending        = false;
        _autoResizeAnimating      = false;
        _footerOnlyRestorePending = false;
        if (! footerOnly)
        {
            _maxAutoSizedWindowHeight = std::max(_maxAutoSizedWindowHeight, static_cast<int>(windowRc.bottom - windowRc.top));
        }
        return;
    }

    if (_footerOnlyRestorePending && ! footerOnly)
    {
        _footerOnlyRestorePending = false;
        _autoResizePending        = false;
        _autoResizeAnimating      = false;
        _maxAutoSizedWindowHeight = std::max(_maxAutoSizedWindowHeight, static_cast<int>(windowRc.bottom - windowRc.top));
        return;
    }

    if (footerOnly || reducedMotion)
    {
        _autoResizePending   = false;
        _autoResizeAnimating = false;
        SetWindowPos(hwnd,
                     nullptr,
                     targetRect.left,
                     targetRect.top,
                     targetRect.right - targetRect.left,
                     targetRect.bottom - targetRect.top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        if (! footerOnly)
        {
            _maxAutoSizedWindowHeight = std::max(_maxAutoSizedWindowHeight, static_cast<int>(targetRect.bottom - targetRect.top));
        }
        return;
    }

    constexpr ULONGLONG kResizeDebounceMs = 140ull;
    constexpr ULONGLONG kResizeEaseMs     = 240ull;

    if (! _autoResizeAnimating)
    {
        if (! _autoResizePending || ! RectsNearEqual(_autoResizePendingTargetRect, targetRect))
        {
            _autoResizePending           = true;
            _autoResizePendingTargetRect = targetRect;
            _autoResizePendingDueTick    = nowTick + kResizeDebounceMs;
            return;
        }

        if (nowTick < _autoResizePendingDueTick)
        {
            return;
        }

        _autoResizePending             = false;
        _autoResizeAnimating           = true;
        _autoResizeAnimationStartRect  = windowRc;
        _autoResizeAnimationTargetRect = _autoResizePendingTargetRect;
        _autoResizeAnimationStartTick  = nowTick;
    }
    else if (! RectsNearEqual(_autoResizeAnimationTargetRect, targetRect))
    {
        _autoResizeAnimationStartRect  = windowRc;
        _autoResizeAnimationTargetRect = targetRect;
        _autoResizeAnimationStartTick  = nowTick;
    }

    const float fraction = EaseFileOperationsUiMotionFraction(nowTick - _autoResizeAnimationStartTick, kResizeEaseMs);
    const RECT nextRect = fraction >= 1.0f ? _autoResizeAnimationTargetRect : LerpRect(_autoResizeAnimationStartRect, _autoResizeAnimationTargetRect, fraction);

    SetWindowPos(hwnd, nullptr, nextRect.left, nextRect.top, nextRect.right - nextRect.left, nextRect.bottom - nextRect.top, SWP_NOZORDER | SWP_NOACTIVATE);

    if (fraction >= 1.0f)
    {
        _autoResizeAnimating      = false;
        _maxAutoSizedWindowHeight = std::max(_maxAutoSizedWindowHeight, static_cast<int>(nextRect.bottom - nextRect.top));
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::ArmUiMotionTimer(HWND hwnd) noexcept
{
    if (! hwnd || IsReducedMotionEnabled())
    {
        return;
    }

    _uiMotionTimerDueTick = GetTickCount64() + kFileOperationsPopupUiMotionTimerWindowMs;
    RefreshTelemetryTimer(hwnd);
}

bool FileOperationsPopupInternal::FileOperationsPopupState::HasVisibleUiMotion(HWND hwnd, ULONGLONG nowTick) const noexcept
{
    if (! IsWindowVisible(hwnd) || IsIconic(hwnd) || _failureSurface)
        return false;
    if ((_uiMotionTimerDueTick != 0u && nowTick < _uiMotionTimerDueTick) || _autoResizePending || _autoResizeAnimating)
        return true;
    const auto active = [nowTick](const ScalarAnimation& animation, ULONGLONG duration) noexcept
    {
        // A stale hidden scalar may retain an intermediate value. Its elapsed
        // deadline, rather than that value alone, bounds the last frame's timer.
        return animation.initialized && animation.displayed != animation.target && (nowTick < animation.startTick || nowTick - animation.startTick < duration);
    };
    if (active(_completedGroupVisibilityAnimation, kFileOperationsPopupUiMotionDurationMs) ||
        active(_footerDisclosureAnimation, kFileOperationsPopupUiMotionDurationMs))
        return true;
    for (const auto& [taskId, animation] : _taskHeightAnimations)
    {
        static_cast<void>(taskId);
        if (active(animation.height, kFileOperationsPopupUiMotionDurationMs))
            return true;
    }
    for (const auto& [taskId, animation] : _taskDisclosureAnimations)
    {
        static_cast<void>(taskId);
        if (active(animation, kFileOperationsPopupUiMotionDurationMs))
            return true;
    }
    return std::ranges::any_of(_hostedProgressAnimations, [&](const ScalarAnimation& animation) noexcept { return active(animation, 160u); });
}

void FileOperationsPopupInternal::FileOperationsPopupState::RefreshTelemetryTimer(HWND hwnd) noexcept
{
    if (! hwnd || _deletePending)
        return;
    const ULONGLONG nowTick = GetTickCount64();
    if (_uiMotionTimerDueTick != 0u && nowTick >= _uiMotionTimerDueTick)
        _uiMotionTimerDueTick = 0u;
    using TelemetryWork     = FolderWindow::FileOperationState::PopupTelemetryWork;
    TelemetryWork workState = fileOps && hostLifetime.lock() ? fileOps->GetPopupTelemetryWork() : TelemetryWork::None;
#ifdef ENABLE_TESTS
    if (debugVisualScenario.has_value())
    {
        // The synthetic fixture applies the engine's classification to its declared snapshots.
        workState = TelemetryWork::None;
        for (const TaskSnapshot& task : debugVisualScenario.value())
        {
            if (task.finished)
                continue;
            const bool waiting = task.kind != TaskSnapshot::Kind::Informational && ! task.cancellationRequested &&
                                 (task.paused || task.queuePaused || task.waitingInQueue || (task.conflict.active && ! task.conflict.metadataLoading));
            workState          = waiting ? TelemetryWork::Waiting : TelemetryWork::Active;
            if (! waiting)
                break;
        }
    }
#endif
    // x-perf#5: paused, queued and prompt-blocked tasks change nothing between ticks; they keep
    // the elapsed text and taskbar current at the waiting cadence instead of 10 frames a second.
    const bool work        = workState == TelemetryWork::Active;
    const bool waitingWork = workState == TelemetryWork::Waiting;
    // Active progress keeps its existing telemetry cadence. A changing progress
    // scalar must not repeatedly turn 100 ms sampling into the motion cadence.
    const bool explicitMotion = _uiMotionTimerDueTick != 0u && IsWindowVisible(hwnd) && ! IsIconic(hwnd) && ! _failureSurface;
    // Q6: a skipped frame or a pending re-attach keeps the timer that retries it. A skipped frame
    // retries at the telemetry cadence; a re-attach (one second or more away) wakes only when it
    // is due, at most once a second. x-perf#4: accessible text still waiting for its publication
    // interval keeps the telemetry cadence until a frame publishes it.
    const bool recoveryVisible = IsWindowVisible(hwnd) && ! IsIconic(hwnd);
    UINT idleIntervalMs        = 0u;
    if (recoveryVisible && (_framePrepareRetryTick != 0u || (_hostedAccessibleTextPending && ! _failureSurface)))
        idleIntervalMs = kFileOperationsPopupTimerIntervalMs;
    else if (recoveryVisible && _failureSurface && _hostReattachRetryTick != 0u)
    {
        const ULONGLONG remainingMs = _hostReattachRetryTick > nowTick ? _hostReattachRetryTick - nowTick : 0u;
        idleIntervalMs              = static_cast<UINT>(std::clamp(remainingMs, static_cast<ULONGLONG>(kFileOperationsPopupTimerIntervalMs), 1000ull));
    }
#ifdef ENABLE_TESTS
    // Counterfactual only: retain the normal cadence when production would stop.
    // Work, motion, taskbar, rendering and callback paths remain unchanged.
    if (! _debugIdleSuspensionEnabled)
        idleIntervalMs = kFileOperationsPopupTimerIntervalMs;
#endif
    const UINT waitingIntervalMs =
        idleIntervalMs != 0u ? (std::min)(idleIntervalMs, kFileOperationsPopupWaitingTimerIntervalMs) : kFileOperationsPopupWaitingTimerIntervalMs;
    const UINT desired = explicitMotion                      ? kFileOperationsPopupUiMotionTimerIntervalMs
                         : work                              ? kFileOperationsPopupTimerIntervalMs
                         : HasVisibleUiMotion(hwnd, nowTick) ? kFileOperationsPopupUiMotionTimerIntervalMs
                         : waitingWork                       ? waitingIntervalMs
                                                             : idleIntervalMs;
    if (desired == _telemetryTimerIntervalMs)
        return;
    if (desired != 0u)
    {
        _telemetryTimerIntervalMs = SetTimer(hwnd, kFileOperationsPopupTimerId, desired, nullptr) != 0u ? desired : 0u;
        return;
    }

    KillTimer(hwnd, kFileOperationsPopupTimerId);
    _telemetryTimerIntervalMs = 0u;
    // Clear/reduce the taskbar once at the quiet transition. Retained history is
    // no longer copied on idle ticks; Explorer can replay this cached model.
    if (hostLifetime.lock())
        UpdateTaskbarProgress(hwnd);
    if (! _deletePending && IsWindowVisible(hwnd) && ! IsIconic(hwnd))
        Invalidate(hwnd); // Commit the last animation frame before remaining idle.
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawDxUiButtonChrome(const PopupButton& button,
                                                                                 IDWriteTextFormat* format,
                                                                                 std::wstring_view text,
                                                                                 DxUi::ButtonVariant variant) noexcept
{
    if (! _target || ! _buttonChromeBrush || ! folderWindow || ! hostLifetime.lock())
    {
        return;
    }

    DxUi::ButtonChromeDrawSpec spec{};
    spec.bounds       = button.bounds;
    spec.text         = text;
    spec.variant      = variant;
    spec.scale        = static_cast<float>(_dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    spec.chevronGlyph = FluentIcons::kChevronDown;

    IDWriteTextFormat* iconFormat = _statusIconFormat.get();
    if (! iconFormat || ! DirectWriteFormatHasGlyph(_dwriteFactory.get(), iconFormat, spec.chevronGlyph))
    {
        spec.chevronGlyph = FluentIcons::kFallbackChevronDown;
        iconFormat        = _statusIconFallbackFormat ? _statusIconFallbackFormat.get() : format;
    }

    const AppTheme& appTheme = folderWindow->GetTheme();
    DxUi::DrawButtonChrome(_target.get(), _buttonChromeBrush.get(), format, iconFormat, MakeAppThemeDxPalette(appTheme, appTheme.windowBackground), spec);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawButton(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text) noexcept
{
    DrawDxUiButtonChrome(button, format, text, DxUi::ButtonVariant::Standard);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawFooterQueueModeControl(const PopupButton& button, bool queueMode) noexcept
{
    _footerQueueSegmentRect    = {};
    _footerParallelSegmentRect = {};
    if (button.bounds.right <= button.bounds.left || button.bounds.bottom <= button.bounds.top)
    {
        return;
    }
    DrawDxUiButtonChrome(button,
                         _buttonSmallFormat.get(),
                         queueMode ? std::wstring_view(_footerQueueText) : std::wstring_view(_footerParallelText),
                         DxUi::ButtonVariant::Selector);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawFooterAutoDismissControl(const PopupButton& button, bool enabled) noexcept
{
    _footerAutoDismissLabelVisible = false;
    if (! _target || ! _buttonBgBrush || ! _borderBrush || ! _smallFormat || button.bounds.right <= button.bounds.left ||
        button.bounds.bottom <= button.bounds.top)
    {
        return;
    }

    const float radius = ClampCornerRadius(button.bounds, DipsToPixels(4.0f, _dpi));
    _target->FillRoundedRectangle(D2D1::RoundedRect(button.bounds, radius, radius), _buttonBgBrush.get());
    _target->DrawRoundedRectangle(D2D1::RoundedRect(button.bounds, radius, radius), _borderBrush.get(), 1.0f);

    const float buttonW           = button.bounds.right - button.bounds.left;
    const float padX              = DipsToPixels(8.0f, _dpi);
    const float boxSize           = DipsToPixels(14.0f, _dpi);
    const float gap               = DipsToPixels(7.0f, _dpi);
    const float boxTop            = button.bounds.top + ((button.bounds.bottom - button.bounds.top - boxSize) * 0.5f);
    const std::wstring_view label = enabled ? std::wstring_view(_footerAutoDismissOnText) : std::wstring_view(_footerAutoDismissOffText);
    const float textAvailableW    = std::max(0.0f, buttonW - padX * 2.0f - boxSize - gap);
    const float labelW =
        MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), label, DipsToPixels(4096.0f, _dpi), button.bounds.bottom - button.bounds.top);
    const bool showLabel           = labelW > 0.0f && labelW <= textAvailableW;
    _footerAutoDismissLabelVisible = showLabel;
    const float boxLeft            = showLabel ? button.bounds.left + padX : button.bounds.left + (buttonW - boxSize) * 0.5f;
    const D2D1_RECT_F boxRc        = D2D1::RectF(boxLeft, boxTop, boxLeft + boxSize, boxTop + boxSize);
    DrawCheckboxBox(boxRc, enabled);

    if (! showLabel)
    {
        return;
    }

    const D2D1_RECT_F textRc = D2D1::RectF(
        boxRc.right + gap, button.bounds.top + DipsToPixels(4.0f, _dpi), button.bounds.right - padX, button.bounds.bottom - DipsToPixels(4.0f, _dpi));
    ID2D1Brush* textBrush = enabled && _textBrush ? _textBrush.get() : (_subTextBrush ? _subTextBrush.get() : _textBrush.get());
    if (textBrush && textRc.right > textRc.left)
    {
        _target->DrawTextW(label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), textRc, textBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }
}

wchar_t FileOperationsPopupInternal::FileOperationsPopupState::DrawCenteredGlyph(const D2D1_RECT_F& rc,
                                                                                 wchar_t fluentGlyph,
                                                                                 wchar_t fallbackGlyph,
                                                                                 ID2D1Brush* glyphBrush) noexcept
{
    ID2D1Brush* brush = glyphBrush ? glyphBrush : _textBrush.get();
    if (! _target || ! brush || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return L'\0';
    }

    const bool useFluentGlyph = _statusIconFormat != nullptr && DirectWriteFormatHasGlyph(_dwriteFactory.get(), _statusIconFormat.get(), fluentGlyph);
    const wchar_t glyph       = useFluentGlyph ? fluentGlyph : fallbackGlyph;
    IDWriteTextFormat* format =
        useFluentGlyph ? _statusIconFormat.get() : (_statusIconFallbackFormat ? _statusIconFallbackFormat.get() : _buttonSmallFormat.get());
    if (! format || glyph == 0)
    {
        return L'\0';
    }

    const wchar_t text[2]{glyph, 0};
    _target->DrawTextW(text, 1u, format, rc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
    return glyph;
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawMenuButton(const PopupButton& button,
                                                                           IDWriteTextFormat* format,
                                                                           std::wstring_view text) noexcept
{
    DrawDxUiButtonChrome(button, format, text, DxUi::ButtonVariant::DropDown);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawSelectorButton(const PopupButton& button,
                                                                               IDWriteTextFormat* format,
                                                                               std::wstring_view text) noexcept
{
    DrawDxUiButtonChrome(button, format, text, DxUi::ButtonVariant::Selector);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawCheckboxBox(const D2D1_RECT_F& rect, bool checked) noexcept
{
    if (! _target)
    {
        return;
    }

    const float size = std::max(0.0f, std::min(rect.right - rect.left, rect.bottom - rect.top));
    if (size <= 1.0f)
    {
        return;
    }

    const float left = rect.left + (rect.right - rect.left - size) * 0.5f;
    const float top  = rect.top + (rect.bottom - rect.top - size) * 0.5f;

    const D2D1_RECT_F boxRc = D2D1::RectF(left, top, left + size, top + size);

    ID2D1Brush* base = _buttonBgBrush ? _buttonBgBrush.get() : (_bgBrush ? _bgBrush.get() : nullptr);
    if (base)
    {
        _target->FillRectangle(boxRc, base);
    }

    if (checked && _checkboxFillBrush)
    {
        _target->FillRectangle(boxRc, _checkboxFillBrush.get());
    }

    if (_borderBrush)
    {
        _target->DrawRectangle(boxRc, _borderBrush.get(), 1.0f);
    }

    if (! checked)
    {
        return;
    }

    ID2D1Brush* checkBrush = _checkboxCheckBrush ? _checkboxCheckBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
    if (! checkBrush)
    {
        return;
    }

#ifdef ENABLE_TESTS
    _debugLastDrawnCheckboxGlyph = DrawCenteredGlyph(boxRc, FluentIcons::kCheckMark, FluentIcons::kFallbackCheckMark, checkBrush);
#else
    static_cast<void>(DrawCenteredGlyph(boxRc, FluentIcons::kCheckMark, FluentIcons::kFallbackCheckMark, checkBrush));
#endif
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawDisclosureChevron(const D2D1_RECT_F& rc, float expandedProgress) noexcept
{
    if (! _target || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    const DxUi::DisclosureChevronVisualState visual = DxUi::ResolveDisclosureChevronVisualState(expandedProgress, DxUi::ChevronDirection::Left);
    const wchar_t fluentGlyph                       = visual.direction == DxUi::ChevronDirection::Left ? FluentIcons::kChevronLeft : FluentIcons::kChevronDown;
    const wchar_t fallbackGlyph = visual.direction == DxUi::ChevronDirection::Left ? FluentIcons::kFallbackChevronLeft : FluentIcons::kFallbackChevronDown;
    if (visual.rotationDegrees == 0.0f)
    {
        static_cast<void>(DrawCenteredGlyph(rc, fluentGlyph, fallbackGlyph));
        return;
    }

    D2D1_MATRIX_3X2_F previousTransform{};
    _target->GetTransform(&previousTransform);
    const D2D1_POINT_2F center = D2D1::Point2F((rc.left + rc.right) * 0.5f, (rc.top + rc.bottom) * 0.5f);
    _target->SetTransform(D2D1::Matrix3x2F::Rotation(visual.rotationDegrees, center) * previousTransform);
    auto restoreTransform = wil::scope_exit([&] { _target->SetTransform(previousTransform); });
    static_cast<void>(DrawCenteredGlyph(rc, fluentGlyph, fallbackGlyph));
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawDiscoveryActivityIndicator(const D2D1_RECT_F& rc, ULONGLONG tick, bool reducedMotion) noexcept
{
    if (! _target || ! _graphDynamicBrush || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    const AppTheme* theme         = folderWindow && hostLifetime.lock() ? &folderWindow->GetTheme() : nullptr;
    D2D1_COLOR_F accent           = theme ? theme->accent : D2D1::ColorF(D2D1::ColorF::DodgerBlue);
    const float centerY           = (rc.top + rc.bottom) * 0.5f;
    const float dotRadius         = DipsToPixels(2.0f, _dpi);
    const float dotStep           = DipsToPixels(8.0f, _dpi);
    constexpr ULONGLONG kPeriodMs = 900ull;

    for (size_t index = 0u; index < 3u; ++index)
    {
        float strength = 0.58f;
        if (! reducedMotion)
        {
            const ULONGLONG offset = static_cast<ULONGLONG>(index) * (kPeriodMs / 3ull);
            const float phase      = static_cast<float>((tick + kPeriodMs - offset) % kPeriodMs) / static_cast<float>(kPeriodMs);
            strength               = 0.28f + 0.72f * std::max(0.0f, std::sin(phase * 2.0f * 3.14159265f));
        }
        accent.a = strength;
        _graphDynamicBrush->SetColor(accent);
        const float x = rc.left + dotRadius + static_cast<float>(index) * dotStep;
        _target->FillEllipse(D2D1::Ellipse(D2D1::Point2F(x, centerY), dotRadius, dotRadius), _graphDynamicBrush.get());
    }
}

float FileOperationsPopupInternal::FileOperationsPopupState::StalledRateLabelSlotWidth() noexcept
{
    if (! _dwriteFactory || ! _smallFormat)
        return 0.0f;
    if (_stalledRateSlotFormat.get() == _smallFormat.get())
        return _stalledRateSlotWidth;
    // A wide speed in each unit, with each stall name: the widest of them is the slot.
    constexpr std::array<double, 3> kWideRates{{999.9 * 1024.0, 999.9 * 1024.0 * 1024.0, 999.9 * 1024.0 * 1024.0 * 1024.0}};
    constexpr std::array<RateStallKind, 4> kStalls{
        {RateStallKind::Finalizing, RateStallKind::WaitingForDestination, RateStallKind::WaitingForSource, RateStallKind::WaitingForStorage}};
    float widest = 0.0f;
    for (const double rate : kWideRates)
    {
        for (const RateStallKind stall : kStalls)
        {
            const std::wstring text = FormatRateLabelText(rate, stall);
            wil::com_ptr<IDWriteTextLayout> layout;
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(
                    _dwriteFactory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), _smallFormat.get(), 16'384.0f, 4096.0f, layout.put())) &&
                SUCCEEDED(layout->GetMetrics(&metrics)))
                widest = std::max(widest, std::ceil(metrics.widthIncludingTrailingWhitespace));
        }
    }
    _stalledRateSlotFormat = _smallFormat;
    _stalledRateSlotWidth  = widest;
    return widest;
}

float FileOperationsPopupInternal::FileOperationsPopupState::InPlotGraphLabelRowBottom(const D2D1_RECT_F& plot,
                                                                                       std::wstring_view leading,
                                                                                       std::wstring_view trailing) noexcept
{
    const float top = plot.top + DipsToPixels(3.0f, _dpi);
    if (leading.empty() || ! _dwriteFactory || ! _smallFormat)
        return std::min(plot.bottom, top);
    const float inset         = DipsToPixels(6.0f, _dpi);
    const float gap           = DipsToPixels(8.0f, _dpi);
    const float center        = (plot.left + plot.right) * 0.5f;
    const float leadingWidth  = std::max(1.0f, (trailing.empty() ? plot.right - inset : center - gap * 0.5f) - (plot.left + inset));
    const float trailingWidth = std::max(1.0f, plot.right - inset - (center + gap * 0.5f));
    const auto height         = [&](std::wstring_view text, float width) noexcept
    {
        wil::com_ptr<IDWriteTextLayout> layout;
        DWRITE_TEXT_METRICS metrics{};
        if (text.empty() ||
            FAILED(_dwriteFactory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), _smallFormat.get(), width, 4096.0f, layout.put())) ||
            FAILED(layout->GetMetrics(&metrics)))
            return 0.0f;
        return std::ceil(metrics.height);
    };
    const float tallest = std::max(height(leading, leadingWidth), height(trailing, trailingWidth));
    return std::min(plot.bottom, top + tallest);
}

void FileOperationsPopupInternal::FileOperationsPopupState::PrepareFactPair(
    GraphLabelLayout& pair, std::wstring leading, std::wstring trailing, IDWriteTextFormat* format, float width, float leadingSlot)
{
    if (! _dwriteFactory || ! format || width <= 0.0f)
    {
        pair = {};
        return;
    }
    if (pair.width == width && pair.format.get() == format && pair.rateText == leading && pair.etaText == trailing && pair.leadingSlot == leadingSlot)
        return;
    Debug::Perf::Scope measurement(L"FileOps.Popup.FactPair.MeasureUs");
    GraphLabelLayout next{};
    next.width         = width;
    next.format        = format;
    next.rateText      = std::move(leading);
    next.etaText       = std::move(trailing);
    next.leadingSlot   = leadingSlot;
    const auto prepare = [&](std::wstring_view text, wil::com_ptr<IDWriteTextLayout>& layout, float& textWidth, float& textHeight)
    {
        if (text.empty())
            return true;
        if (FAILED(_dwriteFactory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, width, 4096.0f, layout.put())) ||
            FAILED(layout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP)) || FAILED(layout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR)))
            return false;
        DWRITE_TEXT_METRICS metrics{};
        if (FAILED(layout->GetMetrics(&metrics)))
            return false;
        textWidth  = std::ceil(metrics.widthIncludingTrailingWhitespace);
        textHeight = std::ceil(metrics.height);
        // The measured paragraph owns its height, including explicit newlines and long notes.
        return SUCCEEDED(layout->SetMaxHeight(textHeight));
    };
    if (! prepare(next.rateText, next.rate, next.rateWidth, next.rateHeight) || ! prepare(next.etaText, next.eta, next.etaWidth, next.etaHeight))
    {
        pair = {};
        return;
    }
    next.stacked = std::max(next.rateWidth, leadingSlot) + next.etaWidth + DipsToPixels(8.0f, _dpi) > width && next.rate && next.eta;
    if (next.rate)
        static_cast<void>(next.rate->SetMaxWidth(next.stacked || ! next.eta ? width : std::max(1.0f, next.rateWidth)));
    if (next.eta)
    {
        static_cast<void>(next.eta->SetMaxWidth(next.stacked ? width : std::max(1.0f, next.etaWidth)));
        static_cast<void>(next.eta->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING));
    }
    pair = std::move(next);
}

float FileOperationsPopupInternal::FileOperationsPopupState::FactPairHeight(const GraphLabelLayout& pair) const noexcept
{
    return pair.stacked ? pair.rateHeight + pair.etaHeight + DipsToPixels(2.0f, _dpi) : std::max(pair.rateHeight, pair.etaHeight);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawFactPair(
    const GraphLabelLayout& pair, const D2D1_RECT_F& bounds, uint64_t taskId, std::wstring_view leadingId, std::wstring_view trailingId)
{
    const auto draw = [&](IDWriteTextLayout* layout, std::wstring_view text, const D2D1_RECT_F& rect, std::wstring_view suffix)
    {
        if (! layout)
            return;
        _target->DrawTextLayout(D2D1::Point2F(rect.left, rect.top),
                                layout,
                                pair.format.get() == _headerFormat.get() || suffix == L"Throughput.Label" ? _textBrush.get() : _subTextBrush.get(),
                                D2D1_DRAW_TEXT_OPTIONS_CLIP);
        _hostedTooltipDescriptors.push_back({.bounds       = rect,
                                             .text         = std::wstring(text),
                                             .automationId = std::format(L"FileOperations.Task.{}.{}", taskId, suffix),
                                             .taskId       = taskId,
                                             .liveText     = true});
    };
    draw(pair.rate.get(),
         pair.rateText,
         D2D1::RectF(bounds.left, bounds.top, bounds.left + (pair.stacked || ! pair.eta ? pair.width : pair.rateWidth), bounds.top + pair.rateHeight),
         leadingId);
    const float etaTop = pair.stacked ? bounds.top + pair.rateHeight + DipsToPixels(2.0f, _dpi) : bounds.top;
    draw(pair.eta.get(),
         pair.etaText,
         D2D1::RectF(bounds.right - (pair.stacked ? pair.width : pair.etaWidth), etaTop, bounds.right, etaTop + pair.etaHeight),
         trailingId);
}

void FileOperationsPopupInternal::FileOperationsPopupState::PrepareCardInformation(const std::vector<TaskSnapshot>& snapshot, float width)
{
    _cardInformationLayouts.resize(snapshot.size());
    for (size_t index = 0u; index < snapshot.size(); ++index)
    {
        const auto& task  = snapshot[index];
        auto& information = _cardInformationLayouts[index];
        if (task.kind != TaskSnapshot::Kind::FileOperation || task.conflict.active ||
            IsTaskCollapsedForDisplay(task.taskId, fileOps && fileOps->GetPopupCompactDensity()))
            continue;
        const auto presentation = ResolveWholeTaskProgressPresentation(task);
        std::wstring progress;
        std::wstring percent;
        if (presentation.showWorkFacts && ! TaskUsesSkippedCountingWorkSummary(task))
        {
            if (task.discoveryClosed && task.totalBytes > 0u && (task.operation != FILESYSTEM_COPY || task.discoveredBytesKnown))
                progress = FormatStringResource(
                    nullptr, IDS_FMT_FILEOPS_TRANSFER_AMOUNT, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.totalBytes));
            else if (task.discoveryClosed && task.operation == FILESYSTEM_COPY && task.copiedFileCountKnown && task.discoveredFileCount > 0u)
                progress = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COPIED_FILES, task.copiedFileCount, task.discoveredFileCount);
            else
                progress = FormatStringResource(nullptr, IDS_FMT_FILEOPS_WRITTEN_BYTES, FormatBytesCompact(task.completedBytes));
            if (task.discoveryClosed && presentation.determinate)
                percent = FormatStringResource(
                    nullptr, IDS_FMT_FILEOPS_COMPACT_PERCENT, static_cast<unsigned>(std::lround(Clamp01(static_cast<float>(presentation.fraction)) * 100.0f)));
        }
        PrepareFactPair(information.progress, std::move(progress), std::move(percent), _headerFormat.get(), width);
        std::wstring rate;
        std::wstring eta;
        if (presentation.showThroughputGraph)
        {
            const auto found = _rates.find(task.taskId);
            const RateHistory empty{};
            const auto& history = found != _rates.end() ? found->second : empty;
            rate                = FormatRateLabelText(CurrentBandwidthForGraphMarker(history), history.stallKind);
            if (! task.cancellationRequested && task.discoveryClosed && task.totalBytes > 0u && history.hasSmoothedEta && history.smoothedEtaSeconds > 0.0 &&
                task.completedBytes <= task.totalBytes)
                eta = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA, FormatDurationHms(SaturatingCeilNonNegativeToUint64(history.smoothedEtaSeconds)));
            else if (! task.cancellationRequested && ! task.discoveryClosed && ! task.discoverySkipped && history.hasProvisionalEta &&
                     history.provisionalEtaSeconds > 0.0)
                eta = FormatStringResource(
                    nullptr, IDS_FMT_FILEOPS_ETA_PROVISIONAL, FormatDurationHms(SaturatingCeilNonNegativeToUint64(history.provisionalEtaSeconds)));
            else
                eta = LoadStringResource(nullptr, IDS_FILEOPS_ETA_UNAVAILABLE);
#ifdef ENABLE_TESTS
            if (debugVisualScenario && ! debugGraphLabelText[0].empty() && ! debugGraphLabelText[1].empty())
            {
                rate = debugGraphLabelText[0];
                eta  = debugGraphLabelText[1];
            }
#endif
        }
        // Popup presentation uses only published recursive facts. A selected folder's 0/1
        // counter cannot stand in for the files and folders inside it, especially for Move
        // routes which do not publish type-specific completion receipts.
        // Counting marks provisional totals; plain labels and completion fractions
        // share the same row without adding "found" to either kind of count.
        std::wstring counts;
        if (presentation.showWorkFacts && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) &&
            ! TaskUsesSkippedCountingWorkSummary(task))
        {
            if (! task.discoveryClosed)
                counts = task.discoveredFileCount > 0u || task.discoveredDirectoryCount > 0u
                             ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNTING_FOUND, task.discoveredDirectoryCount, task.discoveredFileCount)
                             : LoadStringResource(nullptr, IDS_FILEOPS_COUNTING_FILES_FOLDERS);
            else if (task.operation == FILESYSTEM_COPY && task.copiedFileCountKnown && task.copiedFolderCountKnown)
                counts = FormatStringResource(nullptr,
                                              IDS_FMT_FILEOPS_FOLDER_FILE_PROGRESS,
                                              task.copiedFolderCount,
                                              task.discoveredDirectoryCount,
                                              task.copiedFileCount,
                                              task.discoveredFileCount);
            else if (task.operation == FILESYSTEM_COPY && task.copiedFileCountKnown)
                counts = FormatStringResource(
                    nullptr, IDS_FMT_FILEOPS_FOUND_FOLDERS_FILE_PROGRESS, task.discoveredDirectoryCount, task.copiedFileCount, task.discoveredFileCount);
            else if (task.discoveredFileCount > 0u || task.discoveredDirectoryCount > 0u)
                counts = FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOLDER_FILE_TOTALS, task.discoveredDirectoryCount, task.discoveredFileCount);
            else
                counts = LoadStringResource(nullptr, IDS_FILEOPS_FILE_FOLDER_COUNTS_UNAVAILABLE);
        }
        PrepareFactPair(information.counts,
                        std::move(counts),
                        std::move(eta),
                        _smallFormat.get(),
                        std::max(1.0f, width - (presentation.showDiscoveryActivity ? DipsToPixels(30.0f, _dpi) : 0.0f)));
        PrepareFactPair(information.throughput,
                        presentation.showThroughputGraph ? LoadStringResource(nullptr, IDS_FILEOPS_THROUGHPUT) : std::wstring{},
                        std::move(rate),
                        _bodyFormat.get(),
                        std::max(1.0f, width - DipsToPixels(28.0f, _dpi)));
        std::wstring concurrency;
        if (presentation.showWorkFacts && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) && task.autoConcurrencyUsed &&
            task.effectiveConcurrencyBudget > 0u)
            concurrency = FormatStringResource(nullptr, IDS_FMT_FILEOPS_PARALLEL_TRANSFERS, task.effectiveConcurrencyBudget);
        PrepareFactPair(information.concurrency, std::move(concurrency), {}, _smallFormat.get(), width);
    }
}
void FileOperationsPopupInternal::FileOperationsPopupState::DrawBandwidthGraph(const D2D1_RECT_F& rect,
                                                                               uint64_t taskId,
                                                                               const RateHistory& history,
                                                                               std::wstring_view currentValueTrailingText,
                                                                               std::wstring_view overlayText,
                                                                               bool rainbowMode,
                                                                               bool perStreamBands,
                                                                               bool showLabels,
                                                                               const D2D1_RECT_F* rateBounds,
                                                                               const D2D1_RECT_F* etaBounds) noexcept
{
    // C4: the hosted ThroughputGraph is the only graph painter; this builds its descriptor.
    HostedGraphDescriptor hostedGraph{};
    hostedGraph.bounds                         = rect;
    hostedGraph.taskId                         = taskId;
    hostedGraph.overlayText                    = overlayText;
    hostedGraph.rainbowMode                    = rainbowMode;
    hostedGraph.perStreamBands                 = perStreamBands;
    hostedGraph.currentBandwidthBytesPerSecond = CurrentBandwidthForGraphMarker(history);
    hostedGraph.currentBandwidthText           = FormatRateLabelText(hostedGraph.currentBandwidthBytesPerSecond, history.stallKind);
    hostedGraph.currentBandwidthTrailingText =
        currentValueTrailingText.empty() ? LoadStringResource(nullptr, IDS_FILEOPS_ETA_UNAVAILABLE) : std::wstring(currentValueTrailingText);
#ifdef ENABLE_TESTS
    if (debugVisualScenario && ! debugGraphLabelText[0].empty() && ! debugGraphLabelText[1].empty())
    {
        hostedGraph.currentBandwidthText         = debugGraphLabelText[0];
        hostedGraph.currentBandwidthTrailingText = debugGraphLabelText[1];
    }
#endif
    const size_t labelIndex = _hostedGraphDescriptors.size();
    if (_graphLabelLayouts.size() <= labelIndex)
        _graphLabelLayouts.resize(labelIndex + 1u);
    auto& labels      = _graphLabelLayouts[labelIndex];
    const float inset = DipsToPixels(6.0f, _dpi);
    const float width = std::max(0.0f, rect.right - rect.left - inset * 2.0f);
    if (showLabels)
        PrepareFactPair(
            labels, hostedGraph.currentBandwidthText, hostedGraph.currentBandwidthTrailingText, _smallFormat.get(), width, StalledRateLabelSlotWidth());
    if (showLabels && labels.rate && labels.eta && labels.width == width && labels.format.get() == _smallFormat.get() &&
        labels.rateText == hostedGraph.currentBandwidthText && labels.etaText == hostedGraph.currentBandwidthTrailingText)
    {
        const float rateWidth       = labels.stacked ? width : labels.rateWidth;
        const float etaWidth        = labels.stacked ? width : labels.etaWidth;
        const float etaTop          = labels.stacked ? rect.top + labels.rateHeight + inset : rect.top;
        hostedGraph.rateLabelBounds = D2D1::RectF(rect.left + inset, rect.top, rect.left + inset + rateWidth, rect.top + labels.rateHeight);
        hostedGraph.etaLabelBounds  = D2D1::RectF(rect.right - inset - etaWidth, etaTop, rect.right - inset, etaTop + labels.etaHeight);
        const float plotTop         = std::max(hostedGraph.rateLabelBounds.bottom, hostedGraph.etaLabelBounds.bottom) + inset;
        if (plotTop < rect.bottom)
        {
            hostedGraph.bounds.top       = plotTop;
            hostedGraph.separateLabelRow = true;
            _target->DrawTextLayout(D2D1::Point2F(hostedGraph.rateLabelBounds.left, hostedGraph.rateLabelBounds.top), labels.rate.get(), _subTextBrush.get());
            _target->DrawTextLayout(D2D1::Point2F(hostedGraph.etaLabelBounds.left, hostedGraph.etaLabelBounds.top), labels.eta.get(), _subTextBrush.get());
            _hostedTooltipDescriptors.push_back({.bounds       = hostedGraph.rateLabelBounds,
                                                 .text         = hostedGraph.currentBandwidthText,
                                                 .automationId = std::format(L"FileOperations.Task.{}.Throughput.Rate", taskId),
                                                 .liveText     = true});
            _hostedTooltipDescriptors.push_back({.bounds       = hostedGraph.etaLabelBounds,
                                                 .text         = hostedGraph.currentBandwidthTrailingText,
                                                 .automationId = std::format(L"FileOperations.Task.{}.Throughput.Eta", taskId),
                                                 .liveText     = true});
        }
    }
    if (! showLabels && rateBounds && etaBounds)
    {
        hostedGraph.separateLabelRow    = true; // Speed is beside Throughput; ETA is beside the counts above the paths.
        hostedGraph.labelsInTaskSummary = true;
        hostedGraph.rateLabelBounds     = *rateBounds;
        hostedGraph.etaLabelBounds      = *etaBounds;
    }
    hostedGraph.labelRowBottom =
        hostedGraph.separateLabelRow
            ? std::max(hostedGraph.rateLabelBounds.bottom, hostedGraph.etaLabelBounds.bottom)
            : InPlotGraphLabelRowBottom(hostedGraph.bounds, hostedGraph.currentBandwidthText, hostedGraph.currentBandwidthTrailingText);
    // The stall band hatches the plot below the labels: when they have no row above the plot, the graph
    // draws them inside it and the band starts under them.
    hostedGraph.stallBandBounds = hostedGraph.bounds;
    if (! hostedGraph.separateLabelRow)
        hostedGraph.stallBandBounds.top =
            std::clamp(hostedGraph.labelRowBottom, hostedGraph.bounds.top, std::max(hostedGraph.bounds.top, hostedGraph.bounds.bottom));
    hostedGraph.samples.reserve(history.count);
    hostedGraph.verificationSamples.reserve(history.count);
    hostedGraph.stalledSamples.reserve(history.count);
    for (size_t sampleOffset = 0u; sampleOffset < history.count; ++sampleOffset)
    {
        const size_t index = (history.writeIndex + RateHistory::kMaxSamples - history.count + sampleOffset) % RateHistory::kMaxSamples;
        hostedGraph.stalledSamples.push_back(history.stalledSamples[index]);
        DxUi::ThroughputGraphSample sample{};
        sample.value          = std::max(0.0f, history.samples[index]);
        sample.hueDegrees     = history.hues[index];
        sample.hueWeightCount = std::min<size_t>(history.hueWeightCounts[index], sample.hueWeights.size());
        for (size_t hueIndex = 0u; hueIndex < sample.hueWeightCount; ++hueIndex)
        {
            sample.hueWeights[hueIndex].hueDegrees = history.hueWeights[index][hueIndex].hue;
            sample.hueWeights[hueIndex].weight     = history.hueWeights[index][hueIndex].weight;
            sample.hueWeights[hueIndex].colorSlot  = history.hueWeights[index][hueIndex].colorSlot;
        }
        if (sample.hueWeightCount > 0u)
        {
            sample.colorSlot = sample.hueWeights[0].colorSlot;
        }
        hostedGraph.samples.push_back(sample);
        hostedGraph.verificationSamples.push_back(std::max(0.0f, history.verificationSamples[index]));
    }
    _hostedGraphDescriptors.push_back(std::move(hostedGraph));
}

std::wstring FileOperationsPopupInternal::FileOperationsPopupState::HostedControlText(const PopupHitTest& hit) const
{
#ifdef ENABLE_TESTS
    if (debugVisualScenario)
    {
        const auto fixture = std::find_if(
            debugVisualScenario->begin(), debugVisualScenario->end(), [&](const TaskSnapshot& task) noexcept { return task.taskId == hit.taskId; });
        if (fixture != debugVisualScenario->end())
        {
            if (hit.kind == PopupHitTest::Kind::TaskPause)
                return LoadStringResource(nullptr, fixture->paused ? IDS_FILEOP_BTN_RESUME : IDS_FILEOP_BTN_PAUSE);
            if (hit.kind == PopupHitTest::Kind::TaskSpeedLimit)
                return FormatSpeedLimitSelectorText(fixture->desiredSpeedLimitBytesPerSecond);
            if (hit.kind == PopupHitTest::Kind::TaskConflictAction)
                return ConflictActionText(static_cast<ConflictAction>(hit.data), static_cast<ConflictBucket>(fixture->conflict.bucket), fixture->operation);
        }
    }
#endif
    switch (hit.kind)
    {
        case PopupHitTest::Kind::FooterCancelAll: return LoadStringResource(nullptr, IDS_FILEOPS_BTN_CANCEL_ALL);
        case PopupHitTest::Kind::FooterPauseResumeAll:
            return LoadStringResource(nullptr, hit.data == kFooterPauseResumeAllPauseAction ? IDS_FILEOPS_BTN_PAUSE_ALL : IDS_FILEOPS_BTN_RESUME_ALL);
        case PopupHitTest::Kind::FooterAutoDismiss:
            return LoadStringResource(nullptr,
                                      fileOps && fileOps->GetAutoDismissSuccess() ? IDS_FILEOPS_CHECK_AUTODISMISS_ON : IDS_FILEOPS_CHECK_AUTODISMISS_OFF);
        case PopupHitTest::Kind::FooterQueueMode:
            return LoadStringResource(nullptr,
                                      hit.data == kFooterQueueModeParallelAction || (hit.data == 0u && fileOps && ! fileOps->GetQueueNewTasks())
                                          ? IDS_FILEOPS_BTN_MODE_PARALLEL
                                          : IDS_FILEOPS_BTN_MODE_QUEUE);
        case PopupHitTest::Kind::FooterDensity:
            return LoadStringResource(nullptr,
                                      fileOps && fileOps->GetPopupCompactDensity() ? IDS_FILEOPS_BTN_DENSITY_COMPACT : IDS_FILEOPS_BTN_DENSITY_EXPANDED);
        case PopupHitTest::Kind::FooterToggleDetails: return L"\u2304";
        case PopupHitTest::Kind::CompletedGroupToggle: return L"\u2304";
        case PopupHitTest::Kind::CompletedGroupClear: return LoadStringResource(nullptr, IDS_FILEOPS_BTN_CLEAR_COMPLETED);
        case PopupHitTest::Kind::TaskToggleCollapse: return L"\u2304";
        case PopupHitTest::Kind::TaskToggleThroughput: return LoadStringResource(nullptr, IDS_FILEOPS_THROUGHPUT);
        case PopupHitTest::Kind::TaskStartNow: return LoadStringResource(nullptr, IDS_FILEOPS_BTN_START_NOW);
        case PopupHitTest::Kind::TaskQueueMoveUp: return L"\u2191";
        case PopupHitTest::Kind::TaskQueueMoveDown: return L"\u2193";
        case PopupHitTest::Kind::TaskPause:
        {
            const auto* task = fileOps ? fileOps->FindTask(hit.taskId) : nullptr;
            return LoadStringResource(nullptr, task && task->IsPaused() ? IDS_FILEOP_BTN_RESUME : IDS_FILEOP_BTN_PAUSE);
        }
        case PopupHitTest::Kind::TaskCancel: return LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);
        case PopupHitTest::Kind::TaskSkip: return LoadStringResource(nullptr, IDS_FILEOPS_BTN_SKIP);
        case PopupHitTest::Kind::TaskSpeedLimit:
        {
            const auto* task = fileOps ? fileOps->FindTask(hit.taskId) : nullptr;
            return task ? FormatSpeedLimitSelectorText(task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire))
                        : LoadStringResource(nullptr, IDS_FILEOP_BTN_SPEED_LIMIT);
        }
        case PopupHitTest::Kind::TaskShowLog: return LoadStringResource(nullptr, IDS_FILEOP_BTN_SHOW_LOG);
        case PopupHitTest::Kind::TaskExportIssues: return LoadStringResource(nullptr, IDS_FILEOP_BTN_EXPORT_ISSUES);
        case PopupHitTest::Kind::TaskCompletedMore: return L"\u2026";
        case PopupHitTest::Kind::TaskConflictMore: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_MORE);
        case PopupHitTest::Kind::TaskConflictInspectSource: return LoadStringResource(nullptr, IDS_FILEOPS_INSPECTION_SOURCE);
        case PopupHitTest::Kind::TaskConflictInspectDestination: return LoadStringResource(nullptr, IDS_FILEOPS_INSPECTION_DESTINATION);
        case PopupHitTest::Kind::TaskConflictToggleApplyToAll: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_APPLY_TO_ALL);
        case PopupHitTest::Kind::TaskConflictAction:
        {
            ConflictBucket bucket         = ConflictBucket::Unknown;
            FileSystemOperation operation = FILESYSTEM_COPY;
            if (const auto* task = fileOps ? fileOps->FindTask(hit.taskId) : nullptr)
            {
                std::scoped_lock lock(task->_conflictArbiter.mutex);
                bucket    = task->_conflictArbiter.prompt.bucket;
                operation = task->GetOperation();
            }
            return ConflictActionText(static_cast<ConflictAction>(hit.data), bucket, operation);
        }
        case PopupHitTest::Kind::TaskIssues: return LoadStringResource(nullptr, IDS_FILEOPS_BTN_ISSUES);
        case PopupHitTest::Kind::TaskDismiss: return LoadStringResource(nullptr, IDS_FILEOP_BTN_DISMISS);
        case PopupHitTest::Kind::FooterOptions:
            return _controlsHost.HasFluentIconFont() ? std::wstring(1u, FluentIcons::kMore) : std::wstring(1u, FluentIcons::kFallbackMore);
        case PopupHitTest::Kind::None: break;
    }
    return {};
}

std::wstring FileOperationsPopupInternal::FileOperationsPopupState::HostedControlAccessibleName(const PopupHitTest& hit) const
{
    if (hit.kind == PopupHitTest::Kind::FooterToggleDetails)
    {
        return LoadStringResource(nullptr, IDS_FILEOPS_TOGGLE_DETAILS);
    }
    if (hit.kind == PopupHitTest::Kind::CompletedGroupToggle)
    {
        return LoadStringResource(nullptr, IDS_FILEOPS_COMPLETED_GROUP_ACCESSIBLE);
    }
    if (hit.kind == PopupHitTest::Kind::TaskToggleCollapse)
    {
        return LoadStringResource(nullptr, IDS_FILEOPS_TOGGLE_TASK_DETAILS);
    }
    if (hit.kind == PopupHitTest::Kind::TaskQueueMoveUp)
    {
        return LoadStringResource(nullptr, IDS_FILEOPS_QUEUE_MOVE_UP);
    }
    if (hit.kind == PopupHitTest::Kind::TaskQueueMoveDown)
    {
        return LoadStringResource(nullptr, IDS_FILEOPS_QUEUE_MOVE_DOWN);
    }
    if (hit.kind == PopupHitTest::Kind::TaskCompletedMore || hit.kind == PopupHitTest::Kind::TaskConflictMore)
    {
        return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_MORE);
    }
    if (hit.kind == PopupHitTest::Kind::FooterOptions)
    {
        return LoadStringResource(nullptr, IDS_FILEOPS_FOOTER_OPTIONS);
    }
    return HostedControlText(hit);
}

FileOperationsPopupInternal::FileOperationsPopupState::ConflictBodyLayout* FileOperationsPopupInternal::FileOperationsPopupState::FindConflictBody(
    uint64_t taskId) noexcept
{
    const auto found = std::ranges::find(_conflictBodies, taskId, &ConflictBodyLayout::taskId);
    return found == _conflictBodies.end() ? nullptr : &*found;
}

const FileOperationsPopupInternal::FileOperationsPopupState::ConflictBodyLayout* FileOperationsPopupInternal::FileOperationsPopupState::FindConflictBody(
    uint64_t taskId) const noexcept
{
    const auto found = std::ranges::find(_conflictBodies, taskId, &ConflictBodyLayout::taskId);
    return found == _conflictBodies.end() ? nullptr : &*found;
}

void FileOperationsPopupInternal::FileOperationsPopupState::EnsureHostedFocusVisible(HWND hwnd, DxUi::Control* control) noexcept
{
    // popup-b#3: a programmatic restore after its control was retired never undoes the user's scroll.
    if (! control || ! _controlsBody || _restoringHostedFocus || ! hostLifetime.lock())
        return;
    const auto button     = std::ranges::find(_hostedButtons, control);
    const size_t index    = static_cast<size_t>(std::distance(_hostedButtons.begin(), button));
    const bool taskButton = index < _hostedButtonHits.size() &&
                            (_hostedButtonHits[index].taskId != 0u || _hostedButtonHits[index].kind == PopupHitTest::Kind::CompletedGroupToggle ||
                             _hostedButtonHits[index].kind == PopupHitTest::Kind::CompletedGroupClear);
    if (! taskButton && std::ranges::find(_hostedTooltipRegions, control) == _hostedTooltipRegions.end())
        return;
    auto bounds                     = control->GetBounds();
    const float scale               = static_cast<float>(_dpi) / 96.0f;
    const uint64_t comparisonTaskId = index < _hostedButtonHits.size() && (_hostedButtonHits[index].kind == PopupHitTest::Kind::TaskConflictInspectSource ||
                                                                           _hostedButtonHits[index].kind == PopupHitTest::Kind::TaskConflictInspectDestination)
                                          ? _hostedButtonHits[index].taskId
                                          : 0u;
    for (auto& body : _conflictBodies)
    {
        if (! body.visible || ! body.panel || comparisonTaskId != body.taskId)
            continue;
        const float visibleTop     = bounds.top - body.offsetDip - 4.0f;
        const float visibleBottom  = bounds.bottom - body.offsetDip + 4.0f;
        const float viewportTop    = body.viewport.top / scale;
        const float viewportBottom = body.viewport.bottom / scale;
        const float adjustment = visibleTop < viewportTop ? visibleTop - viewportTop : (visibleBottom > viewportBottom ? visibleBottom - viewportBottom : 0.0f);
        body.panel->SetScrollOffset(body.offsetDip + adjustment);
        bounds.top -= body.offsetDip;
        bounds.bottom -= body.offsetDip;
        break;
    }
    // Bounds belong to the last painted scroll origin. Another focus request
    // may already have changed the pending native offset without repainting.
    // Translate before testing visibility so successive focus requests do not
    // accumulate offsets from different frames or undo an already visible one.
    const float pendingScrollTranslation = _scrollY - static_cast<float>(_scrollPos);
    // Focus chrome extends outside the control bounds; reveal that too.
    const float top    = (bounds.top - 4.0f) * scale + pendingScrollTranslation;
    const float bottom = (bounds.bottom + 4.0f) * scale + pendingScrollTranslation;
    const float delta =
        top < _listViewportRect.top ? top - _listViewportRect.top : (bottom > _listViewportRect.bottom ? bottom - _listViewportRect.bottom : 0.0f);
    const int maximum  = static_cast<int>(std::ceil(std::max(0.0f, _contentHeight - (_listViewportRect.bottom - _listViewportRect.top))));
    const int position = std::clamp(_scrollPos + static_cast<int>(std::ceil(delta)), 0, maximum);
    if (position != _scrollPos)
    {
        _scrollPos = position;
        Invalidate(hwnd); // Reflow through the normal frame; never recurse into painting from focus callbacks.
    }
}

FileOperationsPopupInternal::HostedTaskSpaceKeyRoute FileOperationsPopupInternal::ResolveHostedTaskSpaceKeyRoute(PopupHitTest::Kind kind,
                                                                                                                 bool controlEnabled,
                                                                                                                 bool modified) noexcept
{
    // Decision, inspection, scope and disclosure controls own Space, as do Cancel, Skip, the speed
    // selector and every other enabled button (popup-b#4). Only the task's own Pause/Resume button
    // or a control that cannot act keeps the pause shortcut, and a Ctrl/Alt chord is never it.
    if (kind == PopupHitTest::Kind::TaskConflictAction || kind == PopupHitTest::Kind::TaskConflictMore ||
        kind == PopupHitTest::Kind::TaskConflictInspectSource || kind == PopupHitTest::Kind::TaskConflictInspectDestination ||
        kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll || kind == PopupHitTest::Kind::TaskToggleThroughput ||
        kind == PopupHitTest::Kind::TaskToggleCollapse)
        return HostedTaskSpaceKeyRoute::ControlActivation;
    if (modified)
        return HostedTaskSpaceKeyRoute::Ignore;
    return kind == PopupHitTest::Kind::TaskPause || ! controlEnabled ? HostedTaskSpaceKeyRoute::PauseOrResume : HostedTaskSpaceKeyRoute::ControlActivation;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::HandleHostedKeyDown(HWND hwnd, UINT virtualKey, LPARAM keyData, bool systemKey) noexcept
{
    // A held activation must not act on a newly published decision or repeatedly
    // toggle/cancel a task. Navigation keys still retain their normal autorepeat.
    constexpr LPARAM previousKeyState = LPARAM{1} << 30;
    if ((keyData & previousKeyState) != 0 && (virtualKey == VK_ESCAPE || virtualKey == VK_RETURN || virtualKey == VK_SPACE || virtualKey == VK_DELETE))
        return true;

    if (! _pathDetailId.empty())
    {
        if (virtualKey == VK_ESCAPE)
        {
            ClosePathDetail(hwnd, true);
            return true;
        }
        // The read-only editor owns navigation and copy. Never fall through to
        // a hidden task's Delete/Space/default-decision accelerator.
        return false;
    }
    auto* const focused  = _controlsHost.GetFocusControl();
    const auto focusedIt = std::ranges::find(_hostedButtons, focused);
    if (focusedIt == _hostedButtons.end())
    {
        return false;
    }

    const size_t index = static_cast<size_t>(std::distance(_hostedButtons.begin(), focusedIt));
    if (index >= _hostedButtonHits.size())
    {
        return false;
    }
    const PopupHitTest hit = _hostedButtonHits[index];

    if (virtualKey == VK_LEFT || virtualKey == VK_UP || virtualKey == VK_RIGHT || virtualKey == VK_DOWN)
    {
        if (hit.taskId == 0u)
        {
            return false;
        }
        const int direction = (virtualKey == VK_LEFT || virtualKey == VK_UP) ? -1 : 1;
        for (size_t offset = 1u; offset < _hostedButtonHits.size(); ++offset)
        {
            const ptrdiff_t candidateIndex = static_cast<ptrdiff_t>(index) + direction * static_cast<ptrdiff_t>(offset);
            if (candidateIndex < 0 || candidateIndex >= static_cast<ptrdiff_t>(_hostedButtonHits.size()))
            {
                break;
            }
            const size_t candidate = static_cast<size_t>(candidateIndex);
            if (_hostedButtonHits[candidate].taskId == hit.taskId)
            {
                _controlsHost.SetFocusControl(_hostedButtons[candidate]);
                return true;
            }
        }
        return true;
    }

    // Ctrl/Alt chords and system keys are never the task's Delete or Space shortcut.
    const bool modified = systemKey || GetKeyState(VK_CONTROL) < 0 || GetKeyState(VK_MENU) < 0;
    if (virtualKey == VK_DELETE && hit.taskId != 0u)
    {
        if (! modified && IsTaskDecisionAvailable(hit.taskId))
        {
            if (auto* task = fileOps->FindTask(hit.taskId))
            {
                task->RequestCancel();
                fileOps->CancelConflictInspectionForTask(hit.taskId);
                Invalidate(hwnd);
            }
        }
        return true;
    }

    // Decision and inspection controls own their keyboard activation. In particular,
    // Space must not pause a transfer instead of inspecting its conflicting item.
    if (hit.kind == PopupHitTest::Kind::TaskConflictAction || hit.kind == PopupHitTest::Kind::TaskConflictMore ||
        hit.kind == PopupHitTest::Kind::TaskConflictInspectSource || hit.kind == PopupHitTest::Kind::TaskConflictInspectDestination)
        return false;

    // Repeat scope, disclosures and enabled action buttons own Space through their hosted
    // controls; activating them must not also change the task's pause state.
    if (virtualKey == VK_SPACE && hit.taskId != 0u)
    {
        const HostedTaskSpaceKeyRoute route = ResolveHostedTaskSpaceKeyRoute(hit.kind, focused->IsEnabled(), modified);
        if (route == HostedTaskSpaceKeyRoute::Ignore)
            return true;
        if (route == HostedTaskSpaceKeyRoute::ControlActivation)
            return false;
        if (auto* task = fileOps ? fileOps->FindTask(hit.taskId) : nullptr; task && task->HasStarted())
        {
            task->TogglePause();
            Invalidate(hwnd);
            return true;
        }
    }
    return false;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::SubmitPublishedConflictEscapeAction(HWND hwnd) noexcept
{
    if (! fileOps || _conflictEscapeTaskId == 0u)
        return false;
    auto* task = fileOps->FindTask(_conflictEscapeTaskId);
    if (! task || ! task->SubmitConflictDecision(
                      _conflictEscapeGeneration, ConflictActionFromRaw(_conflictEscapeAction), std::nullopt, FileOperations::ConflictActionPlacement::Escape))
        return false;
    Invalidate(hwnd);
    return true;
}
HRESULT FileOperationsPopupInternal::FileOperationsPopupState::MeasureConflictActions(const TaskSnapshot& task,
                                                                                      float widthPx,
                                                                                      bool rightToLeft,
                                                                                      ConflictActionsLayout& cached) noexcept
{
    const size_t primaryCount = std::min(task.conflict.primaryActionCount, task.conflict.primaryActions.size());
    const bool hasMore        = task.conflict.overflowActionCount != 0u;
    const size_t count        = primaryCount + (hasMore ? 1u : 0u);
    std::array<std::wstring, TaskSnapshot::kMaxConflictActions + 1u> labels;
    for (size_t index = 0u; index < primaryCount; ++index)
        labels[index] =
            ConflictActionText(ConflictActionFromRaw(task.conflict.primaryActions[index]), static_cast<ConflictBucket>(task.conflict.bucket), task.operation);
    if (hasMore)
        labels[primaryCount] = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_MORE);
    auto* format = _controlsHost.GetTextFormat(DxUi::FontRole::Body,
                                               DWRITE_TEXT_ALIGNMENT_CENTER,
                                               DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                                               true,
                                               rightToLeft ? DWRITE_READING_DIRECTION_RIGHT_TO_LEFT : DWRITE_READING_DIRECTION_LEFT_TO_RIGHT);
    if (! _controlsHost.GetHwnd())
        return E_PENDING;
    const std::wstring scopeLabel = task.conflict.applyToAllEligible ? LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_APPLY_TO_ALL) : std::wstring{};
    const float widthDip          = _controlsHost.PixelsToDip(widthPx);
    if (cached.format.get() == format && cached.widthDip == widthDip && cached.dpi == _dpi && cached.rightToLeft == rightToLeft && cached.count == count &&
        cached.labels == labels && cached.scopeLabel == scopeLabel)
        return S_OK;
    std::array<D2D1_SIZE_F, TaskSnapshot::kMaxConflictActions + 1u> sizes{};
    float scopeHeightDip = 0.0f;
    std::array<D2D1_RECT_F, TaskSnapshot::kMaxConflictActions + 1u> bounds{};
    DxUi::MeasuredActionLayout arranged{};
    const HRESULT hr = [&]() noexcept -> HRESULT
    {
        if (! format || ! _dwriteFactory)
            return E_FAIL;
        if (widthDip <= 36.0f)
            return E_INVALIDARG;
#ifdef ENABLE_TESTS
        if (ConsumeForcedFailure(g_fileOperationsLayoutMeasurementForcedFailures))
            return E_OUTOFMEMORY;
#endif
        const float scale = static_cast<float>(_dpi) / 96.0f;
        for (size_t index = 0u; index < count; ++index)
        {
            wil::com_ptr<IDWriteTextLayout> text;
            RETURN_IF_FAILED(_dwriteFactory->CreateTextLayout(
                labels[index].data(), static_cast<UINT32>(labels[index].size()), format, widthDip - 24.0f, 4096.0f, text.put()));
            DWRITE_TEXT_METRICS metrics{};
            RETURN_IF_FAILED(text->GetMetrics(&metrics));
            const float horizontalPadding = index == primaryCount ? 52.0f : 24.0f;
            sizes[index] =
                D2D1::SizeF(std::min(widthDip, std::max(80.0f, std::ceil(metrics.widthIncludingTrailingWhitespace * scale) / scale + horizontalPadding)),
                            std::max(32.0f, std::ceil(metrics.height * scale) / scale + 16.0f));
        }
        if (! scopeLabel.empty())
        {
            wil::com_ptr<IDWriteTextLayout> scope;
            RETURN_IF_FAILED(
                _dwriteFactory->CreateTextLayout(scopeLabel.data(), static_cast<UINT32>(scopeLabel.size()), format, widthDip - 36.0f, 4096.0f, scope.put()));
            DWRITE_TEXT_METRICS metrics{};
            RETURN_IF_FAILED(scope->GetMetrics(&metrics));
            scopeHeightDip = std::max(24.0f, std::ceil(metrics.height * scale) / scale + 6.0f);
        }
        return DxUi::ArrangeMeasuredActions(std::span(sizes).first(count),
                                            widthDip,
                                            D2D1::SizeF(8.0f, 6.0f),
                                            bounds,
                                            arranged,
                                            rightToLeft ? DxUi::FlowDirection::RightToLeft : DxUi::FlowDirection::LeftToRight);
    }();
    if (FAILED(hr))
    {
        // x-arch#2: a failed measurement or fit check must not hide a decision. Each action keeps
        // a full-width 32-DIP row; the empty format key measures again on the next paint.
        cached = {};
        for (size_t index = 0u; index < count; ++index)
        {
            const float top      = static_cast<float>(index) * (32.0f + 6.0f);
            cached.bounds[index] = D2D1::RectF(0.0f, top, std::max(1.0f, widthDip), top + 32.0f);
            cached.heightDip     = top + 32.0f;
        }
        cached.scopeLabel     = scopeLabel;
        cached.scopeHeightDip = scopeLabel.empty() ? 0.0f : 24.0f;
        cached.labels         = std::move(labels);
        cached.widthDip       = widthDip;
        cached.dpi            = _dpi;
        cached.count          = count;
        cached.rightToLeft    = rightToLeft;
        return hr;
    }
    cached.scopeLabel     = scopeLabel;
    cached.scopeHeightDip = scopeHeightDip;
    cached.labels         = std::move(labels);
    cached.bounds         = bounds;
    cached.format         = format;
    cached.widthDip       = widthDip;
    cached.heightDip      = arranged.heightDip;
    cached.dpi            = _dpi;
    cached.count          = count;
    cached.rightToLeft    = rightToLeft;
    return S_OK;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ReconcileHostedControls(HWND hwnd, bool rightToLeft) noexcept
{
    // popup-b#1: the descriptors changed shape (a stream row, a card entering the viewport, a new
    // decision). Keep every control whose identity is still published, create only the new ones and
    // retire only the vanished ones, so a pressed, hovered or focused control survives untouched.
    ++_hostedTreeCounters.reconciles;
    const auto sameHit = [](const PopupHitTest& left, const PopupHitTest& right) noexcept
    { return left.kind == right.kind && left.taskId == right.taskId && left.data == right.data && left.promptGeneration == right.promptGeneration; };
    const auto comparisonForHit = [](const PopupHitTest& hit) noexcept -> uint64_t
    { return hit.kind == PopupHitTest::Kind::TaskConflictInspectSource || hit.kind == PopupHitTest::Kind::TaskConflictInspectDestination ? hit.taskId : 0u; };
    std::vector<RetiredHostedControl> retired;
    uint64_t created = 0u;
    uint64_t revived = 0u;
    using Role       = RetiredHostedControl::Role;
    // A control parked during a capture whose identity is published again under the same parent is
    // shown again instead of recreated, so a card that flaps while a press is held neither grows the
    // tree nor pays for new controls (every new control republishes the accessibility snapshot).
    const auto revive = [&](Role role, const DxUi::Panel* parent, const auto& sameIdentity) noexcept -> DxUi::Control*
    {
        const auto parked = std::ranges::find_if(_retiredHostedControls, [&](const RetiredHostedControl& entry) noexcept {
            return entry.role == role && entry.parent == parent && sameIdentity(entry);
        });
        if (parked == _retiredHostedControls.end())
            return nullptr;
        DxUi::Control* const control = parked->control;
        _retiredHostedControls.erase(parked);
        control->SetVisible(true);
        ++revived;
        return control;
    };
    const auto sameAutomationId = [](std::wstring_view automationId) noexcept
    { return [automationId](const RetiredHostedControl& entry) noexcept { return entry.control->GetAccessibleAutomationId() == automationId; }; };

    // Conflict comparison bodies: one retained panel per visible body, created eagerly so a visible
    // body without comparison children never asks for another pass (popup-b#5).
    std::vector<DxUi::ScrollPanel*> bodyPanels;
    std::vector<uint64_t> bodyIds;
    for (auto& body : _conflictBodies)
    {
        if (! body.visible)
        {
            body.panel = nullptr; // Retired below with the other unreferenced panels.
            continue;
        }
        if (! body.panel)
        {
            auto* panel = _controlsBody->AddChild<FileOperationsHostedScrollPanel>();
            panel->SetScrollStepDip(36.0f);
            panel->SetAccessibleAutomationId(std::format(L"FileOperations.Task.{}.Conflict.Comparison", body.taskId));
            panel->SetOnScrollChanged([this, hwnd, taskId = body.taskId](float offset)
            {
                if (_syncingConflictBodies || ! hostLifetime.lock())
                    return;
                if (auto* current = FindConflictBody(taskId); current && current->visible)
                {
                    current->offsetDip = offset;
                    Invalidate(hwnd);
                }
            });
            body.panel = panel;
            ++created;
        }
        bodyPanels.push_back(body.panel);
        bodyIds.push_back(body.taskId);
    }
    for (DxUi::ScrollPanel* panel : _hostedConflictBodyPanels)
        if (std::ranges::find(bodyPanels, panel) == bodyPanels.end())
            retired.push_back({panel, _controlsBody});
    const auto parentFor = [&](uint64_t comparisonTaskId) noexcept -> DxUi::Panel*
    {
        const auto* body = FindConflictBody(comparisonTaskId);
        return body && body->visible && body->panel ? static_cast<DxUi::Panel*>(body->panel) : static_cast<DxUi::Panel*>(_controlsBody);
    };

    // Buttons keep their action identity (kind, task, data, prompt generation) and parent.
    std::vector<DxUi::Button*> buttons(_buttons.size(), nullptr);
    std::vector<DxUi::Panel*> buttonParents(_buttons.size(), nullptr);
    std::vector<uint8_t> buttonKept(_hostedButtons.size(), 0u);
    for (size_t index = 0u; index < _buttons.size(); ++index)
    {
        const PopupHitTest hit = _buttons[index].hit;
        const bool inBody = hit.taskId != 0u || hit.kind == PopupHitTest::Kind::CompletedGroupToggle || hit.kind == PopupHitTest::Kind::CompletedGroupClear;
        DxUi::Panel* const parent = inBody ? parentFor(comparisonForHit(hit)) : _controlsRoot;
        const auto reusable       = [&](size_t old) noexcept
        {
            return old < _hostedButtons.size() && buttonKept[old] == 0u && _hostedButtons[old] && old < _hostedButtonParents.size() &&
                   _hostedButtonParents[old] == parent && old < _hostedButtonHits.size() && sameHit(_hostedButtonHits[old], hit);
        };
        size_t match = reusable(index) ? index : _hostedButtons.size();
        for (size_t old = 0u; match == _hostedButtons.size() && old < _hostedButtons.size(); ++old)
            if (reusable(old))
                match = old;
        if (match < _hostedButtons.size())
        {
            buttonKept[match] = 1u;
            buttons[index]    = _hostedButtons[match];
        }
        else if (DxUi::Control* const parked =
                     revive(Role::Button, parent, [&hit, &sameHit](const RetiredHostedControl& entry) noexcept { return sameHit(entry.hit, hit); }))
        {
            buttons[index] = static_cast<DxUi::Button*>(parked);
        }
        else
        {
            const std::wstring controlText    = HostedControlText(hit);
            const std::wstring accessibleName = HostedControlAccessibleName(hit);
            DxUi::Button* const button        = hit.kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll ? parent->AddChild<DxUi::Checkbox>(controlText)
                                                                                                             : parent->AddChild<DxUi::Button>(controlText);
            button->SetAccessibleName(accessibleName);
            button->SetAccessibleAutomationId(PopupActionAutomationId(hit));
            button->SetTooltipText(accessibleName != controlText ? accessibleName : std::wstring{});
            button->SetDensity(DxUi::Density::Compact);
            // A retired button stays alive while a capture is active and a held press can still
            // release on it; only a published button acts.
            const auto activate = [this, hwnd, hit, button]
            {
                if (IsWindow(hwnd) != FALSE && hostLifetime.lock() && _pathDetailId.empty() &&
                    std::ranges::find(_hostedButtons, button) != _hostedButtons.end())
                {
#ifdef ENABLE_TESTS
                    ++_debugHostedActionCount;
#endif
                    static_cast<void>(OnActivatedHit(hwnd, hit));
                }
            };
            if (hit.kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll)
                static_cast<DxUi::Checkbox*>(button)->SetOnToggled([activate](bool) { activate(); });
            else if (OpensButtonFlyout(hit.kind))
                button->SetOnDropDownClick(activate);
            else
                button->SetOnClick(activate);
            buttons[index] = button;
            ++created;
        }
        buttonParents[index] = parent;
    }
    for (size_t old = 0u; old < _hostedButtons.size(); ++old)
        if (buttonKept[old] == 0u && _hostedButtons[old])
            retired.push_back({_hostedButtons[old],
                               old < _hostedButtonParents.size() ? _hostedButtonParents[old] : _controlsBody,
                               Role::Button,
                               old < _hostedButtonHits.size() ? _hostedButtonHits[old] : PopupHitTest{}});

    // Progress bars and graphs keep their role/task automation ID.
    const auto progressId = [](uint64_t taskId)
    { return taskId == 0u ? std::wstring(L"FileOperations.Aggregate.Progress") : std::format(L"FileOperations.Task.{}.Progress", taskId); };
    std::vector<DxUi::ProgressBar*> progressBars(_hostedProgressDescriptors.size(), nullptr);
    std::vector<ScalarAnimation> progressAnimations(_hostedProgressDescriptors.size());
    std::vector<uint8_t> progressKept(_hostedProgressBars.size(), 0u);
    for (size_t index = 0u; index < _hostedProgressDescriptors.size(); ++index)
    {
        const uint64_t taskId = _hostedProgressDescriptors[index].taskId;
        const std::wstring id = progressId(taskId);
        for (size_t old = 0u; old < _hostedProgressBars.size(); ++old)
        {
            if (progressKept[old] == 0u && _hostedProgressBars[old] && _hostedProgressBars[old]->GetAccessibleAutomationId() == id)
            {
                progressKept[old]         = 1u;
                progressBars[index]       = _hostedProgressBars[old];
                progressAnimations[index] = old < _hostedProgressAnimations.size() ? _hostedProgressAnimations[old] : ScalarAnimation{};
                break;
            }
        }
        DxUi::Panel* const parent = taskId == 0u ? static_cast<DxUi::Panel*>(_controlsRoot) : static_cast<DxUi::Panel*>(_controlsBody);
        if (! progressBars[index])
            progressBars[index] = static_cast<DxUi::ProgressBar*>(revive(Role::Progress, parent, sameAutomationId(id)));
        if (! progressBars[index])
        {
            progressBars[index] = parent->AddChild<DxUi::ProgressBar>();
            progressBars[index]->SetAccessibleAutomationId(id);
            ++created;
        }
    }
    for (size_t old = 0u; old < _hostedProgressBars.size(); ++old)
        if (progressKept[old] == 0u && _hostedProgressBars[old])
            retired.push_back({_hostedProgressBars[old],
                               _hostedProgressBars[old]->GetAccessibleAutomationId() == progressId(0u) ? static_cast<DxUi::Panel*>(_controlsRoot)
                                                                                                       : static_cast<DxUi::Panel*>(_controlsBody),
                               Role::Progress});

    // I35 (owner report 2026-10-13): each bar's verification-half marks keep the bar's identity and parent.
    const auto marksId = [&](uint64_t taskId) { return progressId(taskId) + L".VerificationMarks"; };
    std::vector<DxUi::Control*> progressMarks(_hostedProgressDescriptors.size(), nullptr);
    std::vector<uint8_t> marksKept(_hostedProgressVerificationMarks.size(), 0u);
    for (size_t index = 0u; index < _hostedProgressDescriptors.size(); ++index)
    {
        const uint64_t taskId = _hostedProgressDescriptors[index].taskId;
        const std::wstring id = marksId(taskId);
        for (size_t old = 0u; old < _hostedProgressVerificationMarks.size(); ++old)
        {
            if (marksKept[old] == 0u && _hostedProgressVerificationMarks[old] && _hostedProgressVerificationMarks[old]->GetAccessibleAutomationId() == id)
            {
                marksKept[old]       = 1u;
                progressMarks[index] = _hostedProgressVerificationMarks[old];
                break;
            }
        }
        DxUi::Panel* const parent = taskId == 0u ? static_cast<DxUi::Panel*>(_controlsRoot) : static_cast<DxUi::Panel*>(_controlsBody);
        if (! progressMarks[index])
            progressMarks[index] = revive(Role::ProgressVerificationMarks, parent, sameAutomationId(id));
        if (! progressMarks[index])
        {
            progressMarks[index] = parent->AddChild<FileOperationsVerificationHalfMarks>();
            progressMarks[index]->SetAccessibleAutomationId(id);
            ++created;
        }
    }
    for (size_t old = 0u; old < _hostedProgressVerificationMarks.size(); ++old)
        if (marksKept[old] == 0u && _hostedProgressVerificationMarks[old])
            retired.push_back({_hostedProgressVerificationMarks[old],
                               _hostedProgressVerificationMarks[old]->GetAccessibleAutomationId() == marksId(0u) ? static_cast<DxUi::Panel*>(_controlsRoot)
                                                                                                                 : static_cast<DxUi::Panel*>(_controlsBody),
                               Role::ProgressVerificationMarks});

    std::vector<DxUi::ThroughputGraph*> graphs(_hostedGraphDescriptors.size(), nullptr);
    std::vector<uint8_t> graphKept(_hostedGraphs.size(), 0u);
    for (size_t index = 0u; index < _hostedGraphDescriptors.size(); ++index)
    {
        const std::wstring id = std::format(L"FileOperations.Task.{}.ThroughputGraph", _hostedGraphDescriptors[index].taskId);
        for (size_t old = 0u; old < _hostedGraphs.size(); ++old)
        {
            if (graphKept[old] == 0u && _hostedGraphs[old] && _hostedGraphs[old]->GetAccessibleAutomationId() == id)
            {
                graphKept[old] = 1u;
                graphs[index]  = _hostedGraphs[old];
                break;
            }
        }
        if (! graphs[index])
            graphs[index] = static_cast<DxUi::ThroughputGraph*>(revive(Role::Graph, _controlsBody, sameAutomationId(id)));
        if (! graphs[index])
        {
            graphs[index] = _controlsBody->AddChild<DxUi::ThroughputGraph>();
            graphs[index]->SetAccessibleAutomationId(id);
            ++created;
        }
    }
    for (size_t old = 0u; old < _hostedGraphs.size(); ++old)
        if (graphKept[old] == 0u && _hostedGraphs[old])
            retired.push_back({_hostedGraphs[old], _controlsBody, Role::Graph});

    // I35 R10.2: each graph's stall band keeps the graph's task identity.
    std::vector<DxUi::Control*> stallBands(_hostedGraphDescriptors.size(), nullptr);
    std::vector<uint8_t> stallBandKept(_hostedGraphStallBands.size(), 0u);
    for (size_t index = 0u; index < _hostedGraphDescriptors.size(); ++index)
    {
        const std::wstring id = std::format(L"FileOperations.Task.{}.ThroughputGraph.StallBand", _hostedGraphDescriptors[index].taskId);
        for (size_t old = 0u; old < _hostedGraphStallBands.size(); ++old)
        {
            if (stallBandKept[old] == 0u && _hostedGraphStallBands[old] && _hostedGraphStallBands[old]->GetAccessibleAutomationId() == id)
            {
                stallBandKept[old] = 1u;
                stallBands[index]  = _hostedGraphStallBands[old];
                break;
            }
        }
        if (! stallBands[index])
            stallBands[index] = revive(Role::GraphStallBand, _controlsBody, sameAutomationId(id));
        if (! stallBands[index])
        {
            stallBands[index] = _controlsBody->AddChild<FileOperationsGraphStallBand>();
            stallBands[index]->SetAccessibleAutomationId(id);
            ++created;
        }
    }
    for (size_t old = 0u; old < _hostedGraphStallBands.size(); ++old)
        if (stallBandKept[old] == 0u && _hostedGraphStallBands[old])
            retired.push_back({_hostedGraphStallBands[old], _controlsBody, Role::GraphStallBand});

    // Tooltip, status and path regions keep their automation ID, kind and parent.
    std::vector<DxUi::Control*> regions(_hostedTooltipDescriptors.size(), nullptr);
    std::vector<DxUi::Panel*> regionParents(_hostedTooltipDescriptors.size(), nullptr);
    std::vector<uint8_t> regionKept(_hostedTooltipRegions.size(), 0u);
    for (size_t index = 0u; index < _hostedTooltipDescriptors.size(); ++index)
    {
        const HostedTooltipDescriptor& descriptor = _hostedTooltipDescriptors[index];
        const bool pathRegion                     = ! descriptor.copyPath.empty();
        DxUi::Panel* const parent                 = parentFor(descriptor.comparisonTaskId);
        const auto reusable                       = [&](size_t old) noexcept
        {
            return old < _hostedTooltipRegions.size() && regionKept[old] == 0u && _hostedTooltipRegions[old] && old < _hostedTooltipParents.size() &&
                   _hostedTooltipParents[old] == parent && _hostedTooltipRegions[old]->IsFocusable() == pathRegion &&
                   _hostedTooltipRegions[old]->GetAccessibleAutomationId() == descriptor.automationId;
        };
        size_t match = reusable(index) ? index : _hostedTooltipRegions.size();
        for (size_t old = 0u; match == _hostedTooltipRegions.size() && old < _hostedTooltipRegions.size(); ++old)
            if (reusable(old))
                match = old;
        if (match < _hostedTooltipRegions.size())
        {
            regionKept[match] = 1u;
            regions[index]    = _hostedTooltipRegions[match];
        }
        else if (DxUi::Control* const parked = revive(pathRegion ? Role::PathRegion : Role::Region, parent, sameAutomationId(descriptor.automationId)))
        {
            regions[index] = parked;
        }
        else
        {
            // popup-a#2: an informational region is a status; a path row is an invokable Button.
            DxUi::Control* const region = pathRegion ? static_cast<DxUi::Control*>(parent->AddChild<FileOperationsPathRegion>())
                                                     : static_cast<DxUi::Control*>(parent->AddChild<DxUi::Label>());
            if (! pathRegion)
                region->SetAccessibilityRole(DxUi::AccessibilityRole::Status);
            region->SetAccessibleAutomationId(descriptor.automationId);
            if (pathRegion)
            {
                const auto open = [this, hwnd, id = descriptor.automationId] { OpenPathDetail(hwnd, id); };
                static_cast<FileOperationsPathRegion*>(region)->SetOnOpenDetail(open);
                region->SetAccessibleInvoke([open](DxUi::WindowHost&) { open(); });
            }
            regions[index] = region;
            ++created;
        }
        regionParents[index] = parent;
    }
    for (size_t old = 0u; old < _hostedTooltipRegions.size(); ++old)
        if (regionKept[old] == 0u && _hostedTooltipRegions[old])
            retired.push_back({_hostedTooltipRegions[old],
                               old < _hostedTooltipParents.size() ? _hostedTooltipParents[old] : _controlsBody,
                               dynamic_cast<FileOperationsPathRegion*>(_hostedTooltipRegions[old]) ? Role::PathRegion : Role::Region});

    // Retire. A control under a retired panel goes with that panel.
    const auto panelRetired = [&](const DxUi::Panel* parent) noexcept
    { return std::ranges::any_of(retired, [parent](const RetiredHostedControl& entry) noexcept { return entry.control == parent; }); };
    std::erase_if(retired, [&](const RetiredHostedControl& entry) noexcept { return panelRetired(entry.parent); });
    // Counted once, here: controls carried over from an earlier capture were counted when they retired.
    const size_t newlyRetired = retired.size();
    // Controls hidden during an earlier capture leave now unless a capture is still active.
    const bool captured = _controlsHost.GetCapturedControl() != nullptr;
    if (! captured)
    {
        for (const RetiredHostedControl& entry : _retiredHostedControls)
            if (! panelRetired(entry.parent))
                retired.push_back(entry);
        _retiredHostedControls.clear();
    }
    std::erase_if(_retiredHostedControls, [&](const RetiredHostedControl& entry) noexcept { return panelRetired(entry.parent); });
    // A focused control under a retired comparison panel is retired with it.
    DxUi::Control* const focus = _controlsHost.GetFocusControl();
    bool focusRetired          = false;
    for (const RetiredHostedControl& entry : retired)
        focusRetired = focusRetired || entry.control == focus;
    for (size_t index = 0u; ! focusRetired && focus && index < _hostedButtons.size() && index < _hostedButtonParents.size(); ++index)
        focusRetired = _hostedButtons[index] == focus && panelRetired(_hostedButtonParents[index]);
    for (size_t index = 0u; ! focusRetired && focus && index < _hostedTooltipRegions.size() && index < _hostedTooltipParents.size(); ++index)
        focusRetired = _hostedTooltipRegions[index] == focus && panelRetired(_hostedTooltipParents[index]);
    const bool focusedRegionRetired = focusRetired && std::ranges::find(_hostedTooltipRegions, focus) != _hostedTooltipRegions.end();
    std::vector<std::unique_ptr<DxUi::Control>> graveyard;
    bool hoverRetired = false;
    if (captured)
    {
        // The host and a ScrollPanel route the release to the pressed control by raw pointer. Keep
        // every retired control alive and hidden until no capture is active. A release never
        // reaches a retired action: only a published button acts, including the buttons of a
        // retired comparison panel. Hiding is the only change, as each one republishes the
        // accessibility snapshot.
        for (const RetiredHostedControl& entry : retired)
        {
            entry.control->SetVisible(false);
            _retiredHostedControls.push_back(entry);
        }
    }
    else if (! retired.empty())
    {
        // A ScrollPanel keeps its captured and hovered child by raw pointer, and the host can drop
        // its capture without telling the panel (ReleaseMouseCapture and the SetFocusControl prune
        // notify nobody). No capture is active here, so a captured child a panel still holds is
        // stale. Every surviving panel that loses a child drops both pointers through its own
        // handlers before any child leaves (OnMouseLeave keeps the hovered child while a captured
        // one is set), whatever the retired child's own hover flag says.
        std::vector<DxUi::ScrollPanel*> releasing;
        for (const RetiredHostedControl& entry : retired)
            if (entry.parent != _controlsRoot && std::ranges::find(releasing, entry.parent) == releasing.end())
                releasing.push_back(static_cast<DxUi::ScrollPanel*>(entry.parent));
        for (DxUi::ScrollPanel* panel : releasing)
            panel->OnCaptureLost(_controlsHost);
        for (const RetiredHostedControl& entry : retired)
            hoverRetired = hoverRetired || entry.control->IsHovered();
        for (DxUi::ScrollPanel* panel : releasing)
            panel->OnMouseLeave(_controlsHost);
        for (const RetiredHostedControl& entry : retired)
        {
            for (auto& slot : entry.parent->GetChildren())
            {
                if (slot.get() == entry.control)
                {
                    graveyard.push_back(std::move(slot));
                    break;
                }
            }
        }
        static_cast<FileOperationsPopupRoot*>(_controlsRoot)->CompactChildren();
        static_cast<FileOperationsHostedScrollPanel*>(_controlsBody)->CompactChildren();
        for (DxUi::ScrollPanel* panel : bodyPanels)
            static_cast<FileOperationsHostedScrollPanel*>(panel)->CompactChildren();
#ifdef ENABLE_TESTS
        if (_debugRetainRetiredControls)
        {
            for (auto& control : graveyard)
            {
                if (auto* region = dynamic_cast<FileOperationsPathRegion*>(control.get()))
                    region->DebugMarkDetached(&_debugDetachedControlInputCount);
                _debugRetiredControlQuarantine.push_back(std::move(control));
            }
        }
#endif
    }
    // The host drops the focus and hover pointers it no longer finds in the tree (or finds hidden)
    // while the retired controls are still alive. Nothing else is reset: a surviving capture, hover
    // or focus stays exactly where it is.
    if (focusRetired)
        _controlsHost.SetFocusControl(nullptr, false);
    else if (hoverRetired && ! _controlsHost.HasActiveTextInput())
        _controlsHost.SetFocusControl(focus, false);
    if (hoverRetired || focusedRegionRetired)
        static_cast<void>(_controlsHost.ClearTooltip());
    _hostedTreeCounters.controlsCreated += created;
    _hostedTreeCounters.controlsRevived += revived;
    _hostedTreeCounters.controlsRetired += newlyRetired;

    _hostedButtons       = std::move(buttons);
    _hostedButtonParents = std::move(buttonParents);
    _hostedButtonHits.resize(_buttons.size());
    for (size_t index = 0u; index < _buttons.size(); ++index)
        _hostedButtonHits[index] = _buttons[index].hit;
    _hostedProgressBars              = std::move(progressBars);
    _hostedProgressAnimations        = std::move(progressAnimations);
    _hostedGraphs                    = std::move(graphs);
    _hostedGraphStallBands           = std::move(stallBands);
    _hostedProgressVerificationMarks = std::move(progressMarks);
    _hostedTooltipRegions            = std::move(regions);
    _hostedTooltipParents            = std::move(regionParents);
    _hostedTooltipComparisonIds.clear();
    _hostedTooltipTaskIds.clear();
    for (const auto& descriptor : _hostedTooltipDescriptors)
    {
        _hostedTooltipComparisonIds.push_back(descriptor.comparisonTaskId);
        _hostedTooltipTaskIds.push_back(descriptor.taskId);
    }
    _hostedConflictBodyPanels = std::move(bodyPanels);
    _hostedConflictBodyIds    = std::move(bodyIds);
    _hostedRightToLeft        = rightToLeft;
    // Publication ticks belong to live controls only, so a new control never inherits a retired one's.
    std::erase_if(_hostedAccessibleTextTicks,
                  [&](const auto& entry) noexcept
    {
        const DxUi::Control* control = entry.first;
        return std::ranges::find(_hostedTooltipRegions, control) == _hostedTooltipRegions.end() &&
               std::ranges::find(_hostedProgressBars, control) == _hostedProgressBars.end() && std::ranges::find(_hostedGraphs, control) == _hostedGraphs.end();
    });

    // One traversal order from final geometry (UI_FileOperationsPopup.md): top-to-bottom, then leading
    // to trailing, with a comparison body's controls kept together at the body's top.
    struct HostedVisualOrder final
    {
        D2D1_RECT_F bounds{};
        const DxUi::Control* control = nullptr;
        float groupTop               = 0.0f;
    };
    std::vector<HostedVisualOrder> visualOrder;
    visualOrder.reserve(_buttons.size() + _hostedProgressDescriptors.size() + 2u * _hostedGraphDescriptors.size() + _hostedTooltipDescriptors.size());
    const auto groupTopFor = [&](uint64_t comparisonTaskId, const D2D1_RECT_F& bounds) noexcept
    {
        const auto* body = FindConflictBody(comparisonTaskId);
        return body && body->visible ? body->bounds.top : bounds.top;
    };
    for (size_t index = 0u; index < _buttons.size(); ++index)
        visualOrder.push_back({_buttons[index].bounds, _hostedButtons[index], groupTopFor(comparisonForHit(_buttons[index].hit), _buttons[index].bounds)});
    for (size_t index = 0u; index < _hostedProgressDescriptors.size(); ++index)
    {
        visualOrder.push_back({_hostedProgressDescriptors[index].bounds, _hostedProgressBars[index], _hostedProgressDescriptors[index].bounds.top});
        // The verification marks paint over their bar: same geometry, so the stable sort keeps them right after.
        if (index < _hostedProgressVerificationMarks.size())
            visualOrder.push_back(
                {_hostedProgressDescriptors[index].bounds, _hostedProgressVerificationMarks[index], _hostedProgressDescriptors[index].bounds.top});
    }
    for (size_t index = 0u; index < _hostedGraphDescriptors.size(); ++index)
    {
        visualOrder.push_back({_hostedGraphDescriptors[index].bounds, _hostedGraphs[index], _hostedGraphDescriptors[index].bounds.top});
        // The stall band paints over its graph: same geometry, so the stable sort keeps it right after.
        if (index < _hostedGraphStallBands.size())
            visualOrder.push_back({_hostedGraphDescriptors[index].bounds, _hostedGraphStallBands[index], _hostedGraphDescriptors[index].bounds.top});
    }
    for (size_t index = 0u; index < _hostedTooltipDescriptors.size(); ++index)
        visualOrder.push_back({_hostedTooltipDescriptors[index].bounds,
                               _hostedTooltipRegions[index],
                               groupTopFor(_hostedTooltipDescriptors[index].comparisonTaskId, _hostedTooltipDescriptors[index].bounds)});
    std::stable_sort(visualOrder.begin(),
                     visualOrder.end(),
                     [rightToLeft](const HostedVisualOrder& left, const HostedVisualOrder& right) noexcept
    {
        if (left.groupTop != right.groupTop)
            return left.groupTop < right.groupTop;
        if (left.bounds.top != right.bounds.top)
            return left.bounds.top < right.bounds.top;
        return rightToLeft ? left.bounds.right > right.bounds.right : left.bounds.left < right.bounds.left;
    });
    std::unordered_map<const DxUi::Control*, size_t> rank;
    rank.reserve(visualOrder.size() + _hostedConflictBodyPanels.size() + 1u);
    rank.emplace(_controlsBody, 0u);
    for (size_t index = 0u; index < visualOrder.size(); ++index)
        rank.emplace(visualOrder[index].control, index + 1u);
    for (size_t index = 0u; index < _hostedTooltipRegions.size() + _hostedButtons.size(); ++index)
    {
        // A comparison panel takes the place of its first control, as its group is contiguous.
        const bool button              = index < _hostedButtons.size();
        DxUi::Panel* const parent      = button ? _hostedButtonParents[index] : _hostedTooltipParents[index - _hostedButtons.size()];
        const DxUi::Control* const own = button ? static_cast<const DxUi::Control*>(_hostedButtons[index])
                                                : static_cast<const DxUi::Control*>(_hostedTooltipRegions[index - _hostedButtons.size()]);
        if (parent == _controlsRoot || parent == _controlsBody)
            continue;
        const auto found = rank.find(own);
        if (found == rank.end())
            continue;
        const size_t ownRank         = found->second;
        const auto [panel, inserted] = rank.try_emplace(parent, ownRank);
        if (! inserted)
            panel->second = std::min(panel->second, ownRank);
    }
    const auto orderChildren = [&](DxUi::Panel& panel) noexcept
    {
        const auto children = panel.GetChildren();
        const auto rankOf   = [&](const std::unique_ptr<DxUi::Control>& child) noexcept
        {
            const auto found = rank.find(child.get());
            return found == rank.end() ? std::numeric_limits<size_t>::max() : found->second;
        };
        std::stable_sort(children.begin(), children.end(), [&](const auto& left, const auto& right) noexcept { return rankOf(left) < rankOf(right); });
    };
    orderChildren(*_controlsRoot);
    orderChildren(*_controlsBody);
    for (DxUi::ScrollPanel* panel : _hostedConflictBodyPanels)
        orderChildren(*panel);
}

void FileOperationsPopupInternal::FileOperationsPopupState::PublishHostedAccessibleText(
    DxUi::Control& control, std::wstring name, std::wstring helpText, bool liveText, ULONGLONG now)
{
    if (control.GetAccessibleName() == name && control.GetAccessibleHelpText() == helpText)
        return;
    const auto [published, first] = _hostedAccessibleTextTicks.try_emplace(&control, now);
    if (liveText && ! first && ! _hostedAccessibleTextFlush && now - published->second < kFileOperationsPopupAccessibleTextIntervalMs)
    {
        // The popup timer keeps painting while text is pending (RefreshTelemetryTimer).
        _hostedAccessibleTextPending = true;
        ++_hostedTreeCounters.accessibleTextDeferrals;
        return;
    }
    published->second = now;
    ++_hostedTreeCounters.accessibleTextUpdates;
    control.SetAccessibleName(std::move(name));
    control.SetAccessibleHelpText(std::move(helpText));
}

uint64_t FileOperationsPopupInternal::FileOperationsPopupState::FocusedHostedTaskId() const noexcept
{
    const DxUi::Control* const focused = _controlsHost.GetFocusControl();
    if (! focused)
        return 0u;
    for (size_t index = 0u; index < _hostedButtons.size() && index < _hostedButtonHits.size(); ++index)
        if (_hostedButtons[index] == focused)
            return _hostedButtonHits[index].taskId;
    for (size_t index = 0u; index < _hostedTooltipRegions.size() && index < _hostedTooltipTaskIds.size(); ++index)
        if (_hostedTooltipRegions[index] == focused)
            return _hostedTooltipTaskIds[index];
    return 0u;
}

const FileOperationsPopupInternal::FileOperationsPopupState::PublishedConflictKeyboardBinding* FileOperationsPopupInternal::FileOperationsPopupState::
    SyncHostedConflictKeyboardBindings(uint64_t fallbackTaskId) noexcept
{
    uint64_t focusedTaskId = FocusedHostedTaskId();
    if (! _controlsHost.GetFocusControl())
        focusedTaskId = fallbackTaskId;

    const PublishedConflictKeyboardBinding* decisionTask = nullptr;
    bool focusedTaskPresent                              = false;
    if (focusedTaskId != 0u)
    {
        const auto focusedTask = std::ranges::find(_publishedConflictKeyboardBindings, focusedTaskId, &PublishedConflictKeyboardBinding::taskId);
        focusedTaskPresent     = focusedTask != _publishedConflictKeyboardBindings.end();
        if (focusedTask != _publishedConflictKeyboardBindings.end() && focusedTask->actionable)
        {
            decisionTask = &*focusedTask;
        }
    }
    if (! decisionTask && ! focusedTaskPresent)
    {
        // Global/no task focus can use the sole actionable decision. A focused task
        // with no published decision must never route Escape to another task, including
        // the safe disclosure fallback while that task's metadata is loading.
        size_t actionableDecisions = 0u;
        for (const auto& task : _publishedConflictKeyboardBindings)
        {
            if (task.actionable)
            {
                ++actionableDecisions;
                decisionTask = &task;
            }
        }
        if (actionableDecisions != 1u)
        {
            decisionTask = nullptr;
        }
    }

    const auto findPublishedActionButton = [&](uint64_t taskId, uint8_t rawAction) noexcept -> DxUi::Button*
    {
        for (size_t index = 0u; index < _hostedButtonHits.size() && index < _hostedButtons.size(); ++index)
        {
            const PopupHitTest& hit = _hostedButtonHits[index];
            if (hit.kind == PopupHitTest::Kind::TaskConflictAction && hit.taskId == taskId && hit.data == static_cast<uint32_t>(rawAction) && decisionTask &&
                hit.promptGeneration == decisionTask->generation)
            {
                return _hostedButtons[index];
            }
        }
        return nullptr;
    };

    DxUi::Button* defaultButton = nullptr;
    DxUi::Button* escapeButton  = nullptr;
    _conflictEscapeTaskId       = 0u;
    _conflictEscapeAction       = 0u;
    _conflictEscapeGeneration   = 0u;
    if (decisionTask)
    {
        defaultButton             = findPublishedActionButton(decisionTask->taskId, decisionTask->defaultAction);
        escapeButton              = findPublishedActionButton(decisionTask->taskId, decisionTask->escapeAction);
        _conflictEscapeTaskId     = decisionTask->taskId;
        _conflictEscapeAction     = decisionTask->escapeAction;
        _conflictEscapeGeneration = decisionTask->generation;
    }

    _controlsHost.SetDefaultButton(defaultButton);
    _controlsHost.SetCancelButton(escapeButton);

    return decisionTask;
}

void FileOperationsPopupInternal::FileOperationsPopupState::SyncHostedControls(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept
{
    // Retain at most two layouts per currently visible expanded graph, never per
    // historical task. Their text/font/width keys handle slot reuse and DPI changes.
    _graphLabelLayouts.resize(_hostedGraphDescriptors.size());
    if (! _controlsRoot || _controlsHost.GetHwnd() != hwnd)
    {
        return;
    }

    const uint64_t startedUs = PerfNowUs();
    const bool syncingBefore = _syncingHostedControls;
    _syncingHostedControls   = true;
    auto restoreSyncing      = wil::scope_exit([&] { _syncingHostedControls = syncingBefore; });
    _publishedConflictKeyboardBindings.clear();
    _publishedConflictKeyboardBindings.reserve(snapshot.size());
    for (const auto& task : snapshot)
        _publishedConflictKeyboardBindings.push_back({task.taskId,
                                                      task.conflict.generation,
                                                      task.conflict.defaultAction,
                                                      task.conflict.escapeAction,
                                                      task.conflict.bucket,
                                                      task.conflict.active && task.conflict.buttonsPublishable});

    ++_hostedTreeCounters.syncs;
    const uint64_t accessibleUpdatesBefore   = _hostedTreeCounters.accessibleTextUpdates;
    const uint64_t accessibleDeferralsBefore = _hostedTreeCounters.accessibleTextDeferrals;
    _hostedAccessibleTextPending             = false;
    _controlsRoot->SetBounds(_controlsHost.GetClientBoundsDip());

    std::optional<PopupHitTest> focusedHit;
    std::wstring focusedPathId;
    uint64_t focusedRegionTaskId = 0u;
    if (auto* focused = _controlsHost.GetFocusControl())
    {
        const auto focusedIt = std::ranges::find(_hostedButtons, focused);
        const auto pathIt    = std::ranges::find(_hostedTooltipRegions, focused);
        if (pathIt != _hostedTooltipRegions.end() && focused->IsFocusable())
        {
            focusedPathId = focused->GetAccessibleAutomationId();
            // Owner recorded with the existing regions; the new descriptors may already be reindexed.
            const size_t pathIndex = static_cast<size_t>(std::distance(_hostedTooltipRegions.begin(), pathIt));
            if (pathIndex < _hostedTooltipTaskIds.size())
                focusedRegionTaskId = _hostedTooltipTaskIds[pathIndex];
        }
        if (focusedIt != _hostedButtons.end())
        {
            const size_t focusedIndex = static_cast<size_t>(std::distance(_hostedButtons.begin(), focusedIt));
            if (focusedIndex < _hostedButtonHits.size())
            {
                focusedHit = _hostedButtonHits[focusedIndex];
            }
        }
    }

    const auto sameHit = [](const PopupHitTest& left, const PopupHitTest& right) noexcept
    { return left.kind == right.kind && left.taskId == right.taskId && left.data == right.data && left.promptGeneration == right.promptGeneration; };
    const bool rightToLeft = (GetWindowLongPtrW(hwnd, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) != 0;
    bool initialBuild      = ! _controlsBody;
    // popup-b#1: any difference between the published descriptors and the retained controls is
    // reconciled by identity. Only the first build after the host attaches creates the whole tree.
    bool structural  = initialBuild || _buttons.size() != _hostedButtonHits.size() || _hostedProgressDescriptors.size() != _hostedProgressBars.size() ||
                       _hostedGraphDescriptors.size() != _hostedGraphs.size() || _hostedGraphDescriptors.size() != _hostedGraphStallBands.size() ||
                       _hostedProgressDescriptors.size() != _hostedProgressVerificationMarks.size() ||
                       _hostedTooltipDescriptors.size() != _hostedTooltipRegions.size() || rightToLeft != _hostedRightToLeft ||
                       (! _retiredHostedControls.empty() && _controlsHost.GetCapturedControl() == nullptr);
    size_t bodyIndex = 0u;
    for (const auto& body : _conflictBodies)
    {
        if (! body.visible)
            continue;
        // A new prompt can reuse a task ID; its old panel must not survive context replacement.
        if (! body.panel || bodyIndex >= _hostedConflictBodyIds.size() || _hostedConflictBodyIds[bodyIndex] != body.taskId)
            structural = true;
        ++bodyIndex;
    }
    structural = structural || bodyIndex != _hostedConflictBodyIds.size();
    if (! structural)
    {
        for (size_t index = 0u; index < _hostedTooltipDescriptors.size(); ++index)
        {
            if (index >= _hostedTooltipComparisonIds.size() || _hostedTooltipComparisonIds[index] != _hostedTooltipDescriptors[index].comparisonTaskId ||
                _hostedTooltipRegions[index]->IsFocusable() != ! _hostedTooltipDescriptors[index].copyPath.empty() ||
                _hostedTooltipRegions[index]->GetAccessibleAutomationId() != _hostedTooltipDescriptors[index].automationId)
                structural = true;
        }
        for (size_t index = 0u; index < _buttons.size(); ++index)
        {
            if (! sameHit(_buttons[index].hit, _hostedButtonHits[index]))
            {
                structural = true;
                break;
            }
        }
    }
#ifdef ENABLE_TESTS
    initialBuild = initialBuild || (debugRebuildOnStructuralChange && structural);
#endif

    std::wstring detailTextBeforeRebuild;
    std::wstring detailFocusBeforeRebuild;
    std::optional<std::pair<size_t, size_t>> detailSelectionBeforeRebuild;
    size_t detailCaretBeforeRebuild = 0u;
    if (initialBuild && _pathDetailField && ! _pathDetailId.empty())
    {
        detailTextBeforeRebuild      = _pathDetailField->GetText();
        detailSelectionBeforeRebuild = _pathDetailField->GetSelectionRange();
        detailCaretBeforeRebuild     = _pathDetailField->GetCaretIndex();
        if (const auto* focus = _controlsHost.GetFocusControl())
        {
            detailFocusBeforeRebuild = focus->GetAccessibleAutomationId();
            if (detailFocusBeforeRebuild.starts_with(L"FileOperations.PathDetail."))
                _pathDetailLogicalFocusId = detailFocusBeforeRebuild;
        }
    }
    if (initialBuild)
    {
        // The first build after the host attaches: no hosted control exists yet. The host owns
        // non-owning focus/hover/capture pointers into the retained tree; clear them before the
        // root's children go.
        ++_hostedTreeCounters.rebuilds;
        _controlsHost.ResetInteractionState();
        _controlsRoot->ClearChildren();
        _pathDetailPanel = nullptr;
        _pathDetailTitle = nullptr;
        _pathDetailField = nullptr;
        _pathDetailCopy  = nullptr;
        _pathDetailClose = nullptr;
        _hostedConflictBodyIds.clear();
        _hostedConflictBodyPanels.clear();
        for (auto& body : _conflictBodies)
            body.panel = nullptr;
        // Native scrolling owns the offset; ScrollPanel supplies the same body
        // clip to retained paint, pointer dispatch and accessibility hit testing.
        _controlsBody = _controlsRoot->AddChild<FileOperationsHostedScrollPanel>();
        _controlsBody->SetInternalScrollbarEnabled(false);
        _hostedButtons.clear();
        _hostedButtonHits.clear();
        _hostedButtonParents.clear();
        _hostedProgressBars.clear();
        _hostedProgressAnimations.clear();
        _hostedGraphs.clear();
        _hostedGraphStallBands.clear();
        _hostedProgressVerificationMarks.clear();
        _hostedTooltipRegions.clear();
        _hostedTooltipParents.clear();
        _hostedTooltipComparisonIds.clear();
        _hostedTooltipTaskIds.clear();
        _retiredHostedControls.clear();
        _hostedAccessibleTextTicks.clear();
    }
    if (structural)
        ReconcileHostedControls(hwnd, rightToLeft);
    const auto comparisonForHit = [](const PopupHitTest& hit) noexcept -> uint64_t
    { return hit.kind == PopupHitTest::Kind::TaskConflictInspectSource || hit.kind == PopupHitTest::Kind::TaskConflictInspectDestination ? hit.taskId : 0u; };

    _controlsBody->SetBounds(D2D1::RectF(_controlsHost.PixelsToDip(_listViewportRect.left),
                                         _controlsHost.PixelsToDip(_listViewportRect.top),
                                         _controlsHost.PixelsToDip(_listViewportRect.right),
                                         _controlsHost.PixelsToDip(_listViewportRect.bottom)));
    _controlsBody->SetContentHeight(_controlsHost.PixelsToDip(std::max(0.0f, _listViewportRect.bottom - _listViewportRect.top)));

    _syncingConflictBodies = true;
    for (auto& body : _conflictBodies)
    {
        if (! body.visible || ! body.panel)
            continue;
        const float requestedOffset = body.offsetDip;
        body.panel->SetBounds(D2D1::RectF(_controlsHost.PixelsToDip(body.bounds.left),
                                          _controlsHost.PixelsToDip(body.bounds.top),
                                          _controlsHost.PixelsToDip(body.bounds.right),
                                          _controlsHost.PixelsToDip(body.bounds.bottom)));
        body.panel->SetContentHeight(_controlsHost.PixelsToDip(body.contentHeightPx));
        body.panel->SetScrollOffset(requestedOffset);
        body.offsetDip = body.panel->GetScrollOffset();
    }
    _syncingConflictBodies = false;

    for (size_t index = 0u; index < _buttons.size() && index < _hostedButtons.size(); ++index)
    {
        D2D1_RECT_F bounds = _buttons[index].bounds;
        if (const auto* body = FindConflictBody(comparisonForHit(_buttons[index].hit)); body && body->visible)
        {
            const float offset = DipsToPixels(body->offsetDip, _dpi);
            bounds.top += offset;
            bounds.bottom += offset;
        }
        DxUi::Button* const button        = _hostedButtons[index];
        const PopupHitTest& hit           = _buttons[index].hit;
        const std::wstring controlText    = HostedControlText(hit);
        const std::wstring accessibleName = HostedControlAccessibleName(hit);
        button->SetBounds(D2D1::RectF(_controlsHost.PixelsToDip(bounds.left),
                                      _controlsHost.PixelsToDip(bounds.top),
                                      _controlsHost.PixelsToDip(bounds.right),
                                      _controlsHost.PixelsToDip(bounds.bottom)));
        button->SetText(controlText);
        button->SetAccessibleName(accessibleName);
        button->SetTooltipText(accessibleName != controlText ? accessibleName : std::wstring{});
        button->SetVariant(HostedButtonVariant(hit.kind));
        button->SetMultiline(hit.kind == PopupHitTest::Kind::TaskConflictAction || hit.kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll ||
                             hit.kind == PopupHitTest::Kind::TaskIssues || hit.kind == PopupHitTest::Kind::TaskDismiss ||
                             hit.kind == PopupHitTest::Kind::TaskCompletedMore);
        if (hit.kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll)
        {
            const auto task      = std::ranges::find(snapshot, hit.taskId, &TaskSnapshot::taskId);
            auto* const checkbox = static_cast<DxUi::Checkbox*>(button);
            checkbox->SetChecked(task != snapshot.end() && task->conflict.applyToAllChecked);
            checkbox->SetEnabled(task != snapshot.end() && task->conflict.buttonsPublishable && task->conflict.applyToAllEligible);
        }

        if (hit.kind == PopupHitTest::Kind::TaskConflictInspectSource || hit.kind == PopupHitTest::Kind::TaskConflictInspectDestination)
            button->SetEnabled(! _inspectionPresentation || ! _inspectionPresentation->pending || sameHit(hit, _inspectionPresentation->invokingHit));

        if (hit.kind == PopupHitTest::Kind::FooterToggleDetails)
        {
            button->SetDisclosureCollapsedDirection(DxUi::ChevronDirection::Left);
            button->SetDisclosureExpanded(fileOps ? ! fileOps->GetPopupFooterOnly() : true);
        }
        else if (hit.kind == PopupHitTest::Kind::CompletedGroupToggle)
        {
            button->SetDisclosureCollapsedDirection(DxUi::ChevronDirection::Left);
            button->SetDisclosureExpanded(_completedGroupExpanded);
        }
        else if (hit.kind == PopupHitTest::Kind::TaskToggleCollapse)
        {
            button->SetDisclosureCollapsedDirection(DxUi::ChevronDirection::Left);
            button->SetDisclosureExpanded(! IsTaskCollapsedForDisplay(hit.taskId, fileOps ? fileOps->GetPopupCompactDensity() : false));
        }
        else if (hit.kind == PopupHitTest::Kind::TaskToggleThroughput)
        {
            button->SetDisclosureCollapsedDirection(DxUi::ChevronDirection::Right);
            button->SetDisclosureExpanded(IsTaskThroughputExpanded(hit.taskId));
        }
        else
        {
            button->ClearDisclosureState();
        }
    }

    const auto boundsToDip = [&](const D2D1_RECT_F& bounds) noexcept
    {
        return D2D1::RectF(_controlsHost.PixelsToDip(bounds.left),
                           _controlsHost.PixelsToDip(bounds.top),
                           _controlsHost.PixelsToDip(bounds.right),
                           _controlsHost.PixelsToDip(bounds.bottom));
    };
    const auto taskStatusName = [&](uint64_t taskId)
    {
        if (taskId == 0u)
        {
            return LoadStringResource(nullptr, IDS_FILEOPS_AGGREGATE_PROGRESS);
        }
        const auto taskIt = std::ranges::find(snapshot, taskId, &TaskSnapshot::taskId);
        return taskIt != snapshot.end() ? StatusTextForTask(*taskIt, taskIt->statusKind, GetTickCount64())
                                        : LoadStringResource(nullptr, IDS_FILEOPS_STATUS_RUNNING);
    };
    const ULONGLONG animationTick = GetTickCount64();
    const bool reducedMotion      = IsReducedMotionEnabled();
    // Motion follows the snapshot's existing task-status policy, independently of
    // whether a truthful total is available. An unanswered, loaded prompt is idle.
    const auto taskMayAnimate = [](const TaskSnapshot& task) noexcept
    {
        const auto status = ResolveTaskStatusKind(task);
        return ! task.finished &&
               ((StatusCountsAsRunning(status) && status != TaskStatusKind::Paused) || (status == TaskStatusKind::Conflict && task.conflict.metadataLoading));
    };
    const bool aggregateMayAnimate = std::ranges::any_of(snapshot, taskMayAnimate);
    for (size_t index = 0u; index < _hostedProgressDescriptors.size() && index < _hostedProgressBars.size(); ++index)
    {
        const HostedProgressDescriptor& descriptor = _hostedProgressDescriptors[index];
        auto* progress                             = _hostedProgressBars[index];
        progress->SetBounds(boundsToDip(descriptor.bounds));
        progress->SetTrackHeightDip(_controlsHost.PixelsToDip(std::max(0.0f, descriptor.bounds.bottom - descriptor.bounds.top)));
        progress->SetMinimum(0.0);
        progress->SetMaximum(100.0);
        progress->SetValue(ResolveHostedProgressValue(index, descriptor.value, animationTick, reducedMotion));
        const auto task = std::ranges::find(snapshot, descriptor.taskId, &TaskSnapshot::taskId);
        progress->SetAnimationPaused(descriptor.taskId == 0u ? ! aggregateMayAnimate : task == snapshot.end() || ! taskMayAnimate(*task));
        progress->SetIndeterminate(descriptor.indeterminate);
        std::wstring helpText;
        if (descriptor.segmented && ! descriptor.indeterminate && folderWindow)
        {
            progress->SetSegmentedValues(
                descriptor.primarySegmentValue, descriptor.secondarySegmentValue, folderWindow->GetTheme().fileOperations.progressVerify);
            helpText = FormatStringResource(nullptr,
                                            IDS_FMT_FILEOPS_PROGRESS_SEGMENTS_ACCESSIBLE,
                                            static_cast<unsigned int>(std::lround(descriptor.primarySegmentValue)),
                                            static_cast<unsigned int>(std::lround(descriptor.secondarySegmentValue)));
        }
        else
        {
            progress->ClearSegmentedValues();
        }
        if (index < _hostedProgressVerificationMarks.size() && _hostedProgressVerificationMarks[index])
        {
            // Created only as FileOperationsVerificationHalfMarks (ReconcileHostedControls).
            auto* const marks = static_cast<FileOperationsVerificationHalfMarks*>(_hostedProgressVerificationMarks[index]);
            marks->SetBounds(boundsToDip(descriptor.bounds));
            marks->SetMarks(descriptor.segmented && ! descriptor.indeterminate && folderWindow,
                            descriptor.secondarySegmentValue,
                            folderWindow ? D2D1_COLOR_F(folderWindow->GetTheme().fileOperations.progressVerify) : D2D1_COLOR_F{},
                            descriptor.verificationSettled);
        }
        // Only the segment percentages are volatile. A status change (Paused, Stopping, a decision)
        // publishes at once; the Name itself changes at most once a second (elapsed time).
        std::wstring name   = taskStatusName(descriptor.taskId);
        const bool liveText = progress->GetAccessibleName() == name;
        PublishHostedAccessibleText(*progress, std::move(name), std::move(helpText), liveText, animationTick);
    }
    for (size_t index = 0u; index < _hostedGraphDescriptors.size() && index < _hostedGraphs.size(); ++index)
    {
        const HostedGraphDescriptor& descriptor = _hostedGraphDescriptors[index];
        auto* graph                             = _hostedGraphs[index];
        graph->SetBounds(boundsToDip(descriptor.bounds));
        graph->SetSamples(descriptor.samples);
        graph->SetSecondarySamples(descriptor.verificationSamples);
        if (folderWindow)
        {
            graph->SetSecondarySeriesColor(folderWindow->GetTheme().fileOperations.graphVerify);
        }
        graph->SetLimit(0.0);
        graph->SetCurrentValueMarker(descriptor.currentBandwidthBytesPerSecond,
                                     descriptor.separateLabelRow ? std::wstring{} : descriptor.currentBandwidthText,
                                     descriptor.separateLabelRow ? std::wstring{} : descriptor.currentBandwidthTrailingText);
        graph->SetOverlayText(descriptor.overlayText);
        graph->SetRainbowMode(descriptor.rainbowMode);
        graph->SetPerStreamBands(descriptor.perStreamBands);
        if (index < _hostedGraphStallBands.size() && _hostedGraphStallBands[index])
        {
            // Created only as FileOperationsGraphStallBand (ReconcileHostedControls).
            auto* const band = static_cast<FileOperationsGraphStallBand*>(_hostedGraphStallBands[index]);
            band->SetBounds(boundsToDip(descriptor.stallBandBounds));
            band->SetStalledSamples(descriptor.stalledSamples);
        }
        PublishHostedAccessibleText(*graph,
                                    LoadStringResource(nullptr, IDS_FILEOPS_THROUGHPUT_GRAPH),
                                    descriptor.currentBandwidthTrailingText.empty()
                                        ? descriptor.currentBandwidthText
                                        : std::format(L"{0} {1}", descriptor.currentBandwidthText, descriptor.currentBandwidthTrailingText),
                                    true,
                                    animationTick);
    }
    for (size_t index = 0u; index < _hostedTooltipDescriptors.size() && index < _hostedTooltipRegions.size(); ++index)
    {
        const HostedTooltipDescriptor& descriptor = _hostedTooltipDescriptors[index];
        auto* region                              = _hostedTooltipRegions[index];
        auto bounds                               = boundsToDip(descriptor.bounds);
        if (const auto* body = FindConflictBody(descriptor.comparisonTaskId); body && body->visible)
        {
            bounds.top += body->offsetDip;
            bounds.bottom += body->offsetDip;
        }
        region->SetBounds(bounds);
        region->SetAccessibleAutomationId(descriptor.automationId);
        const std::wstring& helpText = descriptor.helpText.empty() ? descriptor.text : descriptor.helpText;
        PublishHostedAccessibleText(*region, descriptor.text, helpText, descriptor.liveText, animationTick);
        region->SetTooltipText(helpText);
        if (! descriptor.copyPath.empty())
            static_cast<FileOperationsPathRegion*>(region)->SetPath(_controlsHost, descriptor.copyPath);
    }

    const auto restoreRebuiltFocus = [&](const PopupHitTest& hit) noexcept
    {
        for (size_t index = 0u; index < _hostedButtonHits.size(); ++index)
        {
            if (sameHit(_hostedButtonHits[index], hit) && _hostedButtons[index]->IsFocusable())
            {
                // Rebuilding a background popup must not reclaim the user's native focus.
                _controlsHost.SetFocusControl(_hostedButtons[index], false);
                return true;
            }
        }
        return false;
    };
    const auto restoreSafeRebuiltFocus = [&](uint64_t taskId) noexcept
    {
        // Disclosure is non-destructive and retains the task identity. If that task
        // disappeared, use global Options; never choose another task's decision.
        if (taskId == 0u || ! restoreRebuiltFocus({PopupHitTest::Kind::TaskToggleCollapse, taskId, 0u}))
        {
            static_cast<void>(restoreRebuiltFocus({PopupHitTest::Kind::FooterOptions, 0u, 0u}));
        }
    };
    // A focused control that survived keeps its focus untouched. Only a retired one is replaced, by
    // identity, and the replacement never scrolls the list (popup-b#3).
    const bool focusLost = (focusedHit.has_value() || ! focusedPathId.empty()) && _controlsHost.GetFocusControl() == nullptr;
    if (focusLost)
    {
        const bool restoringBefore = _restoringHostedFocus;
        _restoringHostedFocus      = true;
        if (focusedHit.has_value() && ! restoreRebuiltFocus(focusedHit.value()))
            restoreSafeRebuiltFocus(focusedHit->taskId);
        if (! focusedPathId.empty())
        {
            const auto path = std::ranges::find_if(_hostedTooltipRegions, [&](const DxUi::Control* region) noexcept {
                return region->IsFocusable() && region->GetAccessibleAutomationId() == focusedPathId;
            });
            if (path != _hostedTooltipRegions.end())
                _controlsHost.SetFocusControl(*path, false);
            else
                restoreSafeRebuiltFocus(focusedRegionTaskId);
        }
        _restoringHostedFocus = restoringBefore;
    }

    bool inspectionFocusRestored = false;
    if (_inspectionReturn && _inspectionReturn->restoreFocus)
    {
        if (fileOps && fileOps->IsPendingConflictInspectionCurrent(
                           _inspectionReturn->invokingHit.taskId, _inspectionReturn->invokingHit.promptGeneration, _inspectionReturn->context))
        {
            for (size_t index = 0u; index < _hostedButtonHits.size() && index < _hostedButtons.size(); ++index)
            {
                if (sameHit(_hostedButtonHits[index], _inspectionReturn->invokingHit))
                {
                    // Preserve the return control here without defeating ShowPopup's
                    // foreground/input witness while its no-activate show is in progress.
                    _controlsHost.SetFocusControl(_hostedButtons[index], false);
                    inspectionFocusRestored = _controlsHost.GetFocusControl() == _hostedButtons[index];
                    break;
                }
            }
        }
        _inspectionReturn.reset();
    }

    // Rebind after reconciliation and inspection restoration, using identities in
    // the current tree rather than a captured control slot.
    const auto* decisionTask  = SyncHostedConflictKeyboardBindings(focusedHit.has_value() ? focusedHit->taskId : focusedRegionTaskId);
    auto* const defaultButton = _controlsHost.GetDefaultButton();

    const bool recycleEscalation = decisionTask &&
                                   (static_cast<ConflictBucket>(decisionTask->bucket) == ConflictBucket::RecycleFailed ||
                                    static_cast<ConflictBucket>(decisionTask->bucket) == ConflictBucket::RecycleFailedByName) &&
                                   decisionTask->defaultAction == RawConflictAction(ConflictAction::Cancel);
    if (recycleEscalation && defaultButton != nullptr)
    {
        if (! inspectionFocusRestored && (_recycleEscalationFocusedTaskId != decisionTask->taskId ||
                                          _recycleEscalationFocusedGeneration != decisionTask->generation || (initialBuild && ! focusedHit.has_value())))
        {
            // Presentation selects the safe retained default; only the explicit
            // View command owns activation of the native popup window.
            _controlsHost.SetFocusControl(defaultButton, false);
        }
        _recycleEscalationFocusedTaskId     = decisionTask->taskId;
        _recycleEscalationFocusedGeneration = decisionTask->generation;
    }
    else
    {
        _recycleEscalationFocusedTaskId     = 0u;
        _recycleEscalationFocusedGeneration = 0u;
    }

    for (const TaskSnapshot& task : snapshot)
    {
        const auto previous                = _announcedTaskStatuses.find(task.taskId);
        const auto previousConflictLoading = _announcedConflictMetadataLoading.find(task.taskId);
        const bool isAnnounceableStatus    = task.statusKind == TaskSnapshot::StatusKind::Done || task.statusKind == TaskSnapshot::StatusKind::Partial ||
                                             task.statusKind == TaskSnapshot::StatusKind::Failed || task.statusKind == TaskSnapshot::StatusKind::Conflict ||
                                             task.statusKind == TaskSnapshot::StatusKind::Stopping || task.statusKind == TaskSnapshot::StatusKind::Canceled ||
                                             task.statusKind == TaskSnapshot::StatusKind::Indeterminate;
        const bool statusChanged           = previous != _announcedTaskStatuses.end() && previous->second != task.statusKind;
        const bool initialAttention        = previous == _announcedTaskStatuses.end() && task.statusKind == TaskSnapshot::StatusKind::Conflict;
        // x-ux#1: a task that ended Failed, Partial or Outcome unknown before its card first appeared
        // is announced once when that card appears, also on a new popup's first synchronization (it
        // was created for that card). A restart notice of an interrupted Move is not a new result.
        const bool initialProblemResult = previous == _announcedTaskStatuses.end() && task.finished && ! task.interruptedMoveNotice &&
                                          (task.statusKind == TaskSnapshot::StatusKind::Failed || task.statusKind == TaskSnapshot::StatusKind::Partial ||
                                           task.statusKind == TaskSnapshot::StatusKind::Indeterminate);
        const bool conflictDecisionStateChanged = task.statusKind == TaskSnapshot::StatusKind::Conflict &&
                                                  previousConflictLoading != _announcedConflictMetadataLoading.end() &&
                                                  previousConflictLoading->second != task.conflict.metadataLoading;
        // I35 R8.12: a published decision is announced once: by the decision mirror when it took the focus
        // over a pane, otherwise here, once the mirror's reconcile decided (until then the transition is kept
        // for the next synchronization). The text is the one both surfaces share.
        bool deferAnnouncement = false;
        if (isAnnounceableStatus && (statusChanged || initialAttention || initialProblemResult || conflictDecisionStateChanged) &&
            (_hostedAccessibilityInitialized || initialAttention || initialProblemResult))
        {
            std::wstring announcement = StatusTextForTask(task, task.statusKind, GetTickCount64());
            bool announce             = true;
            if (task.statusKind == TaskSnapshot::StatusKind::Conflict && ! task.conflict.metadataLoading)
            {
                using AnnouncementClaim = FolderWindow::FileOperationState::PromptAnnouncementClaim;
                announcement            = FileOperations::BuildConflictPromptAnnouncement(ConflictPromptTextFactsFor(task));
                const AnnouncementClaim claim =
                    fileOps ? fileOps->ClaimPromptAnnouncementForCard(task.taskId, task.conflict.generation) : AnnouncementClaim::Announce;
                announce          = claim == AnnouncementClaim::Announce;
                deferAnnouncement = claim == AnnouncementClaim::Defer;
            }
            if (announce && DxUi::RaiseWindowHostAccessibilityNotification(hwnd, announcement, std::format(L"FileOperations.Task.{}", task.taskId)))
            {
                ++_hostedAccessibilityNotificationCount;
            }
        }
        if (deferAnnouncement)
        {
            continue;
        }
        _announcedTaskStatuses[task.taskId] = task.statusKind;
        if (task.statusKind == TaskSnapshot::StatusKind::Conflict)
        {
            _announcedConflictMetadataLoading[task.taskId] = task.conflict.metadataLoading;
        }
        else
        {
            _announcedConflictMetadataLoading.erase(task.taskId);
        }
    }
    _hostedAccessibilityInitialized = true;
    SyncPathDetail(hwnd);
    if (initialBuild && _pathDetailField && _pathDetailField->GetText() == detailTextBeforeRebuild)
    {
        if (detailSelectionBeforeRebuild)
        {
            const auto [start, end] = detailSelectionBeforeRebuild.value();
            _pathDetailField->SetSelectionRange(detailCaretBeforeRebuild == start ? end : start, detailCaretBeforeRebuild);
        }
        else
            _pathDetailField->SetSelectionRange(detailCaretBeforeRebuild, detailCaretBeforeRebuild);
        // Logical focus is restored by SyncPathDetail under its native-focus policy.
    }
    _controlsHost.RefreshAccessibilitySnapshot();

    const uint64_t syncUs = PerfElapsedUs(startedUs);
    if (initialBuild)
        _hostedTreeCounters.rebuildUs += syncUs;
    else if (structural)
        _hostedTreeCounters.reconcileUs += syncUs;
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Popup.HostedControls.SyncUs",
                          initialBuild ? L"rebuild"
                          : structural ? L"reconcile"
                                       : L"update",
                          syncUs,
                          _buttons.size(),
                          snapshot.size(),
                          S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.HostedControls.AccessibleTextUpdates",
                          L"applied, deferred",
                          0u,
                          _hostedTreeCounters.accessibleTextUpdates - accessibleUpdatesBefore,
                          _hostedTreeCounters.accessibleTextDeferrals - accessibleDeferralsBefore,
                          S_OK);
    }
    _hostedAccessibleTextFlush = false;
    // Pending live text keeps the popup timer running until a frame publishes it.
    if (_hostedAccessibleTextPending)
        RefreshTelemetryTimer(hwnd);
}

void FileOperationsPopupInternal::FileOperationsPopupState::Render(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const auto paintFailureSurfaceGround = [hwnd]() noexcept
    {
        PAINTSTRUCT ps{};
        wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
        FillRect(hdc.get(), &ps.rcPaint, GetSysColorBrush(COLOR_WINDOW));
    };

    if (_failureSurface)
    {
        if (! _failureSurfaceReady && ! EnterFailureSurface(hwnd))
        {
            paintFailureSurfaceGround();
            return;
        }

        // C4: the native children paint themselves; the popup only clears its ground.
        paintFailureSurfaceGround();
        return;
    }

    // A suspended host has no drawable client target. Leave its retained resources
    // intact; a later nonzero WM_SIZE resumes preparation instead of entering failure UI.
    if (_clientSize.cx <= 0 || _clientSize.cy <= 0 || IsIconic(hwnd))
        return;

    if (! hostLifetime.lock())
    {
        return;
    }

    const FolderWindow* folderWindowPtr = folderWindow;
    if (! folderWindowPtr)
    {
        return;
    }
    const AppTheme& appTheme = folderWindowPtr->GetTheme();

    // Q6: a frame that could not be prepared is retried from the popup timer once its backoff
    // elapses. Until then the frame is skipped; the host's WM_PAINT consumes the update region.
    if (_framePrepareRetryTick != 0u && GetTickCount64() < _framePrepareRetryTick)
        return;

    EnsureTarget(hwnd);
    EnsureTextFormats();
    EnsureBrushes();

    if (! _target || ! _bgBrush || ! _textBrush || ! _borderBrush)
    {
        // C4: without Direct2D there is no hand-painted popup. A transient failure skips the frame;
        // only a detached host or repeated failures select the native failure surface.
        if (OnFramePrepareFailed(hwnd, _target ? E_OUTOFMEMORY : D2DERR_RECREATE_TARGET))
            paintFailureSurfaceGround();
        return;
    }
#ifdef ENABLE_TESTS
    _debugLastDrawnCheckboxGlyph = L'\0';
#endif

    const bool capturePerf                   = Debug::Perf::IsCaptureEnabled();
    const bool reducedMotion                 = IsReducedMotionEnabled();
    const uint64_t renderStartedUs           = capturePerf ? PerfNowUs() : 0u;
    const uint64_t snapshotStartedUs         = capturePerf ? PerfNowUs() : 0u;
    const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
    const uint64_t snapshotUs                = capturePerf ? PerfElapsedUs(snapshotStartedUs) : 0u;
    CleanupCollapsedTasks(snapshot);
    AutoCollapseCompletedTasks(snapshot);
    UpdateCaptionStatus(hwnd, snapshot);
    const bool footerOnly                                   = fileOps ? fileOps->GetPopupFooterOnly() : false;
    const bool compactDensity                               = fileOps ? fileOps->GetPopupCompactDensity() : false;
    GlobalFileOperationsStatusSummary globalSummary         = BuildGlobalStatusSummary(snapshot, &_rates);
    const std::wstring globalSummaryText                    = FormatGlobalStatusSummaryText(globalSummary);
    const bool showPauseResumeAll                           = HasFooterPauseResumeAllControl(globalSummary);
    const bool showCancelAll                                = HasCancelableTasks();
    const CompletedGroupResultSummary completedGroupSummary = BuildCompletedGroupResultSummary(snapshot);
    const uint32_t completedGroupCount                      = completedGroupSummary.total;
    const bool showCompletedGroup                           = ShouldShowCompletedGroup(completedGroupCount);
    ResolveBulkCancelGroupCollapse(hwnd, snapshot, showCompletedGroup);
    constexpr ULONGLONG kCompletedInFlightGraceMs = 300ull;
    const ULONGLONG renderTick                    = GetTickCount64();
    const float completedGroupVisibility          = ResolveCompletedGroupVisibility(renderTick, reducedMotion);

    std::vector<PopupDisplayRow> displayRows;
    if (! footerOnly)
    {
        displayRows.reserve(snapshot.size() + (showCompletedGroup ? 1u : 0u));
        bool completedGroupInserted = false;
        for (size_t i = 0; i < snapshot.size(); ++i)
        {
            const TaskSnapshot& task = snapshot[i];
            if (showCompletedGroup && IsCompletedGroupTask(task))
            {
                if (! completedGroupInserted)
                {
                    displayRows.push_back(PopupDisplayRow{.kind = PopupDisplayRowKind::CompletedGroup});
                    completedGroupInserted = true;
                }
                if (! _completedGroupExpanded && completedGroupVisibility <= 0.001f)
                {
                    continue;
                }
            }

            displayRows.push_back(PopupDisplayRow{.kind = PopupDisplayRowKind::Task, .taskIndex = i});
        }
    }

    float width  = static_cast<float>(_clientSize.cx);
    float height = static_cast<float>(_clientSize.cy);
    // Measure labels once per render, retaining the fixed vocabulary cache. The
    // post-resize layout pass only rearranges the five measured controls.
    if (showPauseResumeAll || showCancelAll)
        MeasureFooterActionWidths();
    LayoutChrome(width, height, showPauseResumeAll, showCancelAll, globalSummaryText);

    const float padding = DipsToPixels(10.0f, _dpi);
    const float cardGap = DipsToPixels(10.0f, _dpi);

    const float expandedCardH         = DipsToPixels(kFileOperationsExpandedCardHeightDip, _dpi);
    const float expandedTransferCardH = DipsToPixels(kFileOperationsExpandedTransferCardHeightDip, _dpi);
    const float collapsedCardH        = DipsToPixels(44.0f, _dpi);
    const float baseLineH             = DipsToPixels(18.0f, _dpi);
    const float fromToGapY            = DipsToPixels(4.0f, _dpi);
    const float promptRowGap          = DipsToPixels(4.0f, _dpi);
    const float promptMeasureWidth    = std::max(1.0f, static_cast<float>(_clientSize.cx) - padding * 4.0f);

    const uint64_t cardLayoutStartedUs = capturePerf ? PerfNowUs() : 0u;
    // Reserve the scrollbar gutter even before it appears. This keeps action wrapping
    // stable through the width/height dependency of native scrollbar creation.
    const float scrollGutterPx = static_cast<float>(GetSystemMetricsForDpi(SM_CXVSCROLL, _dpi));
    const float conflictActionWidthPx =
        std::max(1.0f, static_cast<float>(_clientSize.cx) + (_scrollBarVisible ? scrollGutterPx : 0.0f) - scrollGutterPx - padding * 4.0f);
    const bool actionRightToLeft = (GetWindowLongPtrW(hwnd, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) != 0;
    std::erase_if(_conflictBodies,
                  [&](const ConflictBodyLayout& body)
    {
        const auto task = std::ranges::find(snapshot, body.taskId, &TaskSnapshot::taskId);
        return task == snapshot.end() || ! task->conflict.active || task->conflict.inspection != body.context;
    });
    for (auto& body : _conflictBodies)
        body.visible = false;
    _conflictActionMeasurements.resize(snapshot.size());
    _speedLimitWidthMeasurements.resize(snapshot.size());
    for (size_t index = 0u; index < snapshot.size(); ++index)
    {
        if (footerOnly || IsTaskCollapsedForDisplay(snapshot[index].taskId, compactDensity) || ! snapshot[index].conflict.active ||
            ! snapshot[index].conflict.buttonsPublishable)
        {
            _conflictActionMeasurements[index] = {};
            continue;
        }
        const HRESULT layoutHr = MeasureConflictActions(snapshot[index], conflictActionWidthPx, actionRightToLeft, _conflictActionMeasurements[index]);
        if (layoutHr == E_PENDING)
            return;
        // x-arch#2: a measurement or fit failure keeps every action on a fallback layout.
        if (FAILED(layoutHr))
            ReportDegradedLayout(0u, layoutHr);
    }

    if (! footerOnly && showCompletedGroup)
    {
        const std::array counts{completedGroupSummary.completed,
                                completedGroupSummary.partial,
                                completedGroupSummary.indeterminate,
                                completedGroupSummary.failed,
                                completedGroupSummary.canceled};
        const HRESULT groupHr = PrepareCompletedGroup(conflictActionWidthPx, completedGroupCount, counts);
        if (groupHr == E_PENDING)
            return;
        if (FAILED(groupHr))
            ReportDegradedLayout(1u, groupHr);
    }
    else
        _completedGroupMeasurement = {};
    _terminalSummaries.resize(snapshot.size());
    _scanStatusLayouts.resize(snapshot.size());
    for (auto& entry : _scanStatusLayouts)
        entry.usedThisPaint = false;
    const auto releaseUnusedScanLayouts = wil::scope_exit([&]
    {
        for (auto& entry : _scanStatusLayouts)
            if (! entry.usedThisPaint)
                entry = {};
    });
    _terminalFacts.resize(snapshot.size());
    _terminalActions.resize(snapshot.size());
    _cardInformationLayouts.resize(snapshot.size());
    const uint64_t languageRevision = Localization::GetLanguagePreferenceRevision();
    for (size_t index = 0u; index < snapshot.size(); ++index)
    {
        const auto& task       = snapshot[index];
        const bool groupHidden = showCompletedGroup && ! _completedGroupExpanded && IsCompletedGroupTask(task);
        if (footerOnly || groupHidden || task.kind != TaskSnapshot::Kind::Informational || IsTaskCollapsedForDisplay(task.taskId, compactDensity))
            _scanStatusLayouts[index] = {};
        if (footerOnly || groupHidden || task.kind != TaskSnapshot::Kind::FileOperation || ! task.finished ||
            IsTaskCollapsedForDisplay(task.taskId, compactDensity))
        {
            _terminalSummaries[index]               = {};
            _terminalFacts[index]                   = {};
            _terminalActions[index]                 = {};
            _cardInformationLayouts[index].lastNote = {};
            continue;
        }
        const auto& note = EffectiveLastNote(task);
        if (note.empty())
            _cardInformationLayouts[index].lastNote = {};
        else
            PrepareFactPair(_cardInformationLayouts[index].lastNote,
                            FormatStringResource(nullptr, IDS_FMT_FILEOPS_LAST_NOTE, note),
                            {},
                            _smallFormat.get(),
                            conflictActionWidthPx);
        // x-arch#2: failed actions keep a fallback layout; failed facts or summary text is skipped.
        if (const HRESULT actionsHr = PrepareTerminalActions(index, task, conflictActionWidthPx, actionRightToLeft); FAILED(actionsHr))
            ReportDegradedLayout(2u, actionsHr);
        if (TaskHasTerminalFacts(task))
        {
            if (const HRESULT factsHr = PrepareTerminalFacts(index, task, conflictActionWidthPx, languageRevision); FAILED(factsHr))
            {
                _terminalFacts[index] = {};
                ReportDegradedLayout(3u, factsHr);
            }
        }
        else
            _terminalFacts[index] = {};
        if (task.resultSummary.empty())
        {
            _terminalSummaries[index] = {};
            continue;
        }
        if (const HRESULT summaryHr = PrepareTerminalSummary(index, task, conflictActionWidthPx); FAILED(summaryHr))
        {
            _terminalSummaries[index] = {};
            ReportDegradedLayout(4u, summaryHr);
        }
    }

    if (! footerOnly)
        PrepareCardInformation(snapshot, conflictActionWidthPx);
    std::vector<float> cardHeights;
    cardHeights.reserve(displayRows.size());
    if (! footerOnly)
    {
        for (const PopupDisplayRow& row : displayRows)
        {
            if (row.kind == PopupDisplayRowKind::CompletedGroup)
            {
                cardHeights.push_back(DipsToPixels(_completedGroupMeasurement.layout.heightDip + 14.0f, _dpi));
                continue;
            }

            const TaskSnapshot& task = snapshot[row.taskIndex];
            if (task.kind == TaskSnapshot::Kind::Informational)
            {
                const FolderWindow::InformationalTaskUpdate& info = task.informational;
                const float expandedBase                          = task.finished ? DipsToPixels(180.0f, _dpi) : DipsToPixels(210.0f, _dpi);
                const bool displayCollapsed                       = IsTaskCollapsedForDisplay(task.taskId, compactDensity);
                float h                                           = displayCollapsed ? collapsedCardH : expandedBase;

                if (! displayCollapsed && ! task.finished && info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories &&
                    info.contentActive && info.contentInFlightCount > 1u)
                {
                    size_t activeInFlightCount = 0;
                    for (size_t i = 0; i < info.contentInFlightCount; ++i)
                    {
                        const auto& entry          = info.contentInFlight[i];
                        const bool active          = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                        const bool recentCompleted = ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes &&
                                                     entry.lastUpdateTick != 0 && renderTick >= entry.lastUpdateTick &&
                                                     (renderTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
                        if (active || recentCompleted)
                        {
                            ++activeInFlightCount;
                        }
                    }

                    const size_t lineCount = std::max<size_t>(1u, activeInFlightCount);
                    if (lineCount > 1u)
                    {
                        h += static_cast<float>(lineCount - 1u) * baseLineH;
                    }
                }

                h = ResolveAnimatedTaskHeight(task.taskId, h, renderTick, reducedMotion, ! displayCollapsed);
                if (showCompletedGroup && IsCompletedGroupTask(task))
                {
                    h *= completedGroupVisibility;
                }
                cardHeights.push_back(h);
                continue;
            }

            const bool displayCollapsed = IsTaskCollapsedForDisplay(task.taskId, compactDensity);
            float h                     = displayCollapsed                                                           ? collapsedCardH
                                          : (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) ? expandedTransferCardH
                                                                                                                     : expandedCardH;
            if (! displayCollapsed && task.finished)
            {
                // I35 P5/R9.2/R9.3: no zero counters and no raw HRESULT on the card.
                const float diagnosticLines = (task.interruptedMoveNotice || ! TaskShowsCountersLine(task)) ? 0.0f : 1.0f;
                const float pathLines       = task.operation == FILESYSTEM_DELETE ? 1.0f : 2.0f;
                // Top/header, measured facts, exact body rows, then progress,
                // action gap, dismiss action and bottom inset. Outer scrolling owns overflow.
                h = DipsToPixels(8.0f + 18.0f + 24.0f + _terminalActions[row.taskIndex].heightDip, _dpi) + (diagnosticLines + pathLines) * baseLineH +
                    FactPairHeight(_cardInformationLayouts[row.taskIndex].lastNote);
                if (const auto& summary = _terminalSummaries[row.taskIndex]; summary.layout)
                    h += summary.heightPx + DipsToPixels(10.0f, _dpi);
                if (const auto& facts = _terminalFacts[row.taskIndex].prepared; facts.layout)
                    h += facts.heightPx + DipsToPixels(10.0f, _dpi);
            }
            else if (! displayCollapsed && task.operation == FILESYSTEM_DELETE)
            {
                // Keep every painted body row above the bottom-anchored progress and actions.
                // Open discovery used to paint two extra rows into a fixed 152-DIP card.
                const bool preparing      = TaskShowsPreparingStatus(task);
                const bool summary        = TaskUsesSkippedCountingWorkSummary(task);
                const float discoveryRows = ResolveWholeTaskProgressPresentation(task).showDiscoveryActivity ? 2.0f : 0.0f;
                const float bodyRows      = 1.0f + discoveryRows + (preparing ? 1.0f : summary ? 1.0f : 2.0f);
                h = std::max(DipsToPixels(kFileOperationsExpandedDeleteCardHeightDip, _dpi), DipsToPixels(82.0f, _dpi) + bodyRows * baseLineH);
            }
            if (! displayCollapsed && ! task.conflict.active && task.operation != FILESYSTEM_COPY && task.operation != FILESYSTEM_MOVE &&
                TaskUsesSkippedCountingWorkSummary(task))
            {
                const std::wstring summary = FormatDiscoveryProgressText(task);
                const float summaryHeight  = MeasureWrappedTextHeight(_dwriteFactory.get(), _bodyFormat.get(), summary, conflictActionWidthPx, baseLineH);
                // Transfer cards formerly reserved two discovery rows. A skipped scan has one
                // semantic summary, which may wrap at narrow widths or larger text sizes.
                h += summaryHeight - (task.operation == FILESYSTEM_COPY ? 2.0f : 0.0f) * baseLineH;
            }
            if (! displayCollapsed && ! task.finished && task.conflict.active)
            {
                const ConflictPromptContent promptContent = BuildConflictPromptContent(task, ConflictInspectionResultText(task));
                const float promptContentHeight           = MeasureConflictPromptContentHeight(_dwriteFactory.get(),
                                                                                               _headerFormat.get(),
                                                                                               _bodyFormat.get(),
                                                                                               _smallFormat.get(),
                                                                                               promptContent,
                                                                                               promptMeasureWidth,
                                                                                               baseLineH,
                                                                                               promptRowGap,
                                                                                               DipsToPixels(24.0f, _dpi));
                // Header (8 + 18), body gap (4), comparison focus inset (8), action gap (8), bottom pad (10).
                // Active decisions do not reserve the ordinary transfer/graph body.
                const float decisionActionsHeight =
                    task.conflict.buttonsPublishable ? DipsToPixels(_conflictActionMeasurements[row.taskIndex].heightDip, _dpi) : 0.0f;
                const float decisionScopeHeight = task.conflict.buttonsPublishable && task.conflict.applyToAllEligible
                                                      ? DipsToPixels(_conflictActionMeasurements[row.taskIndex].scopeHeightDip + 6.0f, _dpi)
                                                      : 0.0f;
                h                               = DipsToPixels(56.0f, _dpi) + promptContentHeight + decisionActionsHeight + decisionScopeHeight;
                const float comparisonHeight    = MeasureConflictPromptContentHeight(_dwriteFactory.get(),
                                                                                     _headerFormat.get(),
                                                                                     _bodyFormat.get(),
                                                                                     _smallFormat.get(),
                                                                                     promptContent,
                                                                                     promptMeasureWidth,
                                                                                     baseLineH,
                                                                                     promptRowGap,
                                                                                     DipsToPixels(24.0f, _dpi),
                                                                                     true);
                // Retain the subject and complete decisions. Extremely small work areas still use outer
                // scrolling when even this minimum cannot fit, rather than overlapping controls.
                const float insetComparisonHeight = MeasureConflictPromptContentHeight(_dwriteFactory.get(),
                                                                                       _headerFormat.get(),
                                                                                       _bodyFormat.get(),
                                                                                       _smallFormat.get(),
                                                                                       promptContent,
                                                                                       std::max(1.0f, promptMeasureWidth - DipsToPixels(8.0f, _dpi)),
                                                                                       baseLineH,
                                                                                       promptRowGap,
                                                                                       DipsToPixels(24.0f, _dpi),
                                                                                       true);
                h += insetComparisonHeight - comparisonHeight;
                const float minimumHeight   = h - insetComparisonHeight + std::min(insetComparisonHeight, DipsToPixels(48.0f, _dpi));
                const float availableHeight = static_cast<float>(_clientSize.cy) - FooterHeightPixels() - padding * 2.0f;
                h                           = std::min(h, std::max(minimumHeight, availableHeight));
            }
            if (! displayCollapsed && ! task.conflict.active)
            {
                const auto& information = _cardInformationLayouts[row.taskIndex];
                if (! task.finished && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE))
                {
                    const auto progress = ResolveWholeTaskProgressPresentation(task);
                    const float summaryHeight =
                        ! progress.showWorkFacts ? 0.0f
                        : TaskUsesSkippedCountingWorkSummary(task)
                            ? MeasureWrappedTextHeight(
                                  _dwriteFactory.get(), _bodyFormat.get(), FormatDiscoveryProgressText(task), conflictActionWidthPx, baseLineH)
                            : FactPairHeight(information.progress) + DipsToPixels(4.0f, _dpi);
                    size_t streams = 0u;
                    for (size_t stream = 0u; progress.showTransferProgress && progress.transferActive && stream < task.inFlightFileCount; ++stream)
                        if (FileProgressIsVisible(task.inFlightFiles[stream], renderTick))
                            ++streams;
                    const float pathHeight = (static_cast<float>(std::max<size_t>(1u, streams)) + 1.0f) * baseLineH + fromToGapY;
                    const float countsHeight =
                        information.counts.rate || information.counts.eta ? FactPairHeight(information.counts) + DipsToPixels(4.0f, _dpi) : 0.0f;
                    const float progressHeight = progress.showTransferProgress ? DipsToPixels(26.0f, _dpi) : DipsToPixels(8.0f, _dpi);
                    const float plotHeight     = progress.showThroughputGraph
                                                     ? std::max(DipsToPixels(24.0f, _dpi), FactPairHeight(information.throughput)) +
                                                           DipsToPixels(8.0f + (IsTaskThroughputExpanded(task.taskId) ? 80.0f : 0.0f), _dpi)
                                                     : 0.0f;
                    h = DipsToPixels(8.0f + 18.0f + 8.0f + 24.0f + 10.0f, _dpi) + summaryHeight + countsHeight + progressHeight + pathHeight + plotHeight;
                }
                if (information.concurrency.rate)
                    h += FactPairHeight(information.concurrency) + DipsToPixels(6.0f, _dpi);
            }
            // Decision content changes immediately (metadata/inspection outcomes).
            // Snap its measured height too; interpolating a shorter previous card
            // would paint the new result through the scope checkbox/action row.
            h = ResolveAnimatedTaskHeight(task.taskId, h, renderTick, reducedMotion || (task.conflict.active && ! displayCollapsed), ! displayCollapsed);
            if (showCompletedGroup && IsCompletedGroupTask(task))
            {
                h *= completedGroupVisibility;
            }
            cardHeights.push_back(h);
        }
    }
    const uint64_t cardLayoutUs = capturePerf ? PerfElapsedUs(cardLayoutStartedUs) : 0u;

    const size_t rowCount = footerOnly ? 0u : displayRows.size();
    if (footerOnly)
    {
        _contentHeight = 0.0f;
    }
    else if (rowCount == 0)
    {
        _contentHeight = padding * 2.0f;
    }
    else
    {
        float sumHeights = 0.0f;
        for (const float h : cardHeights)
        {
            sumHeights += h;
        }
        _contentHeight = padding * 2.0f + sumHeights + static_cast<float>(rowCount - 1u) * cardGap;
    }

    // Auto-resize window to fit content (limited to screen height)
    const uint64_t autoResizeStartedUs = capturePerf ? PerfNowUs() : 0u;
    if (_preserveWindowSizeForThroughputChange)
    {
        // Disclosure changes content, not the user's saved outer geometry. A subsequent unrelated
        // task/phase change still follows the normal auto-sizing policy.
        _lastAutoSizedContentHeight            = _contentHeight;
        _lastTaskCount                         = rowCount;
        _preserveWindowSizeForThroughputChange = false;
    }
    else
    {
        AutoResizeWindow(hwnd, _contentHeight, rowCount, footerOnly, reducedMotion);
    }
    const uint64_t autoResizeUs = capturePerf ? PerfElapsedUs(autoResizeStartedUs) : 0u;

    const uint64_t scrollLayoutStartedUs = capturePerf ? PerfNowUs() : 0u;
    bool scrollReady                     = false;
    for (int pass = 0; pass < 2; ++pass)
    {
        RECT clientRc{};
        GetClientRect(hwnd, &clientRc);
        const UINT clientW = static_cast<UINT>(std::max(0L, clientRc.right - clientRc.left));
        const UINT clientH = static_cast<UINT>(std::max(0L, clientRc.bottom - clientRc.top));

        if (_target && (_clientSize.cx != static_cast<LONG>(clientW) || _clientSize.cy != static_cast<LONG>(clientH)))
        {
            _clientSize.cx = static_cast<LONG>(clientW);
            _clientSize.cy = static_cast<LONG>(clientH);
        }

        width  = static_cast<float>(clientW);
        height = static_cast<float>(clientH);

        LayoutChrome(width, height, showPauseResumeAll, showCancelAll, globalSummaryText);

        const float viewH              = std::max(0.0f, _listViewportRect.bottom - _listViewportRect.top);
        const bool shouldShowScrollBar = ! footerOnly && _contentHeight > viewH;
        if (shouldShowScrollBar != _scrollBarVisible)
        {
            _scrollBarVisible = shouldShowScrollBar;
            if (! shouldShowScrollBar)
            {
                _scrollPos = 0;
                _scrollY   = 0.0f;
            }

            ShowScrollBar(hwnd, SB_VERT, shouldShowScrollBar ? TRUE : FALSE);

            SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
            continue;
        }

        const float maxScroll = std::max(0.0f, _contentHeight - viewH);
        _scrollPos            = std::clamp(_scrollPos, 0, static_cast<int>(std::ceil(maxScroll)));
        UpdateScrollBar(hwnd, viewH, _contentHeight);
        _scrollY    = static_cast<float>(_scrollPos);
        scrollReady = true;
        break;
    }

    if (! scrollReady)
    {
        const float viewH     = std::max(0.0f, _listViewportRect.bottom - _listViewportRect.top);
        const float maxScroll = std::max(0.0f, _contentHeight - viewH);
        _scrollPos            = std::clamp(_scrollPos, 0, static_cast<int>(std::ceil(maxScroll)));
        UpdateScrollBar(hwnd, viewH, _contentHeight);
        _scrollY = static_cast<float>(_scrollPos);
    }
    const uint64_t scrollLayoutUs = capturePerf ? PerfElapsedUs(scrollLayoutStartedUs) : 0u;

    _buttons.clear();
    _completedGroupClearTaskIds.clear();
    _publishedInspectionContexts.clear();
    _hostedProgressDescriptors.clear();
    _hostedGraphDescriptors.clear();
    _hostedTooltipDescriptors.clear();
#ifdef ENABLE_TESTS
    _paintedTaskBodiesForSelfTest.clear();
    _paintedTaskCardsForSelfTest.clear();
    _paintedLiveProgressForSelfTest.clear();
    _paintedStatusChipsForSelfTest.clear();
    _graphRowColorMatchCountsForSelfTest.clear();
    _graphRowColorMismatchCountsForSelfTest.clear();
#endif

    const uint64_t drawStartedUs = capturePerf ? PerfNowUs() : 0u;
    wil::com_ptr<ID2D1CommandList> preparedBackground;
    HRESULT hrEndDraw = _target->CreateCommandList(preparedBackground.put());
    if (FAILED(hrEndDraw))
    {
        // D2DERR_RECREATE_TARGET and DXGI device removal or reset are recoverable: drop this
        // device generation and prepare again after the backoff.
        DiscardDeviceResources();
        if (OnFramePrepareFailed(hwnd, hrEndDraw))
            paintFailureSurfaceGround();
        return;
    }
    {
        // Prepare legacy chrome/text before traversing the retained tree. Recording does
        // not present or mutate that tree, and all borrowed host drawing state is restored.
        wil::com_ptr<ID2D1Image> previousTarget;
        _target->GetTarget(previousTarget.put());
        float previousDpiX = 96.0f;
        float previousDpiY = 96.0f;
        _target->GetDpi(&previousDpiX, &previousDpiY);
        D2D1_MATRIX_3X2_F previousTransform{};
        _target->GetTransform(&previousTransform);
        const auto previousTextAntialiasMode = _target->GetTextAntialiasMode();
        auto restoreTarget                   = wil::scope_exit([&]
        {
            _target->SetTarget(previousTarget.get());
            _target->SetDpi(previousDpiX, previousDpiY);
            _target->SetTransform(previousTransform);
            _target->SetTextAntialiasMode(previousTextAntialiasMode);
        });
        _target->SetTarget(preparedBackground.get());
        _target->SetDpi(96.0f, 96.0f);
        _target->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
        _target->BeginDraw();
        auto endDraw = wil::scope_exit([&] { hrEndDraw = _target->EndDraw(); });

        _target->SetTransform(D2D1::Matrix3x2F::Identity());

        const D2D1_RECT_F clientRect = D2D1::RectF(0.0f, 0.0f, width, height);
        _target->FillRectangle(clientRect, _bgBrush.get());

        const float footerH          = FooterHeightPixels();
        const float footerTop        = std::max(0.0f, height - footerH);
        const D2D1_RECT_F footerRect = D2D1::RectF(0.0f, footerTop, width, height);
        _target->DrawRectangle(footerRect, _borderBrush.get(), 1.0f);

        const GlobalTaskbarProgressModel taskbarProgress = BuildGlobalTaskbarProgressModel(globalSummary);
        ApplyTaskbarProgress(hwnd, taskbarProgress.state, taskbarProgress.completed, taskbarProgress.total);

        if (HasGlobalAggregateProgress(globalSummary) && _footerAggregateProgressRect.right > _footerAggregateProgressRect.left &&
            _footerAggregateProgressRect.bottom > _footerAggregateProgressRect.top)
        {
            _hostedProgressDescriptors.push_back(HostedProgressDescriptor{_footerAggregateProgressRect,
                                                                          0u,
                                                                          GlobalAggregateProgressFraction(globalSummary) * 100.0,
                                                                          ! HasDeterminateGlobalAggregateProgress(globalSummary)});
        }

        const auto rectHasArea = [](const D2D1_RECT_F& rc) noexcept { return rc.right > rc.left && rc.bottom > rc.top; };

        const bool pauseResumeAllPauses = FooterPauseResumeAllShouldPause(globalSummary);
        PopupButton pauseResumeAllBtn{};
        pauseResumeAllBtn.bounds   = _footerPauseResumeAllRect;
        pauseResumeAllBtn.hit.kind = PopupHitTest::Kind::FooterPauseResumeAll;
        pauseResumeAllBtn.hit.data = pauseResumeAllPauses ? kFooterPauseResumeAllPauseAction : kFooterPauseResumeAllResumeAction;
        if (showPauseResumeAll && rectHasArea(pauseResumeAllBtn.bounds))
        {
            _buttons.push_back(pauseResumeAllBtn);
        }

        PopupButton cancelAllBtn{};
        cancelAllBtn.bounds   = _footerCancelAllRect;
        cancelAllBtn.hit.kind = PopupHitTest::Kind::FooterCancelAll;
        if (showCancelAll && rectHasArea(cancelAllBtn.bounds))
        {
            _buttons.push_back(cancelAllBtn);
        }

        PopupButton queueBtn{};
        queueBtn.bounds   = _footerQueueModeRect;
        queueBtn.hit.kind = PopupHitTest::Kind::FooterQueueMode;
        if (rectHasArea(queueBtn.bounds))
        {
            _buttons.push_back(queueBtn);
        }

        PopupButton optionsBtn{};
        optionsBtn.bounds   = _footerOptionsRect;
        optionsBtn.hit.kind = PopupHitTest::Kind::FooterOptions;
        if (rectHasArea(optionsBtn.bounds))
        {
            _buttons.push_back(optionsBtn);
        }

        PopupButton detailsBtn{};
        detailsBtn.bounds   = _footerDetailsToggleRect;
        detailsBtn.hit.kind = PopupHitTest::Kind::FooterToggleDetails;
        if (rectHasArea(detailsBtn.bounds))
        {
            _buttons.push_back(detailsBtn);
        }

        if (showPauseResumeAll && rectHasArea(pauseResumeAllBtn.bounds))
        {
            DrawButton(pauseResumeAllBtn,
                       _buttonSmallFormat.get(),
                       LoadStringResource(nullptr, pauseResumeAllPauses ? IDS_FILEOPS_BTN_PAUSE_ALL : IDS_FILEOPS_BTN_RESUME_ALL));
        }

        if (showCancelAll && rectHasArea(cancelAllBtn.bounds))
        {
            DrawButton(cancelAllBtn, _buttonFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_BTN_CANCEL_ALL));
        }

        if (rectHasArea(detailsBtn.bounds))
        {
            DrawDisclosureChevron(detailsBtn.bounds, ResolveDisclosureProgress(_footerDisclosureAnimation, ! footerOnly, renderTick, reducedMotion));
        }

        const bool queueMode = fileOps ? fileOps->GetQueueNewTasks() : true;
        if (rectHasArea(queueBtn.bounds))
        {
            DrawFooterQueueModeControl(queueBtn, queueMode);
        }

        if (rectHasArea(optionsBtn.bounds))
        {
            DrawButton(optionsBtn, _buttonSmallFormat.get(), {});
            static_cast<void>(DrawCenteredGlyph(optionsBtn.bounds, FluentIcons::kMore, FluentIcons::kFallbackMore));
        }

        if (_footerSummaryLayout && _subTextBrush && _footerSummaryRect.right > _footerSummaryRect.left && _footerSummaryRect.bottom > _footerSummaryRect.top)
        {
            _target->DrawTextLayout(
                D2D1::Point2F(_footerSummaryRect.left, _footerSummaryRect.top), _footerSummaryLayout.get(), _subTextBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }

        float y = _listViewportRect.top + padding - _scrollY;

        const float cardW = std::max(0.0f, width - padding * 2.0f);
        // popup-b#3: the focused task's card keeps its hosted controls while it is scrolled out of
        // view, so scrolling never moves keyboard focus. It paints clipped by the list viewport.
        const uint64_t focusedTaskId = FocusedHostedTaskId();

        _target->PushAxisAlignedClip(_listViewportRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        for (size_t rowIndex = 0; rowIndex < rowCount; ++rowIndex)
        {
            const PopupDisplayRow& row = displayRows[rowIndex];
            const float taskCardH      = cardHeights[rowIndex];
            const D2D1_RECT_F cardRect = D2D1::RectF(padding, y, padding + cardW, y + taskCardH);

            // Keep the height transition, but retire collapsed task content and its
            // pointer/UIA controls on the first frame, not at the animation's end.
            const bool hiddenByGroup =
                showCompletedGroup && ! _completedGroupExpanded && row.kind == PopupDisplayRowKind::Task && IsCompletedGroupTask(snapshot[row.taskIndex]);
            const bool focusedCard = focusedTaskId != 0u && row.kind == PopupDisplayRowKind::Task && snapshot[row.taskIndex].taskId == focusedTaskId;
            const bool visible     = ! hiddenByGroup && ((cardRect.bottom >= _listViewportRect.top && cardRect.top <= _listViewportRect.bottom) || focusedCard);
            if (visible)
            {
#ifdef ENABLE_TESTS
                if (row.kind == PopupDisplayRowKind::Task)
                    _paintedTaskCardsForSelfTest[snapshot[row.taskIndex].taskId] = cardRect;
#endif
                _target->DrawRoundedRectangle(D2D1::RoundedRect(cardRect, DipsToPixels(2.0f, _dpi), DipsToPixels(2.0f, _dpi)), _borderBrush.get(), 1.0f);

                _target->PushAxisAlignedClip(cardRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                auto popCardClip = wil::scope_exit([&] { _target->PopAxisAlignedClip(); });

                const float padX         = DipsToPixels(10.0f, _dpi);
                const float textX        = cardRect.left + padX;
                const float contentRight = cardRect.right - padX;
                const float lineH        = DipsToPixels(18.0f, _dpi);
                float textY              = cardRect.top + DipsToPixels(8.0f, _dpi);
                const float textMaxW     = std::max(0.0f, contentRight - textX);

                if (row.kind == PopupDisplayRowKind::CompletedGroup)
                {
                    const auto& group         = _completedGroupMeasurement;
                    const auto placeGroupRect = [&](const D2D1_RECT_F& local) noexcept
                    {
                        const float top = cardRect.top + DipsToPixels(7.0f, _dpi);
                        return D2D1::RectF(textX + DipsToPixels(local.left, _dpi),
                                           top + DipsToPixels(local.top, _dpi),
                                           textX + DipsToPixels(local.right, _dpi),
                                           top + DipsToPixels(local.bottom, _dpi));
                    };
                    const D2D1_RECT_F chevronRc = placeGroupRect(group.layout.disclosure);
                    PopupButton clearBtn{};
                    clearBtn.bounds   = placeGroupRect(group.layout.clear);
                    clearBtn.hit.kind = PopupHitTest::Kind::CompletedGroupClear;

                    PopupButton groupToggle{};
                    groupToggle.bounds   = chevronRc;
                    groupToggle.hit.kind = PopupHitTest::Kind::CompletedGroupToggle;
                    if (groupToggle.bounds.right > groupToggle.bounds.left && groupToggle.bounds.bottom > groupToggle.bounds.top)
                    {
                        _buttons.push_back(groupToggle);
                    }

                    if (clearBtn.bounds.right > clearBtn.bounds.left && clearBtn.bounds.bottom > clearBtn.bounds.top)
                    {
                        _buttons.push_back(clearBtn);
                        DrawButton(clearBtn, _buttonSmallFormat.get(), group.clearText);
                        for (const TaskSnapshot& member : snapshot)
                        {
                            if (IsCompletedGroupTask(member))
                                _completedGroupClearTaskIds.push_back(member.taskId);
                        }
                    }
                    if (_completedGroupExpanded)
                    {
                        const D2D1_RECT_F textRc = placeGroupRect(group.layout.heading);
                        if (group.headingLayout && _textBrush)
                            _target->DrawTextLayout(
                                D2D1::Point2F(textRc.left, textRc.top), group.headingLayout.get(), _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }
                    else
                    {
                        struct ResultBadge final
                        {
                            uint32_t count             = 0u;
                            PopupStatusVisualTone tone = PopupStatusVisualTone::None;
                            wchar_t fluentGlyph        = 0;
                            wchar_t fallbackGlyph      = 0;
                            UINT tooltipResourceId     = 0u;
                            std::wstring_view automationSuffix;
                        };
                        const std::array badges = {
                            ResultBadge{completedGroupSummary.completed,
                                        PopupStatusVisualTone::Ok,
                                        FluentIcons::kCheckMark,
                                        FluentIcons::kFallbackCheckMark,
                                        IDS_FMT_FILEOPS_BADGE_COMPLETED_TOOLTIP,
                                        L"Completed"},
                            ResultBadge{completedGroupSummary.partial,
                                        PopupStatusVisualTone::Warning,
                                        FluentIcons::kWarning,
                                        FluentIcons::kFallbackWarning,
                                        IDS_FMT_FILEOPS_BADGE_PARTIAL_TOOLTIP,
                                        L"Partial"},
                            ResultBadge{completedGroupSummary.indeterminate,
                                        PopupStatusVisualTone::Warning,
                                        FluentIcons::kInfo,
                                        L'?',
                                        IDS_FMT_FILEOPS_BADGE_INDETERMINATE_TOOLTIP,
                                        L"Indeterminate"},
                            ResultBadge{completedGroupSummary.failed,
                                        PopupStatusVisualTone::Error,
                                        FluentIcons::kError,
                                        FluentIcons::kFallbackError,
                                        IDS_FMT_FILEOPS_BADGE_FAILED_TOOLTIP,
                                        L"Failed"},
                            ResultBadge{completedGroupSummary.canceled,
                                        PopupStatusVisualTone::Neutral,
                                        FluentIcons::kClear,
                                        L'\u00D7',
                                        IDS_FMT_FILEOPS_BADGE_CANCELED_TOOLTIP,
                                        L"Canceled"},
                        };
                        size_t visibleBadgeIndex = 0u;
                        for (const ResultBadge& badge : badges)
                        {
                            if (badge.count == 0u)
                            {
                                continue;
                            }
                            // A degraded group layout (x-arch#2) arranges no badges.
                            if (visibleBadgeIndex >= group.layout.badgeCount)
                                break;

                            const D2D1_RECT_F badgeRc = placeGroupRect(group.layout.badges[visibleBadgeIndex++]);
                            D2D1_COLOR_F color        = StatusVisualColorForTone(appTheme, badge.tone);
                            if (_graphDynamicBrush)
                            {
                                color.a = appTheme.highContrast ? 0.30f : 0.14f;
                                _graphDynamicBrush->SetColor(color);
                                const float radius = ClampCornerRadius(badgeRc, DipsToPixels(7.0f, _dpi));
                                _target->FillRoundedRectangle(D2D1::RoundedRect(badgeRc, radius, radius), _graphDynamicBrush.get());
                                color.a = appTheme.highContrast ? 1.0f : 0.65f;
                                _graphDynamicBrush->SetColor(color);
                                _target->DrawRoundedRectangle(D2D1::RoundedRect(badgeRc, radius, radius), _graphDynamicBrush.get(), 1.0f);
                            }

                            const float iconW        = DipsToPixels(18.0f, _dpi);
                            const D2D1_RECT_F iconRc = D2D1::RectF(
                                badgeRc.left + DipsToPixels(4.0f, _dpi), badgeRc.top, badgeRc.left + DipsToPixels(4.0f, _dpi) + iconW, badgeRc.bottom);
                            const bool useFluent =
                                _statusIconFormat && DirectWriteFormatHasGlyph(_dwriteFactory.get(), _statusIconFormat.get(), badge.fluentGlyph);
                            const wchar_t glyph           = useFluent ? badge.fluentGlyph : badge.fallbackGlyph;
                            IDWriteTextFormat* iconFormat = useFluent ? _statusIconFormat.get() : _statusIconFallbackFormat.get();
                            color.a                       = 1.0f;
#ifdef ENABLE_TESTS
                            wchar_t paintedGlyph = L'\0';
#endif
                            if (iconFormat && glyph != 0 && _graphDynamicBrush)
                            {
                                _graphDynamicBrush->SetColor(color);
                                const wchar_t iconText[2]{glyph, 0};
                                _target->DrawTextW(iconText, 1u, iconFormat, iconRc, _graphDynamicBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
#ifdef ENABLE_TESTS
                                paintedGlyph = glyph;
#endif
                            }
                            const std::wstring countText = std::to_wstring(badge.count);
                            const D2D1_RECT_F countRc    = D2D1::RectF(iconRc.right, badgeRc.top, badgeRc.right - DipsToPixels(4.0f, _dpi), badgeRc.bottom);
                            if (_smallFormat && _textBrush)
                            {
                                _target->DrawTextW(countText.data(),
                                                   static_cast<UINT32>(countText.size()),
                                                   _smallFormat.get(),
                                                   countRc,
                                                   _textBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            }
                            _hostedTooltipDescriptors.push_back(
                                HostedTooltipDescriptor{.bounds = badgeRc,
                                                        .text = FormatStringResource(nullptr, badge.tooltipResourceId, static_cast<unsigned long>(badge.count)),
                                                        .automationId = std::format(L"FileOperations.CompletedGroup.Badge.{}", badge.automationSuffix)});
#ifdef ENABLE_TESTS
                            _hostedTooltipDescriptors.back().paintedGlyph       = paintedGlyph;
                            _hostedTooltipDescriptors.back().paintedLabel       = countText;
                            _hostedTooltipDescriptors.back().paintedLabelBounds = countRc;
#endif
                        }
                    }

                    const float gapAfter = (rowIndex + 1u < rowCount) ? cardGap : 0.0f;
                    y += taskCardH + gapAfter;
                    continue;
                }

                const TaskSnapshot& task   = snapshot[row.taskIndex];
                const bool isCollapsedTask = IsTaskCollapsedForDisplay(task.taskId, compactDensity);

                if (task.kind == TaskSnapshot::Kind::Informational)
                {
                    const FolderWindow::InformationalTaskUpdate& info = task.informational;
                    const ULONGLONG nowTick                           = renderTick;

                    std::wstring headerText = info.title;
                    if (info.finished && ! info.title.empty())
                    {
                        const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        std::wstring statusText;
                        if (SUCCEEDED(info.resultHr))
                        {
                            statusText = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
                        }
                        else if (info.resultHr == cancelledHr || info.resultHr == E_ABORT)
                        {
                            statusText = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_CANCELED);
                        }
                        else
                        {
                            // P5: no raw HRESULT on the card; it stays in Technical details and exports.
                            statusText = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_FAILED_SHORT);
                        }
                        headerText = FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, info.title, statusText);
                    }

                    const float collapseBtnSize = DipsToPixels(18.0f, _dpi);
                    const float collapseBtnGap  = DipsToPixels(6.0f, _dpi);

                    const float headerTop    = isCollapsedTask ? cardRect.top + (taskCardH - lineH) * 0.5f : textY;
                    const float headerBottom = headerTop + lineH;
                    const float collapseTop  = headerTop + (lineH - collapseBtnSize) * 0.5f;
                    const float collapseLeft = std::max(textX, contentRight - collapseBtnSize);

                    PopupButton collapseBtn{};
                    collapseBtn.bounds     = D2D1::RectF(collapseLeft, collapseTop, contentRight, collapseTop + collapseBtnSize);
                    collapseBtn.hit.kind   = PopupHitTest::Kind::TaskToggleCollapse;
                    collapseBtn.hit.taskId = task.taskId;
                    _buttons.push_back(collapseBtn);

                    const float headerRight = std::max(textX, collapseBtn.bounds.left - collapseBtnGap);
                    float headerLeft        = textX;

                    CaptionStatus statusIcon = CaptionStatus::None;
                    {
                        const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        if (info.finished && FAILED(info.resultHr) && info.resultHr != cancelledHr && info.resultHr != E_ABORT)
                        {
                            statusIcon = CaptionStatus::Error;
                        }
                        else if (info.finished && SUCCEEDED(info.resultHr))
                        {
                            statusIcon = CaptionStatus::Ok;
                        }
                    }

                    if (statusIcon != CaptionStatus::None && _target)
                    {
                        const float iconSize = DipsToPixels(16.0f, _dpi);
                        const float iconGap  = DipsToPixels(6.0f, _dpi);

                        D2D1_RECT_F iconRc = D2D1::RectF(textX, headerTop, textX + iconSize, headerBottom);
                        iconRc.right       = std::min(iconRc.right, headerRight);

                        wchar_t fluentGlyph = 0;
                        wchar_t fallback    = 0;
                        ID2D1Brush* brush   = _textBrush.get();
                        switch (statusIcon)
                        {
                            case CaptionStatus::Ok:
                                fluentGlyph = FluentIcons::kCheckMark;
                                fallback    = FluentIcons::kFallbackCheckMark;
                                brush       = _statusOkBrush ? _statusOkBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                                break;
                            case CaptionStatus::Error:
                                fluentGlyph = FluentIcons::kError;
                                fallback    = FluentIcons::kFallbackError;
                                brush       = _statusErrorBrush ? _statusErrorBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                                break;
                            case CaptionStatus::Warning:
                            case CaptionStatus::None:
                            default: break;
                        }

                        const bool useFluentFormat =
                            _statusIconFormat != nullptr && DirectWriteFormatHasGlyph(_dwriteFactory.get(), _statusIconFormat.get(), fluentGlyph);
                        const wchar_t glyph       = useFluentFormat ? fluentGlyph : fallback;
                        IDWriteTextFormat* format = useFluentFormat ? _statusIconFormat.get() : _statusIconFallbackFormat.get();

                        if (format && brush && glyph != 0 && iconRc.right > iconRc.left)
                        {
                            const wchar_t text[2]{glyph, 0};
                            _target->DrawText(text, 1u, format, iconRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            headerLeft = std::min(headerRight, iconRc.right + iconGap);
                        }
                    }

                    if (_headerFormat)
                    {
                        const D2D1_RECT_F headerRc = D2D1::RectF(headerLeft, headerTop, headerRight, headerBottom);
                        _target->DrawTextW(headerText.data(),
                                           static_cast<UINT32>(headerText.size()),
                                           _headerFormat.get(),
                                           headerRc,
                                           _textBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }

                    if (isCollapsedTask)
                    {
                        const float gapAfter = (rowIndex + 1u < rowCount) ? cardGap : 0.0f;
                        y += taskCardH + gapAfter;
                        continue;
                    }

                    textY = headerBottom + DipsToPixels(6.0f, _dpi);

                    const auto drawLabeledPathLine = [&](UINT labelId, const std::filesystem::path& path) noexcept
                    {
                        if (! _dwriteFactory || ! _smallFormat || ! _bodyFormat || ! _textBrush || ! _subTextBrush)
                        {
                            return;
                        }

                        if (textY + lineH > cardRect.bottom)
                        {
                            return;
                        }

                        const std::wstring label = LoadStringResource(nullptr, labelId);
                        if (label.empty())
                        {
                            return;
                        }

                        const float labelW   = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), label, textMaxW, lineH);
                        const float labelGap = DipsToPixels(6.0f, _dpi);
                        const float pathLeft = textX + labelW + labelGap;
                        const float pathW    = std::max(0.0f, contentRight - pathLeft);

                        const D2D1_RECT_F labelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(
                            label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                        const std::wstring pathText = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), path.native(), pathW, lineH);
                        const D2D1_RECT_F pathRc    = D2D1::RectF(pathLeft, textY, contentRight, textY + lineH);
                        _target->DrawTextW(
                            pathText.data(), static_cast<UINT32>(pathText.size()), _bodyFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                        textY += lineH;
                    };

                    // Keep the localized label and counters visible while abbreviating only
                    // the changing path. The full status remains available on its painted row.
                    const auto drawScanStatus = [&](const std::filesystem::path& path, uint64_t folders, uint64_t entries) noexcept
                    {
                        if (! _dwriteFactory || ! _smallFormat || ! _textBrush || textY + lineH > cardRect.bottom)
                            return;
                        const float width           = std::max(0.0f, contentRight - textX);
                        const std::wstring rawPath  = path.empty() ? std::wstring(L".") : path.native();
                        const std::wstring fullText = FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, rawPath, folders, entries);
                        auto& cached                = _scanStatusLayouts[row.taskIndex];
                        auto& prepared              = cached.prepared;
                        if (! prepared.layout || prepared.taskId != task.taskId || prepared.text != fullText || prepared.widthPx != width ||
                            prepared.heightPx != lineH || prepared.format.get() != _smallFormat.get())
                        {
                            // x-perf#7: the counters change on every scan frame, but the path truncation
                            // depends only on the path, the row and the scaffold width. Digits share one
                            // width, so the scaffold with its digits normalized keys the measured path.
                            const std::wstring scaffold = FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, L"", folders, entries);
                            std::wstring scaffoldKey    = scaffold;
                            std::ranges::replace_if(scaffoldKey, [](wchar_t ch) noexcept { return ch >= L'0' && ch <= L'9'; }, L'0');
                            const bool truncationCurrent = prepared.layout && prepared.taskId == task.taskId && prepared.widthPx == width &&
                                                           prepared.heightPx == lineH && prepared.format.get() == _smallFormat.get() &&
                                                           ! cached.wholeLineAbbreviated && cached.measuredPath == rawPath &&
                                                           cached.measuredScaffold == scaffoldKey;
                            std::wstring painted;
                            if (truncationCurrent)
                                painted = FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, cached.paintedPath, folders, entries);
                            else
                            {
                                const float fixedWidth = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), scaffold, 65536.0f, lineH);
                                cached.paintedPath     = TruncatePathMiddleToWidth(
                                    _dwriteFactory.get(), _smallFormat.get(), rawPath, std::max(0.0f, width - fixedWidth - 1.0f), lineH);
                                painted = FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, cached.paintedPath, folders, entries);
                                // Shaping and a very narrow row can exceed the fixed scaffold budget.
                                cached.wholeLineAbbreviated = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), painted, 65536.0f, lineH) > width;
                                if (cached.wholeLineAbbreviated)
                                    painted = TruncateTextMiddleToWidth(_dwriteFactory.get(), _smallFormat.get(), painted, width, lineH, kEllipsisText, 0u);
                                cached.measuredPath     = rawPath;
                                cached.measuredScaffold = std::move(scaffoldKey);
                                ++_scanStatusMeasurementCount;
                            }
                            wil::com_ptr<IDWriteTextLayout> layout;
                            if (FAILED(_dwriteFactory->CreateTextLayout(
                                    painted.data(), static_cast<UINT32>(painted.size()), _smallFormat.get(), width, lineH, layout.put())))
                            {
                                // Keep the measured row and its accessible text even if
                                // retention fails; the next paint retries layout creation.
                                prepared = {};
                            }
                            else
                            {
                                prepared.taskId   = task.taskId;
                                prepared.text     = fullText;
                                prepared.widthPx  = width;
                                prepared.heightPx = lineH;
                                prepared.format   = _smallFormat;
                                prepared.layout   = std::move(layout);
                            }
                            cached.paintedText = std::move(painted);
                        }
                        const D2D1_RECT_F bounds = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                        cached.usedThisPaint     = true;
                        if (prepared.layout)
                            _target->DrawTextLayout(D2D1::Point2F(textX, textY), prepared.layout.get(), _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        else
                            _target->DrawTextW(cached.paintedText.data(),
                                               static_cast<UINT32>(cached.paintedText.size()),
                                               _smallFormat.get(),
                                               bounds,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        _hostedTooltipDescriptors.push_back({bounds,
                                                             fullText,
                                                             std::format(L"FileOperations.Task.{}.ScanStatus", task.taskId),
                                                             path.native(),
                                                             0u,
                                                             task.taskId,
                                                             {},
                                                             ! task.informational.finished});
#ifdef ENABLE_TESTS
                        _hostedTooltipDescriptors.back().paintedPath            = cached.paintedText;
                        _hostedTooltipDescriptors.back().paintedContainerBounds = cardRect;
#endif
                        textY += lineH;
                    };
                    if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories)
                    {
                        drawLabeledPathLine(IDS_PREFS_PANES_HEADER_LEFT, info.leftRoot);
                        drawLabeledPathLine(IDS_PREFS_PANES_HEADER_RIGHT, info.rightRoot);

                        if (_smallFormat && _textBrush && (info.scanActive || info.scanFolderCount > 0 || info.scanEntryCount > 0))
                        {
                            drawScanStatus(info.scanCurrentRelative, info.scanFolderCount, info.scanEntryCount);
                        }

                        if (_smallFormat && _subTextBrush && ! info.finished && info.scanElapsedSeconds.has_value())
                        {
                            if (textY + lineH <= cardRect.bottom)
                            {
                                const std::wstring duration = FormatDurationHms(info.scanElapsedSeconds.value());
                                if (! duration.empty())
                                {
                                    const std::wstring elapsedText = FormatStringResource(nullptr, IDS_FMT_COMPARE_ELAPSED, duration);
                                    const D2D1_RECT_F elapsedRc    = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                    _target->DrawTextW(elapsedText.data(),
                                                       static_cast<UINT32>(elapsedText.size()),
                                                       _smallFormat.get(),
                                                       elapsedRc,
                                                       _subTextBrush.get(),
                                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                    textY += lineH;
                                }
                            }
                        }

                        if (_smallFormat && _subTextBrush && (info.scanCandidateFileCount > 0 || info.scanCandidateTotalBytes > 0))
                        {
                            if (textY + lineH <= cardRect.bottom)
                            {
                                const std::wstring totalBytes = FormatBytesCompact(info.scanCandidateTotalBytes);
                                const std::wstring candidateText =
                                    FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_CANDIDATES_STATUS, info.scanCandidateFileCount, totalBytes);
                                const D2D1_RECT_F candidatesRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                _target->DrawTextW(candidateText.data(),
                                                   static_cast<UINT32>(candidateText.size()),
                                                   _smallFormat.get(),
                                                   candidatesRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                textY += lineH;
                            }
                        }
                    }
                    else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeCase)
                    {
                        if (! info.changeCaseCurrentPath.empty())
                        {
                            drawLabeledPathLine(IDS_FILEOPS_LABEL_FROM, info.changeCaseCurrentPath);
                        }

                        if (_smallFormat && _textBrush &&
                            (info.changeCaseEnumerating || info.changeCaseScannedFolders > 0 || info.changeCaseScannedEntries > 0))
                        {
                            drawScanStatus(info.changeCaseCurrentPath, info.changeCaseScannedFolders, info.changeCaseScannedEntries);
                        }

                        if (_smallFormat && _subTextBrush &&
                            (info.changeCaseRenaming || info.changeCasePlannedRenames > 0 || info.changeCaseCompletedRenames > 0))
                        {
                            if (textY + lineH <= cardRect.bottom)
                            {
                                const std::wstring countsText =
                                    info.changeCasePlannedRenames > 0
                                        ? FormatEmbeddedStringResource(
                                              nullptr, IDS_FMT_FILEOPS_OP_COUNTS, info.title, info.changeCaseCompletedRenames, info.changeCasePlannedRenames)
                                        : FormatEmbeddedStringResource(
                                              nullptr, IDS_FMT_FILEOPS_OP_COUNTS_UNKNOWN_TOTAL, info.title, info.changeCaseCompletedRenames);
                                const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                _target->DrawTextW(countsText.data(),
                                                   static_cast<UINT32>(countsText.size()),
                                                   _smallFormat.get(),
                                                   countsRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                textY += lineH;
                            }
                        }
                    }
                    else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeAttributes)
                    {
                        if (! info.changeAttributesCurrentPath.empty())
                        {
                            drawLabeledPathLine(IDS_FILEOPS_LABEL_FROM, info.changeAttributesCurrentPath);
                        }

                        if (_smallFormat && _textBrush &&
                            (info.changeAttributesEnumerating || info.changeAttributesScannedFolders > 0 || info.changeAttributesScannedEntries > 0))
                        {
                            drawScanStatus(info.changeAttributesCurrentPath, info.changeAttributesScannedFolders, info.changeAttributesScannedEntries);
                        }

                        if (_smallFormat && _subTextBrush &&
                            (info.changeAttributesApplying || info.changeAttributesPlannedItems > 0 || info.changeAttributesCompletedItems > 0))
                        {
                            if (textY + lineH <= cardRect.bottom)
                            {
                                const std::wstring countsText =
                                    info.changeAttributesPlannedItems > 0
                                        ? FormatEmbeddedStringResource(nullptr,
                                                                       IDS_FMT_FILEOPS_OP_COUNTS,
                                                                       info.title,
                                                                       info.changeAttributesCompletedItems,
                                                                       info.changeAttributesPlannedItems)
                                        : FormatEmbeddedStringResource(
                                              nullptr, IDS_FMT_FILEOPS_OP_COUNTS_UNKNOWN_TOTAL, info.title, info.changeAttributesCompletedItems);
                                const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                _target->DrawTextW(countsText.data(),
                                                   static_cast<UINT32>(countsText.size()),
                                                   _smallFormat.get(),
                                                   countsRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                textY += lineH;
                            }
                        }
                    }
                    else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::MakeFileList)
                    {
                        if (! info.makeFileListCurrentPath.empty())
                        {
                            drawLabeledPathLine(IDS_FILEOPS_LABEL_FROM, info.makeFileListCurrentPath);
                        }

                        if (_smallFormat && _textBrush &&
                            (info.makeFileListCollecting || info.makeFileListScannedFolders > 0u || info.makeFileListScannedEntries > 0u))
                        {
                            drawScanStatus(info.makeFileListCurrentPath, info.makeFileListScannedFolders, info.makeFileListScannedEntries);
                        }

                        if (_smallFormat && _subTextBrush && (info.makeFileListRendering || info.makeFileListWriting || info.makeFileListTotalEntries > 0u))
                        {
                            const std::wstring countsText =
                                info.makeFileListTotalEntries > 0u
                                    ? FormatEmbeddedStringResource(
                                          nullptr, IDS_FMT_FILEOPS_OP_COUNTS, info.title, info.makeFileListRenderedEntries, info.makeFileListTotalEntries)
                                    : FormatEmbeddedStringResource(
                                          nullptr, IDS_FMT_FILEOPS_OP_COUNTS_UNKNOWN_TOTAL, info.title, info.makeFileListRenderedEntries);
                            const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(countsText.data(),
                                               static_cast<UINT32>(countsText.size()),
                                               _smallFormat.get(),
                                               countsRc,
                                               _subTextBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }
                    }
                    if (_smallFormat && _textBrush && info.contentActive)
                    {
                        std::array<size_t, FolderWindow::InformationalTaskUpdate::kMaxContentInFlightFiles> activeInFlightIndices{};
                        size_t activeInFlightCount = 0;
                        for (size_t i = 0; i < info.contentInFlightCount && activeInFlightCount < activeInFlightIndices.size(); ++i)
                        {
                            const auto& entry          = info.contentInFlight[i];
                            const bool active          = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                            const bool recentCompleted = ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes &&
                                                         entry.lastUpdateTick != 0 && nowTick >= entry.lastUpdateTick &&
                                                         (nowTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
                            if (active || recentCompleted)
                            {
                                activeInFlightIndices[activeInFlightCount] = i;
                                ++activeInFlightCount;
                            }
                        }

                        if (activeInFlightCount == 0u)
                        {
                            const std::wstring contentPath = info.contentCurrentRelative.empty() ? std::wstring{} : info.contentCurrentRelative.native();
                            const std::wstring bytesRead   = FormatBytesCompact(info.contentCurrentCompletedBytes);
                            std::wstring contentText;
                            if (info.contentCurrentTotalBytes > 0)
                            {
                                const std::wstring bytesTotal = FormatBytesCompact(info.contentCurrentTotalBytes);
                                contentText = FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_STATUS, contentPath, bytesRead, bytesTotal);
                            }
                            else
                            {
                                contentText = FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_STATUS_UNKNOWN, contentPath, bytesRead);
                            }

                            const D2D1_RECT_F contentRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(contentText.data(),
                                               static_cast<UINT32>(contentText.size()),
                                               _smallFormat.get(),
                                               contentRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }
                        else
                        {
                            const float rightEdge       = textX + textMaxW;
                            const float miniBarGap      = DipsToPixels(8.0f, _dpi);
                            const float miniBarWDesired = DipsToPixels(92.0f, _dpi);
                            const float miniBarH        = DipsToPixels(6.0f, _dpi);

                            for (size_t i = 0; i < activeInFlightCount; ++i)
                            {
                                if (textY + lineH > cardRect.bottom)
                                {
                                    break;
                                }

                                const auto& entry = info.contentInFlight[activeInFlightIndices[i]];

                                const std::wstring_view sourcePathText = entry.relativePath.native();
                                const uint64_t fileTotalBytes          = entry.totalBytes;
                                const uint64_t fileCompletedBytes      = entry.completedBytes;

                                const float availableW     = std::max(0.0f, rightEdge - textX);
                                const float miniBarWMin    = DipsToPixels(40.0f, _dpi);
                                const float minTextW       = DipsToPixels(48.0f, _dpi);
                                float miniBarW             = std::min(miniBarWDesired, availableW);
                                const float maxBarWithText = std::max(0.0f, availableW - miniBarGap - minTextW);
                                if (maxBarWithText > 0.0f)
                                {
                                    miniBarW = std::clamp(miniBarW, std::min(miniBarWMin, maxBarWithText), maxBarWithText);
                                }

                                if (fileTotalBytes > 0u && fileCompletedBytes >= fileTotalBytes)
                                {
                                    miniBarW = 0.0f;
                                }

                                const float barRight  = rightEdge;
                                const float barLeft   = barRight - miniBarW;
                                const float pathRight = (miniBarW > 0.0f) ? std::max(textX, barLeft - miniBarGap) : rightEdge;
                                const float pathW     = std::max(0.0f, pathRight - textX);

                                const std::wstring fromPath = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), sourcePathText, pathW, lineH);
                                const D2D1_RECT_F pathRc    = D2D1::RectF(textX, textY, textX + pathW, textY + lineH);
                                _target->DrawTextW(fromPath.data(),
                                                   static_cast<UINT32>(fromPath.size()),
                                                   _bodyFormat.get(),
                                                   pathRc,
                                                   _textBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);

                                if (miniBarW > 0.0f && _progressBgBrush && _progressItemBrush)
                                {
                                    const float barTop          = textY + (lineH - miniBarH) * 0.5f;
                                    const D2D1_RECT_F miniBarRc = D2D1::RectF(barLeft, barTop, barRight, barTop + miniBarH);

                                    const float radiusTrack = ClampCornerRadius(miniBarRc, DipsToPixels(2.0f, _dpi));
                                    _target->FillRoundedRectangle(D2D1::RoundedRect(miniBarRc, radiusTrack, radiusTrack), _progressBgBrush.get());

                                    const bool hasTotal = fileTotalBytes > 0u;
                                    const float frac =
                                        hasTotal && fileCompletedBytes <= fileTotalBytes
                                            ? Clamp01(static_cast<float>(static_cast<double>(fileCompletedBytes) / static_cast<double>(fileTotalBytes)))
                                            : 0.0f;

                                    if (appTheme.menu.rainbowMode)
                                    {
                                        const D2D1_COLOR_F rainbow = RainbowProgressColor(appTheme, sourcePathText);
                                        _progressItemBrush->SetColor(rainbow);
                                    }
                                    else
                                    {
                                        _progressItemBrush->SetColor(_progressItemBaseColor);
                                    }

                                    const D2D1_RECT_F fill =
                                        hasTotal
                                            ? D2D1::RectF(
                                                  miniBarRc.left, miniBarRc.top, miniBarRc.left + (miniBarRc.right - miniBarRc.left) * frac, miniBarRc.bottom)
                                            : ComputeIndeterminateBarFill(miniBarRc, nowTick, reducedMotion);
                                    const float radiusFill = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                                    _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radiusFill, radiusFill), _progressItemBrush.get());
                                }

                                textY += lineH;
                            }
                        }
                    }

                    if (_smallFormat && _subTextBrush && (info.contentPendingCount > 0 || info.contentCompletedCount > 0))
                    {
                        if (textY + lineH <= cardRect.bottom)
                        {
                            const std::wstring countsText =
                                FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_COUNTS_STATUS, info.contentPendingCount, info.contentCompletedCount);
                            const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(countsText.data(),
                                               static_cast<UINT32>(countsText.size()),
                                               _smallFormat.get(),
                                               countsRc,
                                               _subTextBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }
                    }

                    if (_smallFormat && _subTextBrush && info.contentTotalBytes > 0)
                    {
                        if (textY + lineH <= cardRect.bottom)
                        {
                            const std::wstring completedBytes = FormatBytesCompact(info.contentCompletedBytes);
                            const std::wstring totalBytes     = FormatBytesCompact(info.contentTotalBytes);
                            const std::wstring bytesText =
                                FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_TOTAL_BYTES_STATUS, completedBytes, totalBytes);
                            const D2D1_RECT_F bytesRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(bytesText.data(),
                                               static_cast<UINT32>(bytesText.size()),
                                               _smallFormat.get(),
                                               bytesRc,
                                               _subTextBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }
                    }

                    if (_smallFormat && _subTextBrush && ! info.finished && info.contentEtaSeconds.has_value())
                    {
                        if (textY + lineH <= cardRect.bottom)
                        {
                            const std::wstring duration = FormatDurationHms(info.contentEtaSeconds.value());
                            if (! duration.empty())
                            {
                                const std::wstring etaText = FormatStringResource(nullptr, IDS_FMT_COMPARE_ETA, duration);
                                const D2D1_RECT_F etaRc    = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                                _target->DrawTextW(etaText.data(),
                                                   static_cast<UINT32>(etaText.size()),
                                                   _smallFormat.get(),
                                                   etaRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                                textY += lineH;
                            }
                        }
                    }

                    if (_smallFormat && _subTextBrush && info.finished && ! info.doneSummary.empty())
                    {
                        const D2D1_RECT_F doneRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                        _target->DrawTextW(info.doneSummary.data(),
                                           static_cast<UINT32>(info.doneSummary.size()),
                                           _smallFormat.get(),
                                           doneRc,
                                           _subTextBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        textY += lineH;
                    }

                    const bool showProgressBar =
                        ! info.finished &&
                        ((info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories && (info.scanActive || info.contentActive)) ||
                         (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeCase && (info.changeCaseEnumerating || info.changeCaseRenaming)) ||
                         (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeAttributes &&
                          (info.changeAttributesEnumerating || info.changeAttributesApplying)) ||
                         (info.kind == FolderWindow::InformationalTaskUpdate::Kind::MakeFileList &&
                          (info.makeFileListCollecting || info.makeFileListRendering || info.makeFileListWriting)));
                    if (showProgressBar)
                    {
                        const float barH        = DipsToPixels(8.0f, _dpi);
                        const float bottomPad   = DipsToPixels(10.0f, _dpi);
                        const float barBottom   = cardRect.bottom - bottomPad;
                        const float barTop      = barBottom - barH;
                        const D2D1_RECT_F barRc = D2D1::RectF(textX, barTop, contentRight, barBottom);

                        bool hostedHasTotal  = false;
                        float hostedFraction = 0.0f;
                        if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories)
                        {
                            hostedHasTotal = info.contentTotalBytes > 0 && info.contentCompletedBytes <= info.contentTotalBytes;
                            hostedFraction =
                                hostedHasTotal
                                    ? Clamp01(static_cast<float>(static_cast<double>(info.contentCompletedBytes) / static_cast<double>(info.contentTotalBytes)))
                                    : 0.0f;
                        }
                        else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeCase)
                        {
                            hostedHasTotal = info.changeCasePlannedRenames > 0 && info.changeCaseCompletedRenames <= info.changeCasePlannedRenames;
                            hostedFraction = hostedHasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.changeCaseCompletedRenames) /
                                                                                         static_cast<double>(info.changeCasePlannedRenames)))
                                                            : 0.0f;
                        }
                        else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeAttributes)
                        {
                            hostedHasTotal = info.changeAttributesPlannedItems > 0 && info.changeAttributesCompletedItems <= info.changeAttributesPlannedItems;
                            hostedFraction = hostedHasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.changeAttributesCompletedItems) /
                                                                                         static_cast<double>(info.changeAttributesPlannedItems)))
                                                            : 0.0f;
                        }
                        else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::MakeFileList)
                        {
                            hostedHasTotal = info.makeFileListRendering && info.makeFileListTotalEntries > 0u &&
                                             info.makeFileListRenderedEntries <= info.makeFileListTotalEntries;
                            hostedFraction = hostedHasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.makeFileListRenderedEntries) /
                                                                                         static_cast<double>(info.makeFileListTotalEntries)))
                                                            : 0.0f;
                        }
                        _hostedProgressDescriptors.push_back(
                            HostedProgressDescriptor{barRc, task.taskId, static_cast<double>(hostedFraction) * 100.0, ! hostedHasTotal});
                    }

                    if (info.finished)
                    {
                        const float dismissButtonH         = DipsToPixels(24.0f, _dpi);
                        const float dismissButtonBottomPad = DipsToPixels(8.0f, _dpi);
                        const float dismissButtonTop       = cardRect.bottom - dismissButtonBottomPad - dismissButtonH;

                        PopupButton dismissBtn{};
                        dismissBtn.bounds     = D2D1::RectF(textX, dismissButtonTop, contentRight, dismissButtonTop + dismissButtonH);
                        dismissBtn.hit.kind   = PopupHitTest::Kind::TaskDismiss;
                        dismissBtn.hit.taskId = task.taskId;
                        _buttons.push_back(dismissBtn);
                        DrawButton(dismissBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOP_BTN_DISMISS));
                    }

                    const float gapAfter = (rowIndex + 1u < rowCount) ? cardGap : 0.0f;
                    y += taskCardH + gapAfter;
                    continue;
                }

                const UINT pauseId           = task.paused ? static_cast<UINT>(IDS_FILEOP_BTN_RESUME) : static_cast<UINT>(IDS_FILEOP_BTN_PAUSE);
                const std::wstring pauseText = LoadStringResource(nullptr, pauseId);

                const std::wstring cancelText = LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);

                const bool showCopyMoveControls = task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE;
                std::wstring speedLimitText;
                if (showCopyMoveControls)
                {
                    speedLimitText = FormatSpeedLimitSelectorText(task.desiredSpeedLimitBytesPerSecond);
                }
                const float speedLimitWidthPx =
                    showCopyMoveControls && ! task.finished && ! task.cancellationRequested && ! task.conflict.active && ! isCollapsedTask
                        ? DipsToPixels(MeasureSpeedLimitWidthDip(row.taskIndex, speedLimitText), _dpi)
                        : 0.0f;

                const UINT opTextId = [&]() -> UINT
                {
                    switch (task.operation)
                    {
                        case FILESYSTEM_COPY: return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
                        case FILESYSTEM_MOVE: return static_cast<UINT>(IDS_FILEOP_OPERATION_MOVE);
                        case FILESYSTEM_DELETE: return static_cast<UINT>(IDS_FILEOP_OPERATION_DELETE);
                        case FILESYSTEM_RENAME: return static_cast<UINT>(IDS_FILEOP_OPERATION_RENAME);
                        case FILESYSTEM_CREATE_DIRECTORY: return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
                    }
                    return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
                }();

                const std::wstring opText = LoadStringResource(nullptr, opTextId);
                const ULONGLONG nowTick   = renderTick;

                const TaskStatusKind taskStatus = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
                std::wstring headerText = taskStatus == TaskStatusKind::Conflict && ! isCollapsedTask ? opText : BuildTaskHeaderText(task, opText, nowTick);
                const PopupStatusVisualTone statusTone = StatusVisualToneForTaskStatus(taskStatus);
                if (statusTone != PopupStatusVisualTone::None && _graphDynamicBrush)
                {
                    D2D1::ColorF stripeColor = StatusVisualColorForTone(appTheme, statusTone);
                    stripeColor.a            = 1.0f;
                    _graphDynamicBrush->SetColor(stripeColor);

                    const float stripeW        = DipsToPixels(3.0f, _dpi);
                    const D2D1_RECT_F stripeRc = D2D1::RectF(cardRect.left + DipsToPixels(1.0f, _dpi),
                                                             cardRect.top + DipsToPixels(1.0f, _dpi),
                                                             cardRect.left + DipsToPixels(1.0f, _dpi) + stripeW,
                                                             cardRect.bottom - DipsToPixels(1.0f, _dpi));
                    _target->FillRectangle(stripeRc, _graphDynamicBrush.get());
                }

                const float collapseBtnSize = DipsToPixels(18.0f, _dpi);
                const float collapseBtnGap  = DipsToPixels(6.0f, _dpi);

                const float headerTop    = isCollapsedTask ? cardRect.top + (taskCardH - lineH) * 0.5f : textY;
                const float headerBottom = headerTop + lineH;
                const float collapseTop  = headerTop + (lineH - collapseBtnSize) * 0.5f;
                const float collapseLeft = std::max(textX, contentRight - collapseBtnSize);

                PopupButton collapseBtn{};
                collapseBtn.bounds     = D2D1::RectF(collapseLeft, collapseTop, contentRight, collapseTop + collapseBtnSize);
                collapseBtn.hit.kind   = PopupHitTest::Kind::TaskToggleCollapse;
                collapseBtn.hit.taskId = task.taskId;
                _buttons.push_back(collapseBtn);

                float headerControlsLeft = collapseBtn.bounds.left;
                if (! task.started && ! task.overlapQueued && (task.waitingForOthers || task.waitingInQueue))
                {
                    const float queueButtonGap = DipsToPixels(3.0f, _dpi);
                    if (task.canMoveQueueDown)
                    {
                        PopupButton moveDownBtn{};
                        moveDownBtn.bounds     = D2D1::RectF(headerControlsLeft - queueButtonGap - collapseBtnSize,
                                                             collapseTop,
                                                             headerControlsLeft - queueButtonGap,
                                                             collapseTop + collapseBtnSize);
                        moveDownBtn.hit.kind   = PopupHitTest::Kind::TaskQueueMoveDown;
                        moveDownBtn.hit.taskId = task.taskId;
                        _buttons.push_back(moveDownBtn);
                        DrawButton(moveDownBtn, _buttonSmallFormat.get(), L"\u2193");
                        headerControlsLeft = moveDownBtn.bounds.left;
                    }
                    if (task.canMoveQueueUp)
                    {
                        PopupButton moveUpBtn{};
                        moveUpBtn.bounds     = D2D1::RectF(headerControlsLeft - queueButtonGap - collapseBtnSize,
                                                           collapseTop,
                                                           headerControlsLeft - queueButtonGap,
                                                           collapseTop + collapseBtnSize);
                        moveUpBtn.hit.kind   = PopupHitTest::Kind::TaskQueueMoveUp;
                        moveUpBtn.hit.taskId = task.taskId;
                        _buttons.push_back(moveUpBtn);
                        DrawButton(moveUpBtn, _buttonSmallFormat.get(), L"\u2191");
                        headerControlsLeft = moveUpBtn.bounds.left;
                    }
                }

                float headerRight = std::max(textX, headerControlsLeft - collapseBtnGap);
                float headerLeft  = textX;

                CaptionStatus statusIcon = CaptionStatus::None;
                {
                    if (StatusIsError(taskStatus))
                    {
                        statusIcon = CaptionStatus::Error;
                    }
                    else if (StatusIsWarning(taskStatus))
                    {
                        statusIcon = CaptionStatus::Warning;
                    }
                    else if (StatusIsOk(taskStatus))
                    {
                        statusIcon = CaptionStatus::Ok;
                    }
                }

                if (statusIcon != CaptionStatus::None && _target)
                {
                    const float iconSize = DipsToPixels(16.0f, _dpi);
                    const float iconGap  = DipsToPixels(6.0f, _dpi);

                    D2D1_RECT_F iconRc = D2D1::RectF(textX, headerTop, textX + iconSize, headerBottom);
                    iconRc.right       = std::min(iconRc.right, headerRight);

                    wchar_t fluentGlyph = 0;
                    wchar_t fallback    = 0;
                    ID2D1Brush* brush   = _textBrush.get();
                    switch (statusIcon)
                    {
                        case CaptionStatus::Ok:
                            fluentGlyph = FluentIcons::kCheckMark;
                            fallback    = FluentIcons::kFallbackCheckMark;
                            brush       = _statusOkBrush ? _statusOkBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                            break;
                        case CaptionStatus::Warning:
                            fluentGlyph = FluentIcons::kWarning;
                            fallback    = FluentIcons::kFallbackWarning;
                            brush       = _statusWarningBrush ? _statusWarningBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                            break;
                        case CaptionStatus::Error:
                            fluentGlyph = FluentIcons::kError;
                            fallback    = FluentIcons::kFallbackError;
                            brush       = _statusErrorBrush ? _statusErrorBrush.get() : (_textBrush ? _textBrush.get() : nullptr);
                            break;
                        case CaptionStatus::None:
                        default: break;
                    }

                    const bool useFluentFormat =
                        _statusIconFormat != nullptr && DirectWriteFormatHasGlyph(_dwriteFactory.get(), _statusIconFormat.get(), fluentGlyph);
                    const wchar_t glyph       = useFluentFormat ? fluentGlyph : fallback;
                    IDWriteTextFormat* format = useFluentFormat ? _statusIconFormat.get() : _statusIconFallbackFormat.get();

                    if (format && brush && glyph != 0 && iconRc.right > iconRc.left)
                    {
                        const wchar_t text[2]{glyph, 0};
                        _target->DrawTextW(text, 1u, format, iconRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        headerLeft = std::min(headerRight, iconRc.right + iconGap);
                    }
                }

                bool statusChipPainted      = false;
                const std::wstring chipText = isCollapsedTask && taskStatus == TaskStatusKind::Conflict && task.selection.itemCount > 0u
                                                  ? LoadStringResource(nullptr, IDS_FILEOPS_STATUS_NEEDS_ATTENTION)
                                                  : StatusChipTextForTask(task, taskStatus, nowTick);
                if (! chipText.empty() && statusTone != PopupStatusVisualTone::None && _smallFormat && _graphDynamicBrush)
                {
                    const float availableW = std::max(0.0f, headerRight - headerLeft);
                    const float chipH      = DipsToPixels(18.0f, _dpi);
                    const float chipHPad   = DipsToPixels(7.0f, _dpi);
                    const float chipGap    = DipsToPixels(6.0f, _dpi);
                    const float chipMaxW   = std::min(DipsToPixels(132.0f, _dpi), availableW * 0.42f);
                    if (chipMaxW >= DipsToPixels(44.0f, _dpi) && availableW > DipsToPixels(92.0f, _dpi))
                    {
                        const float textW = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), chipText, chipMaxW, chipH);
                        const float chipW = std::clamp(textW + chipHPad * 2.0f, DipsToPixels(44.0f, _dpi), chipMaxW);
                        if (chipW + chipGap < availableW)
                        {
                            D2D1::ColorF chipColor   = StatusVisualColorForTone(appTheme, statusTone);
                            chipColor.a              = appTheme.highContrast ? 0.28f : 0.14f;
                            const float chipTop      = headerTop + (lineH - chipH) * 0.5f;
                            const D2D1_RECT_F chipRc = D2D1::RectF(headerRight - chipW, chipTop, headerRight, chipTop + chipH);
                            const float chipRadius   = ClampCornerRadius(chipRc, DipsToPixels(5.0f, _dpi));

                            _graphDynamicBrush->SetColor(chipColor);
                            _target->FillRoundedRectangle(D2D1::RoundedRect(chipRc, chipRadius, chipRadius), _graphDynamicBrush.get());

                            chipColor.a = appTheme.highContrast ? 1.0f : 0.55f;
                            _graphDynamicBrush->SetColor(chipColor);
                            _target->DrawRoundedRectangle(D2D1::RoundedRect(chipRc, chipRadius, chipRadius), _graphDynamicBrush.get(), 1.0f);

                            const D2D1_RECT_F chipTextRc = D2D1::RectF(chipRc.left + chipHPad, chipRc.top, chipRc.right - chipHPad, chipRc.bottom);
                            _target->DrawTextW(chipText.data(),
                                               static_cast<UINT32>(chipText.size()),
                                               _smallFormat.get(),
                                               chipTextRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
#ifdef ENABLE_TESTS
                            auto& paintedChip         = _paintedStatusChipsForSelfTest[task.taskId];
                            paintedChip.painted       = true;
                            paintedChip.chrome        = chipRc;
                            paintedChip.textRect      = chipTextRc;
                            paintedChip.text          = chipText;
                            paintedChip.paintedFormat = _smallFormat;
#endif
                            headerRight       = std::max(headerLeft, chipRc.left - chipGap);
                            statusChipPainted = true;
                            if (taskStatus == TaskStatusKind::Stopping)
                                _hostedTooltipDescriptors.push_back({chipRc, chipText, std::format(L"FileOperations.Task.{}.Status", task.taskId), {}});
                        }
                    }
                }

                if (task.finished && statusChipPainted && task.selection.itemCount == 0u)
                    headerText = opText;
                if (task.finished && ! task.resultSummary.empty() && ! statusChipPainted && task.selection.itemCount == 0u)
                    headerText = FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, opText, chipText);
                if (taskStatus == TaskStatusKind::Stopping && ! statusChipPainted && task.selection.itemCount == 0u)
                    headerText = FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_OP_STATUS, opText, chipText);
                float headerTextRight                                           = headerRight;
                const WholeTaskProgressPresentation compactProgressPresentation = ResolveWholeTaskProgressPresentation(task);
                const bool showCompactProgress = (task.finished && TaskHasKnownCompactProgress(task)) ||
                                                 (compactProgressPresentation.showTransferProgress &&
                                                  (TaskHasKnownCompactProgress(task) || compactProgressPresentation.showTransferMarquee));
                if (isCollapsedTask && showCompactProgress && _progressBgBrush && _progressGlobalBrush && _smallFormat && _subTextBrush)
                {
                    const float availableHeaderW = std::max(0.0f, headerRight - headerLeft);
                    const float progressGap      = DipsToPixels(8.0f, _dpi);
                    const bool showPercent       = task.discoveryClosed && compactProgressPresentation.determinate;
                    // I35 (owner report 2026-10-13): a finished Verify On task that did not verify every byte names the
                    // share verified in place of its percent, in the verification color; the meter grows by the label.
                    const std::wstring verificationLabel = showPercent ? FinishedVerificationLabel(task) : std::wstring{};
                    const float labelW =
                        verificationLabel.empty() || ! _dwriteFactory
                            ? 0.0f
                            : std::ceil(MeasureTextWidth(
                                  _dwriteFactory.get(), _smallFormat.get(), verificationLabel, DipsToPixels(4096.0f, _dpi), DipsToPixels(128.0f, _dpi)));
                    const float percentW        = showPercent ? std::max(DipsToPixels(42.0f, _dpi), labelW) : 0.0f;
                    const float preferredMeterW = DipsToPixels(118.0f, _dpi) + (showPercent ? percentW - DipsToPixels(42.0f, _dpi) : 0.0f);
                    const float compactMeterW   = std::min(preferredMeterW, availableHeaderW * 0.40f);
                    const float percentGap      = showPercent ? DipsToPixels(6.0f, _dpi) : 0.0f;
                    const float barW            = compactMeterW - percentW - percentGap;
                    if (barW >= DipsToPixels(36.0f, _dpi) && compactMeterW + progressGap < availableHeaderW)
                    {
                        const float meterRight = headerRight;
                        const float meterLeft  = meterRight - compactMeterW;
                        headerTextRight        = std::max(headerLeft, meterLeft - progressGap);

                        const float fraction    = task.discoveryClosed ? ComputeFileOperationsTaskCompleteFractionForDisplay(task)
                                                                       : static_cast<float>(compactProgressPresentation.fraction);
                        const uint32_t percent  = static_cast<uint32_t>(std::lround(Clamp01(fraction) * 100.0f));
                        const float barH        = DipsToPixels(6.0f, _dpi);
                        const float barTop      = headerTop + (lineH - barH) * 0.5f;
                        const D2D1_RECT_F barRc = D2D1::RectF(meterLeft, barTop, meterLeft + barW, barTop + barH);
                        HostedProgressDescriptor progressDescriptor{
                            barRc, task.taskId, static_cast<double>(Clamp01(fraction)) * 100.0, ! compactProgressPresentation.determinate};
                        progressDescriptor.segmented             = task.verificationRequested && compactProgressPresentation.determinate;
                        progressDescriptor.primarySegmentValue   = TransferSegmentPercent(task);
                        progressDescriptor.secondarySegmentValue = VerificationSegmentPercent(task);
                        progressDescriptor.verificationSettled   = task.finished;
                        _hostedProgressDescriptors.push_back(progressDescriptor);

                        if (showPercent)
                        {
                            const std::wstring percentText =
                                verificationLabel.empty() ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_PERCENT, percent) : verificationLabel;
                            const D2D1_RECT_F percentRc = D2D1::RectF(barRc.right + percentGap, headerTop, meterRight, headerBottom);
                            ID2D1SolidColorBrush* const percentBrush =
                                ! verificationLabel.empty() && _progressVerifyBrush ? _progressVerifyBrush.get() : _subTextBrush.get();
                            _target->DrawTextW(percentText.data(),
                                               static_cast<UINT32>(percentText.size()),
                                               _smallFormat.get(),
                                               percentRc,
                                               percentBrush,
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
#ifdef ENABLE_TESTS
                            _paintedCompactMeterLabelsForSelfTest[task.taskId] = percentText;
#endif
                        }
                    }
                }

                if (_headerFormat)
                {
                    const D2D1_RECT_F headerRc = D2D1::RectF(headerLeft, headerTop, headerTextRight, headerBottom);
                    auto& information          = _cardInformationLayouts[row.taskIndex];
                    const float availableWidth = std::max(0.0f, headerRc.right - headerRc.left);
                    if (information.headerText != headerText || information.headerWidth != availableWidth ||
                        information.headerFormat.get() != _headerFormat.get())
                    {
                        information.headerText    = headerText;
                        information.headerWidth   = availableWidth;
                        information.headerFormat  = _headerFormat;
                        information.paintedHeader = TruncateTextMiddleToWidth(_dwriteFactory.get(),
                                                                              _headerFormat.get(),
                                                                              headerText,
                                                                              availableWidth,
                                                                              lineH,
                                                                              kEllipsisText,
                                                                              std::min(headerText.size(), opText.size() + 11u));
                    }
                    _target->DrawTextW(information.paintedHeader.data(),
                                       static_cast<UINT32>(information.paintedHeader.size()),
                                       _headerFormat.get(),
                                       headerRc,
                                       _textBrush.get(),
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    _hostedTooltipDescriptors.push_back({.bounds       = headerRc,
                                                         .text         = headerText,
                                                         .automationId = std::format(L"FileOperations.Task.{}.Subject", task.taskId),
                                                         .taskId       = task.taskId});
                    if (taskStatus == TaskStatusKind::Stopping && ! statusChipPainted)
                        _hostedTooltipDescriptors.push_back({headerRc, chipText, std::format(L"FileOperations.Task.{}.Status", task.taskId), {}});
#ifdef ENABLE_TESTS
                    _paintedStatusChipsForSelfTest[task.taskId].headerText        = headerText;
                    _paintedStatusChipsForSelfTest[task.taskId].paintedHeaderText = information.paintedHeader;
                    _paintedStatusChipsForSelfTest[task.taskId].headerRect        = headerRc;
#endif
                }

                if (isCollapsedTask)
                {
                    const float gapAfter = (rowIndex + 1u < rowCount) ? cardGap : 0.0f;
                    y += taskCardH + gapAfter;
                    continue;
                }

                textY = headerBottom;

                if (task.finished)
                {
                    const auto& summary = _terminalSummaries[row.taskIndex];
                    if (summary.layout)
                    {
                        textY += DipsToPixels(4.0f, _dpi);
                        const D2D1_RECT_F summaryBounds = D2D1::RectF(textX, textY, textX + summary.widthPx, textY + summary.heightPx);
                        _target->DrawTextLayout(D2D1::Point2F(textX, textY), summary.layout.get(), _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        _hostedTooltipDescriptors.push_back(
                            HostedTooltipDescriptor{.bounds       = summaryBounds,
                                                    .text         = summary.text,
                                                    .automationId = std::format(L"FileOperations.Task.{}.ResultSummary", task.taskId),
                                                    .taskId       = task.taskId});
                        textY = summaryBounds.bottom + DipsToPixels(6.0f, _dpi);
                    }
                    const auto& facts = _terminalFacts[row.taskIndex].prepared;
                    if (facts.layout)
                    {
                        textY += DipsToPixels(4.0f, _dpi);
                        const D2D1_RECT_F factsBounds = D2D1::RectF(textX, textY, textX + facts.widthPx, textY + facts.heightPx);
                        _target->DrawTextLayout(D2D1::Point2F(textX, textY), facts.layout.get(), _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        _hostedTooltipDescriptors.push_back(
                            HostedTooltipDescriptor{.bounds       = factsBounds,
                                                    .text         = facts.text,
                                                    .automationId = std::format(L"FileOperations.Task.{}.TerminalFacts", task.taskId),
                                                    .taskId       = task.taskId});
                        textY = factsBounds.bottom + DipsToPixels(6.0f, _dpi);
                    }
                }

                const auto addPathRegion = [&](const D2D1_RECT_F& bounds,
                                               std::wstring_view label,
                                               std::wstring_view path,
                                               std::wstring_view suffix,
                                               const D2D1_RECT_F& labelBounds,
                                               std::wstring_view paintedPath,
                                               std::wstring helpText = {})
                {
                    if (path.empty() || bounds.right <= bounds.left || bounds.bottom <= bounds.top)
                        return;
                    // Keep the exact value on the same rectangle as its abbreviated display.
                    // These bounded row slots are refreshed with the current snapshot, not cached by path.
                    _hostedTooltipDescriptors.push_back({bounds,
                                                         std::format(L"{} {}", label, path),
                                                         std::format(L"FileOperations.Task.{}.Path.{}", task.taskId, suffix),
                                                         std::wstring(path),
                                                         0u,
                                                         task.taskId,
                                                         std::move(helpText),
                                                         ! task.finished});
#ifdef ENABLE_TESTS
                    _hostedTooltipDescriptors.back().paintedLabelBounds = labelBounds;
                    _hostedTooltipDescriptors.back().paintedLabel       = label;
                    _hostedTooltipDescriptors.back().paintedPath        = paintedPath;
#else
                    static_cast<void>(labelBounds);
                    static_cast<void>(paintedPath);
#endif
                };

                if (task.finished)
                {
                    // I35 P5/R9.2/R9.3: the counters line names only nonzero counters, and the raw
                    // result HRESULT stays in Technical details, the log and the export.
                    if (! task.interruptedMoveNotice && TaskShowsCountersLine(task))
                    {
                        const std::wstring diagCounts = FormatTaskCountersText(task);
                        const D2D1_RECT_F countsRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(diagCounts.data(), static_cast<UINT32>(diagCounts.size()), _bodyFormat.get(), countsRc, _subTextBrush.get());
                        textY += lineH;
                    }

                    const bool deleting                 = task.operation == FILESYSTEM_DELETE;
                    const std::wstring sourceLabel      = LoadStringResource(nullptr, deleting ? IDS_FILEOPS_LABEL_DELETING : IDS_FILEOPS_LABEL_FROM);
                    const std::wstring destinationLabel = deleting ? std::wstring{} : LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO);
                    const float labelWDesired           = (std::max)({DipsToPixels(56.0f, _dpi),
                                                                      MeasurePathLabelWidth(deleting ? 2u : 0u, sourceLabel),
                                                                      deleting ? 0.0f : MeasurePathLabelWidth(1u, destinationLabel)});
                    const float labelGapDesired         = DipsToPixels(6.0f, _dpi);
                    const float labelW                  = std::min(labelWDesired, textMaxW);
                    const float labelGap                = (labelW < textMaxW) ? std::min(labelGapDesired, textMaxW - labelW) : 0.0f;
                    const float pathW                   = std::max(0.0f, textMaxW - labelW - labelGap);

                    if (task.operation == FILESYSTEM_DELETE)
                    {
                        const std::wstring& label = sourceLabel;
                        const D2D1_RECT_F labelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get());

                        const std::wstring path  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentSourcePath, pathW, lineH);
                        const D2D1_RECT_F pathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(
                            path.data(), static_cast<UINT32>(path.size()), _bodyFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        addPathRegion(pathRc, label, task.currentSourcePath, L"Source.0", labelRc, path);
                        textY += lineH;
                    }
                    else
                    {
                        const std::wstring& fromLabel = sourceLabel;
                        const D2D1_RECT_F fromLabelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(fromLabel.data(), static_cast<UINT32>(fromLabel.size()), _smallFormat.get(), fromLabelRc, _subTextBrush.get());

                        const std::wstring fromPath  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentSourcePath, pathW, lineH);
                        const D2D1_RECT_F fromPathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(fromPath.data(),
                                           static_cast<UINT32>(fromPath.size()),
                                           _bodyFormat.get(),
                                           fromPathRc,
                                           _textBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        addPathRegion(fromPathRc, fromLabel, task.currentSourcePath, L"Source.0", fromLabelRc, fromPath);
                        textY += lineH;

                        const std::wstring& toLabel = destinationLabel;
                        const D2D1_RECT_F toLabelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(toLabel.data(), static_cast<UINT32>(toLabel.size()), _smallFormat.get(), toLabelRc, _subTextBrush.get());

                        const std::wstring toPath =
                            TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentDestinationPath, pathW, lineH);
                        const D2D1_RECT_F toPathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(
                            toPath.data(), static_cast<UINT32>(toPath.size()), _bodyFormat.get(), toPathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        addPathRegion(toPathRc, toLabel, task.currentDestinationPath, L"Destination", toLabelRc, toPath);
                        textY += lineH;
                    }

                    if (const auto& note = _cardInformationLayouts[row.taskIndex].lastNote; note.rate)
                    {
                        const auto bounds = D2D1::RectF(textX, textY, contentRight, textY + FactPairHeight(note));
                        DrawFactPair(note, bounds, task.taskId, L"LastNote", {});
                        textY = bounds.bottom;
                    }

                    const float dismissButtonH         = DipsToPixels(_terminalActions[row.taskIndex].heightDip, _dpi);
                    const float dismissButtonBottomPad = DipsToPixels(8.0f, _dpi);
                    const float dismissButtonTop       = std::max(textY + DipsToPixels(4.0f, _dpi), cardRect.bottom - dismissButtonBottomPad - dismissButtonH);

#ifdef ENABLE_TESTS
                    _paintedTaskBodiesForSelfTest[task.taskId] = {
                        cardRect, textY, D2D1::RectF(textX, dismissButtonTop, contentRight, dismissButtonTop + dismissButtonH)};
#endif

                    const float progressBarH         = DipsToPixels(8.0f, _dpi);
                    const float progressBarBottomPad = DipsToPixels(6.0f, _dpi);
                    const float progressBarTop       = std::max(textY + DipsToPixels(2.0f, _dpi), dismissButtonTop - progressBarBottomPad - progressBarH);
                    const D2D1_RECT_F progressRc     = D2D1::RectF(textX, progressBarTop, contentRight, progressBarTop + progressBarH);

                    HostedProgressDescriptor progressDescriptor{
                        progressRc, task.taskId, static_cast<double>(Clamp01(ComputeFileOperationsTaskCompleteFractionForDisplay(task))) * 100.0, false};
                    progressDescriptor.segmented             = task.verificationRequested;
                    progressDescriptor.primarySegmentValue   = TransferSegmentPercent(task);
                    progressDescriptor.secondarySegmentValue = VerificationSegmentPercent(task);
                    progressDescriptor.verificationSettled   = task.finished;
                    _hostedProgressDescriptors.push_back(progressDescriptor);

                    const auto& actions = _terminalActions[row.taskIndex];
                    constexpr std::array<PopupHitTest::Kind, 3> kinds{
                        {PopupHitTest::Kind::TaskDismiss, PopupHitTest::Kind::TaskIssues, PopupHitTest::Kind::TaskCompletedMore}};
                    for (size_t index = 0u; index < kinds.size(); ++index)
                    {
                        if (actions.labels[index].empty())
                            continue;
                        const auto& local = actions.bounds[index];
                        PopupButton action{};
                        action.bounds = D2D1::RectF(textX + DipsToPixels(local.left, _dpi),
                                                    dismissButtonTop + DipsToPixels(local.top, _dpi),
                                                    textX + DipsToPixels(local.right, _dpi),
                                                    dismissButtonTop + DipsToPixels(local.bottom, _dpi));
                        action.hit    = {kinds[index], task.taskId, index == 2u ? 2u : 0u};
                        _buttons.push_back(action);
                        // Retained DxUi buttons own the measured, wrapped Body caption.
                        // Paint only the fallback chrome here, without a duplicate small label.
                        if (index == 2u)
                            DrawMenuButton(action, nullptr, {});
                        else
                            DrawButton(action, nullptr, {});
                    }

                    const float gapAfter = (rowIndex + 1u < rowCount) ? cardGap : 0.0f;
                    y += taskCardH + gapAfter;
                    continue;
                }

                const AppTheme& theme = folderWindow->GetTheme();

                const RateHistory* history = nullptr;
                const auto historyIt       = _rates.find(task.taskId);
                if (historyIt != _rates.end())
                {
                    history = &historyIt->second;
                }
                const bool graphBandsActive = RateHistoryGraphBandsAreActive(history, theme);

                const auto& information     = _cardInformationLayouts[row.taskIndex];
                const bool priorityTransfer = ! task.conflict.active && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE);
                float primaryProgressTop    = 0.0f;
                D2D1_RECT_F primaryEtaRect{};
                const bool skippedCountingSummary = TaskUsesSkippedCountingWorkSummary(task);
                if (priorityTransfer && information.progress.rate && ! skippedCountingSummary)
                {
                    textY += DipsToPixels(4.0f, _dpi);
                    const auto bounds = D2D1::RectF(textX, textY, contentRight, textY + FactPairHeight(information.progress));
                    DrawFactPair(information.progress, bounds, task.taskId, L"Progress.Amount", L"Progress.Percent");
#ifdef ENABLE_TESTS
                    _paintedLiveProgressForSelfTest[task.taskId].transferText           = information.progress.rateText;
                    _paintedLiveProgressForSelfTest[task.taskId].transferRect           = bounds;
                    _paintedLiveProgressForSelfTest[task.taskId].transferCompletedBytes = task.completedBytes;
#endif
                    textY = bounds.bottom;
                }
                if (skippedCountingSummary && compactProgressPresentation.showWorkFacts)
                {
                    const std::wstring summaryText = FormatDiscoveryProgressText(task);
                    const float summaryHeight      = MeasureWrappedTextHeight(_dwriteFactory.get(), _bodyFormat.get(), summaryText, textMaxW, lineH);
                    const D2D1_RECT_F summaryRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + summaryHeight);
                    static_cast<void>(
                        DrawWrappedText(_dwriteFactory.get(), _target.get(), _bodyFormat.get(), _subTextBrush.get(), summaryText, summaryRc, lineH));
                    _hostedTooltipDescriptors.push_back({.bounds       = summaryRc,
                                                         .text         = summaryText,
                                                         .automationId = std::format(L"FileOperations.Task.{}.WorkSummary", task.taskId),
                                                         .taskId       = task.taskId,
                                                         .liveText     = true});
#ifdef ENABLE_TESTS
                    _paintedLiveProgressForSelfTest[task.taskId].transferText  = summaryText;
                    _paintedLiveProgressForSelfTest[task.taskId].transferRect  = summaryRc;
                    _paintedLiveProgressForSelfTest[task.taskId].transferColor = _subTextBrush->GetColor();
                    _paintedLiveProgressForSelfTest[task.taskId].transferCompletedBytes =
                        task.operation == FILESYSTEM_DELETE ? task.completedWorkBytes : task.completedBytes;
#endif
                    textY += summaryHeight;
                }
                const bool discoveryInProgress = compactProgressPresentation.showDiscoveryActivity;
                if (discoveryInProgress && ! priorityTransfer)
                {
                    const std::wstring discoveryText = FormatDiscoveredSoFarText(task);
                    const float indicatorW           = DipsToPixels(24.0f, _dpi);
                    const float indicatorGap         = DipsToPixels(6.0f, _dpi);
                    const D2D1_RECT_F discoveryRc    = D2D1::RectF(textX, textY, std::max(textX, contentRight - indicatorW - indicatorGap), textY + lineH);
                    _target->DrawTextW(discoveryText.data(),
                                       static_cast<UINT32>(discoveryText.size()),
                                       _bodyFormat.get(),
                                       discoveryRc,
                                       _subTextBrush.get(),
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    DrawDiscoveryActivityIndicator(D2D1::RectF(contentRight - indicatorW, textY, contentRight, textY + lineH), nowTick, reducedMotion);
#ifdef ENABLE_TESTS
                    _paintedLiveProgressForSelfTest[task.taskId].discoveryText  = discoveryText;
                    _paintedLiveProgressForSelfTest[task.taskId].discoveryRect  = discoveryRc;
                    _paintedLiveProgressForSelfTest[task.taskId].discoveryColor = _subTextBrush->GetColor();
#endif
                    textY += lineH;

                    // D2-A08: exact discovered/processed counters instead of a provisional fraction.
                    const std::wstring transferText = FormatDiscoveryProgressText(task);
                    const D2D1_RECT_F transferRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(transferText.data(),
                                       static_cast<UINT32>(transferText.size()),
                                       _bodyFormat.get(),
                                       transferRc,
                                       _subTextBrush.get(),
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
#ifdef ENABLE_TESTS
                    _paintedLiveProgressForSelfTest[task.taskId].transferText           = transferText;
                    _paintedLiveProgressForSelfTest[task.taskId].transferRect           = transferRc;
                    _paintedLiveProgressForSelfTest[task.taskId].transferCompletedBytes = task.completedBytes;
#endif
                    textY += lineH;
                }

                if (! task.conflict.active && task.operation == FILESYSTEM_DELETE && (! skippedCountingSummary || task.completedWorkCountsKnown))
                {
                    const bool showPreparing = TaskShowsPreparingStatus(task);

                    if (showPreparing)
                    {
                        const ULONGLONG opStartTick = task.operationStartTick;
                        const uint64_t elapsedSec =
                            (opStartTick > 0 && nowTick >= opStartTick) ? static_cast<uint64_t>((nowTick - opStartTick) / 1000ull) : 0ull;
                        const std::wstring prepText = elapsedSec > 0
                                                          ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_PREPARING_TIME, FormatDurationHms(elapsedSec))
                                                          : LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
                        const D2D1_RECT_F prepRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(prepText.data(), static_cast<UINT32>(prepText.size()), _bodyFormat.get(), prepRc, _subTextBrush.get());
                        textY += lineH;
                    }
                    else
                    {
                        const double itemsPerSec     = history ? history->displayedItemsPerSec : 0.0;
                        const std::wstring speedText = FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_ITEMS, itemsPerSec);
                        const D2D1_RECT_F speedRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(speedText.data(), static_cast<UINT32>(speedText.size()), _bodyFormat.get(), speedRc, _subTextBrush.get());
                        _hostedTooltipDescriptors.push_back({.bounds       = speedRc,
                                                             .text         = speedText,
                                                             .automationId = std::format(L"FileOperations.Task.{}.DeletionRate", task.taskId),
                                                             .liveText     = true});
                        textY += lineH;

                        const bool showSizeProgress =
                            ! skippedCountingSummary && task.discoveryClosed && task.discoveredTotalBytes > 0 && task.completedBytes > 0;
                        if (showSizeProgress)
                        {
                            const std::wstring sizeProgressText = FormatEmbeddedStringResource(
                                nullptr, IDS_FMT_FILEOPS_SIZE_PROGRESS, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.discoveredTotalBytes));
                            const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                sizeProgressText.data(), static_cast<UINT32>(sizeProgressText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
                            textY += lineH;
                        }
                        else if (! skippedCountingSummary && task.discoveryClosed && task.totalItems > 0)
                        {
                            const std::wstring itemsProgressText = FormatStringResource(nullptr, IDS_FMT_FILEOP_ITEMS, task.completedItems, task.totalItems);
                            const D2D1_RECT_F itemsRc            = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                itemsProgressText.data(), static_cast<UINT32>(itemsProgressText.size()), _bodyFormat.get(), itemsRc, _subTextBrush.get());
                            textY += lineH;
                        }
                        else if (! skippedCountingSummary)
                        {
                            const std::wstring itemsProgressText = FormatStringResource(nullptr, IDS_FMT_FILEOP_ITEMS_UNKNOWN_TOTAL, task.completedItems);
                            const D2D1_RECT_F itemsRc            = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                itemsProgressText.data(), static_cast<UINT32>(itemsProgressText.size()), _bodyFormat.get(), itemsRc, _subTextBrush.get());
                            textY += lineH;
                        }
                    }
                }
                else if (! task.conflict.active && ! skippedCountingSummary && ! priorityTransfer)
                {
                    if (task.operation == FILESYSTEM_COPY && task.discoveryClosed)
                    {
                        const std::wstring foundText =
                            FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOUND_FILE_COUNTS, task.discoveredFileCount, task.discoveredDirectoryCount);
                        const D2D1_RECT_F foundRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(foundText.data(),
                                           static_cast<UINT32>(foundText.size()),
                                           _bodyFormat.get(),
                                           foundRc,
                                           _subTextBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
#ifdef ENABLE_TESTS
                        _paintedLiveProgressForSelfTest[task.taskId].discoveryText  = foundText;
                        _paintedLiveProgressForSelfTest[task.taskId].discoveryRect  = foundRc;
                        _paintedLiveProgressForSelfTest[task.taskId].discoveryColor = _subTextBrush->GetColor();
#endif
                        textY += lineH;
                        if (task.copiedFileCountKnown)
                        {
                            const std::wstring copiedText =
                                FormatStringResource(nullptr, IDS_FMT_FILEOPS_COPIED_FILES, task.copiedFileCount, task.discoveredFileCount);
                            const D2D1_RECT_F copiedRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(copiedText.data(),
                                               static_cast<UINT32>(copiedText.size()),
                                               _bodyFormat.get(),
                                               copiedRc,
                                               _subTextBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
                        }
                    }
                    // Show size progress (transferred / total) if we have data
                    if (task.discoveryClosed && (task.operation == FILESYSTEM_COPY ? task.discoveredBytesKnown : task.totalBytes > 0u))
                    {
                        const std::wstring sizeProgressText = FormatEmbeddedStringResource(
                            nullptr, IDS_FMT_FILEOPS_SIZE_PROGRESS, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.totalBytes));
                        const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(
                            sizeProgressText.data(), static_cast<UINT32>(sizeProgressText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
#ifdef ENABLE_TESTS
                        _paintedLiveProgressForSelfTest[task.taskId].transferText           = sizeProgressText;
                        _paintedLiveProgressForSelfTest[task.taskId].transferRect           = sizeRc;
                        _paintedLiveProgressForSelfTest[task.taskId].transferCompletedBytes = task.completedBytes;
#endif
                        textY += lineH;
                    }
                    else if (task.operation == FILESYSTEM_COPY && task.discoveryClosed)
                    {
                        const auto writtenText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_WRITTEN_BYTES, FormatBytesCompact(task.completedBytes));
                        const auto writtenRc   = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(writtenText.data(), static_cast<UINT32>(writtenText.size()), _bodyFormat.get(), writtenRc, _subTextBrush.get());
                        textY += lineH;
                    }
                }

                if (priorityTransfer)
                {
                    if (information.counts.rate || information.counts.eta)
                    {
                        textY += DipsToPixels(4.0f, _dpi);
                        const float countsLeft = textX + (discoveryInProgress ? DipsToPixels(30.0f, _dpi) : 0.0f);
                        const auto bounds      = D2D1::RectF(countsLeft, textY, contentRight, textY + FactPairHeight(information.counts));
                        DrawFactPair(information.counts, bounds, task.taskId, L"Progress.Counts", L"Throughput.Eta");
                        const float etaTop = information.counts.stacked ? textY + information.counts.rateHeight + DipsToPixels(2.0f, _dpi) : textY;
                        primaryEtaRect     = D2D1::RectF(contentRight - (information.counts.stacked ? information.counts.width : information.counts.etaWidth),
                                                         etaTop,
                                                         contentRight,
                                                         etaTop + information.counts.etaHeight);
#ifdef ENABLE_TESTS
                        const auto countsRect = D2D1::RectF(countsLeft,
                                                            textY,
                                                            countsLeft + (information.counts.stacked ? information.counts.width : information.counts.rateWidth),
                                                            textY + information.counts.rateHeight);
                        auto& painted         = _paintedLiveProgressForSelfTest[task.taskId];
                        painted.discoveryText = information.counts.rateText;
                        // Preserve numeric paint inputs independently of the formatted label so live
                        // tests can compare it with their own known totals without a snapshot race.
                        painted.recursiveCounts = {task.discoveredFileCount,
                                                   task.discoveredDirectoryCount,
                                                   task.copiedFileCount,
                                                   task.copiedFolderCount,
                                                   task.copiedFileCountKnown,
                                                   task.copiedFolderCountKnown,
                                                   task.discoveryClosed};
                        painted.discoveryRect   = countsRect;
                        painted.discoveryColor  = _subTextBrush->GetColor();
#endif
                        if (discoveryInProgress)
                            DrawDiscoveryActivityIndicator(D2D1::RectF(textX, textY, textX + DipsToPixels(24.0f, _dpi), textY + lineH), nowTick, reducedMotion);
                        textY = bounds.bottom;
                    }
                    primaryProgressTop = textY + DipsToPixels(6.0f, _dpi);
                    textY += compactProgressPresentation.showTransferProgress ? DipsToPixels(26.0f, _dpi) : DipsToPixels(8.0f, _dpi);
                }
                // A decision replaces ordinary transfer details. Keeping both here
                // exposes stale paths and a moving mini-progress above the conflict.
                if (! task.conflict.active)
                {
                    const bool deleting                 = task.operation == FILESYSTEM_DELETE;
                    const UINT sourceLabelId            = static_cast<UINT>(deleting                  ? IDS_FILEOPS_LABEL_DELETING
                                                                            : task.verificationActive ? IDS_FILEOPS_STATUS_VERIFYING
                                                                                                      : IDS_FILEOPS_LABEL_FROM);
                    const std::wstring sourceLabel      = LoadStringResource(nullptr, sourceLabelId);
                    const std::wstring destinationLabel = deleting ? std::wstring{} : LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO);
                    float phaseLabelWidth               = 0.0f;
                    if (task.operation == FILESYSTEM_COPY)
                    {
                        phaseLabelWidth = (std::max)(MeasurePathLabelWidth(4u, LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FINALIZING)),
                                                     MeasurePathLabelWidth(5u, LoadStringResource(nullptr, IDS_FILEOPS_LABEL_PUBLISHING)));
                        phaseLabelWidth = (std::max)(phaseLabelWidth, MeasurePathLabelWidth(3u, LoadStringResource(nullptr, IDS_FILEOPS_STATUS_VERIFYING)));
                    }
                    const float labelWDesired   = (std::max)({DipsToPixels(56.0f, _dpi),
                                                              phaseLabelWidth,
                                                              MeasurePathLabelWidth(deleting                  ? 2u
                                                                                    : task.verificationActive ? 3u
                                                                                                              : 0u,
                                                                                    sourceLabel),
                                                              deleting ? 0.0f : MeasurePathLabelWidth(1u, destinationLabel)});
                    const float labelGapDesired = DipsToPixels(6.0f, _dpi);
                    const float labelW          = std::min(labelWDesired, textMaxW);
                    const float labelGap        = (labelW < textMaxW) ? std::min(labelGapDesired, textMaxW - labelW) : 0.0f;

                    if (task.operation == FILESYSTEM_DELETE)
                    {
                        const std::wstring& label = sourceLabel;
                        const D2D1_RECT_F labelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get());

                        const float pathW        = std::max(0.0f, textMaxW - labelW - labelGap);
                        const std::wstring path  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentSourcePath, pathW, lineH);
                        const D2D1_RECT_F pathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(
                            path.data(), static_cast<UINT32>(path.size()), _bodyFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        addPathRegion(pathRc, label, task.currentSourcePath, L"Source.0", labelRc, path);
                        textY += lineH;
                    }
                    else
                    {
                        const std::wstring& fromLabel = sourceLabel;
                        const float miniBarGap        = DipsToPixels(8.0f, _dpi);
                        const float miniBarWDesired   = DipsToPixels(92.0f, _dpi);
                        const float miniBarH          = DipsToPixels(6.0f, _dpi);

                        const float pathLeft  = textX + labelW + labelGap;
                        const float rightEdge = textX + textMaxW;

                        const bool showInFlightFiles = compactProgressPresentation.showTransferProgress && compactProgressPresentation.transferActive &&
                                                       (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE);

                        std::array<size_t, TaskSnapshot::kMaxInFlightFiles> activeInFlightIndices{};
                        size_t activeInFlightCount = 0;
                        if (showInFlightFiles)
                        {
                            for (size_t j = 0; j < task.inFlightFileCount && activeInFlightCount < activeInFlightIndices.size(); ++j)
                            {
                                const auto& entry = task.inFlightFiles[j];
                                if (FileProgressIsVisible(entry, nowTick))
                                {
                                    activeInFlightIndices[activeInFlightCount] = j;
                                    ++activeInFlightCount;
                                }
                            }
                        }

                        const size_t inFlightCount = showInFlightFiles ? std::max<size_t>(1u, activeInFlightCount) : 1u;
                        // Q7: the task publishes one stable source region, on the first row. The other
                        // active streams are painted rows and are named in that region's help text.
                        std::wstring otherStreams;
                        for (size_t i = 1u; showInFlightFiles && i < activeInFlightCount; ++i)
                        {
                            if (! otherStreams.empty())
                                otherStreams += L'\n';
                            otherStreams += task.inFlightFiles[activeInFlightIndices[i]].sourcePath;
                        }

                        for (size_t i = 0; i < inFlightCount; ++i)
                        {
                            std::wstring_view sourcePathText;
                            uint64_t fileTotalBytes                                 = 0;
                            uint64_t fileCompletedBytes                             = 0;
                            const TaskSnapshot::InFlightFileSnapshot* inFlightEntry = nullptr;

                            const bool hasActiveInFlight = showInFlightFiles && activeInFlightCount > 0;
                            const bool useInFlightEntry  = hasActiveInFlight && i < activeInFlightCount;

                            if (useInFlightEntry)
                            {
                                const auto& entry  = task.inFlightFiles[activeInFlightIndices[i]];
                                inFlightEntry      = &entry;
                                sourcePathText     = entry.sourcePath;
                                fileTotalBytes     = entry.totalBytes;
                                fileCompletedBytes = entry.completedBytes;
                            }
                            else
                            {
                                sourcePathText     = task.currentSourcePath;
                                fileTotalBytes     = task.verificationActive ? task.verificationItemTotalBytes : task.itemTotalBytes;
                                fileCompletedBytes = task.verificationActive ? task.verificationItemCompletedBytes : task.itemCompletedBytes;
                            }
                            std::wstring rowLabel = fromLabel;
                            if (inFlightEntry && inFlightEntry->explicitLifecycle && ! inFlightEntry->terminal)
                            {
                                if (inFlightEntry->phase == FileSystemFileWorkPhase::Finalizing)
                                    rowLabel = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FINALIZING);
                                else if (inFlightEntry->phase == FileSystemFileWorkPhase::Publishing)
                                    rowLabel = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_PUBLISHING);
                                else if (inFlightEntry->phase == FileSystemFileWorkPhase::Verifying)
                                    rowLabel = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_VERIFYING);
                            }
                            if (i == 0u || rowLabel != fromLabel)
                            {
                                const D2D1_RECT_F fromRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                                _target->DrawTextW(rowLabel.data(),
                                                   static_cast<UINT32>(rowLabel.size()),
                                                   _smallFormat.get(),
                                                   fromRc,
                                                   _subTextBrush.get(),
                                                   D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            }

                            const float availableW     = std::max(0.0f, rightEdge - pathLeft);
                            const float miniBarWMin    = DipsToPixels(40.0f, _dpi);
                            const float minTextW       = DipsToPixels(48.0f, _dpi);
                            float miniBarW             = std::min(miniBarWDesired, availableW);
                            const float maxBarWithText = std::max(0.0f, availableW - miniBarGap - minTextW);
                            if (maxBarWithText > 0.0f)
                            {
                                miniBarW = std::clamp(miniBarW, std::min(miniBarWMin, maxBarWithText), maxBarWithText);
                            }

                            // If nothing is actively copying (e.g., end-of-file or finalization), avoid showing a "stuck at 100%" mini bar.
                            if (! showInFlightFiles || (! useInFlightEntry && fileTotalBytes > 0 && fileCompletedBytes >= fileTotalBytes))
                            {
                                miniBarW = 0.0f;
                            }

                            const float barRight = rightEdge;
                            const float barLeft  = barRight - miniBarW;

                            const float pathRight = (miniBarW > 0.0f) ? std::max(pathLeft, barLeft - miniBarGap) : rightEdge;
                            const float pathW     = std::max(0.0f, pathRight - pathLeft);

                            const std::wstring fromPath  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), sourcePathText, pathW, lineH);
                            const D2D1_RECT_F fromPathRc = D2D1::RectF(pathLeft, textY, pathLeft + pathW, textY + lineH);
                            _target->DrawTextW(fromPath.data(),
                                               static_cast<UINT32>(fromPath.size()),
                                               _bodyFormat.get(),
                                               fromPathRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);

                            if (i == 0u)
                                addPathRegion(
                                    fromPathRc,
                                    fromLabel,
                                    sourcePathText,
                                    L"Source.0",
                                    D2D1::RectF(textX, textY, textX + labelW, textY + lineH),
                                    fromPath,
                                    otherStreams.empty()
                                        ? std::wstring{}
                                        : FormatStringResource(
                                              nullptr, IDS_FMT_FILEOPS_PATH_OTHER_STREAMS, fromLabel, sourcePathText, activeInFlightCount - 1u, otherStreams));

                            if (miniBarW > 0.0f && _progressBgBrush && _progressGlobalBrush && _progressItemBrush && _progressVerifyBrush)
                            {
                                const float barTop          = textY + (lineH - miniBarH) * 0.5f;
                                const D2D1_RECT_F miniBarRc = D2D1::RectF(barLeft, barTop, barRight, barTop + miniBarH);

                                const float radiusTrack = ClampCornerRadius(miniBarRc, DipsToPixels(2.0f, _dpi));
                                _target->FillRoundedRectangle(D2D1::RoundedRect(miniBarRc, radiusTrack, radiusTrack), _progressBgBrush.get());

                                const bool hasTotal = fileTotalBytes > 0 || (inFlightEntry && inFlightEntry->explicitLifecycle && inFlightEntry->sizeKnown);
                                const float frac =
                                    fileTotalBytes > 0 && fileCompletedBytes <= fileTotalBytes
                                        ? Clamp01(static_cast<float>(static_cast<double>(fileCompletedBytes) / static_cast<double>(fileTotalBytes)))
                                    : hasTotal ? 1.0f
                                               : 0.0f;

                                const bool verificationMiniBar = task.verificationActive && ! useInFlightEntry;
                                if (! verificationMiniBar && graphBandsActive && inFlightEntry)
                                {
                                    const D2D1_COLOR_F rowColor = RainbowProgressColor(theme, sourcePathText, history, inFlightEntry);
                                    _progressItemBrush->SetColor(rowColor);
#ifdef ENABLE_TESTS
                                    const float assignedHue      = AssignedRateStreamHue(history, inFlightEntry);
                                    const D2D1_COLOR_F bandColor = DxUi::ThroughputGraphColorFromHue(assignedHue, theme.dark);
                                    if (assignedHue >= 0.0f && rowColor.r == bandColor.r && rowColor.g == bandColor.g && rowColor.b == bandColor.b &&
                                        rowColor.a == bandColor.a)
                                    {
                                        ++_graphRowColorMatchCountsForSelfTest[task.taskId];
                                    }
                                    else
                                    {
                                        ++_graphRowColorMismatchCountsForSelfTest[task.taskId];
                                    }
#endif
                                }
                                else
                                {
                                    _progressItemBrush->SetColor(_progressItemBaseColor);
                                }

                                const D2D1_RECT_F fill =
                                    hasTotal ? D2D1::RectF(
                                                   miniBarRc.left, miniBarRc.top, miniBarRc.left + (miniBarRc.right - miniBarRc.left) * frac, miniBarRc.bottom)
                                             : ComputeIndeterminateBarFill(miniBarRc, nowTick, reducedMotion);
                                const float radiusFill                = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                                ID2D1SolidColorBrush* const fillBrush = verificationMiniBar ? _progressVerifyBrush.get() : _progressItemBrush.get();
                                _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radiusFill, radiusFill), fillBrush);
#ifdef ENABLE_TESTS
                                auto& paintedProgress = _paintedLiveProgressForSelfTest[task.taskId];
                                if (paintedProgress.fileCount < paintedProgress.files.size())
                                {
                                    auto& paintedFile            = paintedProgress.files[paintedProgress.fileCount++];
                                    paintedFile.sourcePath       = sourcePathText;
                                    paintedFile.progressStreamId = inFlightEntry ? inFlightEntry->progressStreamId : 0u;
                                    paintedFile.totalBytes       = fileTotalBytes;
                                    paintedFile.completedBytes   = fileCompletedBytes;
                                    paintedFile.track            = miniBarRc;
                                    paintedFile.fill             = fill;
                                    paintedFile.trackColor       = _progressBgBrush->GetColor();
                                    paintedFile.fillColor        = fillBrush->GetColor();
                                    paintedFile.determinate      = hasTotal;
                                }
#endif
                                if (verificationMiniBar && fill.right > fill.left)
                                {
                                    _target->PushAxisAlignedClip(fill, D2D1_ANTIALIAS_MODE_ALIASED);
                                    const float hatchStep = std::max(3.0f, DipsToPixels(5.0f, _dpi));
                                    const float hatchRise = fill.bottom - fill.top;
                                    for (float x = fill.left - hatchRise; x < fill.right; x += hatchStep)
                                    {
                                        _target->DrawLine(
                                            D2D1::Point2F(x, fill.bottom), D2D1::Point2F(x + hatchRise, fill.top), _progressGlobalBrush.get(), 1.0f);
                                    }
                                    _target->PopAxisAlignedClip();
                                }
                            }

                            textY += lineH;
                        }

                        textY += fromToGapY;

                        const std::wstring& toLabel = destinationLabel;
                        const D2D1_RECT_F toRc      = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(toLabel.data(), static_cast<UINT32>(toLabel.size()), _smallFormat.get(), toRc, _subTextBrush.get());

                        const std::wstring destText = task.destinationFolder.wstring();

                        const float toPathLeft = textX + labelW + labelGap;
                        const float toRight    = textX + textMaxW;

                        const float toPathW        = std::max(0.0f, toRight - toPathLeft);
                        const std::wstring toPath  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), destText, toPathW, lineH);
                        const D2D1_RECT_F toPathRc = D2D1::RectF(toPathLeft, textY, toPathLeft + toPathW, textY + lineH);
                        _target->DrawTextW(
                            toPath.data(), static_cast<UINT32>(toPath.size()), _bodyFormat.get(), toPathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        addPathRegion(toPathRc, toLabel, destText, L"Destination", toRc, toPath);

                        textY += lineH;
                    }
                }

                const float barInsetX = DipsToPixels(10.0f, _dpi);
                const float barW      = std::max(0.0f, cardRect.right - cardRect.left - barInsetX * 2.0f);
                const float barX      = cardRect.left + barInsetX;

                const float barHItem = DipsToPixels(10.0f, _dpi);

                const bool hasConflictPrompt = task.conflict.active;
                const ConflictPromptContent conflictPromptContent =
                    hasConflictPrompt ? BuildConflictPromptContent(task, ConflictInspectionResultText(task)) : ConflictPromptContent{};

                const float barsHeight    = barHItem;
                const float bottomPadding = DipsToPixels(10.0f, _dpi);
                const float buttonGapY    = DipsToPixels(8.0f, _dpi);
                const float buttonH       = DipsToPixels(24.0f, _dpi);

                const float conflictRowGapY = DipsToPixels(6.0f, _dpi);
                const float conflictButtonsHeight =
                    hasConflictPrompt && task.conflict.buttonsPublishable ? DipsToPixels(_conflictActionMeasurements[row.taskIndex].heightDip, _dpi) : 0.0f;
                const float conflictApplyLineHeight = hasConflictPrompt && task.conflict.buttonsPublishable && task.conflict.applyToAllEligible
                                                          ? DipsToPixels(_conflictActionMeasurements[row.taskIndex].scopeHeightDip, _dpi) + conflictRowGapY
                                                          : 0.0f;
                const float buttonsHeight           = hasConflictPrompt ? (conflictButtonsHeight + conflictApplyLineHeight) : buttonH;

                const float noteHeight =
                    ! hasConflictPrompt && information.concurrency.rate ? FactPairHeight(information.concurrency) + DipsToPixels(6.0f, _dpi) : 0.0f;
                const float buttonRowBottom = cardRect.bottom - bottomPadding - noteHeight;
                const float buttonRowTop    = buttonRowBottom - buttonsHeight;
#ifdef ENABLE_TESTS
                if (! hasConflictPrompt)
                    _paintedTaskBodiesForSelfTest[task.taskId] = {
                        cardRect, textY, D2D1::RectF(cardRect.left + padX, buttonRowTop, contentRight, buttonRowBottom)};
#endif

                const float barsBottom = buttonRowTop - buttonGapY;
                const float barsTop    = priorityTransfer ? primaryProgressTop : barsBottom - barsHeight;

                const auto drawConflictPromptInfo = [&](const D2D1_RECT_F& rc) noexcept
                {
                    if (! _bodyFormat || ! _smallFormat || ! _dwriteFactory || ! _target)
                    {
                        return;
                    }

                    _publishedInspectionContexts.push_back({task.taskId, task.conflict.generation, task.conflict.inspection});
                    D2D1_RECT_F contentBounds = rc;
                    uint64_t comparisonTaskId = 0u;
                    float yPrompt             = rc.top;
                    float maxW                = std::max(0.0f, contentBounds.right - contentBounds.left);
                    const float rowGap        = DipsToPixels(4.0f, _dpi);
                    bool hasPriorContent      = false;

                    const auto addSemanticRegion = [&](const D2D1_RECT_F& bounds, std::wstring_view text, std::wstring_view suffix)
                    {
                        if (text.empty())
                        {
                            return;
                        }
                        _hostedTooltipDescriptors.push_back(HostedTooltipDescriptor{
                            .bounds           = bounds,
                            .text             = std::wstring(text),
                            .automationId     = std::format(L"FileOperations.Task.{0}.Conflict.{1}", task.taskId, suffix),
                            .comparisonTaskId = comparisonTaskId,
                        });
                    };

                    const auto drawTextRegion = [&](std::wstring_view text, IDWriteTextFormat* format, ID2D1Brush* brush, std::wstring_view suffix) noexcept
                    {
                        if (text.empty())
                        {
                            return;
                        }
                        if (hasPriorContent)
                        {
                            yPrompt += rowGap;
                        }
                        const float textHeight = MeasureWrappedTextHeight(_dwriteFactory.get(), format, text, maxW, lineH);
                        const D2D1_RECT_F textRc =
                            D2D1::RectF(contentBounds.left, yPrompt, contentBounds.right, std::min(contentBounds.bottom, yPrompt + textHeight));
                        static_cast<void>(DrawWrappedText(_dwriteFactory.get(), _target.get(), format, brush, text, textRc, lineH));
                        addSemanticRegion(textRc, text, suffix);
                        yPrompt += textHeight;
                        hasPriorContent = true;
                    };

                    const auto drawContext = [&](std::wstring_view label,
                                                 std::wstring_view path,
                                                 std::wstring_view name,
                                                 std::wstring_view parent,
                                                 std::wstring_view metadata,
                                                 std::wstring_view suffix,
                                                 PopupHitTest::Kind inspectionKind) noexcept
                    {
                        if (label.empty() || path.empty())
                        {
                            return;
                        }
                        if (hasPriorContent)
                        {
                            yPrompt += rowGap;
                        }
                        const float contextTop    = yPrompt;
                        const D2D1_RECT_F labelRc = D2D1::RectF(contentBounds.left, yPrompt, contentBounds.right, yPrompt + lineH);
                        _target->DrawTextW(
                            label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        yPrompt += lineH;

                        const auto drawPathRow = [&](std::wstring_view value, std::wstring_view part, IDWriteTextFormat* format) noexcept
                        {
                            const float height = MeasureWrappedTextHeight(_dwriteFactory.get(), format, value, maxW, lineH);
                            const auto bounds = D2D1::RectF(contentBounds.left, yPrompt, contentBounds.right, std::min(contentBounds.bottom, yPrompt + height));
                            static_cast<void>(DrawWrappedText(_dwriteFactory.get(), _target.get(), format, _textBrush.get(), value, bounds, lineH));
                            if (! part.empty())
                                addSemanticRegion(bounds, value, std::format(L"{0}.{1}", suffix, part));
                            yPrompt += height;
                        };
                        if (! name.empty() && ! parent.empty())
                        {
                            if (conflictPromptContent.commonName.empty())
                            {
                                drawPathRow(name, L"Name", _headerFormat.get());
                                yPrompt += rowGap;
                            }
                            drawPathRow(parent, L"Parent", _smallFormat.get());
                        }
                        else
                            drawPathRow(path, L"", _smallFormat.get());
                        addSemanticRegion(D2D1::RectF(contentBounds.left, contextTop, contentBounds.right, std::min(contentBounds.bottom, yPrompt)),
                                          std::format(L"{0} {1}", label, name.empty() || parent.empty() ? path : parent),
                                          suffix);

                        // Comparison facts have their own complete row. Do not trade away
                        // path/heading width or truncate the value to a metadata column.
                        if (! metadata.empty())
                        {
                            yPrompt += rowGap;
                            const float metadataHeight = MeasureWrappedTextHeight(_dwriteFactory.get(), _smallFormat.get(), metadata, maxW, lineH);
                            const D2D1_RECT_F metadataRc =
                                D2D1::RectF(contentBounds.left, yPrompt, contentBounds.right, std::min(contentBounds.bottom, yPrompt + metadataHeight));
                            static_cast<void>(
                                DrawWrappedText(_dwriteFactory.get(), _target.get(), _smallFormat.get(), _textBrush.get(), metadata, metadataRc, lineH));
                            addSemanticRegion(metadataRc, metadata, std::format(L"{0}.Metadata", suffix));
                            yPrompt += metadataHeight;
                        }
                        yPrompt += rowGap;
                        const PopupButton inspectButton{D2D1::RectF(contentBounds.left, yPrompt, contentBounds.right, yPrompt + buttonH),
                                                        PopupHitTest{inspectionKind, task.taskId, 0u, task.conflict.generation}};
                        _buttons.push_back(inspectButton);
                        // The retained button owns paint, clipping, focus, and accessibility.
                        yPrompt += buttonH;
                        hasPriorContent = true;
                    };

                    drawTextRegion(
                        conflictPromptContent.state, _bodyFormat.get(), _statusWarningBrush ? _statusWarningBrush.get() : _textBrush.get(), L"State");
                    drawTextRegion(conflictPromptContent.commonName, _headerFormat.get(), _textBrush.get(), L"Name");
                    drawTextRegion(conflictPromptContent.question, _bodyFormat.get(), _textBrush.get(), L"Question");
                    drawTextRegion(conflictPromptContent.facts, _smallFormat.get(), _subTextBrush.get(), L"Facts");
                    auto* body = FindConflictBody(task.taskId);
                    if (! body)
                    {
                        _conflictBodies.push_back(ConflictBodyLayout{.taskId = task.taskId, .context = task.conflict.inspection});
                        body = &_conflictBodies.back();
                    }
                    const float inset   = DipsToPixels(4.0f, _dpi);
                    const float bodyTop = yPrompt + (hasPriorContent ? rowGap : 0.0f);
                    const float resultHeight =
                        conflictPromptContent.inspectionResult.empty()
                            ? 0.0f
                            : MeasureWrappedTextHeight(_dwriteFactory.get(), _smallFormat.get(), conflictPromptContent.inspectionResult, maxW, lineH);
                    const float bodyBottom = rc.bottom - (resultHeight > 0.0f ? resultHeight + rowGap : 0.0f);
                    body->bounds           = D2D1::RectF(rc.left, bodyTop, rc.right, bodyBottom);
                    body->viewport         = body->bounds;
                    const float viewHeight = std::max(0.0f, bodyBottom - bodyTop);
                    const auto measureBody = [&](float widthPx) noexcept
                    {
                        return MeasureConflictPromptContentHeight(_dwriteFactory.get(),
                                                                  _headerFormat.get(),
                                                                  _bodyFormat.get(),
                                                                  _smallFormat.get(),
                                                                  conflictPromptContent,
                                                                  std::max(1.0f, widthPx),
                                                                  lineH,
                                                                  rowGap,
                                                                  buttonH,
                                                                  true) +
                               inset * 2.0f;
                    };
                    body->contentHeightPx = measureBody(rc.right - rc.left - inset * 2.0f);
                    if (body->contentHeightPx > viewHeight)
                    {
                        const float gutter = DipsToPixels(_conflictBodyScrollbarDip, _dpi);
                        if (actionRightToLeft)
                            body->viewport.left += gutter;
                        else
                            body->viewport.right -= gutter;
                    }
                    contentBounds.left    = body->viewport.left + inset;
                    contentBounds.right   = body->viewport.right - inset;
                    maxW                  = std::max(1.0f, contentBounds.right - contentBounds.left);
                    body->contentHeightPx = measureBody(maxW);
                    body->visible         = viewHeight > 0.0f;
                    body->offsetDip       = std::clamp(body->offsetDip, 0.0f, _controlsHost.PixelsToDip(std::max(0.0f, body->contentHeightPx - viewHeight)));
                    yPrompt               = bodyTop + inset - DipsToPixels(body->offsetDip, _dpi);
                    contentBounds.bottom  = yPrompt + body->contentHeightPx;
                    hasPriorContent       = false;
                    comparisonTaskId      = task.taskId;
                    _target->PushAxisAlignedClip(body->viewport, D2D1_ANTIALIAS_MODE_ALIASED);
                    drawContext(conflictPromptContent.sourceLabel,
                                conflictPromptContent.sourcePath,
                                conflictPromptContent.sourceName,
                                conflictPromptContent.sourceParentPath,
                                conflictPromptContent.sourceMetadata,
                                task.operation == FILESYSTEM_DELETE ? L"Deleting" : L"From",
                                PopupHitTest::Kind::TaskConflictInspectSource);
                    drawContext(conflictPromptContent.destinationLabel,
                                conflictPromptContent.destinationPath,
                                conflictPromptContent.destinationName,
                                conflictPromptContent.destinationParentPath,
                                conflictPromptContent.destinationMetadata,
                                L"To",
                                PopupHitTest::Kind::TaskConflictInspectDestination);
                    drawTextRegion(conflictPromptContent.lastNote, _smallFormat.get(), _subTextBrush.get(), L"LastNote");
                    _target->PopAxisAlignedClip();
                    // Inspection feedback stays beside the decisions, even when comparison details scroll.
                    contentBounds    = rc;
                    maxW             = std::max(1.0f, rc.right - rc.left);
                    yPrompt          = rc.bottom - resultHeight;
                    hasPriorContent  = false;
                    comparisonTaskId = 0u;
                    drawTextRegion(conflictPromptContent.inspectionResult, _smallFormat.get(), _textBrush.get(), L"InspectionResult");
                };

                std::wstring graphEtaText;
                if (! task.cancellationRequested && task.discoveryClosed && task.totalBytes > 0 && history && history->hasSmoothedEta &&
                    history->smoothedEtaSeconds > 0.0 && task.completedBytes <= task.totalBytes)
                {
                    const uint64_t seconds = SaturatingCeilNonNegativeToUint64(history->smoothedEtaSeconds);
                    graphEtaText           = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA, FormatDurationHms(seconds));
                }
                else if (! task.cancellationRequested && ! task.discoveryClosed && ! task.discoverySkipped && history && history->hasProvisionalEta &&
                         history->provisionalEtaSeconds > 0.0)
                {
                    // Clearly provisional: discovery is still open, so the value may rise or fall.
                    const uint64_t seconds = SaturatingCeilNonNegativeToUint64(history->provisionalEtaSeconds);
                    graphEtaText           = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA_PROVISIONAL, FormatDurationHms(seconds));
                }

                if (hasConflictPrompt || compactProgressPresentation.showThroughputGraph)
                {
                    const float graphTop    = textY + DipsToPixels(4.0f, _dpi);
                    const float graphBottom = hasConflictPrompt  ? barsBottom
                                              : priorityTransfer ? buttonRowTop - DipsToPixels(8.0f, _dpi)
                                                                 : (barsTop - DipsToPixels(6.0f, _dpi));
                    if (hasConflictPrompt)
                    {
                        drawConflictPromptInfo(D2D1::RectF(barX, graphTop, barX + barW, graphBottom));
                    }
                    else
                    {
                        const float disclosureSize = DipsToPixels(24.0f, _dpi);
                        PopupButton disclosure{};
                        disclosure.bounds = D2D1::RectF(barX, graphTop, barX + disclosureSize, graphTop + disclosureSize);
                        disclosure.hit    = {PopupHitTest::Kind::TaskToggleThroughput, task.taskId, 0u};
                        _buttons.push_back(disclosure);
                        const float labelLeft     = barX + disclosureSize + DipsToPixels(4.0f, _dpi);
                        const float headingHeight = priorityTransfer ? std::max(disclosureSize, FactPairHeight(information.throughput)) : disclosureSize;
                        D2D1_RECT_F rateRect{};
                        if (priorityTransfer)
                        {
                            DrawFactPair(information.throughput,
                                         D2D1::RectF(labelLeft, graphTop, barX + barW, graphTop + headingHeight),
                                         task.taskId,
                                         L"Throughput.Label",
                                         L"Throughput.Rate");
                            const auto& pair    = information.throughput;
                            const float rateTop = pair.stacked ? graphTop + pair.rateHeight + DipsToPixels(2.0f, _dpi) : graphTop;
                            rateRect = D2D1::RectF(barX + barW - (pair.stacked ? pair.width : pair.etaWidth), rateTop, barX + barW, rateTop + pair.etaHeight);
                        }
                        else
                        {
                            const auto label = LoadStringResource(nullptr, IDS_FILEOPS_THROUGHPUT);
                            _target->DrawTextW(label.data(),
                                               static_cast<UINT32>(label.size()),
                                               _bodyFormat.get(),
                                               D2D1::RectF(labelLeft, graphTop, barX + barW, graphTop + headingHeight),
                                               _textBrush.get());
                        }

                        const RateHistory empty{};
                        const RateHistory& graphHistory = history ? *history : empty;
                        const float bodyTop             = graphTop + headingHeight + DipsToPixels(4.0f, _dpi);
                        const D2D1_RECT_F bodyRect      = D2D1::RectF(barX, bodyTop, barX + barW, graphBottom);
                        if (IsTaskThroughputExpanded(task.taskId))
                        {
                            bool showAnimation             = false;
                            const std::wstring overlayText = GraphOverlayTextForStatus(task, taskStatus, showAnimation);
                            const bool rainbowMode         = folderWindow && folderWindow->GetTheme().menu.rainbowMode;
                            DrawBandwidthGraph(bodyRect,
                                               task.taskId,
                                               graphHistory,
                                               graphEtaText,
                                               overlayText,
                                               rainbowMode,
                                               true,
                                               ! priorityTransfer,
                                               priorityTransfer ? &rateRect : nullptr,
                                               priorityTransfer ? &primaryEtaRect : nullptr);
                        }
                        else if (! priorityTransfer)
                        {
                            const std::wstring rateText = FormatRateLabelText(CurrentBandwidthForGraphMarker(graphHistory), graphHistory.stallKind);
                            const std::wstring etaText  = graphEtaText.empty() ? LoadStringResource(nullptr, IDS_FILEOPS_ETA_UNAVAILABLE) : graphEtaText;
                            const std::wstring summary  = std::format(L"{0} · {1}", rateText, etaText);
                            static_cast<void>(
                                DrawWrappedText(_dwriteFactory.get(), _target.get(), _smallFormat.get(), _subTextBrush.get(), summary, bodyRect, lineH));
                            _hostedTooltipDescriptors.push_back({.bounds       = bodyRect,
                                                                 .text         = summary,
                                                                 .automationId = std::format(L"FileOperations.Task.{}.ThroughputSummary", task.taskId),
                                                                 .liveText     = true});
                        }
                    }
                }

                const WholeTaskProgressPresentation progressPresentation = ResolveWholeTaskProgressPresentation(task);
                if (progressPresentation.showTransferProgress && ! task.discoveryClosed)
                {
                    const D2D1_RECT_F barRc = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);
                    _hostedProgressDescriptors.push_back(
                        HostedProgressDescriptor{barRc, task.taskId, progressPresentation.fraction * 100.0, ! progressPresentation.determinate});
                }
                else if (hasConflictPrompt || ! progressPresentation.showTransferProgress)
                {
                    // Conflict prompt uses the progress bar area so actions and the apply-to-all toggle sit close together.
                }
                else if (task.operation == FILESYSTEM_DELETE)
                {
                    const D2D1_RECT_F totalBarRc   = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);
                    const bool hostedHasTotalBytes = task.totalBytes > 0 && task.completedBytes <= task.totalBytes;
                    // A closed, known denominator is determinate even before the first item completes.
                    // Preserve the existing byte-first-after-progress / item fallback policy.
                    const bool hostedHasUsefulItems = task.totalItems > 0;
                    const bool hostedUseBytes       = hostedHasTotalBytes && (task.completedBytes > 0 || ! hostedHasUsefulItems);
                    const bool hostedUseItems       = ! hostedUseBytes && hostedHasUsefulItems;
                    double hostedFraction           = 0.0;
                    if (hostedUseBytes)
                    {
                        hostedFraction = static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes);
                    }
                    else if (hostedUseItems)
                    {
                        hostedFraction = static_cast<double>(std::min(task.completedItems, task.totalItems)) / static_cast<double>(task.totalItems);
                    }
                    _hostedProgressDescriptors.push_back(
                        HostedProgressDescriptor{totalBarRc, task.taskId, std::clamp(hostedFraction, 0.0, 1.0) * 100.0, ! hostedUseBytes && ! hostedUseItems});
                }
                else
                {
                    const D2D1_RECT_F totalBarRc = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);
                    HostedProgressDescriptor progressDescriptor{
                        totalBarRc, task.taskId, progressPresentation.fraction * 100.0, ! progressPresentation.determinate};
                    progressDescriptor.segmented             = task.verificationRequested;
                    progressDescriptor.primarySegmentValue   = TransferSegmentPercent(task);
                    progressDescriptor.secondarySegmentValue = VerificationSegmentPercent(task);
                    progressDescriptor.verificationSettled   = task.finished;
                    _hostedProgressDescriptors.push_back(progressDescriptor);
                }

                {
                    const float btnGapX = DipsToPixels(8.0f, _dpi);
                    const float rowW    = std::max(0.0f, contentRight - textX);
                    if (rowW > 1.0f && ! task.cancellationRequested)
                    {
                        const float rowTop    = buttonRowTop;
                        const float rowBottom = buttonRowBottom;

                        if (hasConflictPrompt && task.conflict.buttonsPublishable)
                        {
                            const float applyTop = rowTop;
                            float buttonsTop     = rowTop;
                            if (task.conflict.applyToAllEligible)
                            {
                                // The scope control is rendered only when the typed class, item
                                // kinds, and endpoint profiles can be cached safely.
                                const float applyBottom = applyTop + DipsToPixels(_conflictActionMeasurements[row.taskIndex].scopeHeightDip, _dpi);
                                buttonsTop              = applyBottom + conflictRowGapY;

                                // The hosted Checkbox owns both the indicator and label. Painting a second
                                // checkbox below an opaque hosted Button concealed its checked state.
                                PopupButton applyBtn{};
                                applyBtn.bounds               = D2D1::RectF(textX, applyTop, contentRight, applyBottom);
                                applyBtn.hit.kind             = PopupHitTest::Kind::TaskConflictToggleApplyToAll;
                                applyBtn.hit.taskId           = task.taskId;
                                applyBtn.hit.promptGeneration = task.conflict.generation;
                                _buttons.push_back(applyBtn);
                            }

                            const auto& actions             = _conflictActionMeasurements[row.taskIndex];
                            const size_t primaryActionCount = std::min(task.conflict.primaryActionCount, task.conflict.primaryActions.size());
                            for (size_t index = 0u; index < actions.count; ++index)
                            {
                                const auto local = actions.bounds[index];
                                PopupButton button{};
                                button.bounds               = D2D1::RectF(textX + DipsToPixels(local.left, _dpi),
                                                                          buttonsTop + DipsToPixels(local.top, _dpi),
                                                                          textX + DipsToPixels(local.right, _dpi),
                                                                          buttonsTop + DipsToPixels(local.bottom, _dpi));
                                button.hit.taskId           = task.taskId;
                                button.hit.promptGeneration = task.conflict.generation;
                                if (index < primaryActionCount)
                                {
                                    button.hit.kind = PopupHitTest::Kind::TaskConflictAction;
                                    button.hit.data = static_cast<uint32_t>(RawConflictAction(ConflictActionFromRaw(task.conflict.primaryActions[index])));
                                    _buttons.push_back(button);
                                    DrawButton(button, _buttonSmallFormat.get(), actions.labels[index]);
                                }
                                else
                                {
                                    button.hit.kind = PopupHitTest::Kind::TaskConflictMore;
                                    button.hit.data = static_cast<uint32_t>(std::min(task.conflict.overflowActionCount, task.conflict.overflowActions.size()));
                                    _buttons.push_back(button);
                                    DrawMenuButton(button, _buttonSmallFormat.get(), actions.labels[index]);
                                }
                            }
                        }
                        else if (hasConflictPrompt)
                        {
                            // Metadata loading is decision preparation. It exposes no stale
                            // transfer controls and no action whose required facts are incomplete.
                        }
                        // During copy/move discovery-ahead, keep the transfer controls available before bytes start moving.
                        else if (task.discoveryAheadActive)
                        {
                            const bool showStartNow     = ! task.started && ! task.overlapQueued && (task.waitingForOthers || task.waitingInQueue);
                            const std::wstring skipText = LoadStringResource(nullptr, IDS_FILEOPS_BTN_SKIP);
                            if (showStartNow && showCopyMoveControls && ! speedLimitText.empty())
                            {
                                const float available = std::max(0.0f, rowW - btnGapX * 2.0f);
                                const float minEach   = DipsToPixels(68.0f, _dpi);

                                float startNowW = DipsToPixels(96.0f, _dpi);
                                float cancelW   = DipsToPixels(84.0f, _dpi);
                                float limitW    = std::max(0.0f, available - startNowW - cancelW);

                                if (available < minEach * 3.0f)
                                {
                                    const float eachW = available / 3.0f;
                                    startNowW         = eachW;
                                    cancelW           = eachW;
                                    limitW            = eachW;
                                }
                                else
                                {
                                    const float minLimitW = speedLimitWidthPx;
                                    if (limitW < minLimitW)
                                    {
                                        const float minSideW          = DipsToPixels(72.0f, _dpi);
                                        const float remainingForSides = std::max(0.0f, available - minLimitW);
                                        const float sideW             = std::max(minSideW, remainingForSides / 2.0f);
                                        startNowW                     = std::min(startNowW, sideW);
                                        cancelW                       = std::min(cancelW, sideW);
                                        limitW                        = std::max(0.0f, available - startNowW - cancelW);
                                    }

                                    limitW = std::min(limitW, minLimitW);
                                }

                                float xBtn = textX;

                                PopupButton startNowBtn{};
                                startNowBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + startNowW, rowBottom);
                                startNowBtn.hit.kind   = PopupHitTest::Kind::TaskStartNow;
                                startNowBtn.hit.taskId = task.taskId;
                                _buttons.push_back(startNowBtn);
                                DrawButton(startNowBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_BTN_START_NOW));
                                xBtn += startNowW + btnGapX;

                                PopupButton limitBtn{};
                                limitBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + limitW, rowBottom);
                                limitBtn.hit.kind   = PopupHitTest::Kind::TaskSpeedLimit;
                                limitBtn.hit.taskId = task.taskId;
                                _buttons.push_back(limitBtn);
                                DrawSelectorButton(limitBtn, _buttonSmallFormat.get(), speedLimitText);
                                xBtn += limitW + btnGapX;

                                PopupButton cancelBtn{};
                                cancelBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + cancelW, rowBottom);
                                cancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                                cancelBtn.hit.taskId = task.taskId;
                                _buttons.push_back(cancelBtn);
                                DrawButton(cancelBtn, _buttonSmallFormat.get(), cancelText);
                            }
                            else if (showStartNow)
                            {
                                const float startNowW = std::max(0.0f, (rowW - btnGapX) * 0.5f);
                                const float cancelW   = std::max(0.0f, rowW - btnGapX - startNowW);

                                PopupButton startNowBtn{};
                                startNowBtn.bounds     = D2D1::RectF(textX, rowTop, textX + startNowW, rowBottom);
                                startNowBtn.hit.kind   = PopupHitTest::Kind::TaskStartNow;
                                startNowBtn.hit.taskId = task.taskId;
                                _buttons.push_back(startNowBtn);
                                DrawButton(startNowBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_BTN_START_NOW));

                                PopupButton cancelBtn{};
                                cancelBtn.bounds     = D2D1::RectF(textX + startNowW + btnGapX, rowTop, textX + startNowW + btnGapX + cancelW, rowBottom);
                                cancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                                cancelBtn.hit.taskId = task.taskId;
                                _buttons.push_back(cancelBtn);
                                DrawButton(cancelBtn, _buttonSmallFormat.get(), cancelText);
                            }
                            else if (showCopyMoveControls && ! speedLimitText.empty())
                            {
                                const float available = std::max(0.0f, rowW - btnGapX * 2.0f);
                                const float minEach   = DipsToPixels(68.0f, _dpi);

                                // Retained Buttons paint Body typography, not the legacy 11-DIP
                                // small format. Keep the translated discovery label and 12-DIP
                                // padding on both sides instead of clipping it into an 84-DIP slot.
                                IDWriteTextFormat* const actionFormat = _controlsHost.GetTextFormat(DxUi::FontRole::Body);
                                if (actionFormat && (_discoverySkipMeasuredFormat.get() != actionFormat || _discoverySkipMeasuredText != skipText))
                                {
                                    const uint64_t measureStartUs = PerfNowUs();
                                    const float labelWidthDip     = MeasureTextWidth(_dwriteFactory.get(), actionFormat, skipText, 4096.0f, 128.0f);
                                    if (labelWidthDip > 0.0f)
                                    {
                                        _discoverySkipWidthDip       = std::max(84.0f, std::ceil(labelWidthDip + 24.0f));
                                        _discoverySkipMeasuredText   = skipText;
                                        _discoverySkipMeasuredFormat = actionFormat;
                                    }
                                    Debug::Perf::Emit(L"FileOps.Popup.DiscoveryActions.MeasureUs",
                                                      L"discovery label cache miss",
                                                      PerfElapsedUs(measureStartUs),
                                                      1u,
                                                      0u,
                                                      S_OK);
                                }
                                float skipW   = DipsToPixels(_discoverySkipWidthDip, _dpi);
                                float cancelW = DipsToPixels(84.0f, _dpi);
                                float limitW  = std::max(0.0f, available - skipW - cancelW);

                                if (available < minEach * 3.0f)
                                {
                                    const float eachW = available / 3.0f;
                                    skipW             = eachW;
                                    cancelW           = eachW;
                                    limitW            = eachW;
                                }
                                else
                                {
                                    const float minLimitW = speedLimitWidthPx;
                                    if (limitW < minLimitW)
                                    {
                                        const float minSideW          = DipsToPixels(72.0f, _dpi);
                                        const float remainingForSides = std::max(0.0f, available - minLimitW);
                                        const float sideW             = std::max(minSideW, remainingForSides / 2.0f);
                                        skipW                         = std::min(skipW, sideW);
                                        cancelW                       = std::min(cancelW, sideW);
                                        limitW                        = std::max(0.0f, available - skipW - cancelW);
                                    }

                                    // Speed limit is an auxiliary control; it must never out-weigh
                                    // the primary actions by absorbing the rest of the row.
                                    limitW = std::min(limitW, minLimitW);
                                }

                                float xBtn = textX;

                                PopupButton skipBtn{};
                                skipBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + skipW, rowBottom);
                                skipBtn.hit.kind   = PopupHitTest::Kind::TaskSkip;
                                skipBtn.hit.taskId = task.taskId;
                                _buttons.push_back(skipBtn);
                                DrawButton(skipBtn, _buttonSmallFormat.get(), skipText);
                                xBtn += skipW + btnGapX;

                                PopupButton limitBtn{};
                                limitBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + limitW, rowBottom);
                                limitBtn.hit.kind   = PopupHitTest::Kind::TaskSpeedLimit;
                                limitBtn.hit.taskId = task.taskId;
                                _buttons.push_back(limitBtn);
                                DrawSelectorButton(limitBtn, _buttonSmallFormat.get(), speedLimitText);
                                xBtn += limitW + btnGapX;

                                PopupButton calcCancelBtn{};
                                calcCancelBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + cancelW, rowBottom);
                                calcCancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                                calcCancelBtn.hit.taskId = task.taskId;
                                _buttons.push_back(calcCancelBtn);
                                DrawButton(calcCancelBtn, _buttonSmallFormat.get(), cancelText);
                            }
                            else
                            {
                                const float skipW       = std::max(0.0f, (rowW - btnGapX) * 0.5f);
                                const float calcCancelW = std::max(0.0f, rowW - btnGapX - skipW);

                                PopupButton skipBtn{};
                                skipBtn.bounds     = D2D1::RectF(textX, rowTop, textX + skipW, rowBottom);
                                skipBtn.hit.kind   = PopupHitTest::Kind::TaskSkip;
                                skipBtn.hit.taskId = task.taskId;
                                _buttons.push_back(skipBtn);
                                DrawButton(skipBtn, _buttonSmallFormat.get(), skipText);

                                PopupButton calcCancelBtn{};
                                calcCancelBtn.bounds     = D2D1::RectF(textX + skipW + btnGapX, rowTop, textX + skipW + btnGapX + calcCancelW, rowBottom);
                                calcCancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                                calcCancelBtn.hit.taskId = task.taskId;
                                _buttons.push_back(calcCancelBtn);
                                DrawButton(calcCancelBtn, _buttonSmallFormat.get(), cancelText);
                            }
                        }
                        else if (showCopyMoveControls && ! speedLimitText.empty())
                        {
                            // A queued task has nothing to pause yet; Cancel and the speed limit
                            // (which applies once it starts) are the only meaningful controls,
                            // plus Start now when the task is explicitly waiting behind others.
                            const bool showPause    = task.started;
                            const bool showStartNow = ! task.started && ! task.overlapQueued && (task.waitingForOthers || task.waitingInQueue);
                            const float actionCount = (showPause ? 1.0f : 0.0f) + (showStartNow ? 1.0f : 0.0f) + 2.0f;
                            const float available   = std::max(0.0f, rowW - btnGapX * std::max(0.0f, actionCount - 1.0f));
                            const float minEach     = DipsToPixels(68.0f, _dpi);

                            float startNowW = showStartNow ? DipsToPixels(96.0f, _dpi) : 0.0f;
                            float pauseW    = showPause ? DipsToPixels(84.0f, _dpi) : 0.0f;
                            float cancelW   = DipsToPixels(84.0f, _dpi);
                            float limitW    = std::max(0.0f, available - startNowW - pauseW - cancelW);

                            if (available < minEach * actionCount)
                            {
                                const float eachW = actionCount > 0.0f ? (available / actionCount) : 0.0f;
                                startNowW         = showStartNow ? eachW : 0.0f;
                                pauseW            = showPause ? eachW : 0.0f;
                                cancelW           = eachW;
                                limitW            = eachW;
                            }
                            else
                            {
                                const float minLimitW = speedLimitWidthPx;
                                if (limitW < minLimitW)
                                {
                                    const float minSideW          = DipsToPixels(72.0f, _dpi);
                                    const float remainingForSides = std::max(0.0f, available - minLimitW);
                                    const float sideCount         = (showStartNow ? 1.0f : 0.0f) + (showPause ? 1.0f : 0.0f) + 1.0f;
                                    const float sideW             = std::max(minSideW, remainingForSides / sideCount);
                                    startNowW                     = showStartNow ? std::min(startNowW, sideW) : 0.0f;
                                    pauseW                        = showPause ? std::min(pauseW, sideW) : 0.0f;
                                    cancelW                       = std::min(cancelW, sideW);
                                    limitW                        = std::max(0.0f, available - startNowW - pauseW - cancelW);
                                }

                                // Speed limit is an auxiliary control; it must never out-weigh
                                // the primary actions by absorbing the rest of the row.
                                limitW = std::min(limitW, minLimitW);
                            }

                            float xBtn = textX;

                            if (showStartNow)
                            {
                                PopupButton startNowBtn{};
                                startNowBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + startNowW, rowBottom);
                                startNowBtn.hit.kind   = PopupHitTest::Kind::TaskStartNow;
                                startNowBtn.hit.taskId = task.taskId;
                                _buttons.push_back(startNowBtn);
                                DrawButton(startNowBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_BTN_START_NOW));
                                xBtn += startNowW + btnGapX;
                            }

                            if (showPause)
                            {
                                PopupButton pauseBtn{};
                                pauseBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + pauseW, rowBottom);
                                pauseBtn.hit.kind   = PopupHitTest::Kind::TaskPause;
                                pauseBtn.hit.taskId = task.taskId;
                                _buttons.push_back(pauseBtn);
                                DrawButton(pauseBtn, _buttonSmallFormat.get(), pauseText);
                                xBtn += pauseW + btnGapX;
                            }

                            PopupButton limitBtn{};
                            limitBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + limitW, rowBottom);
                            limitBtn.hit.kind   = PopupHitTest::Kind::TaskSpeedLimit;
                            limitBtn.hit.taskId = task.taskId;
                            _buttons.push_back(limitBtn);
                            DrawSelectorButton(limitBtn, _buttonSmallFormat.get(), speedLimitText);
                            xBtn += limitW + btnGapX;

                            PopupButton cancelBtn{};
                            cancelBtn.bounds     = D2D1::RectF(xBtn, rowTop, xBtn + cancelW, rowBottom);
                            cancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                            cancelBtn.hit.taskId = task.taskId;
                            _buttons.push_back(cancelBtn);
                            DrawButton(cancelBtn, _buttonSmallFormat.get(), cancelText);
                        }
                        else
                        {
                            const bool showPause    = task.started;
                            const bool showStartNow = ! task.started && ! task.overlapQueued && (task.waitingForOthers || task.waitingInQueue);
                            if (showPause)
                            {
                                const float pauseW  = std::max(0.0f, (rowW - btnGapX) * 0.5f);
                                const float cancelW = std::max(0.0f, rowW - btnGapX - pauseW);

                                PopupButton pauseBtn{};
                                pauseBtn.bounds     = D2D1::RectF(textX, rowTop, textX + pauseW, rowBottom);
                                pauseBtn.hit.kind   = PopupHitTest::Kind::TaskPause;
                                pauseBtn.hit.taskId = task.taskId;
                                _buttons.push_back(pauseBtn);
                                DrawButton(pauseBtn, _buttonSmallFormat.get(), pauseText);

                                PopupButton cancelBtn{};
                                cancelBtn.bounds     = D2D1::RectF(textX + pauseW + btnGapX, rowTop, textX + pauseW + btnGapX + cancelW, rowBottom);
                                cancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                                cancelBtn.hit.taskId = task.taskId;
                                _buttons.push_back(cancelBtn);
                                DrawButton(cancelBtn, _buttonSmallFormat.get(), cancelText);
                            }
                            else if (showStartNow)
                            {
                                const float startNowW = std::max(0.0f, (rowW - btnGapX) * 0.5f);
                                const float cancelW   = std::max(0.0f, rowW - btnGapX - startNowW);

                                PopupButton startNowBtn{};
                                startNowBtn.bounds     = D2D1::RectF(textX, rowTop, textX + startNowW, rowBottom);
                                startNowBtn.hit.kind   = PopupHitTest::Kind::TaskStartNow;
                                startNowBtn.hit.taskId = task.taskId;
                                _buttons.push_back(startNowBtn);
                                DrawButton(startNowBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_BTN_START_NOW));

                                PopupButton cancelBtn{};
                                cancelBtn.bounds     = D2D1::RectF(textX + startNowW + btnGapX, rowTop, textX + startNowW + btnGapX + cancelW, rowBottom);
                                cancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                                cancelBtn.hit.taskId = task.taskId;
                                _buttons.push_back(cancelBtn);
                                DrawButton(cancelBtn, _buttonSmallFormat.get(), cancelText);
                            }
                            else
                            {
                                PopupButton cancelBtn{};
                                cancelBtn.bounds     = D2D1::RectF(textX, rowTop, textX + rowW, rowBottom);
                                cancelBtn.hit.kind   = PopupHitTest::Kind::TaskCancel;
                                cancelBtn.hit.taskId = task.taskId;
                                _buttons.push_back(cancelBtn);
                                DrawButton(cancelBtn, _buttonSmallFormat.get(), cancelText);
                            }
                        }
                    }
                }
                if (! hasConflictPrompt && information.concurrency.rate)
                    DrawFactPair(information.concurrency,
                                 D2D1::RectF(textX, buttonRowBottom + DipsToPixels(6.0f, _dpi), contentRight, buttonRowBottom + noteHeight),
                                 task.taskId,
                                 L"Concurrency",
                                 {});
            }

            const float gapAfter = (rowIndex + 1u < rowCount) ? cardGap : 0.0f;
            y += taskCardH + gapAfter;
        }

        _target->PopAxisAlignedClip();
    }
    if (SUCCEEDED(hrEndDraw))
        hrEndDraw = preparedBackground->Close();
    if (SUCCEEDED(hrEndDraw))
    {
        _preparedBackground = std::move(preparedBackground);
        // A prepared frame ends any recovery backoff.
        _framePrepareFailureCount     = 0u;
        _framePrepareFirstFailureTick = 0u;
        _framePrepareRetryTick        = 0u;
        _hostReattachAttemptCount     = 0u;
    }
    else
        _preparedBackground.reset();
    const uint64_t drawUs = capturePerf ? PerfElapsedUs(drawStartedUs) : 0u;
    if (capturePerf)
    {
        Debug::Perf::Emit(L"FileOps.Popup.Render.PreparedBackground",
                          L"host command list",
                          0u,
                          _preparedBackground ? 1u : 0u,
                          _target.get() == _controlsHost.GetDeviceContext() ? 1u : 0u,
                          hrEndDraw);
        Debug::Perf::Emit(L"FileOps.Popup.Render.ClientPixels",
                          L"",
                          0u,
                          static_cast<uint64_t>(std::max(0L, _clientSize.cx)),
                          static_cast<uint64_t>(std::max(0L, _clientSize.cy)),
                          hrEndDraw);
    }
    SyncHostedControls(hwnd, snapshot);

    if (capturePerf)
    {
        uint64_t informationalTaskCount = 0u;
        for (const TaskSnapshot& task : snapshot)
        {
            if (task.kind == TaskSnapshot::Kind::Informational)
            {
                ++informationalTaskCount;
            }
        }

        Debug::Perf::Emit(L"FileOps.Popup.Render.BuildSnapshotUs", L"", snapshotUs, informationalTaskCount, static_cast<uint64_t>(rowCount), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Render.CardLayoutUs", L"", cardLayoutUs, informationalTaskCount, static_cast<uint64_t>(rowCount), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Render.AutoResizeUs", L"", autoResizeUs, informationalTaskCount, static_cast<uint64_t>(rowCount), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Render.ScrollLayoutUs", L"", scrollLayoutUs, informationalTaskCount, static_cast<uint64_t>(rowCount), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Render.DrawUs", L"", drawUs, informationalTaskCount, static_cast<uint64_t>(rowCount), hrEndDraw);
        Debug::Perf::Emit(
            L"FileOps.Popup.Render.TotalUs", L"", PerfElapsedUs(renderStartedUs), informationalTaskCount, static_cast<uint64_t>(rowCount), hrEndDraw);
    }

    if (hrEndDraw == D2DERR_RECREATE_TARGET || hrEndDraw == DXGI_ERROR_DEVICE_REMOVED || hrEndDraw == DXGI_ERROR_DEVICE_RESET)
    {
        // A lost device is recoverable: prepare the next generation after the backoff.
        DiscardDeviceResources();
        if (OnFramePrepareFailed(hwnd, hrEndDraw))
            paintFailureSurfaceGround();
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateLastPopupRect(HWND hwnd) noexcept
{
    if (! hwnd || ! fileOps)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    if (! IsWindowVisible(hwnd) || IsIconic(hwnd))
    {
        return;
    }

    RECT rc{};
    if (! GetWindowRect(hwnd, &rc))
    {
        return;
    }

    fileOps->UpdateLastPopupRect(rc);
    fileOps->SavePopupPlacement(hwnd);
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateCaptionStatus(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept
{
    CaptionStatus computed = snapshot.empty() ? CaptionStatus::None : CaptionStatus::Ok;

    bool sawWarning = false;
    for (const TaskSnapshot& task : snapshot)
    {
        const TaskStatusKind status = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
        if (StatusIsError(status))
        {
            computed = CaptionStatus::Error;
            break;
        }

        if (StatusIsWarning(status))
        {
            sawWarning = true;
        }
    }

    if (computed != CaptionStatus::Error && sawWarning)
    {
        computed = CaptionStatus::Warning;
    }

    if (_captionStatus == computed)
    {
        return;
    }

    _captionStatus = computed;

    if (hwnd)
    {
        RedrawWindow(hwnd, nullptr, nullptr, RDW_FRAME | RDW_NOERASE | RDW_NOCHILDREN);
    }
}

bool FileOperationsPopupInternal::FileOperationsPopupState::EnsureTaskbarList() noexcept
{
    if (_taskbarList)
    {
        return true;
    }

    const ULONGLONG nowTick = GetTickCount64();
    if (_taskbarListRetryAfterTick != 0 && nowTick < _taskbarListRetryAfterTick)
    {
        return false;
    }

    const auto scheduleRetry = [&]() noexcept
    {
        _taskbarList.reset();
        _taskbarListRetryAfterTick = nowTick + kTaskbarListRetryDelayMs;
        return false;
    };

    ++_taskbarListAttemptCount;
#ifdef ENABLE_TESTS
    unsigned int forcedFailures = g_fileOperationsTaskbarListForcedFailures.load(std::memory_order_acquire);
    while (forcedFailures > 0)
    {
        if (g_fileOperationsTaskbarListForcedFailures.compare_exchange_weak(
                forcedFailures, forcedFailures - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            return scheduleRetry();
        }
    }
#endif

    wil::com_ptr<ITaskbarList3> taskbar;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(taskbar.put()));
    if (FAILED(hr) || ! taskbar)
    {
        return scheduleRetry();
    }

    hr = taskbar->HrInit();
    if (FAILED(hr))
    {
        return scheduleRetry();
    }

    _taskbarList               = std::move(taskbar);
    _taskbarListRetryAfterTick = 0;
    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateTaskbarProgress(HWND hwnd) noexcept
{
    const GlobalFileOperationsStatusSummary summary = BuildGlobalStatusSummary(BuildSnapshot());
    const GlobalTaskbarProgressModel model          = BuildGlobalTaskbarProgressModel(summary);
    ApplyTaskbarProgress(hwnd, model.state, model.completed, model.total);
}

void FileOperationsPopupInternal::FileOperationsPopupState::ApplyTaskbarProgress(HWND hwnd, uint32_t state, uint64_t completed, uint64_t total) noexcept
{
    ++_taskbarUpdateCount;
    _cachedTaskbarState     = state;
    _cachedTaskbarCompleted = completed;
    _cachedTaskbarTotal     = total;
#ifdef ENABLE_TESTS
    _lastTaskbarState = state;
#endif
    if (! hwnd || ! _taskbarButtonReady || ! EnsureTaskbarList())
    {
        return;
    }

    const auto progressState = static_cast<TBPFLAG>(state);
    if (progressState == TBPF_NOPROGRESS || progressState == TBPF_INDETERMINATE || total == 0)
    {
        const HRESULT stateHr = _taskbarList->SetProgressState(hwnd, progressState);
#ifdef ENABLE_TESTS
        if (SUCCEEDED(stateHr))
            ++_taskbarSuccessCount;
#else
        static_cast<void>(stateHr);
#endif
        return;
    }

    const HRESULT valueHr = _taskbarList->SetProgressValue(hwnd, std::min(completed, total), total);
    const HRESULT stateHr = _taskbarList->SetProgressState(hwnd, progressState);
#ifdef ENABLE_TESTS
    if (SUCCEEDED(valueHr) && SUCCEEDED(stateHr))
        ++_taskbarSuccessCount;
#else
    static_cast<void>(valueHr);
    static_cast<void>(stateHr);
#endif
}

void FileOperationsPopupInternal::FileOperationsPopupState::ClearTaskbarProgress(HWND hwnd) noexcept
{
    if (! hwnd || ! _taskbarList)
    {
        return;
    }

    static_cast<void>(_taskbarList->SetProgressState(hwnd, TBPF_NOPROGRESS));
}

void FileOperationsPopupInternal::FileOperationsPopupState::PaintCaptionStatusGlyph(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    if (! folderWindow)
    {
        return;
    }

    if (_captionStatus == CaptionStatus::None)
    {
        return;
    }

    const AppTheme& theme = folderWindow->GetTheme();
    if (theme.highContrast)
    {
        return;
    }

    RECT windowScreen{};
    if (! GetWindowRect(hwnd, &windowScreen))
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(hwnd, &client))
    {
        return;
    }

    POINT clientTopLeftScreen{0, 0};
    if (! ClientToScreen(hwnd, &clientTopLeftScreen))
    {
        return;
    }

    const int windowW         = std::max(0, static_cast<int>(windowScreen.right - windowScreen.left));
    const int clientW         = std::max(0, static_cast<int>(client.right - client.left));
    const int nonClientTopH   = std::max(0, static_cast<int>(clientTopLeftScreen.y - windowScreen.top));
    const int nonClientRightW = std::max(0, static_cast<int>(windowScreen.right - (clientTopLeftScreen.x + clientW)));

    if (windowW <= 0 || nonClientTopH <= 0)
    {
        return;
    }

    const UINT dpi    = GetDpiForWindow(hwnd);
    const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const bool hasSys = (style & WS_SYSMENU) != 0;
    const bool hasMin = (style & WS_MINIMIZEBOX) != 0;
    const bool hasMax = (style & WS_MAXIMIZEBOX) != 0;
    const int buttonW = GetSystemMetricsForDpi(SM_CXSIZE, dpi);

    int buttonCount = 0;
    if (hasSys)
    {
        buttonCount += 1; // Close
    }
    if (hasMax)
    {
        buttonCount += 1;
    }
    if (hasMin)
    {
        buttonCount += 1;
    }

    if (buttonCount <= 0 || buttonW <= 0)
    {
        return;
    }

    const int iconSize = DipsToPixels(20, dpi);
    const int gap      = DipsToPixels(8, dpi);

    const int buttonsLeft = windowW - nonClientRightW - buttonW * buttonCount;
    const int iconRight   = buttonsLeft - gap;
    const int iconLeft    = iconRight - iconSize;
    const int iconTop     = std::max(0, (nonClientTopH - iconSize) / 2);

    if (iconRight <= iconLeft || iconTop + iconSize <= 0)
    {
        return;
    }

    RECT iconRc{iconLeft, iconTop, iconRight, iconTop + iconSize};
    RECT targetRc{0, 0, windowW, nonClientTopH};

    wchar_t fluentGlyph = 0;
    wchar_t fallback    = 0;
    D2D1::ColorF color  = ColorFromCOLORREF(theme.menu.text);

    switch (_captionStatus)
    {
        case CaptionStatus::Ok:
            fluentGlyph = FluentIcons::kCheckMark;
            fallback    = FluentIcons::kFallbackCheckMark;
            color       = theme.accent;
            break;
        case CaptionStatus::Warning:
            fluentGlyph = FluentIcons::kWarning;
            fallback    = FluentIcons::kFallbackWarning;
            color       = theme.folderView.warningText;
            break;
        case CaptionStatus::Error:
            fluentGlyph = FluentIcons::kError;
            fallback    = FluentIcons::kFallbackError;
            color       = theme.folderView.errorText;
            break;
        case CaptionStatus::None:
        default: return;
    }

    if (! EnsureCaptionGlyphTarget(dpi))
    {
        return;
    }

    const bool useFluentGlyph = _captionGlyphFormat != nullptr && DirectWriteFormatHasGlyph(_dwriteFactory.get(), _captionGlyphFormat.get(), fluentGlyph);
    IDWriteTextFormat* format = useFluentGlyph ? _captionGlyphFormat.get() : _captionGlyphFallbackFormat.get();
    const wchar_t glyph       = useFluentGlyph ? fluentGlyph : fallback;
    if (! format || glyph == 0)
    {
        return;
    }

    wil::unique_hdc_window hdc = wil::GetWindowDC(hwnd);
    if (! hdc)
    {
        return;
    }

    const HRESULT bindHr = _captionGlyphTarget->BindDC(hdc.get(), &targetRc);
    if (FAILED(bindHr))
    {
        _captionGlyphBrush.reset();
        _captionGlyphTarget.reset();
        return;
    }

    _captionGlyphBrush->SetColor(color);
    const D2D1_RECT_F glyphRc = RectPixelsToDips(iconRc, dpi);
    const wchar_t glyphText[2]{glyph, 0};

    _captionGlyphTarget->BeginDraw();
    _captionGlyphTarget->DrawText(glyphText, 1u, format, glyphRc, _captionGlyphBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP, DWRITE_MEASURING_MODE_NATURAL);
    const HRESULT endHr = _captionGlyphTarget->EndDraw();
    if (endHr == D2DERR_RECREATE_TARGET)
    {
        _captionGlyphBrush.reset();
        _captionGlyphTarget.reset();
    }
}

PopupHitTest FileOperationsPopupInternal::FileOperationsPopupState::HitTest(float x, float y) const noexcept
{
    for (auto it = _buttons.rbegin(); it != _buttons.rend(); ++it)
    {
        if (! PointInRectF(it->bounds, x, y))
        {
            continue;
        }

        // Task-card buttons live inside the scrolled list viewport; a card scrolled under the
        // footer must never steal clicks from the footer controls.
        const bool isFooterButton = it->hit.kind == PopupHitTest::Kind::FooterCancelAll || it->hit.kind == PopupHitTest::Kind::FooterPauseResumeAll ||
                                    it->hit.kind == PopupHitTest::Kind::FooterAutoDismiss || it->hit.kind == PopupHitTest::Kind::FooterQueueMode ||
                                    it->hit.kind == PopupHitTest::Kind::FooterDensity || it->hit.kind == PopupHitTest::Kind::FooterOptions ||
                                    it->hit.kind == PopupHitTest::Kind::FooterToggleDetails;
        if (! isFooterButton && ! PointInRectF(_listViewportRect, x, y))
        {
            continue;
        }

        if (it->hit.kind == PopupHitTest::Kind::TaskConflictInspectSource || it->hit.kind == PopupHitTest::Kind::TaskConflictInspectDestination)
        {
            const auto* body = FindConflictBody(it->hit.taskId);
            if (body && body->visible && ! PointInRectF(body->viewport, x, y))
                continue;
        }
        return it->hit;
    }
    return {};
}

std::optional<FileOperationsPopupInternal::PopupMenuAnchor> FileOperationsPopupInternal::FileOperationsPopupState::ResolveButtonMenuAnchor(
    HWND hwnd, const PopupHitTest& hit, DxUi::ContextMenuRootVerticalPlacement placement) const noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return std::nullopt;
    }

    const PopupButton* match = nullptr;
    for (auto it = _buttons.rbegin(); it != _buttons.rend(); ++it)
    {
        const PopupButton& button = *it;
        if (button.hit.kind != hit.kind || button.hit.taskId != hit.taskId || button.hit.data != hit.data ||
            button.hit.promptGeneration != hit.promptGeneration)
        {
            continue;
        }

        if (button.bounds.right <= button.bounds.left || button.bounds.bottom <= button.bounds.top)
        {
            continue;
        }

        match = &button;
        break;
    }

    if (! match)
    {
        return std::nullopt;
    }

    RECT client{};
    static_cast<void>(GetClientRect(hwnd, &client));
    const float clientMidX = static_cast<float>(client.right - client.left) * 0.5f;
    const float buttonMidX = (match->bounds.left + match->bounds.right) * 0.5f;
    const bool alignEnd    = buttonMidX >= clientMidX;

    POINT anchor{
        static_cast<LONG>(std::lround(alignEnd ? match->bounds.right : match->bounds.left)),
        static_cast<LONG>(std::lround(placement == DxUi::ContextMenuRootVerticalPlacement::Above ? match->bounds.top : match->bounds.bottom)),
    };
    if (ClientToScreen(hwnd, &anchor) == FALSE)
    {
        return std::nullopt;
    }

    PopupMenuAnchor result{};
    result.screenPoint = anchor;
    result.sessionCallbacks.rootHorizontalAlignment =
        alignEnd ? DxUi::ContextMenuRootHorizontalAlignment::End : DxUi::ContextMenuRootHorizontalAlignment::Start;
    result.sessionCallbacks.rootVerticalPlacement = placement;
    result.sessionCallbacks.minRootWidthDip       = PixelsToDips(match->bounds.right - match->bounds.left, _dpi);
    return result;
}

void FileOperationsPopupInternal::FileOperationsPopupState::Invalidate(HWND hwnd) const noexcept
{
    if (hwnd)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

bool FileOperationsPopupInternal::FileOperationsPopupState::ConfirmCancelAll(HWND hwnd) noexcept
{
    if (! hwnd || ! folderWindow)
    {
        return false;
    }

    if (! hostLifetime.lock())
    {
        return true;
    }

    if (! HasCancelableTasks())
    {
        return true;
    }

    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_FILEOPS_CANCEL_ALL);
    const std::wstring message = LoadStringResource(nullptr, IDS_MSG_FILEOPS_CANCEL_ALL_POPUP);

    HostPromptRequest prompt{};
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_INFO;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = hwnd;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_OK;

    const FolderWindow::FileOperationPromptDispatchScope promptDispatch(*folderWindow);
    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
    {
        return false;
    }

    if (hostLifetime.lock() && fileOps)
    {
        // A second Cancel all adds to a cohort that has not resolved yet; it never drops it.
        for (const uint64_t taskId : fileOps->CancelAll())
        {
            if (std::ranges::find(_bulkCancelCohort, taskId) == _bulkCancelCohort.end())
                _bulkCancelCohort.push_back(taskId);
        }
    }

    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowQueueModeMenu(HWND hwnd) noexcept
{
    if (! hwnd || ! fileOps || ! folderWindow || ! hostLifetime.lock())
    {
        return;
    }

    constexpr int kQueueCommand    = 1;
    constexpr int kParallelCommand = 2;
    const bool queueMode           = fileOps->GetQueueNewTasks();
    std::array<DxUi::MenuFlyoutItem, 2u> items{};
    items[0].kind      = DxUi::MenuItemKind::Radio;
    items[0].text      = LoadStringResource(nullptr, IDS_FILEOPS_BTN_MODE_QUEUE);
    items[0].commandId = kQueueCommand;
    items[0].checked   = queueMode;
    items[1].kind      = DxUi::MenuItemKind::Radio;
    items[1].text      = LoadStringResource(nullptr, IDS_FILEOPS_BTN_MODE_PARALLEL);
    items[1].commandId = kParallelCommand;
    items[1].checked   = ! queueMode;

    DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor = ResolveButtonMenuAnchor(hwnd, PopupHitTest{PopupHitTest::Kind::FooterQueueMode, 0u, 0u}, DxUi::ContextMenuRootVerticalPlacement::Above);
    const POINT point = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(DxUi::ContextMenu::ShowAsync(hwnd,
                                                   point,
                                                   items,
                                                   MakeAppThemeDxPalette(folderWindow->GetTheme()),
                                                   [this, hwnd](std::optional<int> chosen) noexcept
    {
        if (! chosen.has_value() || ! hwnd || IsWindow(hwnd) == FALSE || ! hostLifetime.lock() || ! fileOps)
        {
            return;
        }

        if (chosen.value() == kQueueCommand || chosen.value() == kParallelCommand)
        {
            fileOps->ApplyQueueMode(chosen.value() == kQueueCommand);
            Invalidate(hwnd);
        }
    },
                                                   sessionCallbacks));
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowFooterOptionsMenu(HWND hwnd) noexcept
{
    if (! hwnd || ! fileOps || ! folderWindow || ! hostLifetime.lock())
    {
        return;
    }

    constexpr int kAutoDismissCommand = 1;
    constexpr int kCompactCommand     = 2;
    constexpr int kExpandedCommand    = 3;
    const bool autoDismiss            = fileOps->GetAutoDismissSuccess();
    const bool compactDensity         = fileOps->GetPopupCompactDensity();

    std::array<DxUi::MenuFlyoutItem, 4u> items{};
    items[0].kind      = DxUi::MenuItemKind::Toggle;
    items[0].text      = LoadStringResource(nullptr, autoDismiss ? IDS_FILEOPS_CHECK_AUTODISMISS_ON : IDS_FILEOPS_CHECK_AUTODISMISS_OFF);
    items[0].commandId = kAutoDismissCommand;
    items[0].checked   = autoDismiss;
    items[1].kind      = DxUi::MenuItemKind::Separator;
    items[2].kind      = DxUi::MenuItemKind::Radio;
    items[2].text      = LoadStringResource(nullptr, IDS_FILEOPS_BTN_DENSITY_COMPACT);
    items[2].commandId = kCompactCommand;
    items[2].checked   = compactDensity;
    items[3].kind      = DxUi::MenuItemKind::Radio;
    items[3].text      = LoadStringResource(nullptr, IDS_FILEOPS_BTN_DENSITY_EXPANDED);
    items[3].commandId = kExpandedCommand;
    items[3].checked   = ! compactDensity;

    DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor = ResolveButtonMenuAnchor(hwnd, PopupHitTest{PopupHitTest::Kind::FooterOptions, 0u, 0u}, DxUi::ContextMenuRootVerticalPlacement::Above);
    const POINT point = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(DxUi::ContextMenu::ShowAsync(hwnd,
                                                   point,
                                                   items,
                                                   MakeAppThemeDxPalette(folderWindow->GetTheme()),
                                                   [this, hwnd](std::optional<int> chosen) noexcept
    {
        if (! chosen.has_value() || ! hwnd || IsWindow(hwnd) == FALSE || ! hostLifetime.lock() || ! fileOps)
        {
            return;
        }

        if (chosen.value() == kAutoDismissCommand)
        {
            fileOps->SetAutoDismissSuccess(! fileOps->GetAutoDismissSuccess());
        }
        else if (chosen.value() == kCompactCommand || chosen.value() == kExpandedCommand)
        {
            fileOps->SetPopupCompactDensity(chosen.value() == kCompactCommand);
            _maxAutoSizedWindowHeight   = 0;
            _lastAutoSizedContentHeight = -1.0f;
        }
        Invalidate(hwnd);
    },
                                                   sessionCallbacks));
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowSpeedLimitMenu(HWND hwnd, uint64_t taskId) noexcept
{
    if (! IsTaskDecisionAvailable(taskId))
        return;
    if (! hwnd || ! fileOps || ! folderWindow)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    FolderWindow::FileOperationState::Task* task = fileOps->FindTask(taskId);
    const TaskSnapshot* fixture                  = nullptr;
#ifdef ENABLE_TESTS
    if (debugVisualScenario)
    {
        const auto found = std::find_if(
            debugVisualScenario->begin(), debugVisualScenario->end(), [taskId](const TaskSnapshot& value) noexcept { return value.taskId == taskId; });
        if (found != debugVisualScenario->end())
            fixture = &*found;
    }
#endif
    if (! task && ! fixture)
    {
        return;
    }

    const FileSystemOperation operation = fixture ? fixture->operation : task->GetOperation();
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return;
    }

    const uint64_t currentLimit = fixture ? fixture->desiredSpeedLimitBytesPerSecond : task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

    constexpr UINT kCmdUnlimited  = 1u;
    constexpr UINT kCmdCustom     = 2u;
    constexpr UINT kCmdPresetBase = 10u;

    static constexpr std::array<uint64_t, 6> kPresets = {{
        1ull * 1024ull * 1024ull,
        5ull * 1024ull * 1024ull,
        10ull * 1024ull * 1024ull,
        50ull * 1024ull * 1024ull,
        100ull * 1024ull * 1024ull,
        1ull * 1024ull * 1024ull * 1024ull,
    }};

    std::vector<DxUi::MenuFlyoutItem> items;
    items.reserve(kPresets.size() + 4u);

    auto appendRadioItem = [&](UINT commandId, std::wstring text, bool checked) noexcept
    {
        DxUi::MenuFlyoutItem item{};
        item.kind      = DxUi::MenuItemKind::Radio;
        item.text      = std::move(text);
        item.commandId = static_cast<int>(commandId);
        item.checked   = checked;
        items.push_back(std::move(item));
    };

    appendRadioItem(kCmdUnlimited, LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_MENU_UNLIMITED), currentLimit == 0);

    DxUi::MenuFlyoutItem separator{};
    separator.kind = DxUi::MenuItemKind::Separator;
    items.push_back(separator);

    for (size_t i = 0; i < kPresets.size(); ++i)
    {
        const uint64_t bytesPerSecond = kPresets[i];
        appendRadioItem(kCmdPresetBase + static_cast<UINT>(i),
                        FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_MENU_BYTES, FormatBytesCompact(bytesPerSecond)),
                        currentLimit == bytesPerSecond);
    }

    items.push_back(separator);

    DxUi::MenuFlyoutItem customItem{};
    customItem.text      = LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_MENU_CUSTOM);
    customItem.commandId = static_cast<int>(kCmdCustom);
    items.push_back(std::move(customItem));

    DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor =
        ResolveButtonMenuAnchor(hwnd, PopupHitTest{PopupHitTest::Kind::TaskSpeedLimit, taskId, 0u}, DxUi::ContextMenuRootVerticalPlacement::Below);
    const POINT pt = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(DxUi::ContextMenu::ShowAsync(hwnd,
                                                   pt,
                                                   items,
                                                   MakeAppThemeDxPalette(folderWindow->GetTheme()),
                                                   [this, hwnd, taskId](std::optional<int> chosenOpt) noexcept
    {
        if (! chosenOpt.has_value())
        {
            return;
        }

        if (! hwnd || IsWindow(hwnd) == FALSE || ! hostLifetime.lock() || ! fileOps || ! folderWindow)
        {
            return;
        }

        FolderWindow::FileOperationState::Task* selectedTask = fileOps->FindTask(taskId);
        if (! selectedTask || selectedTask->CancelRequestedDurable())
        {
            return;
        }

        const FileSystemOperation selectedOperation = selectedTask->GetOperation();
        if (selectedOperation != FILESYSTEM_COPY && selectedOperation != FILESYSTEM_MOVE)
        {
            return;
        }

        const UINT chosen      = static_cast<UINT>(chosenOpt.value());
        uint64_t newLimit      = selectedTask->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        const uint64_t current = newLimit;
        if (chosen == kCmdUnlimited)
        {
            newLimit = 0;
        }
        else if (chosen >= kCmdPresetBase && chosen < (kCmdPresetBase + static_cast<UINT>(kPresets.size())))
        {
            const size_t index = static_cast<size_t>(chosen - kCmdPresetBase);
            newLimit           = kPresets[index];
        }
        else if (chosen == kCmdCustom)
        {
            // The custom prompt pumps nested thread messages; keep FileOperationState alive across
            // that pump and re-find the task by id afterwards (same fence as the deferred path).
            const FolderWindow::FileOperationPromptDispatchScope promptDispatch(*folderWindow);
            const auto promptResult = ShowCustomSpeedLimitPrompt(hwnd, folderWindow->GetTheme(), current);
            if (! promptResult.has_value())
            {
                return;
            }

            newLimit     = promptResult.value();
            selectedTask = fileOps->FindTask(taskId);
            if (! selectedTask || selectedTask->CancelRequestedDurable())
            {
                return;
            }

            const FileSystemOperation postPromptOperation = selectedTask->GetOperation();
            if (postPromptOperation != FILESYSTEM_COPY && postPromptOperation != FILESYSTEM_MOVE)
            {
                return;
            }
        }
        else
        {
            return;
        }

        selectedTask->SetDesiredSpeedLimit(newLimit);
        Invalidate(hwnd);
    },
                                                   sessionCallbacks));
}

bool FileOperationsPopupInternal::FileOperationsPopupState::ShowCustomSpeedLimitPromptForTask(HWND hwnd, uint64_t requestedTaskId) noexcept
{
    if (! fileOps || ! folderWindow || ! hostLifetime.lock())
    {
        return false;
    }

    uint64_t resolvedTaskId                      = requestedTaskId;
    FolderWindow::FileOperationState::Task* task = resolvedTaskId != 0 ? fileOps->FindTask(resolvedTaskId) : nullptr;
    if (! task)
    {
        const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
        const auto isActiveFileOperation         = [](const TaskSnapshot& candidate) noexcept
        { return candidate.kind == TaskSnapshot::Kind::FileOperation && ! candidate.finished && ! candidate.cancellationRequested && candidate.taskId != 0; };

        const auto activeIt = std::find_if(snapshot.begin(), snapshot.end(), isActiveFileOperation);
        if (activeIt != snapshot.end())
        {
            resolvedTaskId = activeIt->taskId;
            task           = fileOps->FindTask(resolvedTaskId);
        }
    }

    if (! task || task->CancelRequestedDurable())
    {
        return false;
    }

    const uint64_t taskId       = resolvedTaskId;
    const uint64_t currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    const FolderWindow::FileOperationPromptDispatchScope promptDispatch(*folderWindow);
    const auto promptResult = ShowCustomSpeedLimitPrompt(hwnd, folderWindow->GetTheme(), currentLimit);
    if (promptResult.has_value() && hostLifetime.lock())
    {
        if (FolderWindow::FileOperationState::Task* const currentTask = fileOps->FindTask(taskId))
        {
            if (! currentTask->CancelRequestedDurable())
                currentTask->SetDesiredSpeedLimit(promptResult.value());
        }
    }
    if (IsWindow(hwnd) != FALSE)
    {
        Invalidate(hwnd);
    }
    return true;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::SubmitCompletedOverflowAction(HWND hwnd,
                                                                                          uint64_t taskId,
                                                                                          uint32_t action,
                                                                                          bool openExportAfterWrite) noexcept
{
    if (! fileOps)
    {
        return false;
    }

    bool handled = false;
    if (action == kCompletedOverflowActionShowLog)
    {
        handled = fileOps->OpenDiagnosticsLogForTask(taskId);
    }
    else if (action == kCompletedOverflowActionExportIssues)
    {
        handled = fileOps->ExportTaskIssuesReport(taskId, nullptr, openExportAfterWrite);
        if (! handled && openExportAfterWrite && hwnd && IsWindow(hwnd))
        {
            const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            const std::wstring message = LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_EXPORT_FAILED);
            HostAlertRequest alert{};
            alert.sizeBytes    = sizeof(alert);
            alert.scope        = HOST_ALERT_SCOPE_WINDOW;
            alert.modality     = HOST_ALERT_MODELESS;
            alert.severity     = HOST_ALERT_ERROR;
            alert.targetWindow = hwnd;
            alert.title        = title.c_str();
            alert.message      = message.c_str();
            alert.closable     = TRUE;
            static_cast<void>(HostShowAlert(alert));
        }
    }
    else if (action == kCompletedOverflowActionFailedItems)
    {
        handled = fileOps->OpenIssuesForTask(taskId);
    }
    else if (action == kCompletedOverflowActionOpenSource)
    {
        handled = fileOps->OpenCompletedTaskSource(taskId);
    }
    else if (action == kCompletedOverflowActionSelectRetained)
    {
        handled = fileOps->SelectCompletedTaskRetainedSources(taskId);
    }
    else if (action == kCompletedOverflowActionCutRetainedAgain)
    {
        handled = fileOps->CutCompletedTaskRetainedSourcesAgain(taskId);
    }
    else if (action == kCompletedOverflowActionOpenDestination || action == kCompletedOverflowActionRevealDestination)
    {
        if (! folderWindow)
        {
            return false;
        }

        const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
        const auto taskIt                        = std::find_if(snapshot.begin(), snapshot.end(), [taskId](const TaskSnapshot& task) noexcept {
            return task.kind == TaskSnapshot::Kind::FileOperation && task.taskId == taskId && task.finished;
        });
        if (taskIt == snapshot.end())
        {
            return false;
        }

        if (action == kCompletedOverflowActionOpenDestination && CompletedTaskCanUseDestinationActions(*taskIt))
        {
            handled = SUCCEEDED(folderWindow->ExecuteInPaneLocation(taskIt->destinationPane.value(),
                                                                    taskIt->destinationPluginId,
                                                                    taskIt->destinationPluginShortId,
                                                                    taskIt->destinationInstanceContext,
                                                                    taskIt->destinationFolder,
                                                                    {},
                                                                    0u,
                                                                    true));
        }
        else if (action == kCompletedOverflowActionRevealDestination)
        {
            const std::optional<CompletedTaskRevealLocation> revealLocation = ResolveCompletedTaskRevealLocation(*taskIt);
            if (revealLocation.has_value() && taskIt->destinationPane.has_value())
            {
                handled = SUCCEEDED(folderWindow->ExecuteInPaneLocation(taskIt->destinationPane.value(),
                                                                        taskIt->destinationPluginId,
                                                                        taskIt->destinationPluginShortId,
                                                                        taskIt->destinationInstanceContext,
                                                                        revealLocation->folder,
                                                                        revealLocation->leaf,
                                                                        0u,
                                                                        true));
            }
        }
    }

    if (handled)
    {
        Invalidate(hwnd);
    }
    return handled;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowCompletedOverflowMenu(HWND hwnd, uint64_t taskId) noexcept
{
    if (! hwnd || ! fileOps || ! folderWindow)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
    const auto taskIt                        = std::find_if(snapshot.begin(), snapshot.end(), [taskId](const TaskSnapshot& task) noexcept {
        return task.kind == TaskSnapshot::Kind::FileOperation && task.taskId == taskId && task.finished;
    });
    if (taskIt == snapshot.end() || ! CompletedTaskHasOverflowActions(*taskIt))
    {
        return;
    }

    std::vector<DxUi::MenuFlyoutItem> items;
    items.reserve(8u);

    // I35 R5.3: Show log and Export issues follow any retained diagnostic (notes included), never the
    // informational-notes display setting; the export always contains everything.
    const bool hasDiagnostics = ! taskIt->interruptedMoveNotice && (taskIt->warningCount > 0 || taskIt->errorCount > 0 || taskIt->hasRetainedDiagnostics);

    if (ResolveCompletedTaskRevealLocation(*taskIt).has_value())
    {
        DxUi::MenuFlyoutItem revealItem{};
        revealItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_REVEAL_ITEM);
        revealItem.commandId = static_cast<int>(kCompletedOverflowActionRevealDestination);
        items.push_back(std::move(revealItem));
    }

    if (taskIt->hasSourceActionPaths)
    {
        DxUi::MenuFlyoutItem openSourceItem{};
        openSourceItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_OPEN_SOURCE);
        openSourceItem.commandId = static_cast<int>(kCompletedOverflowActionOpenSource);
        items.push_back(std::move(openSourceItem));
    }

    if (taskIt->exactRetainedSourceActionAvailable)
    {
        DxUi::MenuFlyoutItem selectRetainedItem{};
        selectRetainedItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_SELECT_RETAINED);
        selectRetainedItem.commandId = static_cast<int>(kCompletedOverflowActionSelectRetained);
        items.push_back(std::move(selectRetainedItem));

        DxUi::MenuFlyoutItem cutRetainedItem{};
        cutRetainedItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_CUT_RETAINED_AGAIN);
        cutRetainedItem.commandId = static_cast<int>(kCompletedOverflowActionCutRetainedAgain);
        items.push_back(std::move(cutRetainedItem));
    }

    if (CompletedTaskCanUseDestinationActions(*taskIt))
    {
        DxUi::MenuFlyoutItem openDestinationItem{};
        openDestinationItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_OPEN_DESTINATION);
        openDestinationItem.commandId = static_cast<int>(kCompletedOverflowActionOpenDestination);
        items.push_back(std::move(openDestinationItem));
    }

    if (hasDiagnostics)
    {
        DxUi::MenuFlyoutItem showLogItem{};
        showLogItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_SHOW_LOG);
        showLogItem.commandId = static_cast<int>(kCompletedOverflowActionShowLog);
        items.push_back(std::move(showLogItem));

        DxUi::MenuFlyoutItem exportIssuesItem{};
        exportIssuesItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_EXPORT_ISSUES);
        exportIssuesItem.commandId = static_cast<int>(kCompletedOverflowActionExportIssues);
        items.push_back(std::move(exportIssuesItem));
    }

    DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor =
        ResolveButtonMenuAnchor(hwnd, PopupHitTest{PopupHitTest::Kind::TaskCompletedMore, taskId, 2u}, DxUi::ContextMenuRootVerticalPlacement::Above);
    const POINT pt = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(DxUi::ContextMenu::ShowAsync(hwnd,
                                                   pt,
                                                   items,
                                                   MakeAppThemeDxPalette(folderWindow->GetTheme()),
                                                   [this, hwnd, taskId](std::optional<int> chosenOpt) noexcept
    {
        if (! chosenOpt.has_value())
        {
            return;
        }

        if (! hwnd || IsWindow(hwnd) == FALSE || ! hostLifetime.lock())
        {
            return;
        }

        const UINT chosen = static_cast<UINT>(chosenOpt.value());
        if (chosen == kCompletedOverflowActionShowLog || chosen == kCompletedOverflowActionExportIssues || chosen == kCompletedOverflowActionFailedItems ||
            chosen == kCompletedOverflowActionOpenDestination || chosen == kCompletedOverflowActionRevealDestination ||
            chosen == kCompletedOverflowActionOpenSource || chosen == kCompletedOverflowActionSelectRetained ||
            chosen == kCompletedOverflowActionCutRetainedAgain)
        {
            static_cast<void>(SubmitCompletedOverflowAction(hwnd, taskId, chosen, true));
        }
    },
                                                   sessionCallbacks));
}

bool FileOperationsPopupInternal::FileOperationsPopupState::SubmitConflictOverflowAction(HWND hwnd,
                                                                                         uint64_t taskId,
                                                                                         uint32_t rawAction,
                                                                                         uint64_t generation) noexcept
{
    auto* task = fileOps ? fileOps->FindTask(taskId) : nullptr;
    if (! task ||
        ! task->SubmitConflictDecision(generation, static_cast<ConflictAction>(rawAction), std::nullopt, FileOperations::ConflictActionPlacement::Overflow))
        return false;
    Invalidate(hwnd);
    return true;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::SubmitConflictPrimaryAction(HWND hwnd,
                                                                                        uint64_t taskId,
                                                                                        uint32_t rawAction,
                                                                                        uint64_t generation) noexcept
{
    auto* task = fileOps ? fileOps->FindTask(taskId) : nullptr;
    if (! task ||
        ! task->SubmitConflictDecision(generation, static_cast<ConflictAction>(rawAction), std::nullopt, FileOperations::ConflictActionPlacement::Primary))
        return false;
    Invalidate(hwnd);
    return true;
}
void FileOperationsPopupInternal::FileOperationsPopupState::ShowConflictOverflowMenu(HWND hwnd, uint64_t taskId, uint64_t generation) noexcept
{
    if (! hwnd || ! fileOps || ! folderWindow)
    {
        return;
    }

    if (! hostLifetime.lock())
    {
        return;
    }

    FolderWindow::FileOperationState::Task* task = fileOps->FindTask(taskId);
    const TaskSnapshot* fixture                  = nullptr;
#ifdef ENABLE_TESTS
    if (debugVisualScenario)
    {
        const auto found = std::find_if(
            debugVisualScenario->begin(), debugVisualScenario->end(), [taskId](const TaskSnapshot& value) noexcept { return value.taskId == taskId; });
        if (found != debugVisualScenario->end())
            fixture = &*found;
    }
#endif
    if (! task && ! fixture)
    {
        return;
    }

    TaskSnapshot::ConflictPromptSnapshot conflict{};
    if (fixture)
    {
        conflict = fixture->conflict;
    }
    else
    {
        std::scoped_lock lock(task->_conflictArbiter.mutex);
        if (generation == 0u || task->_conflictArbiter.prompt.generation != generation || ! task->_conflictArbiter.prompt.active ||
            ! task->_conflictArbiter.prompt.buttonsPublishable)
        {
            return;
        }

        conflict.active          = true;
        conflict.bucket          = static_cast<uint8_t>(task->_conflictArbiter.prompt.bucket);
        conflict.status          = task->_conflictArbiter.prompt.status;
        conflict.sourcePath      = task->_conflictArbiter.prompt.sourcePath;
        conflict.destinationPath = task->_conflictArbiter.prompt.destinationPath;
        CopyPublishedConflictPolicy(task->_conflictArbiter.prompt, conflict);
        conflict.applyToAllChecked = task->_conflictArbiter.prompt.applyToAllChecked;
        conflict.retryFailed       = task->_conflictArbiter.prompt.retryFailed;
        conflict.attemptCount      = task->_conflictArbiter.prompt.attemptCount;
    }

    if (conflict.overflowActionCount == 0u)
    {
        return;
    }

    constexpr UINT kCmdOverflowBase = 1u;
    std::vector<DxUi::MenuFlyoutItem> items;
    items.reserve(std::min(conflict.overflowActionCount, conflict.overflowActions.size()));
    std::vector<uint32_t> overflowActions;
    overflowActions.reserve(std::min(conflict.overflowActionCount, conflict.overflowActions.size()));
    for (size_t i = 0; i < conflict.overflowActionCount && i < conflict.overflowActions.size(); ++i)
    {
        DxUi::MenuFlyoutItem item{};
        const ConflictAction action = ConflictActionFromRaw(conflict.overflowActions[i]);
        item.text      = ConflictActionText(action, static_cast<ConflictBucket>(conflict.bucket), fixture ? fixture->operation : task->GetOperation());
        item.commandId = static_cast<int>(kCmdOverflowBase + static_cast<UINT>(i));
        items.push_back(std::move(item));
        overflowActions.push_back(static_cast<uint32_t>(RawConflictAction(action)));
    }

    DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const uint32_t overflowData = static_cast<uint32_t>(std::min<size_t>(conflict.overflowActionCount, std::numeric_limits<uint32_t>::max()));
    const auto anchor           = ResolveButtonMenuAnchor(
        hwnd, PopupHitTest{PopupHitTest::Kind::TaskConflictMore, taskId, overflowData, generation}, DxUi::ContextMenuRootVerticalPlacement::Below);
    const POINT pt = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(
        DxUi::ContextMenu::ShowAsync(hwnd,
                                     pt,
                                     items,
                                     MakeAppThemeDxPalette(folderWindow->GetTheme()),
                                     [this, hwnd, taskId, generation, overflowActions = std::move(overflowActions)](std::optional<int> chosenOpt) noexcept
    {
        if (! chosenOpt.has_value())
        {
            return;
        }

        if (! hwnd || IsWindow(hwnd) == FALSE || ! hostLifetime.lock())
        {
            return;
        }

        const UINT chosen = static_cast<UINT>(chosenOpt.value());
        if (chosen < kCmdOverflowBase || chosen >= kCmdOverflowBase + static_cast<UINT>(overflowActions.size()))
        {
            return;
        }

        const size_t index = static_cast<size_t>(chosen - kCmdOverflowBase);
        static_cast<void>(SubmitConflictOverflowAction(hwnd, taskId, overflowActions[index], generation));
    },
                                     sessionCallbacks));
}

// C4: graphics-independent failure surface. One native text, Cancel all, and Close; no Direct2D,
// DirectWrite, DxUi host, hit-testing, task graph, queue/speed menu, or conflict action. Waiting tasks
// keep waiting and running tasks keep running unless Cancel all is confirmed.
namespace
{
constexpr INT_PTR kFailureSurfaceTextId              = 0x4C01;
constexpr INT_PTR kFailureSurfaceCancelAllId         = 0x4C02;
constexpr INT_PTR kFailureSurfaceCloseId             = 0x4C03;
constexpr wchar_t kFailureSurfaceButtonWndProcProp[] = L"RedSalamander.FileOperations.FailureSurfaceButton.WndProc";

[[nodiscard]] HWND CreateFailureSurfaceChild(
    unsigned int ordinal, const wchar_t* className, const wchar_t* text, DWORD style, HWND parent, INT_PTR controlId, HINSTANCE instance) noexcept
{
#ifdef ENABLE_TESTS
    if (ShouldFailFailureSurfaceChildCreation(ordinal))
        return nullptr;
#else
    static_cast<void>(ordinal);
#endif
    return CreateWindowExW(0, className, text, style, 0, 0, 0, 0, parent, reinterpret_cast<HMENU>(controlId), instance, nullptr);
}
} // namespace

LRESULT CALLBACK FileOperationsPopupInternal::FileOperationsPopupState::FailureSurfaceButtonProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    const WNDPROC original = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kFailureSurfaceButtonWndProcProp);
    if (msg == WM_NCDESTROY)
    {
        const HWND popup = GetParent(hwnd);
        auto* state      = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA)) : nullptr;
        if (state != nullptr)
            state->_failureCancelAllEnterArmed = false;
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kFailureSurfaceButtonWndProcProp);
    }
    else if (msg == WM_KEYDOWN || msg == WM_KEYUP || msg == WM_KILLFOCUS)
    {
        const HWND popup = GetParent(hwnd);
        auto* state      = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA)) : nullptr;
        if (state != nullptr && msg == WM_KILLFOCUS)
        {
            state->_failureCancelAllEnterArmed = false;
        }
        else if (state != nullptr && ((msg == WM_KEYDOWN && state->HandleFailureSurfaceKeyDown(popup, hwnd, static_cast<UINT>(wp), lp)) ||
                                      (msg == WM_KEYUP && state->HandleFailureSurfaceKeyUp(hwnd, static_cast<UINT>(wp)))))
        {
            return 0;
        }
    }
    return original != nullptr ? RedSalamander::Win32Callback::CallWindowProcNoThrow(original, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
}

bool FileOperationsPopupInternal::FileOperationsPopupState::EnterFailureSurface(HWND hwnd) noexcept
{
    if (_failureSurfaceReady)
    {
        return true;
    }
    if (_failureSurfaceInitializing)
    {
#ifdef ENABLE_TESTS
        ++_debugFailureSurfaceReentrantEntryCount;
#endif
        // Native layout/theme changes can synchronously paint while this transaction
        // is detaching the retained host. Only the outer entry may publish children.
        return false;
    }
    _failureSurfaceInitializing     = true;
    _failureCancelAllEnterArmed     = false;
    const auto finishInitialization = wil::scope_exit([this]() noexcept { _failureSurfaceInitializing = false; });

    if (! _failureSurface)
    {
        _failureSurface      = true;
        _failureSurfaceReady = false;

        _pathDetailId.clear();
        _pathDetailReturnId.clear();
        _pathDetailLogicalFocusId.clear();
        _pathDetailTaskId     = 0u;
        _pathDetailNeedsFocus = false;
        _pathDetailPanel      = nullptr;
        _pathDetailTitle      = nullptr;
        _pathDetailField      = nullptr;
        _pathDetailCopy       = nullptr;
        _pathDetailClose      = nullptr;
#ifdef ENABLE_TESTS
        _debugRetiredControlQuarantine.clear();
#endif
        _controlsHost.Detach();
        _controlsRoot = nullptr;
        _controlsBody = nullptr;
        _hostedButtons.clear();
        _hostedButtonHits.clear();
        _hostedProgressBars.clear();
        _hostedProgressAnimations.clear();
        _hostedGraphs.clear();
        _hostedGraphStallBands.clear();
        _hostedProgressVerificationMarks.clear();
        _hostedTooltipRegions.clear();
        _hostedTooltipComparisonIds.clear();
        _hostedTooltipTaskIds.clear();
        _hostedButtonParents.clear();
        _hostedTooltipParents.clear();
        _retiredHostedControls.clear();
        _hostedAccessibleTextTicks.clear();
        _hostedAccessibleTextPending = false;
        _conflictBodies.clear();
        _hostedConflictBodyIds.clear();
        _hostedConflictBodyPanels.clear();
        _buttons.clear();
        _completedGroupClearTaskIds.clear();
        // The failure surface shows no decision, so Escape must not keep one bound.
        _conflictEscapeTaskId     = 0u;
        _conflictEscapeAction     = 0u;
        _conflictEscapeGeneration = 0u;
        DiscardDeviceResources();
        ShowScrollBar(hwnd, SB_VERT, FALSE);
        _scrollBarVisible = false;
        // Q6: the popup timer re-attempts the host with backoff (reopening retries at once).
        _framePrepareFailureCount     = 0u;
        _framePrepareFirstFailureTick = 0u;
        _framePrepareRetryTick        = 0u;
        const uint32_t shift          = std::min(_hostReattachAttemptCount, 6u);
        _hostReattachRetryTick = GetTickCount64() + std::min(kFileOperationsPopupHostReattachRetryBaseMs << shift, kFileOperationsPopupHostReattachRetryMaxMs);
    }

    const HINSTANCE instance     = reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(hwnd, GWLP_HINSTANCE));
    const std::wstring cancelAll = LoadStringResource(nullptr, IDS_FILEOPS_BTN_CANCEL_ALL);
    const std::wstring close     = LoadStringResource(nullptr, IDS_FILEOPS_BTN_CLOSE);
    wil::unique_hwnd failureText(
        CreateFailureSurfaceChild(1u, L"STATIC", L"", WS_CHILD | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL, hwnd, kFailureSurfaceTextId, instance));
    wil::unique_hwnd failureCancelAll(
        CreateFailureSurfaceChild(2u, L"BUTTON", cancelAll.c_str(), WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON, hwnd, kFailureSurfaceCancelAllId, instance));
    wil::unique_hwnd failureClose(
        CreateFailureSurfaceChild(3u, L"BUTTON", close.c_str(), WS_CHILD | WS_TABSTOP | BS_DEFPUSHBUTTON, hwnd, kFailureSurfaceCloseId, instance));
    const auto installButtonHook = [](HWND button, unsigned int ordinal) noexcept
    {
#ifdef ENABLE_TESTS
        if (ShouldFailFailureSurfaceButtonSetup(ordinal))
            return false;
#else
        static_cast<void>(ordinal);
#endif
        return RedSalamander::Win32Callback::InstallWndProcHook(button, kFailureSurfaceButtonWndProcProp, &FailureSurfaceButtonProc);
    };
    if (! failureText || ! failureCancelAll || ! failureClose || ! installButtonHook(failureCancelAll.get(), 1u) || ! installButtonHook(failureClose.get(), 2u))
    {
        if (! _failureSurfaceFailureReported)
        {
            // The existing host error surface carries the explanation; Cancel all stays reachable
            // from the main window's File Operations command while the native surface retries later.
            Debug::Error(L"File operations popup could not create and initialize its failure surface.");
            const std::wstring message = LoadStringResource(nullptr, IDS_FILEOPS_FAILURE_SURFACE_TEXT);
            HostAlertRequest alert{};
            alert.sizeBytes    = sizeof(alert);
            alert.scope        = HOST_ALERT_SCOPE_WINDOW;
            alert.modality     = HOST_ALERT_MODELESS;
            alert.severity     = HOST_ALERT_ERROR;
            alert.targetWindow = folderWindow ? folderWindow->GetHwnd() : hwnd;
            alert.title        = nullptr;
            alert.message      = message.c_str();
            alert.closable     = TRUE;
            static_cast<void>(HostShowAlert(alert));
            _failureSurfaceFailureReported = true;
        }
        return false;
    }

    _failureText      = failureText.release();
    _failureCancelAll = failureCancelAll.release();
    _failureClose     = failureClose.release();
    LayoutFailureSurface(hwnd);
    RefreshFailureSurfaceText();
    _failureSurfaceReady = true;
    for (const HWND child : {_failureText, _failureCancelAll, _failureClose})
    {
        ShowWindow(child, SW_SHOWNA);
    }
    if (GetActiveWindow() == hwnd && GetForegroundWindow() == hwnd)
    {
        SetFocus(_failureClose);
    }
    InvalidateRect(hwnd, nullptr, TRUE);
    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::LayoutFailureSurface(HWND hwnd) noexcept
{
    if (! _failureSurface || ! _failureText || ! _failureCancelAll || ! _failureClose)
    {
        return;
    }
    if (! _failureFont || _failureFontDpi != _dpi)
    {
        NONCLIENTMETRICSW metrics{};
        metrics.cbSize = sizeof(metrics);
        if (SystemParametersInfoForDpi(SPI_GETNONCLIENTMETRICS, sizeof(metrics), &metrics, 0, _dpi) != FALSE)
        {
            _failureFont.reset(CreateFontIndirectW(&metrics.lfMessageFont));
        }
        _failureFontDpi  = _dpi;
        const HFONT font = _failureFont ? _failureFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        for (const HWND child : {_failureText, _failureCancelAll, _failureClose})
        {
            SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
        }
    }

    RECT client{};
    static_cast<void>(GetClientRect(hwnd, &client));
    const int margin     = static_cast<int>(DipsToPixels(12.0f, _dpi));
    const int gap        = static_cast<int>(DipsToPixels(8.0f, _dpi));
    const int buttonW    = static_cast<int>(DipsToPixels(96.0f, _dpi));
    const int buttonH    = static_cast<int>(DipsToPixels(28.0f, _dpi));
    const int width      = static_cast<int>(std::max(0L, client.right - client.left));
    const int height     = static_cast<int>(std::max(0L, client.bottom - client.top));
    const int buttonsTop = std::max(margin, height - margin - buttonH);
    const int closeLeft  = std::max(margin, width - margin - buttonW);
    const int cancelLeft = std::max(margin, closeLeft - gap - buttonW);
    SetWindowPos(_failureText, nullptr, margin, margin, std::max(0, width - 2 * margin), std::max(0, buttonsTop - gap - margin), SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(_failureCancelAll, nullptr, cancelLeft, buttonsTop, buttonW, buttonH, SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(_failureClose, nullptr, closeLeft, buttonsTop, buttonW, buttonH, SWP_NOZORDER | SWP_NOACTIVATE);
}

void FileOperationsPopupInternal::FileOperationsPopupState::RefreshFailureSurfaceText() noexcept
{
    if (! _failureSurface || ! _failureText)
    {
        return;
    }
    const std::vector<TaskSnapshot> snapshot        = BuildSnapshot();
    const GlobalFileOperationsStatusSummary summary = BuildGlobalStatusSummary(snapshot, &_rates);
    const bool canCancel                            = HasCancelableTasks();
    if (_failureCancelAll && (IsWindowEnabled(_failureCancelAll) != FALSE) != canCancel)
    {
        // popup-c#4: disabling the focused window leaves no keyboard focus; move it to Close first.
        if (! canCancel && _failureClose && GetFocus() == _failureCancelAll)
            SetFocus(_failureClose);
        EnableWindow(_failureCancelAll, canCancel ? TRUE : FALSE);
    }
    if (! canCancel)
        _failureCancelAllEnterArmed = false;
    std::wstring text              = LoadStringResource(nullptr, IDS_FILEOPS_FAILURE_SURFACE_TEXT);
    const std::wstring summaryText = FormatGlobalStatusSummaryText(summary);
    if (! summaryText.empty())
    {
        text.append(L"\r\n\r\n").append(summaryText);
    }
    // One line per operation: what it is and where it stands, without the card machinery.
    const ULONGLONG nowTick = GetTickCount64();
    size_t listed           = 0u;
    for (const TaskSnapshot& task : snapshot)
    {
        if (task.kind != TaskSnapshot::Kind::FileOperation)
        {
            continue;
        }
        if (listed == 8u)
        {
            text.append(L"\r\n…");
            break;
        }
        UINT operationId = IDS_FILEOP_OPERATION_COPY;
        switch (task.operation)
        {
            case FILESYSTEM_MOVE: operationId = IDS_FILEOP_OPERATION_MOVE; break;
            case FILESYSTEM_DELETE: operationId = IDS_FILEOP_OPERATION_DELETE; break;
            case FILESYSTEM_RENAME: operationId = IDS_FILEOP_OPERATION_RENAME; break;
            case FILESYSTEM_COPY:
            case FILESYSTEM_CREATE_DIRECTORY:
            default: break;
        }
        text.append(L"\r\n")
            .append(LoadStringResource(nullptr, operationId))
            .append(L": ")
            .append(StatusTextForTask(task, ResolveTaskStatusKind(task), nowTick));
        ++listed;
    }
    if (text != _failureTextValue)
    {
        _failureTextValue = std::move(text);
        SetWindowTextW(_failureText, _failureTextValue.c_str());
    }
}

bool FileOperationsPopupInternal::FileOperationsPopupState::HandleFailureSurfaceKeyDown(HWND hwnd, HWND focused, UINT virtualKey, LPARAM keyData) noexcept
{
    if (! _failureSurface || ! _failureCancelAll || ! _failureClose)
    {
        return false;
    }
    constexpr LPARAM previousKeyState = LPARAM{1} << 30;
    if ((keyData & previousKeyState) != 0 && (virtualKey == VK_ESCAPE || virtualKey == VK_RETURN))
    {
        return true;
    }
    if (virtualKey == VK_ESCAPE)
    {
        static_cast<void>(OnClose(hwnd));
        return true;
    }
    if (virtualKey == VK_TAB)
    {
#ifdef ENABLE_TESTS
        ++_debugFailureSurfaceTabCount;
#endif
        // popup-c#4: a disabled Cancel all is not a tab stop.
        const HWND next = focused == _failureClose && IsWindowEnabled(_failureCancelAll) != FALSE ? _failureCancelAll : _failureClose;
        if (next != focused)
            SetFocus(next);
        return true;
    }
    if (virtualKey == VK_RETURN)
    {
        // Release Enter before opening consent: otherwise the held key's repeat
        // can reach the newly opened dialog's default OK button.
        _failureCancelAllEnterArmed = focused == _failureCancelAll;
        if (! _failureCancelAllEnterArmed)
            SendMessageW(_failureClose, BM_CLICK, 0, 0);
        return true;
    }
    return false;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::HandleFailureSurfaceKeyUp(HWND focused, UINT virtualKey) noexcept
{
    if (! _failureSurface || virtualKey != VK_RETURN)
        return false;
    const bool invokeCancelAll  = _failureCancelAllEnterArmed && focused == _failureCancelAll;
    _failureCancelAllEnterArmed = false;
    if (invokeCancelAll)
        SendMessageW(_failureCancelAll, BM_CLICK, 0, 0);
    return true;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    RefreshLocalizedFooterText();

    if (folderWindow && hostLifetime.lock())
    {
        const AppTheme& theme = folderWindow->GetTheme();
        ApplyWindowChromeTheme(hwnd, theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd);
        if (! TryAttachHost(hwnd))
            static_cast<void>(EnterFailureSurface(hwnd));
    }
    ApplyScrollBarTheme(hwnd);
    ShowScrollBar(hwnd, SB_VERT, FALSE);
    _scrollBarVisible = false;

    UpdateLastPopupRect(hwnd);

    RefreshTelemetryTimer(hwnd);
    return 0;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::TryAttachHost(HWND hwnd) noexcept
{
    if (! hwnd || ! folderWindow || ! hostLifetime.lock())
        return false;
    const AppTheme& theme = folderWindow->GetTheme();
    const auto palette    = MakeAppThemeDxPalette(theme, theme.windowBackground);
    _reducedMotion        = palette.reducedMotion;
    DxUi::WindowHost::AttachOptions options{};
    options.presentationMode = DxUi::WindowHost::PresentationMode::CompositionSwapChain;
    bool forceAttachFailure  = false;
#ifdef ENABLE_TESTS
    forceAttachFailure = _debugFailureSurfacePinned || ConsumeForcedFailure(g_fileOperationsDxUiHostAttachForcedFailures);
#endif
    if (forceAttachFailure || ! _controlsHost.Attach(hwnd, options))
    {
        if (! forceAttachFailure)
            Debug::Error(L"File operations popup could not attach its DxUi control host.");
        return false;
    }
    _controlsHost.SetTheme(palette);
    auto root     = std::make_unique<FileOperationsPopupRoot>(_target, _preparedBackground);
    _controlsRoot = root.get();
    _controlsHost.SetRoot(std::move(root));
    _controlsHost.SetOnFocusChanged([this, hwnd](DxUi::Control* control)
    {
#ifdef ENABLE_TESTS
        ++_debugFocusChangeCount;
        _debugPreviousFocusTargetId = std::move(_debugLastFocusTargetId);
        _debugLastFocusTargetId     = control ? control->GetAccessibleAutomationId() : std::wstring{};
#endif
        // Focus is observable before the next frame. Enter/Escape must change
        // authority in the same transaction, including UIA and inspection return.
        // Reconciliation binds once after its control arrays are consistent.
        if (! _syncingHostedControls && _pathDetailId.empty())
            static_cast<void>(SyncHostedConflictKeyboardBindings());
        EnsureHostedFocusVisible(hwnd, control);
    });
    _controlsHost.SetOnEscape([this, hwnd]
    {
        if (_conflictEscapeTaskId != 0u)
        {
            // A retained conflict escape binding belongs to the generation
            // last synchronized into the host. Consume it even if that
            // generation became stale; closing would hide the newer prompt.
            static_cast<void>(SubmitPublishedConflictEscapeAction(hwnd));
            return true;
        }
        static_cast<void>(OnClose(hwnd));
        return true;
    });
    return true;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::TryLeaveFailureSurface(HWND hwnd) noexcept
{
    if (! _failureSurface || _failureSurfaceInitializing || ! hwnd)
        return false;
    // Every attempt backs off further until a frame is prepared again, so a device that keeps
    // failing after a successful attach cannot flip between the two surfaces at a fixed rate.
    const uint32_t shift   = std::min(++_hostReattachAttemptCount, 6u);
    _hostReattachRetryTick = GetTickCount64() + std::min(kFileOperationsPopupHostReattachRetryBaseMs << shift, kFileOperationsPopupHostReattachRetryMaxMs);
    if (! TryAttachHost(hwnd))
        return false;

    const HWND focus               = GetFocus();
    const bool focusOnSurface      = focus != nullptr && (focus == _failureText || focus == _failureCancelAll || focus == _failureClose);
    _failureSurface                = false;
    _failureSurfaceReady           = false;
    _failureSurfaceFailureReported = false;
    _failureCancelAllEnterArmed    = false;
    _hostReattachRetryTick         = 0u;
    // Escape binds a decision again only once the next frame publishes it.
    _conflictEscapeTaskId     = 0u;
    _conflictEscapeAction     = 0u;
    _conflictEscapeGeneration = 0u;
    if (focusOnSurface)
        SetFocus(hwnd);
    for (HWND* child : {&_failureText, &_failureCancelAll, &_failureClose})
    {
        if (*child)
        {
            RedSalamander::Win32Callback::RestoreWndProcHook(*child, kFailureSurfaceButtonWndProcProp);
            DestroyWindow(*child);
            *child = nullptr;
        }
    }
    _failureFont.reset();
    _failureFontDpi = 0u;
    _failureTextValue.clear();
    Invalidate(hwnd);
    RefreshTelemetryTimer(hwnd);
    return true;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::OnFramePrepareFailed(HWND hwnd, HRESULT hr) noexcept
{
    const ULONGLONG now = GetTickCount64();
    if (_framePrepareFailureCount == 0u)
        _framePrepareFirstFailureTick = now;
    if (_framePrepareFailureCount < kFileOperationsPopupFramePrepareFailureLimit)
        ++_framePrepareFailureCount;
    // Q6: only a host that is not attached, or repeated failures over time, select the
    // failure surface. A transient failure waits for the backoff and the popup timer.
    if (_controlsHost.GetHwnd() != hwnd || (_framePrepareFailureCount >= kFileOperationsPopupFramePrepareFailureLimit &&
                                            now - _framePrepareFirstFailureTick >= kFileOperationsPopupFramePrepareFailureWindowMs))
    {
        Debug::Error(L"File operations popup could not prepare a frame ({} attempts over {} ms, last hr=0x{:08X}); showing the failure surface.",
                     _framePrepareFailureCount,
                     now - _framePrepareFirstFailureTick,
                     static_cast<unsigned long>(hr));
        static_cast<void>(EnterFailureSurface(hwnd));
        return true;
    }
    if (_framePrepareFailureCount == 1u)
        Debug::Warning(L"File operations popup skipped a frame it could not prepare (hr=0x{:08X}); retrying.", static_cast<unsigned long>(hr));
    const uint32_t shift   = std::min(_framePrepareFailureCount - 1u, 4u);
    _framePrepareRetryTick = now + std::min(kFileOperationsPopupFramePrepareRetryBaseMs << shift, kFileOperationsPopupFramePrepareRetryMaxMs);
    RefreshTelemetryTimer(hwnd);
    return false;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ReportDegradedLayout(uint32_t step, HRESULT hr) noexcept
{
    static constexpr std::array<const wchar_t*, 5> kSteps{
        {L"conflict actions", L"completed group", L"terminal actions", L"terminal facts", L"terminal summary"}};
    if (step >= kSteps.size())
        return;
    Debug::Perf::Emit(L"FileOps.Popup.Layout.Degraded", kSteps[step], 0u, step, 0u, hr);
    const uint32_t bit = 1u << step;
    if ((_degradedLayoutReportedSteps & bit) != 0u)
        return;
    _degradedLayoutReportedSteps |= bit;
    Debug::Warning(L"File operations popup could not measure its {} (hr=0x{:08X}); showing a fallback layout.", kSteps[step], static_cast<unsigned long>(hr));
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnThemeChanged(HWND hwnd) noexcept
{
    if (_inThemeChange)
    {
        return 0;
    }

    _inThemeChange      = true;
    auto clearThemeFlag = wil::scope_exit([&] { _inThemeChange = false; });

    // Live settings reload uses the theme notification after changing language.
    // The native caption must follow the same satellite as the retained content.
    const std::wstring title = LoadStringResource(nullptr, IDS_FILEOPS_POPUP_TITLE);
    SetWindowTextW(hwnd, title.c_str());
    _scanStatusLayouts.clear();
    RefreshLocalizedFooterText();
    DiscardDeviceResources();
    // WM_SETTINGCHANGE can replace the number-format locale without changing the
    // selected language. Reformat localized count fragments on the next paint.
    _terminalFacts.clear();
    _terminalActions.clear();

    if (folderWindow && hostLifetime.lock())
    {
        const AppTheme& theme = folderWindow->GetTheme();
        const auto palette    = MakeAppThemeDxPalette(theme, theme.windowBackground);
        _reducedMotion        = palette.reducedMotion;
        _controlsHost.SetTheme(palette);
        ApplyWindowChromeTheme(hwnd, theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd);
        FileOperationsSpeedLimitPromptWindow::RefreshOwnedPrompts(hwnd, theme);
    }
    ApplyScrollBarTheme(hwnd);

    RedrawWindow(hwnd, nullptr, nullptr, RDW_FRAME | RDW_NOERASE | RDW_NOCHILDREN);
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcDestroy(HWND hwnd) noexcept
{
    _failureCancelAllEnterArmed = false;
    _inspectionLifetime.reset();
    KillTimer(hwnd, kFileOperationsPopupTimerId);
    _telemetryTimerIntervalMs = 0u;
    _pathDetailTaskId         = 0u;
    _pathDetailId.clear();
    _pathDetailReturnId.clear();
    _pathDetailPanel = nullptr;
    _pathDetailTitle = nullptr;
    _pathDetailField = nullptr;
    _pathDetailCopy  = nullptr;
    _pathDetailClose = nullptr;
#ifdef ENABLE_TESTS
    _debugRetiredControlQuarantine.clear();
#endif
    _controlsHost.Detach();
    _controlsRoot     = nullptr;
    _controlsBody     = nullptr;
    _failureText      = nullptr;
    _failureCancelAll = nullptr;
    _failureClose     = nullptr;
    _failureFont.reset();
    _hostedButtons.clear();
    _hostedButtonHits.clear();
    _hostedProgressBars.clear();
    _hostedProgressAnimations.clear();
    _hostedGraphs.clear();
    _hostedGraphStallBands.clear();
    _hostedTooltipRegions.clear();
    _hostedTooltipComparisonIds.clear();
    _hostedTooltipTaskIds.clear();
    _hostedButtonParents.clear();
    _hostedTooltipParents.clear();
    _retiredHostedControls.clear();
    _hostedAccessibleTextTicks.clear();
    _conflictBodies.clear();
    _hostedConflictBodyIds.clear();
    _hostedConflictBodyPanels.clear();
    ClearTaskbarProgress(hwnd);
    _taskbarList.reset();

    if (fileOps && hostLifetime.lock())
    {
        fileOps->OnPopupDestroyed(hwnd);
    }

    DiscardDeviceResources();

    _headerFormat.reset();
    _bodyFormat.reset();
    _smallFormat.reset();
    _footerSummaryLayout.reset();
    _footerSummaryFormat.reset();
    _footerSummaryText.clear();
    _terminalSummaries.clear();
    _cardInformationLayouts.clear();
    _scanStatusLayouts.clear();
    _terminalFacts.clear();
    _terminalActions.clear();
    _completedGroupMeasurement = {};
    _graphTrailingLabelFormat.reset();
    _buttonFormat.reset();
    _buttonSmallFormat.reset();
    _graphOverlayFormat.reset();
    _statusIconFormat.reset();
    _statusIconFallbackFormat.reset();
    _captionGlyphFormat.reset();
    _captionGlyphFallbackFormat.reset();
    _dwriteFactory.reset();
    _d2dFactory.reset();

    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    _deletePending = true;
    if (_dispatchDepth == 0u)
    {
        delete this;
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnPaint(HWND hwnd) noexcept
{
#ifdef ENABLE_TESTS
    ++_paintGenerationForSelfTest;
#endif
    const bool capturePerf   = Debug::Perf::IsCaptureEnabled();
    const uint64_t startedUs = capturePerf ? PerfNowUs() : 0u;
    Render(hwnd);
    const uint64_t hostedStartedUs = capturePerf ? PerfNowUs() : 0u;
    RECT hostedUpdateRect{};
    const bool hostedHasUpdate = capturePerf && GetUpdateRect(hwnd, &hostedUpdateRect, FALSE);
    bool hostedHandled         = false;
    static_cast<void>(_controlsHost.HandleMessage(hwnd, WM_PAINT, 0, 0, hostedHandled));
    // A skipped frame without an attached host must still consume its update region.
    if (! hostedHandled)
        ValidateRect(hwnd, nullptr);
    if (capturePerf)
    {
        Debug::Perf::Emit(L"FileOps.Popup.HostedPaintUs", L"", PerfElapsedUs(hostedStartedUs), hostedHasUpdate ? 1u : 0u, hostedHandled ? 1u : 0u, S_OK);
    }
    RefreshTelemetryTimer(hwnd);
    if (capturePerf)
        Debug::Perf::Emit(L"FileOps.Popup.WmPaintUs", L"", PerfElapsedUs(startedUs), 0u, 0u, S_OK);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnSize(HWND hwnd, WPARAM sizeType, UINT width, UINT height) noexcept
{
    _clientSize.cx = static_cast<LONG>(width);
    _clientSize.cy = static_cast<LONG>(height);
    if (_failureSurface)
    {
        LayoutFailureSurface(hwnd);
    }

    bool hostedHandled = false;
    static_cast<void>(_controlsHost.HandleMessage(hwnd, WM_SIZE, sizeType, MAKELPARAM(width, height), hostedHandled));

    UpdateLastPopupRect(hwnd);
    Invalidate(hwnd);
    RefreshTelemetryTimer(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT& suggested) noexcept
{
    _dpi = newDpi;

    bool hostedHandled = false;
    static_cast<void>(_controlsHost.HandleMessage(hwnd, WM_DPICHANGED, MAKEWPARAM(newDpi, newDpi), reinterpret_cast<LPARAM>(&suggested), hostedHandled));

    _headerFormat.reset();
    _bodyFormat.reset();
    _smallFormat.reset();
    _footerSummaryLayout.reset();
    _footerSummaryFormat.reset();
    _footerSummaryText.clear();
    _terminalSummaries.clear();
    _cardInformationLayouts.clear();
    _scanStatusLayouts.clear();
    _terminalFacts.clear();
    _terminalActions.clear();
    _completedGroupMeasurement = {};
    _graphTrailingLabelFormat.reset();
    _buttonFormat.reset();
    _buttonSmallFormat.reset();
    _graphOverlayFormat.reset();
    _statusIconFormat.reset();
    _statusIconFallbackFormat.reset();
    _captionGlyphFormat.reset();
    _captionGlyphFallbackFormat.reset();
    _captionGlyphDpi = 0;

    SetWindowPos(hwnd,
                 nullptr,
                 suggested.left,
                 suggested.top,
                 std::max(0L, suggested.right - suggested.left),
                 std::max(0L, suggested.bottom - suggested.top),
                 SWP_NOZORDER | SWP_NOACTIVATE);

    _maxAutoSizedWindowHeight = std::max(0L, suggested.bottom - suggested.top);

    UpdateLastPopupRect(hwnd);
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnGetMinMaxInfo(HWND hwnd, MINMAXINFO* info) noexcept
{
    if (! hwnd || ! info)
    {
        return 0;
    }

    const UINT dpiForWindow = GetDpiForWindow(hwnd);

    constexpr int kMinClientWidthDip = FileOperationsFooterLayout::kMinimumClientWidthDip;
    const int minClientHeightDip =
        (fileOps && fileOps->GetPopupFooterOnly()) ? kFileOperationsPopupFooterOnlyMinClientHeightDip : kFileOperationsPopupMinClientHeightDip;

    const int minClientW = DipsToPixels(kMinClientWidthDip, dpiForWindow);
    FileOperationsFooterLayout::Layout minimumFooter{};
    static_cast<void>(FileOperationsFooterLayout::Arrange(static_cast<float>(kMinClientWidthDip) - 20.0f,
                                                          _footerPauseResumeAllRect.right > _footerPauseResumeAllRect.left ? _footerActionWidthsDip[0] : 0.0f,
                                                          _footerCancelAllRect.right > _footerCancelAllRect.left ? _footerActionWidthsDip[1] : 0.0f,
                                                          minimumFooter,
                                                          _controlsHost.PixelsToDip(_footerMinimumSummaryHeightPx)));
    const int minClientH =
        std::max(DipsToPixels(minClientHeightDip, dpiForWindow), static_cast<int>(std::ceil(DipsToPixels(minimumFooter.heightDip + 8.0f, dpiForWindow))));

    const DWORD style   = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const DWORD exStyle = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE));

    RECT rc{0, 0, minClientW, minClientH};
    AdjustWindowRectExForDpi(&rc, style, FALSE, exStyle, dpiForWindow);

    const long minTrackW = std::max(0L, rc.right - rc.left);
    const long minTrackH = std::max(0L, rc.bottom - rc.top);

    info->ptMinTrackSize.x = std::max(static_cast<LONG>(info->ptMinTrackSize.x), minTrackW);
    info->ptMinTrackSize.y = std::max(static_cast<LONG>(info->ptMinTrackSize.y), minTrackH);

    static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMove(HWND hwnd) noexcept
{
    UpdateLastPopupRect(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnTimer(HWND hwnd, UINT_PTR timerId) noexcept
{
    if (timerId == kFileOperationsPopupTimerId)
    {
#ifdef ENABLE_TESTS
        ++_telemetryTimerTickCount;
#endif
        if (! hostLifetime.lock())
        {
            DestroyWindow(hwnd);
            return 0;
        }
        RefreshTelemetryTimer(hwnd);
        if (_deletePending)
            return 0;
        if (_telemetryTimerIntervalMs == 0u)
        {
            // The tick that stops the timer still asks for one frame: a motion whose deadline passed since the last
            // frame (a late tick on a busy UI thread) settles at its target only when a frame resolves it, and nothing
            // else may paint again (I35, owner report 2026-10-13: a collapsed card kept its expanded height).
            if (IsWindowVisible(hwnd) && ! IsIconic(hwnd) && ! _failureSurface)
                Invalidate(hwnd);
            return 0;
        }

        if (! IsWindowVisible(hwnd) || IsIconic(hwnd))
        {
            UpdateTaskbarProgress(hwnd);
            return 0;
        }
        if (_failureSurface)
        {
            UpdateTaskbarProgress(hwnd);
            // Q6: re-attempt the host with backoff; success restores the hosted popup.
            if (_hostReattachRetryTick != 0u && GetTickCount64() >= _hostReattachRetryTick && TryLeaveFailureSurface(hwnd))
                return 0;
            RefreshFailureSurfaceText();
            return 0;
        }

        UpdateRates();
        Invalidate(hwnd);
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnEnterSizeMove(HWND hwnd) noexcept
{
    static_cast<void>(hwnd);
    _inSizeMove          = true;
    _autoResizePending   = false;
    _autoResizeAnimating = false;
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnExitSizeMove(HWND hwnd) noexcept
{
    if (hwnd)
    {
        RECT rc{};
        GetWindowRect(hwnd, &rc);
        _maxAutoSizedWindowHeight = std::max(0L, rc.bottom - rc.top);
    }
    // A user resize establishes the new geometry for this measured content;
    // queued automatic sizing must not expand it again on the next paint.
    _lastAutoSizedContentHeight = _contentHeight;
    _inSizeMove                 = false;
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnVScroll(HWND hwnd, UINT request) noexcept
{
    if (! _pathDetailId.empty())
        return 0;
    if (! hwnd)
    {
        return 0;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    if (! GetScrollInfo(hwnd, SB_VERT, &si))
    {
        return 0;
    }

    const int page     = std::max(1, static_cast<int>(si.nPage));
    const int maxPos   = std::max(0, si.nMax - page + 1);
    const int lineStep = std::max(1, DipsToPixels(36, _dpi));
    const int pageStep = page;

    int newPos = _scrollPos;
    switch (request)
    {
        case SB_TOP: newPos = 0; break;
        case SB_BOTTOM: newPos = maxPos; break;
        case SB_LINEUP: newPos -= lineStep; break;
        case SB_LINEDOWN: newPos += lineStep; break;
        case SB_PAGEUP: newPos -= pageStep; break;
        case SB_PAGEDOWN: newPos += pageStep; break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: newPos = si.nTrackPos; break;
        default: return 0;
    }

    newPos = std::clamp(newPos, 0, maxPos);
    if (newPos == _scrollPos)
    {
        return 0;
    }

    _scrollPos = newPos;

    SCROLLINFO set{};
    set.cbSize = sizeof(set);
    set.fMask  = SIF_POS;
    set.nPos   = _scrollPos;
    SetScrollInfo(hwnd, SB_VERT, &set, TRUE);

    Invalidate(hwnd);
    return 0;
}

std::wstring FileOperationsPopupInternal::FileOperationsPopupState::ConflictInspectionResultText(const TaskSnapshot& task) const
{
    if (! _inspectionPresentation || _inspectionPresentation->invokingHit.taskId != task.taskId ||
        _inspectionPresentation->invokingHit.promptGeneration != task.conflict.generation || _inspectionPresentation->context != task.conflict.inspection ||
        ! task.conflict.active)
        return {};
    return LoadStringResource(nullptr, _inspectionPresentation->messageId);
}

void FileOperationsPopupInternal::FileOperationsPopupState::BeginConflictInspection(HWND hwnd, const PopupHitTest& hit) noexcept
{
    if (! fileOps || ! hit.promptGeneration || ! hostLifetime.lock() || (_inspectionPresentation && _inspectionPresentation->pending))
        return;
    const auto published = std::ranges::find_if(
        _publishedInspectionContexts, [&](const auto& entry) noexcept { return entry.taskId == hit.taskId && entry.promptGeneration == hit.promptGeneration; });
    if (published == _publishedInspectionContexts.end())
        return;
    _inspectionReturn.reset();
    const auto context      = published->context;
    _inspectionPresentation = InspectionPresentation{hit, context, IDS_FILEOPS_INSPECTION_PENDING, true};
    if (! _inspectionLifetime)
        _inspectionLifetime = std::make_shared<uint8_t>(uint8_t{0});
    const std::weak_ptr<void> lifetime = _inspectionLifetime;
    const auto complete                = [this, hwnd, lifetime, context, hit](const FileOperations::ConflictInspectionNavigationResult& result)
    {
        if (lifetime.expired() || ! hostLifetime.lock() || ! _inspectionPresentation || _inspectionPresentation->context != context ||
            _inspectionPresentation->invokingHit.taskId != hit.taskId || _inspectionPresentation->invokingHit.promptGeneration != hit.promptGeneration ||
            ! _inspectionPresentation->pending)
            return;
        using Status   = FileOperations::ConflictInspectionNavigationStatus;
        UINT messageId = IDS_FILEOPS_INSPECTION_UNAVAILABLE;
        switch (result.status)
        {
            case Status::Selected: messageId = IDS_FILEOPS_INSPECTION_SELECTED; break;
            case Status::MissingOrFiltered: messageId = IDS_FILEOPS_INSPECTION_MISSING; break;
            case Status::Ambiguous: messageId = IDS_FILEOPS_INSPECTION_AMBIGUOUS; break;
            case Status::Unavailable: break;
            case Status::Stale: messageId = IDS_FILEOPS_INSPECTION_STALE; break;
            case Status::NavigationFailed: messageId = IDS_FILEOPS_INSPECTION_FAILED; break;
            case Status::RollbackFailed: messageId = IDS_FILEOPS_INSPECTION_ROLLBACK_FAILED; break;
            case Status::Canceled: messageId = IDS_FILEOPS_INSPECTION_CANCELED; break;
        }
        if (result.error == HRESULT_FROM_WIN32(ERROR_BUSY))
            messageId = IDS_FILEOPS_INSPECTION_BUSY;
        _inspectionPresentation->messageId = messageId;
        _inspectionPresentation->pending   = false;
        Invalidate(hwnd);
        if (result.status == Status::Selected && result.focusTicket)
        {
            const auto focused         = std::ranges::find(_hostedButtons, _controlsHost.GetFocusControl());
            const size_t index         = static_cast<size_t>(std::distance(_hostedButtons.begin(), focused));
            const bool invokingControl = index < _hostedButtonHits.size() && _hostedButtonHits[index].kind == hit.kind &&
                                         _hostedButtonHits[index].taskId == hit.taskId && _hostedButtonHits[index].data == hit.data &&
                                         _hostedButtonHits[index].promptGeneration == hit.promptGeneration;
            const int returnScroll     = _scrollPos;
            if (invokingControl)
            {
                const bool committed = fileOps->CommitPendingConflictInspectionFocus(result.focusTicket, hwnd);
                if (lifetime.expired() || ! hostLifetime.lock())
                    return;
                if (committed && fileOps->IsPendingConflictInspectionCurrent(hit.taskId, hit.promptGeneration, context))
                {
                    _inspectionReturn = InspectionReturn{hit, context, returnScroll, false};
                    static_cast<void>(OnClose(hwnd));
                    return;
                }
            }
        }
        // Only the current prompt owns this result. A replacement prompt cannot
        // inherit stale status text or an accessibility announcement.
        if (fileOps->IsPendingConflictInspectionCurrent(hit.taskId, hit.promptGeneration, context))
            static_cast<void>(DxUi::RaiseWindowHostAccessibilityNotification(
                hwnd, LoadStringResource(nullptr, messageId), std::format(L"FileOperations.Task.{}.Inspection", hit.taskId)));
    };
    Invalidate(hwnd);
    const HRESULT status = context ? fileOps->BeginPendingConflictInspection(
                                         hit.taskId, hit.promptGeneration, context, hit.kind == PopupHitTest::Kind::TaskConflictInspectDestination, complete)
                                   : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    if (lifetime.expired())
        return;
    // Begin can complete synchronously and then return a failure after provider
    // reentry. Never overwrite a callback that has already settled this request.
    if (FAILED(status) && _inspectionPresentation && _inspectionPresentation->pending)
    {
        complete(
            {status == E_ABORT ? FileOperations::ConflictInspectionNavigationStatus::Stale : FileOperations::ConflictInspectionNavigationStatus::Unavailable,
             status});
    }
}

bool FileOperationsPopupInternal::FileOperationsPopupState::HasCancelableTasks() const noexcept
{
    const auto lifetime = hostLifetime.lock();
    if (! lifetime)
        return false;
#ifdef ENABLE_TESTS
    if (debugVisualScenario)
        return std::ranges::any_of(debugVisualScenario.value(), [](const TaskSnapshot& task) noexcept {
            return task.kind == TaskSnapshot::Kind::FileOperation && ! task.finished && ! task.cancellationRequested;
        });
#endif
    // Engine truth includes presentation-hidden tasks; global cancellation must
    // not depend on reveal deadlines or copy paths/history just to decide consent.
    return fileOps && fileOps->HasCancelableOperations();
}

bool FileOperationsPopupInternal::FileOperationsPopupState::IsTaskDecisionAvailable(uint64_t taskId) const noexcept
{
    const auto lifetime = hostLifetime.lock();
    if (! lifetime)
        return false;
#ifdef ENABLE_TESTS
    if (debugVisualScenario)
        for (const auto& task : debugVisualScenario.value())
            if (task.taskId == taskId)
                return ! task.finished && ! task.cancellationRequested;
#endif
    const auto* task = fileOps ? fileOps->FindTask(taskId) : nullptr;
    return task != nullptr && ! task->CancelRequestedDurable() && ! task->_taskFinished.load(std::memory_order_acquire);
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnActivatedHit(HWND hwnd, const PopupHitTest& hit) noexcept
{
    if (! _pathDetailId.empty())
        return 0;
    if (! hostLifetime.lock())
    {
        DestroyWindow(hwnd);
        return 0;
    }
    if (IsTaskDecisionHit(hit.kind) && ! IsTaskDecisionAvailable(hit.taskId))
        return 0;
    if (_failureSurface && hit.kind != PopupHitTest::Kind::FooterCancelAll)
    {
        // C4: the failure surface offers Cancel all and Close only; no other action is reachable.
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskConflictInspectSource || hit.kind == PopupHitTest::Kind::TaskConflictInspectDestination)
    {
        BeginConflictInspection(hwnd, hit);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterCancelAll)
    {
        static_cast<void>(ConfirmCancelAll(hwnd));
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterAutoDismiss)
    {
        if (fileOps)
        {
            fileOps->SetAutoDismissSuccess(! fileOps->GetAutoDismissSuccess());
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterPauseResumeAll)
    {
        if (fileOps)
        {
            bool pause = hit.data == kFooterPauseResumeAllPauseAction;
            if (hit.data != kFooterPauseResumeAllPauseAction && hit.data != kFooterPauseResumeAllResumeAction)
            {
                pause = FooterPauseResumeAllShouldPause(BuildGlobalStatusSummary(BuildSnapshot()));
            }

            fileOps->SetAllRunningTasksPaused(pause);
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterQueueMode)
    {
        if (fileOps)
        {
            if (hit.data == kFooterQueueModeQueueAction || hit.data == kFooterQueueModeParallelAction)
            {
                fileOps->ApplyQueueMode(hit.data == kFooterQueueModeQueueAction);
            }
            else
            {
                ShowQueueModeMenu(hwnd);
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterOptions)
    {
        ShowFooterOptionsMenu(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterDensity)
    {
        if (fileOps)
        {
            fileOps->SetPopupCompactDensity(! fileOps->GetPopupCompactDensity());
        }
        _maxAutoSizedWindowHeight   = 0;
        _lastAutoSizedContentHeight = -1.0f;
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterToggleDetails)
    {
        if (fileOps)
        {
            const bool wasFooterOnly = fileOps->GetPopupFooterOnly();
            if (! wasFooterOnly)
            {
                RECT windowRc{};
                if (GetWindowRect(hwnd, &windowRc) != FALSE)
                {
                    _footerOnlyRestoreWindowRect = windowRc;
                }
                fileOps->SavePopupExpandedPlacement(hwnd);
            }

            fileOps->SetPopupFooterOnly(! wasFooterOnly);
            if (wasFooterOnly && _footerOnlyRestoreWindowRect.has_value())
            {
                const RECT restoreRc = _footerOnlyRestoreWindowRect.value();
                SetWindowPos(hwnd,
                             nullptr,
                             restoreRc.left,
                             restoreRc.top,
                             std::max(0L, restoreRc.right - restoreRc.left),
                             std::max(0L, restoreRc.bottom - restoreRc.top),
                             SWP_NOZORDER | SWP_NOACTIVATE);
                _footerOnlyRestorePending = true;
                _maxAutoSizedWindowHeight = std::max(0L, restoreRc.bottom - restoreRc.top);
            }
        }
        _autoResizePending          = false;
        _autoResizeAnimating        = false;
        _maxAutoSizedWindowHeight   = fileOps && fileOps->GetPopupFooterOnly() ? 0 : _maxAutoSizedWindowHeight;
        _lastAutoSizedContentHeight = -1.0f;
        _lastTaskCount              = std::numeric_limits<size_t>::max();
        ArmUiMotionTimer(hwnd);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::CompletedGroupToggle)
    {
        // The user's own expand or collapse wins over a pending one-time collapse after Cancel all.
        _bulkCancelCohort.clear();
        _completedGroupExpanded     = ! _completedGroupExpanded;
        _maxAutoSizedWindowHeight   = 0;
        _lastAutoSizedContentHeight = -1.0f;
        ArmUiMotionTimer(hwnd);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::CompletedGroupClear)
    {
        // Clear only what the last published header covered; a task that finished after that
        // paint stays for the user to see.
        const std::vector<uint64_t> covered = std::exchange(_completedGroupClearTaskIds, {});
        if (fileOps)
        {
            fileOps->ClearCompletedTasks(covered);
        }
        _maxAutoSizedWindowHeight   = 0;
        _lastAutoSizedContentHeight = -1.0f;
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskToggleThroughput)
    {
        const auto card = _throughputExpandedTasks.find(hit.taskId);
        const bool published =
            std::ranges::any_of(_buttons, [&](const PopupButton& button) { return button.hit.kind == hit.kind && button.hit.taskId == hit.taskId; });
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(std::format(L"Throughput activation: task={} card={} published={} before={} nativeFocus=0x{:X}",
                                                  hit.taskId,
                                                  card != _throughputExpandedTasks.end(),
                                                  published,
                                                  card != _throughputExpandedTasks.end() && card->second,
                                                  reinterpret_cast<UINT_PTR>(GetFocus())));
#endif
        if (card == _throughputExpandedTasks.end() || ! published)
            return 0;
        card->second = ! card->second;
        if (fileOps)
            fileOps->SetPopupThroughputExpanded(card->second);
        _taskHeightAnimations.erase(hit.taskId);
        _preserveWindowSizeForThroughputChange = true;
        _autoResizePending                     = false;
        _autoResizeAnimating                   = false;
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskToggleCollapse)
    {
        ToggleTaskCollapsed(hit.taskId, fileOps ? fileOps->GetPopupCompactDensity() : false);
        ArmUiMotionTimer(hwnd);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskPause)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->TogglePause();
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskStartNow)
    {
        if (fileOps)
        {
            static_cast<void>(fileOps->RunQueuedTaskNow(hit.taskId));
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskQueueMoveUp || hit.kind == PopupHitTest::Kind::TaskQueueMoveDown)
    {
        if (fileOps)
        {
            static_cast<void>(fileOps->MoveQueuedTask(hit.taskId, hit.kind == PopupHitTest::Kind::TaskQueueMoveUp));
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskCancel)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->RequestCancel();
                fileOps->CancelConflictInspectionForTask(hit.taskId);
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskDismiss)
    {
        if (fileOps)
        {
            fileOps->DismissCompletedTask(hit.taskId);
            fileOps->DismissInformationalTask(hit.taskId);
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskIssues)
    {
        if (fileOps)
            static_cast<void>(fileOps->OpenIssuesForTask(hit.taskId));
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskCompletedMore)
    {
        ShowCompletedOverflowMenu(hwnd, hit.taskId);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskShowLog)
    {
        if (fileOps)
        {
            static_cast<void>(fileOps->OpenDiagnosticsLogForTask(hit.taskId));
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskExportIssues)
    {
        if (fileOps)
        {
            static_cast<void>(fileOps->ExportTaskIssuesReport(hit.taskId));
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskSkip)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->SkipDiscovery();
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskSpeedLimit)
    {
        ShowSpeedLimitMenu(hwnd, hit.taskId);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->ToggleConflictApplyToAllChecked(hit.promptGeneration);
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskConflictMore)
    {
        ShowConflictOverflowMenu(hwnd, hit.taskId, hit.promptGeneration);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskConflictAction)
    {
        static_cast<void>(SubmitConflictPrimaryAction(hwnd, hit.taskId, hit.data, hit.promptGeneration));
        return 0;
    }

    return 0;
}

#ifdef ENABLE_TESTS
LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnSelfTestInvoke(HWND hwnd, const PopupSelfTestInvoke* payload) noexcept
{
    if (! payload)
    {
        return 0;
    }

    if (IsTaskDecisionHit(payload->kind) && ! IsTaskDecisionAvailable(payload->taskId))
        return 0;

    if (payload->kind == PopupHitTest::Kind::None && payload->data == kPopupSelfTestDestroyOnNextShowData)
    {
        _destroyOnNextShowForSelfTest = true;
        return 1;
    }

    uint64_t generation = payload->promptGeneration;
    if (generation == 0u && fileOps)
    {
        if (auto* task = fileOps->FindTask(payload->taskId))
        {
            std::scoped_lock lock(task->_conflictArbiter.mutex);
            generation = task->_conflictArbiter.prompt.generation;
        }
    }
    if (_failureSurface && payload->kind != PopupHitTest::Kind::None && payload->kind != PopupHitTest::Kind::FooterCancelAll)
    {
        // C4: the failure surface offers Cancel all and Close only.
        return 0;
    }

    if (payload->kind == PopupHitTest::Kind::TaskSpeedLimit && payload->data == 1u)
    {
        // Let self-test callers advance their state before the modal prompt loop starts.
        return PostMessageW(hwnd, kFileOperationsPopupDeferredSpeedLimitPromptMessage, 0, static_cast<LPARAM>(payload->taskId)) != FALSE ? 1 : 0;
    }

    if (payload->kind == PopupHitTest::Kind::TaskConflictMore && payload->data != 0u)
    {
        return SubmitConflictOverflowAction(hwnd, payload->taskId, payload->data, generation) ? 1 : 0;
    }

    if (payload->kind == PopupHitTest::Kind::TaskCompletedMore && payload->data != 0u)
    {
        return SubmitCompletedOverflowAction(hwnd, payload->taskId, payload->data, false) ? 1 : 0;
    }

    if (payload->kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll || payload->kind == PopupHitTest::Kind::TaskConflictAction ||
        payload->kind == PopupHitTest::Kind::FooterToggleDetails || payload->kind == PopupHitTest::Kind::FooterQueueMode ||
        payload->kind == PopupHitTest::Kind::FooterDensity || payload->kind == PopupHitTest::Kind::FooterAutoDismiss ||
        payload->kind == PopupHitTest::Kind::FooterPauseResumeAll || payload->kind == PopupHitTest::Kind::FooterOptions ||
        payload->kind == PopupHitTest::Kind::CompletedGroupToggle || payload->kind == PopupHitTest::Kind::CompletedGroupClear ||
        payload->kind == PopupHitTest::Kind::TaskToggleCollapse || payload->kind == PopupHitTest::Kind::TaskToggleThroughput ||
        payload->kind == PopupHitTest::Kind::TaskStartNow || payload->kind == PopupHitTest::Kind::TaskQueueMoveUp ||
        payload->kind == PopupHitTest::Kind::TaskQueueMoveDown || payload->kind == PopupHitTest::Kind::TaskConflictInspectSource ||
        payload->kind == PopupHitTest::Kind::TaskConflictInspectDestination || payload->kind == PopupHitTest::Kind::TaskIssues)
    {
        // OnActivatedHit returns 0 for these kinds even on success; self-test callers need a
        // dispatched-successfully signal.
        static_cast<void>(OnActivatedHit(hwnd, PopupHitTest{payload->kind, payload->taskId, payload->data, generation}));
        return 1;
    }

    return OnActivatedHit(hwnd, PopupHitTest{payload->kind, payload->taskId, payload->data, generation});
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnTaskSnapshotRequest(const PopupTaskSnapshotRequest* request) const noexcept
{
    if (! request)
    {
        return 0;
    }

    PopupTaskSnapshotRequest& mutableRequest = *const_cast<PopupTaskSnapshotRequest*>(request);
    mutableRequest.found                     = false;

    const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
    auto it = mutableRequest.taskId == 0
                  ? std::find_if(snapshot.begin(),
                                 snapshot.end(),
                                 [](const TaskSnapshot& task) noexcept { return task.kind == TaskSnapshot::Kind::FileOperation && ! task.finished; })
                  : std::find_if(snapshot.begin(), snapshot.end(), [&mutableRequest](const TaskSnapshot& task) noexcept {
        return task.taskId == mutableRequest.taskId;
    });
    if (mutableRequest.taskId == 0 && it == snapshot.end())
    {
        it = std::find_if(snapshot.begin(), snapshot.end(), [](const TaskSnapshot& task) noexcept { return task.kind == TaskSnapshot::Kind::FileOperation; });
    }
    if (it == snapshot.end())
    {
        return 0;
    }

    mutableRequest.snapshot = *it;
    mutableRequest.found    = true;
    return 1;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnCaptionGlyphSnapshotRequest(CaptionGlyphDebugSnapshot* snapshot) const noexcept
{
    if (! snapshot)
    {
        return 0;
    }

    const bool highContrast                 = folderWindow && folderWindow->GetTheme().highContrast;
    snapshot->statusVisible                 = _captionStatus != CaptionStatus::None && ! highContrast;
    snapshot->highContrastSuppressed        = _captionStatus != CaptionStatus::None && highContrast;
    snapshot->usesDirectWriteGlyphRendering = true;
    snapshot->usesGdiTextFallback           = false;
    return 1;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnLayoutSnapshotRequest(PopupLayoutDebugSnapshot* snapshot) const noexcept
{
    if (! snapshot)
    {
        return 0;
    }

    PopupLayoutDebugSnapshot result{};
    result.taskId                     = snapshot->taskId;
    result.hostedKeyboardFocusVisible = _controlsHost.IsKeyboardFocusVisible();
    result.hostedReducedMotion        = _controlsHost.GetTheme().reducedMotion;

    const auto rectHasArea = [](const D2D1_RECT_F& rc) noexcept { return rc.right > rc.left && rc.bottom > rc.top; };

    const std::vector<TaskSnapshot> taskSnapshot            = BuildSnapshot();
    GlobalFileOperationsStatusSummary globalSummary         = BuildGlobalStatusSummary(taskSnapshot, &_rates);
    const GlobalTaskbarProgressModel taskbarProgress        = BuildGlobalTaskbarProgressModel(globalSummary);
    const AppTheme appTheme                                 = folderWindow ? folderWindow->GetTheme() : AppTheme{};
    const bool highContrast                                 = appTheme.highContrast;
    const bool reducedMotion                                = IsReducedMotionEnabled();
    const CompletedGroupResultSummary completedGroupSummary = BuildCompletedGroupResultSummary(taskSnapshot);
    const uint32_t completedGroupCount                      = completedGroupSummary.total;
    const bool completedGroupVisible                        = ShouldShowCompletedGroup(completedGroupCount);
    result.globalRunningCount                               = globalSummary.running;
    result.globalWaitingCount                               = globalSummary.waiting;
    result.globalNeedAttentionCount                         = globalSummary.needAttention;
    result.globalSummaryText                                = FormatGlobalStatusSummaryText(globalSummary);
    result.globalSummaryVisible                             = ! result.globalSummaryText.empty() && rectHasArea(_footerSummaryRect);
    result.footerPauseResumeAllVisible                      = rectHasArea(_footerPauseResumeAllRect) && HasFooterPauseResumeAllControl(globalSummary);
    result.footerPauseResumeAllPauses                       = FooterPauseResumeAllShouldPause(globalSummary);
    result.footerQueueModeSegmentedVisible                  = false;
    result.footerQueueModeSelectorVisible                   = rectHasArea(_footerQueueModeRect);
    result.footerQueueModeIsParallel                        = fileOps ? ! fileOps->GetQueueNewTasks() : false;
    result.footerQueueSegmentRect                           = _footerQueueSegmentRect;
    result.footerParallelSegmentRect                        = _footerParallelSegmentRect;
    result.footerSummaryRect                                = _footerSummaryRect;
    if (_footerSummaryLayout)
    {
        DWRITE_TEXT_METRICS summaryMetrics{};
        if (SUCCEEDED(_footerSummaryLayout->GetMetrics(&summaryMetrics)))
        {
            result.footerSummaryRequiredHeightPx = summaryMetrics.height;
            result.footerSummaryLineCount        = summaryMetrics.lineCount;
            result.footerSummaryFits             = summaryMetrics.height <= _footerSummaryRect.bottom - _footerSummaryRect.top + 0.5f &&
                                                   summaryMetrics.width <= _footerSummaryRect.right - _footerSummaryRect.left + 0.5f &&
                                                   _footerSummaryRect.bottom <= static_cast<float>(_clientSize.cy) && _footerSummaryRect.top >= 0.0f;
            for (const auto& command : {_footerPauseResumeAllRect, _footerCancelAllRect, _footerQueueModeRect, _footerOptionsRect, _footerDetailsToggleRect})
                if (rectHasArea(command))
                    result.footerSummaryFits =
                        result.footerSummaryFits && command.top >= _footerSummaryRect.bottom && command.bottom <= static_cast<float>(_clientSize.cy);
        }
    }
    if (result.footerQueueModeSelectorVisible)
    {
        const float hitX                              = (_footerQueueModeRect.left + _footerQueueModeRect.right) * 0.5f;
        const float hitY                              = (_footerQueueModeRect.top + _footerQueueModeRect.bottom) * 0.5f;
        const PopupHitTest hit                        = HitTest(hitX, hitY);
        result.footerQueueModeSelectorHitTargetActive = hit.kind == PopupHitTest::Kind::FooterQueueMode && hit.data == 0u;
    }
    result.footerAutoDismissVisible      = false;
    result.footerAutoDismissLabelVisible = _footerAutoDismissLabelVisible;
    result.footerAutoDismissEnabled      = fileOps ? fileOps->GetAutoDismissSuccess() : false;
    result.checkboxGlyph                 = _debugLastDrawnCheckboxGlyph;
    result.footerDensityToggleVisible    = false;
    if (result.footerDensityToggleVisible)
    {
        const float hitX                    = (_footerDensityRect.left + _footerDensityRect.right) * 0.5f;
        const float hitY                    = (_footerDensityRect.top + _footerDensityRect.bottom) * 0.5f;
        result.footerDensityHitTargetActive = HitTest(hitX, hitY).kind == PopupHitTest::Kind::FooterDensity;
    }
    result.footerOptionsVisible = rectHasArea(_footerOptionsRect);
    if (result.footerOptionsVisible)
    {
        const float hitX                    = (_footerOptionsRect.left + _footerOptionsRect.right) * 0.5f;
        const float hitY                    = (_footerOptionsRect.top + _footerOptionsRect.bottom) * 0.5f;
        result.footerOptionsHitTargetActive = HitTest(hitX, hitY).kind == PopupHitTest::Kind::FooterOptions;
    }
    result.popupCompactDensity                = fileOps ? fileOps->GetPopupCompactDensity() : false;
    result.footerAggregateProgressVisible     = HasGlobalAggregateProgress(globalSummary);
    result.footerAggregateProgressDeterminate = HasDeterminateGlobalAggregateProgress(globalSummary);
    result.footerAggregateProgressFraction    = GlobalAggregateProgressFraction(globalSummary);
    result.footerOpenDiscoveryTasks           = globalSummary.openDiscoveryTasks;
    result.footerAggregateCompletedBytes      = globalSummary.completedBytes;
    result.footerAggregateTotalBytes          = globalSummary.totalBytes;
    result.footerAggregateCompletedItems      = globalSummary.completedItems;
    result.footerAggregateTotalItems          = globalSummary.totalItems;
    result.footerAggregateBytesPerSecond      = globalSummary.displayedBytesPerSec;
    result.footerAggregateEtaVisible          = globalSummary.hasAggregateEta;
    result.footerAggregateEtaSeconds          = globalSummary.hasAggregateEta ? SaturatingCeilNonNegativeToUint64(globalSummary.aggregateEtaSeconds) : 0ull;
    result.taskbarProgressState               = taskbarProgress.state;
    result.taskbarProgressCompleted           = taskbarProgress.completed;
    result.taskbarProgressTotal               = taskbarProgress.total;
    result.taskbarUpdateCount                 = _taskbarUpdateCount;
    result.telemetryTimerTickCount            = _telemetryTimerTickCount;
    result.taskSnapshotBuildCount             = _taskSnapshotBuildCount;
    result.taskbarButtonReady                 = _taskbarButtonReady;
    result.taskbarListAvailable               = static_cast<bool>(_taskbarList);
    result.taskbarListAttemptCount            = _taskbarListAttemptCount;
    if (_taskbarListRetryAfterTick != 0)
    {
        const ULONGLONG nowTick = GetTickCount64();
        if (nowTick < _taskbarListRetryAfterTick)
        {
            result.taskbarListRetryPending = true;
            result.taskbarListRetryDelayMs = static_cast<uint64_t>(_taskbarListRetryAfterTick - nowTick);
        }
    }
    result.footerOnly                 = fileOps ? fileOps->GetPopupFooterOnly() : false;
    result.footerDetailsToggleVisible = rectHasArea(_footerDetailsToggleRect);
    result.footerDetailsToggleRightAligned =
        result.footerDetailsToggleVisible && _footerDetailsToggleRect.right >= static_cast<float>(_clientSize.cx) - DipsToPixels(12.0f, _dpi);
    result.footerDetailsToggleUsesDisclosureChrome = false;
    result.highContrastEnabled                     = highContrast;
    result.reducedMotionEnabled                    = reducedMotion;
    const auto scalarAnimationActive               = [](const ScalarAnimation& animation) noexcept
    { return animation.initialized && std::abs(animation.displayed - animation.target) > 0.01; };
    result.autoResizeMotionActive                 = _autoResizePending || _autoResizeAnimating;
    result.completedGroupLayoutMotionActive       = completedGroupVisible && scalarAnimationActive(_completedGroupVisibilityAnimation);
    result.layoutMotionActive                     = result.autoResizeMotionActive || result.completedGroupLayoutMotionActive;
    result.autoResizeAnimationEnabled             = ! reducedMotion;
    result.footerQueueModeAnimationEnabled        = false;
    result.taskCardAnimationEnabled               = ! reducedMotion;
    result.hostedProgressAnimationEnabled         = ! reducedMotion;
    result.usesDxUiHost                           = _controlsHost.GetHwnd() != nullptr && _controlsRoot != nullptr;
    result.surfaceResources                       = FileOperationsSelfTest::CaptureSurfaceResources(_controlsHost);
    result.hostedActionControlCount               = _hostedButtons.size();
    result.hostedProgressControlCount             = _hostedProgressBars.size();
    result.hostedGraphControlCount                = _hostedGraphs.size();
    result.hostedGraphStallBandCount              = _hostedGraphStallBands.size();
    result.hostedTooltipRegionCount               = _hostedTooltipRegions.size();
    result.completedGroupBadgeTooltipsInformative = ! _hostedTooltipDescriptors.empty() && _hostedTooltipDescriptors.size() == _hostedTooltipRegions.size();
    for (size_t index = 0u; result.completedGroupBadgeTooltipsInformative && index < _hostedTooltipDescriptors.size(); ++index)
    {
        const HostedTooltipDescriptor& descriptor = _hostedTooltipDescriptors[index];
        const DxUi::Control* region               = _hostedTooltipRegions[index];
        result.completedGroupBadgeTooltipsInformative =
            region && ! descriptor.text.empty() && descriptor.text.find_first_not_of(L"0123456789 \t") != std::wstring::npos &&
            region->GetTooltipText() == descriptor.text && region->GetAccessibleName() == descriptor.text && region->GetAccessibleHelpText() == descriptor.text;
    }
    result.hostedAccessibilityNotificationCount = _hostedAccessibilityNotificationCount;
    result.hostedNativeTabCount                 = _debugNativeTabCount;
    result.failureSurfaceTabCount               = _debugFailureSurfaceTabCount;
    result.failureSurfaceReentrantEntryCount    = _debugFailureSurfaceReentrantEntryCount;
    result.hostedFocusChangeCount               = _debugFocusChangeCount;
    result.hostedPreviousFocusTargetId          = _debugPreviousFocusTargetId;
    result.hostedLastFocusTargetId              = _debugLastFocusTargetId;
    result.scrollPosition                       = _scrollPos;
    result.listViewportRect                     = _listViewportRect;
    if (const auto* focusedControl = _controlsHost.GetFocusControl())
    {
        result.hostedFocusedAutomationId = focusedControl->GetAccessibleAutomationId();
        result.hostedFocusedHasFocus     = focusedControl->HasFocus();
    }
    if (const auto* defaultButton = _controlsHost.GetDefaultButton())
    {
        result.hostedDefaultAutomationId = defaultButton->GetAccessibleAutomationId();
    }
    if (const auto* cancelButton = _controlsHost.GetCancelButton())
    {
        result.hostedCancelAutomationId = cancelButton->GetAccessibleAutomationId();
    }
    result.completedGroupVisible          = completedGroupVisible;
    result.completedGroupExpanded         = completedGroupVisible && _completedGroupExpanded;
    result.completedGroupCount            = completedGroupVisible ? completedGroupCount : 0u;
    result.completedGroupCompletedCount   = completedGroupVisible ? completedGroupSummary.completed : 0u;
    result.completedGroupPartialCount     = completedGroupVisible ? completedGroupSummary.partial : 0u;
    result.completedGroupFailedCount      = completedGroupVisible ? completedGroupSummary.failed : 0u;
    result.completedGroupCanceledCount    = completedGroupVisible ? completedGroupSummary.canceled : 0u;
    result.completedGroupVisibleTaskCount = result.completedGroupExpanded ? completedGroupCount : 0u;
    result.completedGroupAnimationEnabled = completedGroupVisible && ! reducedMotion;
    if (completedGroupVisible)
    {
        for (const uint32_t count : {completedGroupSummary.completed,
                                     completedGroupSummary.partial,
                                     completedGroupSummary.indeterminate,
                                     completedGroupSummary.failed,
                                     completedGroupSummary.canceled})
            result.completedGroupExpectedBadges += count != 0u ? 1u : 0u;
        result.completedGroupBadgeCountsFit = true;
        constexpr std::array<std::wstring_view, 5> badgeSuffixes{{L"Completed", L"Partial", L"Indeterminate", L"Failed", L"Canceled"}};
        for (size_t regionIndex = 0u; regionIndex < _hostedTooltipDescriptors.size(); ++regionIndex)
        {
            const auto& descriptor = _hostedTooltipDescriptors[regionIndex];
            if (! descriptor.automationId.starts_with(L"FileOperations.CompletedGroup.Badge.") || descriptor.bounds.right <= descriptor.bounds.left ||
                descriptor.bounds.bottom <= descriptor.bounds.top)
                continue;
            ++result.completedGroupPaintedBadges;
            for (size_t index = 0u; index < badgeSuffixes.size(); ++index)
                if (descriptor.automationId == std::format(L"FileOperations.CompletedGroup.Badge.{}", badgeSuffixes[index]))
                {
                    result.completedGroupPaintedBadgeMask |= 1u << static_cast<unsigned int>(index);
                    result.completedGroupPaintedGlyphs[index] = descriptor.paintedGlyph;
                    if (regionIndex < _hostedTooltipRegions.size() && _hostedTooltipRegions[regionIndex])
                        result.completedGroupHostedNames[index] = _hostedTooltipRegions[regionIndex]->GetAccessibleName();
                }
            const float labelWidth              = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), descriptor.paintedLabel, 4096.0f, 128.0f);
            result.completedGroupBadgeCountsFit = result.completedGroupBadgeCountsFit && labelWidth > 0.0f &&
                                                  labelWidth <= descriptor.paintedLabelBounds.right - descriptor.paintedLabelBounds.left;
        }
    }

    for (const TaskSnapshot& task : taskSnapshot)
    {
        if (task.finished && IsTaskCollapsed(task.taskId))
        {
            ++result.completedAutoCollapsedCount;
        }
    }

    result.scanStatusMeasurementCount = _scanStatusMeasurementCount;
    result.scanStatusLayoutSlots      = _scanStatusLayouts.size();
    result.scanStatusRetainedLayouts =
        static_cast<size_t>(std::ranges::count_if(_scanStatusLayouts, [](const auto& entry) noexcept { return ! ! entry.prepared.layout; }));
    // Informational tasks are intentionally outside the file-operation snapshot projection.
    if (result.taskId != 0u)
    {
        result.found                    = std::ranges::any_of(taskSnapshot, [&result](const TaskSnapshot& task) noexcept {
            return task.kind == TaskSnapshot::Kind::Informational && task.taskId == result.taskId;
        });
        const std::wstring scanStatusId = std::format(L"FileOperations.Task.{}.ScanStatus", result.taskId);
        for (size_t index = 0u; index < _hostedTooltipDescriptors.size(); ++index)
        {
            const auto& region = _hostedTooltipDescriptors[index];
            if (region.automationId != scanStatusId)
                continue;
            const float width            = region.bounds.right - region.bounds.left;
            const float paintedWidth     = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), region.paintedPath, 65536.0f, 128.0f);
            result.taskScanStatusVisible = true;
            result.taskScanStatusFits = ! region.paintedPath.empty() && paintedWidth <= width + 0.5f &&
                                        region.bounds.left >= region.paintedContainerBounds.left &&
                                        region.bounds.right <= region.paintedContainerBounds.right && region.bounds.top >= region.paintedContainerBounds.top &&
                                        region.bounds.bottom <= region.paintedContainerBounds.bottom;
            result.taskScanStatusValidUnicode      = Common::Strings::TryUtf8FromUtf16Strict(region.paintedPath).has_value();
            result.taskScanStatusAbbreviated       = region.paintedPath.find(kEllipsisText) != std::wstring::npos;
            result.taskScanStatusNeedsAbbreviation = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), region.text, 65536.0f, 128.0f) > width;
            result.taskScanStatusText              = region.text;
            result.taskScanStatusCopyPath          = region.copyPath;
            result.taskScanStatusTooltipComplete =
                index < _hostedTooltipRegions.size() && _hostedTooltipRegions[index] && _hostedTooltipRegions[index]->GetTooltipText() == region.text;
        }
    }

    auto taskIt = result.taskId == 0
                      ? std::find_if(taskSnapshot.begin(),
                                     taskSnapshot.end(),
                                     [](const TaskSnapshot& task) noexcept { return task.kind == TaskSnapshot::Kind::FileOperation && ! task.finished; })
                      : std::find_if(taskSnapshot.begin(), taskSnapshot.end(), [&result](const TaskSnapshot& task) noexcept {
        return task.kind == TaskSnapshot::Kind::FileOperation && task.taskId == result.taskId;
    });
    if (result.taskId == 0 && taskIt == taskSnapshot.end())
    {
        taskIt = std::find_if(
            taskSnapshot.begin(), taskSnapshot.end(), [](const TaskSnapshot& task) noexcept { return task.kind == TaskSnapshot::Kind::FileOperation; });
    }
    if (taskIt != taskSnapshot.end())
    {
        result.taskId = taskIt->taskId;
        if (const auto* body = FindConflictBody(result.taskId))
        {
            result.conflictBodyBounds          = body->bounds;
            result.conflictPanelBounds         = body->panel ? body->panel->GetBounds() : D2D1::RectF();
            result.conflictBodyContentHeightPx = body->contentHeightPx;
            result.conflictBodyOffsetDip       = body->offsetDip;
        }
        result.found = true;
        // Observe the retained DxUi control after the renderer has synchronized it.
        // taskId selects the card; the value and state are never recomputed from the model.
        const auto progressId = std::format(L"FileOperations.Task.{}.Progress", taskIt->taskId);
        for (size_t i = 0u; i < _hostedProgressDescriptors.size() && i < _hostedProgressBars.size(); ++i)
        {
            const DxUi::ProgressBar* progress = _hostedProgressBars[i];
            if (! progress || _hostedProgressDescriptors[i].taskId != taskIt->taskId || progress->GetAccessibleAutomationId() != std::wstring_view(progressId))
                continue;
            result.taskHostedProgressPresent         = true;
            result.taskHostedProgressValue           = progress->GetValue();
            result.taskHostedProgressIndeterminate   = progress->IsIndeterminate();
            result.taskHostedProgressAnimationPaused = progress->IsAnimationPaused();
            result.taskHostedProgressEnabled         = progress->IsEnabled();
            result.taskHostedProgressSegmented       = progress->HasSegmentedValues();
            result.taskHostedProgressHelpText        = std::wstring(progress->GetAccessibleHelpText());
            if (i < _hostedProgressVerificationMarks.size() && _hostedProgressVerificationMarks[i])
            {
                const auto* const marks              = static_cast<const FileOperationsVerificationHalfMarks*>(_hostedProgressVerificationMarks[i]);
                result.taskVerificationMarksVisible  = marks->MarksShown();
                result.taskVerificationMarksPercent  = marks->VerifiedPercent();
                result.taskVerificationMarksOutlined = marks->OutlinesPending();
            }
            break;
        }
        D2D1_RECT_F paintedSummaryBounds{};
        const auto summaryId = std::format(L"FileOperations.Task.{}.ResultSummary", taskIt->taskId);
        for (const auto& descriptor : _hostedTooltipDescriptors)
        {
            if (descriptor.automationId != summaryId)
                continue;
            result.taskTerminalSummaryVisible = descriptor.text == taskIt->resultSummary && ! descriptor.text.empty();
            const auto& bounds                = descriptor.bounds;
            paintedSummaryBounds              = bounds;
            wil::com_ptr<IDWriteTextLayout> measured;
            DWRITE_TEXT_METRICS metrics{};
            const float availableWidth = bounds.right - bounds.left;
            const bool measuredOk =
                _dwriteFactory && _bodyFormat && availableWidth > 0.0f &&
                SUCCEEDED(_dwriteFactory->CreateTextLayout(
                    descriptor.text.data(), static_cast<UINT32>(descriptor.text.size()), _bodyFormat.get(), availableWidth, 4096.0f, measured.put())) &&
                measured && SUCCEEDED(measured->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP)) && SUCCEEDED(measured->GetMetrics(&metrics));
            result.taskTerminalSummaryFits =
                measuredOk && metrics.height > 0.0f && metrics.height <= bounds.bottom - bounds.top + 0.5f && metrics.width <= availableWidth + 0.5f;
            if (const auto painted = _paintedTaskBodiesForSelfTest.find(taskIt->taskId); painted != _paintedTaskBodiesForSelfTest.end())
            {
                const auto& card                          = painted->second.card;
                result.taskTerminalSummaryFits            = result.taskTerminalSummaryFits && bounds.left >= card.left && bounds.right <= card.right &&
                                                            bounds.top >= card.top && bounds.bottom <= card.bottom;
                result.taskTerminalSummaryPrecedesActions = bounds.bottom <= painted->second.actions.top;
            }
        }
        const auto factsId      = std::format(L"FileOperations.Task.{}.TerminalFacts", taskIt->taskId);
        size_t factsRegionCount = 0u;
        for (size_t index = 0u; index < _hostedTooltipDescriptors.size(); ++index)
        {
            const HostedTooltipDescriptor& descriptor = _hostedTooltipDescriptors[index];
            if (descriptor.automationId != factsId)
                continue;
            ++factsRegionCount;
            const DxUi::Control* region         = index < _hostedTooltipRegions.size() ? _hostedTooltipRegions[index] : nullptr;
            const D2D1_RECT_F& bounds           = descriptor.bounds;
            result.taskTerminalFactsText        = descriptor.text;
            result.taskTerminalFactsPaintedRect = bounds;
            if (region)
                result.taskTerminalFactsHostedName = std::wstring(region->GetAccessibleName());
            const auto close               = [](float a, float b) noexcept { return std::abs(a - b) <= 0.5f; };
            const D2D1_RECT_F hostedBounds = region ? region->GetBounds() : D2D1::RectF();
            const bool hostedBoundsMatch =
                region && close(hostedBounds.left, _controlsHost.PixelsToDip(bounds.left)) && close(hostedBounds.top, _controlsHost.PixelsToDip(bounds.top)) &&
                close(hostedBounds.right, _controlsHost.PixelsToDip(bounds.right)) && close(hostedBounds.bottom, _controlsHost.PixelsToDip(bounds.bottom));
            result.taskTerminalFactsVisible = region && ! descriptor.text.empty() && rectHasArea(bounds) && hostedBoundsMatch &&
                                              region->GetAccessibleAutomationId() == factsId && region->GetAccessibleName() == descriptor.text &&
                                              region->GetAccessibleHelpText() == descriptor.text && region->GetTooltipText() == descriptor.text;
            wil::com_ptr<IDWriteTextLayout> measured;
            DWRITE_TEXT_METRICS metrics{};
            const float availableWidth = bounds.right - bounds.left;
            const bool measuredOk =
                _dwriteFactory && _bodyFormat && availableWidth > 0.0f &&
                SUCCEEDED(_dwriteFactory->CreateTextLayout(
                    descriptor.text.data(), static_cast<UINT32>(descriptor.text.size()), _bodyFormat.get(), availableWidth, 4096.0f, measured.put())) &&
                measured && SUCCEEDED(measured->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP)) && SUCCEEDED(measured->GetMetrics(&metrics));
            result.taskTerminalFactsFits = result.taskTerminalFactsVisible && measuredOk && metrics.height > 0.0f &&
                                           metrics.height <= bounds.bottom - bounds.top + 0.5f && metrics.width <= availableWidth + 0.5f;
            if (const auto painted = _paintedTaskBodiesForSelfTest.find(taskIt->taskId); painted != _paintedTaskBodiesForSelfTest.end())
            {
                const auto& card                        = painted->second.card;
                result.taskTerminalFactsFits            = result.taskTerminalFactsFits && bounds.left >= card.left && bounds.right <= card.right &&
                                                          bounds.top >= card.top && bounds.bottom <= card.bottom;
                result.taskTerminalFactsFollowsSummary  = taskIt->resultSummary.empty()
                                                              ? bounds.top >= card.top + DipsToPixels(8.0f + 18.0f, _dpi)
                                                              : result.taskTerminalSummaryVisible && paintedSummaryBounds.bottom <= bounds.top;
                result.taskTerminalFactsPrecedesActions = bounds.bottom <= painted->second.actions.top;
            }
        }
        if (factsRegionCount != 1u)
            result.taskTerminalFactsVisible = result.taskTerminalFactsFits = false;
        if (const auto painted = _paintedTaskBodiesForSelfTest.find(taskIt->taskId); painted != _paintedTaskBodiesForSelfTest.end())
        {
            result.taskPaintedBodyBoundsValid = true;
            result.taskPaintedCardRect        = painted->second.card;
            result.taskPaintedBodyBottomPx    = painted->second.bodyBottom;
            result.taskPaintedActionRowRect   = painted->second.actions;
        }
        if (const auto painted = _paintedTaskCardsForSelfTest.find(taskIt->taskId); painted != _paintedTaskCardsForSelfTest.end())
            result.taskPaintedCardRect = painted->second;
        const auto noteId = std::format(L"FileOperations.Task.{}.LastNote", taskIt->taskId);
        for (const auto& descriptor : _hostedTooltipDescriptors)
        {
            if (descriptor.automationId != noteId)
                continue;
            result.taskLastNoteText = descriptor.text;
            result.taskLastNoteRect = descriptor.bounds;
            const auto& bounds      = descriptor.bounds;
            wil::com_ptr<IDWriteTextLayout> measured;
            DWRITE_TEXT_METRICS metrics{};
            const float width = bounds.right - bounds.left;
            const bool measuredOk =
                _dwriteFactory && _smallFormat && width > 0.0f &&
                SUCCEEDED(_dwriteFactory->CreateTextLayout(
                    descriptor.text.data(), static_cast<UINT32>(descriptor.text.size()), _smallFormat.get(), width, 4096.0f, measured.put())) &&
                SUCCEEDED(measured->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP)) && SUCCEEDED(measured->GetMetrics(&metrics));
            result.taskLastNoteLineCount = metrics.lineCount;
            result.taskLastNoteFits      = measuredOk && metrics.height > 0.0f && metrics.height <= bounds.bottom - bounds.top + 0.5f &&
                                           metrics.width <= width + 0.5f && bounds.left >= result.taskPaintedCardRect.left &&
                                           bounds.right <= result.taskPaintedCardRect.right && bounds.bottom <= result.taskPaintedCardRect.bottom &&
                                           bounds.bottom <= result.taskPaintedActionRowRect.top;
            for (const auto& progress : _hostedProgressDescriptors)
                if (progress.taskId == taskIt->taskId)
                    result.taskLastNotePrecedesProgress = bounds.bottom <= progress.bounds.top && progress.bounds.bottom <= result.taskPaintedActionRowRect.top;
        }
        if (const auto painted = _paintedLiveProgressForSelfTest.find(taskIt->taskId); painted != _paintedLiveProgressForSelfTest.end())
        {
            result.taskPaintGeneration               = _paintGenerationForSelfTest;
            result.taskPaintedFileProgress           = painted->second.files;
            result.taskPaintedFileProgressCount      = painted->second.fileCount;
            result.taskPaintedDiscoveryText          = painted->second.discoveryText;
            result.taskPaintedRecursiveCounts        = painted->second.recursiveCounts;
            result.taskPaintedTransferText           = painted->second.transferText;
            result.taskPaintedDiscoveryRect          = painted->second.discoveryRect;
            result.taskPaintedDiscoveryColor         = painted->second.discoveryColor;
            result.taskPaintedTransferRect           = painted->second.transferRect;
            result.taskPaintedTransferColor          = painted->second.transferColor;
            result.taskPaintedTransferCompletedBytes = painted->second.transferCompletedBytes;
        }
        if (const auto animation = _taskHeightAnimations.find(taskIt->taskId); animation != _taskHeightAnimations.end())
        {
            result.selectedTaskHeightMotionActive = scalarAnimationActive(animation->second.height);
            result.taskHeightAnimationUnsettled = animation->second.height.initialized && animation->second.height.displayed != animation->second.height.target;
            result.layoutMotionActive           = result.layoutMotionActive || result.selectedTaskHeightMotionActive;
        }
        result.taskDiscoveryAheadActive                              = taskIt->discoveryAheadActive;
        const WholeTaskProgressPresentation taskProgressPresentation = ResolveWholeTaskProgressPresentation(*taskIt);
        result.taskDiscoveryIndicatorVisible                         = taskProgressPresentation.showDiscoveryActivity;
        const std::wstring pathPrefix                                = std::format(L"FileOperations.Task.{}.Path.", taskIt->taskId);
        for (const auto& region : _hostedTooltipDescriptors)
        {
            if (! region.automationId.starts_with(pathPrefix))
                continue;
            const float paintedWidth = MeasureTextWidth(_dwriteFactory.get(), _bodyFormat.get(), region.paintedPath, 65536.0f, 128.0f);
            ++result.taskPaintedPathCount;
            result.taskPaintedPathsFit =
                result.taskPaintedPathsFit && ! region.paintedPath.empty() && paintedWidth <= region.bounds.right - region.bounds.left + 0.5f;
            result.taskPaintedPathsValidUnicode =
                result.taskPaintedPathsValidUnicode && Common::Strings::TryUtf8FromUtf16Strict(region.paintedPath).has_value();
            if (region.copyPath.ends_with(L".raw"))
                result.taskPaintedPathsRetainRawExtension = result.taskPaintedPathsRetainRawExtension && region.paintedPath.ends_with(L".raw");
            if (region.paintedLabel.empty() || ! rectHasArea(region.paintedLabelBounds))
                continue;
            const auto& labelBounds   = region.paintedLabelBounds;
            const float requiredWidth = MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), region.paintedLabel, 4096.0f, 128.0f);
            const float slack         = labelBounds.right - labelBounds.left - requiredWidth;
            result.taskPaintedPathLabelMinimumSlackPx =
                result.taskPaintedPathLabelCount == 0u ? slack : (std::min)(result.taskPaintedPathLabelMinimumSlackPx, slack);
            ++result.taskPaintedPathLabelCount;
            result.taskPaintedPathLabelsFit = result.taskPaintedPathLabelsFit && requiredWidth > 0.0f && slack >= -0.5f &&
                                              labelBounds.right <= region.bounds.left && labelBounds.top == region.bounds.top &&
                                              labelBounds.bottom == region.bounds.bottom;
        }
        result.taskTransferProgressVisible        = taskProgressPresentation.showTransferProgress;
        result.taskTransferProgressDeterminate    = result.taskTransferProgressVisible && taskProgressPresentation.determinate;
        result.taskTransferProgressProvisional    = result.taskTransferProgressDeterminate && taskProgressPresentation.provisional;
        result.taskTransferProgressMarqueeVisible = taskProgressPresentation.showTransferMarquee;
        const TaskStatusKind status               = taskIt->statusKind != TaskStatusKind::None ? taskIt->statusKind : ResolveTaskStatusKind(*taskIt);
        result.taskStatusKind                     = status;
        result.taskStatusActiveStateCount         = SurfacedTaskStatusCount(status);
        bool graphShowAnimation                   = false;
        static_cast<void>(GraphOverlayTextForStatus(*taskIt, status, graphShowAnimation));
        result.graphStatusAnimationEnabled        = graphShowAnimation && ! reducedMotion;
        result.conflictStackedPathRows            = taskIt->conflict.active;
        result.conflictDefaultAction              = taskIt->conflict.defaultAction;
        result.conflictEscapeAction               = taskIt->conflict.escapeAction;
        result.conflictButtonsPublishable         = taskIt->conflict.buttonsPublishable;
        result.conflictSkipAllEligible            = taskIt->conflict.skipAllEligible;
        result.conflictSourceMetadataVisible      = taskIt->conflict.active && taskIt->conflict.sourceMetadata.available;
        result.conflictDestinationMetadataVisible = taskIt->conflict.active && taskIt->conflict.destinationMetadata.available;
        result.conflictMetadataSizeCompareVisible =
            taskIt->conflict.active && taskIt->conflict.sourceMetadata.sizeKnown && taskIt->conflict.destinationMetadata.sizeKnown;
        result.conflictMetadataDateCompareVisible =
            taskIt->conflict.active && taskIt->conflict.sourceMetadata.lastWriteTime > 0 && taskIt->conflict.destinationMetadata.lastWriteTime > 0;
        result.conflictDiscoveryIndicatorVisible = taskIt->conflict.active && taskProgressPresentation.showDiscoveryActivity;
        result.conflictTransferProgressVisible   = taskIt->conflict.active && taskProgressPresentation.showTransferProgress;
        if (taskIt->conflict.active)
        {
            const std::wstring semanticPrefix = std::format(L"FileOperations.Task.{}.Conflict.", taskIt->taskId);
            const auto findSemanticRegion     = [&](std::wstring_view suffix) noexcept -> const HostedTooltipDescriptor*
            {
                const std::wstring automationId = std::format(L"{0}{1}", semanticPrefix, suffix);
                const auto descriptor           = std::ranges::find_if(_hostedTooltipDescriptors, [&](const HostedTooltipDescriptor& value) noexcept {
                    return value.automationId == automationId && rectHasArea(value.bounds) && ! value.text.empty();
                });
                return descriptor != _hostedTooltipDescriptors.end() ? &*descriptor : nullptr;
            };

            const HostedTooltipDescriptor* stateRegion    = findSemanticRegion(L"State");
            const HostedTooltipDescriptor* questionRegion = findSemanticRegion(L"Question");
            const HostedTooltipDescriptor* factsRegion    = findSemanticRegion(L"Facts");
            const HostedTooltipDescriptor* fromRegion     = findSemanticRegion(taskIt->operation == FILESYSTEM_DELETE ? L"Deleting" : L"From");
            const HostedTooltipDescriptor* toRegion       = findSemanticRegion(L"To");
            const HostedTooltipDescriptor* lastNoteRegion = findSemanticRegion(L"LastNote");

            if (const auto* inspectionRegion = findSemanticRegion(L"InspectionResult"))
            {
                const auto& bounds         = inspectionRegion->bounds;
                const float requiredHeight = MeasureWrappedTextHeight(
                    _dwriteFactory.get(), _smallFormat.get(), inspectionRegion->text, bounds.right - bounds.left, DipsToPixels(18.0f, _dpi));
                result.taskInspectionResultFits = requiredHeight <= bounds.bottom - bounds.top + 0.5f;
            }
            result.conflictWaitingForDecisionVisible     = ! taskIt->conflict.metadataLoading && stateRegion != nullptr && questionRegion != nullptr;
            result.conflictDecisionDetailsLoadingVisible = taskIt->conflict.metadataLoading && stateRegion != nullptr;
            result.conflictPromptVisibleWithoutClipping  = stateRegion != nullptr && (taskIt->conflict.metadataLoading || questionRegion != nullptr);
            result.conflictContextVisibleWithoutClipping =
                fromRegion != nullptr && (taskIt->operation == FILESYSTEM_DELETE || taskIt->conflict.destinationPath.empty() || toRegion != nullptr) &&
                ((! taskIt->conflict.factItemCountKnown && ! taskIt->conflict.factBytesKnown) || factsRegion != nullptr);
            result.lastNoteVisibleWithoutClipping = taskIt->lastDiagnosticMessage.empty() || lastNoteRegion != nullptr;

            float promptBottom = 0.0f;
            for (const HostedTooltipDescriptor& descriptor : _hostedTooltipDescriptors)
            {
                if (descriptor.automationId.starts_with(semanticPrefix))
                {
                    result.conflictPromptContentBottomPx = std::max(result.conflictPromptContentBottomPx, descriptor.bounds.bottom);
                    auto visibleBounds                   = descriptor.bounds;
                    if (descriptor.comparisonTaskId != 0u)
                    {
                        // Comparison content uses scroll-content coordinates; the
                        // renderer clips it to this viewport before drawing actions.
                        // Off-screen rows must not be mistaken for visible overlap.
                        const auto* body = FindConflictBody(descriptor.comparisonTaskId);
                        if (! body || ! body->visible)
                            continue;
                        visibleBounds.left   = std::max(visibleBounds.left, body->viewport.left);
                        visibleBounds.top    = std::max(visibleBounds.top, body->viewport.top);
                        visibleBounds.right  = std::min(visibleBounds.right, body->viewport.right);
                        visibleBounds.bottom = std::min(visibleBounds.bottom, body->viewport.bottom);
                    }
                    if (rectHasArea(visibleBounds))
                        promptBottom = std::max(promptBottom, visibleBounds.bottom);
                }
            }
            float actionTop = (std::numeric_limits<float>::max)();
            for (const PopupButton& button : _buttons)
            {
                if (button.hit.taskId == taskIt->taskId &&
                    (button.hit.kind == PopupHitTest::Kind::TaskConflictAction || button.hit.kind == PopupHitTest::Kind::TaskConflictMore ||
                     button.hit.kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll))
                {
                    actionTop = std::min(actionTop, button.bounds.top);
                }
            }
            result.conflictPromptVisibleBottomPx = promptBottom;
            result.conflictActionsTopPx          = actionTop == (std::numeric_limits<float>::max)() ? 0.0f : actionTop;
            result.conflictPromptPrecedesActions =
                promptBottom > 0.0f && (actionTop == (std::numeric_limits<float>::max)() || promptBottom <= actionTop + 0.5f);
        }
        result.taskDuplicateUnderGraphItemBarVisible = false;
        const PopupStatusVisualTone statusTone       = StatusVisualToneForTaskStatus(status);
        const std::wstring statusChipText            = StatusChipTextForTask(*taskIt, status, GetTickCount64());
        result.taskStatusVisualTone                  = static_cast<uint32_t>(statusTone);
        result.taskStatusVisualColorRef              = ColorToCOLORREF(StatusVisualColorForTone(appTheme, statusTone));
        result.taskStatusStripeVisible               = statusTone != PopupStatusVisualTone::None;
        result.taskStatusChipVisible                 = result.taskStatusStripeVisible && ! statusChipText.empty();
        if (const auto label = _paintedCompactMeterLabelsForSelfTest.find(taskIt->taskId); label != _paintedCompactMeterLabelsForSelfTest.end())
            result.taskCompactMeterLabel = label->second;
        if (const auto painted = _paintedStatusChipsForSelfTest.find(taskIt->taskId); painted != _paintedStatusChipsForSelfTest.end())
        {
            const auto& chip                     = painted->second;
            result.taskStatusHeaderPaintedText   = chip.headerText;
            result.taskSubjectPaintedText        = chip.paintedHeaderText;
            result.taskSubjectRect               = chip.headerRect;
            result.taskStatusChipActuallyPainted = chip.painted;
            if (chip.painted)
            {
                result.taskStatusChipPaintedRect                 = chip.chrome;
                result.taskStatusChipPaintedTextRect             = chip.textRect;
                result.taskStatusChipPaintedText                 = chip.text;
                const float width                                = chip.textRect.right - chip.textRect.left;
                const float height                               = chip.textRect.bottom - chip.textRect.top;
                result.taskStatusChipPaintedTextAvailableWidthPx = width;
                if (_dwriteFactory && chip.paintedFormat && ! chip.text.empty() && width > 0.0f && height > 0.0f)
                {
                    wil::com_ptr<IDWriteTextLayout> measured;
                    DWRITE_TEXT_METRICS metrics{};
                    const bool measuredOk =
                        SUCCEEDED(_dwriteFactory->CreateTextLayout(
                            chip.text.data(), static_cast<UINT32>(chip.text.size()), chip.paintedFormat.get(), width, height, measured.put())) &&
                        measured && SUCCEEDED(measured->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP)) && SUCCEEDED(measured->GetMetrics(&metrics));
                    if (measuredOk)
                    {
                        result.taskStatusChipPaintedTextWidthPx = metrics.widthIncludingTrailingWhitespace;
                        result.taskStatusChipPaintedTextFits    = metrics.lineCount == 1u && metrics.height <= height + 0.5f && metrics.left >= -0.5f &&
                                                                  metrics.left + metrics.widthIncludingTrailingWhitespace <= width + 0.5f &&
                                                                  chip.textRect.left >= chip.chrome.left && chip.textRect.top >= chip.chrome.top &&
                                                                  chip.textRect.right <= chip.chrome.right && chip.textRect.bottom <= chip.chrome.bottom;
                    }
                }
            }
        }
        result.taskStatusGlyphSignalVisible = StatusIsOk(status) || StatusIsWarning(status) || StatusIsError(status);
        result.taskStatusTextSignalVisible  = ! statusChipText.empty() || ! StatusTextForTask(*taskIt, status, GetTickCount64()).empty();
        result.taskStatusColorBlindSafeEncoding =
            statusTone == PopupStatusVisualTone::None || result.taskStatusGlyphSignalVisible || result.taskStatusTextSignalVisible;
        // The old Failed items entry moved out of More. Do not infer menu membership
        // from warning/error counters in this rendered-controls observer.
        result.completedFailedItemsActionVisible       = false;
        result.completedOpenDestinationActionVisible   = CompletedTaskCanUseDestinationActions(*taskIt);
        result.completedRevealDestinationActionVisible = ResolveCompletedTaskRevealLocation(*taskIt).has_value();
        result.taskHiddenByCompletedGroup              = completedGroupVisible && ! _completedGroupExpanded && IsCompletedGroupTask(*taskIt);
        result.taskCollapsed                           = IsTaskCollapsed(taskIt->taskId);
        result.taskThroughputExpanded                  = IsTaskThroughputExpanded(taskIt->taskId);
        result.taskThroughputSummaryVisible            = std::ranges::any_of(_hostedTooltipDescriptors, [&](const HostedTooltipDescriptor& descriptor) {
            return descriptor.automationId == std::format(L"FileOperations.Task.{}.ThroughputSummary", taskIt->taskId);
        });
        result.taskCompactRow                          = IsTaskCollapsedForDisplay(taskIt->taskId, result.popupCompactDensity);
        result.taskAutoCollapsedOnCompletion           = taskIt->finished && result.taskCollapsed;
        result.taskCompactProgressVisible =
            result.taskCompactRow && taskIt->kind == TaskSnapshot::Kind::FileOperation &&
            ((taskIt->finished && TaskHasKnownCompactProgress(*taskIt)) ||
             (taskProgressPresentation.showTransferProgress && (TaskHasKnownCompactProgress(*taskIt) || taskProgressPresentation.showTransferMarquee)));
        result.taskInlineSpeedRowVisible = std::ranges::any_of(_hostedTooltipDescriptors, [&](const HostedTooltipDescriptor& descriptor) {
            return descriptor.automationId == std::format(L"FileOperations.Task.{}.DeletionRate", taskIt->taskId);
        });
        result.taskInlineEtaRowVisible   = false;
        result.taskExpandedBaseHeightDip =
            taskIt->finished
                ? (result.taskPaintedBodyBoundsValid ? _controlsHost.PixelsToDip(result.taskPaintedCardRect.bottom - result.taskPaintedCardRect.top) : 0.0f)
            : taskIt->operation == FILESYSTEM_DELETE                                         ? kFileOperationsExpandedDeleteCardHeightDip
            : (taskIt->operation == FILESYSTEM_COPY || taskIt->operation == FILESYSTEM_MOVE) ? kFileOperationsExpandedTransferCardHeightDip
                                                                                             : kFileOperationsExpandedCardHeightDip;
        if (taskProgressPresentation.showTransferProgress)
        {
            result.taskUnderGraphProgressBarCount = 1u;
        }

        result.conflictOverflowActionCount = std::min(taskIt->conflict.overflowActionCount, taskIt->conflict.overflowActions.size());
        for (size_t i = 0; i < result.conflictOverflowActionCount && i < result.conflictOverflowActions.size(); ++i)
        {
            result.conflictOverflowActions[i] = taskIt->conflict.overflowActions[i];
        }

        // Aggregate the live graph hue distribution so the fairness selftest can assert on the
        // REAL pipeline (samples, eviction, multi-bucket flushes) instead of synthetic weights.
        if (const auto rateIt = _rates.find(result.taskId); rateIt != _rates.end())
        {
            PopulateGraphHueDebugSummary(rateIt->second, result);
            result.graphCurrentBandwidthBytesPerSecond = CurrentBandwidthForGraphMarker(rateIt->second);
            result.graphCurrentBandwidthLineVisible    = result.graphCurrentBandwidthBytesPerSecond > 0.0;
            for (size_t offset = 0u; offset < rateIt->second.count; ++offset)
            {
                const size_t slot = (rateIt->second.writeIndex + RateHistory::kMaxSamples - 1u - offset) % RateHistory::kMaxSamples;
                result.graphStalledSampleCount += rateIt->second.stalledSamples[slot] != 0u ? 1u : 0u;
            }
        }
        const auto graphIt = std::find_if(_hostedGraphDescriptors.begin(),
                                          _hostedGraphDescriptors.end(),
                                          [&result](const HostedGraphDescriptor& descriptor) noexcept { return descriptor.taskId == result.taskId; });
        if (graphIt != _hostedGraphDescriptors.end())
        {
            result.graphCurrentBandwidthLabelVisible = ! graphIt->currentBandwidthText.empty();
            result.graphEtaLabelVisible              = ! graphIt->currentBandwidthTrailingText.empty();
            result.graphEtaLabelRightAligned         = result.graphEtaLabelVisible;
            result.graphPlotRect                     = graphIt->bounds;
            result.graphRateLabelRect                = graphIt->rateLabelBounds;
            result.graphEtaLabelRect                 = graphIt->etaLabelBounds;
            result.graphStallBandRect                = graphIt->stallBandBounds;
            result.graphLabelRowBottomPx             = graphIt->labelRowBottom;
            result.graphLabelsOutsidePlot            = graphIt->separateLabelRow && rectHasArea(graphIt->bounds) && rectHasArea(graphIt->rateLabelBounds) &&
                                                       rectHasArea(graphIt->etaLabelBounds) && graphIt->rateLabelBounds.bottom < graphIt->bounds.top &&
                                                       graphIt->etaLabelBounds.bottom < graphIt->bounds.top;
            const size_t index                       = static_cast<size_t>(std::distance(_hostedGraphDescriptors.begin(), graphIt));
            if (index < _graphLabelLayouts.size())
            {
                const auto fits = [&result](IDWriteTextLayout* layout, const D2D1_RECT_F& bounds)
                {
                    DWRITE_TEXT_METRICS metrics{};
                    if (! layout || FAILED(layout->GetMetrics(&metrics)))
                        return false;
                    // DirectWrite includes line-ending spaces in widthIncludingTrailingWhitespace.
                    // Visible ink must fit; invisible wrap whitespace is recorded separately.
                    result.graphLabelVisibleOverhangPx = std::max({result.graphLabelVisibleOverhangPx,
                                                                   metrics.left + metrics.width - (bounds.right - bounds.left),
                                                                   metrics.top + metrics.height - (bounds.bottom - bounds.top)});
                    result.graphLabelWhitespaceOverhangPx =
                        std::max(result.graphLabelWhitespaceOverhangPx, metrics.left + metrics.widthIncludingTrailingWhitespace - (bounds.right - bounds.left));
                    return metrics.left >= 0.0f && metrics.top >= 0.0f && metrics.left + metrics.width <= bounds.right - bounds.left + 0.5f &&
                           metrics.top + metrics.height <= bounds.bottom - bounds.top + 0.5f;
                };
                const size_t taskIndex = static_cast<size_t>(std::distance(taskSnapshot.begin(), taskIt));
                if (graphIt->labelsInTaskSummary && taskIndex < _cardInformationLayouts.size())
                {
                    const auto& information = _cardInformationLayouts[taskIndex];
                    result.graphLabelsFit =
                        fits(information.throughput.eta.get(), graphIt->rateLabelBounds) && fits(information.counts.eta.get(), graphIt->etaLabelBounds);
                }
                else
                    result.graphLabelsFit = fits(_graphLabelLayouts[index].rate.get(), graphIt->rateLabelBounds) &&
                                            fits(_graphLabelLayouts[index].eta.get(), graphIt->etaLabelBounds);
            }
        }

        result.failureSurfaceActive     = _failureSurface;
        result.failureSurfaceReady      = _failureSurfaceReady;
        result.failureSurfaceText       = _failureTextValue;
        result.failureSurfaceChildCount = static_cast<size_t>(IsActuallyVisibleChildWindow(_failureText)) +
                                          static_cast<size_t>(IsActuallyVisibleChildWindow(_failureCancelAll)) +
                                          static_cast<size_t>(IsActuallyVisibleChildWindow(_failureClose));
        if (const auto matches = _graphRowColorMatchCountsForSelfTest.find(result.taskId); matches != _graphRowColorMatchCountsForSelfTest.end())
        {
            result.graphRowColorMatchCount = matches->second;
        }
        if (const auto mismatches = _graphRowColorMismatchCountsForSelfTest.find(result.taskId); mismatches != _graphRowColorMismatchCountsForSelfTest.end())
        {
            result.graphRowColorMismatchCount = mismatches->second;
        }
    }

    const auto appendPrimaryAction = [&](uint8_t rawAction) noexcept
    {
        if (result.conflictPrimaryActionCount < result.conflictPrimaryActions.size())
        {
            result.conflictPrimaryActions[result.conflictPrimaryActionCount] = rawAction;
        }
        ++result.conflictPrimaryActionCount;
    };

    const auto rectsOverlap = [&](const D2D1_RECT_F& lhs, const D2D1_RECT_F& rhs) noexcept
    {
        if (! rectHasArea(lhs) || ! rectHasArea(rhs))
        {
            return false;
        }

        return lhs.left < rhs.right && rhs.left < lhs.right && lhs.top < rhs.bottom && rhs.top < lhs.bottom;
    };

    const auto visibleButtonBounds = [&](const PopupButton& button) noexcept
    {
        auto bounds = button.bounds;
        if (button.hit.kind == PopupHitTest::Kind::TaskConflictInspectSource || button.hit.kind == PopupHitTest::Kind::TaskConflictInspectDestination)
        {
            if (const auto* body = FindConflictBody(button.hit.taskId); body && body->visible)
            {
                bounds.left   = std::max(bounds.left, body->viewport.left);
                bounds.top    = std::max(bounds.top, body->viewport.top);
                bounds.right  = std::min(bounds.right, body->viewport.right);
                bounds.bottom = std::min(bounds.bottom, body->viewport.bottom);
            }
        }
        return bounds;
    };
    result.visibleButtonCount     = _buttons.size();
    const auto hasActiveHitTarget = [&](const PopupButton& button) noexcept
    {
        if (! rectHasArea(button.bounds))
        {
            return false;
        }
        const PopupHitTest hit = HitTest((button.bounds.left + button.bounds.right) * 0.5f, (button.bounds.top + button.bounds.bottom) * 0.5f);
        return hit.kind == button.hit.kind && hit.taskId == button.hit.taskId && hit.data == button.hit.data &&
               hit.promptGeneration == button.hit.promptGeneration;
    };
    const auto hostedLabelFits = [&](const DxUi::Button& action, const D2D1_RECT_F& bounds, bool scope) noexcept
    {
        // The same negative witness builds against the original pin and
        // the qualified multiline API, without changing the assertion.
        const bool multiline = []<typename T>(const T& control) noexcept
        {
            if constexpr (requires { control.IsMultiline(); })
                return control.IsMultiline();
            else
                return false;
        }(action);
        const float widthDip  = _controlsHost.PixelsToDip(bounds.right - bounds.left) - (scope ? 36.0f : (multiline ? 24.0f : 0.0f));
        const float heightDip = _controlsHost.PixelsToDip(bounds.bottom - bounds.top) - (scope ? 6.0f : (multiline ? 16.0f : 0.0f));
        const auto text       = action.GetText();
        auto* format          = _controlsHost.GetTextFormat(DxUi::FontRole::Body, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, multiline);
        wil::com_ptr<IDWriteTextLayout> layout;
        DWRITE_TEXT_METRICS metrics{};
        return widthDip > 0 && heightDip > 0 && format && _dwriteFactory &&
               SUCCEEDED(_dwriteFactory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, widthDip, 4096.0f, layout.put())) &&
               SUCCEEDED(layout->GetMetrics(&metrics)) && metrics.widthIncludingTrailingWhitespace <= widthDip + 0.5f && metrics.height <= heightDip + 0.5f;
    };
    const auto completedActionFits = [&](size_t index, const PopupButton& button, bool& labelFits) noexcept
    {
        labelFits = false;
        if (index >= _hostedButtons.size() || ! _hostedButtons[index] || ! rectHasArea(button.bounds))
            return false;
        const DxUi::Button& hosted      = *_hostedButtons[index];
        const bool more                 = button.hit.kind == PopupHitTest::Kind::TaskCompletedMore;
        const UINT labelId              = static_cast<UINT>(button.hit.kind == PopupHitTest::Kind::TaskIssues ? IDS_FILEOPS_BTN_ISSUES
                                                            : more                                            ? IDS_FILEOPS_CONFLICT_BTN_MORE
                                                                                                              : IDS_FILEOP_BTN_DISMISS);
        const std::wstring label        = LoadStringResource(nullptr, labelId);
        const std::wstring expectedText = more ? L"\u2026" : label;
        const auto actualBounds         = hosted.GetBounds();
        const auto boundsNear           = [](float a, float b) noexcept { return std::abs(a - b) <= 0.5f; };
        const bool hostedBoundsMatch    = boundsNear(actualBounds.left, _controlsHost.PixelsToDip(button.bounds.left)) &&
                                          boundsNear(actualBounds.top, _controlsHost.PixelsToDip(button.bounds.top)) &&
                                          boundsNear(actualBounds.right, _controlsHost.PixelsToDip(button.bounds.right)) &&
                                          boundsNear(actualBounds.bottom, _controlsHost.PixelsToDip(button.bounds.bottom));
        const auto painted              = _paintedTaskBodiesForSelfTest.find(button.hit.taskId);
        const bool insideCard = painted != _paintedTaskBodiesForSelfTest.end() && button.bounds.left >= painted->second.actions.left - 0.5f &&
                                button.bounds.top >= painted->second.actions.top - 0.5f && button.bounds.right <= painted->second.actions.right + 0.5f &&
                                button.bounds.bottom <= painted->second.actions.bottom + 0.5f && button.bounds.left >= painted->second.card.left - 0.5f &&
                                button.bounds.right <= painted->second.card.right + 0.5f && button.bounds.bottom <= painted->second.card.bottom + 0.5f;
        if (more)
        {
            const auto chrome = DxUi::ComputeButtonChromeLayout(actualBounds, DxUi::ButtonVariant::DropDown);
            const float width =
                MeasureTextWidth(_dwriteFactory.get(),
                                 _controlsHost.GetTextFormat(DxUi::FontRole::Body, DWRITE_TEXT_ALIGNMENT_CENTER, DWRITE_PARAGRAPH_ALIGNMENT_CENTER, false),
                                 hosted.GetText(),
                                 4096.0f,
                                 128.0f);
            labelFits = width > 0.0f && width <= chrome.textRect.right - chrome.textRect.left + 0.5f;
        }
        else
            labelFits = hostedLabelFits(hosted, button.bounds, false);
        const std::wstring automationId = PopupActionAutomationId(button.hit);
        return ! label.empty() && hostedBoundsMatch && insideCard && hasActiveHitTarget(button) && labelFits &&
               hosted.GetVariant() == (more ? DxUi::ButtonVariant::DropDown : DxUi::ButtonVariant::Standard) && hosted.GetText() == expectedText &&
               hosted.GetAccessibleName() == label && hosted.GetAccessibleAutomationId() == automationId;
    };
    size_t completedActionOrdinal = 0u;
    for (size_t i = 0; i < _buttons.size(); ++i)
    {
        const PopupButton& button = _buttons[i];
        switch (button.hit.kind)
        {
            case PopupHitTest::Kind::FooterCancelAll:
            case PopupHitTest::Kind::FooterPauseResumeAll:
                ++result.footerVisibleButtonCount;
                result.footerActionHitTargetsActive = result.footerActionHitTargetsActive && hasActiveHitTarget(button);
                if (i < _hostedButtons.size() && _hostedButtons[i])
                {
                    const float labelWidthDip = MeasureTextWidth(
                        _dwriteFactory.get(), _controlsHost.GetTextFormat(DxUi::FontRole::Body), _hostedButtons[i]->GetText(), 4096.0f, 128.0f);
                    const float paddingPx = button.bounds.right - button.bounds.left - DipsToPixels(labelWidthDip, _dpi);
                    result.footerActionMinimumPaddingPx =
                        result.footerMeasuredActionCount == 0u ? paddingPx : std::min(result.footerActionMinimumPaddingPx, paddingPx);
                    ++result.footerMeasuredActionCount;
                }
                break;
            case PopupHitTest::Kind::FooterAutoDismiss:
            case PopupHitTest::Kind::FooterDensity:
            case PopupHitTest::Kind::FooterOptions: ++result.footerVisibleButtonCount; break;
            case PopupHitTest::Kind::FooterToggleDetails:
                ++result.footerVisibleButtonCount;
                result.footerDetailsToggleUsesDisclosureChrome =
                    i < _hostedButtons.size() && _hostedButtons[i] && _hostedButtons[i]->GetVariant() == DxUi::ButtonVariant::Disclosure;
                break;
            case PopupHitTest::Kind::FooterQueueMode:
                ++result.footerVisibleButtonCount;
                result.footerQueueModeUsesSelectorChrome =
                    i < _hostedButtons.size() && _hostedButtons[i] && _hostedButtons[i]->GetVariant() == DxUi::ButtonVariant::Selector;
                break;
            case PopupHitTest::Kind::CompletedGroupToggle: result.completedGroupToggleVisible = true; break;
            case PopupHitTest::Kind::CompletedGroupClear:
                result.completedGroupClearVisible      = true;
                result.completedGroupClearTooltipEmpty = i < _hostedButtons.size() && _hostedButtons[i] && _hostedButtons[i]->GetTooltipText().empty();
                if (i < _hostedButtons.size() && _hostedButtons[i])
                {
                    const auto* control = _hostedButtons[i];
                    const auto chrome   = DxUi::ComputeButtonChromeLayout(control->GetBounds(), control->GetVariant());
                    const float labelWidth =
                        MeasureTextWidth(_dwriteFactory.get(), _controlsHost.GetTextFormat(DxUi::FontRole::Body), control->GetText(), 4096.0f, 128.0f);
                    result.completedGroupClearLabelFits = labelWidth > 0.0f && labelWidth <= chrome.textRect.right - chrome.textRect.left;
                }
                break;
            case PopupHitTest::Kind::TaskConflictToggleApplyToAll:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.conflictApplyToAllVisible = true;
                    if (i < _hostedButtons.size())
                    {
                        if (const auto* checkbox = dynamic_cast<const DxUi::Checkbox*>(_hostedButtons[i]))
                        {
                            result.conflictApplyToAllIsCheckbox = true;
                            result.conflictApplyToAllChecked    = checkbox->IsChecked();
                            result.conflictScopeLabelFits       = hostedLabelFits(*checkbox, button.bounds, true) &&
                                                                  checkbox->GetText() == LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_APPLY_TO_ALL);
                        }
                    }
                }
                break;
            case PopupHitTest::Kind::TaskConflictAction:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    appendPrimaryAction(static_cast<uint8_t>(button.hit.data));
                    if (i < _hostedButtons.size() && _hostedButtons[i])
                    {
                        const auto& action              = *_hostedButtons[i];
                        result.conflictPrimaryLabelsFit = result.conflictPrimaryLabelsFit && hostedLabelFits(action, button.bounds, false);
                        ++result.conflictMeasuredPrimaryCount;
                    }
                }
                break;
            case PopupHitTest::Kind::TaskConflictMore:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.conflictMoreVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskShowLog:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.completedShowLogVisible = true;
                    ++result.completedVisibleActionCount;
                }
                break;
            case PopupHitTest::Kind::TaskExportIssues:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.completedExportIssuesVisible = true;
                    ++result.completedVisibleActionCount;
                }
                break;
            case PopupHitTest::Kind::TaskCompletedMore:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.completedActionRowFits = result.completedActionRowFits && completedActionOrdinal == (result.completedIssuesVisible ? 2u : 1u);
                    ++completedActionOrdinal;
                    result.completedDiagnosticsMoreVisible           = true;
                    result.completedDiagnosticsMoreButtonRectVisible = true;
                    result.completedDiagnosticsMoreButtonRect        = button.bounds;
                    bool labelFits                                   = false;
                    result.completedActionRowFits                    = result.completedActionRowFits && completedActionFits(i, button, labelFits);
                    ++result.completedVisibleActionCount;
                }
                break;
            case PopupHitTest::Kind::TaskIssues:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.completedActionRowFits = result.completedActionRowFits && completedActionOrdinal == 1u;
                    ++completedActionOrdinal;
                    result.completedIssuesVisible         = true;
                    result.completedIssuesButtonRect      = button.bounds;
                    result.completedIssuesHitTargetActive = hasActiveHitTarget(button);
                    bool labelFits                        = false;
                    const bool actionFits                 = completedActionFits(i, button, labelFits);
                    result.completedIssuesLabelFits       = labelFits;
                    result.completedActionRowFits         = result.completedActionRowFits && actionFits;
                    if (i < _hostedButtons.size() && _hostedButtons[i])
                    {
                        result.completedIssuesHostedText = std::wstring(_hostedButtons[i]->GetText());
                        result.completedIssuesHostedName = std::wstring(_hostedButtons[i]->GetAccessibleName());
                    }
                    ++result.completedVisibleActionCount;
                }
                break;
            case PopupHitTest::Kind::TaskDismiss:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.completedActionRowFits = result.completedActionRowFits && completedActionOrdinal == 0u;
                    ++completedActionOrdinal;
                    result.completedDismissVisible = true;
                    bool labelFits                 = false;
                    result.completedActionRowFits  = result.completedActionRowFits && completedActionFits(i, button, labelFits);
                    ++result.completedVisibleActionCount;
                }
                break;
            case PopupHitTest::Kind::TaskToggleCollapse:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskToggleCollapseVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskToggleThroughput:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskThroughputDisclosureVisible = rectHasArea(button.bounds);
                    result.taskThroughputDisclosureRect    = button.bounds;
                }
                break;
            case PopupHitTest::Kind::TaskStartNow:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskStartNowVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskQueueMoveUp:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskQueueMoveUpVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskQueueMoveDown:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskQueueMoveDownVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskPause:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskPauseVisible         = rectHasArea(button.bounds);
                    result.taskPauseHitTargetActive = hasActiveHitTarget(button);
                }
                break;
            case PopupHitTest::Kind::TaskCancel:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskCancelVisible         = rectHasArea(button.bounds);
                    result.taskCancelHitTargetActive = hasActiveHitTarget(button);
                }
                break;
            case PopupHitTest::Kind::TaskSkip:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskSkipVisible         = rectHasArea(button.bounds);
                    result.taskSkipHitTargetActive = hasActiveHitTarget(button);
                    result.taskSkipControlWidthPx  = button.bounds.right - button.bounds.left;
                    if (i < _hostedButtons.size() && _hostedButtons[i])
                    {
                        // Measure the actual retained Button's body typography, not the legacy
                        // small format or a duplicate of the allocation algorithm being tested.
                        const float labelWidthDip = MeasureTextWidth(
                            _dwriteFactory.get(), _controlsHost.GetTextFormat(DxUi::FontRole::Body), _hostedButtons[i]->GetText(), 4096.0f, 128.0f);
                        result.taskSkipLabelWidthPx = DipsToPixels(labelWidthDip, _dpi);
                    }
                }
                break;
            case PopupHitTest::Kind::TaskSpeedLimit:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskSpeedLimitVisible         = rectHasArea(button.bounds);
                    result.taskSpeedLimitHitTargetActive = hasActiveHitTarget(button);
                    result.taskSpeedLimitUsesSelectorChrome =
                        i < _hostedButtons.size() && _hostedButtons[i] && _hostedButtons[i]->GetVariant() == DxUi::ButtonVariant::Selector;
                    result.taskSpeedLimitShowsCurrentValue =
                        i < _hostedButtons.size() && _hostedButtons[i] && taskIt != taskSnapshot.end() &&
                        _hostedButtons[i]->GetText() == FormatSpeedLimitSelectorText(taskIt->desiredSpeedLimitBytesPerSecond);
                    if (i < _hostedButtons.size() && _hostedButtons[i])
                    {
                        const auto* control = _hostedButtons[i];
                        const auto chrome   = DxUi::ComputeButtonChromeLayout(control->GetBounds(), control->GetVariant());
                        const float labelWidthDip =
                            MeasureTextWidth(_dwriteFactory.get(), _controlsHost.GetTextFormat(DxUi::FontRole::Body), control->GetText(), 4096.0f, 128.0f);
                        result.taskSpeedLimitLabelWidthPx = DipsToPixels(labelWidthDip, _dpi);
                        result.taskSpeedLimitTextWidthPx  = DipsToPixels(std::max(0.0f, chrome.textRect.right - chrome.textRect.left), _dpi);
                        result.taskSpeedLimitLabelFits    = labelWidthDip > 0.0f && result.taskSpeedLimitLabelWidthPx <= result.taskSpeedLimitTextWidthPx;
                    }
                }
                break;
            case PopupHitTest::Kind::TaskConflictInspectSource:
            case PopupHitTest::Kind::TaskConflictInspectDestination:
                if (result.found && button.hit.taskId == result.taskId && i < _hostedButtons.size() && _hostedButtons[i])
                {
                    ++result.taskInspectionControlCount;
                    const float labelWidth = MeasureTextWidth(
                        _dwriteFactory.get(), _controlsHost.GetTextFormat(DxUi::FontRole::Body), _hostedButtons[i]->GetText(), 4096.0f, 128.0f);
                    result.taskInspectionLabelsFit =
                        result.taskInspectionLabelsFit && DipsToPixels(labelWidth + 16.0f, _dpi) <= button.bounds.right - button.bounds.left;
                }
                break;
            case PopupHitTest::Kind::None:
            default: break;
        }

        for (size_t j = i + 1u; j < _buttons.size(); ++j)
        {
            if (rectsOverlap(visibleButtonBounds(button), visibleButtonBounds(_buttons[j])))
            {
                result.hasVisibleButtonOverlap = true;
                if (result.found && button.hit.taskId == result.taskId && _buttons[j].hit.taskId == result.taskId)
                {
                    result.taskHasVisibleButtonOverlap = true;
                }
            }
        }
    }

    result.completedActionRowFits = result.completedActionRowFits && ! result.taskHasVisibleButtonOverlap;

    for (const auto& descriptor : _hostedTooltipDescriptors)
    {
        if (descriptor.taskId != result.taskId)
            continue;
        if (descriptor.automationId.ends_with(L".Progress.Amount"))
            result.taskAmountRect = descriptor.bounds;
        else if (descriptor.automationId.ends_with(L".Progress.Percent"))
            result.taskPercentRect = descriptor.bounds;
        else if (descriptor.automationId.ends_with(L".Throughput.Rate"))
            result.taskRateRect = descriptor.bounds;
        else if (descriptor.automationId.ends_with(L".Throughput.Eta"))
            result.taskEtaRect = descriptor.bounds;
        else if (descriptor.automationId.ends_with(L".Progress.Counts"))
        {
            result.taskCountsRect = descriptor.bounds;
            result.taskCountsText = descriptor.text;
        }
        else if (descriptor.automationId.ends_with(L".Throughput.Label"))
            result.taskThroughputLabelRect = descriptor.bounds;
        else if (descriptor.automationId.ends_with(L".Concurrency"))
        {
            result.taskConcurrencyRect = descriptor.bounds;
            result.taskConcurrencyText = descriptor.text;
        }
    }
    *snapshot = result;
    return result.found ? 1 : 0;
}
#endif

#ifdef ENABLE_TESTS
FileOperationsPopupInternal::PopupPointerDebugSnapshot FileOperationsPopupInternal::FileOperationsPopupState::DebugPointerSnapshotForSelfTest(
    POINT clientPoint, PopupHitTest::Kind kind) noexcept
{
    PopupPointerDebugSnapshot result;
    result.mouseMoveCount = _debugMouseMoveCount;
    result.lastMouseMove  = _debugLastMouseMove;
    result.hostDpi        = _controlsHost.GetDpi();
    result.hostHwnd       = _controlsHost.GetHwnd();
    result.nativeCapture  = GetCapture();
    if (const auto* captured = _controlsHost.GetCapturedControl())
        result.capturedAutomationId = captured->GetAccessibleAutomationId();
    if (const auto* hit = _controlsHost.DebugHitTestControl(
            D2D1::Point2F(_controlsHost.PixelsToDip(static_cast<float>(clientPoint.x)), _controlsHost.PixelsToDip(static_cast<float>(clientPoint.y)))))
    {
        result.hitAutomationId = hit->GetAccessibleAutomationId();
        result.hitBoundsDip    = hit->GetBounds();
        result.hitVisible      = hit->IsVisible();
        result.hitEnabled      = hit->IsEnabled();
    }
    for (size_t index = 0u; index < _hostedButtonHits.size() && index < _hostedButtons.size(); ++index)
        if (_hostedButtonHits[index].kind == kind && _hostedButtons[index])
        {
            const auto* target     = _hostedButtons[index];
            result.targetBoundsDip = target->GetBounds();
            result.targetVisible   = target->IsVisible();
            result.targetEnabled   = target->IsEnabled();
            result.targetHovered   = target->IsHovered();
            break;
        }
    return result;
}
#endif

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMouseWheel(HWND hwnd, int delta, POINT screenPoint) noexcept
{
    POINT point = screenPoint;
    ScreenToClient(hwnd, &point);
    if (! _pathDetailId.empty())
    {
        if (_pathDetailField)
            static_cast<void>(_pathDetailField->OnMouseWheel(
                _controlsHost,
                D2D1::Point2F(_controlsHost.PixelsToDip(static_cast<float>(point.x)), _controlsHost.PixelsToDip(static_cast<float>(point.y))),
                static_cast<float>(delta),
                0u));
        return 0;
    }
    for (auto& body : _conflictBodies)
    {
        if (! body.visible || ! body.panel || ! PointInRectF(_listViewportRect, static_cast<float>(point.x), static_cast<float>(point.y)) ||
            ! PointInRectF(body.bounds, static_cast<float>(point.x), static_cast<float>(point.y)))
            continue;
        const float previous = body.offsetDip;
        body.panel->SetScrollOffset(previous - static_cast<float>(delta) / static_cast<float>(WHEEL_DELTA) * 36.0f);
        if (body.offsetDip != previous)
            return 0;
        break; // At the inner boundary, continue normal outer-list scrolling.
    }
    const int step = std::max(1, DipsToPixels(36, _dpi));
    _mouseWheelRemainder += delta;

    const int steps      = _mouseWheelRemainder / WHEEL_DELTA;
    _mouseWheelRemainder = _mouseWheelRemainder % WHEEL_DELTA;

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    if (GetScrollInfo(hwnd, SB_VERT, &si))
    {
        const int page   = std::max(1, static_cast<int>(si.nPage));
        const int maxPos = std::max(0, si.nMax - page + 1);
        _scrollPos       = std::clamp(_scrollPos - steps * step, 0, maxPos);

        SCROLLINFO set{};
        set.cbSize = sizeof(set);
        set.fMask  = SIF_POS;
        set.nPos   = _scrollPos;
        SetScrollInfo(hwnd, SB_VERT, &set, TRUE);
    }

    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnShowWindow(HWND hwnd, bool showing) noexcept
{
    if (! showing)
    {
        _failureCancelAllEnterArmed = false;
        ClosePathDetail(hwnd, false);
        SyncPathDetail(hwnd);
    }
#ifdef ENABLE_TESTS
    if (showing && _destroyOnNextShowForSelfTest)
    {
        _destroyOnNextShowForSelfTest = false;
        DestroyWindow(hwnd);
        return 0;
    }
#endif
    // Q6: reopening re-attempts the host; when it attaches, the hosted popup returns.
    if (showing && _failureSurface)
        static_cast<void>(TryLeaveFailureSurface(hwnd));
    if (showing && _inspectionReturn)
    {
        if (hostLifetime.lock() && fileOps &&
            fileOps->IsPendingConflictInspectionCurrent(
                _inspectionReturn->invokingHit.taskId, _inspectionReturn->invokingHit.promptGeneration, _inspectionReturn->context))
        {
            _scrollPos                      = _inspectionReturn->scrollPos;
            _inspectionReturn->restoreFocus = true;
            Invalidate(hwnd);
        }
        else
            _inspectionReturn.reset();
    }
    RefreshTelemetryTimer(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnClose(HWND hwnd) noexcept
{
    _failureCancelAllEnterArmed = false;
    if (fileOps)
    {
        fileOps->OnPopupHiddenByUser();
    }
    ShowWindow(hwnd, SW_HIDE);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcPaint(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    const bool capturePerf   = Debug::Perf::IsCaptureEnabled();
    const uint64_t startedUs = capturePerf ? PerfNowUs() : 0u;
    const LRESULT result     = DefWindowProcW(hwnd, WM_NCPAINT, wParam, lParam);
    PaintCaptionStatusGlyph(hwnd);
    if (capturePerf)
    {
        Debug::Perf::Emit(L"FileOps.Popup.WmNcPaintUs", L"", PerfElapsedUs(startedUs), 0u, 0u, result >= 0 ? S_OK : E_FAIL);
    }
    return result;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcActivate(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    const bool capturePerf   = Debug::Perf::IsCaptureEnabled();
    const uint64_t startedUs = capturePerf ? PerfNowUs() : 0u;
    if (folderWindow && hostLifetime.lock())
    {
        ApplyTitleBarTheme(hwnd, folderWindow->GetTheme(), wParam != FALSE);
    }

    const LRESULT result = DefWindowProcW(hwnd, WM_NCACTIVATE, wParam, lParam);
    PaintCaptionStatusGlyph(hwnd);
    if (capturePerf)
    {
        Debug::Perf::Emit(L"FileOps.Popup.WmNcActivateUs", L"", PerfElapsedUs(startedUs), wParam != FALSE ? 1u : 0u, 0u, result >= 0 ? S_OK : E_FAIL);
    }
    return result;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
#ifdef ENABLE_TESTS
    if (msg == WM_MOUSEMOVE)
    {
        ++_debugMouseMoveCount;
        _debugLastMouseMove = POINT{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
    }
    if ((msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) && wp == VK_TAB)
        ++_debugNativeTabCount;
    if ((msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN || msg == WM_KEYUP || msg == WM_SYSKEYUP) && (wp == VK_SPACE || wp == VK_RETURN))
    {
        const auto* focused = _controlsHost.GetFocusControl();
        SelfTest::AppendSelfTestTrace(std::format(L"Popup key delivery: message=0x{:X} key={} extra=0x{:X} hwnd=0x{:X} nativeFocus=0x{:X} "
                                                  L"foreground=0x{:X} active=0x{:X} hostedFocus='{}' enabled={} visible={} modifiers=0x{:X}",
                                                  msg,
                                                  wp,
                                                  static_cast<UINT_PTR>(GetMessageExtraInfo()),
                                                  reinterpret_cast<UINT_PTR>(hwnd),
                                                  reinterpret_cast<UINT_PTR>(GetFocus()),
                                                  reinterpret_cast<UINT_PTR>(GetForegroundWindow()),
                                                  reinterpret_cast<UINT_PTR>(GetActiveWindow()),
                                                  focused ? focused->GetAccessibleAutomationId() : std::wstring_view{},
                                                  focused && focused->IsEnabled(),
                                                  focused && focused->IsVisible(),
                                                  _controlsHost.DebugGetModifierState()));
    }
#endif
    const UINT taskbarButtonCreatedMessage = FileOperationsTaskbarButtonCreatedMessage();
    if (taskbarButtonCreatedMessage != 0 && msg == taskbarButtonCreatedMessage)
    {
        _taskbarButtonReady        = true;
        _taskbarListRetryAfterTick = 0;
        _taskbarListAttemptCount   = 0u;
        _taskbarList.reset();
        ApplyTaskbarProgress(hwnd, _cachedTaskbarState, _cachedTaskbarCompleted, _cachedTaskbarTotal);
        Invalidate(hwnd);
        return 0;
    }

    if (_failureSurface && ((msg == WM_KEYDOWN && HandleFailureSurfaceKeyDown(hwnd, GetFocus(), static_cast<UINT>(wp), lp)) ||
                            (msg == WM_KEYUP && HandleFailureSurfaceKeyUp(GetFocus(), static_cast<UINT>(wp)))))
    {
        return 0;
    }

    if ((msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) && HandleHostedKeyDown(hwnd, static_cast<UINT>(wp), lp, msg == WM_SYSKEYDOWN))
    {
        return 0;
    }

    // A window-scoped report error belongs to this popup. Clear its host-owned
    // alert before the HWND can be recycled and before the anchor's NC destroy.
    if (msg == WM_DESTROY)
        static_cast<void>(HostClearAlert(HOST_ALERT_SCOPE_WINDOW, reinterpret_cast<void*>(hwnd)));

    if (msg != WM_CREATE && msg != WM_NCDESTROY && msg != WM_PAINT && msg != WM_SIZE && msg != WM_DPICHANGED && msg != WM_MOUSEWHEEL)
    {
        bool hostedHandled         = false;
        const LRESULT hostedResult = _controlsHost.HandleMessage(hwnd, msg, wp, lp, hostedHandled);
        if (hostedHandled)
        {
            return hostedResult;
        }
    }

    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd);
        case WM_SHOWWINDOW: return OnShowWindow(hwnd, wp != FALSE);
        case WM_NCDESTROY: return OnNcDestroy(hwnd);
        case WM_NCACTIVATE: return OnNcActivate(hwnd, wp, lp);
        case WM_NCPAINT: return OnNcPaint(hwnd, wp, lp);
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: return OnPaint(hwnd);
        case WM_SIZE: return OnSize(hwnd, wp, LOWORD(lp), HIWORD(lp));
        case WM_MOVE: return OnMove(hwnd);
        case WM_GETMINMAXINFO: return OnGetMinMaxInfo(hwnd, reinterpret_cast<MINMAXINFO*>(lp));
        case WM_ENTERSIZEMOVE: return OnEnterSizeMove(hwnd);
        case WM_EXITSIZEMOVE: return OnExitSizeMove(hwnd);
        case WM_TIMER: return OnTimer(hwnd, static_cast<UINT_PTR>(wp));
        case WndMsg::kFileOperationPresentationChanged: RefreshTelemetryTimer(hwnd); return 0;
        case WM_VSCROLL: return OnVScroll(hwnd, static_cast<UINT>(LOWORD(wp)));
        case WM_MOUSEWHEEL: return OnMouseWheel(hwnd, GET_WHEEL_DELTA_WPARAM(wp), POINT{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_COMMAND:
            if (_failureSurface && HIWORD(wp) == BN_CLICKED)
            {
                const HWND control = reinterpret_cast<HWND>(lp);
                if (control == _failureCancelAll)
                {
                    _failureCancelAllEnterArmed = false;
                    static_cast<void>(ConfirmCancelAll(hwnd));
                    return 0;
                }
                if (control == _failureClose)
                {
                    return OnClose(hwnd);
                }
            }
            break;
        case WM_CTLCOLORSTATIC:
            if (_failureSurface)
            {
                HDC dc = reinterpret_cast<HDC>(wp);
                SetBkColor(dc, GetSysColor(COLOR_WINDOW));
                SetTextColor(dc, GetSysColor(COLOR_WINDOWTEXT));
                return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_WINDOW));
            }
            break;
        case WM_SETFOCUS:
            if (_failureSurface && _failureClose)
            {
                SetFocus(_failureClose);
                return 0;
            }
            break;
        case WM_KILLFOCUS: _failureCancelAllEnterArmed = false; break;
        case kFileOperationsPopupDeferredSpeedLimitPromptMessage:
        {
            return ShowCustomSpeedLimitPromptForTask(hwnd, static_cast<uint64_t>(lp)) ? 1 : 0;
        }
        case WM_DPICHANGED:
        {
            const auto* suggested = reinterpret_cast<const RECT*>(lp);
            return suggested ? OnDpiChanged(hwnd, LOWORD(wp), *suggested) : 0;
        }
        case WM_THEMECHANGED:
        case WM_SYSCOLORCHANGE:
        case WM_SETTINGCHANGE: return OnThemeChanged(hwnd);
        case WM_CLOSE: return OnClose(hwnd);
#ifdef ENABLE_TESTS
        case WndMsg::kFileOpsPopupSelfTestInvoke: return OnSelfTestInvoke(hwnd, reinterpret_cast<const PopupSelfTestInvoke*>(lp));
        case WndMsg::kFileOpsPopupSelfTestSnapshot: return OnTaskSnapshotRequest(reinterpret_cast<const PopupTaskSnapshotRequest*>(lp));
        case WndMsg::kFileOpsPopupCaptionGlyphSnapshot: return OnCaptionGlyphSnapshotRequest(reinterpret_cast<CaptionGlyphDebugSnapshot*>(lp));
        case WndMsg::kFileOpsPopupLayoutSnapshot: return OnLayoutSnapshotRequest(reinterpret_cast<PopupLayoutDebugSnapshot*>(lp));
        case WndMsg::kFileOpsPopupTaskbarUpdate: UpdateTaskbarProgress(hwnd); return TRUE;
#endif
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

namespace
{
} // namespace

LRESULT CALLBACK FileOperationsPopupInternal::FileOperationsPopupState::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    FileOperationsPopupState* state = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        state    = cs ? reinterpret_cast<FileOperationsPopupState*>(cs->lpCreateParams) : nullptr;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
    }
    else
    {
        state = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (! state)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    ++state->_dispatchDepth;
    const auto finishDispatch = wil::scope_exit([state]() noexcept
    {
        if (state->_dispatchDepth > 0u)
        {
            --state->_dispatchDepth;
        }
        if (state->_dispatchDepth == 0u && state->_deletePending)
        {
            delete state;
        }
    });

    return state->WndProc(hwnd, msg, wp, lp);
}

void FileOperationsPopup::RefreshTelemetry(HWND popup) noexcept
{
    if (popup && IsWindow(popup) && GetWindowThreadProcessId(popup, nullptr) == GetCurrentThreadId())
    {
        // Normal HWND dispatch retains the popup state across reentrant COM or
        // owner teardown, unlike directly borrowing its GWLP_USERDATA pointer.
        SendMessageW(popup, WndMsg::kFileOperationPresentationChanged, 0, 0);
    }
}

HWND FileOperationsPopup::Create(FolderWindow::FileOperationState* fileOps,
                                 FolderWindow* folderWindow,
                                 HWND ownerWindow,
                                 std::weak_ptr<void> hostLifetime) noexcept
{
    if (! fileOps || ! folderWindow)
    {
        return nullptr;
    }

    if (hostLifetime.expired())
    {
        return nullptr;
    }

    if (! RegisterFileOperationsPopupWndClass(GetModuleHandleW(nullptr)))
    {
        return nullptr;
    }

    auto statePtr          = std::make_unique<FileOperationsPopupInternal::FileOperationsPopupState>();
    statePtr->fileOps      = fileOps;
    statePtr->folderWindow = folderWindow;
    statePtr->hostLifetime = std::move(hostLifetime);

    const UINT ownerDpi           = ownerWindow ? GetDpiForWindow(ownerWindow) : USER_DEFAULT_SCREEN_DPI;
    const int desiredClientWidth  = DipsToPixels(480, ownerDpi);
    const int desiredClientHeight = DipsToPixels(460, ownerDpi);

    const DWORD style   = WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_VSCROLL;
    const DWORD exStyle = WS_EX_APPWINDOW;

    int width  = 0;
    int height = 0;
    int x      = 0;
    int y      = 0;

    bool useSavedPlacement = false;
    RECT savedRect{};
    [[maybe_unused]] bool startMaximized = false;
    RECT expandedRect{};
    const bool hasExpandedPlacement = fileOps && fileOps->TryGetPopupExpandedPlacement(expandedRect, ownerDpi);
    if (hasExpandedPlacement)
    {
        statePtr->SetExpandedPlacementForRestore(expandedRect);
    }
    if (fileOps)
    {
        if (! fileOps->GetPopupFooterOnly() && hasExpandedPlacement)
        {
            savedRect         = expandedRect;
            useSavedPlacement = true;
        }
        else if (fileOps->TryGetPopupPlacement(savedRect, startMaximized, ownerDpi))
        {
            useSavedPlacement = true;
        }
        else
        {
            const std::optional<RECT> lastRectOpt = fileOps->GetLastPopupRect();
            if (lastRectOpt.has_value() && IsRectFullyVisible(lastRectOpt.value()))
            {
                savedRect         = lastRectOpt.value();
                useSavedPlacement = true;
            }
        }
    }

    if (useSavedPlacement)
    {
        width  = std::max(0L, savedRect.right - savedRect.left);
        height = std::max(0L, savedRect.bottom - savedRect.top);
        x      = static_cast<int>(savedRect.left);
        y      = static_cast<int>(savedRect.top);
        if (fileOps && ! fileOps->GetPopupFooterOnly())
        {
            statePtr->SetExpandedPlacementForRestore(savedRect, true);
        }
    }
    else
    {
        const HWND monitorOwner = ownerWindow ? ownerWindow : folderWindow->GetHwnd();
        HMONITOR monitor        = MonitorFromWindow(monitorOwner, MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (! GetMonitorInfoW(monitor, &mi))
        {
            return nullptr;
        }

        const RECT work = mi.rcWork;

        RECT desiredWindowRect{0, 0, desiredClientWidth, desiredClientHeight};
        AdjustWindowRectExForDpi(&desiredWindowRect, style, FALSE, exStyle, ownerDpi);
        width  = std::max(0L, desiredWindowRect.right - desiredWindowRect.left);
        height = std::max(0L, desiredWindowRect.bottom - desiredWindowRect.top);

        RECT ownerRect{};
        bool useOwnerCenter = false;
        if (ownerWindow && ! IsIconic(ownerWindow) && GetWindowRect(ownerWindow, &ownerRect))
        {
            useOwnerCenter = true;
        }

        int centerX = work.left + (work.right - work.left - width) / 2;
        int centerY = work.top + (work.bottom - work.top - height) / 2;

        if (useOwnerCenter)
        {
            const int ownerW = std::max(0L, ownerRect.right - ownerRect.left);
            const int ownerH = std::max(0L, ownerRect.bottom - ownerRect.top);
            centerX          = ownerRect.left + (ownerW - width) / 2;
            centerY          = ownerRect.top + (ownerH - height) / 2;
        }

        const int maxX = work.right - width;
        if (maxX >= work.left)
        {
            x = std::clamp(centerX, static_cast<int>(work.left), maxX);
        }
        else
        {
            x = work.left;
        }

        const int maxY = work.bottom - height;
        if (maxY >= work.top)
        {
            y = std::clamp(centerY, static_cast<int>(work.top), maxY);
        }
        else
        {
            y = work.top;
        }
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_FILEOPS_POPUP_TITLE);

    // Transfer ownership to window - it will delete itself in WM_DESTROY
    auto* state = statePtr.release();
    HWND popup =
        CreateWindowExW(exStyle, kFileOperationsPopupClassName, title.c_str(), style, x, y, width, height, nullptr, nullptr, GetModuleHandleW(nullptr), state);

    if (! popup)
    {
        // Reclaim ownership via unique_ptr destructor
        std::unique_ptr<FileOperationsPopupInternal::FileOperationsPopupState> reclaimed(state);
        return nullptr;
    }

    return popup;
}

#ifdef ENABLE_TESTS
HWND FindFileOperationsDebugWindowForCurrentProcess(const wchar_t* className) noexcept
{
    if (! className || *className == L'\0')
    {
        return nullptr;
    }

    struct SearchState
    {
        DWORD processId          = 0;
        const wchar_t* className = nullptr;
        HWND hwnd                = nullptr;
    } state{GetCurrentProcessId(), className, nullptr};

    EnumWindows(
        [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* state = reinterpret_cast<SearchState*>(lParam);
        if (! state || ! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
        {
            return TRUE;
        }

        DWORD processId = 0;
        GetWindowThreadProcessId(hwnd, &processId);
        if (processId != state->processId)
        {
            return TRUE;
        }

        wchar_t windowClass[128]{};
        if (GetClassNameW(hwnd, windowClass, static_cast<int>(std::size(windowClass))) == 0)
        {
            return TRUE;
        }

        if (wcscmp(windowClass, state->className) != 0)
        {
            return TRUE;
        }

        state->hwnd = hwnd;
        return FALSE;
    },
        reinterpret_cast<LPARAM>(&state));

    return state.hwnd;
}

bool DebugInvokeFileOperationsPopup(HWND popup, const FileOperationsPopupInternal::PopupSelfTestInvoke& invoke) noexcept
{
    // Keep the selftest invoke synchronous: the stack payload must remain live while
    // the popup opens any modal test surface.
    return popup && IsWindow(popup) != FALSE && SendMessageW(popup, WndMsg::kFileOpsPopupSelfTestInvoke, 0, reinterpret_cast<LPARAM>(&invoke)) != 0;
}

bool DebugGetFileOperationsPopupTaskSnapshot(HWND popup, uint64_t taskId, FileOperationsPopupInternal::TaskSnapshot& out) noexcept
{
    if (! popup || IsWindow(popup) == FALSE)
    {
        return false;
    }

    FileOperationsPopupInternal::PopupTaskSnapshotRequest request{};
    request.taskId = taskId;
    const bool ok  = SendMessageW(popup, WndMsg::kFileOpsPopupSelfTestSnapshot, 0, reinterpret_cast<LPARAM>(&request)) != FALSE;
    if (ok && request.found)
    {
        out = std::move(request.snapshot);
    }
    return ok && request.found;
}

bool DebugGetFileOperationsPopupCaptionGlyphSnapshot(HWND popup, FileOperationsPopupInternal::CaptionGlyphDebugSnapshot& out) noexcept
{
    if (! popup || IsWindow(popup) == FALSE)
    {
        return false;
    }

    FileOperationsPopupInternal::CaptionGlyphDebugSnapshot snapshot{};
    const bool ok = SendMessageW(popup, WndMsg::kFileOpsPopupCaptionGlyphSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
    if (ok)
    {
        out = snapshot;
    }
    return ok;
}

bool DebugGetFileOperationsPopupLayoutSnapshot(HWND popup, FileOperationsPopupInternal::PopupLayoutDebugSnapshot& out) noexcept
{
    if (! popup || IsWindow(popup) == FALSE)
    {
        return false;
    }

    FileOperationsPopupInternal::PopupLayoutDebugSnapshot snapshot{};
    snapshot.taskId = out.taskId;
    const bool ok   = SendMessageW(popup, WndMsg::kFileOpsPopupLayoutSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
    if (ok && snapshot.found)
    {
        out = snapshot;
    }
    return ok && snapshot.found;
}

bool DebugUpdateFileOperationsPopupTaskbarProgress(HWND popup) noexcept
{
    return popup && IsWindow(popup) != FALSE && SendMessageW(popup, WndMsg::kFileOpsPopupTaskbarUpdate, 0, 0) != FALSE;
}

bool DebugBuildFileOperationsPopupGlobalSummarySnapshot(const std::vector<FileOperationsPopupInternal::TaskSnapshot>& tasks,
                                                        FileOperationsPopupInternal::PopupLayoutDebugSnapshot& out,
                                                        double displayedBytesPerSecOverride,
                                                        double aggregateEtaSecondsOverride) noexcept
{
    GlobalFileOperationsStatusSummary summary = BuildGlobalStatusSummary(tasks);
    if (displayedBytesPerSecOverride >= 0.0)
    {
        summary.displayedBytesPerSec = displayedBytesPerSecOverride;
    }
    if (aggregateEtaSecondsOverride >= 0.0)
    {
        summary.hasAggregateEta     = true;
        summary.aggregateEtaSeconds = aggregateEtaSecondsOverride;
    }

    const GlobalTaskbarProgressModel taskbarProgress = BuildGlobalTaskbarProgressModel(summary);

    out.globalRunningCount                 = summary.running;
    out.globalWaitingCount                 = summary.waiting;
    out.globalNeedAttentionCount           = summary.needAttention;
    out.globalSummaryText                  = FormatGlobalStatusSummaryText(summary);
    out.globalSummaryVisible               = ! tasks.empty();
    out.footerAggregateProgressVisible     = HasGlobalAggregateProgress(summary);
    out.footerAggregateProgressDeterminate = HasDeterminateGlobalAggregateProgress(summary);
    out.footerAggregateProgressFraction    = GlobalAggregateProgressFraction(summary);
    out.footerOpenDiscoveryTasks           = summary.openDiscoveryTasks;
    out.footerAggregateCompletedBytes      = summary.completedBytes;
    out.footerAggregateTotalBytes          = summary.totalBytes;
    out.footerAggregateCompletedItems      = summary.completedItems;
    out.footerAggregateTotalItems          = summary.totalItems;
    out.footerAggregateBytesPerSecond      = summary.displayedBytesPerSec;
    out.footerAggregateEtaVisible          = summary.hasAggregateEta;
    out.footerAggregateEtaSeconds          = summary.hasAggregateEta ? SaturatingCeilNonNegativeToUint64(summary.aggregateEtaSeconds) : 0ull;
    out.footerPauseResumeAllVisible        = HasFooterPauseResumeAllControl(summary);
    out.footerPauseResumeAllPauses         = FooterPauseResumeAllShouldPause(summary);
    out.taskbarProgressState               = taskbarProgress.state;
    out.taskbarProgressCompleted           = taskbarProgress.completed;
    out.taskbarProgressTotal               = taskbarProgress.total;
    return true;
}

void DebugFailNextFileOperationsTaskbarListAttempts(unsigned int attempts) noexcept
{
    g_fileOperationsTaskbarListForcedFailures.store(attempts, std::memory_order_release);
}

void DebugFailNextFileOperationsDxUiHostAttachAttempts(unsigned int attempts) noexcept
{
    g_fileOperationsDxUiHostAttachForcedFailures.store(attempts, std::memory_order_release);
}

void DebugFailNextFileOperationsD2DTargetAttempts(unsigned int attempts) noexcept
{
    g_fileOperationsD2DTargetForcedFailures.store(attempts, std::memory_order_release);
}

void DebugFailNextFileOperationsPopupLayoutMeasurements(unsigned int attempts) noexcept
{
    g_fileOperationsLayoutMeasurementForcedFailures.store(attempts, std::memory_order_release);
}

void DebugFailNextFileOperationsFailureSurfaceChildCreation(unsigned int ordinal) noexcept
{
    g_fileOperationsFailureSurfaceChildCreationOrdinal.store(ordinal <= 3u ? ordinal : 0u, std::memory_order_release);
}

void DebugFailNextFileOperationsFailureSurfaceButtonSetup(unsigned int ordinal) noexcept
{
    g_fileOperationsFailureSurfaceButtonSetupOrdinal.store(ordinal <= 2u ? ordinal : 0u, std::memory_order_release);
}

bool DebugEnterFileOperationsPopupFailureSurface(HWND popup) noexcept
{
    if (! popup || IsWindow(popup) == FALSE || GetWindowThreadProcessId(popup, nullptr) != GetCurrentThreadId())
        return false;
    auto* state = reinterpret_cast<FileOperationsPopupInternal::FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA));
    return state && state->DebugEnterFailureSurface(popup);
}

bool DebugSetFileOperationsPopupIdleSuspensionForF6(HWND popup, FolderWindow::FileOperationState* expectedOwner, bool enabled) noexcept
{
    if (! expectedOwner || ! popup || IsWindow(popup) == FALSE || GetWindowThreadProcessId(popup, nullptr) != GetCurrentThreadId())
        return false;
    auto* state = reinterpret_cast<FileOperationsPopupInternal::FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA));
    if (! state || state->fileOps != expectedOwner)
        return false;
    state->DebugSetIdleSuspensionEnabled(popup, enabled);
    return true;
}

bool DebugGetFileOperationsPopupUiMotionForF6(HWND popup, FolderWindow::FileOperationState* expectedOwner, bool& visibleMotion) noexcept
{
    visibleMotion = false;
    if (! expectedOwner || ! popup || IsWindow(popup) == FALSE || GetWindowThreadProcessId(popup, nullptr) != GetCurrentThreadId())
        return false;
    const auto* state = reinterpret_cast<const FileOperationsPopupInternal::FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA));
    if (! state || state->fileOps != expectedOwner)
        return false;
    visibleMotion = state->DebugHasVisibleUiMotion(popup);
    return true;
}

bool DebugGetFileOperationsPopupTelemetrySnapshot(HWND popup, FileOperationsPopupInternal::PopupTelemetryDebugSnapshot& out) noexcept
{
    out = {};
    if (! popup || IsWindow(popup) == FALSE || GetWindowThreadProcessId(popup, nullptr) != GetCurrentThreadId())
        return false;
    const auto* state = reinterpret_cast<const FileOperationsPopupInternal::FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA));
    if (! state)
        return false;
    out = state->DebugTelemetrySnapshot();
    return true;
}

DxUi::ButtonVariant DebugResolveFileOperationsHostedButtonVariant(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
{
    return HostedButtonVariant(kind);
}

std::wstring DebugFormatFileOperationsSpeedLimitSelectorText(uint64_t bytesPerSecond)
{
    return FormatSpeedLimitSelectorText(bytesPerSecond);
}

bool DebugBuildFileOperationsGraphFairColorWeightSnapshot(FileOperationsPopupInternal::GraphHueWeightDebugSnapshot& out) noexcept
{
    out = {};

    RateHistory history{};
    AddPendingHueWeight(history, 0u, 10.0f, 100.0);
    AddPendingHueWeight(history, 1u, 100.0f, 100.0);
    AddPendingHueWeight(history, 2u, 190.0f, 100.0);
    AddPendingHueWeight(history, 3u, 280.0f, 100.0);
    AppendResampledRateSamples(history, kRateSampleBucketMs, 400.0, 0.0, -1.0f);

    if (history.count == 0u)
    {
        return false;
    }

    const size_t sampleIndex = (history.writeIndex + RateHistory::kMaxSamples - 1u) % RateHistory::kMaxSamples;
    out.hueCount             = std::min<size_t>(history.hueWeightCounts[sampleIndex], out.hues.size());
    for (size_t i = 0; i < out.hueCount; ++i)
    {
        out.hues[i]    = history.hueWeights[sampleIndex][i].hue;
        out.weights[i] = history.hueWeights[sampleIndex][i].weight;
        out.totalWeight += out.weights[i];
    }
    return true;
}

bool DebugBuildFileOperationsGraphFairnessHistorySnapshot(FileOperationsPopupInternal::PopupLayoutDebugSnapshot& out) noexcept
{
    out = {};

    RateHistory history{};
    constexpr uint64_t kHistoryBucketCount = static_cast<uint64_t>(RateHistory::kMaxSamples) + 7u;
    for (uint64_t bucket = 1; bucket <= kHistoryBucketCount; ++bucket)
    {
        RateSnapshot task{};
        task.taskId                   = 1;
        task.started                  = true;
        task.inFlightFileCount        = 4u;
        task.completedBytes           = bucket * 400u;
        task.lastProgressCallbackTick = bucket * kRateSampleBucketMs;
        task.progressStateChangeTick  = task.lastProgressCallbackTick;
        for (size_t i = 0; i < task.inFlightFileCount; ++i)
        {
            auto& stream            = task.inFlightFiles[i];
            stream.cookieKey        = reinterpret_cast<const void*>(static_cast<uintptr_t>(i + 1u));
            stream.progressStreamId = static_cast<uint64_t>(i + 1u);
            stream.sourcePath       = std::format(L"synthetic-stream-{}.bin", i + 1u);
            stream.totalBytes       = (kHistoryBucketCount + 1u) * 100u;
            stream.completedBytes   = bucket * 100u;
            stream.lastUpdateTick   = task.lastProgressCallbackTick;
        }

        AccumulateStreamHueWeights(history, task, 400u, -1.0f);
        AppendResampledRateSamples(history, kRateSampleBucketMs, static_cast<double>(bucket), 0.0, -1.0f);
        AppendRateWindowSample(history, task, task.lastProgressCallbackTick);
    }

    out.taskId                   = 1;
    out.found                    = true;
    const size_t oldestIndex     = (history.writeIndex + RateHistory::kMaxSamples - history.count) % RateHistory::kMaxSamples;
    const size_t newestIndex     = (history.writeIndex + RateHistory::kMaxSamples - 1u) % RateHistory::kMaxSamples;
    out.graphHistoryCount        = history.count;
    out.graphHistoryWriteIndex   = history.writeIndex;
    out.graphHistoryOldestSample = history.samples[oldestIndex];
    out.graphHistoryNewestSample = history.samples[newestIndex];
    PopulateGraphHueDebugSummary(history, out);
    // The marker shows the rate the production window measured over the fixture's last 5 s (the task
    // delivered 400 bytes per 100 ms bucket: 4,000 B/s), never a plotted sample.
    history.displayedBytesPerSec            = ComputeWindowedRates(history, kHistoryBucketCount * kRateSampleBucketMs).bytesPerSec;
    out.graphCurrentBandwidthBytesPerSecond = CurrentBandwidthForGraphMarker(history);
    out.graphCurrentBandwidthLineVisible    = out.graphCurrentBandwidthBytesPerSecond > 0.0;

    AppTheme darkTheme{};
    darkTheme.dark = true;
    for (size_t i = 0; i < history.streamProgressCount && i < history.streamProgress.size(); ++i)
    {
        const auto& graphStream = history.streamProgress[i];
        TaskSnapshot::InFlightFileSnapshot rowStream{};
        rowStream.cookieKey        = graphStream.cookieKey;
        rowStream.progressStreamId = graphStream.progressStreamId;
        rowStream.sourcePath       = graphStream.sourcePath;

        const float rowHue            = AssignedRateStreamHue(&history, &rowStream);
        const D2D1_COLOR_F rowColor   = RainbowProgressColor(darkTheme, rowStream.sourcePath, &history, &rowStream);
        const D2D1_COLOR_F graphColor = DxUi::ThroughputGraphColorFromHue(graphStream.assignedHue, darkTheme.dark);
        const bool exactColorMatch    = rowHue == graphStream.assignedHue && rowColor.r == graphColor.r && rowColor.g == graphColor.g &&
                                        rowColor.b == graphColor.b && rowColor.a == graphColor.a;
        if (exactColorMatch)
        {
            ++out.graphRowColorMatchCount;
        }
        else
        {
            ++out.graphRowColorMismatchCount;
        }
    }

    return out.graphMultiHueBucketCount >= 10u && out.graphDistinctHueCount == 4u && out.graphMinHueShare >= 0.20 && out.graphMaxHueShare <= 0.30 &&
           out.graphRowColorMatchCount == 4u && out.graphRowColorMismatchCount == 0u;
}

bool DebugFileOperationsStreamColorSlotsRemainUniqueUnderPermutation() noexcept
{
    RateHistory history{};
    const auto makeSnapshot = [](std::initializer_list<size_t> order, uint64_t bytes) noexcept
    {
        RateSnapshot task{};
        task.taskId            = 1;
        task.started           = true;
        task.inFlightFileCount = order.size();
        size_t write           = 0u;
        for (const size_t id : order)
        {
            auto& stream            = task.inFlightFiles[write++];
            stream.cookieKey        = reinterpret_cast<const void*>(static_cast<uintptr_t>(id));
            stream.progressStreamId = static_cast<uint64_t>(id);
            stream.sourcePath       = std::format(L"permuted-stream-{}.bin", id);
            stream.totalBytes       = 4096u;
            stream.completedBytes   = bytes;
            stream.lastUpdateTick   = bytes;
        }
        return task;
    };

    AccumulateStreamHueWeights(history, makeSnapshot({1u, 2u, 3u}, 100u), 100u, -1.0f);
    std::array<uint8_t, 3> firstSlots{};
    for (size_t i = 0; i < 3u; ++i)
    {
        firstSlots[i] = history.streamProgress[i].assignedColorSlot;
        if (firstSlots[i] >= RateHistory::kMaxHueWeightsPerSample)
        {
            return false;
        }
    }
    if (firstSlots[0] == firstSlots[1] || firstSlots[0] == firstSlots[2] || firstSlots[1] == firstSlots[2])
    {
        return false;
    }

    AccumulateStreamHueWeights(history, makeSnapshot({3u, 1u, 2u}, 200u), 100u, -1.0f);
    std::array<bool, RateHistory::kMaxHueWeightsPerSample> occupied{};
    uint8_t slotFor1 = DxUi::ThroughputGraphSample::kInvalidColorSlot;
    uint8_t slotFor2 = DxUi::ThroughputGraphSample::kInvalidColorSlot;
    uint8_t slotFor3 = DxUi::ThroughputGraphSample::kInvalidColorSlot;
    for (size_t i = 0; i < history.streamProgressCount; ++i)
    {
        const auto& stream = history.streamProgress[i];
        if (stream.assignedColorSlot >= occupied.size() || occupied[stream.assignedColorSlot])
        {
            return false;
        }
        occupied[stream.assignedColorSlot] = true;
        if (stream.progressStreamId == 1u)
        {
            slotFor1 = stream.assignedColorSlot;
        }
        else if (stream.progressStreamId == 2u)
        {
            slotFor2 = stream.assignedColorSlot;
        }
        else if (stream.progressStreamId == 3u)
        {
            slotFor3 = stream.assignedColorSlot;
        }
    }
    if (slotFor1 != firstSlots[0] || slotFor2 != firstSlots[1] || slotFor3 != firstSlots[2])
    {
        return false;
    }

    AccumulateStreamHueWeights(history, makeSnapshot({4u, 2u, 1u, 3u}, 300u), 100u, -1.0f);
    occupied = {};
    for (size_t i = 0; i < history.streamProgressCount; ++i)
    {
        const auto& stream = history.streamProgress[i];
        if (stream.assignedColorSlot >= occupied.size() || occupied[stream.assignedColorSlot])
        {
            return false;
        }
        occupied[stream.assignedColorSlot] = true;
        if (stream.progressStreamId == 1u && stream.assignedColorSlot != slotFor1)
        {
            return false;
        }
        if (stream.progressStreamId == 2u && stream.assignedColorSlot != slotFor2)
        {
            return false;
        }
        if (stream.progressStreamId == 3u && stream.assignedColorSlot != slotFor3)
        {
            return false;
        }
    }
    return history.streamProgressCount == 4u;
}

bool DebugValidateFileOperationsStoppingPresentation() noexcept
{
    using FileOperationsPopupInternal::RateHistory;
    using Snapshot = FileOperationsPopupInternal::TaskSnapshot;
    using Status   = Snapshot::StatusKind;
    Snapshot task{};
    task.started            = true;
    task.operationStartTick = 1u;
    task.discoveryClosed    = true;
    // This fixture models a known 100-byte Copy; the default is deliberately size-unknown.
    task.discoveredBytesKnown = true;
    task.operation            = FILESYSTEM_COPY;
    task.totalBytes           = 100u;
    task.completedBytes       = 25u;
    task.totalItems           = 4u;
    task.completedItems       = 1u;
    task.lifecyclePhase       = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Stopping);
    // The normal finalizer also enters Stopping; phase is not cancellation authority.
    if (ResolveTaskStatusKind(task) != Status::Running)
        return false;

    const std::array<Status, 8> liveStatuses{
        {Status::Running, Status::Conflict, Status::Waiting, Status::Paused, Status::Verifying, Status::Discovering, Status::Preparing, Status::Running}};
    for (size_t index = 0u; index < liveStatuses.size(); ++index)
    {
        auto candidate                 = task;
        candidate.conflict.active      = index == 1u;
        candidate.queuePaused          = index == 2u;
        candidate.paused               = index == 3u;
        candidate.verificationActive   = index == 4u;
        candidate.discoveryAheadActive = index == 5u;
        candidate.operationStartTick   = index == 5u ? 0u : 1u;
        candidate.lifecyclePhase =
            static_cast<uint8_t>(index == 6u ? FileOperations::TaskLifecyclePhase::Preparing : FileOperations::TaskLifecyclePhase::Stopping);
        candidate.statusKind = liveStatuses[index];
        if (ResolveTaskStatusKind(candidate) != liveStatuses[index])
            return false;
        candidate.cancellationRequested = true;
        candidate.canMoveQueueUp        = true;
        candidate.canMoveQueueDown      = true;
        candidate.statusKind            = Status::Running; // Stale presentation cannot defeat durable authority.
        if (BuildGlobalStatusSummary({candidate}).stopping != 1u)
            return false;
        ProjectTaskCancellationForDisplay(candidate);
        bool animate       = true;
        const auto summary = BuildGlobalStatusSummary({candidate});
        if (candidate.statusKind != Status::Stopping || candidate.conflict.active || candidate.canMoveQueueUp || candidate.canMoveQueueDown ||
            StatusVisualToneForTaskStatus(candidate.statusKind) != FileOperationsPopupInternal::PopupStatusVisualTone::Neutral ||
            StatusChipTextForTask(candidate, candidate.statusKind, 0u) != LoadStringResource(nullptr, IDS_FILEOPS_STATUS_STOPPING) ||
            StatusTextForTask(candidate, candidate.statusKind, 0u) != LoadStringResource(nullptr, IDS_FILEOPS_STATUS_STOPPING) ||
            ! GraphOverlayTextForStatus(candidate, candidate.statusKind, animate).empty() || animate ||
            BuildTaskHeaderText(candidate, L"Copy", 0u).find(LoadStringResource(nullptr, IDS_FILEOPS_STATUS_STOPPING)) != std::wstring::npos ||
            summary.running != 1u || summary.stopping != 1u || summary.activeRunning != 1u || summary.needAttention != 0u || summary.waiting != 0u ||
            summary.pauseEligibleRunning != 0u || summary.pauseEligiblePaused != 0u || HasFooterPauseResumeAllControl(summary) ||
            FormatGlobalStatusSummaryText(summary) != FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_STOPPING_COUNT, 1ul) ||
            candidate.completedBytes != task.completedBytes || candidate.totalBytes != task.totalBytes)
            return false;
    }

    task.cancellationRequested = true;
    task.paused                = true;
    task.conflict.active       = true;
    ProjectTaskCancellationForDisplay(task);
    const auto progress = ResolveWholeTaskProgressPresentation(task);
    RateHistory history{};
    history.displayedBytesPerSec = 10.0;
    std::unordered_map<uint64_t, RateHistory> rates;
    rates.emplace(task.taskId, std::move(history));
    const auto summary               = BuildGlobalStatusSummary({task}, &rates);
    auto uncanceled                  = task;
    uncanceled.cancellationRequested = false;
    uncanceled.paused                = false;
    uncanceled.statusKind            = Status::Running;
    if (! BuildGlobalStatusSummary({uncanceled}, &rates).hasAggregateEta)
        return false;
    const auto mixed = BuildGlobalStatusSummary({task, uncanceled});
    if (mixed.running != 2u || mixed.stopping != 1u || mixed.pauseEligibleRunning != 1u || mixed.pauseEligiblePaused != 0u ||
        ! HasFooterPauseResumeAllControl(mixed))
        return false;
    const auto taskbar = BuildGlobalTaskbarProgressModel(summary);
    if (! progress.includeInAggregateCohort || ! progress.showThroughputGraph || ! progress.determinate || progress.fraction != 0.25 ||
        summary.displayedBytesPerSec != 10.0 || summary.hasAggregateEta || summary.completedBytes != 25u || summary.totalBytes != 100u ||
        taskbar.state != static_cast<uint32_t>(TBPF_NORMAL) || taskbar.completed != 25u || taskbar.total != 100u)
        return false;
    auto unknownSize                 = task;
    unknownSize.discoveredBytesKnown = false;
    const auto unknownProgress       = ResolveWholeTaskProgressPresentation(unknownSize);
    const auto unknownSummary        = BuildGlobalStatusSummary({unknownSize}, &rates);
    if (! unknownProgress.includeInAggregateCohort || unknownProgress.determinate || unknownProgress.showTransferMarquee || unknownSummary.totalBytes != 0u ||
        unknownSummary.hasAggregateEta)
        return false;
    task.discoveryClosed      = false;
    task.discoveryAheadActive = true;
    const auto openProgress   = ResolveWholeTaskProgressPresentation(task);
    if (openProgress.showDiscoveryActivity || openProgress.showTransferMarquee || ! openProgress.showThroughputGraph)
        return false;

    const std::array<std::pair<HRESULT, Status>, 6> terminal{{{S_OK, Status::Done},
                                                              {S_FALSE, Status::Done},
                                                              {HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), Status::Partial},
                                                              {HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), Status::Indeterminate},
                                                              {E_FAIL, Status::Failed},
                                                              {HRESULT_FROM_WIN32(ERROR_CANCELLED), Status::Canceled}}};
    task.finished = true;
    for (const auto& [result, expected] : terminal)
    {
        task.resultHr   = result;
        task.statusKind = Status::Running;
        if (BuildGlobalStatusSummary({task}).running != 0u)
            return false;
        ProjectTaskCancellationForDisplay(task);
        if (task.statusKind != expected || BuildGlobalStatusSummary({task}).running != 0u)
            return false;
    }
    task.resultHr     = S_OK;
    task.warningCount = 1u;
    if (ResolveTaskStatusKind(task) != Status::Partial)
        return false;
    task.hasIndeterminateResult = true;
    if (ResolveTaskStatusKind(task) != Status::Indeterminate)
        return false;

    // A cancel that left an item's publication or source unknown is not a clean Canceled: its
    // chip, tone and completed-group badge say Indeterminate. A clean cancel stays Canceled.
    for (const HRESULT canceledHr : {HRESULT_FROM_WIN32(ERROR_CANCELLED), E_ABORT})
    {
        task.resultHr               = canceledHr;
        task.warningCount           = 1u;
        task.hasIndeterminateResult = true;
        ProjectTaskCancellationForDisplay(task);
        const CompletedGroupResultSummary unknown = BuildCompletedGroupResultSummary({task});
        if (task.statusKind != Status::Indeterminate || StatusVisualToneForTaskStatus(task.statusKind) != PopupStatusVisualTone::Warning ||
            unknown.indeterminate != 1u || unknown.canceled != 0u)
            return false;
        task.warningCount           = 0u;
        task.hasIndeterminateResult = false;
        ProjectTaskCancellationForDisplay(task);
        const CompletedGroupResultSummary clean = BuildCompletedGroupResultSummary({task});
        if (task.statusKind != Status::Canceled || clean.canceled != 1u || clean.indeterminate != 0u)
            return false;
    }

    // Move sources a cancel or stop kept before publishing were never copied: no downgrade badge.
    Snapshot stopped{};
    stopped.finished                    = true;
    stopped.operation                   = FILESYSTEM_MOVE;
    stopped.retainedSourceCount         = 6u;
    stopped.canceledRetainedSourceCount = 6u;
    std::array<TerminalAxisBadge, 3u> badges{};
    if (BuildTerminalAxisBadges(stopped, badges) != 0u || TaskHasTerminalAxisBadges(stopped))
        return false;
    stopped.retainedSourceCount = 7u;
    return BuildTerminalAxisBadges(stopped, badges) == 1u && badges[0].text == LoadStringResource(nullptr, IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT) &&
           TaskHasTerminalAxisBadges(stopped);
}

bool DebugValidateFileOperationsVerificationPresentation() noexcept
{
    TaskSnapshot task{};
    task.taskId                     = 1u;
    task.operation                  = FILESYSTEM_COPY;
    task.started                    = true;
    task.discoveryClosed            = true;
    task.discoveredBytesKnown       = true;
    task.totalBytes                 = 100u;
    task.completedBytes             = 100u;
    task.verificationRequested      = true;
    task.verificationActive         = true;
    task.verificationTotalBytes     = 100u;
    task.verificationCompletedBytes = 50u;
    if (ResolveTaskStatusKind(task) != TaskStatusKind::Verifying || std::abs(ComputeFileOperationsTaskCompleteFractionForDisplay(task) - 0.75f) > 0.001f ||
        std::abs(VerificationTransferProgressPercent(task) - 100.0) > 0.001 || std::abs(VerificationProofProgressPercent(task) - 50.0) > 0.001)
    {
        return false;
    }

    TaskSnapshot unknownBytes         = task;
    unknownBytes.discoveredBytesKnown = false;
    if (TaskHasKnownCompactProgress(unknownBytes) || ComputeFileOperationsTaskCompleteFractionForDisplay(unknownBytes) != 0.0f)
        return false;

    task.discoveryClosed = false;
    if (TaskHasKnownCompactProgress(task))
    {
        return false;
    }
    task.discoveryClosed = true;

    // A finished, cleanly verified Copy reads 100%, so both halves of its bar are full, also when its
    // completed record carried no verification bytes; a verification problem keeps the real counters.
    TaskSnapshot finishedVerified               = task;
    finishedVerified.finished                   = true;
    finishedVerified.resultHr                   = S_OK;
    finishedVerified.verificationActive         = false;
    finishedVerified.verificationState          = static_cast<uint8_t>(FileOperations::VerificationState::Verified);
    finishedVerified.verificationTotalBytes     = 0u;
    finishedVerified.verificationCompletedBytes = 0u;
    NormalizeCompletedTaskSnapshotForDisplay(finishedVerified);
    if (std::abs(ComputeFileOperationsTaskCompleteFractionForDisplay(finishedVerified) - 1.0f) > 0.001f ||
        std::abs(VerificationTransferProgressPercent(finishedVerified) - 100.0) > 0.001 ||
        std::abs(VerificationProofProgressPercent(finishedVerified) - 100.0) > 0.001)
    {
        return false;
    }
    TaskSnapshot finishedUnavailable               = finishedVerified;
    finishedUnavailable.verificationState          = static_cast<uint8_t>(FileOperations::VerificationState::Unavailable);
    finishedUnavailable.verificationTotalBytes     = 100u;
    finishedUnavailable.verificationCompletedBytes = 40u;
    NormalizeCompletedTaskSnapshotForDisplay(finishedUnavailable);
    if (std::abs(VerificationProofProgressPercent(finishedUnavailable) - 40.0) > 0.001)
    {
        return false;
    }
    // I35 (owner report 2026-10-13): a finished task reads 100% only when every byte was verified. Otherwise its
    // transfer half is full, its verification half and whole value show the share verified, and its compact label
    // names that share instead of a percent.
    if (std::abs(ComputeFileOperationsTaskCompleteFractionForDisplay(finishedUnavailable) - 0.7f) > 0.001f ||
        std::abs(TransferSegmentPercent(finishedUnavailable) - 100.0) > 0.001 || std::abs(VerificationSegmentPercent(finishedUnavailable) - 40.0) > 0.001 ||
        FinishedVerificationLabel(finishedUnavailable) != FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_VERIFIED_PERCENT, 40u) ||
        ! FinishedVerificationLabel(finishedVerified).empty())
    {
        return false;
    }
    TaskSnapshot finishedNotVerified               = finishedUnavailable;
    finishedNotVerified.verificationCompletedBytes = 0u;
    if (std::abs(ComputeFileOperationsTaskCompleteFractionForDisplay(finishedNotVerified) - 0.5f) > 0.001f ||
        FinishedVerificationLabel(finishedNotVerified) != LoadStringResource(nullptr, IDS_FILEOPS_COMPACT_NOT_VERIFIED))
    {
        return false;
    }
    // A Verify On copy of empty files only: verified by its record, though no byte was read back.
    TaskSnapshot emptyVerified               = finishedVerified;
    emptyVerified.totalBytes                 = 0u;
    emptyVerified.completedBytes             = 0u;
    emptyVerified.verificationTotalBytes     = 0u;
    emptyVerified.verificationCompletedBytes = 0u;
    if (std::abs(ComputeFileOperationsTaskCompleteFractionForDisplay(emptyVerified) - 1.0f) > 0.001f || ! FinishedVerificationLabel(emptyVerified).empty())
    {
        return false;
    }

    TaskSnapshot streamingTask{};
    streamingTask.taskId                                    = 2u;
    streamingTask.operation                                 = FILESYSTEM_COPY;
    streamingTask.started                                   = true;
    streamingTask.operationStartTick                        = 1u;
    streamingTask.hasProgressCallbacks                      = true;
    streamingTask.discoveryAheadActive                      = true;
    streamingTask.discoveryClosed                           = false;
    streamingTask.discoveredTotalBytes                      = 200u;
    streamingTask.completedBytes                            = 50u;
    const WholeTaskProgressPresentation provisionalProgress = ResolveWholeTaskProgressPresentation(streamingTask);
    bool discoveryGraphAnimation                            = true;
    const std::wstring discoveryGraphOverlay = GraphOverlayTextForStatus(streamingTask, ResolveTaskStatusKind(streamingTask), discoveryGraphAnimation);
    if (ResolveTaskStatusKind(streamingTask) != TaskStatusKind::Running || provisionalProgress.determinate || ! provisionalProgress.provisional ||
        ! provisionalProgress.transferActive || ! provisionalProgress.showDiscoveryActivity || ! provisionalProgress.showTransferProgress ||
        ! provisionalProgress.showTransferMarquee || ! provisionalProgress.showThroughputGraph || ! provisionalProgress.includeInAggregateCohort ||
        provisionalProgress.fraction != 0.0 || discoveryGraphAnimation || ! discoveryGraphOverlay.empty())
    {
        return false;
    }

    streamingTask.completedBytes                       = 0u;
    const WholeTaskProgressPresentation startupMarquee = ResolveWholeTaskProgressPresentation(streamingTask);
    if (startupMarquee.determinate || ! startupMarquee.provisional || ! startupMarquee.transferActive || ! startupMarquee.showTransferMarquee)
    {
        return false;
    }

    streamingTask.completedBytes                            = 240u;
    streamingTask.discoveredTotalBytes                      = 200u;
    const WholeTaskProgressPresentation rebasedOpenProgress = ResolveWholeTaskProgressPresentation(streamingTask);
    if (rebasedOpenProgress.determinate || ! rebasedOpenProgress.showTransferMarquee)
    {
        return false;
    }
    streamingTask.discoveryClosed                      = true;
    streamingTask.totalBytes                           = 240u;
    streamingTask.discoveredBytesKnown                 = true;
    const WholeTaskProgressPresentation closedProgress = ResolveWholeTaskProgressPresentation(streamingTask);
    if (! closedProgress.determinate || closedProgress.provisional || std::abs(closedProgress.fraction - 1.0) > 0.001)
    {
        return false;
    }
    streamingTask.discoveryClosed = false;
    streamingTask.totalBytes      = 0u;

    streamingTask.waitingForOthers                          = true;
    const WholeTaskProgressPresentation waitingPresentation = ResolveWholeTaskProgressPresentation(streamingTask);
    if (waitingPresentation.showDiscoveryActivity || waitingPresentation.showTransferProgress || waitingPresentation.showTransferMarquee ||
        waitingPresentation.showThroughputGraph || waitingPresentation.includeInAggregateCohort ||
        HasGlobalAggregateProgress(BuildGlobalStatusSummary({streamingTask})))
    {
        return false;
    }

    streamingTask.waitingForOthers                                  = false;
    streamingTask.conflict.active                                   = true;
    streamingTask.conflict.metadataLoading                          = true;
    const WholeTaskProgressPresentation loadingConflictPresentation = ResolveWholeTaskProgressPresentation(streamingTask);
    if (loadingConflictPresentation.showDiscoveryActivity || loadingConflictPresentation.showTransferProgress ||
        loadingConflictPresentation.showTransferMarquee || loadingConflictPresentation.showThroughputGraph ||
        loadingConflictPresentation.includeInAggregateCohort ||
        StatusTextForTask(streamingTask, TaskStatusKind::Conflict, 0u) != LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_READING_DETAILS))
    {
        return false;
    }

    streamingTask.conflict.metadataLoading = false;
    if (StatusTextForTask(streamingTask, TaskStatusKind::Conflict, 0u) != LoadStringResource(nullptr, IDS_FILEOPS_STATUS_WAITING_FOR_DECISION))
    {
        return false;
    }

    RateHistory history{};
    history.displayedBytesPerSec             = 10.0;
    history.displayedVerificationBytesPerSec = 5.0;
    AppendRateSample(history, 10.0, 5.0, -1.0f);
    const size_t newestIndex = (history.writeIndex + RateHistory::kMaxSamples - 1u) % RateHistory::kMaxSamples;
    if (history.count != 1u || history.samples[newestIndex] != 10.0f || history.verificationSamples[newestIndex] != 5.0f ||
        CurrentBandwidthForGraphMarker(history) != 15.0)
    {
        return false;
    }

    // Logical provider-proof completion advances the verification bar but never the verification
    // rate: the rate window measures only physical exact-object readback (I35 R10.2).
    RateSnapshot providerProofOnly{};
    providerProofOnly.taskId  = 1u;
    providerProofOnly.started = true;
    RateHistory providerProofHistory{};
    RateUpdateTotals providerProofTotals{};
    AdvanceRateHistory(providerProofHistory, providerProofOnly, 10'000ull, providerProofTotals);
    providerProofOnly.verificationCompletedBytes = 100ull * 1024ull * 1024ull;
    AdvanceRateHistory(providerProofHistory, providerProofOnly, 12'000ull, providerProofTotals);
    const double providerProofRate          = providerProofHistory.displayedVerificationBytesPerSec;
    providerProofOnly.verificationReadBytes = 40ull * 1024ull * 1024ull;
    AdvanceRateHistory(providerProofHistory, providerProofOnly, 14'000ull, providerProofTotals);
    if (providerProofRate != 0.0 || providerProofHistory.displayedVerificationBytesPerSec <= 0.0)
    {
        return false;
    }

    task.completedBytes = 80u;
    std::vector<TaskSnapshot> tasks{task};
    std::unordered_map<uint64_t, RateHistory> rates;
    rates.emplace(task.taskId, history);
    const GlobalFileOperationsStatusSummary summary = BuildGlobalStatusSummary(tasks, &rates);
    // Transfer is sequentially followed by verification: 20/10 + 50/5 = 12 seconds.
    if (! summary.hasAggregateEta || std::abs(summary.aggregateEtaSeconds - 12.0) >= 0.001)
    {
        return false;
    }

    TaskSnapshot terminalTask{};
    terminalTask.finished            = true;
    terminalTask.operation           = FILESYSTEM_MOVE;
    terminalTask.verificationState   = static_cast<uint8_t>(FileOperations::VerificationState::Failed);
    terminalTask.retainedSourceCount = 1u;
    terminalTask.unknownSourceCount  = 1u;
    std::array<TerminalAxisBadge, 3u> terminalBadges{};
    const size_t terminalBadgeCount = BuildTerminalAxisBadges(terminalTask, terminalBadges);
    if (terminalBadgeCount != 3u || terminalBadges[0].text != LoadStringResource(nullptr, IDS_FILEOPS_RESULT_VERIFICATION_FAILED) ||
        terminalBadges[1].text != LoadStringResource(nullptr, IDS_FILEOPS_RESULT_SOURCE_UNKNOWN) ||
        terminalBadges[2].text != LoadStringResource(nullptr, IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT) ||
        terminalBadges[0].tone != PopupStatusVisualTone::Error || terminalBadges[1].tone != PopupStatusVisualTone::Error ||
        terminalBadges[2].tone != PopupStatusVisualTone::Warning || ! TaskHasTerminalAxisBadges(terminalTask))
    {
        return false;
    }

    terminalTask.verificationState   = static_cast<uint8_t>(FileOperations::VerificationState::NotApplicable);
    terminalTask.retainedSourceCount = 0u;
    terminalTask.unknownSourceCount  = 0u;
    return BuildTerminalAxisBadges(terminalTask, terminalBadges) == 0u && ! TaskHasTerminalAxisBadges(terminalTask);
}

bool DebugValidateFileOperationsMissingSourcePresentation(std::wstring& failure) noexcept
{
    using FileOperationsPopupInternal::TaskSnapshot;
    failure.clear();
    TaskSnapshot terminalTask{};
    terminalTask.finished           = true;
    terminalTask.operation          = FILESYSTEM_MOVE;
    terminalTask.resultHr           = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    terminalTask.missingSourceCount = 1u;
    const auto missingFact          = FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_SOURCE_MISSING, 1ul);
    const auto retainedFact         = FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_SOURCE_RETAINED, 1ul);
    const auto missingFacts         = BuildTerminalFactsText(terminalTask);
    if (! TaskHasTerminalFacts(terminalTask) || missingFacts.find(missingFact) == std::wstring::npos || missingFacts.find(retainedFact) != std::wstring::npos)
    {
        failure = std::format(L"Missing-source card fact mismatch: actual='{}', missing='{}', retained='{}'.", missingFacts, missingFact, retainedFact);
        return false;
    }
    // A mixed result must retain both facts without claiming that every item moved.
    terminalTask.publishedItemCount = 1u;
    terminalTask.removedSourceCount = 1u;
    const auto mixedFacts           = BuildTerminalFactsText(terminalTask);
    if (MoveRemovedEveryPublishedSource(TerminalFactCounts(terminalTask)) || mixedFacts.find(missingFact) == std::wstring::npos ||
        mixedFacts.find(FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_SOURCE_REMOVED, 1ul)) == std::wstring::npos)
    {
        failure = std::format(L"Mixed missing-source card fact mismatch: actual='{}'.", mixedFacts);
        return false;
    }
    if (TaskHasTerminalAxisBadges(terminalTask))
    {
        failure = L"A missing source must not show a retained-source or verification badge.";
        return false;
    }
    return true;
}

float DebugComputeFileOperationsTaskCompleteFraction(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return ComputeFileOperationsTaskCompleteFractionForDisplay(task);
}

void DebugPublishFileOperationsPlannedItemTotalAfterDiscovery(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    PublishPlannedItemTotalAfterDiscovery(task);
}

bool DebugFileOperationsTaskHasKnownCompactProgress(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return TaskHasKnownCompactProgress(task);
}

void DebugResolveFileOperationsWholeTaskProgress(const FileOperationsPopupInternal::TaskSnapshot& task,
                                                 FileOperationsPopupInternal::TaskProgressPresentationDebugSnapshot& out) noexcept
{
    const WholeTaskProgressPresentation presentation = ResolveWholeTaskProgressPresentation(task);
    out.fraction                                     = presentation.fraction;
    out.determinate                                  = presentation.determinate;
    out.provisional                                  = presentation.provisional;
    out.transferActive                               = presentation.transferActive;
    out.showDiscoveryActivity                        = presentation.showDiscoveryActivity;
    out.showTransferProgress                         = presentation.showTransferProgress;
    out.showTransferMarquee                          = presentation.showTransferMarquee;
    out.showThroughputGraph                          = presentation.showThroughputGraph;
    out.includeInAggregateCohort                     = presentation.includeInAggregateCohort;
}

std::wstring DebugBuildFileOperationsTaskHeaderText(const FileOperationsPopupInternal::TaskSnapshot& task, ULONGLONG nowTick)
{
    const UINT opTextId = [&]() -> UINT
    {
        switch (task.operation)
        {
            case FILESYSTEM_COPY: return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
            case FILESYSTEM_MOVE: return static_cast<UINT>(IDS_FILEOP_OPERATION_MOVE);
            case FILESYSTEM_DELETE: return static_cast<UINT>(IDS_FILEOP_OPERATION_DELETE);
            case FILESYSTEM_RENAME: return static_cast<UINT>(IDS_FILEOP_OPERATION_RENAME);
            case FILESYSTEM_CREATE_DIRECTORY: return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
        }
        return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
    }();
    return BuildTaskHeaderText(task, LoadStringResource(nullptr, opTextId), nowTick);
}

std::wstring DebugFormatFileOperationsConflictTimestamp(__int64 fileTime) noexcept
{
    return FormatFileTimeLocalCompact(fileTime);
}

std::wstring DebugFormatFileOperationsDiscoveryProgress(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    return FormatDiscoveryProgressText(task);
}

std::wstring DebugFormatFileOperationsDiscoveredSoFar(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    return FormatDiscoveredSoFarText(task);
}

bool DebugRunFileOperationsRateModel(FileSystemOperation operation,
                                     std::span<const FileOperationsPopupInternal::RateModelDebugTick> ticks,
                                     std::vector<FileOperationsPopupInternal::RateModelDebugObservation>& out) noexcept
{
    out.clear();
    out.reserve(ticks.size());
    RateHistory history{};
    RateUpdateTotals totals{};
    for (const FileOperationsPopupInternal::RateModelDebugTick& tick : ticks)
    {
        RateSnapshot task{};
        task.taskId                   = 1u;
        task.operation                = operation;
        task.started                  = true;
        task.discoveryClosed          = tick.totalBytes != 0u;
        task.totalBytes               = tick.totalBytes;
        task.completedBytes           = tick.completedBytes;
        task.verificationReadBytes    = tick.verificationReadBytes;
        task.completedItems           = tick.completedItems;
        task.lastProgressCallbackTick = tick.lastProgressCallbackTick;
        task.lastByteProgressTick     = tick.lastByteProgressTick;
        task.progressStateChangeTick  = tick.progressStateChangeTick;
        task.stallKind                = tick.stallKind;
        task.speedLimitBytesPerSecond = tick.speedLimitBytesPerSecond;
        task.paused                   = tick.paused;
        task.waitingForDecision       = tick.waitingForDecision;
        task.stopping                 = tick.stopping;
        AdvanceRateHistory(history, task, tick.tick, totals);

        FileOperationsPopupInternal::RateModelDebugObservation observation{};
        observation.displayedBytesPerSec             = history.displayedBytesPerSec;
        observation.displayedVerificationBytesPerSec = history.displayedVerificationBytesPerSec;
        observation.displayedItemsPerSec             = history.displayedItemsPerSec;
        observation.stallKind                        = history.stallKind;
        observation.rateLabelText                    = FormatRateLabelText(CurrentBandwidthForGraphMarker(history), history.stallKind);
        observation.hasSmoothedEta                   = history.hasSmoothedEta;
        observation.smoothedEtaSeconds               = history.smoothedEtaSeconds;
        observation.sampleCount                      = history.count;
        for (size_t offset = 0u; offset < history.count; ++offset)
        {
            const size_t slot = (history.writeIndex + RateHistory::kMaxSamples - 1u - offset) % RateHistory::kMaxSamples;
            observation.stalledSampleCount += history.stalledSamples[slot] != 0u ? 1u : 0u;
        }
        if (history.count > 0u)
        {
            const size_t newest             = (history.writeIndex + RateHistory::kMaxSamples - 1u) % RateHistory::kMaxSamples;
            observation.newestSample        = history.samples[newest];
            observation.newestSampleStalled = history.stalledSamples[newest] != 0u;
        }
        out.push_back(std::move(observation));
    }
    return true;
}

double DebugWindowedRateAfterSilence(double steadyBytesPerSecond, ULONGLONG silenceMs) noexcept
{
    constexpr ULONGLONG kStartTick = 100'000ull;
    constexpr ULONGLONG kSteadyMs  = 10'000ull;
    constexpr ULONGLONG kStepMs    = 100ull;
    const auto bytesAt             = [steadyBytesPerSecond](ULONGLONG elapsedMs) noexcept
    { return static_cast<uint64_t>(std::llround(ClampFiniteNonNegative(steadyBytesPerSecond) * static_cast<double>(elapsedMs) / 1000.0)); };
    std::vector<FileOperationsPopupInternal::RateModelDebugTick> ticks;
    for (ULONGLONG elapsedMs = 0u; elapsedMs <= kSteadyMs; elapsedMs += kStepMs)
    {
        ticks.push_back({.tick                     = kStartTick + elapsedMs,
                         .completedBytes           = bytesAt(elapsedMs),
                         .lastProgressCallbackTick = kStartTick + elapsedMs,
                         .lastByteProgressTick     = kStartTick + elapsedMs});
    }
    const ULONGLONG silenceStart = kStartTick + kSteadyMs;
    for (ULONGLONG elapsedMs = kStepMs; elapsedMs <= silenceMs + kStepMs - 1u; elapsedMs += kStepMs)
    {
        ticks.push_back({.tick                     = silenceStart + std::min(elapsedMs, silenceMs),
                         .completedBytes           = bytesAt(kSteadyMs),
                         .lastProgressCallbackTick = silenceStart,
                         .lastByteProgressTick     = silenceStart});
    }
    std::vector<FileOperationsPopupInternal::RateModelDebugObservation> observations;
    if (! DebugRunFileOperationsRateModel(FILESYSTEM_COPY, ticks, observations) || observations.empty())
        return -1.0;
    return observations.back().displayedBytesPerSec;
}

std::wstring DebugFormatFileOperationsRateLabelText(double bytesPerSecond, FileOperationsPopupInternal::RateStallKind stall)
{
    return FormatRateLabelText(bytesPerSecond, stall);
}

double DebugSmoothEtaSecondsForDisplay(double previousEtaSeconds, double sampleEtaSeconds, ULONGLONG elapsedMs) noexcept
{
    return SmoothEtaSecondsForDisplay(previousEtaSeconds, sampleEtaSeconds, elapsedMs);
}

float DebugEaseFileOperationsGraphLatestPointYForDisplay(float previousY, float targetY, ULONGLONG elapsedMs) noexcept
{
    return EaseGraphLatestPointYForDisplay(previousY, targetY, elapsedMs);
}

float DebugEaseFileOperationsAutoResizeFraction(ULONGLONG elapsedMs, ULONGLONG durationMs) noexcept
{
    return EaseFileOperationsUiMotionFraction(elapsedMs, durationMs);
}

D2D1_RECT_F DebugComputeFileOperationsIndeterminateBarFill(const D2D1_RECT_F& bar, ULONGLONG tick, bool reducedMotion) noexcept
{
    return ComputeIndeterminateBarFill(bar, tick, reducedMotion);
}

HWND GetFileOperationsSpeedLimitPromptHandle() noexcept
{
    const HWND hwnd = FindFileOperationsDebugWindowForCurrentProcess(kFileOperationsSpeedLimitPromptClassName);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

void DebugShowFileOperationsSpeedLimitPromptForGallery(HWND owner, const AppTheme& theme) noexcept
{
    static_cast<void>(ShowCustomSpeedLimitPrompt(owner, theme, 80ull * 1024ull * 1024ull));
}

bool DebugGetFileOperationsSpeedLimitPromptSnapshot(FileOperationsSpeedLimitPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto snapshot = std::make_unique<FileOperationsSpeedLimitPromptDebugSnapshot>();
    const bool ok = SendMessageW(hwnd,
                                 kFileOperationsSpeedLimitPromptDebugMessage,
                                 static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::GetSnapshot),
                                 reinterpret_cast<LPARAM>(snapshot.get())) != FALSE;
    if (ok)
    {
        out = std::move(*snapshot);
    }
    return ok;
}

bool DebugSetFileOperationsSpeedLimitPromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    if (! hwnd)
    {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(text);
    return SendMessageW(hwnd,
                        kFileOperationsSpeedLimitPromptDebugMessage,
                        static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::SetText),
                        reinterpret_cast<LPARAM>(payload.get())) != FALSE;
}

bool DebugSetFileOperationsSpeedLimitPromptClientHeightLimit(std::optional<float> heightDip) noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    return hwnd && SendMessageW(hwnd,
                                kFileOperationsSpeedLimitPromptDebugMessage,
                                static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::SetClientHeightLimit),
                                reinterpret_cast<LPARAM>(&heightDip)) != FALSE;
}

bool DebugSetFileOperationsSpeedLimitPromptSelection(size_t start, size_t end) noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    const std::pair<size_t, size_t> selection{start, end};
    return hwnd && SendMessageW(hwnd,
                                kFileOperationsSpeedLimitPromptDebugMessage,
                                static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::SetSelection),
                                reinterpret_cast<LPARAM>(&selection)) != FALSE;
}

bool DebugConfirmFileOperationsSpeedLimitPrompt() noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    return PostDxUiPromptCloseDebugCommand(
        hwnd, kFileOperationsSpeedLimitPromptDebugMessage, static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::Confirm));
}

bool DebugCancelFileOperationsSpeedLimitPrompt() noexcept
{
    const HWND hwnd = GetFileOperationsSpeedLimitPromptHandle();
    return PostDxUiPromptCloseDebugCommand(
        hwnd, kFileOperationsSpeedLimitPromptDebugMessage, static_cast<WPARAM>(FileOperationsSpeedLimitPromptWindow::DebugCommand::Cancel));
}
#endif
