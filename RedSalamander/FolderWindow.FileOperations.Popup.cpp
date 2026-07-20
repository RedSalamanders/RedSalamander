#include "FolderWindow.FileOperations.Popup.h"

#include "DxUi/DxUi.Typography.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "FluentIcons.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "ThroughputParsing.h"
#include "WindowMaximizeBehavior.h"
#include "WindowSizing.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <d2d1.h>
#include <dwrite.h>
#include <limits>
#include <shobjidl.h>
#include <unordered_map>
#include <windowsx.h>

namespace
{
constexpr wchar_t kFileOperationsPopupClassName[] = L"RedSalamander.FileOperationsPopup";

constexpr UINT_PTR kFileOperationsPopupTimerId                      = 1;
constexpr UINT kFileOperationsPopupTimerIntervalMs                  = 100;
constexpr ULONGLONG kRateSampleBucketMs                             = 100ull;
constexpr UINT kFileOperationsPopupDeferredSpeedLimitPromptMessage  = WM_APP + 0x71;
constexpr wchar_t kFileOperationsSpeedLimitPromptClassName[]        = L"RedSalamander.FileOperations.SpeedLimitPrompt";
constexpr UINT kFileOperationsSpeedLimitPromptDeferredActionMessage = WM_APP + 0x74;
constexpr std::wstring_view kEllipsisText                           = L"\u2026";
constexpr float kFileOperationsPopupFooterHeightDip                 = 88.0f;
constexpr int kFileOperationsPopupMinClientHeightDip                = 320;
constexpr int kFileOperationsPopupFooterOnlyMinClientHeightDip      = 96;
constexpr uint32_t kCompletedOverflowActionShowLog                  = 1u;
constexpr uint32_t kCompletedOverflowActionExportIssues             = 2u;
constexpr uint32_t kCompletedOverflowActionFailedItems              = 3u;
constexpr uint32_t kCompletedOverflowActionOpenDestination          = 4u;
constexpr uint32_t kCompletedOverflowActionRevealDestination        = 5u;
constexpr uint32_t kFooterPauseResumeAllPauseAction                 = 1u;
constexpr uint32_t kFooterPauseResumeAllResumeAction                = 2u;
constexpr uint32_t kFooterQueueModeQueueAction                      = 1u;
constexpr uint32_t kFooterQueueModeParallelAction                   = 2u;
constexpr ULONGLONG kTaskbarListRetryDelayMs                        = 1000ull;

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

float FileOperationsPopupFooterHeightPixels(UINT dpi) noexcept
{
    return DipsToPixels(kFileOperationsPopupFooterHeightDip, dpi);
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

float ComputeFileOperationsTaskCompleteFractionForDisplay(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (task.finished && SUCCEEDED(task.resultHr))
    {
        return 1.0f;
    }

    if (task.operation == FILESYSTEM_DELETE)
    {
        if (task.totalBytes > 0 && task.completedBytes > 0)
        {
            return Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
        }
        if (task.totalItems > 0)
        {
            return Clamp01(static_cast<float>(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems)));
        }
        return 0.0f;
    }

    if (task.totalBytes > 0)
    {
        return Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
    }
    if (task.totalItems > 0)
    {
        return Clamp01(static_cast<float>(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems)));
    }

    return 0.0f;
}

void NormalizeCompletedTaskSnapshotForDisplay(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    // Succeeded-with-warnings tasks keep their REAL counters: forcing "N of N"/100% would hide
    // that children were skipped. Only a clean success rounds the display up.
    if (! task.finished || FAILED(task.resultHr) || task.warningCount > 0 || task.errorCount > 0)
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
    return task.warningCount > 0 || task.errorCount > 0 || CompletedTaskCanUseDestinationActions(task) || ResolveCompletedTaskRevealLocation(task).has_value();
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

struct ConflictActionLayout
{
    std::array<ConflictAction, FileOperationsPopupInternal::TaskSnapshot::kMaxConflictActions> primary{};
    std::array<ConflictAction, FileOperationsPopupInternal::TaskSnapshot::kMaxConflictActions> overflow{};
    size_t primaryCount  = 0u;
    size_t overflowCount = 0u;
};

[[nodiscard]] ConflictAction ConflictActionFromRaw(uint8_t rawAction) noexcept
{
    return static_cast<ConflictAction>(rawAction);
}

[[nodiscard]] uint8_t RawConflictAction(ConflictAction action) noexcept
{
    return static_cast<uint8_t>(action);
}

[[nodiscard]] bool ConflictPromptHasAction(const FileOperationsPopupInternal::TaskSnapshot::ConflictPromptSnapshot& conflict, ConflictAction action) noexcept
{
    for (size_t i = 0; i < conflict.actionCount && i < conflict.actions.size(); ++i)
    {
        if (ConflictActionFromRaw(conflict.actions[i]) == action)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool ConflictLayoutContains(const std::array<ConflictAction, FileOperationsPopupInternal::TaskSnapshot::kMaxConflictActions>& actions,
                                          size_t actionCount,
                                          ConflictAction action) noexcept
{
    for (size_t i = 0; i < actionCount && i < actions.size(); ++i)
    {
        if (actions[i] == action)
        {
            return true;
        }
    }

    return false;
}

void AppendConflictLayoutAction(std::array<ConflictAction, FileOperationsPopupInternal::TaskSnapshot::kMaxConflictActions>& actions,
                                size_t& actionCount,
                                ConflictAction action) noexcept
{
    if (action == ConflictAction::None || ConflictLayoutContains(actions, actionCount, action))
    {
        return;
    }

    if (actionCount < actions.size())
    {
        actions[actionCount] = action;
    }
    ++actionCount;
}

void AppendPrimaryConflictActionIfAvailable(ConflictActionLayout& layout,
                                            const FileOperationsPopupInternal::TaskSnapshot::ConflictPromptSnapshot& conflict,
                                            ConflictAction action) noexcept
{
    if (layout.primaryCount >= 3u || ! ConflictPromptHasAction(conflict, action))
    {
        return;
    }

    AppendConflictLayoutAction(layout.primary, layout.primaryCount, action);
}

[[nodiscard]] ConflictActionLayout BuildConflictActionLayout(const FileOperationsPopupInternal::TaskSnapshot::ConflictPromptSnapshot& conflict) noexcept
{
    ConflictActionLayout layout{};
    if (! conflict.active)
    {
        return layout;
    }

    switch (static_cast<ConflictBucket>(conflict.bucket))
    {
        case ConflictBucket::Exists:
        case ConflictBucket::NonEmptyDirectory:
        case ConflictBucket::ReparsePoint: AppendPrimaryConflictActionIfAvailable(layout, conflict, ConflictAction::Overwrite); break;
        case ConflictBucket::RecycleBinFailed: AppendPrimaryConflictActionIfAvailable(layout, conflict, ConflictAction::PermanentDelete); break;
        case ConflictBucket::AccessDenied:
        case ConflictBucket::SharingViolation:
        case ConflictBucket::DiskFull:
        case ConflictBucket::PathTooLong:
        case ConflictBucket::NetworkOffline:
        case ConflictBucket::UnsupportedReparse:
        case ConflictBucket::Unknown:
        case ConflictBucket::Count: AppendPrimaryConflictActionIfAvailable(layout, conflict, ConflictAction::Retry); break;
        case ConflictBucket::ReadOnly:
        default: AppendPrimaryConflictActionIfAvailable(layout, conflict, ConflictAction::Retry); break;
    }

    AppendPrimaryConflictActionIfAvailable(layout, conflict, ConflictAction::Skip);
    AppendPrimaryConflictActionIfAvailable(layout, conflict, ConflictAction::Cancel);

    for (size_t i = 0; i < conflict.actionCount && i < conflict.actions.size(); ++i)
    {
        const ConflictAction action = ConflictActionFromRaw(conflict.actions[i]);
        if (action == ConflictAction::None || ConflictLayoutContains(layout.primary, layout.primaryCount, action))
        {
            continue;
        }

        AppendConflictLayoutAction(layout.overflow, layout.overflowCount, action);
    }

    return layout;
}

[[nodiscard]] std::wstring ConflictActionText(ConflictAction action) noexcept
{
    switch (action)
    {
        case ConflictAction::Overwrite: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_OVERWRITE);
        case ConflictAction::ReplaceReadOnly: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_REPLACE_READONLY);
        case ConflictAction::PermanentDelete: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_PERMANENT_DELETE);
        case ConflictAction::Retry: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_RETRY);
        case ConflictAction::Skip: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_SKIP);
        case ConflictAction::Cancel: return LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);
        case ConflictAction::None:
        default: break;
    }

    return LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL);
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
    uint32_t running                   = 0;
    uint32_t waiting                   = 0;
    uint32_t needAttention             = 0;
    uint32_t activeRunning             = 0;
    uint32_t paused                    = 0;
    uint32_t pauseEligibleRunning      = 0;
    uint32_t pauseEligiblePaused       = 0;
    uint64_t completedBytes            = 0;
    uint64_t totalBytes                = 0;
    uint64_t completedItems            = 0;
    uint64_t totalItems                = 0;
    bool hasUnknownActiveProgress      = false;
    bool hasUnknownActiveTransferBytes = false;
    bool hasActiveTaskbarProgress      = false;
    double displayedBytesPerSec        = 0.0;
    bool hasAggregateEta               = false;
    double aggregateEtaSeconds         = 0.0;
};

struct GlobalTaskbarProgressModel
{
    uint32_t state     = static_cast<uint32_t>(TBPF_NOPROGRESS);
    uint64_t completed = 0;
    uint64_t total     = 0;
};

[[nodiscard]] bool TaskHasPublishedProgressNumbers(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.completedItems > 0 || task.completedBytes > 0 || task.totalItems > 0 || task.totalBytes > 0;
}

[[nodiscard]] bool TaskHasKnownCompactProgress(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.totalBytes > 0 || task.totalItems > 0;
}

void PublishPlannedItemTotalAfterPreCalculation(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (! task.preCalcInProgress && task.totalItems == 0 && task.operation != FILESYSTEM_DELETE)
    {
        task.totalItems = task.plannedItems;
    }
}

[[nodiscard]] bool TaskShowsPreparingStatus(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return ! task.started || ! task.hasProgressCallbacks || (task.operation == FILESYSTEM_DELETE && ! TaskHasPublishedProgressNumbers(task));
}

[[nodiscard]] TaskStatusKind ResolveTaskStatusKind(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    const HRESULT partialHr   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);

    if (task.finished)
    {
        if (task.resultHr == cancelledHr || task.resultHr == E_ABORT)
        {
            return TaskStatusKind::Canceled;
        }
        if (task.resultHr == partialHr || (SUCCEEDED(task.resultHr) && (task.warningCount > 0 || task.errorCount > 0)))
        {
            return TaskStatusKind::Partial;
        }
        if (FAILED(task.resultHr))
        {
            return TaskStatusKind::Failed;
        }
        return TaskStatusKind::Done;
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
    // 5F early admission: pre-calc now runs concurrently with the transfer. The blocking
    // "Calculating" status is only truthful before the transfer has actually started; once
    // ExecuteOperation is underway (operationStartTick set) show the live transfer status, whose
    // ETA already reads "estimating" until pre-calc totals settle.
    if (task.preCalcInProgress && task.operationStartTick == 0)
    {
        return TaskStatusKind::Calculating;
    }
    if (TaskShowsPreparingStatus(task))
    {
        return TaskStatusKind::Preparing;
    }
    return TaskStatusKind::Running;
}

void EnsureFinishedTaskDiagnosticAffordance(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    if (! task.finished || task.warningCount > 0 || task.errorCount > 0)
    {
        return;
    }

    const HRESULT partialHr   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    if (task.resultHr == partialHr)
    {
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
            task.lastDiagnosticMessage = FormatStringResource(nullptr, IDS_FMT_FILEOPS_DIAG_FAILED_NO_DETAILS, static_cast<unsigned long>(task.resultHr));
        }
    }
}

[[nodiscard]] bool StatusNeedsAttention(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Conflict || status == TaskStatusKind::Partial || status == TaskStatusKind::Failed;
}

[[nodiscard]] bool StatusCountsAsRunning(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Calculating || status == TaskStatusKind::Preparing || status == TaskStatusKind::Running ||
           status == TaskStatusKind::Paused;
}

[[nodiscard]] bool StatusCountsAsWaiting(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Waiting;
}

[[nodiscard]] bool StatusIsWarning(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Conflict || status == TaskStatusKind::Partial;
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
        case TaskStatusKind::Calculating:
        case TaskStatusKind::Preparing: return PopupStatusVisualTone::Accent;
        case TaskStatusKind::Waiting:
        case TaskStatusKind::Canceled: return PopupStatusVisualTone::Neutral;
        case TaskStatusKind::Paused: return PopupStatusVisualTone::Muted;
        case TaskStatusKind::Conflict:
        case TaskStatusKind::Partial: return PopupStatusVisualTone::Warning;
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
    switch (status)
    {
        case TaskStatusKind::Waiting: return LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_WAITING);
        case TaskStatusKind::Paused: return LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_PAUSED);
        case TaskStatusKind::Conflict: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_NEEDS_ATTENTION);
        case TaskStatusKind::Calculating:
        {
            const uint64_t elapsedSec = task.preCalcElapsedMs / 1000;
            return elapsedSec > 0 ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_CALCULATING_TIME, FormatDurationHms(elapsedSec))
                                  : LoadStringResource(nullptr, IDS_FILEOPS_CALCULATING);
        }
        case TaskStatusKind::Preparing:
        {
            const ULONGLONG opStartTick = task.operationStartTick;
            const uint64_t elapsedSec   = (opStartTick > 0 && nowTick >= opStartTick) ? static_cast<uint64_t>((nowTick - opStartTick) / 1000ull) : 0ull;
            return elapsedSec > 0 ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_PREPARING_TIME, FormatDurationHms(elapsedSec))
                                  : LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
        }
        case TaskStatusKind::Done: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
        case TaskStatusKind::Partial: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_PARTIAL);
        case TaskStatusKind::Failed: return FormatStringResource(nullptr, IDS_FMT_FILEOPS_STATUS_FAILED, static_cast<unsigned long>(task.resultHr));
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
        case TaskStatusKind::Waiting: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_WAITING);
        case TaskStatusKind::Partial: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_PARTIAL_SHORT);
        case TaskStatusKind::Failed: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_FAILED_SHORT);
        case TaskStatusKind::Paused:
        case TaskStatusKind::Conflict:
        case TaskStatusKind::Calculating:
        case TaskStatusKind::Preparing:
        case TaskStatusKind::Done:
        case TaskStatusKind::Canceled: return StatusTextForTask(task, status, nowTick);
        case TaskStatusKind::None:
        default: break;
    }

    return {};
}

[[nodiscard]] std::wstring BuildTaskHeaderText(const FileOperationsPopupInternal::TaskSnapshot& task, std::wstring_view operationText, ULONGLONG nowTick)
{
    const TaskStatusKind status = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
    if (status == TaskStatusKind::Running || status == TaskStatusKind::None)
    {
        if (task.totalItems > 0)
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
        case TaskStatusKind::Calculating: showAnimation = true; return LoadStringResource(nullptr, IDS_FILEOPS_GRAPH_CALCULATING);
        case TaskStatusKind::Preparing: showAnimation = true; return LoadStringResource(nullptr, IDS_FILEOPS_PREPARING);
        case TaskStatusKind::Conflict: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_NEEDS_ATTENTION);
        case TaskStatusKind::Running:
        case TaskStatusKind::Done:
        case TaskStatusKind::Partial:
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
    double aggregateBytesPerSec    = 0.0;
    double etaBytesPerSec          = 0.0;
    double aggregateRemainingBytes = 0.0;
    bool etaInputsComplete         = true;
    for (const auto& task : snapshot)
    {
        if (task.kind != FileOperationsPopupInternal::TaskSnapshot::Kind::FileOperation)
        {
            continue;
        }

        const TaskStatusKind status = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
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
                if (task.started)
                {
                    ++summary.pauseEligibleRunning;
                }
            }
        }

        if (StatusNeedsAttention(status) || StatusCountsAsWaiting(status) || StatusCountsAsRunning(status))
        {
            summary.hasActiveTaskbarProgress = true;
        }

        if (task.totalBytes > 0)
        {
            summary.totalBytes += task.totalBytes;
            summary.completedBytes += std::min(task.completedBytes, task.totalBytes);
        }
        if (task.totalItems > 0)
        {
            summary.totalItems += task.totalItems;
            summary.completedItems += std::min<uint64_t>(task.completedItems, task.totalItems);
        }
        if (! task.finished && task.totalBytes == 0 && task.totalItems == 0 && (StatusCountsAsRunning(status) || StatusCountsAsWaiting(status)))
        {
            summary.hasUnknownActiveProgress = true;
        }
        if ((task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) && task.totalBytes == 0 &&
            (StatusCountsAsRunning(status) || StatusCountsAsWaiting(status)))
        {
            summary.hasUnknownActiveTransferBytes = true;
        }

        if (! rates || task.finished || (task.operation != FILESYSTEM_COPY && task.operation != FILESYSTEM_MOVE))
        {
            continue;
        }

        const auto rateIt = rates->find(task.taskId);
        if (rateIt == rates->end())
        {
            if (task.totalBytes > task.completedBytes)
            {
                etaInputsComplete = false;
            }
            continue;
        }

        const double bytesPerSec = ClampFiniteNonNegative(rateIt->second.displayedBytesPerSec);
        aggregateBytesPerSec += bytesPerSec;
        if (task.totalBytes > 0 && task.completedBytes < task.totalBytes)
        {
            if (! IsByteRateUsableForEta(bytesPerSec))
            {
                etaInputsComplete = false;
                continue;
            }
            aggregateRemainingBytes += static_cast<double>(task.totalBytes - task.completedBytes);
            etaBytesPerSec += bytesPerSec;
        }
    }

    if (rates)
    {
        summary.displayedBytesPerSec = aggregateBytesPerSec;
        if (! summary.hasUnknownActiveTransferBytes && etaInputsComplete && IsByteRateUsableForEta(etaBytesPerSec) && aggregateRemainingBytes > 0.0)
        {
            summary.hasAggregateEta     = true;
            summary.aggregateEtaSeconds = aggregateRemainingBytes / etaBytesPerSec;
        }
    }

    return summary;
}

[[nodiscard]] bool HasGlobalAggregateProgress(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    return summary.totalBytes > 0 || summary.totalItems > 0 || summary.hasUnknownActiveProgress || summary.running > 0 || summary.waiting > 0 ||
           summary.needAttention > 0;
}

[[nodiscard]] bool HasDeterminateGlobalAggregateProgress(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    return ! summary.hasUnknownActiveProgress && (summary.totalBytes > 0 || summary.totalItems > 0);
}

[[nodiscard]] bool HasFooterPauseResumeAllControl(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    return summary.pauseEligibleRunning > 0 || summary.pauseEligiblePaused > 0;
}

[[nodiscard]] bool FooterPauseResumeAllShouldPause(const GlobalFileOperationsStatusSummary& summary) noexcept
{
    return summary.pauseEligibleRunning > 0;
}

[[nodiscard]] bool IsCompletedGroupTask(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.kind == FileOperationsPopupInternal::TaskSnapshot::Kind::FileOperation && task.finished;
}

[[nodiscard]] uint32_t CountCompletedGroupTasks(const std::vector<FileOperationsPopupInternal::TaskSnapshot>& snapshot) noexcept
{
    uint32_t count = 0u;
    for (const auto& task : snapshot)
    {
        if (IsCompletedGroupTask(task))
        {
            ++count;
        }
    }
    return count;
}

[[nodiscard]] bool ShouldShowCompletedGroup(uint32_t completedCount) noexcept
{
    return completedCount >= 2u;
}

[[nodiscard]] float GlobalAggregateProgressFraction(const GlobalFileOperationsStatusSummary& summary) noexcept
{
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
    const std::wstring statusText = FormatStringResource(nullptr,
                                                         IDS_FMT_FILEOPS_GLOBAL_STATUS_SUMMARY,
                                                         static_cast<unsigned long>(summary.running),
                                                         static_cast<unsigned long>(summary.waiting),
                                                         static_cast<unsigned long>(summary.needAttention));
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
    if (! summary.hasActiveTaskbarProgress)
    {
        return model;
    }

    if (! summary.hasUnknownActiveProgress && summary.totalBytes > 0)
    {
        model.completed = summary.completedBytes;
        model.total     = summary.totalBytes;
    }
    else if (! summary.hasUnknownActiveProgress && summary.totalItems > 0)
    {
        model.completed = summary.completedItems;
        model.total     = summary.totalItems;
    }

    if (summary.needAttention > 0)
    {
        model.state = static_cast<uint32_t>(TBPF_ERROR);
        return model;
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

        const HWND hwnd = CreateWindowExW(exStyle,
                                          kFileOperationsSpeedLimitPromptClassName,
                                          LoadStringResource(nullptr, IDS_CAPTION_FILEOP_SPEED_LIMIT_CUSTOM).c_str(),
                                          style,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
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
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
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
        }

        return _result;
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
    };
#endif

private:
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

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _label = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_LABEL));
        _label->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _field = _root->AddChild<TextField>(_initialLimitBytesPerSecond == 0 ? L"0" : FormatBytesCompact(_initialLimitBytesPerSecond));
        _field->SetOnTextChanged([this](std::wstring_view) { ClearValidation(); });

        _hintLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_HINT));
        _hintLabel->SetMultiline(true);
        _hintLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _hintLabel->SetFontRole(FontRole::Small);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { RequestDeferredAction(DeferredAction::Confirm); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { RequestDeferredAction(DeferredAction::Cancel); });

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

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip       = 16.0f;
        constexpr float kGapDip          = 10.0f;
        constexpr float kFieldHeightDip  = 32.0f;
        constexpr float kButtonHeightDip = 34.0f;
        constexpr float kButtonWidthDip  = 96.0f;
        constexpr float kLabelHeightDip  = 24.0f;
        constexpr float kHintHeightDip   = 46.0f;

        const float left  = client.left + kMarginDip;
        const float right = client.right - kMarginDip;
        float y           = client.top + kMarginDip;

        _label->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        y += kLabelHeightDip + 4.0f;

        _field->SetBounds(D2D1::RectF(left, y, right, y + kFieldHeightDip));
        y += kFieldHeightDip + kGapDip;

        _hintLabel->SetBounds(D2D1::RectF(left, y, right, y + kHintHeightDip));

        const float buttonsTop = client.bottom - kMarginDip - kButtonHeightDip;
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
        _validationText.clear();
        UpdateHintVisuals();
        _dxHost.Invalidate();
    }

    void ShowValidation(UINT messageId) noexcept
    {
        _showingValidationError = true;
        _validationText         = LoadStringResource(nullptr, messageId);
        UpdateHintVisuals();
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
                snapshot->hintText       = _showingValidationError ? std::wstring{} : LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_PROMPT_HINT);
                snapshot->validationText = _validationText;
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
            case DebugCommand::Confirm: RequestDeferredAction(DeferredAction::Confirm); return TRUE;
            case DebugCommand::Cancel: RequestDeferredAction(DeferredAction::Cancel); return TRUE;
        }

        return FALSE;
    }
#endif

    LRESULT WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
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
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root = nullptr;
    RedSalamander::DxUi::ThemePalette _palette{};
    RedSalamander::DxUi::Label* _label         = nullptr;
    RedSalamander::DxUi::TextField* _field     = nullptr;
    RedSalamander::DxUi::Label* _hintLabel     = nullptr;
    RedSalamander::DxUi::Button* _okButton     = nullptr;
    RedSalamander::DxUi::Button* _cancelButton = nullptr;
    std::wstring _validationText;
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

std::wstring TruncateTextMiddleToWidth(IDWriteFactory* factory,
                                       IDWriteTextFormat* format,
                                       std::wstring_view text,
                                       float maxWidth,
                                       float height,
                                       std::wstring_view ellipsisText,
                                       size_t fixedPrefixChars,
                                       size_t minSuffixChars) noexcept
{
    const float fullWidth = MeasureTextWidth(factory, format, text, maxWidth, height);
    if (fullWidth <= maxWidth)
    {
        return std::wstring(text);
    }

    const float dotsWidth = MeasureTextWidth(factory, format, ellipsisText, maxWidth, height);
    if (dotsWidth <= 0.0f || maxWidth <= dotsWidth)
    {
        return std::wstring(ellipsisText);
    }

    fixedPrefixChars = std::min(fixedPrefixChars, text.size());
    minSuffixChars   = std::min(minSuffixChars, text.size());

    if (fixedPrefixChars + minSuffixChars > text.size())
    {
        const size_t overlap = fixedPrefixChars + minSuffixChars - text.size();
        const size_t reduce  = std::min(overlap, fixedPrefixChars);
        fixedPrefixChars -= reduce;
    }

    const std::wstring_view prefix = text.substr(0, fixedPrefixChars);

    const float prefixWidth = MeasureTextWidth(factory, format, prefix, maxWidth, height);
    if (prefixWidth + dotsWidth >= maxWidth)
    {
        return std::wstring(ellipsisText);
    }

    size_t low  = minSuffixChars;
    size_t high = text.size() - fixedPrefixChars;

    while (low < high)
    {
        const size_t mid               = (low + high + 1u) / 2u;
        const std::wstring_view suffix = text.substr(text.size() - mid);

        std::wstring candidate;
        candidate.reserve(prefix.size() + ellipsisText.size() + suffix.size());
        candidate.append(prefix);
        candidate.append(ellipsisText);
        candidate.append(suffix);

        const float w = MeasureTextWidth(factory, format, candidate, maxWidth, height);
        if (w <= maxWidth)
        {
            low = mid;
        }
        else
        {
            high = mid - 1u;
        }
    }

    const std::wstring_view suffix = text.substr(text.size() - low);
    std::wstring result;
    result.reserve(prefix.size() + ellipsisText.size() + suffix.size());
    result.append(prefix);
    result.append(ellipsisText);
    result.append(suffix);
    return result;
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

size_t ComputePathLeafChars(std::wstring_view path) noexcept
{
    std::wstring_view trimmed = path;
    while (! trimmed.empty())
    {
        const wchar_t last = trimmed.back();
        if (last != L'\\' && last != L'/')
        {
            break;
        }
        trimmed.remove_suffix(1);
    }

    const size_t pos = trimmed.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return trimmed.size();
    }

    if (pos + 1u >= trimmed.size())
    {
        return 0u;
    }

    return trimmed.size() - (pos + 1u);
}

D2D1::ColorF RainbowProgressColor(const AppTheme& theme, std::wstring_view seed) noexcept
{
    if (seed.empty())
    {
        return theme.navigationView.accent;
    }

    const uint32_t hash = StableHash32(seed);
    const float hue     = static_cast<float>(hash % 360u);
    const float sat     = 0.85f;
    const float val     = theme.dark ? 0.80f : 0.90f;
    return ColorFromHSV(hue, sat, val, 1.0f);
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

bool IsRateSamplingBlocked(const FileOperationsPopupInternal::RateSnapshot& task) noexcept
{
    return task.paused || task.queuePaused || task.waitingInQueue;
}

void AddHueWeight(std::array<FileOperationsPopupInternal::RateHistory::HueWeight, FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample>& weights,
                  size_t& weightCount,
                  float hue,
                  double weight) noexcept
{
    if (hue < 0.0f || ! std::isfinite(weight) || weight <= 0.0)
    {
        return;
    }

    for (size_t i = 0; i < weightCount && i < weights.size(); ++i)
    {
        if (weights[i].hue == hue)
        {
            weights[i].weight += weight;
            return;
        }
    }

    if (weightCount < weights.size())
    {
        weights[weightCount] = {.hue = hue, .weight = weight};
        ++weightCount;
    }
}

void AddPendingHueWeight(FileOperationsPopupInternal::RateHistory& history, float hue, double weight) noexcept
{
    AddHueWeight(history.pendingHueWeights, history.pendingHueWeightCount, hue, weight);
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

void ResetRateStreamProgress(FileOperationsPopupInternal::RateHistory& history) noexcept
{
    history.streamProgressCount = 0u;
}

void SyncRateStreamBaselines(FileOperationsPopupInternal::RateHistory& history, const FileOperationsPopupInternal::RateSnapshot& task) noexcept
{
    history.streamProgressCount = 0u;
    for (size_t i = 0; i < task.inFlightFileCount && i < task.inFlightFiles.size() && history.streamProgressCount < history.streamProgress.size(); ++i)
    {
        const auto& stream     = task.inFlightFiles[i];
        auto& entry            = history.streamProgress[history.streamProgressCount++];
        entry.cookieKey        = stream.cookieKey;
        entry.progressStreamId = stream.progressStreamId;
        entry.sourcePath       = stream.sourcePath;
        entry.completedBytes   = stream.completedBytes;
        entry.lastUpdateTick   = stream.lastUpdateTick;
    }
}

// Golden-angle sequence keeps concurrent stream hues well separated and deterministic.
[[nodiscard]] float AssignStreamHue(FileOperationsPopupInternal::RateHistory& history) noexcept
{
    constexpr float kGoldenAngleDegrees = 137.5083f;
    const float hue                     = std::fmod(static_cast<float>(history.hueAssignmentCounter) * kGoldenAngleDegrees, 360.0f);
    ++history.hueAssignmentCounter;
    return hue;
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

        const bool sameItem   = baseline && baseline->sourcePath == stream.sourcePath;
        const float streamHue = sameItem && baseline->assignedHue >= 0.0f ? baseline->assignedHue : AssignStreamHue(history);

        if (stream.completedBytes > 0)
        {
            AddPendingHueWeight(history, streamHue, static_cast<double>(stream.completedBytes));
        }

        if (history.streamProgressCount < history.streamProgress.size())
        {
            auto& entry            = history.streamProgress[history.streamProgressCount++];
            entry.cookieKey        = stream.cookieKey;
            entry.progressStreamId = stream.progressStreamId;
            entry.sourcePath       = stream.sourcePath;
            entry.completedBytes   = stream.completedBytes;
            entry.lastUpdateTick   = stream.lastUpdateTick;
            entry.assignedHue      = streamHue;
        }
    }

    // When throttling delays the first byte-progress callback, concurrent streams are still real
    // live graph bands. Preserve their distinct hues instead of collapsing the bucket to the
    // aggregate fallback color.
    if (history.pendingHueWeightCount == 0u && history.streamProgressCount > 1u)
    {
        for (size_t i = 0; i < history.streamProgressCount && i < history.streamProgress.size(); ++i)
        {
            AddPendingHueWeight(history, history.streamProgress[i].assignedHue, 1.0);
        }
    }

    // No stream data at all (e.g. providers that never report streams): fall back to the
    // aggregate so the graph still gets a color. AppendRateSample carries the previous bucket's
    // distribution when even that is absent.
    if (history.pendingHueWeightCount == 0u && aggregateDeltaBytes > 0)
    {
        AddPendingHueWeight(history, fallbackHue, static_cast<double>(aggregateDeltaBytes));
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

[[nodiscard]] double SmoothRateForDisplay(double previousRate, double sampleRate, ULONGLONG elapsedMs) noexcept
{
    constexpr double kRateSmoothingTauMs = 1200.0;
    constexpr double kRateMaxAlpha       = 0.35;
    return SmoothFileOperationValue(previousRate, sampleRate, elapsedMs, kRateSmoothingTauMs, kRateMaxAlpha);
}

[[nodiscard]] double SmoothEtaSecondsForDisplay(double previousEtaSeconds, double sampleEtaSeconds, ULONGLONG elapsedMs) noexcept
{
    constexpr double kEtaSmoothingTauMs = 1500.0;
    constexpr double kEtaMaxAlpha       = 0.45;
    return SmoothFileOperationValue(previousEtaSeconds, sampleEtaSeconds, elapsedMs, kEtaSmoothingTauMs, kEtaMaxAlpha);
}

[[nodiscard]] double DecayRateForCallbackSilence(double smoothedRate, ULONGLONG silenceMs) noexcept
{
    constexpr ULONGLONG kRateSilenceHoldMs = 600ull;
    constexpr double kRateSilenceDecayMs   = 900.0;

    smoothedRate = ClampFiniteNonNegative(smoothedRate);
    if (smoothedRate <= 0.0 || silenceMs <= kRateSilenceHoldMs)
    {
        return smoothedRate;
    }

    const double decayMs     = static_cast<double>(silenceMs - kRateSilenceHoldMs);
    const double decayedRate = smoothedRate * std::exp(-decayMs / kRateSilenceDecayMs);
    return decayedRate >= 1.0 ? decayedRate : 0.0;
}

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

[[nodiscard]] double CurrentBandwidthForGraphMarker(const FileOperationsPopupInternal::RateHistory& history) noexcept
{
    const double displayed = ClampFiniteNonNegative(history.displayedBytesPerSec);
    if (displayed > 0.0)
    {
        return displayed;
    }

    if (history.count == 0)
    {
        return 0.0;
    }

    const size_t newestIndex =
        (history.writeIndex + FileOperationsPopupInternal::RateHistory::kMaxSamples - 1u) % FileOperationsPopupInternal::RateHistory::kMaxSamples;
    return ClampFiniteNonNegative(history.samples[newestIndex]);
}

void AppendRateSample(FileOperationsPopupInternal::RateHistory& history, double sample, float hue) noexcept
{
    const double clampedSample = ClampFiniteNonNegative(sample);
    const size_t slot          = history.writeIndex;

    history.samples[slot] = static_cast<float>(std::min<double>(clampedSample, std::numeric_limits<float>::max()));

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
                history.hueWeights[slot][0] = {.hue = hue, .weight = 1.0};
            }
            history.hues[slot] = hue;
        }
    }

    history.writeIndex = (history.writeIndex + 1u) % FileOperationsPopupInternal::RateHistory::kMaxSamples;
    history.count      = std::min(FileOperationsPopupInternal::RateHistory::kMaxSamples, history.count + 1u);
}

void ResetPendingRateSample(FileOperationsPopupInternal::RateHistory& history) noexcept
{
    history.pendingBucketMs         = 0;
    history.pendingWeightedSampleMs = 0.0;
    history.pendingHue              = -1.0f;
    history.pendingHueWeightCount   = 0u;
}

void AppendResampledRateSamples(FileOperationsPopupInternal::RateHistory& history, ULONGLONG elapsedMs, double sample, float hue) noexcept
{
    const ULONGLONG maxResampleMs = static_cast<ULONGLONG>(FileOperationsPopupInternal::RateHistory::kMaxSamples) * kRateSampleBucketMs;
    ULONGLONG remainingMs         = (std::min)(elapsedMs, maxResampleMs);
    while (remainingMs > 0)
    {
        const ULONGLONG bucketRemainingMs = kRateSampleBucketMs - history.pendingBucketMs;
        const ULONGLONG sliceMs           = std::min(remainingMs, bucketRemainingMs);

        history.pendingWeightedSampleMs += static_cast<double>(sample) * static_cast<double>(sliceMs);
        history.pendingBucketMs += sliceMs;
        if (hue >= 0.0f)
        {
            history.pendingHue = hue;
        }

        remainingMs -= sliceMs;
        if (history.pendingBucketMs < kRateSampleBucketMs)
        {
            continue;
        }

        const float bucketHue     = history.pendingHue >= 0.0f ? history.pendingHue : hue;
        const double bucketSample = history.pendingWeightedSampleMs / static_cast<double>(kRateSampleBucketMs);
        AppendRateSample(history, bucketSample, bucketHue);
        ResetPendingRateSample(history);
    }
}

std::wstring TruncatePathMiddleToWidth(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view path, float maxWidth, float height) noexcept
{
    const size_t prefixChars = ComputePathFixedPrefixChars(path);
    const size_t leafChars   = ComputePathLeafChars(path);

    return TruncateTextMiddleToWidth(factory, format, path, maxWidth, height, kEllipsisText, prefixChars, leafChars);
}

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
    for (const TaskSnapshot& task : snapshot)
    {
        if (task.finished && _collapsedTasks.find(task.taskId) == _collapsedTasks.end())
        {
            _collapsedTasks[task.taskId] = true;
            _compactExpandedTasks.erase(task.taskId);
            addedCollapse = true;
        }
    }

    if (addedCollapse)
    {
        _maxAutoSizedWindowHeight   = 0;
        _lastAutoSizedContentHeight = -1.0f;
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::CleanupCollapsedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept
{
    std::unordered_map<uint64_t, bool> seen;
    seen.reserve(snapshot.size());
    for (const TaskSnapshot& task : snapshot)
    {
        seen[task.taskId] = true;
    }

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
}

void FileOperationsPopupInternal::FileOperationsPopupState::DiscardDeviceResources() noexcept
{
    _target.reset();
    _captionGlyphTarget.reset();

    _bgBrush.reset();
    _textBrush.reset();
    _subTextBrush.reset();
    _borderBrush.reset();
    _progressBgBrush.reset();
    _progressGlobalBrush.reset();
    _progressItemBrush.reset();
    _checkboxFillBrush.reset();
    _checkboxCheckBrush.reset();
    _statusOkBrush.reset();
    _statusWarningBrush.reset();
    _statusErrorBrush.reset();
    _graphBgBrush.reset();
    _graphGridBrush.reset();
    _graphLimitBrush.reset();
    _graphLineBrush.reset();
    _graphFillBrush.reset();
    _graphDynamicBrush.reset();
    _graphTextShadowBrush.reset();
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

    if (_headerFormat && _bodyFormat && _smallFormat && _buttonFormat && _buttonSmallFormat && _graphOverlayFormat && _statusIconFallbackFormat)
    {
        return;
    }

    if (! _headerFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(),
            RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD),
            _headerFormat.put(),
            L""));
    }

    if (! _bodyFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi)), _bodyFormat.put(), L""));
    }

    if (! _smallFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(11.0f, _dpi)), _smallFormat.put(), L""));
    }

    if (! _buttonFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(12.0f, _dpi)), _buttonFormat.put(), L""));
    }

    if (! _buttonSmallFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(11.0f, _dpi)), _buttonSmallFormat.put(), L""));
    }

    if (! _graphOverlayFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(),
            RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(14.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD),
            _graphOverlayFormat.put(),
            L""));
    }

    if (! _statusIconFallbackFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(),
            RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(14.0f, _dpi), DWRITE_FONT_WEIGHT_SEMI_BOLD),
            _statusIconFallbackFormat.put(),
            L""));
    }

    if (! _statusIconFormat)
    {
        // Optional: Segoe Fluent Icons. If missing, fallback format draws standard Unicode glyphs.
        const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiIconSpec(DipsToPixels(14.0f, _dpi)), _statusIconFormat.put(), L"");
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
    configureButtonFormat(_buttonFormat.get());
    configureButtonFormat(_buttonSmallFormat.get());
    configureButtonFormat(_graphOverlayFormat.get());

    configureButtonFormat(_statusIconFormat.get());
    configureButtonFormat(_statusIconFallbackFormat.get());
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
        const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(20.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _captionGlyphFallbackFormat.put(), L"");
        if (FAILED(hr))
        {
            _captionGlyphFallbackFormat.reset();
        }
    }

    if (! _captionGlyphFormat)
    {
        const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiIconSpec(20.0f), _captionGlyphFormat.put(), L"");
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
    if (! _d2dFactory)
    {
        return;
    }

    if (_target)
    {
        return;
    }

    RECT rc{};
    GetClientRect(hwnd, &rc);
    _clientSize.cx = std::max(0L, rc.right - rc.left);
    _clientSize.cy = std::max(0L, rc.bottom - rc.top);

    _dpi = GetDpiForWindow(hwnd);

    const D2D1_SIZE_U size                             = D2D1::SizeU(static_cast<UINT32>(_clientSize.cx), static_cast<UINT32>(_clientSize.cy));
    D2D1_RENDER_TARGET_PROPERTIES props                = D2D1::RenderTargetProperties();
    props.dpiX                                         = 96.0f;
    props.dpiY                                         = 96.0f;
    const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, size);

    wil::com_ptr<ID2D1HwndRenderTarget> rt;
    const HRESULT hr = _d2dFactory->CreateHwndRenderTarget(props, hwndProps, rt.addressof());
    if (FAILED(hr) || ! rt)
    {
        _target.reset();
        return;
    }

    rt->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    _target = std::move(rt);
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
    _progressItemBaseColor            = progressItem;

    const D2D1::ColorF okAccent    = theme.fileOperations.successText;
    const D2D1::ColorF warningText = theme.folderView.warningText;
    const D2D1::ColorF errorText   = theme.folderView.errorText;

    const D2D1::ColorF graphBg    = theme.fileOperations.graphBackground;
    const D2D1::ColorF graphGrid  = theme.fileOperations.graphGrid;
    const D2D1::ColorF graphLimit = theme.fileOperations.graphLimit;
    const D2D1::ColorF graphLine  = theme.fileOperations.graphLine;

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

    if (! _graphBgBrush)
    {
        _target->CreateSolidColorBrush(graphBg, _graphBgBrush.addressof());
    }
    else
    {
        _graphBgBrush->SetColor(graphBg);
    }

    if (! _graphGridBrush)
    {
        _target->CreateSolidColorBrush(graphGrid, _graphGridBrush.addressof());
    }
    else
    {
        _graphGridBrush->SetColor(graphGrid);
    }

    if (! _graphLimitBrush)
    {
        _target->CreateSolidColorBrush(graphLimit, _graphLimitBrush.addressof());
    }
    else
    {
        _graphLimitBrush->SetColor(graphLimit);
    }

    if (! _graphLineBrush)
    {
        _target->CreateSolidColorBrush(graphLine, _graphLineBrush.addressof());
    }
    else
    {
        _graphLineBrush->SetColor(graphLine);
    }

    const float graphFillAlpha   = theme.dark ? 0.22f : 0.18f;
    const D2D1::ColorF graphFill = D2D1::ColorF(graphLine.r, graphLine.g, graphLine.b, graphFillAlpha);
    _graphFillBaseColor          = graphFill;

    if (! _graphFillBrush)
    {
        _target->CreateSolidColorBrush(graphFill, _graphFillBrush.addressof());
    }
    else
    {
        _graphFillBrush->SetColor(graphFill);
    }

    if (! _graphDynamicBrush)
    {
        _target->CreateSolidColorBrush(graphFill, _graphDynamicBrush.addressof());
    }

    // Shadow brush for overlay text - lighter on light themes for subtlety
    const float shadowAlpha        = theme.dark ? 0.6f : 0.25f;
    const D2D1::ColorF shadowColor = D2D1::ColorF(0.0f, 0.0f, 0.0f, shadowAlpha);
    if (! _graphTextShadowBrush)
    {
        _target->CreateSolidColorBrush(shadowColor, _graphTextShadowBrush.addressof());
    }
    else
    {
        _graphTextShadowBrush->SetColor(shadowColor);
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
    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    std::vector<FolderWindow::InformationalTaskUpdate> informationalTasks;
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completedTasks;
    if (fileOps && hostLifetime.lock())
    {
        fileOps->CollectTasks(tasks);
        fileOps->CollectInformationalTasks(informationalTasks);
        fileOps->CollectCompletedTasks(completedTasks);
    }

    std::vector<TaskSnapshot> result;
    result.reserve(tasks.size() + completedTasks.size() + informationalTasks.size());
    std::unordered_map<uint64_t, bool> activeTaskIds;
    activeTaskIds.reserve(tasks.size());
    std::unordered_map<uint64_t, const FolderWindow::FileOperationState::CompletedTaskSummary*> completedTaskById;
    completedTaskById.reserve(completedTasks.size());
    for (const auto& completed : completedTasks)
    {
        completedTaskById.emplace(completed.taskId, &completed);
    }

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
        if (! task)
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.taskId                = task->GetId();
        activeTaskIds[snap.taskId] = true;
        snap.operation             = task->GetOperation();

        snap.totalItems         = task->_publishedProgressTotalItems.load(std::memory_order_acquire);
        snap.completedItems     = task->_publishedProgressCompletedItems.load(std::memory_order_acquire);
        snap.totalBytes         = task->_publishedProgressTotalBytes.load(std::memory_order_acquire);
        snap.completedBytes     = task->_publishedProgressCompletedBytes.load(std::memory_order_acquire);
        snap.itemTotalBytes     = task->_publishedProgressItemTotalBytes.load(std::memory_order_acquire);
        snap.itemCompletedBytes = task->_publishedProgressItemCompletedBytes.load(std::memory_order_acquire);
        snap.completedFiles     = task->_publishedCompletedTopLevelFiles.load(std::memory_order_acquire);
        snap.completedFolders   = task->_publishedCompletedTopLevelFolders.load(std::memory_order_acquire);

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
            snap.inFlightFileCount = std::min(task->_inFlightFileCount, snap.inFlightFiles.size());
            for (size_t i = 0; i < snap.inFlightFileCount; ++i)
            {
                snap.inFlightFiles[i].sourcePath     = task->_inFlightFiles[i].sourcePath;
                snap.inFlightFiles[i].totalBytes     = task->_inFlightFiles[i].totalBytes;
                snap.inFlightFiles[i].completedBytes = task->_inFlightFiles[i].completedBytes;

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
            snap.conflict.bucket                            = static_cast<uint8_t>(task->_conflictArbiter.prompt.bucket);
            snap.conflict.status                            = task->_conflictArbiter.prompt.status;
            snap.conflict.sourcePath                        = task->_conflictArbiter.prompt.sourcePath;
            snap.conflict.destinationPath                   = task->_conflictArbiter.prompt.destinationPath;
            snap.conflict.sourceMetadata.available          = task->_conflictArbiter.prompt.sourceMetadata.available;
            snap.conflict.sourceMetadata.isDirectory        = task->_conflictArbiter.prompt.sourceMetadata.isDirectory;
            snap.conflict.sourceMetadata.sizeKnown          = task->_conflictArbiter.prompt.sourceMetadata.sizeKnown;
            snap.conflict.sourceMetadata.sizeBytes          = task->_conflictArbiter.prompt.sourceMetadata.sizeBytes;
            snap.conflict.sourceMetadata.lastWriteTime      = task->_conflictArbiter.prompt.sourceMetadata.lastWriteTime;
            snap.conflict.sourceMetadata.attributes         = task->_conflictArbiter.prompt.sourceMetadata.attributes;
            snap.conflict.destinationMetadata.available     = task->_conflictArbiter.prompt.destinationMetadata.available;
            snap.conflict.destinationMetadata.isDirectory   = task->_conflictArbiter.prompt.destinationMetadata.isDirectory;
            snap.conflict.destinationMetadata.sizeKnown     = task->_conflictArbiter.prompt.destinationMetadata.sizeKnown;
            snap.conflict.destinationMetadata.sizeBytes     = task->_conflictArbiter.prompt.destinationMetadata.sizeBytes;
            snap.conflict.destinationMetadata.lastWriteTime = task->_conflictArbiter.prompt.destinationMetadata.lastWriteTime;
            snap.conflict.destinationMetadata.attributes    = task->_conflictArbiter.prompt.destinationMetadata.attributes;
            snap.conflict.applyToAllChecked                 = task->_conflictArbiter.prompt.applyToAllChecked;
            snap.conflict.retryFailed                       = task->_conflictArbiter.prompt.retryFailed;

            snap.conflict.actionCount = std::min(task->_conflictArbiter.prompt.actionCount, snap.conflict.actions.size());
            for (size_t i = 0; i < snap.conflict.actionCount; ++i)
            {
                snap.conflict.actions[i] = static_cast<uint8_t>(task->_conflictArbiter.prompt.actions[i]);
            }
        }

        snap.started                    = task->HasStarted();
        snap.paused                     = task->IsPaused();
        snap.waitingForOthers           = task->IsWaitingForOthers();
        snap.waitingInQueue             = task->IsWaitingInQueue();
        snap.queuePaused                = task->IsQueuePaused();
        snap.plannedItems               = task->GetPlannedItemCount();
        snap.destinationFolder          = task->GetDestinationFolder();
        snap.destinationPane            = task->GetDestinationPane();
        snap.destinationPluginId        = task->_destinationPluginId;
        snap.destinationPluginShortId   = task->_destinationPluginShortId;
        snap.destinationInstanceContext = task->_destinationInstanceContext;
        snap.operationStartTick         = task->_operationStartTick.load(std::memory_order_acquire);

        // A task whose thread already completed renders its final status immediately instead of
        // flashing "Running" until the completed-summary row replaces this live row.
        if (task->_taskFinished.load(std::memory_order_acquire))
        {
            snap.finished          = true;
            snap.resultHr          = task->_resultHr.load(std::memory_order_acquire);
            const auto completedIt = completedTaskById.find(snap.taskId);
            if (completedIt != completedTaskById.end() && completedIt->second)
            {
                const auto& completed      = *completedIt->second;
                snap.warningCount          = completed.warningCount;
                snap.errorCount            = completed.errorCount;
                snap.lastDiagnosticMessage = completed.lastDiagnosticMessage;
            }
            else
            {
                fileOps->CollectTaskDiagnosticSnapshot(snap.taskId, snap.warningCount, snap.errorCount, snap.lastDiagnosticMessage);
            }
        }

        snap.desiredSpeedLimitBytesPerSecond       = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        snap.effectiveSpeedLimitBytesPerSecond     = task->_effectiveSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
        snap.autoConcurrencyUsed                   = task->_autoConcurrencyUsed.load(std::memory_order_acquire);
        snap.autoConcurrencyStorageKind            = task->_autoConcurrencyStorageKind.load(std::memory_order_acquire);
        snap.autoConcurrencyDestinationStorageKind = task->_autoConcurrencyDestinationStorageKind.load(std::memory_order_acquire);
        snap.autoTunedConcurrency                  = task->_autoTunedConcurrency.load(std::memory_order_acquire);
        snap.effectiveConcurrencyBudget            = task->_effectiveConcurrencyBudget.load(std::memory_order_acquire);

        // Pre-calculation state
        snap.preCalcInProgress              = task->_preCalcInProgress.load(std::memory_order_acquire);
        snap.preCalcSkipped                 = task->_preCalcSkipped.load(std::memory_order_acquire);
        snap.preCalcCompleted               = task->_preCalcCompleted.load(std::memory_order_acquire);
        snap.earlyAdmissionTransferObserved = task->_transferStartedBeforePreCalcComplete.load(std::memory_order_acquire);
        snap.preCalcTotalBytes              = task->_preCalcTotalBytes.load(std::memory_order_acquire);
        snap.preCalcFileCount               = task->_preCalcFileCount.load(std::memory_order_acquire);
        snap.preCalcDirectoryCount          = task->_preCalcDirectoryCount.load(std::memory_order_acquire);

        const ULONGLONG startTick = task->_preCalcStartTick.load(std::memory_order_acquire);
        if (snap.preCalcInProgress && startTick > 0)
        {
            const ULONGLONG nowTick = GetTickCount64();
            snap.preCalcElapsedMs   = (nowTick >= startTick) ? (nowTick - startTick) : 0;
        }

        PublishPlannedItemTotalAfterPreCalculation(snap);

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
        EnsureFinishedTaskDiagnosticAffordance(snap);
        NormalizeCompletedTaskSnapshotForDisplay(snap);
        snap.statusKind = ResolveTaskStatusKind(snap);

        result.push_back(std::move(snap));
    }

    for (const auto& completed : completedTasks)
    {
        if (activeTaskIds.find(completed.taskId) != activeTaskIds.end())
        {
            continue;
        }

        TaskSnapshot snap{};
        snap.taskId                                = completed.taskId;
        snap.operation                             = completed.operation;
        snap.totalItems                            = completed.totalItems;
        snap.completedItems                        = completed.completedItems;
        snap.totalBytes                            = completed.totalBytes;
        snap.completedBytes                        = completed.completedBytes;
        snap.completedFiles                        = completed.completedFiles;
        snap.completedFolders                      = completed.completedFolders;
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
        snap.preCalcSkipped                        = completed.preCalcSkipped;
        snap.hasProgressCallbacks                  = completed.lastProgressCallbackTick != 0;
        snap.lastProgressCallbackTick              = completed.lastProgressCallbackTick;

        if (snap.totalItems > 0)
        {
            snap.completedItems = std::min(snap.completedItems, snap.totalItems);
        }
        if (snap.totalBytes > 0)
        {
            snap.completedBytes = std::min(snap.completedBytes, snap.totalBytes);
        }
        EnsureFinishedTaskDiagnosticAffordance(snap);
        NormalizeCompletedTaskSnapshotForDisplay(snap);
        snap.statusKind = ResolveTaskStatusKind(snap);

        result.push_back(std::move(snap));
    }

    return result;
}

std::vector<RateSnapshot> FileOperationsPopupInternal::FileOperationsPopupState::BuildRateSnapshot() const
{
    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completedTasks;
    if (fileOps && hostLifetime.lock())
    {
        fileOps->CollectTasks(tasks);
        fileOps->CollectCompletedTasks(completedTasks);
    }

    std::unordered_map<uint64_t, bool> activeTaskIds;
    activeTaskIds.reserve(tasks.size());

    std::vector<RateSnapshot> result;
    result.reserve(tasks.size() + completedTasks.size());

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        RateSnapshot snap{};
        snap.taskId                = task->GetId();
        snap.operation             = task->GetOperation();
        activeTaskIds[snap.taskId] = true;

        snap.completedItems = task->_publishedProgressCompletedItems.load(std::memory_order_acquire);
        snap.totalBytes     = task->_publishedProgressTotalBytes.load(std::memory_order_acquire);
        snap.completedBytes = task->_publishedProgressCompletedBytes.load(std::memory_order_acquire);

        {
            std::scoped_lock lock(task->_progressPathMutex);
            snap.currentSourcePath        = ! task->_lastProgressCallbackSourcePath.empty() ? task->_lastProgressCallbackSourcePath : task->_progressSourcePath;
            snap.lastProgressCallbackTick = task->_lastProgressCallbackTick;
        }

        {
            std::scoped_lock lock(task->_inFlightFilesMutex);
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

        result.push_back(snap);
    }

    for (const auto& completed : completedTasks)
    {
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
        snap.finished                 = true;

        result.push_back(std::move(snap));
    }

    return result;
}

void FileOperationsPopupInternal::FileOperationsPopupState::UpdateRates() noexcept
{
    const bool capturePerf                   = Debug::Perf::IsCaptureEnabled();
    const uint64_t perfStartUs               = capturePerf ? PerfNowUs() : 0u;
    const ULONGLONG nowTick                  = GetTickCount64();
    const std::vector<RateSnapshot> snapshot = BuildRateSnapshot();

    std::unordered_map<uint64_t, bool> seen;
    seen.reserve(snapshot.size());

    uint64_t maxCallbackSilenceMs = 0;
    uint64_t maxDisplayGapMs      = 0;
    uint64_t silentTaskCount      = 0;
    uint64_t updatedTaskCount     = 0;

    for (const RateSnapshot& task : snapshot)
    {
        seen[task.taskId] = true;

        RateHistory& history = _rates[task.taskId];
        const bool blocked   = IsRateSamplingBlocked(task);
        const bool itemRate  = task.operation == FILESYSTEM_DELETE;

        if (! history.initialized)
        {
            history.initialized              = true;
            history.lastBytes                = task.completedBytes;
            history.lastItems                = task.completedItems;
            history.lastProgressCallbackTick = task.lastProgressCallbackTick;
            history.lastDisplaySampleTick    = nowTick;
            history.lastStateChangeTick      = task.progressStateChangeTick;
            SyncRateStreamBaselines(history, task);
            continue;
        }

        history.lastBytes = std::min(history.lastBytes, task.completedBytes);
        history.lastItems = std::min(history.lastItems, task.completedItems);

        if (task.progressStateChangeTick > history.lastStateChangeTick)
        {
            history.lastStateChangeTick   = task.progressStateChangeTick;
            history.resumeTick            = blocked ? 0 : task.progressStateChangeTick;
            history.lastDisplaySampleTick = nowTick;
            ResetPendingRateSample(history);
            ResetRateStreamProgress(history);
            SyncRateStreamBaselines(history, task);

            if (blocked)
            {
                history.lastBytes                = task.completedBytes;
                history.lastItems                = task.completedItems;
                history.lastProgressCallbackTick = task.lastProgressCallbackTick;
                history.hasSmoothedEta           = false;
                continue;
            }
        }

        if (blocked)
        {
            history.resumeTick               = 0;
            history.lastBytes                = task.completedBytes;
            history.lastItems                = task.completedItems;
            history.lastProgressCallbackTick = task.lastProgressCallbackTick;
            history.lastDisplaySampleTick    = nowTick;
            history.hasSmoothedEta           = false;
            ResetPendingRateSample(history);
            ResetRateStreamProgress(history);
            SyncRateStreamBaselines(history, task);
            continue;
        }

        const bool hasNewProgressInput = task.lastProgressCallbackTick != 0 && task.lastProgressCallbackTick > history.lastProgressCallbackTick;
        const float hue                = RateSampleHue(task.currentSourcePath);
        ULONGLONG smoothingElapsedMs   = kFileOperationsPopupTimerIntervalMs;

        if (hasNewProgressInput)
        {
            ULONGLONG baselineTick = history.lastProgressCallbackTick;
            if (baselineTick == 0 || baselineTick > task.lastProgressCallbackTick)
            {
                baselineTick = task.lastProgressCallbackTick;
            }
            if (history.resumeTick > baselineTick && history.resumeTick <= task.lastProgressCallbackTick)
            {
                baselineTick = history.resumeTick;
            }

            const ULONGLONG elapsedMs = std::max<ULONGLONG>(1ull, task.lastProgressCallbackTick - baselineTick);
            const double dtSec        = static_cast<double>(elapsedMs) / 1000.0;
            smoothingElapsedMs        = elapsedMs;

            if (itemRate)
            {
                const unsigned long deltaItems = task.completedItems - history.lastItems;
                if (deltaItems > 0 && dtSec > 0.0)
                {
                    const double instItemsPerSec = static_cast<double>(deltaItems) / dtSec;
                    history.smoothedItemsPerSec  = SmoothRateForDisplay(history.smoothedItemsPerSec, instItemsPerSec, elapsedMs);
                    history.displayedItemsPerSec = history.smoothedItemsPerSec;
                    ++updatedTaskCount;
                }

                history.lastItems = task.completedItems;
            }
            else
            {
                const uint64_t deltaBytes = task.completedBytes - history.lastBytes;
                if (deltaBytes > 0 && dtSec > 0.0)
                {
                    const double instBytesPerSec = static_cast<double>(deltaBytes) / dtSec;
                    history.smoothedBytesPerSec  = SmoothRateForDisplay(history.smoothedBytesPerSec, instBytesPerSec, elapsedMs);
                    history.displayedBytesPerSec = history.smoothedBytesPerSec;
                    ++updatedTaskCount;
                }

                history.lastBytes = task.completedBytes;
            }

            history.lastProgressCallbackTick = task.lastProgressCallbackTick;
            history.resumeTick               = 0;
        }
        else if (! task.finished && task.lastProgressCallbackTick != 0 && nowTick >= task.lastProgressCallbackTick)
        {
            const ULONGLONG silenceMs = nowTick - task.lastProgressCallbackTick;
            maxCallbackSilenceMs      = std::max<uint64_t>(maxCallbackSilenceMs, silenceMs);
            if (silenceMs > 0)
            {
                ++silentTaskCount;
            }

            if (itemRate)
            {
                history.displayedItemsPerSec = DecayRateForCallbackSilence(history.smoothedItemsPerSec, silenceMs);
            }
            else
            {
                history.displayedBytesPerSec = DecayRateForCallbackSilence(history.smoothedBytesPerSec, silenceMs);
            }
        }

        if (! itemRate && task.totalBytes > 0 && task.completedBytes <= task.totalBytes && IsByteRateUsableForEta(history.displayedBytesPerSec))
        {
            const uint64_t remainingBytes = task.totalBytes - task.completedBytes;
            const double rawEtaSeconds    = static_cast<double>(remainingBytes) / history.displayedBytesPerSec;
            history.smoothedEtaSeconds =
                history.hasSmoothedEta ? SmoothEtaSecondsForDisplay(history.smoothedEtaSeconds, rawEtaSeconds, smoothingElapsedMs) : rawEtaSeconds;
            history.hasSmoothedEta = true;
        }
        else if (! itemRate)
        {
            history.hasSmoothedEta = false;
        }

        if (! task.finished)
        {
            // Per-stream hue attribution samples the in-flight cumulative shares every tick,
            // independent of how often the published aggregate or the plugin callbacks advance.
            if (! itemRate)
            {
                AccumulateStreamHueWeights(history, task, 0u, hue);
            }

            if (history.lastDisplaySampleTick == 0 || history.lastDisplaySampleTick > nowTick)
            {
                history.lastDisplaySampleTick = nowTick;
            }
            else
            {
                const ULONGLONG displayElapsedMs = nowTick - history.lastDisplaySampleTick;
                if (displayElapsedMs > 0)
                {
                    maxDisplayGapMs            = std::max<uint64_t>(maxDisplayGapMs, displayElapsedMs);
                    const double displaySample = itemRate ? history.displayedItemsPerSec : history.displayedBytesPerSec;
                    if (displaySample > 0.0 || history.count > 0 || history.pendingHueWeightCount > 0u)
                    {
                        AppendResampledRateSamples(history, displayElapsedMs, displaySample, hue);
                    }
                    history.lastDisplaySampleTick = nowTick;
                }
            }
        }
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
        const std::wstring detail = std::format(L"tasks={} updated={} silent={} rates={}", snapshot.size(), updatedTaskCount, silentTaskCount, _rates.size());
        Debug::Perf::Emit(L"FileOps.Popup.Rate.UpdateUs", detail, PerfElapsedUs(perfStartUs), snapshot.size(), _rates.size(), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Rate.MaxCallbackSilenceMs", detail, maxCallbackSilenceMs, silentTaskCount, snapshot.size(), S_OK);
        Debug::Perf::Emit(L"FileOps.Popup.Rate.MaxDisplayGapMs", detail, maxDisplayGapMs, snapshot.size(), _rates.size(), S_OK);
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::LayoutChrome(float width, float height, bool showPauseResumeAll) noexcept
{
    const float footerH = FileOperationsPopupFooterHeightPixels(_dpi);

    const float footerTop = std::max(0.0f, height - footerH);
    _listViewportRect     = D2D1::RectF(0.0f, 0.0f, width, footerTop);

    const float footerMargin     = DipsToPixels(10.0f, _dpi);
    const float footerGap        = DipsToPixels(8.0f, _dpi);
    const float progressH        = DipsToPixels(6.0f, _dpi);
    const float progressY        = footerTop + DipsToPixels(8.0f, _dpi);
    _footerAggregateProgressRect = D2D1::RectF(footerMargin, progressY, std::max(footerMargin, width - footerMargin), std::min(height, progressY + progressH));

    const float summaryY = footerTop + DipsToPixels(18.0f, _dpi);
    const float summaryH = DipsToPixels(20.0f, _dpi);
    _footerSummaryRect   = D2D1::RectF(footerMargin, summaryY, std::max(footerMargin, width - footerMargin), summaryY + summaryH);

    const float footerBtnH = DipsToPixels(28.0f, _dpi);
    const float footerBtnY = footerTop + DipsToPixels(48.0f, _dpi);
    const float rightEdge  = std::max(footerMargin, width - footerMargin);

    const float cancelW      = width >= DipsToPixels(560.0f, _dpi) ? DipsToPixels(112.0f, _dpi) : DipsToPixels(72.0f, _dpi);
    const float pauseResumeW = showPauseResumeAll ? cancelW : 0.0f;
    const float detailsW     = footerBtnH;
    const float densityW     = width >= DipsToPixels(500.0f, _dpi) ? DipsToPixels(84.0f, _dpi) : footerBtnH;
    const float minQueueW    = DipsToPixels(144.0f, _dpi);
    const float idealQueueW  = DipsToPixels(324.0f, _dpi);
    const float idealAutoW   = width >= DipsToPixels(680.0f, _dpi)
                                   ? DipsToPixels(210.0f, _dpi)
                                   : (width >= DipsToPixels(360.0f, _dpi) ? DipsToPixels(150.0f, _dpi) : DipsToPixels(0.0f, _dpi));
    const float minAutoW     = DipsToPixels(48.0f, _dpi);

    float cursor              = footerMargin;
    _footerCancelAllRect      = D2D1::RectF(cursor, footerBtnY, std::min(rightEdge, cursor + cancelW), footerBtnY + footerBtnH);
    cursor                    = _footerCancelAllRect.right + footerGap;
    _footerPauseResumeAllRect = pauseResumeW > 0.0f ? D2D1::RectF(cursor, footerBtnY, std::min(rightEdge, cursor + pauseResumeW), footerBtnY + footerBtnH)
                                                    : D2D1::RectF(cursor, footerBtnY, cursor, footerBtnY + footerBtnH);
    cursor                    = pauseResumeW > 0.0f ? (_footerPauseResumeAllRect.right + footerGap) : cursor;

    const float detailsLeft  = std::max(footerMargin, rightEdge - detailsW);
    _footerDetailsToggleRect = D2D1::RectF(detailsLeft, footerBtnY, rightEdge, footerBtnY + footerBtnH);

    const float densityLeft = std::max(cursor, detailsLeft - footerGap - densityW);
    _footerDensityRect      = D2D1::RectF(densityLeft, footerBtnY, std::min(detailsLeft - footerGap, densityLeft + densityW), footerBtnY + footerBtnH);

    float rightCursor = std::max(cursor, _footerDensityRect.left - footerGap);
    float available   = std::max(0.0f, rightCursor - cursor);

    float autoW = 0.0f;
    if (idealAutoW > 0.0f && available > minQueueW + footerGap + minAutoW)
    {
        autoW = std::min(idealAutoW, available - minQueueW - footerGap);
    }

    _footerAutoDismissRect = autoW > 0.0f ? D2D1::RectF(cursor, footerBtnY, cursor + autoW, footerBtnY + footerBtnH)
                                          : D2D1::RectF(cursor, footerBtnY, cursor, footerBtnY + footerBtnH);
    cursor                 = autoW > 0.0f ? (_footerAutoDismissRect.right + footerGap) : cursor;
    available              = std::max(0.0f, rightCursor - cursor);

    const float queueW   = std::min(idealQueueW, available);
    _footerQueueModeRect = queueW >= minQueueW ? D2D1::RectF(cursor, footerBtnY, cursor + queueW, footerBtnY + footerBtnH)
                                               : D2D1::RectF(cursor, footerBtnY, cursor, footerBtnY + footerBtnH);
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

    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
}

void FileOperationsPopupInternal::FileOperationsPopupState::AutoResizeWindow(
    HWND hwnd, float desiredContentHeight, size_t taskCount, bool footerOnly, bool reducedMotion) noexcept
{
    if (! hwnd || _inSizeMove)
    {
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

    if (! taskCountChanged && ! contentHeightChanged && ! _autoResizePending && ! _autoResizeAnimating && ! _footerOnlyRestorePending)
    {
        return;
    }

    if (taskCountChanged || contentHeightChanged)
    {
        _lastTaskCount              = taskCount;
        _lastAutoSizedContentHeight = desiredContentHeight;
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
    const float footerH             = FileOperationsPopupFooterHeightPixels(_dpi);
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
    const int minClientH         = DipsToPixels(minClientHeightDip, _dpi);
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

void FileOperationsPopupInternal::FileOperationsPopupState::DrawDxUiButtonChrome(const PopupButton& button,
                                                                                 IDWriteTextFormat* format,
                                                                                 std::wstring_view text,
                                                                                 RedSalamander::DxUi::ButtonVariant variant) noexcept
{
    if (! _target || ! _buttonChromeBrush || ! folderWindow || ! hostLifetime.lock())
    {
        return;
    }

    RedSalamander::DxUi::ButtonChromeDrawSpec spec{};
    spec.bounds       = button.bounds;
    spec.text         = text;
    spec.variant      = variant;
    spec.hovered      = button.hit.kind == _hotHit.kind && button.hit.taskId == _hotHit.taskId && button.hit.data == _hotHit.data;
    spec.pressed      = button.hit.kind == _pressedHit.kind && button.hit.taskId == _pressedHit.taskId && button.hit.data == _pressedHit.data;
    spec.scale        = static_cast<float>(_dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    spec.chevronGlyph = FluentIcons::kChevronDown;

    IDWriteTextFormat* iconFormat = _statusIconFormat.get();
    if (! iconFormat || ! DirectWriteFormatHasGlyph(_dwriteFactory.get(), iconFormat, spec.chevronGlyph))
    {
        spec.chevronGlyph = FluentIcons::kFallbackChevronDown;
        iconFormat        = _statusIconFallbackFormat ? _statusIconFallbackFormat.get() : format;
    }

    const AppTheme& appTheme = folderWindow->GetTheme();
    RedSalamander::DxUi::DrawButtonChrome(
        _target.get(), _buttonChromeBrush.get(), format, iconFormat, MakeAppThemeDxPalette(appTheme, appTheme.windowBackground), spec);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawButton(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text) noexcept
{
    DrawDxUiButtonChrome(button, format, text, RedSalamander::DxUi::ButtonVariant::Standard);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawFooterQueueModeControl(const PopupButton& button, bool queueMode, bool reducedMotion) noexcept
{
    _footerQueueSegmentRect    = {};
    _footerParallelSegmentRect = {};
    if (! _target || ! _buttonBgBrush || ! _borderBrush || ! _smallFormat || ! _buttonSmallFormat || button.bounds.right <= button.bounds.left ||
        button.bounds.bottom <= button.bounds.top)
    {
        return;
    }

    const bool hovered = _hotHit.kind == PopupHitTest::Kind::FooterQueueMode;
    const bool pressed = _pressedHit.kind == PopupHitTest::Kind::FooterQueueMode;

    const float radius = ClampCornerRadius(button.bounds, DipsToPixels(4.0f, _dpi));
    _target->FillRoundedRectangle(D2D1::RoundedRect(button.bounds, radius, radius), _buttonBgBrush.get());
    _target->DrawRoundedRectangle(D2D1::RoundedRect(button.bounds, radius, radius), _borderBrush.get(), pressed ? 1.6f : (hovered ? 1.3f : 1.0f));

    const float padX  = DipsToPixels(8.0f, _dpi);
    D2D1_RECT_F inner = D2D1::RectF(
        button.bounds.left + padX, button.bounds.top + DipsToPixels(4.0f, _dpi), button.bounds.right - padX, button.bounds.bottom - DipsToPixels(4.0f, _dpi));
    if (inner.right <= inner.left || inner.bottom <= inner.top)
    {
        return;
    }

    const std::wstring_view labelText    = _footerNewTasksText;
    const std::wstring_view queueText    = _footerQueueText;
    const std::wstring_view parallelText = _footerParallelText;
    const std::wstring_view selectedText = queueMode ? queueText : parallelText;

    const float innerW = inner.right - inner.left;
    if (innerW < DipsToPixels(112.0f, _dpi) || ! _dwriteFactory)
    {
        _target->DrawTextW(
            selectedText.data(), static_cast<UINT32>(selectedText.size()), _buttonSmallFormat.get(), inner, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
        return;
    }

    float segmentLeft = inner.left;
    if (innerW >= DipsToPixels(184.0f, _dpi) && ! labelText.empty())
    {
        const float maxLabelW = std::min(DipsToPixels(78.0f, _dpi), innerW * 0.38f);
        const float labelW    = std::min(maxLabelW, MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), labelText, maxLabelW, inner.bottom - inner.top));
        const D2D1_RECT_F labelRc = D2D1::RectF(inner.left, inner.top, inner.left + labelW, inner.bottom);
        if (_subTextBrush)
        {
            _target->DrawTextW(
                labelText.data(), static_cast<UINT32>(labelText.size()), _smallFormat.get(), labelRc, _subTextBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }
        segmentLeft = labelRc.right + DipsToPixels(8.0f, _dpi);
    }

    const D2D1_RECT_F segmentRc = D2D1::RectF(segmentLeft, inner.top, inner.right, inner.bottom);
    if (segmentRc.right <= segmentRc.left)
    {
        return;
    }

    const float midX           = (segmentRc.left + segmentRc.right) * 0.5f;
    const D2D1_RECT_F queueRc  = D2D1::RectF(segmentRc.left, segmentRc.top, midX, segmentRc.bottom);
    const D2D1_RECT_F paraRc   = D2D1::RectF(midX, segmentRc.top, segmentRc.right, segmentRc.bottom);
    _footerQueueSegmentRect    = queueRc;
    _footerParallelSegmentRect = paraRc;
    ID2D1Brush* selectedBrush  = _textBrush ? _textBrush.get() : (_subTextBrush ? _subTextBrush.get() : nullptr);
    ID2D1Brush* secondaryBrush = _subTextBrush ? _subTextBrush.get() : (_textBrush ? _textBrush.get() : nullptr);

    const float targetPosition = queueMode ? 0.0f : 1.0f;
    const ULONGLONG nowTick    = GetTickCount64();
    if (! _footerQueueModeAnimationInitialized || reducedMotion)
    {
        _footerQueueModeAnimationPosition    = targetPosition;
        _footerQueueModeAnimationLastTick    = nowTick;
        _footerQueueModeAnimationInitialized = true;
    }
    else
    {
        const ULONGLONG elapsed           = nowTick >= _footerQueueModeAnimationLastTick ? (nowTick - _footerQueueModeAnimationLastTick) : 0ull;
        _footerQueueModeAnimationLastTick = nowTick;

        const float alpha = std::max(0.10f, EaseFileOperationsUiMotionFraction(std::min<ULONGLONG>(elapsed, 180ull), 180ull));
        _footerQueueModeAnimationPosition += (targetPosition - _footerQueueModeAnimationPosition) * alpha;
        if (std::abs(_footerQueueModeAnimationPosition - targetPosition) < 0.01f)
        {
            _footerQueueModeAnimationPosition = targetPosition;
        }
    }

    const float segmentW      = segmentRc.right - segmentRc.left;
    const float thumbW        = segmentW * 0.5f;
    const D2D1_RECT_F thumbRc = D2D1::RectF(segmentRc.left + thumbW * _footerQueueModeAnimationPosition,
                                            segmentRc.top,
                                            segmentRc.left + thumbW * (_footerQueueModeAnimationPosition + 1.0f),
                                            segmentRc.bottom);

    if (_graphDynamicBrush && folderWindow)
    {
        D2D1_COLOR_F thumbColor = folderWindow->GetTheme().accent;
        thumbColor.a            = folderWindow->GetTheme().highContrast ? 1.0f : (folderWindow->GetTheme().dark ? 0.34f : 0.22f);
        _graphDynamicBrush->SetColor(thumbColor);
        const float thumbRadius = ClampCornerRadius(thumbRc, DipsToPixels(4.0f, _dpi));
        _target->FillRoundedRectangle(D2D1::RoundedRect(thumbRc, thumbRadius, thumbRadius), _graphDynamicBrush.get());
    }
    else if (_buttonBgBrush)
    {
        const float thumbRadius = ClampCornerRadius(thumbRc, DipsToPixels(4.0f, _dpi));
        _target->FillRoundedRectangle(D2D1::RoundedRect(thumbRc, thumbRadius, thumbRadius), _buttonBgBrush.get());
    }

    const float segmentRadius = ClampCornerRadius(segmentRc, DipsToPixels(4.0f, _dpi));
    _target->DrawRoundedRectangle(D2D1::RoundedRect(segmentRc, segmentRadius, segmentRadius), _borderBrush.get(), 1.0f);
    _target->DrawLine(D2D1::Point2F(midX, segmentRc.top), D2D1::Point2F(midX, segmentRc.bottom), _borderBrush.get(), 1.0f);

    if (selectedBrush)
    {
        _target->DrawTextW(queueText.data(),
                           static_cast<UINT32>(queueText.size()),
                           _buttonSmallFormat.get(),
                           queueRc,
                           queueMode ? selectedBrush : secondaryBrush,
                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
        _target->DrawTextW(parallelText.data(),
                           static_cast<UINT32>(parallelText.size()),
                           _buttonSmallFormat.get(),
                           paraRc,
                           queueMode ? secondaryBrush : selectedBrush,
                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawFooterAutoDismissControl(const PopupButton& button, bool enabled) noexcept
{
    _footerAutoDismissLabelVisible = false;
    if (! _target || ! _buttonBgBrush || ! _borderBrush || ! _smallFormat || button.bounds.right <= button.bounds.left ||
        button.bounds.bottom <= button.bounds.top)
    {
        return;
    }

    const bool hovered = button.hit.kind == _hotHit.kind && button.hit.taskId == _hotHit.taskId && button.hit.data == _hotHit.data;
    const bool pressed = button.hit.kind == _pressedHit.kind && button.hit.taskId == _pressedHit.taskId && button.hit.data == _pressedHit.data;

    const float radius = ClampCornerRadius(button.bounds, DipsToPixels(4.0f, _dpi));
    _target->FillRoundedRectangle(D2D1::RoundedRect(button.bounds, radius, radius), _buttonBgBrush.get());
    _target->DrawRoundedRectangle(D2D1::RoundedRect(button.bounds, radius, radius), _borderBrush.get(), pressed ? 1.6f : (hovered ? 1.3f : 1.0f));

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

bool FileOperationsPopupInternal::FileOperationsPopupState::DrawCenteredChevronGlyph(const D2D1_RECT_F& rc, wchar_t fluentGlyph, wchar_t fallbackGlyph) noexcept
{
    if (! _target || ! _textBrush || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return false;
    }

    const bool useFluentGlyph = _statusIconFormat != nullptr && DirectWriteFormatHasGlyph(_dwriteFactory.get(), _statusIconFormat.get(), fluentGlyph);
    const wchar_t glyph       = useFluentGlyph ? fluentGlyph : fallbackGlyph;
    IDWriteTextFormat* format =
        useFluentGlyph ? _statusIconFormat.get() : (_statusIconFallbackFormat ? _statusIconFallbackFormat.get() : _buttonSmallFormat.get());
    if (! format || glyph == 0)
    {
        return false;
    }

    const wchar_t text[2]{glyph, 0};
    _target->DrawTextW(text, 1u, format, rc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawMenuButton(const PopupButton& button,
                                                                           IDWriteTextFormat* format,
                                                                           std::wstring_view text) noexcept
{
    DrawDxUiButtonChrome(button, format, text, RedSalamander::DxUi::ButtonVariant::DropDown);
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

    const D2D1_POINT_2F p1{left + size * 0.20f, top + size * 0.55f};
    const D2D1_POINT_2F p2{left + size * 0.42f, top + size * 0.75f};
    const D2D1_POINT_2F p3{left + size * 0.80f, top + size * 0.30f};

    const float thickness = DipsToPixels(1.8f, _dpi);
    _target->DrawLine(p1, p2, checkBrush, thickness);
    _target->DrawLine(p2, p3, checkBrush, thickness);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawCollapseChevron(const D2D1_RECT_F& rc, bool collapsed) noexcept
{
    if (collapsed)
    {
        DrawCenteredChevronGlyph(rc, FluentIcons::kChevronDown, FluentIcons::kFallbackChevronDown);
        return;
    }

    DrawCenteredChevronGlyph(rc, FluentIcons::kChevronUp, FluentIcons::kFallbackChevronUp);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawBandwidthGraph(const D2D1_RECT_F& rect,
                                                                               const RateHistory& history,
                                                                               uint64_t limitBytesPerSecond,
                                                                               std::wstring_view overlayText,
                                                                               bool showAnimation,
                                                                               bool rainbowMode,
                                                                               bool perStreamBands,
                                                                               ULONGLONG tick,
                                                                               bool reducedMotion) noexcept
{
    static_cast<void>(limitBytesPerSecond);

    if (! _target)
    {
        return;
    }

    const float w = rect.right - rect.left;
    const float h = rect.bottom - rect.top;
    if (w <= 0.0f || h <= 0.0f)
    {
        return;
    }

    if (_graphBgBrush)
    {
        _target->FillRectangle(rect, _graphBgBrush.get());
    }

    const AppTheme* theme = folderWindow ? &folderWindow->GetTheme() : nullptr;
    // Full saturation belongs to the opt-in Rainbow theme; in normal themes the per-stream
    // bands use a muted, theme-harmonized palette so parallel streams stay distinguishable
    // without turning the default UI into a rainbow.
    const float bandSat = rainbowMode ? 0.85f : 0.42f;
    const float bandVal = rainbowMode ? ((theme && theme->dark) ? 0.80f : 0.90f) : ((theme && theme->dark) ? 0.68f : 0.80f);

    auto sampleColorFromHue = [&](float hue, float alpha) noexcept -> D2D1_COLOR_F
    {
        if (hue < 0.0f)
        {
            D2D1_COLOR_F c = theme ? theme->navigationView.accent : D2D1::ColorF(D2D1::ColorF::DodgerBlue);
            c.a            = alpha;
            return c;
        }
        return ColorFromHSV(hue, bandSat, bandVal, alpha);
    };

    // Helper to compute rainbow color based on tick
    auto computeRainbowColor = [](ULONGLONG tick, ULONGLONG periodMs, float saturation, float value, float alpha) -> D2D1_COLOR_F
    {
        const float hue = static_cast<float>((tick % periodMs) * 360ull / periodMs);
        return ColorFromHSV(hue, saturation, value, alpha);
    };

    showAnimation = showAnimation && ! reducedMotion;

    // Draw animation for pre-calculation phase
    if (showAnimation && _graphDynamicBrush)
    {
        // Pulsing background effect
        constexpr ULONGLONG kPulsePeriodMs = 1600ull;
        const ULONGLONG pulsePhase         = tick % kPulsePeriodMs;
        const float pulseT                 = static_cast<float>(pulsePhase) / static_cast<float>(kPulsePeriodMs);
        const float pulseAlpha             = 0.15f + 0.15f * std::sin(pulseT * 2.0f * 3.14159265f);

        D2D1_COLOR_F pulseColor = _graphFillBaseColor;
        if (rainbowMode)
        {
            // Rainbow: cycle through hues for pulse background
            constexpr ULONGLONG kRainbowPeriodMs = 3000ull;
            pulseColor                           = computeRainbowColor(tick, kRainbowPeriodMs, 0.6f, 0.8f, pulseAlpha);
        }
        else
        {
            pulseColor.a = pulseAlpha;
        }

        _graphDynamicBrush->SetColor(pulseColor);
        _target->FillRectangle(rect, _graphDynamicBrush.get());

        // Horizontal sweep line effect
        constexpr ULONGLONG kSweepPeriodMs = 1200ull;
        const ULONGLONG sweepPhase         = tick % kSweepPeriodMs;
        const float sweepT                 = static_cast<float>(sweepPhase) / static_cast<float>(kSweepPeriodMs);
        const float sweepX                 = rect.left + w * sweepT;

        D2D1_COLOR_F sweepColor = _graphFillBaseColor;
        if (rainbowMode)
        {
            // Rainbow: sweep line changes color each sweep
            sweepColor = computeRainbowColor(tick, kSweepPeriodMs, 0.85f, 0.9f, 0.7f);
        }
        else
        {
            sweepColor.a = 0.5f;
        }

        const float sweepWidth = DipsToPixels(2.0f, _dpi);
        _graphDynamicBrush->SetColor(sweepColor);
        _target->DrawLine(D2D1::Point2F(sweepX, rect.top), D2D1::Point2F(sweepX, rect.bottom), _graphDynamicBrush.get(), sweepWidth);

        // Spinner dots effect (3 dots bouncing)
        constexpr ULONGLONG kSpinPeriodMs = 1000ull;
        constexpr int kDotCount           = 3;
        const float centerX               = rect.left + w * 0.5f;
        const float centerY               = rect.bottom - h * 0.35f;
        const float dotSpacing            = DipsToPixels(10.0f, _dpi);

        for (int i = 0; i < kDotCount; ++i)
        {
            const float phaseOffset  = static_cast<float>(i) / static_cast<float>(kDotCount);
            const ULONGLONG dotPhase = (tick + static_cast<ULONGLONG>(phaseOffset * kSpinPeriodMs)) % kSpinPeriodMs;
            const float dotT         = static_cast<float>(dotPhase) / static_cast<float>(kSpinPeriodMs);
            const float bounce       = std::abs(std::sin(dotT * 3.14159265f));

            const float dotX      = centerX + (static_cast<float>(i) - 1.0f) * dotSpacing;
            const float dotY      = centerY - bounce * DipsToPixels(8.0f, _dpi);
            const float dotRadius = DipsToPixels(3.0f, _dpi);

            D2D1_COLOR_F dotColor = _graphFillBaseColor;
            if (rainbowMode)
            {
                // Rainbow: each dot has its own hue offset
                constexpr ULONGLONG kDotRainbowPeriodMs = 2000ull;
                const ULONGLONG dotRainbowPhase         = tick + static_cast<ULONGLONG>(i * 667); // 120 degree offset per dot
                dotColor                                = computeRainbowColor(dotRainbowPhase, kDotRainbowPeriodMs, 0.85f, 0.9f, 0.6f + 0.4f * bounce);
            }
            else
            {
                dotColor.a = 0.6f + 0.4f * bounce;
            }

            _graphDynamicBrush->SetColor(dotColor);
            _target->FillEllipse(D2D1::Ellipse(D2D1::Point2F(dotX, dotY), dotRadius, dotRadius), _graphDynamicBrush.get());
        }
    }

    if (_borderBrush)
    {
        _target->DrawRectangle(rect, _borderBrush.get(), 1.0f);
    }

    float maxSpeed = 0.0f;
    for (size_t i = 0; i < history.count; ++i)
    {
        const size_t index = (history.writeIndex + RateHistory::kMaxSamples - history.count + i) % RateHistory::kMaxSamples;
        maxSpeed           = std::max(maxSpeed, history.samples[index]);
    }

    const double currentBandwidthBytesPerSecond = CurrentBandwidthForGraphMarker(history);
    maxSpeed = std::max(maxSpeed, static_cast<float>(std::min<double>(currentBandwidthBytesPerSecond, std::numeric_limits<float>::max())));

    if (maxSpeed <= 0.0f)
    {
        maxSpeed = 1.0f;
    }

    const float axisMax = std::max(1.0f, maxSpeed * 1.10f);

    const bool canDrawSamples = _graphLineBrush && history.count >= 2;

    std::array<D2D1_POINT_2F, RateHistory::kMaxSamples> points{};
    std::array<float, RateHistory::kMaxSamples> sampleHues{};
    std::array<std::array<RateHistory::HueWeight, RateHistory::kMaxHueWeightsPerSample>, RateHistory::kMaxSamples> sampleHueWeights{};
    std::array<uint8_t, RateHistory::kMaxSamples> sampleHueWeightCounts{};
    size_t count  = 0;
    size_t oldest = 0;
    if (canDrawSamples)
    {
        count  = history.count;
        oldest = (history.writeIndex + RateHistory::kMaxSamples - count) % RateHistory::kMaxSamples;

        for (size_t i = 0; i < count; ++i)
        {
            const size_t index       = (oldest + i) % RateHistory::kMaxSamples;
            const float speed        = history.samples[index];
            sampleHues[i]            = history.hues[index];
            sampleHueWeightCounts[i] = history.hueWeightCounts[index];
            sampleHueWeights[i]      = history.hueWeights[index];

            const float xFrac = static_cast<float>(i) / static_cast<float>(count - 1u);
            const float yFrac = Clamp01(speed / axisMax);

            const float x = rect.left + w * xFrac;
            const float y = rect.bottom - h * yFrac;
            points[i]     = D2D1::Point2F(x, y);
        }

        if (! reducedMotion && count >= 2u && history.lastDisplaySampleTick != 0 && tick >= history.lastDisplaySampleTick)
        {
            const size_t newest = count - 1u;
            points[newest].y    = EaseGraphLatestPointYForDisplay(points[newest - 1u].y, points[newest].y, tick - history.lastDisplaySampleTick);
        }

        // In normal themes the bands engage only once the history actually carries multiple
        // streams; single-stream copies keep the classic single-color fill.
        bool historyHasMultiStreamSamples = false;
        for (size_t i = 0; i < count && ! historyHasMultiStreamSamples; ++i)
        {
            historyHasMultiStreamSamples = sampleHueWeightCounts[i] >= 2u;
        }

        if (_graphFillBrush && _d2dFactory)
        {
            if ((rainbowMode || (perStreamBands && historyHasMultiStreamSamples)) && _graphDynamicBrush && count >= 2)
            {
                // Per-stream proportional hue bands per segment trapezoid. Quads are grouped by
                // hue first so each hue fills ONE geometry per frame instead of one geometry per
                // band per segment (the graph redraws every 100ms).
                const float fillAlpha = _graphFillBaseColor.a;

                struct HueQuads
                {
                    float hue = -1.0f;
                    std::vector<std::array<D2D1_POINT_2F, 4>> quads;
                };
                std::vector<HueQuads> hueQuads;
                const auto quadsForHue = [&](float hue) noexcept -> std::vector<std::array<D2D1_POINT_2F, 4>>&
                {
                    for (auto& entry : hueQuads)
                    {
                        if (entry.hue == hue)
                        {
                            return entry.quads;
                        }
                    }
                    hueQuads.push_back(HueQuads{.hue = hue});
                    return hueQuads.back().quads;
                };

                for (size_t i = 1; i < count; ++i)
                {
                    const auto& weights      = sampleHueWeights[i];
                    const size_t weightCount = std::min<size_t>(sampleHueWeightCounts[i], weights.size());
                    double totalWeight       = 0.0;
                    for (size_t band = 0; band < weightCount; ++band)
                    {
                        if (weights[band].weight > 0.0)
                        {
                            totalWeight += weights[band].weight;
                        }
                    }

                    const float leftFilledH  = std::max(0.0f, rect.bottom - points[i - 1u].y);
                    const float rightFilledH = std::max(0.0f, rect.bottom - points[i].y);
                    double lowerShare        = 0.0;

                    const auto addBand = [&](float hue, double upperShare) noexcept
                    {
                        const float lower = Clamp01(static_cast<float>(lowerShare));
                        const float upper = Clamp01(static_cast<float>(upperShare));
                        if (upper <= lower)
                        {
                            lowerShare = upperShare;
                            return;
                        }

                        quadsForHue(hue).push_back({D2D1::Point2F(points[i - 1u].x, rect.bottom - leftFilledH * lower),
                                                    D2D1::Point2F(points[i].x, rect.bottom - rightFilledH * lower),
                                                    D2D1::Point2F(points[i].x, rect.bottom - rightFilledH * upper),
                                                    D2D1::Point2F(points[i - 1u].x, rect.bottom - leftFilledH * upper)});
                        lowerShare = upperShare;
                    };

                    if (weightCount > 0u && totalWeight > 0.0)
                    {
                        for (size_t band = 0; band < weightCount; ++band)
                        {
                            if (weights[band].weight <= 0.0)
                            {
                                continue;
                            }

                            const double upperShare = lowerShare + (weights[band].weight / totalWeight);
                            addBand(weights[band].hue, upperShare);
                        }
                    }
                    else
                    {
                        addBand(sampleHues[i], 1.0);
                    }
                }

                for (const auto& entry : hueQuads)
                {
                    if (entry.quads.empty())
                    {
                        continue;
                    }

                    wil::com_ptr<ID2D1PathGeometry> geometry;
                    if (FAILED(_d2dFactory->CreatePathGeometry(geometry.put())) || ! geometry)
                    {
                        continue;
                    }

                    wil::com_ptr<ID2D1GeometrySink> sink;
                    if (FAILED(geometry->Open(sink.put())) || ! sink)
                    {
                        continue;
                    }

                    sink->SetFillMode(D2D1_FILL_MODE_WINDING);
                    for (const auto& quad : entry.quads)
                    {
                        sink->BeginFigure(quad[0], D2D1_FIGURE_BEGIN_FILLED);
                        sink->AddLine(quad[1]);
                        sink->AddLine(quad[2]);
                        sink->AddLine(quad[3]);
                        sink->EndFigure(D2D1_FIGURE_END_CLOSED);
                    }
                    sink->Close();

                    _graphDynamicBrush->SetColor(sampleColorFromHue(entry.hue, fillAlpha));
                    _target->FillGeometry(geometry.get(), _graphDynamicBrush.get());
                }
            }
            else
            {
                // Non-rainbow: draw single fill geometry
                wil::com_ptr<ID2D1PathGeometry> geometry;
                const HRESULT hrGeo = _d2dFactory->CreatePathGeometry(geometry.put());
                if (SUCCEEDED(hrGeo) && geometry)
                {
                    wil::com_ptr<ID2D1GeometrySink> sink;
                    const HRESULT hrSink = geometry->Open(sink.put());
                    if (SUCCEEDED(hrSink) && sink)
                    {
                        sink->SetFillMode(D2D1_FILL_MODE_WINDING);
                        sink->BeginFigure(points[0], D2D1_FIGURE_BEGIN_FILLED);
                        sink->AddLines(points.data() + 1, static_cast<UINT32>(count - 1u));

                        sink->AddLine(D2D1::Point2F(points[count - 1u].x, rect.bottom));
                        sink->AddLine(D2D1::Point2F(points[0].x, rect.bottom));

                        sink->EndFigure(D2D1_FIGURE_END_CLOSED);
                        sink->Close();

                        _target->FillGeometry(geometry.get(), _graphFillBrush.get());
                    }
                }
            }
        }
    }

    if (_graphGridBrush)
    {
        for (int i = 1; i <= 3; ++i)
        {
            const float frac = static_cast<float>(i) / 4.0f;
            const float y    = rect.bottom - h * frac;
            _target->DrawLine(D2D1::Point2F(rect.left, y), D2D1::Point2F(rect.right, y), _graphGridBrush.get(), 1.0f);
        }
    }

    std::wstring currentBandwidthText;
    if (currentBandwidthBytesPerSecond > 0.0 && _graphLimitBrush)
    {
        const float currentFrac = Clamp01(static_cast<float>(currentBandwidthBytesPerSecond / static_cast<double>(axisMax)));
        const float y           = rect.bottom - h * currentFrac;
        _target->DrawLine(D2D1::Point2F(rect.left, y), D2D1::Point2F(rect.right, y), _graphLimitBrush.get(), 1.0f);

        const uint64_t roundedBytesPerSecond = SaturatingRoundNonNegativeToUint64(currentBandwidthBytesPerSecond);
        currentBandwidthText                 = FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_BYTES, FormatBytesCompact(roundedBytesPerSecond));
    }

    if (canDrawSamples && rainbowMode)
    {
        // Rainbow: draw each line segment with its own hue from the stored per-sample hue
        if (_graphDynamicBrush)
        {
            for (size_t i = 1; i < count; ++i)
            {
                const float hue                = sampleHues[i];
                const D2D1_COLOR_F segmentLine = sampleColorFromHue(hue, 1.0f);
                _graphDynamicBrush->SetColor(segmentLine);
                _target->DrawLine(points[i - 1u], points[i], _graphDynamicBrush.get(), 1.5f);
            }
        }
        else
        {
            for (size_t i = 1; i < count; ++i)
            {
                _target->DrawLine(points[i - 1u], points[i], _graphLineBrush.get(), 1.5f);
            }
        }
    }
    else if (canDrawSamples)
    {
        for (size_t i = 1; i < count; ++i)
        {
            _target->DrawLine(points[i - 1u], points[i], _graphLineBrush.get(), 1.5f);
        }
    }

    if (! currentBandwidthText.empty() && _smallFormat && _textBrush)
    {
        const float inset  = DipsToPixels(6.0f, _dpi);
        const float labelH = DipsToPixels(18.0f, _dpi);
        const D2D1_RECT_F labelRc =
            D2D1::RectF(rect.left + inset, rect.top + DipsToPixels(3.0f, _dpi), rect.right - inset, rect.top + DipsToPixels(3.0f, _dpi) + labelH);
        if (_graphTextShadowBrush)
        {
            const float shadowOffset = DipsToPixels(1.0f, _dpi);
            const D2D1_RECT_F shadowRc =
                D2D1::RectF(labelRc.left + shadowOffset, labelRc.top + shadowOffset, labelRc.right + shadowOffset, labelRc.bottom + shadowOffset);
            _target->DrawTextW(currentBandwidthText.data(),
                               static_cast<UINT32>(currentBandwidthText.size()),
                               _smallFormat.get(),
                               shadowRc,
                               _graphTextShadowBrush.get(),
                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }
        _target->DrawTextW(currentBandwidthText.data(),
                           static_cast<UINT32>(currentBandwidthText.size()),
                           _smallFormat.get(),
                           labelRc,
                           _textBrush.get(),
                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
    }

    if (! overlayText.empty() && _graphOverlayFormat && _textBrush)
    {
        // Draw shadow behind text for better visibility
        if (_graphTextShadowBrush)
        {
            const float shadowOffset = DipsToPixels(1.0f, _dpi);
            const D2D1_RECT_F shadowRect =
                D2D1::RectF(rect.left + shadowOffset, rect.top + shadowOffset, rect.right + shadowOffset, rect.bottom + shadowOffset);
            _target->DrawTextW(overlayText.data(), static_cast<UINT32>(overlayText.size()), _graphOverlayFormat.get(), shadowRect, _graphTextShadowBrush.get());
        }

        // Draw main text
        _target->DrawTextW(overlayText.data(), static_cast<UINT32>(overlayText.size()), _graphOverlayFormat.get(), rect, _textBrush.get());
    }
}

void FileOperationsPopupInternal::FileOperationsPopupState::Render(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    static_cast<void>(hdc.get());
    static_cast<void>(ps);

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

    EnsureTarget(hwnd);
    EnsureTextFormats();
    EnsureBrushes();

    if (! _target || ! _bgBrush || ! _textBrush || ! _borderBrush)
    {
        return;
    }

    const bool capturePerf                   = Debug::Perf::IsCaptureEnabled();
    const bool reducedMotion                 = IsReducedMotionEnabled();
    const uint64_t renderStartedUs           = capturePerf ? PerfNowUs() : 0u;
    const uint64_t snapshotStartedUs         = capturePerf ? PerfNowUs() : 0u;
    const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
    const uint64_t snapshotUs                = capturePerf ? PerfElapsedUs(snapshotStartedUs) : 0u;
    CleanupCollapsedTasks(snapshot);
    AutoCollapseCompletedTasks(snapshot);
    UpdateCaptionStatus(hwnd, snapshot);
    const bool footerOnly                           = fileOps ? fileOps->GetPopupFooterOnly() : false;
    const bool compactDensity                       = fileOps ? fileOps->GetPopupCompactDensity() : false;
    GlobalFileOperationsStatusSummary globalSummary = BuildGlobalStatusSummary(snapshot, &_rates);
    const bool showPauseResumeAll                   = HasFooterPauseResumeAllControl(globalSummary);
    const uint32_t completedGroupCount              = CountCompletedGroupTasks(snapshot);
    const bool showCompletedGroup                   = ShouldShowCompletedGroup(completedGroupCount);

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
                if (! _completedGroupExpanded)
                {
                    continue;
                }
            }

            displayRows.push_back(PopupDisplayRow{.kind = PopupDisplayRowKind::Task, .taskIndex = i});
        }
    }

    constexpr ULONGLONG kCompletedInFlightGraceMs = 300ull;
    const ULONGLONG renderTick                    = GetTickCount64();

    float width  = 0.0f;
    float height = 0.0f;

    const float padding = DipsToPixels(10.0f, _dpi);
    const float cardGap = DipsToPixels(10.0f, _dpi);

    const float expandedCardH  = DipsToPixels(280.0f, _dpi);
    const float collapsedCardH = DipsToPixels(44.0f, _dpi);
    const float baseLineH      = DipsToPixels(18.0f, _dpi);
    const float fromToGapY     = DipsToPixels(4.0f, _dpi);

    const uint64_t cardLayoutStartedUs = capturePerf ? PerfNowUs() : 0u;
    std::vector<float> cardHeights;
    cardHeights.reserve(displayRows.size());
    if (! footerOnly)
    {
        for (const PopupDisplayRow& row : displayRows)
        {
            if (row.kind == PopupDisplayRowKind::CompletedGroup)
            {
                cardHeights.push_back(DipsToPixels(42.0f, _dpi));
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

                cardHeights.push_back(h);
                continue;
            }

            const bool displayCollapsed = IsTaskCollapsedForDisplay(task.taskId, compactDensity);
            float h                     = displayCollapsed ? collapsedCardH : expandedCardH;
            if (! displayCollapsed && task.finished)
            {
                h = DipsToPixels(178.0f, _dpi);
            }
            if (! displayCollapsed && ! task.finished && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE))
            {
                size_t activeInFlightCount = 0;
                for (size_t i = 0; i < task.inFlightFileCount; ++i)
                {
                    const auto& entry          = task.inFlightFiles[i];
                    const bool active          = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                    const bool recentCompleted = ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes && entry.lastUpdateTick != 0 &&
                                                 renderTick >= entry.lastUpdateTick && (renderTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
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
                h += fromToGapY;
            }
            if (! displayCollapsed && ! task.finished && task.conflict.active)
            {
                // Extra room for stacked conflict labels plus full-width path rows.
                h += baseLineH * (task.operation == FILESYSTEM_DELETE ? 3.0f : 5.0f);
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
    AutoResizeWindow(hwnd, _contentHeight, rowCount, footerOnly, reducedMotion);
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
            _target->Resize(D2D1::SizeU(clientW, clientH));
        }

        width  = static_cast<float>(clientW);
        height = static_cast<float>(clientH);

        LayoutChrome(width, height, showPauseResumeAll);

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

            _hotHit     = {};
            _pressedHit = {};

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

    HRESULT hrEndDraw            = S_OK;
    const uint64_t drawStartedUs = capturePerf ? PerfNowUs() : 0u;
    {
        _target->BeginDraw();
        auto endDraw = wil::scope_exit([&] { hrEndDraw = _target->EndDraw(); });

        _target->SetTransform(D2D1::Matrix3x2F::Identity());

        const D2D1_RECT_F clientRect = D2D1::RectF(0.0f, 0.0f, width, height);
        _target->FillRectangle(clientRect, _bgBrush.get());

        const float footerH          = FileOperationsPopupFooterHeightPixels(_dpi);
        const float footerTop        = std::max(0.0f, height - footerH);
        const D2D1_RECT_F footerRect = D2D1::RectF(0.0f, footerTop, width, height);
        _target->DrawRectangle(footerRect, _borderBrush.get(), 1.0f);

        const GlobalTaskbarProgressModel taskbarProgress = BuildGlobalTaskbarProgressModel(globalSummary);
        ApplyTaskbarProgress(hwnd, taskbarProgress.state, taskbarProgress.completed, taskbarProgress.total);

        if (HasGlobalAggregateProgress(globalSummary) && _footerAggregateProgressRect.right > _footerAggregateProgressRect.left &&
            _footerAggregateProgressRect.bottom > _footerAggregateProgressRect.top)
        {
            if (_progressBgBrush)
            {
                const float radius = ClampCornerRadius(_footerAggregateProgressRect, DipsToPixels(3.0f, _dpi));
                _target->FillRoundedRectangle(D2D1::RoundedRect(_footerAggregateProgressRect, radius, radius), _progressBgBrush.get());
            }
            if (_progressGlobalBrush)
            {
                const bool determinate = HasDeterminateGlobalAggregateProgress(globalSummary);
                const D2D1_RECT_F fill =
                    determinate ? D2D1::RectF(_footerAggregateProgressRect.left,
                                              _footerAggregateProgressRect.top,
                                              _footerAggregateProgressRect.left + (_footerAggregateProgressRect.right - _footerAggregateProgressRect.left) *
                                                                                      GlobalAggregateProgressFraction(globalSummary),
                                              _footerAggregateProgressRect.bottom)
                                : ComputeIndeterminateBarFill(_footerAggregateProgressRect, renderTick, reducedMotion);
                if (fill.right > fill.left && fill.bottom > fill.top)
                {
                    const float radius = ClampCornerRadius(fill, DipsToPixels(3.0f, _dpi));
                    _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressGlobalBrush.get());
                }
            }
        }

        const auto rectHasArea = [](const D2D1_RECT_F& rc) noexcept { return rc.right > rc.left && rc.bottom > rc.top; };

        PopupButton cancelAllBtn{};
        cancelAllBtn.bounds   = _footerCancelAllRect;
        cancelAllBtn.hit.kind = PopupHitTest::Kind::FooterCancelAll;
        if (rectHasArea(cancelAllBtn.bounds))
        {
            _buttons.push_back(cancelAllBtn);
        }

        PopupButton autoDismissBtn{};
        autoDismissBtn.bounds   = _footerAutoDismissRect;
        autoDismissBtn.hit.kind = PopupHitTest::Kind::FooterAutoDismiss;
        if (rectHasArea(autoDismissBtn.bounds))
        {
            _buttons.push_back(autoDismissBtn);
        }

        const bool pauseResumeAllPauses = FooterPauseResumeAllShouldPause(globalSummary);
        PopupButton pauseResumeAllBtn{};
        pauseResumeAllBtn.bounds   = _footerPauseResumeAllRect;
        pauseResumeAllBtn.hit.kind = PopupHitTest::Kind::FooterPauseResumeAll;
        pauseResumeAllBtn.hit.data = pauseResumeAllPauses ? kFooterPauseResumeAllPauseAction : kFooterPauseResumeAllResumeAction;
        if (showPauseResumeAll && rectHasArea(pauseResumeAllBtn.bounds))
        {
            _buttons.push_back(pauseResumeAllBtn);
        }

        PopupButton queueBtn{};
        queueBtn.bounds   = _footerQueueModeRect;
        queueBtn.hit.kind = PopupHitTest::Kind::FooterQueueMode;

        PopupButton densityBtn{};
        densityBtn.bounds   = _footerDensityRect;
        densityBtn.hit.kind = PopupHitTest::Kind::FooterDensity;
        if (rectHasArea(densityBtn.bounds))
        {
            _buttons.push_back(densityBtn);
        }

        PopupButton detailsBtn{};
        detailsBtn.bounds   = _footerDetailsToggleRect;
        detailsBtn.hit.kind = PopupHitTest::Kind::FooterToggleDetails;
        if (rectHasArea(detailsBtn.bounds))
        {
            _buttons.push_back(detailsBtn);
        }

        const bool hasActiveOperations = fileOps ? fileOps->HasActiveOperations() : false;
        const UINT footerActionId = hasActiveOperations ? static_cast<UINT>(IDS_FILEOPS_BTN_CANCEL_ALL) : static_cast<UINT>(IDS_FILEOPS_BTN_CLEAR_COMPLETED);
        const std::wstring cancelAllText = LoadStringResource(nullptr, footerActionId);
        if (rectHasArea(cancelAllBtn.bounds))
        {
            DrawButton(cancelAllBtn, _buttonFormat.get(), cancelAllText);
        }

        if (showPauseResumeAll && rectHasArea(pauseResumeAllBtn.bounds))
        {
            DrawButton(pauseResumeAllBtn,
                       _buttonSmallFormat.get(),
                       LoadStringResource(nullptr, pauseResumeAllPauses ? IDS_FILEOPS_BTN_PAUSE_ALL : IDS_FILEOPS_BTN_RESUME_ALL));
        }

        if (rectHasArea(autoDismissBtn.bounds))
        {
            DrawFooterAutoDismissControl(autoDismissBtn, fileOps ? fileOps->GetAutoDismissSuccess() : false);
        }

        if (rectHasArea(detailsBtn.bounds))
        {
            DrawButton(detailsBtn, _buttonSmallFormat.get(), {});
            DrawCenteredChevronGlyph(detailsBtn.bounds,
                                     footerOnly ? FluentIcons::kChevronUp : FluentIcons::kChevronDown,
                                     footerOnly ? FluentIcons::kFallbackChevronUp : FluentIcons::kFallbackChevronDown);
        }

        const bool queueMode = fileOps ? fileOps->GetQueueNewTasks() : true;
        if (rectHasArea(queueBtn.bounds))
        {
            DrawFooterQueueModeControl(queueBtn, queueMode, reducedMotion);
            if (rectHasArea(_footerQueueSegmentRect))
            {
                PopupButton queueSegment{};
                queueSegment.bounds   = _footerQueueSegmentRect;
                queueSegment.hit.kind = PopupHitTest::Kind::FooterQueueMode;
                queueSegment.hit.data = kFooterQueueModeQueueAction;
                _buttons.push_back(queueSegment);
            }
            if (rectHasArea(_footerParallelSegmentRect))
            {
                PopupButton parallelSegment{};
                parallelSegment.bounds   = _footerParallelSegmentRect;
                parallelSegment.hit.kind = PopupHitTest::Kind::FooterQueueMode;
                parallelSegment.hit.data = kFooterQueueModeParallelAction;
                _buttons.push_back(parallelSegment);
            }
        }

        if (rectHasArea(densityBtn.bounds))
        {
            if ((densityBtn.bounds.right - densityBtn.bounds.left) < DipsToPixels(64.0f, _dpi))
            {
                DrawButton(densityBtn, _buttonSmallFormat.get(), {});
                DrawCenteredChevronGlyph(densityBtn.bounds, FluentIcons::kBulletedList, FluentIcons::kFallbackBulletedList);
            }
            else
            {
                DrawButton(densityBtn,
                           _buttonSmallFormat.get(),
                           LoadStringResource(nullptr, compactDensity ? IDS_FILEOPS_BTN_DENSITY_COMPACT : IDS_FILEOPS_BTN_DENSITY_EXPANDED));
            }
        }

        const std::wstring globalSummaryText = FormatGlobalStatusSummaryText(globalSummary);
        if (_smallFormat && _subTextBrush && _footerSummaryRect.right > _footerSummaryRect.left && _footerSummaryRect.bottom > _footerSummaryRect.top)
        {
            _target->DrawTextW(globalSummaryText.data(),
                               static_cast<UINT32>(globalSummaryText.size()),
                               _smallFormat.get(),
                               _footerSummaryRect,
                               _subTextBrush.get(),
                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }

        float y = _listViewportRect.top + padding - _scrollY;

        const float cardW = std::max(0.0f, width - padding * 2.0f);

        _target->PushAxisAlignedClip(_listViewportRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        for (size_t rowIndex = 0; rowIndex < rowCount; ++rowIndex)
        {
            const PopupDisplayRow& row = displayRows[rowIndex];
            const float taskCardH      = cardHeights[rowIndex];
            const D2D1_RECT_F cardRect = D2D1::RectF(padding, y, padding + cardW, y + taskCardH);

            const bool visible = cardRect.bottom >= _listViewportRect.top && cardRect.top <= _listViewportRect.bottom;
            if (visible)
            {
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
                    const float chevronSize     = DipsToPixels(18.0f, _dpi);
                    const float chevronTop      = cardRect.top + (taskCardH - chevronSize) * 0.5f;
                    const D2D1_RECT_F chevronRc = D2D1::RectF(textX, chevronTop, textX + chevronSize, chevronTop + chevronSize);
                    DrawCollapseChevron(chevronRc, ! _completedGroupExpanded);

                    const float clearW = std::min(DipsToPixels(128.0f, _dpi), std::max(0.0f, (contentRight - textX) * 0.34f));
                    PopupButton clearBtn{};
                    clearBtn.bounds   = D2D1::RectF(std::max(textX, contentRight - clearW),
                                                    cardRect.top + DipsToPixels(7.0f, _dpi),
                                                    contentRight,
                                                    cardRect.bottom - DipsToPixels(7.0f, _dpi));
                    clearBtn.hit.kind = PopupHitTest::Kind::CompletedGroupClear;

                    PopupButton groupToggle{};
                    groupToggle.bounds =
                        D2D1::RectF(cardRect.left, cardRect.top, std::max(cardRect.left, clearBtn.bounds.left - DipsToPixels(8.0f, _dpi)), cardRect.bottom);
                    groupToggle.hit.kind = PopupHitTest::Kind::CompletedGroupToggle;
                    if (groupToggle.bounds.right > groupToggle.bounds.left && groupToggle.bounds.bottom > groupToggle.bounds.top)
                    {
                        _buttons.push_back(groupToggle);
                    }

                    if (clearBtn.bounds.right > clearBtn.bounds.left && clearBtn.bounds.bottom > clearBtn.bounds.top)
                    {
                        _buttons.push_back(clearBtn);
                        DrawButton(clearBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_BTN_CLEAR_COMPLETED));
                    }

                    const std::wstring groupText =
                        FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPLETED_GROUP, static_cast<unsigned long>(completedGroupCount));
                    const float textLeft  = chevronRc.right + DipsToPixels(8.0f, _dpi);
                    const float textRight = std::max(textLeft, clearBtn.bounds.left - DipsToPixels(8.0f, _dpi));
                    const D2D1_RECT_F textRc =
                        D2D1::RectF(textLeft, cardRect.top + DipsToPixels(10.0f, _dpi), textRight, cardRect.bottom - DipsToPixels(8.0f, _dpi));
                    IDWriteTextFormat* groupFormat = _headerFormat ? _headerFormat.get() : (_bodyFormat ? _bodyFormat.get() : nullptr);
                    if (groupFormat && _textBrush)
                    {
                        _target->DrawTextW(
                            groupText.data(), static_cast<UINT32>(groupText.size()), groupFormat, textRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
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
                            statusText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_STATUS_FAILED, static_cast<unsigned long>(info.resultHr));
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
                    DrawButton(collapseBtn, nullptr, {});
                    DrawCollapseChevron(collapseBtn.bounds, isCollapsedTask);

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

                    if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories)
                    {
                        drawLabeledPathLine(IDS_PREFS_PANES_HEADER_LEFT, info.leftRoot);
                        drawLabeledPathLine(IDS_PREFS_PANES_HEADER_RIGHT, info.rightRoot);

                        if (_smallFormat && _textBrush && (info.scanActive || info.scanFolderCount > 0 || info.scanEntryCount > 0))
                        {
                            const std::wstring scanPath = info.scanCurrentRelative.empty() ? std::wstring(L".") : info.scanCurrentRelative.native();
                            const std::wstring scanText =
                                FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, scanPath, info.scanFolderCount, info.scanEntryCount);
                            const D2D1_RECT_F scanRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(scanText.data(),
                                               static_cast<UINT32>(scanText.size()),
                                               _smallFormat.get(),
                                               scanRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
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
                            const std::wstring scanPath = info.changeCaseCurrentPath.empty() ? std::wstring(L".") : info.changeCaseCurrentPath.native();
                            const std::wstring scanText = FormatStringResource(
                                nullptr, IDS_FMT_COMPARE_SCAN_STATUS, scanPath, info.changeCaseScannedFolders, info.changeCaseScannedEntries);
                            const D2D1_RECT_F scanRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(scanText.data(),
                                               static_cast<UINT32>(scanText.size()),
                                               _smallFormat.get(),
                                               scanRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
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
                            const std::wstring scanPath =
                                info.changeAttributesCurrentPath.empty() ? std::wstring(L".") : info.changeAttributesCurrentPath.native();
                            const std::wstring scanText = FormatStringResource(
                                nullptr, IDS_FMT_COMPARE_SCAN_STATUS, scanPath, info.changeAttributesScannedFolders, info.changeAttributesScannedEntries);
                            const D2D1_RECT_F scanRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(scanText.data(),
                                               static_cast<UINT32>(scanText.size()),
                                               _smallFormat.get(),
                                               scanRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
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
                            const std::wstring scanPath = info.makeFileListCurrentPath.empty() ? std::wstring(L".") : info.makeFileListCurrentPath.native();
                            const std::wstring scanText = FormatStringResource(
                                nullptr, IDS_FMT_COMPARE_SCAN_STATUS, scanPath, info.makeFileListScannedFolders, info.makeFileListScannedEntries);
                            const D2D1_RECT_F scanRc = D2D1::RectF(textX, textY, contentRight, textY + lineH);
                            _target->DrawTextW(scanText.data(),
                                               static_cast<UINT32>(scanText.size()),
                                               _smallFormat.get(),
                                               scanRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            textY += lineH;
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
                                        const D2D1::ColorF rainbow = RainbowProgressColor(appTheme, sourcePathText);
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

                        if (_progressBgBrush)
                        {
                            const float radius = ClampCornerRadius(barRc, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(barRc, radius, radius), _progressBgBrush.get());
                        }

                        if (_progressGlobalBrush)
                        {
                            bool hasTotal = false;
                            float frac    = 0.0f;
                            if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories)
                            {
                                hasTotal = info.contentTotalBytes > 0 && info.contentCompletedBytes <= info.contentTotalBytes;
                                frac     = hasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.contentCompletedBytes) /
                                                                                 static_cast<double>(info.contentTotalBytes)))
                                                    : 0.0f;
                            }
                            else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeCase)
                            {
                                hasTotal = info.changeCasePlannedRenames > 0 && info.changeCaseCompletedRenames <= info.changeCasePlannedRenames;
                                frac     = hasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.changeCaseCompletedRenames) /
                                                                                 static_cast<double>(info.changeCasePlannedRenames)))
                                                    : 0.0f;
                            }
                            else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::ChangeAttributes)
                            {
                                hasTotal = info.changeAttributesPlannedItems > 0 && info.changeAttributesCompletedItems <= info.changeAttributesPlannedItems;
                                frac     = hasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.changeAttributesCompletedItems) /
                                                                                 static_cast<double>(info.changeAttributesPlannedItems)))
                                                    : 0.0f;
                            }
                            else if (info.kind == FolderWindow::InformationalTaskUpdate::Kind::MakeFileList)
                            {
                                hasTotal = info.makeFileListRendering && info.makeFileListTotalEntries > 0u &&
                                           info.makeFileListRenderedEntries <= info.makeFileListTotalEntries;
                                frac     = hasTotal ? Clamp01(static_cast<float>(static_cast<double>(info.makeFileListRenderedEntries) /
                                                                                 static_cast<double>(info.makeFileListTotalEntries)))
                                                    : 0.0f;
                            }

                            const D2D1_RECT_F fill = hasTotal ? D2D1::RectF(barRc.left, barRc.top, barRc.left + (barRc.right - barRc.left) * frac, barRc.bottom)
                                                              : ComputeIndeterminateBarFill(barRc, nowTick, reducedMotion);
                            const float radius     = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressGlobalBrush.get());
                        }
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
                    if (task.desiredSpeedLimitBytesPerSecond == 0)
                    {
                        speedLimitText = LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_BUTTON_UNLIMITED);
                    }
                    else
                    {
                        speedLimitText =
                            FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_BUTTON_BYTES, FormatBytesCompact(task.desiredSpeedLimitBytesPerSecond));
                    }
                }

                const UINT opTextId = [&]() -> UINT
                {
                    switch (task.operation)
                    {
                        case FILESYSTEM_COPY: return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
                        case FILESYSTEM_MOVE: return static_cast<UINT>(IDS_FILEOP_OPERATION_MOVE);
                        case FILESYSTEM_DELETE: return static_cast<UINT>(IDS_FILEOP_OPERATION_DELETE);
                        case FILESYSTEM_RENAME: return static_cast<UINT>(IDS_FILEOP_OPERATION_RENAME);
                    }
                    return static_cast<UINT>(IDS_FILEOP_OPERATION_COPY);
                }();

                const std::wstring opText = LoadStringResource(nullptr, opTextId);
                const ULONGLONG nowTick   = renderTick;

                const TaskStatusKind taskStatus        = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
                const std::wstring headerText          = BuildTaskHeaderText(task, opText, nowTick);
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
                DrawButton(collapseBtn, nullptr, {});
                DrawCollapseChevron(collapseBtn.bounds, isCollapsedTask);

                const float headerRight = std::max(textX, collapseBtn.bounds.left - collapseBtnGap);
                float headerLeft        = textX;

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

                const std::wstring chipText = StatusChipTextForTask(task, taskStatus, nowTick);
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
                            const D2D1_RECT_F chipRc = D2D1::RectF(headerLeft, chipTop, headerLeft + chipW, chipTop + chipH);
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
                            headerLeft = std::min(headerRight, chipRc.right + chipGap);
                        }
                    }
                }

                float headerTextRight = headerRight;
                if (isCollapsedTask && TaskHasKnownCompactProgress(task) && _progressBgBrush && _progressGlobalBrush && _smallFormat && _subTextBrush)
                {
                    const float availableHeaderW = std::max(0.0f, headerRight - headerLeft);
                    const float progressGap      = DipsToPixels(8.0f, _dpi);
                    const float percentW         = DipsToPixels(42.0f, _dpi);
                    const float preferredMeterW  = DipsToPixels(118.0f, _dpi);
                    const float compactMeterW    = std::min(preferredMeterW, availableHeaderW * 0.40f);
                    const float barW             = compactMeterW - percentW - DipsToPixels(6.0f, _dpi);
                    if (barW >= DipsToPixels(36.0f, _dpi) && compactMeterW + progressGap < availableHeaderW)
                    {
                        const float meterRight = headerRight;
                        const float meterLeft  = meterRight - compactMeterW;
                        headerTextRight        = std::max(headerLeft, meterLeft - progressGap);

                        const float fraction    = ComputeFileOperationsTaskCompleteFractionForDisplay(task);
                        const uint32_t percent  = static_cast<uint32_t>(std::lround(Clamp01(fraction) * 100.0f));
                        const float barH        = DipsToPixels(6.0f, _dpi);
                        const float barTop      = headerTop + (lineH - barH) * 0.5f;
                        const D2D1_RECT_F barRc = D2D1::RectF(meterLeft, barTop, meterLeft + barW, barTop + barH);
                        const float trackRadius = ClampCornerRadius(barRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(barRc, trackRadius, trackRadius), _progressBgBrush.get());

                        const D2D1_RECT_F fillRc =
                            D2D1::RectF(barRc.left, barRc.top, barRc.left + (barRc.right - barRc.left) * Clamp01(fraction), barRc.bottom);
                        const float fillRadius = ClampCornerRadius(fillRc, DipsToPixels(2.0f, _dpi));
                        if (fillRc.right > fillRc.left)
                        {
                            _target->FillRoundedRectangle(D2D1::RoundedRect(fillRc, fillRadius, fillRadius), _progressGlobalBrush.get());
                        }

                        const std::wstring percentText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_PERCENT, percent);
                        const D2D1_RECT_F percentRc    = D2D1::RectF(barRc.right + DipsToPixels(6.0f, _dpi), headerTop, meterRight, headerBottom);
                        _target->DrawTextW(percentText.data(),
                                           static_cast<UINT32>(percentText.size()),
                                           _smallFormat.get(),
                                           percentRc,
                                           _subTextBrush.get(),
                                           D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }
                }

                if (_headerFormat)
                {
                    const D2D1_RECT_F headerRc = D2D1::RectF(headerLeft, headerTop, headerTextRight, headerBottom);
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

                textY = headerBottom;

                if (task.finished)
                {
                    const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    const bool showHrLine   = FAILED(task.resultHr) && task.resultHr != partialHr;

                    const std::wstring diagCounts = FormatStringResource(nullptr, IDS_FMT_FILEOPS_WARNINGS_ERRORS, task.warningCount, task.errorCount);
                    const D2D1_RECT_F countsRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(diagCounts.data(), static_cast<UINT32>(diagCounts.size()), _bodyFormat.get(), countsRc, _subTextBrush.get());
                    textY += lineH;

                    if (showHrLine)
                    {
                        const std::wstring hrText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_RESULT_HRESULT, static_cast<unsigned long>(task.resultHr));
                        const D2D1_RECT_F hrRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(hrText.data(), static_cast<UINT32>(hrText.size()), _bodyFormat.get(), hrRc, _subTextBrush.get());
                        textY += lineH;
                    }

                    const float labelWDesired   = DipsToPixels(56.0f, _dpi);
                    const float labelGapDesired = DipsToPixels(6.0f, _dpi);
                    const float labelW          = std::min(labelWDesired, textMaxW);
                    const float labelGap        = (labelW < textMaxW) ? std::min(labelGapDesired, textMaxW - labelW) : 0.0f;
                    const float pathW           = std::max(0.0f, textMaxW - labelW - labelGap);

                    if (task.operation == FILESYSTEM_DELETE)
                    {
                        const std::wstring label  = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_DELETING);
                        const D2D1_RECT_F labelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get());

                        const std::wstring path  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentSourcePath, pathW, lineH);
                        const D2D1_RECT_F pathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(
                            path.data(), static_cast<UINT32>(path.size()), _bodyFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        textY += lineH;
                    }
                    else
                    {
                        const std::wstring fromLabel  = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FROM);
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
                        textY += lineH;

                        const std::wstring toLabel  = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO);
                        const D2D1_RECT_F toLabelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                        _target->DrawTextW(toLabel.data(), static_cast<UINT32>(toLabel.size()), _smallFormat.get(), toLabelRc, _subTextBrush.get());

                        const std::wstring toPath =
                            TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentDestinationPath, pathW, lineH);
                        const D2D1_RECT_F toPathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                        _target->DrawTextW(
                            toPath.data(), static_cast<UINT32>(toPath.size()), _bodyFormat.get(), toPathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        textY += lineH;
                    }

                    if (! task.lastDiagnosticMessage.empty())
                    {
                        const std::wstring diagText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_LAST_NOTE, task.lastDiagnosticMessage);
                        const D2D1_RECT_F diagRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(diagText.data(), static_cast<UINT32>(diagText.size()), _smallFormat.get(), diagRc, _subTextBrush.get());
                        textY += lineH;
                    }

                    const float dismissButtonH         = DipsToPixels(24.0f, _dpi);
                    const float dismissButtonBottomPad = DipsToPixels(8.0f, _dpi);
                    const float dismissButtonTop       = std::max(textY + DipsToPixels(4.0f, _dpi), cardRect.bottom - dismissButtonBottomPad - dismissButtonH);

                    const float progressBarH         = DipsToPixels(8.0f, _dpi);
                    const float progressBarBottomPad = DipsToPixels(6.0f, _dpi);
                    const float progressBarTop       = std::max(textY + DipsToPixels(2.0f, _dpi), dismissButtonTop - progressBarBottomPad - progressBarH);
                    const D2D1_RECT_F progressRc     = D2D1::RectF(textX, progressBarTop, contentRight, progressBarTop + progressBarH);

                    if (_progressBgBrush)
                    {
                        const float radius = ClampCornerRadius(progressRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(progressRc, radius, radius), _progressBgBrush.get());
                    }

                    const float completeFraction = ComputeFileOperationsTaskCompleteFractionForDisplay(task);

                    if (_progressGlobalBrush)
                    {
                        const D2D1_RECT_F fillRc = D2D1::RectF(
                            progressRc.left, progressRc.top, progressRc.left + (progressRc.right - progressRc.left) * completeFraction, progressRc.bottom);
                        const float radius = ClampCornerRadius(fillRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fillRc, radius, radius), _progressGlobalBrush.get());
                    }

                    if (CompletedTaskHasOverflowActions(task))
                    {
                        const float btnGap         = DipsToPixels(6.0f, _dpi);
                        const float totalW         = std::max(0.0f, contentRight - textX);
                        const float preferredMoreW = DipsToPixels(96.0f, _dpi);
                        const float moreW          = totalW > DipsToPixels(180.0f, _dpi) ? std::min(preferredMoreW, std::max(0.0f, totalW * 0.35f))
                                                                                         : std::max(0.0f, (totalW - btnGap) * 0.5f);
                        const float dismissW       = std::max(0.0f, totalW - btnGap - moreW);

                        PopupButton dismissBtn{};
                        dismissBtn.bounds     = D2D1::RectF(textX, dismissButtonTop, textX + dismissW, dismissButtonTop + dismissButtonH);
                        dismissBtn.hit.kind   = PopupHitTest::Kind::TaskDismiss;
                        dismissBtn.hit.taskId = task.taskId;
                        _buttons.push_back(dismissBtn);
                        DrawButton(dismissBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOP_BTN_DISMISS));

                        PopupButton moreBtn{};
                        moreBtn.bounds     = D2D1::RectF(textX + dismissW + btnGap, dismissButtonTop, contentRight, dismissButtonTop + dismissButtonH);
                        moreBtn.hit.kind   = PopupHitTest::Kind::TaskCompletedMore;
                        moreBtn.hit.taskId = task.taskId;
                        moreBtn.hit.data   = 2u;
                        _buttons.push_back(moreBtn);
                        DrawMenuButton(moreBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_MORE));
                    }
                    else
                    {
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

                const AppTheme& theme = folderWindow->GetTheme();

                const RateHistory* history = nullptr;
                const auto historyIt       = _rates.find(task.taskId);
                if (historyIt != _rates.end())
                {
                    history = &historyIt->second;
                }

                // During pre-calculation, show calculating info instead of speed
                if (task.preCalcInProgress)
                {
                    const std::wstring sizeText = FormatBytesCompact(task.preCalcTotalBytes);
                    const uint64_t totalItems   = static_cast<uint64_t>(task.preCalcFileCount) + static_cast<uint64_t>(task.preCalcDirectoryCount);
                    const std::wstring countsText =
                        FormatStringResource(nullptr, IDS_FMT_FILEOPS_FILES_FOLDERS, totalItems, task.preCalcFileCount, task.preCalcDirectoryCount);
                    const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(countsText.data(), static_cast<UINT32>(countsText.size()), _bodyFormat.get(), countsRc, _subTextBrush.get());
                    textY += lineH;

                    const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(sizeText.data(), static_cast<UINT32>(sizeText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
                    textY += lineH;
                }
                else if (task.operation == FILESYSTEM_DELETE)
                {
                    const bool hasProgressNumbers = task.completedItems > 0 || task.completedBytes > 0 || task.totalItems > 0 || task.totalBytes > 0;
                    const bool showPreparing      = ! hasProgressNumbers;

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
                        textY += lineH;

                        const bool showSizeProgress = task.preCalcCompleted && task.preCalcTotalBytes > 0 && task.completedBytes > 0;
                        if (showSizeProgress)
                        {
                            const std::wstring sizeProgressText = FormatEmbeddedStringResource(
                                nullptr, IDS_FMT_FILEOPS_SIZE_PROGRESS, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.preCalcTotalBytes));
                            const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                sizeProgressText.data(), static_cast<UINT32>(sizeProgressText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
                            textY += lineH;
                        }
                        else if (task.totalItems > 0)
                        {
                            const std::wstring itemsProgressText = FormatStringResource(nullptr, IDS_FMT_FILEOP_ITEMS, task.completedItems, task.totalItems);
                            const D2D1_RECT_F itemsRc            = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                itemsProgressText.data(), static_cast<UINT32>(itemsProgressText.size()), _bodyFormat.get(), itemsRc, _subTextBrush.get());
                            textY += lineH;
                        }
                        else
                        {
                            const std::wstring itemsProgressText = FormatStringResource(nullptr, IDS_FMT_FILEOP_ITEMS_UNKNOWN_TOTAL, task.completedItems);
                            const D2D1_RECT_F itemsRc            = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                itemsProgressText.data(), static_cast<UINT32>(itemsProgressText.size()), _bodyFormat.get(), itemsRc, _subTextBrush.get());
                            textY += lineH;
                        }
                    }
                }
                else
                {
                    if (task.preCalcSkipped && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE))
                    {
                        const uint64_t completedTotal = static_cast<uint64_t>(task.completedFiles) + static_cast<uint64_t>(task.completedFolders);
                        const bool haveBreakdown      = completedTotal == static_cast<uint64_t>(task.completedItems);
                        if (haveBreakdown && task.completedItems > 0)
                        {
                            const std::wstring countsText =
                                FormatStringResource(nullptr, IDS_FMT_FILEOPS_FILES_FOLDERS, completedTotal, task.completedFiles, task.completedFolders);
                            const D2D1_RECT_F countsRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(countsText.data(), static_cast<UINT32>(countsText.size()), _bodyFormat.get(), countsRc, _subTextBrush.get());
                            textY += lineH;
                        }
                    }

                    const double bytesPerSec          = history ? history->displayedBytesPerSec : 0.0;
                    const uint64_t bytesPerSecRounded = SaturatingRoundNonNegativeToUint64(bytesPerSec);
                    const std::wstring bytesText      = FormatBytesCompact(bytesPerSecRounded);
                    const std::wstring speedText      = FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_BYTES, bytesText);
                    const D2D1_RECT_F speedRc         = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(speedText.data(), static_cast<UINT32>(speedText.size()), _bodyFormat.get(), speedRc, _subTextBrush.get());
                    textY += lineH;

                    // Show size progress (transferred / total) if we have data
                    if (task.totalBytes > 0)
                    {
                        const std::wstring sizeProgressText = FormatEmbeddedStringResource(
                            nullptr, IDS_FMT_FILEOPS_SIZE_PROGRESS, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.totalBytes));
                        const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(
                            sizeProgressText.data(), static_cast<UINT32>(sizeProgressText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
                        textY += lineH;
                    }

                    if (task.totalBytes > 0 && history && history->hasSmoothedEta && history->smoothedEtaSeconds > 0.0 &&
                        task.completedBytes <= task.totalBytes)
                    {
                        const uint64_t seconds     = SaturatingCeilNonNegativeToUint64(history->smoothedEtaSeconds);
                        const std::wstring etaText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA, FormatDurationHms(seconds));
                        const D2D1_RECT_F etaRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(etaText.data(), static_cast<UINT32>(etaText.size()), _bodyFormat.get(), etaRc, _subTextBrush.get());
                        textY += lineH;
                    }
                }

                if (task.autoConcurrencyUsed && task.autoTunedConcurrency > 0u)
                {
                    const std::wstring autoConcurrencyText =
                        FormatStringResource(nullptr, IDS_FMT_FILEOPS_AUTO_CONCURRENCY, task.autoTunedConcurrency, task.effectiveConcurrencyBudget);
                    const D2D1_RECT_F autoConcurrencyRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(
                        autoConcurrencyText.data(), static_cast<UINT32>(autoConcurrencyText.size()), _bodyFormat.get(), autoConcurrencyRc, _subTextBrush.get());
                    textY += lineH;
                }

                const float labelWDesired   = DipsToPixels(56.0f, _dpi);
                const float labelGapDesired = DipsToPixels(6.0f, _dpi);
                const float labelW          = std::min(labelWDesired, textMaxW);
                const float labelGap        = (labelW < textMaxW) ? std::min(labelGapDesired, textMaxW - labelW) : 0.0f;

                if (task.operation == FILESYSTEM_DELETE)
                {
                    const std::wstring label  = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_DELETING);
                    const D2D1_RECT_F labelRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                    _target->DrawTextW(label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get());

                    const float pathW        = std::max(0.0f, textMaxW - labelW - labelGap);
                    const std::wstring path  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), task.currentSourcePath, pathW, lineH);
                    const D2D1_RECT_F pathRc = D2D1::RectF(textX + labelW + labelGap, textY, textX + labelW + labelGap + pathW, textY + lineH);
                    _target->DrawTextW(path.data(), static_cast<UINT32>(path.size()), _bodyFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    textY += lineH;
                }
                else
                {
                    const std::wstring fromLabel = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FROM);
                    const float miniBarGap       = DipsToPixels(8.0f, _dpi);
                    const float miniBarWDesired  = DipsToPixels(92.0f, _dpi);
                    const float miniBarH         = DipsToPixels(6.0f, _dpi);

                    const float pathLeft  = textX + labelW + labelGap;
                    const float rightEdge = textX + textMaxW;

                    const bool showInFlightFiles = task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE;

                    std::array<size_t, TaskSnapshot::kMaxInFlightFiles> activeInFlightIndices{};
                    size_t activeInFlightCount = 0;
                    if (showInFlightFiles)
                    {
                        for (size_t j = 0; j < task.inFlightFileCount && activeInFlightCount < activeInFlightIndices.size(); ++j)
                        {
                            const auto& entry          = task.inFlightFiles[j];
                            const bool active          = entry.totalBytes == 0 || entry.completedBytes < entry.totalBytes;
                            const bool recentCompleted = ! active && entry.totalBytes > 0 && entry.completedBytes >= entry.totalBytes &&
                                                         entry.lastUpdateTick != 0 && nowTick >= entry.lastUpdateTick &&
                                                         (nowTick - entry.lastUpdateTick) <= kCompletedInFlightGraceMs;
                            if (active || recentCompleted)
                            {
                                activeInFlightIndices[activeInFlightCount] = j;
                                ++activeInFlightCount;
                            }
                        }
                    }

                    const size_t inFlightCount = showInFlightFiles ? std::max<size_t>(1u, activeInFlightCount) : 1u;

                    for (size_t i = 0; i < inFlightCount; ++i)
                    {
                        if (i == 0u)
                        {
                            const D2D1_RECT_F fromRc = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                            _target->DrawTextW(fromLabel.data(), static_cast<UINT32>(fromLabel.size()), _smallFormat.get(), fromRc, _subTextBrush.get());
                        }

                        std::wstring_view sourcePathText;
                        uint64_t fileTotalBytes     = 0;
                        uint64_t fileCompletedBytes = 0;

                        const bool hasActiveInFlight = showInFlightFiles && activeInFlightCount > 0;
                        const bool useInFlightEntry  = hasActiveInFlight && i < activeInFlightCount;

                        if (useInFlightEntry)
                        {
                            const auto& entry  = task.inFlightFiles[activeInFlightIndices[i]];
                            sourcePathText     = entry.sourcePath;
                            fileTotalBytes     = entry.totalBytes;
                            fileCompletedBytes = entry.completedBytes;
                        }
                        else
                        {
                            sourcePathText     = task.currentSourcePath;
                            fileTotalBytes     = task.itemTotalBytes;
                            fileCompletedBytes = task.itemCompletedBytes;
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
                        if (! useInFlightEntry && fileTotalBytes > 0 && fileCompletedBytes >= fileTotalBytes)
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

                        if (miniBarW > 0.0f && _progressBgBrush && _progressItemBrush)
                        {
                            const float barTop          = textY + (lineH - miniBarH) * 0.5f;
                            const D2D1_RECT_F miniBarRc = D2D1::RectF(barLeft, barTop, barRight, barTop + miniBarH);

                            const float radiusTrack = ClampCornerRadius(miniBarRc, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(miniBarRc, radiusTrack, radiusTrack), _progressBgBrush.get());

                            const bool hasTotal = fileTotalBytes > 0;
                            const float frac = hasTotal && fileCompletedBytes <= fileTotalBytes
                                                   ? Clamp01(static_cast<float>(static_cast<double>(fileCompletedBytes) / static_cast<double>(fileTotalBytes)))
                                                   : 0.0f;

                            if (theme.menu.rainbowMode)
                            {
                                const D2D1::ColorF rainbow = RainbowProgressColor(theme, sourcePathText);
                                _progressItemBrush->SetColor(rainbow);
                            }
                            else
                            {
                                _progressItemBrush->SetColor(_progressItemBaseColor);
                            }

                            const D2D1_RECT_F fill =
                                hasTotal
                                    ? D2D1::RectF(miniBarRc.left, miniBarRc.top, miniBarRc.left + (miniBarRc.right - miniBarRc.left) * frac, miniBarRc.bottom)
                                    : ComputeIndeterminateBarFill(miniBarRc, nowTick, reducedMotion);
                            const float radiusFill = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radiusFill, radiusFill), _progressItemBrush.get());
                        }

                        textY += lineH;
                    }

                    textY += fromToGapY;

                    const std::wstring toLabel = LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO);
                    const D2D1_RECT_F toRc     = D2D1::RectF(textX, textY, textX + labelW, textY + lineH);
                    _target->DrawTextW(toLabel.data(), static_cast<UINT32>(toLabel.size()), _smallFormat.get(), toRc, _subTextBrush.get());

                    const std::wstring destText = task.destinationFolder.wstring();

                    const float toPathLeft = textX + labelW + labelGap;
                    const float toRight    = textX + textMaxW;

                    const bool canSelectDestination =
                        (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) && ! task.started && task.destinationPane.has_value();

                    float destMenuW         = canSelectDestination ? DipsToPixels(28.0f, _dpi) : 0.0f;
                    const float destMenuGap = (destMenuW > 0.0f) ? DipsToPixels(6.0f, _dpi) : 0.0f;

                    const float minPathW = DipsToPixels(80.0f, _dpi);
                    if (destMenuW > 0.0f && (toRight - toPathLeft) < (minPathW + destMenuGap + destMenuW))
                    {
                        destMenuW = 0.0f;
                    }

                    const float toPathRight    = (destMenuW > 0.0f) ? std::max(toPathLeft, toRight - destMenuW - destMenuGap) : toRight;
                    const float toPathW        = std::max(0.0f, toPathRight - toPathLeft);
                    const std::wstring toPath  = TruncatePathMiddleToWidth(_dwriteFactory.get(), _bodyFormat.get(), destText, toPathW, lineH);
                    const D2D1_RECT_F toPathRc = D2D1::RectF(toPathLeft, textY, toPathLeft + toPathW, textY + lineH);
                    _target->DrawTextW(
                        toPath.data(), static_cast<UINT32>(toPath.size()), _bodyFormat.get(), toPathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                    if (destMenuW > 0.0f)
                    {
                        PopupButton destBtn{};
                        destBtn.bounds     = D2D1::RectF(toRight - destMenuW, textY, toRight, textY + lineH);
                        destBtn.hit.kind   = PopupHitTest::Kind::TaskDestination;
                        destBtn.hit.taskId = task.taskId;
                        _buttons.push_back(destBtn);
                        DrawMenuButton(destBtn, nullptr, {});
                    }
                    textY += lineH;
                }

                const float barInsetX = DipsToPixels(10.0f, _dpi);
                const float barW      = std::max(0.0f, cardRect.right - cardRect.left - barInsetX * 2.0f);
                const float barX      = cardRect.left + barInsetX;

                const float barHItem = DipsToPixels(10.0f, _dpi);

                const bool hasConflictPrompt                    = task.conflict.active;
                const ConflictActionLayout conflictActionLayout = hasConflictPrompt ? BuildConflictActionLayout(task.conflict) : ConflictActionLayout{};

                const float barsHeight    = barHItem;
                const float bottomPadding = DipsToPixels(10.0f, _dpi);
                const float buttonGapY    = DipsToPixels(8.0f, _dpi);
                const float buttonH       = DipsToPixels(24.0f, _dpi);

                const float conflictRowGapY = DipsToPixels(6.0f, _dpi);
                const int conflictRows      = 1;
                const float conflictButtonsHeight =
                    buttonH * static_cast<float>(conflictRows) + conflictRowGapY * static_cast<float>(std::max(0, conflictRows - 1));
                const float conflictApplyLineHeight = hasConflictPrompt ? (lineH + conflictRowGapY) : 0.0f;
                const float buttonsHeight           = conflictButtonsHeight + conflictApplyLineHeight;

                const float buttonRowBottom = cardRect.bottom - bottomPadding;
                const float buttonRowTop    = buttonRowBottom - buttonsHeight;

                const float barsBottom = buttonRowTop - buttonGapY;
                const float barsTop    = barsBottom - barsHeight;

                const auto conflictBucketToMessageId = [&](uint8_t bucket) noexcept -> UINT
                {
                    using Bucket = FolderWindow::FileOperationState::Task::ConflictBucket;
                    switch (static_cast<Bucket>(bucket))
                    {
                        case Bucket::Exists: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_EXISTS);
                        case Bucket::NonEmptyDirectory: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_NONEMPTY_DIRECTORY);
                        case Bucket::ReparsePoint: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_REPARSE_POINT);
                        case Bucket::ReadOnly: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_READONLY);
                        case Bucket::AccessDenied: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_ACCESS_DENIED);
                        case Bucket::SharingViolation: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_SHARING);
                        case Bucket::DiskFull: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_DISK_FULL);
                        case Bucket::PathTooLong: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_PATH_TOO_LONG);
                        case Bucket::RecycleBinFailed: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_RECYCLE_BIN);
                        case Bucket::NetworkOffline: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_NETWORK);
                        case Bucket::UnsupportedReparse: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_UNSUPPORTED_REPARSE);
                        case Bucket::Unknown: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_UNKNOWN);
                        case Bucket::Count: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_UNKNOWN);
                        default: return static_cast<UINT>(IDS_FILEOPS_CONFLICT_UNKNOWN);
                    }
                };

                const auto drawConflictPromptInfo = [&](const D2D1_RECT_F& rc) noexcept
                {
                    if (! _bodyFormat || ! _smallFormat)
                    {
                        return;
                    }

                    float yPrompt           = rc.top;
                    const float maxW        = std::max(0.0f, rc.right - rc.left);
                    const float maxDetailsY = rc.bottom;

                    std::wstring message;
                    const auto bucket = static_cast<FolderWindow::FileOperationState::Task::ConflictBucket>(task.conflict.bucket);
                    if (bucket == FolderWindow::FileOperationState::Task::ConflictBucket::Exists ||
                        bucket == FolderWindow::FileOperationState::Task::ConflictBucket::NonEmptyDirectory ||
                        bucket == FolderWindow::FileOperationState::Task::ConflictBucket::ReparsePoint)
                    {
                        // Name the colliding item: merged-folder child conflicts are otherwise
                        // indistinguishable from a folder-level collision.
                        const std::wstring& conflictPath = ! task.conflict.destinationPath.empty() ? task.conflict.destinationPath : task.conflict.sourcePath;
                        std::wstring_view leaf           = conflictPath;
                        if (const size_t separator = leaf.find_last_of(L"\\/"); separator != std::wstring_view::npos)
                        {
                            leaf.remove_prefix(separator + 1);
                        }
                        if (! leaf.empty())
                        {
                            message = FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFLICT_EXISTS_NAMED, std::wstring(leaf));
                            if (bucket != FolderWindow::FileOperationState::Task::ConflictBucket::Exists)
                            {
                                const std::wstring bucketMessage = LoadStringResource(nullptr, conflictBucketToMessageId(task.conflict.bucket));
                                message                          = std::format(L"{} {}", message, bucketMessage);
                            }
                        }
                    }
                    if (message.empty())
                    {
                        message = LoadStringResource(nullptr, conflictBucketToMessageId(task.conflict.bucket));
                    }
                    if (task.conflict.retryFailed)
                    {
                        const std::wstring retryFailed = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_RETRY_FAILED);
                        message                        = std::format(L"{} {}", retryFailed, message);
                    }

                    if (task.conflict.bucket == static_cast<uint8_t>(FolderWindow::FileOperationState::Task::ConflictBucket::Unknown))
                    {
                        message = std::format(L"{} (0x{:08X})", message, static_cast<unsigned long>(task.conflict.status));
                    }

                    const D2D1_RECT_F msgRc = D2D1::RectF(rc.left, yPrompt, rc.left + maxW, yPrompt + lineH);
                    _target->DrawTextW(
                        message.data(), static_cast<UINT32>(message.size()), _bodyFormat.get(), msgRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    yPrompt += lineH;

                    const auto drawConflictPathLine =
                        [&](std::wstring_view label, std::wstring_view path, const TaskSnapshot::ConflictPromptSnapshot::ItemMetadata& metadata) noexcept
                    {
                        if (label.empty() || path.empty() || ! _dwriteFactory)
                        {
                            return;
                        }

                        if (yPrompt + lineH * 2.0f > maxDetailsY)
                        {
                            return;
                        }

                        const std::wstring metadataText = FormatConflictMetadataText(metadata);
                        const float metadataMaxW        = metadataText.empty() ? 0.0f : std::max(0.0f, maxW * 0.48f);
                        const std::wstring metadataDisplay =
                            metadataText.empty()
                                ? std::wstring{}
                                : TruncateTextMiddleToWidth(_dwriteFactory.get(), _smallFormat.get(), metadataText, metadataMaxW, lineH, kEllipsisText, 0u, 6u);
                        const float metadataW =
                            metadataDisplay.empty()
                                ? 0.0f
                                : std::min(metadataMaxW, MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), metadataDisplay, metadataMaxW, lineH));
                        const float metadataGap = metadataW > 0.0f ? DipsToPixels(8.0f, _dpi) : 0.0f;

                        const D2D1_RECT_F labelRc = D2D1::RectF(rc.left, yPrompt, rc.right - metadataW - metadataGap, yPrompt + lineH);
                        _target->DrawTextW(
                            label.data(), static_cast<UINT32>(label.size()), _smallFormat.get(), labelRc, _subTextBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        if (! metadataDisplay.empty())
                        {
                            const D2D1_RECT_F metadataRc = D2D1::RectF(rc.right - metadataW, yPrompt, rc.right, yPrompt + lineH);
                            _target->DrawTextW(metadataDisplay.data(),
                                               static_cast<UINT32>(metadataDisplay.size()),
                                               _smallFormat.get(),
                                               metadataRc,
                                               _subTextBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        }
                        yPrompt += lineH;

                        const std::wstring truncated = TruncatePathMiddleToWidth(_dwriteFactory.get(), _smallFormat.get(), path, maxW, lineH);
                        const D2D1_RECT_F pathRc     = D2D1::RectF(rc.left, yPrompt, rc.right, yPrompt + lineH);
                        _target->DrawTextW(
                            truncated.data(), static_cast<UINT32>(truncated.size()), _smallFormat.get(), pathRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                        yPrompt += lineH;
                    };

                    if (task.operation == FILESYSTEM_DELETE)
                    {
                        drawConflictPathLine(LoadStringResource(nullptr, IDS_FILEOPS_LABEL_DELETING), task.conflict.sourcePath, task.conflict.sourceMetadata);
                    }
                    else
                    {
                        drawConflictPathLine(LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FROM), task.conflict.sourcePath, task.conflict.sourceMetadata);
                        drawConflictPathLine(
                            LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO), task.conflict.destinationPath, task.conflict.destinationMetadata);
                    }
                };

                if (task.operation != FILESYSTEM_DELETE)
                {
                    const float graphTop    = textY + DipsToPixels(4.0f, _dpi);
                    const float graphBottom = hasConflictPrompt ? barsBottom : (barsTop - DipsToPixels(6.0f, _dpi));
                    const float graphMinH   = DipsToPixels(32.0f, _dpi);

                    if ((graphBottom - graphTop) >= graphMinH)
                    {
                        const D2D1_RECT_F graphRc = D2D1::RectF(barX, graphTop, barX + barW, graphBottom);
                        if (hasConflictPrompt)
                        {
                            drawConflictPromptInfo(graphRc);
                        }
                        else
                        {
                            uint64_t limit = 0;
                            if (task.operation != FILESYSTEM_DELETE)
                            {
                                limit =
                                    task.effectiveSpeedLimitBytesPerSecond != 0 ? task.effectiveSpeedLimitBytesPerSecond : task.desiredSpeedLimitBytesPerSecond;
                            }
                            const RateHistory empty{};
                            const RateHistory& graphHistory = history ? *history : empty;
                            bool showAnimation              = false;
                            const std::wstring overlayText  = GraphOverlayTextForStatus(task, taskStatus, showAnimation);
                            const bool rainbowMode          = folderWindow && folderWindow->GetTheme().menu.rainbowMode;
                            DrawBandwidthGraph(graphRc, graphHistory, limit, overlayText, showAnimation, rainbowMode, true, nowTick, reducedMotion);
                        }
                    }
                }
                else
                {
                    const float graphTop    = textY + DipsToPixels(4.0f, _dpi);
                    const float graphBottom = hasConflictPrompt ? barsBottom : (barsTop - DipsToPixels(6.0f, _dpi));
                    const float graphMinH   = DipsToPixels(32.0f, _dpi);

                    if ((graphBottom - graphTop) >= graphMinH)
                    {
                        const D2D1_RECT_F graphRc = D2D1::RectF(barX, graphTop, barX + barW, graphBottom);
                        if (hasConflictPrompt)
                        {
                            drawConflictPromptInfo(graphRc);
                        }
                        else
                        {
                            const RateHistory empty{};
                            const RateHistory& graphHistory = history ? *history : empty;
                            bool showAnimation              = false;
                            const std::wstring overlayText  = GraphOverlayTextForStatus(task, taskStatus, showAnimation);
                            const bool rainbowMode          = folderWindow && folderWindow->GetTheme().menu.rainbowMode;
                            DrawBandwidthGraph(graphRc, graphHistory, 0, overlayText, showAnimation, rainbowMode, true, nowTick, reducedMotion);
                        }
                    }
                }

                // During pre-calculation, show marquee progress bar
                if (task.preCalcInProgress)
                {
                    const D2D1_RECT_F barRc = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);

                    if (_progressBgBrush)
                    {
                        const float radius = ClampCornerRadius(barRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(barRc, radius, radius), _progressBgBrush.get());
                    }

                    if (_progressItemBrush)
                    {
                        const D2D1_RECT_F fill = ComputeIndeterminateBarFill(barRc, nowTick, reducedMotion);
                        const float radius     = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressItemBrush.get());
                    }
                }
                else if (hasConflictPrompt)
                {
                    // Conflict prompt uses the progress bar area so actions and the apply-to-all toggle sit close together.
                }
                else if (task.operation == FILESYSTEM_DELETE)
                {
                    const D2D1_RECT_F totalBarRc = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);

                    if (_progressBgBrush)
                    {
                        const float radius = ClampCornerRadius(totalBarRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(totalBarRc, radius, radius), _progressBgBrush.get());
                    }

                    if (_progressGlobalBrush)
                    {
                        const bool hasTotalBytes  = task.totalBytes > 0 && task.completedBytes <= task.totalBytes;
                        const bool hasUsefulItems = task.totalItems > 1;

                        const bool useBytes = hasTotalBytes && task.completedBytes > 0;
                        const bool useItems = ! useBytes && hasUsefulItems && task.completedItems > 0;

                        float totalFrac = 0.0f;
                        if (useBytes)
                        {
                            totalFrac = Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
                        }
                        else if (useItems)
                        {
                            const double denom = static_cast<double>(task.totalItems);
                            const double numer = static_cast<double>(std::min(task.completedItems, task.totalItems));
                            totalFrac          = Clamp01(static_cast<float>(numer / denom));
                        }

                        const D2D1_RECT_F fill =
                            (useBytes || useItems)
                                ? D2D1::RectF(
                                      totalBarRc.left, totalBarRc.top, totalBarRc.left + (totalBarRc.right - totalBarRc.left) * totalFrac, totalBarRc.bottom)
                                : ComputeIndeterminateBarFill(totalBarRc, nowTick, reducedMotion);
                        const float radius = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressGlobalBrush.get());
                    }
                }
                else
                {
                    const D2D1_RECT_F totalBarRc = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);

                    if (_progressBgBrush)
                    {
                        const float radiusTotal = ClampCornerRadius(totalBarRc, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(totalBarRc, radiusTotal, radiusTotal), _progressBgBrush.get());
                    }

                    const bool hasItemBytes = task.itemTotalBytes > 0;
                    const float itemFrac =
                        hasItemBytes ? Clamp01(static_cast<float>(static_cast<double>(task.itemCompletedBytes) / static_cast<double>(task.itemTotalBytes)))
                                     : 0.0f;

                    float totalFrac          = 0.0f;
                    bool hasDeterminateTotal = false;
                    if (task.totalBytes > 0 && task.completedBytes <= task.totalBytes)
                    {
                        totalFrac           = Clamp01(static_cast<float>(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes)));
                        hasDeterminateTotal = true;
                    }
                    else if (task.totalItems > 0)
                    {
                        const double denom  = static_cast<double>(task.totalItems);
                        const double numer  = static_cast<double>(std::min(task.completedItems, task.totalItems)) + static_cast<double>(itemFrac);
                        totalFrac           = Clamp01(static_cast<float>(numer / denom));
                        hasDeterminateTotal = true;
                    }

                    if (_progressGlobalBrush)
                    {
                        const D2D1_RECT_F fill =
                            hasDeterminateTotal
                                ? D2D1::RectF(
                                      totalBarRc.left, totalBarRc.top, totalBarRc.left + (totalBarRc.right - totalBarRc.left) * totalFrac, totalBarRc.bottom)
                                : ComputeIndeterminateBarFill(totalBarRc, nowTick, reducedMotion);
                        const float radius = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                        _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radius, radius), _progressGlobalBrush.get());
                    }
                }

                {
                    const float btnGapX = DipsToPixels(8.0f, _dpi);
                    const float rowW    = std::max(0.0f, contentRight - textX);
                    if (rowW > 1.0f)
                    {
                        const float rowTop    = buttonRowTop;
                        const float rowBottom = buttonRowBottom;

                        if (hasConflictPrompt)
                        {
                            // "Apply to all" is placed directly above the conflict action buttons so it's easy to notice and use.
                            const float applyTop    = rowTop;
                            const float applyBottom = applyTop + lineH;
                            const float buttonsTop  = applyBottom + conflictRowGapY;

                            const float checkSize     = DipsToPixels(16.0f, _dpi);
                            const float checkTop      = applyTop + (lineH - checkSize) * 0.5f;
                            const D2D1_RECT_F checkRc = D2D1::RectF(textX, checkTop, textX + checkSize, checkTop + checkSize);
                            DrawCheckboxBox(checkRc, task.conflict.applyToAllChecked);

                            const std::wstring applyText = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_APPLY_TO_ALL_SHORT);
                            const float labelLeft        = textX + checkSize + DipsToPixels(8.0f, _dpi);
                            const D2D1_RECT_F labelRc    = D2D1::RectF(labelLeft, applyTop, contentRight, applyBottom);

                            IDWriteTextFormat* applyFormat = _bodyFormat.get();
                            ID2D1Brush* applyBrush         = _textBrush ? _textBrush.get() : (_subTextBrush ? _subTextBrush.get() : nullptr);
                            if (applyFormat && applyBrush && ! applyText.empty())
                            {
                                _target->DrawTextW(
                                    applyText.data(), static_cast<UINT32>(applyText.size()), applyFormat, labelRc, applyBrush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
                            }

                            PopupButton applyBtn{};
                            applyBtn.bounds     = D2D1::RectF(textX, applyTop, contentRight, applyBottom);
                            applyBtn.hit.kind   = PopupHitTest::Kind::TaskConflictToggleApplyToAll;
                            applyBtn.hit.taskId = task.taskId;
                            _buttons.push_back(applyBtn);

                            const float rowY                  = buttonsTop;
                            const float rowYBottom            = rowY + buttonH;
                            const size_t visibleActionButtons = conflictActionLayout.primaryCount + (conflictActionLayout.overflowCount > 0u ? 1u : 0u);
                            if (visibleActionButtons > 0u)
                            {
                                const float totalGapX = btnGapX * static_cast<float>(visibleActionButtons - 1u);
                                const float btnW      = std::max(0.0f, (rowW - totalGapX) / static_cast<float>(visibleActionButtons));

                                float xBtn = textX;
                                for (size_t i = 0; i < conflictActionLayout.primaryCount && i < conflictActionLayout.primary.size(); ++i)
                                {
                                    const ConflictAction action = conflictActionLayout.primary[i];
                                    const std::wstring label    = ConflictActionText(action);

                                    PopupButton btn{};
                                    btn.bounds     = D2D1::RectF(xBtn, rowY, xBtn + btnW, rowYBottom);
                                    btn.hit.kind   = PopupHitTest::Kind::TaskConflictAction;
                                    btn.hit.taskId = task.taskId;
                                    btn.hit.data   = static_cast<uint32_t>(RawConflictAction(action));
                                    _buttons.push_back(btn);
                                    DrawButton(btn, _buttonSmallFormat.get(), label);

                                    xBtn += btnW + btnGapX;
                                }

                                if (conflictActionLayout.overflowCount > 0u)
                                {
                                    PopupButton moreBtn{};
                                    moreBtn.bounds     = D2D1::RectF(xBtn, rowY, xBtn + btnW, rowYBottom);
                                    moreBtn.hit.kind   = PopupHitTest::Kind::TaskConflictMore;
                                    moreBtn.hit.taskId = task.taskId;
                                    moreBtn.hit.data =
                                        static_cast<uint32_t>(std::min<size_t>(conflictActionLayout.overflowCount, std::numeric_limits<uint32_t>::max()));
                                    _buttons.push_back(moreBtn);
                                    DrawMenuButton(moreBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_MORE));
                                }
                            }
                        }
                        // During copy/move pre-calculation, keep the transfer controls available before bytes start moving.
                        else if (task.preCalcInProgress)
                        {
                            const bool showStartNow     = ! task.started && (task.waitingForOthers || task.waitingInQueue);
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
                                    const float minLimitW = DipsToPixels(140.0f, _dpi);
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
                                DrawMenuButton(limitBtn, _buttonSmallFormat.get(), speedLimitText);
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

                                float skipW   = DipsToPixels(84.0f, _dpi);
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
                                    const float minLimitW = DipsToPixels(140.0f, _dpi);
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
                                DrawMenuButton(limitBtn, _buttonSmallFormat.get(), speedLimitText);
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
                            const bool showStartNow = ! task.started && (task.waitingForOthers || task.waitingInQueue);
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
                                const float minLimitW = DipsToPixels(140.0f, _dpi);
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
                            DrawMenuButton(limitBtn, _buttonSmallFormat.get(), speedLimitText);
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
                            const bool showStartNow = ! task.started && (task.waitingForOthers || task.waitingInQueue);
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
            }

            const float gapAfter = (rowIndex + 1u < rowCount) ? cardGap : 0.0f;
            y += taskCardH + gapAfter;
        }

        _target->PopAxisAlignedClip();
    }
    const uint64_t drawUs = capturePerf ? PerfElapsedUs(drawStartedUs) : 0u;

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

    if (hrEndDraw == D2DERR_RECREATE_TARGET)
    {
        DiscardDeviceResources();
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
    if (! hwnd || ! _taskbarButtonReady || ! EnsureTaskbarList())
    {
        return;
    }

    const auto progressState = static_cast<TBPFLAG>(state);
    if (progressState == TBPF_NOPROGRESS || progressState == TBPF_INDETERMINATE || total == 0)
    {
        static_cast<void>(_taskbarList->SetProgressState(hwnd, progressState));
        return;
    }

    static_cast<void>(_taskbarList->SetProgressValue(hwnd, std::min(completed, total), total));
    static_cast<void>(_taskbarList->SetProgressState(hwnd, progressState));
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
                                    it->hit.kind == PopupHitTest::Kind::FooterDensity || it->hit.kind == PopupHitTest::Kind::FooterToggleDetails;
        if (! isFooterButton && ! PointInRectF(_listViewportRect, x, y))
        {
            continue;
        }

        return it->hit;
    }
    return {};
}

std::optional<FileOperationsPopupInternal::PopupMenuAnchor> FileOperationsPopupInternal::FileOperationsPopupState::ResolveButtonMenuAnchor(
    HWND hwnd, const PopupHitTest& hit, RedSalamander::DxUi::ContextMenuRootVerticalPlacement placement) const noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return std::nullopt;
    }

    const PopupButton* match = nullptr;
    for (auto it = _buttons.rbegin(); it != _buttons.rend(); ++it)
    {
        const PopupButton& button = *it;
        if (button.hit.kind != hit.kind || button.hit.taskId != hit.taskId || button.hit.data != hit.data)
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
        static_cast<LONG>(std::lround(placement == RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Above ? match->bounds.top : match->bounds.bottom)),
    };
    if (ClientToScreen(hwnd, &anchor) == FALSE)
    {
        return std::nullopt;
    }

    PopupMenuAnchor result{};
    result.screenPoint = anchor;
    result.sessionCallbacks.rootHorizontalAlignment =
        alignEnd ? RedSalamander::DxUi::ContextMenuRootHorizontalAlignment::End : RedSalamander::DxUi::ContextMenuRootHorizontalAlignment::Start;
    result.sessionCallbacks.rootVerticalPlacement = placement;
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
    if (! hwnd)
    {
        return false;
    }

    if (! hostLifetime.lock())
    {
        return true;
    }

    if (! fileOps || ! fileOps->HasActiveOperations())
    {
        return true;
    }

    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_FILEOPS_CANCEL_ALL);
    const std::wstring message = LoadStringResource(nullptr, IDS_MSG_FILEOPS_CANCEL_ALL_POPUP);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_INFO;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = hwnd;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_OK;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
    {
        return false;
    }

    if (fileOps)
    {
        fileOps->CancelAll();
    }

    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowSpeedLimitMenu(HWND hwnd, uint64_t taskId) noexcept
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
    if (! task)
    {
        return;
    }

    const FileSystemOperation operation = task->GetOperation();
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return;
    }

    const uint64_t currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);

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

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(kPresets.size() + 4u);

    auto appendRadioItem = [&](UINT commandId, std::wstring text, bool checked) noexcept
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
        item.text      = std::move(text);
        item.commandId = static_cast<int>(commandId);
        item.checked   = checked;
        items.push_back(std::move(item));
    };

    appendRadioItem(kCmdUnlimited, LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_MENU_UNLIMITED), currentLimit == 0);

    RedSalamander::DxUi::MenuFlyoutItem separator{};
    separator.kind = RedSalamander::DxUi::MenuItemKind::Separator;
    items.push_back(separator);

    for (size_t i = 0; i < kPresets.size(); ++i)
    {
        const uint64_t bytesPerSecond = kPresets[i];
        appendRadioItem(kCmdPresetBase + static_cast<UINT>(i),
                        FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_MENU_BYTES, FormatBytesCompact(bytesPerSecond)),
                        currentLimit == bytesPerSecond);
    }

    items.push_back(separator);

    RedSalamander::DxUi::MenuFlyoutItem customItem{};
    customItem.text      = LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_MENU_CUSTOM);
    customItem.commandId = static_cast<int>(kCmdCustom);
    items.push_back(std::move(customItem));

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor = ResolveButtonMenuAnchor(
        hwnd, PopupHitTest{PopupHitTest::Kind::TaskSpeedLimit, taskId, 0u}, RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Below);
    const POINT pt = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(RedSalamander::DxUi::ContextMenu::ShowAsync(hwnd,
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
        if (! selectedTask)
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
            const auto promptResult = ShowCustomSpeedLimitPrompt(hwnd, folderWindow->GetTheme(), current);
            if (! promptResult.has_value())
            {
                return;
            }

            newLimit     = promptResult.value();
            selectedTask = fileOps->FindTask(taskId);
            if (! selectedTask)
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
    if (! fileOps || ! folderWindow)
    {
        return false;
    }

    FolderWindow::FileOperationState::Task* task = requestedTaskId != 0 ? fileOps->FindTask(requestedTaskId) : nullptr;
    if (! task)
    {
        const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
        const auto isActiveFileOperation         = [](const TaskSnapshot& candidate) noexcept
        { return candidate.kind == TaskSnapshot::Kind::FileOperation && ! candidate.finished && candidate.taskId != 0; };

        const auto activeIt = std::find_if(snapshot.begin(), snapshot.end(), isActiveFileOperation);
        if (activeIt != snapshot.end())
        {
            task = fileOps->FindTask(activeIt->taskId);
        }
    }

    if (! task)
    {
        return false;
    }

    const uint64_t currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    const auto promptResult     = ShowCustomSpeedLimitPrompt(hwnd, folderWindow->GetTheme(), currentLimit);
    if (promptResult.has_value())
    {
        task->SetDesiredSpeedLimit(promptResult.value());
    }
    Invalidate(hwnd);
    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowDestinationMenu(HWND hwnd, uint64_t taskId) noexcept
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
    if (! task)
    {
        return;
    }

    const FileSystemOperation operation = task->GetOperation();
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return;
    }

    if (task->HasStarted())
    {
        return;
    }

    const std::optional<FolderWindow::Pane> destinationPaneOpt = task->GetDestinationPane();
    if (! destinationPaneOpt.has_value())
    {
        return;
    }

    const FolderWindow::Pane destinationPane                  = destinationPaneOpt.value();
    const std::optional<std::filesystem::path> otherPanelPath = folderWindow->GetCurrentPluginPath(destinationPane);
    const std::vector<std::filesystem::path> history          = folderWindow->GetFolderHistory(destinationPane);

    constexpr UINT kCmdOtherPanel  = 1u;
    constexpr UINT kCmdHistoryBase = 10u;

    const std::filesystem::path currentDestination = task->GetDestinationFolder();

    NavigationLocation::Location destinationLocation;
    const std::optional<std::filesystem::path> displayDestination = folderWindow->GetCurrentPath(destinationPane);
    if (displayDestination.has_value())
    {
        static_cast<void>(NavigationLocation::TryParseLocation(displayDestination.value().wstring(), destinationLocation));
    }

    struct DestinationEntry
    {
        std::filesystem::path folder;
        std::wstring label;
    };

    std::vector<DestinationEntry> entries;
    entries.reserve(history.size());

    for (const auto& h : history)
    {
        if (h.empty())
        {
            continue;
        }

        NavigationLocation::Location parsed;
        if (! NavigationLocation::TryParseLocation(h.wstring(), parsed))
        {
            continue;
        }

        const bool destIsFile  = NavigationLocation::IsFilePluginShortId(destinationLocation.pluginShortId);
        const bool entryIsFile = NavigationLocation::IsFilePluginShortId(parsed.pluginShortId);
        if (destIsFile != entryIsFile)
        {
            continue;
        }

        if (! destIsFile)
        {
            if (! NavigationLocation::EqualsNoCase(parsed.pluginShortId, destinationLocation.pluginShortId))
            {
                continue;
            }

            if (! NavigationLocation::EqualsNoCase(parsed.instanceContext, destinationLocation.instanceContext))
            {
                continue;
            }
        }

        if (parsed.pluginPath.empty())
        {
            continue;
        }

        DestinationEntry entry{};
        entry.folder = parsed.pluginPath;
        entry.label  = h.wstring();
        entries.push_back(std::move(entry));
    }

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(entries.size() + 2u);

    RedSalamander::DxUi::MenuFlyoutItem otherPanelItem{};
    otherPanelItem.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
    otherPanelItem.text      = LoadStringResource(nullptr, IDS_FILEOP_DEST_OTHER_PANEL);
    otherPanelItem.commandId = static_cast<int>(kCmdOtherPanel);
    otherPanelItem.checked   = otherPanelPath.has_value() && otherPanelPath.value() == currentDestination;
    items.push_back(std::move(otherPanelItem));

    if (! entries.empty())
    {
        RedSalamander::DxUi::MenuFlyoutItem separator{};
        separator.kind = RedSalamander::DxUi::MenuItemKind::Separator;
        items.push_back(separator);
    }

    for (size_t i = 0; i < entries.size(); ++i)
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
        item.text      = entries[i].label;
        item.commandId = static_cast<int>(kCmdHistoryBase + static_cast<UINT>(i));
        item.checked   = entries[i].folder == currentDestination;
        items.push_back(std::move(item));
    }

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor = ResolveButtonMenuAnchor(
        hwnd, PopupHitTest{PopupHitTest::Kind::TaskDestination, taskId, 0u}, RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Below);
    const POINT pt = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(
        RedSalamander::DxUi::ContextMenu::ShowAsync(hwnd,
                                                    pt,
                                                    items,
                                                    MakeAppThemeDxPalette(folderWindow->GetTheme()),
                                                    [this, hwnd, taskId, otherPanelPath, entries = std::move(entries)](std::optional<int> chosenOpt) noexcept
    {
        if (! chosenOpt.has_value())
        {
            return;
        }

        if (! hwnd || IsWindow(hwnd) == FALSE || ! hostLifetime.lock() || ! fileOps)
        {
            return;
        }

        FolderWindow::FileOperationState::Task* selectedTask = fileOps->FindTask(taskId);
        if (! selectedTask)
        {
            return;
        }

        const FileSystemOperation selectedOperation = selectedTask->GetOperation();
        if ((selectedOperation != FILESYSTEM_COPY && selectedOperation != FILESYSTEM_MOVE) || selectedTask->HasStarted())
        {
            return;
        }

        const UINT chosen = static_cast<UINT>(chosenOpt.value());
        if (chosen == kCmdOtherPanel)
        {
            if (otherPanelPath.has_value())
            {
                selectedTask->SetDestinationFolder(otherPanelPath.value());
                Invalidate(hwnd);
            }
            return;
        }

        if (chosen >= kCmdHistoryBase && chosen < (kCmdHistoryBase + static_cast<UINT>(entries.size())))
        {
            const size_t index = static_cast<size_t>(chosen - kCmdHistoryBase);
            selectedTask->SetDestinationFolder(entries[index].folder);
            Invalidate(hwnd);
        }
    },
                                                    sessionCallbacks));
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
    }
    else if (action == kCompletedOverflowActionFailedItems)
    {
        if (! fileOps->IsIssuesPaneVisible())
        {
            fileOps->ToggleIssuesPane();
        }
        handled = fileOps->IsIssuesPaneVisible();
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

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(5u);

    const bool hasDiagnostics = taskIt->warningCount > 0 || taskIt->errorCount > 0;
    if (hasDiagnostics)
    {
        RedSalamander::DxUi::MenuFlyoutItem failedItemsItem{};
        failedItemsItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_FAILED_ITEMS);
        failedItemsItem.commandId = static_cast<int>(kCompletedOverflowActionFailedItems);
        items.push_back(std::move(failedItemsItem));
    }

    if (ResolveCompletedTaskRevealLocation(*taskIt).has_value())
    {
        RedSalamander::DxUi::MenuFlyoutItem revealItem{};
        revealItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_REVEAL_ITEM);
        revealItem.commandId = static_cast<int>(kCompletedOverflowActionRevealDestination);
        items.push_back(std::move(revealItem));
    }

    if (CompletedTaskCanUseDestinationActions(*taskIt))
    {
        RedSalamander::DxUi::MenuFlyoutItem openDestinationItem{};
        openDestinationItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_OPEN_DESTINATION);
        openDestinationItem.commandId = static_cast<int>(kCompletedOverflowActionOpenDestination);
        items.push_back(std::move(openDestinationItem));
    }

    if (hasDiagnostics)
    {
        RedSalamander::DxUi::MenuFlyoutItem showLogItem{};
        showLogItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_SHOW_LOG);
        showLogItem.commandId = static_cast<int>(kCompletedOverflowActionShowLog);
        items.push_back(std::move(showLogItem));

        RedSalamander::DxUi::MenuFlyoutItem exportIssuesItem{};
        exportIssuesItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_EXPORT_ISSUES);
        exportIssuesItem.commandId = static_cast<int>(kCompletedOverflowActionExportIssues);
        items.push_back(std::move(exportIssuesItem));
    }

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor = ResolveButtonMenuAnchor(
        hwnd, PopupHitTest{PopupHitTest::Kind::TaskCompletedMore, taskId, 2u}, RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Above);
    const POINT pt = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(RedSalamander::DxUi::ContextMenu::ShowAsync(hwnd,
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
            chosen == kCompletedOverflowActionOpenDestination || chosen == kCompletedOverflowActionRevealDestination)
        {
            static_cast<void>(SubmitCompletedOverflowAction(hwnd, taskId, chosen, true));
        }
    },
                                                                  sessionCallbacks));
}

bool FileOperationsPopupInternal::FileOperationsPopupState::SubmitConflictOverflowAction(HWND hwnd, uint64_t taskId, uint32_t rawAction) noexcept
{
    if (! fileOps)
    {
        return false;
    }

    FolderWindow::FileOperationState::Task* task = fileOps->FindTask(taskId);
    if (! task)
    {
        return false;
    }

    TaskSnapshot::ConflictPromptSnapshot conflict{};
    bool applyToAll = false;
    {
        std::scoped_lock lock(task->_conflictArbiter.mutex);
        if (! task->_conflictArbiter.prompt.active)
        {
            return false;
        }

        conflict.active            = true;
        conflict.bucket            = static_cast<uint8_t>(task->_conflictArbiter.prompt.bucket);
        conflict.status            = task->_conflictArbiter.prompt.status;
        conflict.sourcePath        = task->_conflictArbiter.prompt.sourcePath;
        conflict.destinationPath   = task->_conflictArbiter.prompt.destinationPath;
        conflict.actionCount       = std::min(task->_conflictArbiter.prompt.actionCount, conflict.actions.size());
        conflict.applyToAllChecked = task->_conflictArbiter.prompt.applyToAllChecked;
        conflict.retryFailed       = task->_conflictArbiter.prompt.retryFailed;
        for (size_t i = 0; i < conflict.actionCount; ++i)
        {
            conflict.actions[i] = RawConflictAction(task->_conflictArbiter.prompt.actions[i]);
        }
        applyToAll = task->_conflictArbiter.prompt.applyToAllChecked;
    }

    const ConflictAction requestedAction = static_cast<ConflictAction>(rawAction);
    const ConflictActionLayout layout    = BuildConflictActionLayout(conflict);
    if (! ConflictLayoutContains(layout.overflow, layout.overflowCount, requestedAction))
    {
        return false;
    }

    task->SubmitConflictDecision(requestedAction, applyToAll);
    Invalidate(hwnd);
    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::ShowConflictOverflowMenu(HWND hwnd, uint64_t taskId) noexcept
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
    if (! task)
    {
        return;
    }

    TaskSnapshot::ConflictPromptSnapshot conflict{};
    {
        std::scoped_lock lock(task->_conflictArbiter.mutex);
        if (! task->_conflictArbiter.prompt.active)
        {
            return;
        }

        conflict.active            = true;
        conflict.bucket            = static_cast<uint8_t>(task->_conflictArbiter.prompt.bucket);
        conflict.status            = task->_conflictArbiter.prompt.status;
        conflict.sourcePath        = task->_conflictArbiter.prompt.sourcePath;
        conflict.destinationPath   = task->_conflictArbiter.prompt.destinationPath;
        conflict.actionCount       = std::min(task->_conflictArbiter.prompt.actionCount, conflict.actions.size());
        conflict.applyToAllChecked = task->_conflictArbiter.prompt.applyToAllChecked;
        conflict.retryFailed       = task->_conflictArbiter.prompt.retryFailed;
        for (size_t i = 0; i < conflict.actionCount; ++i)
        {
            conflict.actions[i] = RawConflictAction(task->_conflictArbiter.prompt.actions[i]);
        }
    }

    const ConflictActionLayout layout = BuildConflictActionLayout(conflict);
    if (layout.overflowCount == 0u)
    {
        return;
    }

    constexpr UINT kCmdOverflowBase = 1u;
    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(std::min(layout.overflowCount, layout.overflow.size()));
    std::vector<uint32_t> overflowActions;
    overflowActions.reserve(std::min(layout.overflowCount, layout.overflow.size()));
    for (size_t i = 0; i < layout.overflowCount && i < layout.overflow.size(); ++i)
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.text      = ConflictActionText(layout.overflow[i]);
        item.commandId = static_cast<int>(kCmdOverflowBase + static_cast<UINT>(i));
        items.push_back(std::move(item));
        overflowActions.push_back(static_cast<uint32_t>(RawConflictAction(layout.overflow[i])));
    }

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const uint32_t overflowData = static_cast<uint32_t>(std::min<size_t>(layout.overflowCount, std::numeric_limits<uint32_t>::max()));
    const auto anchor           = ResolveButtonMenuAnchor(
        hwnd, PopupHitTest{PopupHitTest::Kind::TaskConflictMore, taskId, overflowData}, RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Below);
    const POINT pt = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(
        RedSalamander::DxUi::ContextMenu::ShowAsync(hwnd,
                                                    pt,
                                                    items,
                                                    MakeAppThemeDxPalette(folderWindow->GetTheme()),
                                                    [this, hwnd, taskId, overflowActions = std::move(overflowActions)](std::optional<int> chosenOpt) noexcept
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
        static_cast<void>(SubmitConflictOverflowAction(hwnd, taskId, overflowActions[index]));
    },
                                                    sessionCallbacks));
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    RefreshLocalizedFooterText();

    if (folderWindow && hostLifetime.lock())
    {
        const AppTheme& theme = folderWindow->GetTheme();
        _reducedMotion        = MakeAppThemeDxPalette(theme, theme.windowBackground).reducedMotion;
        ApplyWindowChromeTheme(hwnd, theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd);
    }
    ApplyScrollBarTheme(hwnd);
    ShowScrollBar(hwnd, SB_VERT, FALSE);
    _scrollBarVisible = false;

    UpdateLastPopupRect(hwnd);

    SetTimer(hwnd, kFileOperationsPopupTimerId, kFileOperationsPopupTimerIntervalMs, nullptr);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnThemeChanged(HWND hwnd) noexcept
{
    if (_inThemeChange)
    {
        return 0;
    }

    _inThemeChange      = true;
    auto clearThemeFlag = wil::scope_exit([&] { _inThemeChange = false; });

    RefreshLocalizedFooterText();
    DiscardDeviceResources();

    if (folderWindow && hostLifetime.lock())
    {
        const AppTheme& theme = folderWindow->GetTheme();
        _reducedMotion        = MakeAppThemeDxPalette(theme, theme.windowBackground).reducedMotion;
        ApplyWindowChromeTheme(hwnd, theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd);
    }
    ApplyScrollBarTheme(hwnd);

    RedrawWindow(hwnd, nullptr, nullptr, RDW_FRAME | RDW_NOERASE | RDW_NOCHILDREN);
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnNcDestroy(HWND hwnd) noexcept
{
    KillTimer(hwnd, kFileOperationsPopupTimerId);
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

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnSize(HWND hwnd, UINT width, UINT height) noexcept
{
    _clientSize.cx = static_cast<LONG>(width);
    _clientSize.cy = static_cast<LONG>(height);

    if (_target)
    {
        _target->Resize(D2D1::SizeU(width, height));
    }

    UpdateLastPopupRect(hwnd);
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT& suggested) noexcept
{
    _dpi = newDpi;

    _headerFormat.reset();
    _bodyFormat.reset();
    _smallFormat.reset();
    _buttonFormat.reset();
    _buttonSmallFormat.reset();
    _graphOverlayFormat.reset();
    _statusIconFormat.reset();
    _statusIconFallbackFormat.reset();
    _captionGlyphFormat.reset();
    _captionGlyphFallbackFormat.reset();
    _captionGlyphDpi = 0;

    if (_target)
    {
        _target->SetDpi(96.0f, 96.0f);
    }

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

    constexpr int kMinClientWidthDip = 480;
    const int minClientHeightDip =
        (fileOps && fileOps->GetPopupFooterOnly()) ? kFileOperationsPopupFooterOnlyMinClientHeightDip : kFileOperationsPopupMinClientHeightDip;

    const int minClientW = DipsToPixels(kMinClientWidthDip, dpiForWindow);
    const int minClientH = DipsToPixels(minClientHeightDip, dpiForWindow);

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
        if (! hostLifetime.lock())
        {
            DestroyWindow(hwnd);
            return 0;
        }

        if (! IsWindowVisible(hwnd) || IsIconic(hwnd))
        {
            UpdateTaskbarProgress(hwnd);
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
    _inSizeMove = true;
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
    _inSizeMove = false;
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnVScroll(HWND hwnd, UINT request) noexcept
{
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

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMouseMove(HWND hwnd, POINT pt) noexcept
{
    if (! _trackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = hwnd;
        TrackMouseEvent(&tme);
        _trackingMouse = true;
    }

    const PopupHitTest hit = HitTest(static_cast<float>(pt.x), static_cast<float>(pt.y));
    if (hit.kind != _hotHit.kind || hit.taskId != _hotHit.taskId || hit.data != _hotHit.data)
    {
        _hotHit = hit;
        Invalidate(hwnd);
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMouseLeave(HWND hwnd) noexcept
{
    _trackingMouse = false;
    if (_hotHit.kind != PopupHitTest::Kind::None)
    {
        _hotHit = {};
        Invalidate(hwnd);
    }
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnLButtonDown(HWND hwnd, POINT pt) noexcept
{
    SetCapture(hwnd);
    _pressedHit = HitTest(static_cast<float>(pt.x), static_cast<float>(pt.y));
    _hotHit     = _pressedHit;
    Invalidate(hwnd);
    return 0;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnLButtonUp(HWND hwnd, POINT pt) noexcept
{
    ReleaseCapture();

    const PopupHitTest released = HitTest(static_cast<float>(pt.x), static_cast<float>(pt.y));
    const bool activated        = _pressedHit.kind != PopupHitTest::Kind::None && _pressedHit.kind == released.kind && _pressedHit.taskId == released.taskId &&
                                  _pressedHit.data == released.data;
    const PopupHitTest hit      = _pressedHit;
    _pressedHit                 = {};

    if (! activated)
    {
        return 0;
    }

    return OnActivatedHit(hwnd, hit);
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnActivatedHit(HWND hwnd, const PopupHitTest& hit) noexcept
{
    if (! hostLifetime.lock())
    {
        DestroyWindow(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::FooterCancelAll)
    {
        if (fileOps && ! fileOps->HasActiveOperations())
        {
            std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
            fileOps->CollectCompletedTasks(completed);
            for (const auto& summary : completed)
            {
                fileOps->DismissCompletedTask(summary.taskId);
            }
            Invalidate(hwnd);
            return 0;
        }

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
        }
        Invalidate(hwnd);
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
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::CompletedGroupToggle)
    {
        _completedGroupExpanded     = ! _completedGroupExpanded;
        _maxAutoSizedWindowHeight   = 0;
        _lastAutoSizedContentHeight = -1.0f;
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::CompletedGroupClear)
    {
        if (fileOps)
        {
            std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
            fileOps->CollectCompletedTasks(completed);
            for (const auto& summary : completed)
            {
                fileOps->DismissCompletedTask(summary.taskId);
            }
        }
        _maxAutoSizedWindowHeight   = 0;
        _lastAutoSizedContentHeight = -1.0f;
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskToggleCollapse)
    {
        ToggleTaskCollapsed(hit.taskId, fileOps ? fileOps->GetPopupCompactDensity() : false);
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

    if (hit.kind == PopupHitTest::Kind::TaskCancel)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                task->RequestCancel();
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
                task->SkipPreCalculation();
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

    if (hit.kind == PopupHitTest::Kind::TaskDestination)
    {
        ShowDestinationMenu(hwnd, hit.taskId);
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
                task->ToggleConflictApplyToAllChecked();
            }
        }
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskConflictMore)
    {
        ShowConflictOverflowMenu(hwnd, hit.taskId);
        Invalidate(hwnd);
        return 0;
    }

    if (hit.kind == PopupHitTest::Kind::TaskConflictAction)
    {
        if (fileOps)
        {
            FolderWindow::FileOperationState::Task* task = fileOps->FindTask(hit.taskId);
            if (task)
            {
                bool applyToAll = false;
                {
                    std::scoped_lock lock(task->_conflictArbiter.mutex);
                    applyToAll = task->_conflictArbiter.prompt.applyToAllChecked;
                }

                const auto action = static_cast<FolderWindow::FileOperationState::Task::ConflictAction>(hit.data);
                task->SubmitConflictDecision(action, applyToAll);
            }
        }
        Invalidate(hwnd);
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

    if (payload->kind == PopupHitTest::Kind::None && payload->data == kPopupSelfTestDestroyOnNextShowData)
    {
        _destroyOnNextShowForSelfTest = true;
        return 1;
    }

    if (payload->kind == PopupHitTest::Kind::TaskSpeedLimit && payload->data == 1u)
    {
        // Let self-test callers advance their state before the modal prompt loop starts.
        return PostMessageW(hwnd, kFileOperationsPopupDeferredSpeedLimitPromptMessage, 0, static_cast<LPARAM>(payload->taskId)) != FALSE ? 1 : 0;
    }

    if (payload->kind == PopupHitTest::Kind::TaskConflictMore && payload->data != 0u)
    {
        return SubmitConflictOverflowAction(hwnd, payload->taskId, payload->data) ? 1 : 0;
    }

    if (payload->kind == PopupHitTest::Kind::TaskCompletedMore && payload->data != 0u)
    {
        return SubmitCompletedOverflowAction(hwnd, payload->taskId, payload->data, false) ? 1 : 0;
    }

    if (payload->kind == PopupHitTest::Kind::TaskConflictToggleApplyToAll || payload->kind == PopupHitTest::Kind::TaskConflictAction ||
        payload->kind == PopupHitTest::Kind::FooterToggleDetails || payload->kind == PopupHitTest::Kind::FooterQueueMode ||
        payload->kind == PopupHitTest::Kind::FooterDensity || payload->kind == PopupHitTest::Kind::FooterAutoDismiss ||
        payload->kind == PopupHitTest::Kind::FooterPauseResumeAll || payload->kind == PopupHitTest::Kind::CompletedGroupToggle ||
        payload->kind == PopupHitTest::Kind::CompletedGroupClear || payload->kind == PopupHitTest::Kind::TaskToggleCollapse ||
        payload->kind == PopupHitTest::Kind::TaskStartNow)
    {
        // OnActivatedHit returns 0 for these kinds even on success; self-test callers need a
        // dispatched-successfully signal.
        static_cast<void>(OnActivatedHit(hwnd, PopupHitTest{payload->kind, payload->taskId, payload->data}));
        return 1;
    }

    return OnActivatedHit(hwnd, PopupHitTest{payload->kind, payload->taskId, payload->data});
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
    result.taskId = snapshot->taskId;

    const auto rectHasArea = [](const D2D1_RECT_F& rc) noexcept { return rc.right > rc.left && rc.bottom > rc.top; };

    const std::vector<TaskSnapshot> taskSnapshot     = BuildSnapshot();
    GlobalFileOperationsStatusSummary globalSummary  = BuildGlobalStatusSummary(taskSnapshot, &_rates);
    const GlobalTaskbarProgressModel taskbarProgress = BuildGlobalTaskbarProgressModel(globalSummary);
    const AppTheme appTheme                          = folderWindow ? folderWindow->GetTheme() : AppTheme{};
    const bool highContrast                          = appTheme.highContrast;
    const bool reducedMotion                         = IsReducedMotionEnabled();
    const uint32_t completedGroupCount               = CountCompletedGroupTasks(taskSnapshot);
    const bool completedGroupVisible                 = ShouldShowCompletedGroup(completedGroupCount);
    result.globalRunningCount                        = globalSummary.running;
    result.globalWaitingCount                        = globalSummary.waiting;
    result.globalNeedAttentionCount                  = globalSummary.needAttention;
    result.globalSummaryText                         = FormatGlobalStatusSummaryText(globalSummary);
    result.globalSummaryVisible                      = ! taskSnapshot.empty() && rectHasArea(_footerSummaryRect);
    result.footerPauseResumeAllVisible               = rectHasArea(_footerPauseResumeAllRect) && HasFooterPauseResumeAllControl(globalSummary);
    result.footerPauseResumeAllPauses                = FooterPauseResumeAllShouldPause(globalSummary);
    result.footerQueueModeSegmentedVisible           = rectHasArea(_footerQueueModeRect);
    result.footerQueueModeIsParallel                 = fileOps ? ! fileOps->GetQueueNewTasks() : false;
    result.footerQueueSegmentRect                    = _footerQueueSegmentRect;
    result.footerParallelSegmentRect                 = _footerParallelSegmentRect;
    result.footerSummaryRect                         = _footerSummaryRect;
    if (rectHasArea(_footerQueueSegmentRect))
    {
        const float hitX                  = (_footerQueueSegmentRect.left + _footerQueueSegmentRect.right) * 0.5f;
        const float hitY                  = (_footerQueueSegmentRect.top + _footerQueueSegmentRect.bottom) * 0.5f;
        const PopupHitTest hit            = HitTest(hitX, hitY);
        result.footerQueueHitTargetActive = hit.kind == PopupHitTest::Kind::FooterQueueMode && hit.data == kFooterQueueModeQueueAction;
    }
    if (rectHasArea(_footerParallelSegmentRect))
    {
        const float hitX                     = (_footerParallelSegmentRect.left + _footerParallelSegmentRect.right) * 0.5f;
        const float hitY                     = (_footerParallelSegmentRect.top + _footerParallelSegmentRect.bottom) * 0.5f;
        const PopupHitTest hit               = HitTest(hitX, hitY);
        result.footerParallelHitTargetActive = hit.kind == PopupHitTest::Kind::FooterQueueMode && hit.data == kFooterQueueModeParallelAction;
    }
    result.footerAutoDismissVisible      = rectHasArea(_footerAutoDismissRect);
    result.footerAutoDismissLabelVisible = _footerAutoDismissLabelVisible;
    result.footerAutoDismissEnabled      = fileOps ? fileOps->GetAutoDismissSuccess() : false;
    result.footerDensityToggleVisible    = rectHasArea(_footerDensityRect);
    if (result.footerDensityToggleVisible)
    {
        const float hitX                    = (_footerDensityRect.left + _footerDensityRect.right) * 0.5f;
        const float hitY                    = (_footerDensityRect.top + _footerDensityRect.bottom) * 0.5f;
        result.footerDensityHitTargetActive = HitTest(hitX, hitY).kind == PopupHitTest::Kind::FooterDensity;
    }
    result.popupCompactDensity                = fileOps ? fileOps->GetPopupCompactDensity() : false;
    result.footerAggregateProgressVisible     = HasGlobalAggregateProgress(globalSummary);
    result.footerAggregateProgressDeterminate = HasDeterminateGlobalAggregateProgress(globalSummary);
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
    result.highContrastEnabled             = highContrast;
    result.reducedMotionEnabled            = reducedMotion;
    result.autoResizeAnimationEnabled      = ! reducedMotion;
    result.footerQueueModeAnimationEnabled = result.footerQueueModeSegmentedVisible && ! reducedMotion;
    result.completedGroupVisible           = completedGroupVisible;
    result.completedGroupExpanded          = completedGroupVisible && _completedGroupExpanded;
    result.completedGroupCount             = completedGroupVisible ? completedGroupCount : 0u;
    result.completedGroupVisibleTaskCount  = result.completedGroupExpanded ? completedGroupCount : 0u;
    result.completedGroupAnimationEnabled  = false;

    for (const TaskSnapshot& task : taskSnapshot)
    {
        if (task.finished && IsTaskCollapsed(task.taskId))
        {
            ++result.completedAutoCollapsedCount;
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
        result.taskId                     = taskIt->taskId;
        result.found                      = true;
        const TaskStatusKind status       = taskIt->statusKind != TaskStatusKind::None ? taskIt->statusKind : ResolveTaskStatusKind(*taskIt);
        result.taskStatusKind             = status;
        result.taskStatusActiveStateCount = SurfacedTaskStatusCount(status);
        bool graphShowAnimation           = false;
        static_cast<void>(GraphOverlayTextForStatus(*taskIt, status, graphShowAnimation));
        result.graphStatusAnimationEnabled        = graphShowAnimation && ! reducedMotion;
        result.conflictStackedPathRows            = taskIt->conflict.active;
        result.conflictSourceMetadataVisible      = taskIt->conflict.active && taskIt->conflict.sourceMetadata.available;
        result.conflictDestinationMetadataVisible = taskIt->conflict.active && taskIt->conflict.destinationMetadata.available;
        result.conflictMetadataSizeCompareVisible =
            taskIt->conflict.active && taskIt->conflict.sourceMetadata.sizeKnown && taskIt->conflict.destinationMetadata.sizeKnown;
        result.conflictMetadataDateCompareVisible =
            taskIt->conflict.active && taskIt->conflict.sourceMetadata.lastWriteTime > 0 && taskIt->conflict.destinationMetadata.lastWriteTime > 0;
        result.taskDuplicateUnderGraphItemBarVisible = false;
        const PopupStatusVisualTone statusTone       = StatusVisualToneForTaskStatus(status);
        const std::wstring statusChipText            = StatusChipTextForTask(*taskIt, status, GetTickCount64());
        result.taskStatusVisualTone                  = static_cast<uint32_t>(statusTone);
        result.taskStatusVisualColorRef              = ColorToCOLORREF(StatusVisualColorForTone(appTheme, statusTone));
        result.taskStatusStripeVisible               = statusTone != PopupStatusVisualTone::None;
        result.taskStatusChipVisible                 = result.taskStatusStripeVisible && ! statusChipText.empty();
        result.taskStatusGlyphSignalVisible          = StatusIsOk(status) || StatusIsWarning(status) || StatusIsError(status);
        result.taskStatusTextSignalVisible           = ! statusChipText.empty() || ! StatusTextForTask(*taskIt, status, GetTickCount64()).empty();
        result.taskStatusColorBlindSafeEncoding =
            statusTone == PopupStatusVisualTone::None || result.taskStatusGlyphSignalVisible || result.taskStatusTextSignalVisible;
        result.completedFailedItemsActionVisible       = taskIt->finished && (taskIt->warningCount > 0 || taskIt->errorCount > 0);
        result.completedOpenDestinationActionVisible   = CompletedTaskCanUseDestinationActions(*taskIt);
        result.completedRevealDestinationActionVisible = ResolveCompletedTaskRevealLocation(*taskIt).has_value();
        result.taskHiddenByCompletedGroup              = completedGroupVisible && ! _completedGroupExpanded && IsCompletedGroupTask(*taskIt);
        result.taskCollapsed                           = IsTaskCollapsed(taskIt->taskId);
        result.taskCompactRow                          = IsTaskCollapsedForDisplay(taskIt->taskId, result.popupCompactDensity);
        result.taskAutoCollapsedOnCompletion           = taskIt->finished && result.taskCollapsed;
        result.taskCompactProgressVisible = result.taskCompactRow && taskIt->kind == TaskSnapshot::Kind::FileOperation && TaskHasKnownCompactProgress(*taskIt);
        if (! taskIt->finished && ! taskIt->conflict.active &&
            (taskIt->preCalcInProgress || taskIt->operation == FILESYSTEM_COPY || taskIt->operation == FILESYSTEM_MOVE ||
             taskIt->operation == FILESYSTEM_DELETE))
        {
            result.taskUnderGraphProgressBarCount = 1u;
        }

        const ConflictActionLayout conflictLayout = BuildConflictActionLayout(taskIt->conflict);
        result.conflictOverflowActionCount        = conflictLayout.overflowCount;
        for (size_t i = 0; i < conflictLayout.overflowCount && i < conflictLayout.overflow.size() && i < result.conflictOverflowActions.size(); ++i)
        {
            result.conflictOverflowActions[i] = RawConflictAction(conflictLayout.overflow[i]);
        }

        // Aggregate the live graph hue distribution so the fairness selftest can assert on the
        // REAL pipeline (samples, eviction, multi-bucket flushes) instead of synthetic weights.
        if (const auto rateIt = _rates.find(result.taskId); rateIt != _rates.end())
        {
            PopulateGraphHueDebugSummary(rateIt->second, result);
            result.graphCurrentBandwidthBytesPerSecond = CurrentBandwidthForGraphMarker(rateIt->second);
            result.graphCurrentBandwidthLineVisible    = result.graphCurrentBandwidthBytesPerSecond > 0.0;
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

    result.visibleButtonCount = _buttons.size();
    for (size_t i = 0; i < _buttons.size(); ++i)
    {
        const PopupButton& button = _buttons[i];
        switch (button.hit.kind)
        {
            case PopupHitTest::Kind::FooterCancelAll:
            case PopupHitTest::Kind::FooterPauseResumeAll:
            case PopupHitTest::Kind::FooterAutoDismiss:
            case PopupHitTest::Kind::FooterQueueMode:
            case PopupHitTest::Kind::FooterDensity:
            case PopupHitTest::Kind::FooterToggleDetails: ++result.footerVisibleButtonCount; break;
            case PopupHitTest::Kind::CompletedGroupToggle: result.completedGroupToggleVisible = true; break;
            case PopupHitTest::Kind::CompletedGroupClear: result.completedGroupClearVisible = true; break;
            case PopupHitTest::Kind::TaskConflictToggleApplyToAll:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.conflictApplyToAllVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskConflictAction:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    appendPrimaryAction(static_cast<uint8_t>(button.hit.data));
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
                    result.completedDiagnosticsMoreVisible           = true;
                    result.completedDiagnosticsMoreButtonRectVisible = true;
                    result.completedDiagnosticsMoreButtonRect        = button.bounds;
                    ++result.completedVisibleActionCount;
                }
                break;
            case PopupHitTest::Kind::TaskDismiss:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.completedDismissVisible = true;
                    ++result.completedVisibleActionCount;
                }
                break;
            case PopupHitTest::Kind::TaskToggleCollapse:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskToggleCollapseVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskStartNow:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskStartNowVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskPause:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskPauseVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskCancel:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskCancelVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskSkip:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskSkipVisible = true;
                }
                break;
            case PopupHitTest::Kind::TaskSpeedLimit:
                if (result.found && button.hit.taskId == result.taskId)
                {
                    result.taskSpeedLimitVisible = true;
                }
                break;
            case PopupHitTest::Kind::None:
            case PopupHitTest::Kind::TaskDestination:
            default: break;
        }

        for (size_t j = i + 1u; j < _buttons.size(); ++j)
        {
            if (rectsOverlap(button.bounds, _buttons[j].bounds))
            {
                result.hasVisibleButtonOverlap = true;
                if (result.found && button.hit.taskId == result.taskId && _buttons[j].hit.taskId == result.taskId)
                {
                    result.taskHasVisibleButtonOverlap = true;
                }
            }
        }
    }

    *snapshot = result;
    return result.found ? 1 : 0;
}
#endif

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnMouseWheel(HWND hwnd, int delta) noexcept
{
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

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnClose(HWND hwnd) noexcept
{
    if (ConfirmCancelAll(hwnd))
    {
        DestroyWindow(hwnd);
    }
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
    const UINT taskbarButtonCreatedMessage = FileOperationsTaskbarButtonCreatedMessage();
    if (taskbarButtonCreatedMessage != 0 && msg == taskbarButtonCreatedMessage)
    {
        _taskbarButtonReady        = true;
        _taskbarListRetryAfterTick = 0;
        _taskbarListAttemptCount   = 0u;
        _taskbarList.reset();
        Invalidate(hwnd);
        return 0;
    }

    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd);
#ifdef ENABLE_TESTS
        case WM_SHOWWINDOW:
            if (wp != FALSE && _destroyOnNextShowForSelfTest)
            {
                _destroyOnNextShowForSelfTest = false;
                DestroyWindow(hwnd);
                return 0;
            }
            break;
#endif
        case WM_NCDESTROY: return OnNcDestroy(hwnd);
        case WM_NCACTIVATE: return OnNcActivate(hwnd, wp, lp);
        case WM_NCPAINT: return OnNcPaint(hwnd, wp, lp);
        case WM_ERASEBKGND: return 1;
        case WM_PAINT:
        {
            const bool capturePerf   = Debug::Perf::IsCaptureEnabled();
            const uint64_t startedUs = capturePerf ? PerfNowUs() : 0u;
            Render(hwnd);
            if (capturePerf)
            {
                Debug::Perf::Emit(L"FileOps.Popup.WmPaintUs", L"", PerfElapsedUs(startedUs), 0u, 0u, S_OK);
            }
            return 0;
        }
        case WM_SIZE: return OnSize(hwnd, LOWORD(lp), HIWORD(lp));
        case WM_MOVE: return OnMove(hwnd);
        case WM_GETMINMAXINFO: return OnGetMinMaxInfo(hwnd, reinterpret_cast<MINMAXINFO*>(lp));
        case WM_ENTERSIZEMOVE: return OnEnterSizeMove(hwnd);
        case WM_EXITSIZEMOVE: return OnExitSizeMove(hwnd);
        case WM_TIMER: return OnTimer(hwnd, static_cast<UINT_PTR>(wp));
        case WM_VSCROLL: return OnVScroll(hwnd, static_cast<UINT>(LOWORD(wp)));
        case WM_MOUSEMOVE: return OnMouseMove(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_MOUSELEAVE: return OnMouseLeave(hwnd);
        case WM_LBUTTONDOWN: return OnLButtonDown(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_LBUTTONUP: return OnLButtonUp(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_MOUSEWHEEL: return OnMouseWheel(hwnd, GET_WHEEL_DELTA_WPARAM(wp));
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

    RedrawWindow(popup, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOCHILDREN);

    FileOperationsPopupInternal::PopupLayoutDebugSnapshot snapshot{};
    snapshot.taskId = out.taskId;
    const bool ok   = SendMessageW(popup, WndMsg::kFileOpsPopupLayoutSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
    if (ok && snapshot.found)
    {
        out = snapshot;
    }
    return ok && snapshot.found;
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

bool DebugBuildFileOperationsGraphFairColorWeightSnapshot(FileOperationsPopupInternal::GraphHueWeightDebugSnapshot& out) noexcept
{
    out = {};

    RateHistory history{};
    AddPendingHueWeight(history, 10.0f, 100.0);
    AddPendingHueWeight(history, 100.0f, 100.0);
    AddPendingHueWeight(history, 190.0f, 100.0);
    AddPendingHueWeight(history, 280.0f, 100.0);
    AppendResampledRateSamples(history, kRateSampleBucketMs, 400.0, -1.0f);

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
    for (uint64_t bucket = 1; bucket <= 12; ++bucket)
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
            stream.totalBytes       = 4096u;
            stream.completedBytes   = bucket * 100u;
            stream.lastUpdateTick   = task.lastProgressCallbackTick;
        }

        AccumulateStreamHueWeights(history, task, 400u, -1.0f);
        AppendResampledRateSamples(history, kRateSampleBucketMs, 400.0, -1.0f);
    }

    out.taskId = 1;
    out.found  = true;
    PopulateGraphHueDebugSummary(history, out);
    out.graphCurrentBandwidthBytesPerSecond = CurrentBandwidthForGraphMarker(history);
    out.graphCurrentBandwidthLineVisible    = out.graphCurrentBandwidthBytesPerSecond > 0.0;
    return out.graphMultiHueBucketCount >= 10u && out.graphDistinctHueCount == 4u && out.graphMinHueShare >= 0.20 && out.graphMaxHueShare <= 0.30;
}

float DebugComputeFileOperationsTaskCompleteFraction(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return ComputeFileOperationsTaskCompleteFractionForDisplay(task);
}

void DebugPublishFileOperationsPlannedItemTotalAfterPreCalculation(FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    PublishPlannedItemTotalAfterPreCalculation(task);
}

bool DebugFileOperationsTaskHasKnownCompactProgress(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return TaskHasKnownCompactProgress(task);
}

std::wstring DebugFormatFileOperationsConflictTimestamp(__int64 fileTime) noexcept
{
    return FormatFileTimeLocalCompact(fileTime);
}

double DebugSmoothRateForDisplay(double previousRate, double sampleRate, ULONGLONG elapsedMs) noexcept
{
    return SmoothRateForDisplay(previousRate, sampleRate, elapsedMs);
}

double DebugDecayRateForCallbackSilence(double smoothedRate, ULONGLONG silenceMs) noexcept
{
    return DecayRateForCallbackSilence(smoothedRate, silenceMs);
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
