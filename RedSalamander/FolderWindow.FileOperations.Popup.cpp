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
#include "Win32CallbackHelpers.h"
#include "WindowSizing.h"

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
constexpr wchar_t kFileOperationsPopupClassName[] = L"RedSalamander.FileOperationsPopup";

constexpr UINT_PTR kFileOperationsPopupTimerId                      = 1;
constexpr UINT kFileOperationsPopupTimerIntervalMs                  = 100;
constexpr UINT kFileOperationsPopupUiMotionTimerIntervalMs          = 16;
constexpr ULONGLONG kFileOperationsPopupUiMotionDurationMs          = 260ull;
constexpr ULONGLONG kFileOperationsPopupUiMotionTimerWindowMs       = 520ull;
constexpr ULONGLONG kRateSampleBucketMs                             = 100ull;
constexpr UINT kFileOperationsPopupDeferredSpeedLimitPromptMessage  = WM_APP + 0x71;
constexpr wchar_t kFileOperationsSpeedLimitPromptClassName[]        = L"RedSalamander.FileOperations.SpeedLimitPrompt";
constexpr UINT kFileOperationsSpeedLimitPromptDeferredActionMessage = WM_APP + 0x74;
constexpr std::wstring_view kEllipsisText                           = L"\u2026";
constexpr float kFileOperationsPopupFooterHeightDip                 = 88.0f;
constexpr float kFileOperationsExpandedCardHeightDip                = 280.0f;
constexpr float kFileOperationsExpandedTransferCardHeightDip        = 244.0f;
constexpr int kFileOperationsPopupMinClientHeightDip                = 320;
constexpr int kFileOperationsPopupFooterOnlyMinClientHeightDip      = 96;
constexpr uint32_t kCompletedOverflowActionShowLog                  = 1u;
constexpr uint32_t kCompletedOverflowActionExportIssues             = 2u;
constexpr uint32_t kCompletedOverflowActionFailedItems              = 3u;
constexpr uint32_t kCompletedOverflowActionOpenDestination          = 4u;
constexpr uint32_t kCompletedOverflowActionRevealDestination        = 5u;
constexpr uint32_t kCompletedOverflowActionOpenSource               = 6u;
constexpr uint32_t kCompletedOverflowActionSelectRetained           = 7u;
constexpr uint32_t kCompletedOverflowActionCutRetainedAgain         = 8u;
constexpr uint32_t kFooterPauseResumeAllPauseAction                 = 1u;
constexpr uint32_t kFooterPauseResumeAllResumeAction                = 2u;
constexpr uint32_t kFooterQueueModeQueueAction                      = 1u;
constexpr uint32_t kFooterQueueModeParallelAction                   = 2u;
constexpr ULONGLONG kTaskbarListRetryDelayMs                        = 1000ull;

[[nodiscard]] bool OpensButtonFlyout(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
{
    using Kind = FileOperationsPopupInternal::PopupHitTest::Kind;
    return kind == Kind::FooterQueueMode || kind == Kind::TaskDestination || kind == Kind::TaskSpeedLimit || kind == Kind::TaskCompletedMore ||
           kind == Kind::TaskConflictMore;
}

[[nodiscard]] bool UsesSelectorButtonChrome(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
{
    using Kind = FileOperationsPopupInternal::PopupHitTest::Kind;
    return kind == Kind::FooterQueueMode || kind == Kind::TaskSpeedLimit;
}

[[nodiscard]] RedSalamander::DxUi::ButtonVariant HostedButtonVariant(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
{
    using Kind = FileOperationsPopupInternal::PopupHitTest::Kind;
    if (kind == Kind::FooterToggleDetails || kind == Kind::CompletedGroupToggle || kind == Kind::TaskToggleCollapse)
    {
        return RedSalamander::DxUi::ButtonVariant::Disclosure;
    }
    if (UsesSelectorButtonChrome(kind))
    {
        return RedSalamander::DxUi::ButtonVariant::Selector;
    }
    if (OpensButtonFlyout(kind))
    {
        return RedSalamander::DxUi::ButtonVariant::DropDown;
    }
    if (kind == FileOperationsPopupInternal::PopupHitTest::Kind::FooterOptions)
    {
        return RedSalamander::DxUi::ButtonVariant::IconOnly;
    }
    return RedSalamander::DxUi::ButtonVariant::Standard;
}

[[nodiscard]] std::wstring FormatSpeedLimitSelectorText(uint64_t bytesPerSecond)
{
    return bytesPerSecond == 0 ? LoadStringResource(nullptr, IDS_FILEOP_SPEED_LIMIT_BUTTON_UNLIMITED)
                               : FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_LIMIT_BUTTON_BYTES, FormatBytesCompact(bytesPerSecond));
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

    if (task.discoveryClosed && task.totalBytes > 0)
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
    if (task.discoveryClosed && task.totalItems > 0)
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
    return task.warningCount > 0 || task.errorCount > 0 || CompletedTaskCanUseDestinationActions(task) ||
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

[[nodiscard]] UINT ConflictBucketMessageId(uint8_t rawBucket) noexcept
{
    switch (static_cast<ConflictBucket>(rawBucket))
    {
        case ConflictBucket::RegularFileExists: return IDS_FILEOPS_CONFLICT_EXISTS;
        case ConflictBucket::ReadOnlyRegularFile:
        case ConflictBucket::ReadOnlyRegularFileExists: return IDS_FILEOPS_CONFLICT_READONLY;
        case ConflictBucket::TypeMismatch: return IDS_FILEOPS_CONFLICT_TYPE_MISMATCH;
        case ConflictBucket::DestinationLink: return IDS_FILEOPS_CONFLICT_REPARSE_POINT;
        case ConflictBucket::NameNotRepresentable: return IDS_FILEOPS_CONFLICT_NAME_NOT_REPRESENTABLE;
        case ConflictBucket::TargetConflict: return IDS_FILEOPS_CONFLICT_TARGET;
        case ConflictBucket::AccessDenied: return IDS_FILEOPS_CONFLICT_ACCESS_DENIED;
        case ConflictBucket::SharingViolation: return IDS_FILEOPS_CONFLICT_SHARING;
        case ConflictBucket::DiskFull: return IDS_FILEOPS_CONFLICT_DISK_FULL;
        case ConflictBucket::PathTooLong: return IDS_FILEOPS_CONFLICT_PATH_TOO_LONG;
        case ConflictBucket::RecycleFailed: return IDS_FILEOPS_CONFLICT_RECYCLE_BIN;
        case ConflictBucket::InsufficientSpace: return IDS_FILEOPS_CONSENT_LOW_SPACE;
        case ConflictBucket::SpaceUnknown: return IDS_FILEOPS_CONSENT_SPACE_UNKNOWN;
        case ConflictBucket::EfsPlaintext: return IDS_FILEOPS_CONSENT_EFS_PLAINTEXT;
        case ConflictBucket::SparseInflation: return IDS_FILEOPS_CONSENT_SPARSE_INFLATION;
        case ConflictBucket::PlaceholderHydration: return IDS_FILEOPS_CONSENT_PLACEHOLDER_HYDRATION;
        case ConflictBucket::MetadataLoss: return IDS_FILEOPS_CONSENT_METADATA_LOSS;
        case ConflictBucket::SameHostOverlap: return IDS_FILEOPS_CONSENT_SAME_HOST_OVERLAP;
        case ConflictBucket::SameHostLiveOutput: return IDS_FILEOPS_CONSENT_SAME_HOST_LIVE_OUTPUT;
        case ConflictBucket::PermanentDeleteConfirmation: return IDS_FILEOPS_CONSENT_PERMANENT_DELETE;
        case ConflictBucket::NetworkOffline: return IDS_FILEOPS_CONFLICT_NETWORK;
        case ConflictBucket::UnsupportedReparse: return IDS_FILEOPS_CONFLICT_UNSUPPORTED_REPARSE;
        case ConflictBucket::Unknown:
        case ConflictBucket::Count:
        default: return IDS_FILEOPS_CONFLICT_UNKNOWN;
    }
}

[[nodiscard]] std::wstring BuildConflictQuestion(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    const auto bucket = static_cast<ConflictBucket>(task.conflict.bucket);
    std::wstring question;
    if (bucket == ConflictBucket::SameHostOverlap)
    {
        // R4-A02-1: the concrete race, named after the first conflicting task.
        const unsigned long otherTask = static_cast<unsigned long>(task.conflict.overlapTaskId);
        switch (static_cast<FileOperations::SameHostOverlapProblem>(task.conflict.overlapProblem))
        {
            case FileOperations::SameHostOverlapProblem::SameNames: return FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_SAME_NAMES, otherTask);
            case FileOperations::SameHostOverlapProblem::SameMembers: return FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_SAME_MEMBERS, otherTask);
            case FileOperations::SameHostOverlapProblem::RemovesPublished:
                return FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_REMOVES_PUBLISHED, otherTask);
            case FileOperations::SameHostOverlapProblem::ReplacesRead: return FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_REPLACES_READ, otherTask);
            case FileOperations::SameHostOverlapProblem::RemovesRead:
            case FileOperations::SameHostOverlapProblem::None:
            default: return FormatStringResource(nullptr, IDS_FMT_FILEOPS_OVERLAP_REMOVES_READ, otherTask);
        }
    }
    if (bucket == ConflictBucket::SameHostLiveOutput)
    {
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_LIVE_OUTPUT_QUESTION, static_cast<unsigned long>(task.conflict.overlapTaskId));
    }
    if (bucket == ConflictBucket::PermanentDeleteConfirmation)
    {
        // C1: the confirmation that used to be a modal at ingress, in the same words, on the card.
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE, task.conflict.consentDetail, task.conflict.sourcePath);
    }
    if (bucket == ConflictBucket::RegularFileExists || bucket == ConflictBucket::ReadOnlyRegularFileExists || bucket == ConflictBucket::TypeMismatch ||
        bucket == ConflictBucket::DestinationLink)
    {
        const std::wstring& conflictPath = ! task.conflict.destinationPath.empty() ? task.conflict.destinationPath : task.conflict.sourcePath;
        std::wstring_view leaf           = conflictPath;
        if (const size_t separator = leaf.find_last_of(L"\\/"); separator != std::wstring_view::npos)
        {
            leaf.remove_prefix(separator + 1u);
        }
        if (! leaf.empty())
        {
            question = FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFLICT_EXISTS_NAMED, std::wstring(leaf));
            if (bucket != ConflictBucket::RegularFileExists)
            {
                question.append(L" ");
                question.append(LoadStringResource(nullptr, ConflictBucketMessageId(task.conflict.bucket)));
            }
        }
    }
    if (question.empty())
    {
        question = LoadStringResource(nullptr, ConflictBucketMessageId(task.conflict.bucket));
    }
    if (task.conflict.attemptCount > 0u)
    {
        // C2: the human attempt text; Retry stays offered while the failure is a proved no-commit.
        question = std::format(L"{0} {1}",
                               FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFLICT_ATTEMPT_FAILED, static_cast<unsigned long>(task.conflict.attemptCount)),
                               question);
    }
    if (bucket == ConflictBucket::Unknown)
    {
        question = std::format(L"{0} (0x{1:08X})", question, static_cast<unsigned long>(task.conflict.status));
    }
    return question;
}

[[nodiscard]] std::wstring BuildConflictFactText(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    if (task.conflict.factItemCountKnown && task.conflict.factBytesKnown)
    {
        return FormatStringResource(
            nullptr, IDS_FMT_FILEOPS_CONSENT_FACTS_ITEMS_BYTES, task.conflict.factItemCount, FormatBytesCompact(task.conflict.factBytes));
    }
    if (task.conflict.factItemCountKnown)
    {
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONSENT_FACTS_ITEMS, task.conflict.factItemCount);
    }
    if (task.conflict.factBytesKnown)
    {
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONSENT_FACTS_BYTES, FormatBytesCompact(task.conflict.factBytes));
    }
    return {};
}

struct ConflictPromptContent final
{
    std::wstring state;
    std::wstring question;
    std::wstring facts;
    std::wstring sourceLabel;
    std::wstring sourcePath;
    std::wstring sourceMetadata;
    std::wstring destinationLabel;
    std::wstring destinationPath;
    std::wstring destinationMetadata;
    std::wstring lastNote;
};

[[nodiscard]] ConflictPromptContent BuildConflictPromptContent(const FileOperationsPopupInternal::TaskSnapshot& task)
{
    ConflictPromptContent content{};
    if (! task.conflict.active)
    {
        return content;
    }

    content.state = LoadStringResource(nullptr, task.conflict.metadataLoading ? IDS_FILEOPS_CONFLICT_READING_DETAILS : IDS_FILEOPS_STATUS_WAITING_FOR_DECISION);
    if (! task.conflict.metadataLoading)
    {
        content.question = BuildConflictQuestion(task);
    }
    content.facts          = BuildConflictFactText(task);
    content.sourceLabel    = LoadStringResource(nullptr, task.operation == FILESYSTEM_DELETE ? IDS_FILEOPS_LABEL_DELETING : IDS_FILEOPS_LABEL_FROM);
    content.sourcePath     = task.conflict.sourcePath;
    content.sourceMetadata = FormatConflictMetadataText(task.conflict.sourceMetadata);
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
                                                       IDWriteTextFormat* bodyFormat,
                                                       IDWriteTextFormat* smallFormat,
                                                       const ConflictPromptContent& content,
                                                       float maxWidth,
                                                       float lineHeight,
                                                       float rowGap) noexcept
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
    const auto addContext = [&](std::wstring_view label, std::wstring_view path) noexcept
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
        height += MeasureWrappedTextHeight(factory, smallFormat, path, maxWidth, lineHeight);
    };

    addWrapped(content.state, bodyFormat);
    addWrapped(content.question, bodyFormat);
    addWrapped(content.facts, smallFormat);
    addContext(content.sourceLabel, content.sourcePath);
    addContext(content.destinationLabel, content.destinationPath);
    addWrapped(content.lastNote, smallFormat);
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

    copyActions(source.actions, source.actionCount, destination.actions, destination.actionCount);
    copyActions(source.primaryActions, source.primaryActionCount, destination.primaryActions, destination.primaryActionCount);
    copyActions(source.overflowActions, source.overflowActionCount, destination.overflowActions, destination.overflowActionCount);
    destination.defaultAction      = RawConflictAction(source.defaultAction);
    destination.escapeAction       = RawConflictAction(source.escapeAction);
    destination.applyToAllEligible = source.applyToAllEligible;
    destination.skipAllEligible    = source.skipAllEligible;
    destination.buttonsPublishable = source.buttonsPublishable;
}

[[nodiscard]] bool PublishedConflictPlacementContains(const std::array<uint8_t, FileOperationsPopupInternal::TaskSnapshot::kMaxConflictActions>& actions,
                                                      size_t actionCount,
                                                      ConflictAction action) noexcept
{
    for (size_t i = 0; i < actionCount && i < actions.size(); ++i)
    {
        if (ConflictActionFromRaw(actions[i]) == action)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring ConflictActionText(ConflictAction action,
                                              ConflictBucket bucket = ConflictBucket::Unknown,
                                              FileSystemOperation operation = FILESYSTEM_COPY) noexcept
{
    switch (action)
    {
        case ConflictAction::Overwrite: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_OVERWRITE);
        case ConflictAction::ReplaceReadOnly: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_REPLACE_READONLY);
        case ConflictAction::ReplaceLink: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_REPLACE_LINK);
        case ConflictAction::PermanentDelete: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_PERMANENT_DELETE);
        case ConflictAction::Proceed:
            switch (bucket)
            {
                case ConflictBucket::EfsPlaintext: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_PLAINTEXT);
                case ConflictBucket::PlaceholderHydration: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_DOWNLOAD);
                case ConflictBucket::MetadataLoss: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_MOVE_ANYWAY);
                case ConflictBucket::SameHostOverlap: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_QUEUE_AFTER);
                case ConflictBucket::RegularFileExists:
                case ConflictBucket::ReadOnlyRegularFile:
                case ConflictBucket::ReadOnlyRegularFileExists:
                case ConflictBucket::TypeMismatch:
                case ConflictBucket::DestinationLink:
                case ConflictBucket::NameNotRepresentable:
                case ConflictBucket::TargetConflict:
                case ConflictBucket::SharingViolation:
                case ConflictBucket::AccessDenied:
                case ConflictBucket::DiskFull:
                case ConflictBucket::PathTooLong:
                case ConflictBucket::RecycleFailed:
                case ConflictBucket::InsufficientSpace:
                case ConflictBucket::SpaceUnknown:
                case ConflictBucket::SparseInflation:
                case ConflictBucket::SameHostLiveOutput:
                case ConflictBucket::NetworkOffline:
                case ConflictBucket::UnsupportedReparse:
                case ConflictBucket::PermanentDeleteConfirmation:
                case ConflictBucket::Unknown:
                case ConflictBucket::Count: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_CONTINUE);
            }
        case ConflictAction::RetainSource: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_KEEP_SOURCE);
        case ConflictAction::RunConcurrently: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_RUN_CONCURRENT);
        case ConflictAction::QueueUntilOtherTaskFinishes: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_QUEUE_UNTIL_FINISHED);
        case ConflictAction::InvalidateLiveOutput:
            switch (operation)
            {
                case FILESYSTEM_DELETE: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_DELETE_LIVE_OUTPUT);
                case FILESYSTEM_RENAME: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_RENAME_LIVE_OUTPUT);
                case FILESYSTEM_MOVE: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_MOVE_ANYWAY);
                case FILESYSTEM_CREATE_DIRECTORY:
                case FILESYSTEM_COPY:
                default: return LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BTN_REPLACE_LIVE_OUTPUT);
            }
        case ConflictAction::Retry: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_RETRY);
        case ConflictAction::KeepBoth: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_KEEP_BOTH);
        case ConflictAction::Skip: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_SKIP);
        case ConflictAction::SkipAll: return LoadStringResource(nullptr, IDS_FILEOPS_BTN_SKIP_ALL);
        case ConflictAction::Cancel:
            return LoadStringResource(nullptr, bucket == ConflictBucket::SameHostOverlap ? IDS_FILEOPS_CONSENT_BTN_DONT_START : IDS_FILEOP_BTN_CANCEL);
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
    uint32_t openDiscoveryTasks        = 0u;
    // A closed task with no total at all: nothing is known about it and it will not close later,
    // so it keeps the aggregate indeterminate (unlike an open task, which only adds to the count).
    bool hasClosedUnknownTotals        = false;
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

[[nodiscard]] bool TaskHasKnownCompactProgress(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.discoveryClosed && (task.totalBytes > 0 || task.totalItems > 0);
}

struct WholeTaskProgressPresentation final
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

[[nodiscard]] WholeTaskProgressPresentation ResolveWholeTaskProgressPresentation(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    WholeTaskProgressPresentation result{};
    // operationStartTick is published when ExecuteOperation is underway. It makes the startup
    // marquee visible before the first transfer callback without mistaking discovery-only work
    // for an active transfer.
    result.transferActive           = task.started && task.operationStartTick != 0u;
    result.provisional              = ! task.discoveryClosed;
    const bool waitingForAdmission  = task.queuePaused || task.waitingInQueue || task.waitingForOthers;
    result.includeInAggregateCohort = task.started && ! task.finished && ! waitingForAdmission && ! task.conflict.active;
    result.showDiscoveryActivity    = result.includeInAggregateCohort && ! task.discoveryClosed && (task.discoveryAheadActive || task.discoverySkipped);
    result.showTransferProgress     = result.includeInAggregateCohort;
    result.showThroughputGraph      = result.includeInAggregateCohort && result.transferActive;

    // D2-A08: no growing-denominator fraction. Until the total closes the bar is indeterminate (a
    // marquee once transfer is active) and the card shows exact counters instead.
    if (task.discoveryClosed)
    {
        if (task.totalBytes > 0 && task.completedBytes <= task.totalBytes)
        {
            result.fraction    = std::clamp(static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes), 0.0, 1.0);
            result.determinate = true;
            return result;
        }
        if (task.totalItems > 0 && task.completedItems <= task.totalItems)
        {
            result.fraction    = std::clamp(static_cast<double>(task.completedItems) / static_cast<double>(task.totalItems), 0.0, 1.0);
            result.determinate = true;
        }
    }
    result.showTransferMarquee = result.showTransferProgress && ! result.determinate && result.transferActive;
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
    return phase == FileOperations::TaskLifecyclePhase::Preparing ||
           phase == FileOperations::TaskLifecyclePhase::AwaitingAcceptance;
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
        if (task.hasIndeterminateResult || task.resultHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE))
        {
            return TaskStatusKind::Indeterminate;
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
    return status == TaskStatusKind::Conflict || status == TaskStatusKind::Partial || status == TaskStatusKind::Indeterminate ||
           status == TaskStatusKind::Failed;
}

[[nodiscard]] bool StatusCountsAsRunning(TaskStatusKind status) noexcept
{
    return status == TaskStatusKind::Discovering || status == TaskStatusKind::Preparing || status == TaskStatusKind::Running ||
           status == TaskStatusKind::Verifying || status == TaskStatusKind::Paused;
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
        case TaskStatusKind::Done: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
        case TaskStatusKind::Partial: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_PARTIAL);
        case TaskStatusKind::Indeterminate: return LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
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
        case TaskStatusKind::Verifying: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_VERIFYING);
        case TaskStatusKind::Waiting: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_WAITING);
        case TaskStatusKind::Partial: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_PARTIAL_SHORT);
        case TaskStatusKind::Indeterminate: return LoadStringResource(nullptr, IDS_FILEOPS_RESULT_UNKNOWN);
        case TaskStatusKind::Failed: return LoadStringResource(nullptr, IDS_FILEOPS_STATUS_FAILED_SHORT);
        case TaskStatusKind::Paused:
        case TaskStatusKind::Conflict:
        case TaskStatusKind::Discovering:
        case TaskStatusKind::Preparing:
        case TaskStatusKind::Done:
        case TaskStatusKind::Canceled: return StatusTextForTask(task, status, nowTick);
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

[[nodiscard]] bool TaskHasTerminalAxisBadges(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return TaskHasVerificationResultBadge(task) ||
           (task.finished && task.operation == FILESYSTEM_MOVE && (task.retainedSourceCount > 0u || task.unknownSourceCount > 0u));
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
    if (task.operation == FILESYSTEM_MOVE && task.retainedSourceCount > 0u)
    {
        const UINT wording = task.retainedSourceNativeCount == task.retainedSourceCount ? IDS_FILEOPS_RESULT_MOVED_SOURCE_FOLDER_KEPT
                                                                                        : IDS_FILEOPS_RESULT_COPIED_SOURCE_KEPT;
        badges[count++] = TerminalAxisBadge{LoadStringResource(nullptr, wording), PopupStatusVisualTone::Warning, L"SourceRetained"};
    }
    return count;
}

[[nodiscard]] std::wstring BuildTaskHeaderText(const FileOperationsPopupInternal::TaskSnapshot& task, std::wstring_view operationText, ULONGLONG nowTick)
{
    const TaskStatusKind status = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
    if (status == TaskStatusKind::Running || status == TaskStatusKind::Verifying || status == TaskStatusKind::None)
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

        const TaskStatusKind status                      = task.statusKind != TaskStatusKind::None ? task.statusKind : ResolveTaskStatusKind(task);
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

        if (presentation.includeInAggregateCohort)
        {
            summary.hasActiveTaskbarProgress = true;
        }

        const auto accumulateSaturated = [](uint64_t value, uint64_t& aggregate) noexcept
        {
            const uint64_t remaining = (std::numeric_limits<uint64_t>::max)() - aggregate;
            aggregate                = value > remaining ? (std::numeric_limits<uint64_t>::max)() : aggregate + value;
        };

        if (presentation.includeInAggregateCohort && task.discoveryClosed && task.totalBytes > 0)
        {
            accumulateSaturated(task.totalBytes, summary.totalBytes);
            accumulateSaturated((std::min)(task.completedBytes, task.totalBytes), summary.completedBytes);
        }
        if (presentation.includeInAggregateCohort && task.discoveryClosed && task.totalItems > 0)
        {
            accumulateSaturated(task.totalItems, summary.totalItems);
            accumulateSaturated((std::min)(task.completedItems, task.totalItems), summary.completedItems);
        }
        if (presentation.includeInAggregateCohort && (! task.discoveryClosed || (task.totalBytes == 0 && task.totalItems == 0)))
        {
            summary.hasUnknownActiveProgress = true;
        }
        if (presentation.includeInAggregateCohort && ! task.discoveryClosed)
        {
            ++summary.openDiscoveryTasks;
        }
        if (presentation.includeInAggregateCohort && task.discoveryClosed && task.totalBytes == 0 && task.totalItems == 0)
        {
            summary.hasClosedUnknownTotals = true;
        }
        if (presentation.includeInAggregateCohort && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) &&
            (! task.discoveryClosed || task.totalBytes == 0))
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
        if (! summary.hasUnknownActiveTransferBytes && etaInputsComplete && transferEtaAvailable && verificationEtaAvailable &&
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

[[nodiscard]] bool IsCompletedGroupTask(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept
{
    return task.kind == FileOperationsPopupInternal::TaskSnapshot::Kind::FileOperation && task.finished;
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
    statusParts.reserve(5u);
    if (summary.running > 0u)
    {
        statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_RUNNING_COUNT, static_cast<unsigned long>(summary.running)));
    }
    if (summary.waiting > 0u)
    {
        statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_WAITING_COUNT, static_cast<unsigned long>(summary.waiting)));
    }
    if (summary.needAttention > 0u)
    {
        statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_ATTENTION_COUNT, static_cast<unsigned long>(summary.needAttention)));
    }
    if (summary.openDiscoveryTasks > 0u)
    {
        if (HasDeterminateGlobalAggregateProgress(summary))
        {
            const unsigned long knownWorkPercent =
                static_cast<unsigned long>(std::clamp(std::lround(static_cast<double>(GlobalAggregateProgressFraction(summary)) * 100.0), 0L, 100L));
            statusParts.push_back(FormatStringResource(nullptr, IDS_FMT_FILEOPS_KNOWN_WORK, knownWorkPercent));
        }
        statusParts.push_back(
            FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_DISCOVERING_COUNT, static_cast<unsigned long>(summary.openDiscoveryTasks)));
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
        model.completed = summary.completedBytes;
        model.total     = summary.totalBytes;
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
        return RedSalamander::DxUi::ThroughputGraphColorFromHue(hue, theme.dark);
    }
    const D2D1::ColorF accent = theme.navigationView.accent;
    return D2D1_COLOR_F{accent.r, accent.g, accent.b, accent.a};
}

bool IsRateSamplingBlocked(const FileOperationsPopupInternal::RateSnapshot& task) noexcept
{
    return task.paused || task.queuePaused || task.waitingInQueue;
}

struct RateByteDeltas final
{
    uint64_t transfer         = 0u;
    uint64_t verificationRead = 0u;
};

[[nodiscard]] RateByteDeltas ComputeRateByteDeltas(const FileOperationsPopupInternal::RateSnapshot& task,
                                                   const FileOperationsPopupInternal::RateHistory& history) noexcept
{
    return RateByteDeltas{
        .transfer = task.completedBytes - (std::min)(history.lastBytes, task.completedBytes),
        // Logical provider-proof completion advances the verification bar but is deliberately
        // absent here: only physical exact-object readback contributes a graph/rate sample.
        .verificationRead = task.verificationReadBytes - (std::min)(history.lastVerificationBytes, task.verificationReadBytes),
    };
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
    return RedSalamander::DxUi::ShouldRenderThroughputGraphBands(theme.menu.rainbowMode, true, theme.highContrast, maximumConcurrentColorSlots);
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
        entry.assignedColorSlot = RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot;
    }
}

[[nodiscard]] uint8_t AssignFirstFreeStreamColorSlot(std::array<bool, FileOperationsPopupInternal::RateHistory::kMaxHueWeightsPerSample>& occupied) noexcept
{
    const auto available = std::ranges::find(occupied, false);
    return available == occupied.end() ? RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot
                                       : static_cast<uint8_t>(std::distance(occupied.begin(), available));
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
        uint8_t streamColorSlot = RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot;
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
    const double displayed = ClampFiniteNonNegative(history.displayedBytesPerSec) + ClampFiniteNonNegative(history.displayedVerificationBytesPerSec);
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

void AppendRateSample(FileOperationsPopupInternal::RateHistory& history, double sample, double verificationSample, float hue) noexcept
{
    const double clampedSample = ClampFiniteNonNegative(sample);
    const size_t slot          = history.writeIndex;

    history.samples[slot]             = static_cast<float>(std::min<double>(clampedSample, std::numeric_limits<float>::max()));
    history.verificationSamples[slot] = static_cast<float>(std::min<double>(ClampFiniteNonNegative(verificationSample), std::numeric_limits<float>::max()));

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
                    .colorSlot = RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot,
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
    history.pendingHue                          = -1.0f;
    history.pendingHueWeightCount               = 0u;
}

void AppendResampledRateSamples(
    FileOperationsPopupInternal::RateHistory& history, ULONGLONG elapsedMs, double sample, double verificationSample, float hue) noexcept
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
        AppendRateSample(history, bucketSample, bucketVerificationSample, bucketHue);
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

float FileOperationsPopupInternal::FileOperationsPopupState::ResolveAnimatedTaskHeight(uint64_t taskId,
                                                                                       float targetHeight,
                                                                                       ULONGLONG tick,
                                                                                       bool reducedMotion) noexcept
{
    constexpr ULONGLONG kDurationMs = kFileOperationsPopupUiMotionDurationMs;
    ScalarAnimation& animation      = _taskHeightAnimations[taskId];
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
        const float eased       = RedSalamander::DxUi::EvaluateEasing(RedSalamander::DxUi::EasingCurve::PointToPoint, linear);
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

    if (! _graphTrailingLabelFormat)
    {
        static_cast<void>(RedSalamander::DxUi::Typography::CreateTextFormat(
            _dwriteFactory.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(DipsToPixels(11.0f, _dpi)), _graphTrailingLabelFormat.put(), L""));
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
#ifdef ENABLE_TESTS
    unsigned int remainingForcedFailures = g_fileOperationsD2DTargetForcedFailures.load(std::memory_order_acquire);
    while (remainingForcedFailures > 0u)
    {
        if (g_fileOperationsD2DTargetForcedFailures.compare_exchange_weak(
                remainingForcedFailures, remainingForcedFailures - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            return; // the next paint enters the failure surface
        }
    }
#endif

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
    const D2D1::ColorF progressVerify = theme.fileOperations.progressVerify;
    _progressItemBaseColor            = progressItem;

    const D2D1::ColorF okAccent    = theme.fileOperations.successText;
    const D2D1::ColorF warningText = theme.folderView.warningText;
    const D2D1::ColorF errorText   = theme.folderView.errorText;

    const D2D1::ColorF graphLine   = theme.fileOperations.graphLine;

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
            snap.inFlightFileCount = std::min(task->_inFlightFileCount, snap.inFlightFiles.size());
            for (size_t i = 0; i < snap.inFlightFileCount; ++i)
            {
                snap.inFlightFiles[i].cookieKey        = task->_inFlightFiles[i].cookieKey;
                snap.inFlightFiles[i].progressStreamId = task->_inFlightFiles[i].progressStreamId;
                snap.inFlightFiles[i].sourcePath       = task->_inFlightFiles[i].sourcePath;
                snap.inFlightFiles[i].totalBytes       = task->_inFlightFiles[i].totalBytes;
                snap.inFlightFiles[i].completedBytes   = task->_inFlightFiles[i].completedBytes;

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
            snap.conflict.metadataLoading                   = task->_conflictArbiter.prompt.metadataLoading;
            snap.conflict.deferredConsent                   = task->_conflictArbiter.prompt.deferredConsent;
            snap.conflict.bucket                            = static_cast<uint8_t>(task->_conflictArbiter.prompt.bucket);
            snap.conflict.status                            = task->_conflictArbiter.prompt.status;
            snap.conflict.sourcePath                        = task->_conflictArbiter.prompt.sourcePath;
            snap.conflict.destinationPath                   = task->_conflictArbiter.prompt.destinationPath;
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
        snap.plannedItems               = task->GetPlannedItemCount();
        snap.destinationFolder          = task->GetDestinationFolder();
        snap.destinationPane            = task->GetDestinationPane();
        snap.destinationPluginId        = task->_destinationPluginId;
        snap.destinationPluginShortId   = task->_destinationPluginShortId;
        snap.destinationInstanceContext = task->_destinationInstanceContext;
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
                const auto& completed             = *completedIt->second;
                snap.warningCount                 = completed.warningCount;
                snap.errorCount                   = completed.errorCount;
                snap.lastDiagnosticMessage        = completed.lastDiagnosticMessage;
                snap.resultSummary                = completed.resultSummary;
                snap.verifiedItemCount            = completed.verifiedItemCount;
                snap.verificationProblemItemCount = completed.verificationProblemItemCount;
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
                snap.hasSourceActionPaths               = ! completed.retainedSourcePaths.empty() || ! completed.unknownSourcePaths.empty();
                snap.retainedSourceCount                = completed.retainedSourceCount;
                snap.retainedSourceNativeCount          = completed.retainedSourceNativeCount;
                snap.unknownSourceCount                 = completed.unknownSourceCount;
                snap.exactRetainedSourceActionAvailable = completed.clipboardMoveAdmission &&
                                                          NavigationLocation::IsFilePluginShortId(completed.sourcePluginShortId) &&
                                                          completed.retainedSourceCount > 0u && completed.unknownSourceCount == 0u &&
                                                          completed.exactRetainedSourceItems.size() == completed.retainedSourceCount;
            }
            else
            {
                fileOps->CollectTaskDiagnosticSnapshot(snap.taskId, snap.warningCount, snap.errorCount, snap.lastDiagnosticMessage);
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
        snap.discoveryClosed                    = task->_discoveryClosed.load(std::memory_order_acquire);
        snap.firstMutationBeforeDiscoveryClosed = task->_firstMutationBeforeDiscoveryClosed.load(std::memory_order_acquire);
        snap.discoveredTotalBytes               = task->_discoveredTotalBytes.load(std::memory_order_acquire);
        snap.discoveredFileCount                = task->_discoveredFileCount.load(std::memory_order_acquire);
        snap.discoveredDirectoryCount           = task->_discoveredDirectoryCount.load(std::memory_order_acquire);

        const ULONGLONG startTick = task->_discoveryStartTick.load(std::memory_order_acquire);
        if (! snap.discoveryClosed && startTick > 0)
        {
            const ULONGLONG nowTick = GetTickCount64();
            snap.discoveryElapsedMs = (nowTick >= startTick) ? (nowTick - startTick) : 0;
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
        snap.resultSummary                         = completed.resultSummary;
        snap.verificationRequested                 = completed.verifiedItemCount > 0u || completed.verificationProblemItemCount > 0u;
        snap.verifiedItemCount                     = completed.verifiedItemCount;
        snap.verificationProblemItemCount          = completed.verificationProblemItemCount;
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
        snap.hasSourceActionPaths   = ! completed.retainedSourcePaths.empty() || ! completed.unknownSourcePaths.empty();
        snap.retainedSourceCount    = completed.retainedSourceCount;
        snap.retainedSourceNativeCount = completed.retainedSourceNativeCount;
        snap.unknownSourceCount     = completed.unknownSourceCount;
        snap.exactRetainedSourceActionAvailable = completed.clipboardMoveAdmission && NavigationLocation::IsFilePluginShortId(completed.sourcePluginShortId) &&
                                                  completed.retainedSourceCount > 0u && completed.unknownSourceCount == 0u &&
                                                  completed.exactRetainedSourceItems.size() == completed.retainedSourceCount;
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
        snap.discoveredTotalBytes       = task->_discoveredTotalBytes.load(std::memory_order_acquire);
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
        snap.discoveryClosed         = task->_discoveryClosed.load(std::memory_order_acquire);

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
        snap.discoveryClosed          = completed.discoveryClosed;
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
            history.lastVerificationBytes    = task.verificationReadBytes;
            history.lastItems                = task.completedItems;
            history.lastProgressCallbackTick = task.lastProgressCallbackTick;
            history.lastDisplaySampleTick    = nowTick;
            history.lastStateChangeTick      = task.progressStateChangeTick;
            SyncRateStreamBaselines(history, task);
            continue;
        }

        history.lastBytes             = std::min(history.lastBytes, task.completedBytes);
        history.lastVerificationBytes = std::min(history.lastVerificationBytes, task.verificationReadBytes);
        history.lastItems             = std::min(history.lastItems, task.completedItems);

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
                history.lastVerificationBytes    = task.verificationReadBytes;
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
            history.lastVerificationBytes    = task.verificationReadBytes;
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
                const RateByteDeltas deltas           = ComputeRateByteDeltas(task, history);
                const uint64_t deltaBytes             = deltas.transfer;
                const uint64_t deltaVerificationBytes = deltas.verificationRead;
                if (deltaBytes > 0 && dtSec > 0.0)
                {
                    const double instBytesPerSec = static_cast<double>(deltaBytes) / dtSec;
                    history.smoothedBytesPerSec  = SmoothRateForDisplay(history.smoothedBytesPerSec, instBytesPerSec, elapsedMs);
                    history.displayedBytesPerSec = history.smoothedBytesPerSec;
                    ++updatedTaskCount;
                }

                history.lastBytes = task.completedBytes;
                if (deltaVerificationBytes > 0 && dtSec > 0.0)
                {
                    const double instantVerificationBytesPerSec = static_cast<double>(deltaVerificationBytes) / dtSec;
                    history.smoothedVerificationBytesPerSec =
                        SmoothRateForDisplay(history.smoothedVerificationBytesPerSec, instantVerificationBytesPerSec, elapsedMs);
                    history.displayedVerificationBytesPerSec = history.smoothedVerificationBytesPerSec;
                    ++updatedTaskCount;
                }
                history.lastVerificationBytes = task.verificationReadBytes;
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
                history.displayedBytesPerSec             = DecayRateForCallbackSilence(history.smoothedBytesPerSec, silenceMs);
                history.displayedVerificationBytesPerSec = DecayRateForCallbackSilence(history.smoothedVerificationBytesPerSec, silenceMs);
            }
        }

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

        // D2-A08 provisional ETA: only while discovery is open, over the workload discovered so
        // far, from the same usable smoothed rate. It may rise or fall, and it disappears without a
        // usable rate (paused, stalled, no callbacks) or once the total closes.
        if (! itemRate && ! task.discoveryClosed && ! task.paused && task.discoveredTotalBytes > task.completedBytes &&
            IsByteRateUsableForEta(history.displayedBytesPerSec))
        {
            const double rawProvisionalSeconds =
                static_cast<double>(task.discoveredTotalBytes - task.completedBytes) / history.displayedBytesPerSec;
            history.provisionalEtaSeconds = history.hasProvisionalEta
                                                ? SmoothEtaSecondsForDisplay(history.provisionalEtaSeconds, rawProvisionalSeconds, smoothingElapsedMs)
                                                : rawProvisionalSeconds;
            history.hasProvisionalEta     = true;
        }
        else
        {
            history.hasProvisionalEta = false;
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
                    maxDisplayGapMs                        = std::max<uint64_t>(maxDisplayGapMs, displayElapsedMs);
                    const double displayVerificationSample = itemRate ? 0.0 : history.displayedVerificationBytesPerSec;
                    const double displaySample             = itemRate ? history.displayedItemsPerSec : history.displayedBytesPerSec;
                    if (displaySample > 0.0 || history.count > 0 || history.pendingHueWeightCount > 0u)
                    {
                        AppendResampledRateSamples(history, displayElapsedMs, displaySample, displayVerificationSample, hue);
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

void FileOperationsPopupInternal::FileOperationsPopupState::LayoutChrome(float width, float height, bool showPauseResumeAll, bool showCancelAll) noexcept
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

    const float actionW      = width >= DipsToPixels(560.0f, _dpi) ? DipsToPixels(112.0f, _dpi) : DipsToPixels(72.0f, _dpi);
    const float cancelW      = showCancelAll ? actionW : 0.0f;
    const float pauseResumeW = showPauseResumeAll ? actionW : 0.0f;
    const float detailsW     = footerBtnH;
    const float optionsW     = footerBtnH;
    const float minQueueW    = DipsToPixels(112.0f, _dpi);
    const float idealQueueW  = DipsToPixels(168.0f, _dpi);

    float cursor              = footerMargin;
    _footerPauseResumeAllRect = pauseResumeW > 0.0f ? D2D1::RectF(cursor, footerBtnY, std::min(rightEdge, cursor + pauseResumeW), footerBtnY + footerBtnH)
                                                    : D2D1::RectF(cursor, footerBtnY, cursor, footerBtnY + footerBtnH);
    cursor                    = pauseResumeW > 0.0f ? (_footerPauseResumeAllRect.right + footerGap) : cursor;
    _footerCancelAllRect      = cancelW > 0.0f ? D2D1::RectF(cursor, footerBtnY, std::min(rightEdge, cursor + cancelW), footerBtnY + footerBtnH)
                                               : D2D1::RectF(cursor, footerBtnY, cursor, footerBtnY + footerBtnH);
    cursor                    = cancelW > 0.0f ? (_footerCancelAllRect.right + footerGap) : cursor;

    const float detailsLeft  = std::max(footerMargin, rightEdge - detailsW);
    _footerDetailsToggleRect = D2D1::RectF(detailsLeft, footerBtnY, rightEdge, footerBtnY + footerBtnH);

    const float optionsLeft = std::max(cursor, detailsLeft - footerGap - optionsW);
    _footerOptionsRect      = D2D1::RectF(optionsLeft, footerBtnY, std::min(detailsLeft - footerGap, optionsLeft + optionsW), footerBtnY + footerBtnH);
    _footerAutoDismissRect  = {};
    _footerDensityRect      = {};

    const float rightCursor = std::max(cursor, _footerOptionsRect.left - footerGap);
    const float available   = std::max(0.0f, rightCursor - cursor);
    const float queueW      = std::min(idealQueueW, available);
    _footerQueueModeRect    = queueW >= minQueueW ? D2D1::RectF(rightCursor - queueW, footerBtnY, rightCursor, footerBtnY + footerBtnH)
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

void FileOperationsPopupInternal::FileOperationsPopupState::ArmUiMotionTimer(HWND hwnd) noexcept
{
    if (! hwnd || IsReducedMotionEnabled())
    {
        return;
    }

    _uiMotionTimerDueTick = GetTickCount64() + kFileOperationsPopupUiMotionTimerWindowMs;
    SetTimer(hwnd, kFileOperationsPopupTimerId, kFileOperationsPopupUiMotionTimerIntervalMs, nullptr);
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
                         RedSalamander::DxUi::ButtonVariant::Selector);
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
    DrawDxUiButtonChrome(button, format, text, RedSalamander::DxUi::ButtonVariant::DropDown);
}

void FileOperationsPopupInternal::FileOperationsPopupState::DrawSelectorButton(const PopupButton& button,
                                                                               IDWriteTextFormat* format,
                                                                               std::wstring_view text) noexcept
{
    DrawDxUiButtonChrome(button, format, text, RedSalamander::DxUi::ButtonVariant::Selector);
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

    const RedSalamander::DxUi::DisclosureChevronVisualState visual =
        RedSalamander::DxUi::ResolveDisclosureChevronVisualState(expandedProgress, RedSalamander::DxUi::ChevronDirection::Left);
    const wchar_t fluentGlyph = visual.direction == RedSalamander::DxUi::ChevronDirection::Left ? FluentIcons::kChevronLeft : FluentIcons::kChevronDown;
    const wchar_t fallbackGlyph =
        visual.direction == RedSalamander::DxUi::ChevronDirection::Left ? FluentIcons::kFallbackChevronLeft : FluentIcons::kFallbackChevronDown;
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

void FileOperationsPopupInternal::FileOperationsPopupState::DrawBandwidthGraph(const D2D1_RECT_F& rect,
                                                                               uint64_t taskId,
                                                                               const RateHistory& history,
                                                                               std::wstring_view currentValueTrailingText,
                                                                               std::wstring_view overlayText,
                                                                               bool rainbowMode,
                                                                               bool perStreamBands) noexcept
{
    // C4: the hosted ThroughputGraph is the only graph painter; this builds its descriptor.
    HostedGraphDescriptor hostedGraph{};
    hostedGraph.bounds                         = rect;
    hostedGraph.taskId                         = taskId;
    hostedGraph.overlayText                    = overlayText;
    hostedGraph.rainbowMode                    = rainbowMode;
    hostedGraph.perStreamBands                 = perStreamBands;
    hostedGraph.currentBandwidthBytesPerSecond = CurrentBandwidthForGraphMarker(history);
    if (hostedGraph.currentBandwidthBytesPerSecond > 0.0)
    {
        const uint64_t roundedBytesPerSecond     = SaturatingRoundNonNegativeToUint64(hostedGraph.currentBandwidthBytesPerSecond);
        hostedGraph.currentBandwidthText         = FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_BYTES, FormatBytesCompact(roundedBytesPerSecond));
        hostedGraph.currentBandwidthTrailingText = currentValueTrailingText;
    }
    hostedGraph.samples.reserve(history.count);
    hostedGraph.verificationSamples.reserve(history.count);
    for (size_t sampleOffset = 0u; sampleOffset < history.count; ++sampleOffset)
    {
        const size_t index = (history.writeIndex + RateHistory::kMaxSamples - history.count + sampleOffset) % RateHistory::kMaxSamples;
        RedSalamander::DxUi::ThroughputGraphSample sample{};
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
        case PopupHitTest::Kind::TaskDestination: return LoadStringResource(nullptr, IDS_FILEOP_BTN_DESTINATION);
        case PopupHitTest::Kind::TaskSpeedLimit:
        {
            const auto* task = fileOps ? fileOps->FindTask(hit.taskId) : nullptr;
            return task ? FormatSpeedLimitSelectorText(task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire))
                        : LoadStringResource(nullptr, IDS_FILEOP_BTN_SPEED_LIMIT);
        }
        case PopupHitTest::Kind::TaskShowLog: return LoadStringResource(nullptr, IDS_FILEOP_BTN_SHOW_LOG);
        case PopupHitTest::Kind::TaskExportIssues: return LoadStringResource(nullptr, IDS_FILEOP_BTN_EXPORT_ISSUES);
        case PopupHitTest::Kind::TaskCompletedMore:
        case PopupHitTest::Kind::TaskConflictMore: return L"\u2026";
        case PopupHitTest::Kind::TaskConflictToggleApplyToAll: return LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_APPLY_TO_ALL_SHORT);
        case PopupHitTest::Kind::TaskConflictAction:
        {
            ConflictBucket bucket = ConflictBucket::Unknown;
            FileSystemOperation operation = FILESYSTEM_COPY;
            if (const auto* task = fileOps ? fileOps->FindTask(hit.taskId) : nullptr)
            {
                std::scoped_lock lock(task->_conflictArbiter.mutex);
                bucket = task->_conflictArbiter.prompt.bucket;
                operation = task->GetOperation();
            }
            return ConflictActionText(static_cast<ConflictAction>(hit.data), bucket, operation);
        }
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

bool FileOperationsPopupInternal::FileOperationsPopupState::HandleHostedKeyDown(HWND hwnd, UINT virtualKey) noexcept
{
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

    if (virtualKey == VK_DELETE && hit.taskId != 0u)
    {
        if (auto* task = fileOps ? fileOps->FindTask(hit.taskId) : nullptr)
        {
            task->RequestCancel();
            Invalidate(hwnd);
        }
        return true;
    }

    if (virtualKey == VK_SPACE && hit.taskId != 0u)
    {
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
    {
        return false;
    }

    FolderWindow::FileOperationState::Task* const task = fileOps->FindTask(_conflictEscapeTaskId);
    if (! task)
    {
        return false;
    }

    const ConflictAction expectedAction = ConflictActionFromRaw(_conflictEscapeAction);
    bool applyToAll                     = false;
    {
        std::scoped_lock lock(task->_conflictArbiter.mutex);
        const auto& prompt = task->_conflictArbiter.prompt;
        if (! prompt.active || ! prompt.buttonsPublishable || prompt.escapeAction != expectedAction)
        {
            return false;
        }
        applyToAll = prompt.applyToAllChecked;
    }

    task->SubmitConflictDecision(expectedAction, applyToAll);
    Invalidate(hwnd);
    return true;
}

void FileOperationsPopupInternal::FileOperationsPopupState::SyncHostedControls(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept
{
    if (! _controlsRoot || _controlsHost.GetHwnd() != hwnd)
    {
        return;
    }

    const uint64_t startedUs = Debug::Perf::IsCaptureEnabled() ? PerfNowUs() : 0u;
    _controlsRoot->SetBounds(_controlsHost.GetClientBoundsDip());

    std::optional<PopupHitTest> focusedHit;
    if (auto* focused = _controlsHost.GetFocusControl())
    {
        const auto focusedIt = std::ranges::find(_hostedButtons, focused);
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
    { return left.kind == right.kind && left.taskId == right.taskId && left.data == right.data; };
    const bool rightToLeft = (GetWindowLongPtrW(hwnd, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) != 0;
    bool rebuild           = _buttons.size() != _hostedButtonHits.size() || _hostedProgressDescriptors.size() != _hostedProgressBars.size() ||
                             _hostedGraphDescriptors.size() != _hostedGraphs.size() || _hostedTooltipDescriptors.size() != _hostedTooltipRegions.size() ||
                             rightToLeft != _hostedRightToLeft;
    if (! rebuild)
    {
        for (size_t index = 0u; index < _buttons.size(); ++index)
        {
            if (! sameHit(_buttons[index].hit, _hostedButtonHits[index]))
            {
                rebuild = true;
                break;
            }
        }
    }

    if (rebuild)
    {
        // The host owns non-owning focus/hover/capture pointers into the retained tree.
        // Clear them while every old child is still alive, then restore focus by stable hit identity below.
        _controlsHost.ResetInteractionState();
        _controlsRoot->ClearChildren();
        _hostedButtons.assign(_buttons.size(), nullptr);
        _hostedButtonHits.resize(_buttons.size());
        _hostedProgressBars.assign(_hostedProgressDescriptors.size(), nullptr);
        _hostedProgressAnimations.assign(_hostedProgressDescriptors.size(), ScalarAnimation{});
        _hostedGraphs.assign(_hostedGraphDescriptors.size(), nullptr);
        _hostedTooltipRegions.assign(_hostedTooltipDescriptors.size(), nullptr);

        enum class HostedKind : uint8_t
        {
            Button,
            Progress,
            Graph,
            Tooltip,
        };
        struct HostedVisualOrder final
        {
            D2D1_RECT_F bounds{};
            HostedKind kind = HostedKind::Button;
            size_t index    = 0u;
        };
        std::vector<HostedVisualOrder> visualOrder;
        visualOrder.reserve(_buttons.size() + _hostedProgressDescriptors.size() + _hostedGraphDescriptors.size() + _hostedTooltipDescriptors.size());
        for (size_t index = 0u; index < _buttons.size(); ++index)
        {
            visualOrder.push_back({_buttons[index].bounds, HostedKind::Button, index});
        }
        for (size_t index = 0u; index < _hostedProgressDescriptors.size(); ++index)
        {
            visualOrder.push_back({_hostedProgressDescriptors[index].bounds, HostedKind::Progress, index});
        }
        for (size_t index = 0u; index < _hostedGraphDescriptors.size(); ++index)
        {
            visualOrder.push_back({_hostedGraphDescriptors[index].bounds, HostedKind::Graph, index});
        }
        for (size_t index = 0u; index < _hostedTooltipDescriptors.size(); ++index)
        {
            visualOrder.push_back({_hostedTooltipDescriptors[index].bounds, HostedKind::Tooltip, index});
        }
        std::stable_sort(visualOrder.begin(),
                         visualOrder.end(),
                         [rightToLeft](const HostedVisualOrder& left, const HostedVisualOrder& right) noexcept
        {
            if (left.bounds.top != right.bounds.top)
            {
                return left.bounds.top < right.bounds.top;
            }
            return rightToLeft ? left.bounds.right > right.bounds.right : left.bounds.left < right.bounds.left;
        });

        for (const HostedVisualOrder& visual : visualOrder)
        {
            if (visual.kind == HostedKind::Button)
            {
                const PopupHitTest hit            = _buttons[visual.index].hit;
                const std::wstring controlText    = HostedControlText(hit);
                const std::wstring accessibleName = HostedControlAccessibleName(hit);
                auto* button                      = _controlsRoot->AddChild<RedSalamander::DxUi::Button>(controlText);
                button->SetAccessibleName(accessibleName);
                button->SetAccessibleAutomationId(std::format(L"FileOperations.Action.{}.{}.{}", static_cast<uint32_t>(hit.kind), hit.taskId, hit.data));
                button->SetTooltipText(accessibleName != controlText ? accessibleName : std::wstring{});
                button->SetDensity(RedSalamander::DxUi::Density::Compact);
                const auto activate = [this, hwnd, hit]
                {
                    if (IsWindow(hwnd) != FALSE && hostLifetime.lock())
                    {
                        static_cast<void>(OnActivatedHit(hwnd, hit));
                    }
                };
                if (OpensButtonFlyout(hit.kind))
                {
                    button->SetOnDropDownClick(activate);
                }
                else
                {
                    button->SetOnClick(activate);
                }
                _hostedButtons[visual.index]    = button;
                _hostedButtonHits[visual.index] = hit;
            }
            else if (visual.kind == HostedKind::Progress)
            {
                const HostedProgressDescriptor& descriptor = _hostedProgressDescriptors[visual.index];
                auto* progress                             = _controlsRoot->AddChild<RedSalamander::DxUi::ProgressBar>();
                progress->SetAccessibleAutomationId(descriptor.taskId == 0u ? L"FileOperations.Aggregate.Progress"
                                                                            : std::format(L"FileOperations.Task.{}.Progress", descriptor.taskId));
                _hostedProgressBars[visual.index] = progress;
            }
            else if (visual.kind == HostedKind::Graph)
            {
                const HostedGraphDescriptor& descriptor = _hostedGraphDescriptors[visual.index];
                auto* graph                             = _controlsRoot->AddChild<RedSalamander::DxUi::ThroughputGraph>();
                graph->SetAccessibleAutomationId(std::format(L"FileOperations.Task.{}.ThroughputGraph", descriptor.taskId));
                _hostedGraphs[visual.index] = graph;
            }
            else
            {
                const HostedTooltipDescriptor& descriptor = _hostedTooltipDescriptors[visual.index];
                auto* region                              = _controlsRoot->AddChild<RedSalamander::DxUi::Label>();
                region->SetAccessibilityRole(RedSalamander::DxUi::AccessibilityRole::Status);
                region->SetAccessibleAutomationId(descriptor.automationId);
                region->SetAccessibleName(descriptor.text);
                region->SetAccessibleHelpText(descriptor.text);
                region->SetTooltipText(descriptor.text);
                _hostedTooltipRegions[visual.index] = region;
            }
        }
        _hostedRightToLeft = rightToLeft;
    }

    for (size_t index = 0u; index < _buttons.size() && index < _hostedButtons.size(); ++index)
    {
        const D2D1_RECT_F bounds                  = _buttons[index].bounds;
        RedSalamander::DxUi::Button* const button = _hostedButtons[index];
        const PopupHitTest& hit                   = _buttons[index].hit;
        const std::wstring controlText            = HostedControlText(hit);
        const std::wstring accessibleName         = HostedControlAccessibleName(hit);
        button->SetBounds(D2D1::RectF(_controlsHost.PixelsToDip(bounds.left),
                                      _controlsHost.PixelsToDip(bounds.top),
                                      _controlsHost.PixelsToDip(bounds.right),
                                      _controlsHost.PixelsToDip(bounds.bottom)));
        button->SetText(controlText);
        button->SetAccessibleName(accessibleName);
        button->SetTooltipText(accessibleName != controlText ? accessibleName : std::wstring{});
        button->SetVariant(HostedButtonVariant(hit.kind));

        if (hit.kind == PopupHitTest::Kind::FooterToggleDetails)
        {
            button->SetDisclosureCollapsedDirection(RedSalamander::DxUi::ChevronDirection::Left);
            button->SetDisclosureExpanded(fileOps ? ! fileOps->GetPopupFooterOnly() : true);
        }
        else if (hit.kind == PopupHitTest::Kind::CompletedGroupToggle)
        {
            button->SetDisclosureCollapsedDirection(RedSalamander::DxUi::ChevronDirection::Left);
            button->SetDisclosureExpanded(_completedGroupExpanded);
        }
        else if (hit.kind == PopupHitTest::Kind::TaskToggleCollapse)
        {
            button->SetDisclosureCollapsedDirection(RedSalamander::DxUi::ChevronDirection::Left);
            button->SetDisclosureExpanded(! IsTaskCollapsedForDisplay(hit.taskId, fileOps ? fileOps->GetPopupCompactDensity() : false));
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
    for (size_t index = 0u; index < _hostedProgressDescriptors.size() && index < _hostedProgressBars.size(); ++index)
    {
        const HostedProgressDescriptor& descriptor = _hostedProgressDescriptors[index];
        auto* progress                             = _hostedProgressBars[index];
        progress->SetBounds(boundsToDip(descriptor.bounds));
        progress->SetTrackHeightDip(_controlsHost.PixelsToDip(std::max(0.0f, descriptor.bounds.bottom - descriptor.bounds.top)));
        progress->SetMinimum(0.0);
        progress->SetMaximum(100.0);
        progress->SetValue(ResolveHostedProgressValue(index, descriptor.value, animationTick, reducedMotion));
        progress->SetIndeterminate(descriptor.indeterminate);
        if (descriptor.segmented && ! descriptor.indeterminate && folderWindow)
        {
            progress->SetSegmentedValues(
                descriptor.primarySegmentValue, descriptor.secondarySegmentValue, folderWindow->GetTheme().fileOperations.progressVerify);
            progress->SetAccessibleHelpText(FormatStringResource(nullptr,
                                                                 IDS_FMT_FILEOPS_PROGRESS_SEGMENTS_ACCESSIBLE,
                                                                 static_cast<unsigned int>(std::lround(descriptor.primarySegmentValue)),
                                                                 static_cast<unsigned int>(std::lround(descriptor.secondarySegmentValue))));
        }
        else
        {
            progress->ClearSegmentedValues();
            progress->SetAccessibleHelpText({});
        }
        progress->SetAccessibleName(taskStatusName(descriptor.taskId));
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
        graph->SetCurrentValueMarker(descriptor.currentBandwidthBytesPerSecond, descriptor.currentBandwidthText, descriptor.currentBandwidthTrailingText);
        graph->SetOverlayText(descriptor.overlayText);
        graph->SetRainbowMode(descriptor.rainbowMode);
        graph->SetPerStreamBands(descriptor.perStreamBands);
        graph->SetAccessibleName(LoadStringResource(nullptr, IDS_FILEOPS_THROUGHPUT_GRAPH));
        graph->SetAccessibleHelpText(descriptor.currentBandwidthTrailingText.empty()
                                         ? descriptor.currentBandwidthText
                                         : std::format(L"{0} {1}", descriptor.currentBandwidthText, descriptor.currentBandwidthTrailingText));
    }
    for (size_t index = 0u; index < _hostedTooltipDescriptors.size() && index < _hostedTooltipRegions.size(); ++index)
    {
        const HostedTooltipDescriptor& descriptor = _hostedTooltipDescriptors[index];
        auto* region                              = _hostedTooltipRegions[index];
        region->SetBounds(boundsToDip(descriptor.bounds));
        region->SetAccessibleAutomationId(descriptor.automationId);
        region->SetAccessibleName(descriptor.text);
        region->SetAccessibleHelpText(descriptor.text);
        region->SetTooltipText(descriptor.text);
    }

    if (rebuild && focusedHit.has_value())
    {
        for (size_t index = 0u; index < _hostedButtonHits.size(); ++index)
        {
            if (sameHit(_hostedButtonHits[index], focusedHit.value()))
            {
                _controlsHost.SetFocusControl(_hostedButtons[index]);
                break;
            }
        }
    }

    const TaskSnapshot* decisionTask = nullptr;
    if (focusedHit.has_value() && focusedHit->taskId != 0u)
    {
        const auto focusedTask = std::ranges::find(snapshot, focusedHit->taskId, &TaskSnapshot::taskId);
        if (focusedTask != snapshot.end() && focusedTask->conflict.active && focusedTask->conflict.buttonsPublishable)
        {
            decisionTask = &*focusedTask;
        }
    }
    if (! decisionTask)
    {
        // Enter/Escape bind to an unfocused decision only when exactly one is actionable, so a
        // key press never resolves a prompt the user is not looking at.
        size_t actionableDecisions = 0u;
        for (const TaskSnapshot& task : snapshot)
        {
            if (task.conflict.active && task.conflict.buttonsPublishable)
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

    const auto findPublishedActionButton = [&](uint64_t taskId, uint8_t rawAction) noexcept -> RedSalamander::DxUi::Button*
    {
        for (size_t index = 0u; index < _hostedButtonHits.size() && index < _hostedButtons.size(); ++index)
        {
            const PopupHitTest& hit = _hostedButtonHits[index];
            if (hit.kind == PopupHitTest::Kind::TaskConflictAction && hit.taskId == taskId && hit.data == static_cast<uint32_t>(rawAction))
            {
                return _hostedButtons[index];
            }
        }
        return nullptr;
    };

    RedSalamander::DxUi::Button* defaultButton = nullptr;
    RedSalamander::DxUi::Button* escapeButton  = nullptr;
    _conflictEscapeTaskId                      = 0u;
    _conflictEscapeAction                      = 0u;
    if (decisionTask)
    {
        defaultButton         = findPublishedActionButton(decisionTask->taskId, decisionTask->conflict.defaultAction);
        escapeButton          = findPublishedActionButton(decisionTask->taskId, decisionTask->conflict.escapeAction);
        _conflictEscapeTaskId = decisionTask->taskId;
        _conflictEscapeAction = decisionTask->conflict.escapeAction;
    }

    _controlsHost.SetDefaultButton(defaultButton);
    _controlsHost.SetCancelButton(escapeButton);

    const bool recycleEscalation = decisionTask && static_cast<ConflictBucket>(decisionTask->conflict.bucket) == ConflictBucket::RecycleFailed &&
                                   decisionTask->conflict.defaultAction == RawConflictAction(ConflictAction::Cancel);
    if (recycleEscalation && defaultButton != nullptr)
    {
        if (_recycleEscalationFocusedTaskId != decisionTask->taskId || (rebuild && ! focusedHit.has_value()))
        {
            _controlsHost.SetFocusControl(defaultButton);
        }
        _recycleEscalationFocusedTaskId = decisionTask->taskId;
    }
    else
    {
        _recycleEscalationFocusedTaskId = 0u;
    }

    for (const TaskSnapshot& task : snapshot)
    {
        const auto previous                     = _announcedTaskStatuses.find(task.taskId);
        const auto previousConflictLoading      = _announcedConflictMetadataLoading.find(task.taskId);
        const bool isAnnounceableStatus         = task.statusKind == TaskSnapshot::StatusKind::Done || task.statusKind == TaskSnapshot::StatusKind::Partial ||
                                                  task.statusKind == TaskSnapshot::StatusKind::Failed || task.statusKind == TaskSnapshot::StatusKind::Conflict;
        const bool statusChanged                = previous != _announcedTaskStatuses.end() && previous->second != task.statusKind;
        const bool initialAttention             = previous == _announcedTaskStatuses.end() && task.statusKind == TaskSnapshot::StatusKind::Conflict;
        const bool conflictDecisionStateChanged = task.statusKind == TaskSnapshot::StatusKind::Conflict &&
                                                  previousConflictLoading != _announcedConflictMetadataLoading.end() &&
                                                  previousConflictLoading->second != task.conflict.metadataLoading;
        if (isAnnounceableStatus && (statusChanged || initialAttention || conflictDecisionStateChanged) &&
            (_hostedAccessibilityInitialized || initialAttention))
        {
            std::wstring announcement = StatusTextForTask(task, task.statusKind, GetTickCount64());
            if (task.statusKind == TaskSnapshot::StatusKind::Conflict && ! task.conflict.metadataLoading)
            {
                const std::wstring question = BuildConflictQuestion(task);
                if (! question.empty())
                {
                    announcement = std::format(L"{0}. {1}", announcement, question);
                }
            }
            if (RedSalamander::DxUi::RaiseWindowHostAccessibilityNotification(hwnd, announcement, std::format(L"FileOperations.Task.{}", task.taskId)))
            {
                ++_hostedAccessibilityNotificationCount;
            }
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
    _controlsHost.RefreshAccessibilitySnapshot();

    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(
            L"FileOps.Popup.HostedControls.SyncUs", rebuild ? L"rebuild" : L"update", PerfElapsedUs(startedUs), _buttons.size(), snapshot.size(), S_OK);
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
    if (_failureSurface)
    {
        // C4: the native children paint themselves; the popup only clears its ground.
        FillRect(hdc.get(), &ps.rcPaint, GetSysColorBrush(COLOR_WINDOW));
        return;
    }

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
        // C4: no Direct2D means no hand-painted popup; the native failure surface takes over.
        if (EnterFailureSurface(hwnd))
        {
            FillRect(hdc.get(), &ps.rcPaint, GetSysColorBrush(COLOR_WINDOW));
        }
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
    const bool showPauseResumeAll                           = HasFooterPauseResumeAllControl(globalSummary);
    const bool showCancelAll                                = fileOps && fileOps->HasActiveOperations();
    const CompletedGroupResultSummary completedGroupSummary = BuildCompletedGroupResultSummary(snapshot);
    const uint32_t completedGroupCount                      = completedGroupSummary.total;
    const bool showCompletedGroup                           = ShouldShowCompletedGroup(completedGroupCount);
    constexpr ULONGLONG kCompletedInFlightGraceMs           = 300ull;
    const ULONGLONG renderTick                              = GetTickCount64();
    const float completedGroupVisibility                    = ResolveCompletedGroupVisibility(renderTick, reducedMotion);

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

    float width  = 0.0f;
    float height = 0.0f;

    const float padding = DipsToPixels(10.0f, _dpi);
    const float cardGap = DipsToPixels(10.0f, _dpi);

    const float expandedCardH         = DipsToPixels(kFileOperationsExpandedCardHeightDip, _dpi);
    const float expandedTransferCardH = DipsToPixels(kFileOperationsExpandedTransferCardHeightDip, _dpi);
    const float collapsedCardH        = DipsToPixels(44.0f, _dpi);
    const float baseLineH             = DipsToPixels(18.0f, _dpi);
    const float fromToGapY            = DipsToPixels(4.0f, _dpi);
    const float promptRowGap          = DipsToPixels(4.0f, _dpi);
    const float promptMeasureWidth =
        std::max(DipsToPixels(320.0f, _dpi), std::max(static_cast<float>(_clientSize.cx), DipsToPixels(480.0f, _dpi)) - padding * 4.0f);

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

                h = ResolveAnimatedTaskHeight(task.taskId, h, renderTick, reducedMotion);
                if (showCompletedGroup && IsCompletedGroupTask(task))
                {
                    h *= completedGroupVisibility;
                }
                cardHeights.push_back(h);
                continue;
            }

            const bool displayCollapsed = IsTaskCollapsedForDisplay(task.taskId, compactDensity);
            float h                     = displayCollapsed ? collapsedCardH : expandedCardH;
            if (! displayCollapsed && task.finished)
            {
                h = DipsToPixels(TaskHasTerminalAxisBadges(task) ? 202.0f : 178.0f, _dpi);
            }
            else if (! displayCollapsed && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE))
            {
                h = expandedTransferCardH;
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
                const ConflictPromptContent promptContent = BuildConflictPromptContent(task);
                const float promptContentHeight           = MeasureConflictPromptContentHeight(
                    _dwriteFactory.get(), _bodyFormat.get(), _smallFormat.get(), promptContent, promptMeasureWidth, baseLineH, promptRowGap);
                h += promptContentHeight + DipsToPixels(12.0f, _dpi);
            }
            h = ResolveAnimatedTaskHeight(task.taskId, h, renderTick, reducedMotion);
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

        LayoutChrome(width, height, showPauseResumeAll, showCancelAll);

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
    _hostedProgressDescriptors.clear();
    _hostedGraphDescriptors.clear();
    _hostedTooltipDescriptors.clear();
#ifdef ENABLE_TESTS
    _graphRowColorMatchCountsForSelfTest.clear();
    _graphRowColorMismatchCountsForSelfTest.clear();
#endif

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
                    const D2D1_RECT_F chevronRc = D2D1::RectF(contentRight - chevronSize, chevronTop, contentRight, chevronTop + chevronSize);

                    const float clearW = std::min(DipsToPixels(128.0f, _dpi), std::max(0.0f, (contentRight - textX) * 0.34f));
                    PopupButton clearBtn{};
                    clearBtn.bounds   = D2D1::RectF(textX, cardRect.top + DipsToPixels(7.0f, _dpi), textX + clearW, cardRect.bottom - DipsToPixels(7.0f, _dpi));
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
                        DrawButton(clearBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_BTN_CLEAR_COMPLETED));
                    }
                    const float middleLeft  = clearBtn.bounds.right + DipsToPixels(8.0f, _dpi);
                    const float middleRight = std::max(middleLeft, chevronRc.left - DipsToPixels(8.0f, _dpi));
                    if (_completedGroupExpanded)
                    {
                        const std::wstring groupText =
                            FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPLETED_GROUP, static_cast<unsigned long>(completedGroupCount));
                        const D2D1_RECT_F textRc =
                            D2D1::RectF(middleLeft, cardRect.top + DipsToPixels(10.0f, _dpi), middleRight, cardRect.bottom - DipsToPixels(8.0f, _dpi));
                        IDWriteTextFormat* groupFormat = _headerFormat ? _headerFormat.get() : (_bodyFormat ? _bodyFormat.get() : nullptr);
                        if (groupFormat && _textBrush)
                        {
                            _target->DrawTextW(
                                groupText.data(), static_cast<UINT32>(groupText.size()), groupFormat, textRc, _textBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        }
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
                                        FluentIcons::kWarning,
                                        FluentIcons::kFallbackWarning,
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
                        const float badgeGap = DipsToPixels(6.0f, _dpi);
                        const float badgeH   = DipsToPixels(22.0f, _dpi);
                        const float badgeW   = DipsToPixels(48.0f, _dpi);
                        float badgeLeft      = middleLeft;
                        for (const ResultBadge& badge : badges)
                        {
                            if (badge.count == 0u || badgeLeft + badgeW > middleRight)
                            {
                                continue;
                            }

                            const float badgeTop      = cardRect.top + (taskCardH - badgeH) * 0.5f;
                            const D2D1_RECT_F badgeRc = D2D1::RectF(badgeLeft, badgeTop, badgeLeft + badgeW, badgeTop + badgeH);
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
                            if (iconFormat && glyph != 0 && _graphDynamicBrush)
                            {
                                _graphDynamicBrush->SetColor(color);
                                const wchar_t iconText[2]{glyph, 0};
                                _target->DrawTextW(iconText, 1u, iconFormat, iconRc, _graphDynamicBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
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
                            badgeLeft = badgeRc.right + badgeGap;
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

                const float headerRight = std::max(textX, headerControlsLeft - collapseBtnGap);
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
                    const float percentW         = showPercent ? DipsToPixels(42.0f, _dpi) : 0.0f;
                    const float preferredMeterW  = DipsToPixels(118.0f, _dpi);
                    const float compactMeterW    = std::min(preferredMeterW, availableHeaderW * 0.40f);
                    const float percentGap       = showPercent ? DipsToPixels(6.0f, _dpi) : 0.0f;
                    const float barW             = compactMeterW - percentW - percentGap;
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
                        progressDescriptor.primarySegmentValue   = VerificationTransferProgressPercent(task);
                        progressDescriptor.secondarySegmentValue = VerificationProofProgressPercent(task);
                        _hostedProgressDescriptors.push_back(progressDescriptor);

                        if (showPercent)
                        {
                            const std::wstring percentText = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_PERCENT, percent);
                            const D2D1_RECT_F percentRc    = D2D1::RectF(barRc.right + percentGap, headerTop, meterRight, headerBottom);
                            _target->DrawTextW(percentText.data(),
                                               static_cast<UINT32>(percentText.size()),
                                               _smallFormat.get(),
                                               percentRc,
                                               _subTextBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        }
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

                if (task.finished && TaskHasTerminalAxisBadges(task))
                {
                    std::array<TerminalAxisBadge, 3u> badges{};
                    const size_t badgeCount = BuildTerminalAxisBadges(task, badges);
                    const float badgeTop    = textY + DipsToPixels(3.0f, _dpi);
                    const float badgeH      = DipsToPixels(18.0f, _dpi);
                    const float badgeGap    = DipsToPixels(5.0f, _dpi);
                    const float badgeHPad   = DipsToPixels(7.0f, _dpi);
                    float badgeLeft         = textX;
                    for (size_t badgeIndex = 0u; badgeIndex < badgeCount && badgeLeft < contentRight; ++badgeIndex)
                    {
                        const size_t remainingBadgeCount = badgeCount - badgeIndex;
                        const float remainingGap         = badgeGap * static_cast<float>(remainingBadgeCount - 1u);
                        const float availableW           = std::max(0.0f, contentRight - badgeLeft - remainingGap);
                        const float badgeMaxW            = availableW / static_cast<float>(remainingBadgeCount);
                        const float measuredTextW =
                            _smallFormat ? MeasureTextWidth(_dwriteFactory.get(), _smallFormat.get(), badges[badgeIndex].text, badgeMaxW, badgeH) : 0.0f;
                        const float badgeW        = std::clamp(measuredTextW + badgeHPad * 2.0f, std::min(DipsToPixels(48.0f, _dpi), badgeMaxW), badgeMaxW);
                        const D2D1_RECT_F badgeRc = D2D1::RectF(badgeLeft, badgeTop, badgeLeft + badgeW, badgeTop + badgeH);
                        const float badgeRadius   = ClampCornerRadius(badgeRc, DipsToPixels(5.0f, _dpi));
                        if (_graphDynamicBrush)
                        {
                            D2D1::ColorF badgeColor = StatusVisualColorForTone(appTheme, badges[badgeIndex].tone);
                            badgeColor.a            = appTheme.highContrast ? 0.28f : 0.14f;
                            _graphDynamicBrush->SetColor(badgeColor);
                            _target->FillRoundedRectangle(D2D1::RoundedRect(badgeRc, badgeRadius, badgeRadius), _graphDynamicBrush.get());
                            badgeColor.a = appTheme.highContrast ? 1.0f : 0.55f;
                            _graphDynamicBrush->SetColor(badgeColor);
                            _target->DrawRoundedRectangle(D2D1::RoundedRect(badgeRc, badgeRadius, badgeRadius), _graphDynamicBrush.get(), 1.0f);
                        }
                        if (_smallFormat && _textBrush)
                        {
                            const D2D1_RECT_F badgeTextRc = D2D1::RectF(badgeRc.left + badgeHPad, badgeRc.top, badgeRc.right - badgeHPad, badgeRc.bottom);
                            _target->DrawTextW(badges[badgeIndex].text.data(),
                                               static_cast<UINT32>(badges[badgeIndex].text.size()),
                                               _smallFormat.get(),
                                               badgeTextRc,
                                               _textBrush.get(),
                                               D2D1_DRAW_TEXT_OPTIONS_CLIP);
                        }
                        _hostedTooltipDescriptors.push_back(HostedTooltipDescriptor{
                            .bounds       = badgeRc,
                            .text         = badges[badgeIndex].text,
                            .automationId = std::format(L"FileOperations.Task.{}.Result.{}", task.taskId, badges[badgeIndex].automationSuffix),
                        });
                        badgeLeft = badgeRc.right + badgeGap;
                    }
                    textY += DipsToPixels(24.0f, _dpi);
                }

                if (task.finished)
                {
                    const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    const bool showHrLine   = FAILED(task.resultHr) && task.resultHr != partialHr;

                    if (! task.interruptedMoveNotice)
                    {
                        const std::wstring diagCounts = FormatStringResource(nullptr, IDS_FMT_FILEOPS_WARNINGS_ERRORS, task.warningCount, task.errorCount);
                        const D2D1_RECT_F countsRc    = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(diagCounts.data(), static_cast<UINT32>(diagCounts.size()), _bodyFormat.get(), countsRc, _subTextBrush.get());
                        textY += lineH;

                        if (showHrLine)
                        {
                            const std::wstring hrText =
                                FormatStringResource(nullptr, IDS_FMT_FILEOPS_RESULT_HRESULT, static_cast<unsigned long>(task.resultHr));
                            const D2D1_RECT_F hrRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(hrText.data(), static_cast<UINT32>(hrText.size()), _bodyFormat.get(), hrRc, _subTextBrush.get());
                            textY += lineH;
                        }
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

                    HostedProgressDescriptor progressDescriptor{
                        progressRc, task.taskId, static_cast<double>(Clamp01(ComputeFileOperationsTaskCompleteFractionForDisplay(task))) * 100.0, false};
                    progressDescriptor.segmented             = task.verificationRequested;
                    progressDescriptor.primarySegmentValue   = VerificationTransferProgressPercent(task);
                    progressDescriptor.secondarySegmentValue = VerificationProofProgressPercent(task);
                    _hostedProgressDescriptors.push_back(progressDescriptor);

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
                const bool graphBandsActive = RateHistoryGraphBandsAreActive(history, theme);

                const bool discoveryInProgress = compactProgressPresentation.showDiscoveryActivity;
                if (discoveryInProgress)
                {
                    const std::wstring discoveryText =
                        LoadStringResource(nullptr, task.discoverySkipped ? IDS_FILEOPS_DISCOVERING_AS_NEEDED : IDS_FILEOPS_DISCOVERING);
                    const float indicatorW        = DipsToPixels(24.0f, _dpi);
                    const float indicatorGap      = DipsToPixels(6.0f, _dpi);
                    const D2D1_RECT_F discoveryRc = D2D1::RectF(textX, textY, std::max(textX, contentRight - indicatorW - indicatorGap), textY + lineH);
                    _target->DrawTextW(discoveryText.data(),
                                       static_cast<UINT32>(discoveryText.size()),
                                       _bodyFormat.get(),
                                       discoveryRc,
                                       _subTextBrush.get(),
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    DrawDiscoveryActivityIndicator(D2D1::RectF(contentRight - indicatorW, textY, contentRight, textY + lineH), nowTick, reducedMotion);
                    textY += lineH;

                    // D2-A08: exact discovered/processed counters instead of a provisional fraction.
                    const uint64_t discoveredBytes = std::max(task.discoveredTotalBytes, task.completedBytes);
                    const uint64_t discoveredItems = std::max<uint64_t>(
                        static_cast<uint64_t>(task.discoveredFileCount) + static_cast<uint64_t>(task.discoveredDirectoryCount), task.completedItems);
                    const std::wstring transferText =
                        FormatEmbeddedStringResource(nullptr,
                                                     IDS_FMT_FILEOPS_DISCOVERY_PROGRESS,
                                                     FormatBytesCompact(task.completedBytes),
                                                     discoveredBytes > 0 ? FormatBytesCompact(discoveredBytes) : std::wstring(L"?"),
                                                     static_cast<unsigned long>(task.completedItems),
                                                     static_cast<unsigned long>((std::min<uint64_t>)(discoveredItems, ULONG_MAX)));
                    const D2D1_RECT_F transferRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                    _target->DrawTextW(transferText.data(),
                                       static_cast<UINT32>(transferText.size()),
                                       _bodyFormat.get(),
                                       transferRc,
                                       _subTextBrush.get(),
                                       D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    textY += lineH;
                }

                if (! task.conflict.active && task.operation == FILESYSTEM_DELETE)
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
                        textY += lineH;

                        const bool showSizeProgress = task.discoveryClosed && task.discoveredTotalBytes > 0 && task.completedBytes > 0;
                        if (showSizeProgress)
                        {
                            const std::wstring sizeProgressText = FormatEmbeddedStringResource(
                                nullptr, IDS_FMT_FILEOPS_SIZE_PROGRESS, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.discoveredTotalBytes));
                            const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                            _target->DrawTextW(
                                sizeProgressText.data(), static_cast<UINT32>(sizeProgressText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
                            textY += lineH;
                        }
                        else if (task.discoveryClosed && task.totalItems > 0)
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
                else if (! task.conflict.active)
                {
                    // Show size progress (transferred / total) if we have data
                    if (task.discoveryClosed && task.totalBytes > 0)
                    {
                        const std::wstring sizeProgressText = FormatEmbeddedStringResource(
                            nullptr, IDS_FMT_FILEOPS_SIZE_PROGRESS, FormatBytesCompact(task.completedBytes), FormatBytesCompact(task.totalBytes));
                        const D2D1_RECT_F sizeRc = D2D1::RectF(textX, textY, textX + textMaxW, textY + lineH);
                        _target->DrawTextW(
                            sizeProgressText.data(), static_cast<UINT32>(sizeProgressText.size()), _bodyFormat.get(), sizeRc, _subTextBrush.get());
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
                    const std::wstring fromLabel = task.verificationActive ? LoadStringResource(nullptr, IDS_FILEOPS_STATUS_VERIFYING)
                                                                           : LoadStringResource(nullptr, IDS_FILEOPS_LABEL_FROM);
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

                        if (miniBarW > 0.0f && _progressBgBrush && _progressGlobalBrush && _progressItemBrush && _progressVerifyBrush)
                        {
                            const float barTop          = textY + (lineH - miniBarH) * 0.5f;
                            const D2D1_RECT_F miniBarRc = D2D1::RectF(barLeft, barTop, barRight, barTop + miniBarH);

                            const float radiusTrack = ClampCornerRadius(miniBarRc, DipsToPixels(2.0f, _dpi));
                            _target->FillRoundedRectangle(D2D1::RoundedRect(miniBarRc, radiusTrack, radiusTrack), _progressBgBrush.get());

                            const bool hasTotal = fileTotalBytes > 0;
                            const float frac = hasTotal && fileCompletedBytes <= fileTotalBytes
                                                   ? Clamp01(static_cast<float>(static_cast<double>(fileCompletedBytes) / static_cast<double>(fileTotalBytes)))
                                                   : 0.0f;

                            const bool verificationMiniBar = task.verificationActive && ! useInFlightEntry;
                            if (! verificationMiniBar && graphBandsActive && inFlightEntry)
                            {
                                const D2D1_COLOR_F rowColor = RainbowProgressColor(theme, sourcePathText, history, inFlightEntry);
                                _progressItemBrush->SetColor(rowColor);
#ifdef ENABLE_TESTS
                                const float assignedHue      = AssignedRateStreamHue(history, inFlightEntry);
                                const D2D1_COLOR_F bandColor = RedSalamander::DxUi::ThroughputGraphColorFromHue(assignedHue, theme.dark);
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
                                hasTotal
                                    ? D2D1::RectF(miniBarRc.left, miniBarRc.top, miniBarRc.left + (miniBarRc.right - miniBarRc.left) * frac, miniBarRc.bottom)
                                    : ComputeIndeterminateBarFill(miniBarRc, nowTick, reducedMotion);
                            const float radiusFill                = ClampCornerRadius(fill, DipsToPixels(2.0f, _dpi));
                            ID2D1SolidColorBrush* const fillBrush = verificationMiniBar ? _progressVerifyBrush.get() : _progressItemBrush.get();
                            _target->FillRoundedRectangle(D2D1::RoundedRect(fill, radiusFill, radiusFill), fillBrush);
                            if (verificationMiniBar && fill.right > fill.left)
                            {
                                _target->PushAxisAlignedClip(fill, D2D1_ANTIALIAS_MODE_ALIASED);
                                const float hatchStep = std::max(3.0f, DipsToPixels(5.0f, _dpi));
                                const float hatchRise = fill.bottom - fill.top;
                                for (float x = fill.left - hatchRise; x < fill.right; x += hatchStep)
                                {
                                    _target->DrawLine(D2D1::Point2F(x, fill.bottom), D2D1::Point2F(x + hatchRise, fill.top), _progressGlobalBrush.get(), 1.0f);
                                }
                                _target->PopAxisAlignedClip();
                            }
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

                const bool hasConflictPrompt                      = task.conflict.active;
                const ConflictPromptContent conflictPromptContent = hasConflictPrompt ? BuildConflictPromptContent(task) : ConflictPromptContent{};

                const float barsHeight    = barHItem;
                const float bottomPadding = DipsToPixels(10.0f, _dpi);
                const float buttonGapY    = DipsToPixels(8.0f, _dpi);
                const float buttonH       = DipsToPixels(24.0f, _dpi);

                const float conflictRowGapY = DipsToPixels(6.0f, _dpi);
                const int conflictRows      = hasConflictPrompt && task.conflict.buttonsPublishable ? 1 : 0;
                const float conflictButtonsHeight =
                    buttonH * static_cast<float>(conflictRows) + conflictRowGapY * static_cast<float>(std::max(0, conflictRows - 1));
                const float conflictApplyLineHeight =
                    hasConflictPrompt && task.conflict.buttonsPublishable && task.conflict.applyToAllEligible ? (lineH + conflictRowGapY) : 0.0f;
                const float buttonsHeight = conflictButtonsHeight + conflictApplyLineHeight;

                const float buttonRowBottom = cardRect.bottom - bottomPadding;
                const float buttonRowTop    = buttonRowBottom - buttonsHeight;

                const float barsBottom = buttonRowTop - buttonGapY;
                const float barsTop    = barsBottom - barsHeight;

                const auto drawConflictPromptInfo = [&](const D2D1_RECT_F& rc) noexcept
                {
                    if (! _bodyFormat || ! _smallFormat || ! _dwriteFactory || ! _target)
                    {
                        return;
                    }

                    float yPrompt        = rc.top;
                    const float maxW     = std::max(0.0f, rc.right - rc.left);
                    const float rowGap   = DipsToPixels(4.0f, _dpi);
                    bool hasPriorContent = false;

                    const auto addSemanticRegion = [&](const D2D1_RECT_F& bounds, std::wstring_view text, std::wstring_view suffix)
                    {
                        if (text.empty())
                        {
                            return;
                        }
                        _hostedTooltipDescriptors.push_back(HostedTooltipDescriptor{
                            .bounds       = bounds,
                            .text         = std::wstring(text),
                            .automationId = std::format(L"FileOperations.Task.{0}.Conflict.{1}", task.taskId, suffix),
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
                        const float textHeight   = MeasureWrappedTextHeight(_dwriteFactory.get(), format, text, maxW, lineH);
                        const D2D1_RECT_F textRc = D2D1::RectF(rc.left, yPrompt, rc.right, std::min(rc.bottom, yPrompt + textHeight));
                        static_cast<void>(DrawWrappedText(_dwriteFactory.get(), _target.get(), format, brush, text, textRc, lineH));
                        addSemanticRegion(textRc, text, suffix);
                        yPrompt += textHeight;
                        hasPriorContent = true;
                    };

                    const auto drawContext = [&](std::wstring_view label, std::wstring_view path, std::wstring_view metadata, std::wstring_view suffix) noexcept
                    {
                        if (label.empty() || path.empty())
                        {
                            return;
                        }
                        if (hasPriorContent)
                        {
                            yPrompt += rowGap;
                        }
                        const float contextTop   = yPrompt;
                        const float metadataMaxW = metadata.empty() ? 0.0f : std::max(0.0f, maxW * 0.48f);
                        const std::wstring metadataDisplay =
                            metadata.empty()
                                ? std::wstring{}
                                : TruncateTextMiddleToWidth(_dwriteFactory.get(), _smallFormat.get(), metadata, metadataMaxW, lineH, kEllipsisText, 0u, 6u);
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

                        const float pathHeight   = MeasureWrappedTextHeight(_dwriteFactory.get(), _smallFormat.get(), path, maxW, lineH);
                        const D2D1_RECT_F pathRc = D2D1::RectF(rc.left, yPrompt, rc.right, std::min(rc.bottom, yPrompt + pathHeight));
                        static_cast<void>(DrawWrappedText(_dwriteFactory.get(), _target.get(), _smallFormat.get(), _textBrush.get(), path, pathRc, lineH));
                        yPrompt += pathHeight;

                        std::wstring accessibleText = std::format(L"{0} {1}", label, path);
                        if (! metadata.empty())
                        {
                            accessibleText.append(L" ");
                            accessibleText.append(metadata);
                        }
                        addSemanticRegion(D2D1::RectF(rc.left, contextTop, rc.right, std::min(rc.bottom, yPrompt)), accessibleText, suffix);
                        hasPriorContent = true;
                    };

                    drawTextRegion(
                        conflictPromptContent.state, _bodyFormat.get(), _statusWarningBrush ? _statusWarningBrush.get() : _textBrush.get(), L"State");
                    drawTextRegion(conflictPromptContent.question, _bodyFormat.get(), _textBrush.get(), L"Question");
                    drawTextRegion(conflictPromptContent.facts, _smallFormat.get(), _subTextBrush.get(), L"Facts");
                    drawContext(conflictPromptContent.sourceLabel,
                                conflictPromptContent.sourcePath,
                                conflictPromptContent.sourceMetadata,
                                task.operation == FILESYSTEM_DELETE ? L"Deleting" : L"From");
                    drawContext(
                        conflictPromptContent.destinationLabel, conflictPromptContent.destinationPath, conflictPromptContent.destinationMetadata, L"To");
                    drawTextRegion(conflictPromptContent.lastNote, _smallFormat.get(), _subTextBrush.get(), L"LastNote");
                };

                std::wstring graphEtaText;
                if (task.discoveryClosed && task.totalBytes > 0 && history && history->hasSmoothedEta && history->smoothedEtaSeconds > 0.0 &&
                    task.completedBytes <= task.totalBytes)
                {
                    const uint64_t seconds = SaturatingCeilNonNegativeToUint64(history->smoothedEtaSeconds);
                    graphEtaText           = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA, FormatDurationHms(seconds));
                }
                else if (! task.discoveryClosed && history && history->hasProvisionalEta && history->provisionalEtaSeconds > 0.0)
                {
                    // Clearly provisional: discovery is still open, so the value may rise or fall.
                    const uint64_t seconds = SaturatingCeilNonNegativeToUint64(history->provisionalEtaSeconds);
                    graphEtaText           = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ETA_PROVISIONAL, FormatDurationHms(seconds));
                }

                if (task.operation != FILESYSTEM_DELETE && (hasConflictPrompt || compactProgressPresentation.showThroughputGraph))
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
                            DrawBandwidthGraph(graphRc, task.taskId, graphHistory, graphEtaText, overlayText, rainbowMode, true);
                        }
                    }
                }
                else if (task.operation == FILESYSTEM_DELETE && (hasConflictPrompt || compactProgressPresentation.showThroughputGraph))
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
                            DrawBandwidthGraph(graphRc, task.taskId, graphHistory, {}, overlayText, rainbowMode, true);
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
                    const D2D1_RECT_F totalBarRc    = D2D1::RectF(barX, barsTop, barX + barW, barsTop + barHItem);
                    const bool hostedHasTotalBytes  = task.totalBytes > 0 && task.completedBytes <= task.totalBytes;
                    const bool hostedHasUsefulItems = task.totalItems > 1;
                    const bool hostedUseBytes       = hostedHasTotalBytes && task.completedBytes > 0;
                    const bool hostedUseItems       = ! hostedUseBytes && hostedHasUsefulItems && task.completedItems > 0;
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
                    double hostedFraction        = 0.0;
                    bool hostedDeterminate       = false;
                    if (task.totalBytes > 0 && task.completedBytes <= task.totalBytes)
                    {
                        hostedFraction    = static_cast<double>(task.completedBytes) / static_cast<double>(task.totalBytes);
                        hostedDeterminate = true;
                    }
                    else if (task.totalItems > 0)
                    {
                        hostedFraction    = static_cast<double>(std::min(task.completedItems, task.totalItems)) / static_cast<double>(task.totalItems);
                        hostedDeterminate = true;
                    }
                    HostedProgressDescriptor progressDescriptor{totalBarRc, task.taskId, std::clamp(hostedFraction, 0.0, 1.0) * 100.0, ! hostedDeterminate};
                    progressDescriptor.segmented             = task.verificationRequested;
                    progressDescriptor.primarySegmentValue   = VerificationTransferProgressPercent(task);
                    progressDescriptor.secondarySegmentValue = VerificationProofProgressPercent(task);
                    _hostedProgressDescriptors.push_back(progressDescriptor);

                }

                {
                    const float btnGapX = DipsToPixels(8.0f, _dpi);
                    const float rowW    = std::max(0.0f, contentRight - textX);
                    if (rowW > 1.0f)
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
                                const float applyBottom = applyTop + lineH;
                                buttonsTop              = applyBottom + conflictRowGapY;

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
                            }

                            const float rowY                  = buttonsTop;
                            const float rowYBottom            = rowY + buttonH;
                            const size_t primaryActionCount   = std::min(task.conflict.primaryActionCount, task.conflict.primaryActions.size());
                            const size_t overflowActionCount  = std::min(task.conflict.overflowActionCount, task.conflict.overflowActions.size());
                            const size_t visibleActionButtons = primaryActionCount + (overflowActionCount > 0u ? 1u : 0u);
                            if (visibleActionButtons > 0u)
                            {
                                const float totalGapX = btnGapX * static_cast<float>(visibleActionButtons - 1u);
                                const float btnW      = std::max(0.0f, (rowW - totalGapX) / static_cast<float>(visibleActionButtons));

                                float xBtn = textX;
                                for (size_t i = 0; i < primaryActionCount; ++i)
                                {
                                    const ConflictAction action = ConflictActionFromRaw(task.conflict.primaryActions[i]);
                                    const std::wstring label =
                                        ConflictActionText(action, static_cast<ConflictBucket>(task.conflict.bucket), task.operation);

                                    PopupButton btn{};
                                    btn.bounds     = D2D1::RectF(xBtn, rowY, xBtn + btnW, rowYBottom);
                                    btn.hit.kind   = PopupHitTest::Kind::TaskConflictAction;
                                    btn.hit.taskId = task.taskId;
                                    btn.hit.data   = static_cast<uint32_t>(RawConflictAction(action));
                                    _buttons.push_back(btn);
                                    DrawButton(btn, _buttonSmallFormat.get(), label);

                                    xBtn += btnW + btnGapX;
                                }

                                if (overflowActionCount > 0u)
                                {
                                    PopupButton moreBtn{};
                                    moreBtn.bounds     = D2D1::RectF(xBtn, rowY, xBtn + btnW, rowYBottom);
                                    moreBtn.hit.kind   = PopupHitTest::Kind::TaskConflictMore;
                                    moreBtn.hit.taskId = task.taskId;
                                    moreBtn.hit.data   = static_cast<uint32_t>(std::min<size_t>(overflowActionCount, std::numeric_limits<uint32_t>::max()));
                                    _buttons.push_back(moreBtn);
                                    DrawMenuButton(moreBtn, _buttonSmallFormat.get(), LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_MORE));
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
            }

            const float gapAfter = (rowIndex + 1u < rowCount) ? cardGap : 0.0f;
            y += taskCardH + gapAfter;
        }

        _target->PopAxisAlignedClip();
    }
    const uint64_t drawUs = capturePerf ? PerfElapsedUs(drawStartedUs) : 0u;
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
                                    it->hit.kind == PopupHitTest::Kind::FooterDensity || it->hit.kind == PopupHitTest::Kind::FooterOptions ||
                                    it->hit.kind == PopupHitTest::Kind::FooterToggleDetails;
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

    if (! fileOps || ! fileOps->HasActiveOperations())
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
        fileOps->CancelAll();
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
    std::array<RedSalamander::DxUi::MenuFlyoutItem, 2u> items{};
    items[0].kind      = RedSalamander::DxUi::MenuItemKind::Radio;
    items[0].text      = LoadStringResource(nullptr, IDS_FILEOPS_BTN_MODE_QUEUE);
    items[0].commandId = kQueueCommand;
    items[0].checked   = queueMode;
    items[1].kind      = RedSalamander::DxUi::MenuItemKind::Radio;
    items[1].text      = LoadStringResource(nullptr, IDS_FILEOPS_BTN_MODE_PARALLEL);
    items[1].commandId = kParallelCommand;
    items[1].checked   = ! queueMode;

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor =
        ResolveButtonMenuAnchor(hwnd, PopupHitTest{PopupHitTest::Kind::FooterQueueMode, 0u, 0u}, RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Above);
    const POINT point = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(RedSalamander::DxUi::ContextMenu::ShowAsync(hwnd,
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

    std::array<RedSalamander::DxUi::MenuFlyoutItem, 4u> items{};
    items[0].kind      = RedSalamander::DxUi::MenuItemKind::Toggle;
    items[0].text      = LoadStringResource(nullptr, autoDismiss ? IDS_FILEOPS_CHECK_AUTODISMISS_ON : IDS_FILEOPS_CHECK_AUTODISMISS_OFF);
    items[0].commandId = kAutoDismissCommand;
    items[0].checked   = autoDismiss;
    items[1].kind      = RedSalamander::DxUi::MenuItemKind::Separator;
    items[2].kind      = RedSalamander::DxUi::MenuItemKind::Radio;
    items[2].text      = LoadStringResource(nullptr, IDS_FILEOPS_BTN_DENSITY_COMPACT);
    items[2].commandId = kCompactCommand;
    items[2].checked   = compactDensity;
    items[3].kind      = RedSalamander::DxUi::MenuItemKind::Radio;
    items[3].text      = LoadStringResource(nullptr, IDS_FILEOPS_BTN_DENSITY_EXPANDED);
    items[3].commandId = kExpandedCommand;
    items[3].checked   = ! compactDensity;

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const auto anchor =
        ResolveButtonMenuAnchor(hwnd, PopupHitTest{PopupHitTest::Kind::FooterOptions, 0u, 0u}, RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Above);
    const POINT point = anchor.has_value() ? anchor->screenPoint : ResolveOwnerCenterScreenPoint(hwnd);
    if (anchor.has_value())
    {
        sessionCallbacks = anchor->sessionCallbacks;
    }
    static_cast<void>(RedSalamander::DxUi::ContextMenu::ShowAsync(hwnd,
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
    if (! fileOps || ! folderWindow || ! hostLifetime.lock())
    {
        return false;
    }

    uint64_t resolvedTaskId = requestedTaskId;
    FolderWindow::FileOperationState::Task* task = resolvedTaskId != 0 ? fileOps->FindTask(resolvedTaskId) : nullptr;
    if (! task)
    {
        const std::vector<TaskSnapshot> snapshot = BuildSnapshot();
        const auto isActiveFileOperation         = [](const TaskSnapshot& candidate) noexcept
        { return candidate.kind == TaskSnapshot::Kind::FileOperation && ! candidate.finished && candidate.taskId != 0; };

        const auto activeIt = std::find_if(snapshot.begin(), snapshot.end(), isActiveFileOperation);
        if (activeIt != snapshot.end())
        {
            resolvedTaskId = activeIt->taskId;
            task           = fileOps->FindTask(resolvedTaskId);
        }
    }

    if (! task)
    {
        return false;
    }

    const uint64_t taskId = resolvedTaskId;
    const uint64_t currentLimit = task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire);
    const FolderWindow::FileOperationPromptDispatchScope promptDispatch(*folderWindow);
    const auto promptResult = ShowCustomSpeedLimitPrompt(hwnd, folderWindow->GetTheme(), currentLimit);
    if (promptResult.has_value() && hostLifetime.lock())
    {
        if (FolderWindow::FileOperationState::Task* const currentTask = fileOps->FindTask(taskId))
        {
            currentTask->SetDesiredSpeedLimit(promptResult.value());
        }
    }
    if (IsWindow(hwnd) != FALSE)
    {
        Invalidate(hwnd);
    }
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

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(8u);

    const bool hasDiagnostics = ! taskIt->interruptedMoveNotice && (taskIt->warningCount > 0 || taskIt->errorCount > 0);
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

    if (taskIt->hasSourceActionPaths)
    {
        RedSalamander::DxUi::MenuFlyoutItem openSourceItem{};
        openSourceItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_OPEN_SOURCE);
        openSourceItem.commandId = static_cast<int>(kCompletedOverflowActionOpenSource);
        items.push_back(std::move(openSourceItem));
    }

    if (taskIt->exactRetainedSourceActionAvailable)
    {
        RedSalamander::DxUi::MenuFlyoutItem selectRetainedItem{};
        selectRetainedItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_SELECT_RETAINED);
        selectRetainedItem.commandId = static_cast<int>(kCompletedOverflowActionSelectRetained);
        items.push_back(std::move(selectRetainedItem));

        RedSalamander::DxUi::MenuFlyoutItem cutRetainedItem{};
        cutRetainedItem.text      = LoadStringResource(nullptr, IDS_FILEOP_BTN_CUT_RETAINED_AGAIN);
        cutRetainedItem.commandId = static_cast<int>(kCompletedOverflowActionCutRetainedAgain);
        items.push_back(std::move(cutRetainedItem));
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
            chosen == kCompletedOverflowActionOpenDestination || chosen == kCompletedOverflowActionRevealDestination ||
            chosen == kCompletedOverflowActionOpenSource || chosen == kCompletedOverflowActionSelectRetained ||
            chosen == kCompletedOverflowActionCutRetainedAgain)
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
        if (! task->_conflictArbiter.prompt.active || ! task->_conflictArbiter.prompt.buttonsPublishable)
        {
            return false;
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
        applyToAll                 = task->_conflictArbiter.prompt.applyToAllChecked;
    }

    const ConflictAction requestedAction = static_cast<ConflictAction>(rawAction);
    if (! PublishedConflictPlacementContains(conflict.overflowActions, conflict.overflowActionCount, requestedAction))
    {
        return false;
    }

    task->SubmitConflictDecision(requestedAction, applyToAll);
    Invalidate(hwnd);
    return true;
}

bool FileOperationsPopupInternal::FileOperationsPopupState::SubmitConflictPrimaryAction(HWND hwnd, uint64_t taskId, uint32_t rawAction) noexcept
{
    if (! fileOps)
    {
        return false;
    }

    FolderWindow::FileOperationState::Task* const task = fileOps->FindTask(taskId);
    if (! task)
    {
        return false;
    }

    TaskSnapshot::ConflictPromptSnapshot conflict{};
    bool applyToAll = false;
    {
        std::scoped_lock lock(task->_conflictArbiter.mutex);
        if (! task->_conflictArbiter.prompt.active || ! task->_conflictArbiter.prompt.buttonsPublishable)
        {
            return false;
        }
        CopyPublishedConflictPolicy(task->_conflictArbiter.prompt, conflict);
        applyToAll = task->_conflictArbiter.prompt.applyToAllChecked;
    }

    const ConflictAction requestedAction = static_cast<ConflictAction>(rawAction);
    if (! PublishedConflictPlacementContains(conflict.primaryActions, conflict.primaryActionCount, requestedAction))
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
        if (! task->_conflictArbiter.prompt.active || ! task->_conflictArbiter.prompt.buttonsPublishable)
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
    std::vector<RedSalamander::DxUi::MenuFlyoutItem> items;
    items.reserve(std::min(conflict.overflowActionCount, conflict.overflowActions.size()));
    std::vector<uint32_t> overflowActions;
    overflowActions.reserve(std::min(conflict.overflowActionCount, conflict.overflowActions.size()));
    for (size_t i = 0; i < conflict.overflowActionCount && i < conflict.overflowActions.size(); ++i)
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        const ConflictAction action = ConflictActionFromRaw(conflict.overflowActions[i]);
        item.text                   = ConflictActionText(action, static_cast<ConflictBucket>(conflict.bucket), task->GetOperation());
        item.commandId              = static_cast<int>(kCmdOverflowBase + static_cast<UINT>(i));
        items.push_back(std::move(item));
        overflowActions.push_back(static_cast<uint32_t>(RawConflictAction(action)));
    }

    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
    const uint32_t overflowData = static_cast<uint32_t>(std::min<size_t>(conflict.overflowActionCount, std::numeric_limits<uint32_t>::max()));
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

// C4: graphics-independent failure surface. One native text, Cancel all, and Close; no Direct2D,
// DirectWrite, DxUi host, hit-testing, task graph, queue/speed menu, or conflict action. Waiting tasks
// keep waiting and running tasks keep running unless Cancel all is confirmed.
namespace
{
constexpr INT_PTR kFailureSurfaceTextId      = 0x4C01;
constexpr INT_PTR kFailureSurfaceCancelAllId = 0x4C02;
constexpr INT_PTR kFailureSurfaceCloseId     = 0x4C03;
} // namespace

LRESULT CALLBACK FileOperationsPopupInternal::FileOperationsPopupState::FailureSurfaceButtonProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    const WNDPROC original = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCDESTROY)
    {
        if (original != nullptr)
        {
            static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(original)));
        }
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    }
    else if (msg == WM_KEYDOWN)
    {
        const HWND popup = GetParent(hwnd);
        auto* state      = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup, GWLP_USERDATA)) : nullptr;
        if (state != nullptr && state->HandleFailureSurfaceKeyDown(popup, hwnd, static_cast<UINT>(wp)))
        {
            return 0;
        }
    }
    return original != nullptr ? RedSalamander::Win32Callback::CallWindowProcNoThrow(original, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
}

bool FileOperationsPopupInternal::FileOperationsPopupState::EnterFailureSurface(HWND hwnd) noexcept
{
    if (_failureSurface)
    {
        return true;
    }
    _failureSurface = true;

    _controlsHost.Detach();
    _controlsRoot = nullptr;
    _hostedButtons.clear();
    _hostedButtonHits.clear();
    _hostedProgressBars.clear();
    _hostedProgressAnimations.clear();
    _hostedGraphs.clear();
    _hostedTooltipRegions.clear();
    _buttons.clear();
    DiscardDeviceResources();
    ShowScrollBar(hwnd, SB_VERT, FALSE);
    _scrollBarVisible = false;

    const HINSTANCE instance     = reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(hwnd, GWLP_HINSTANCE));
    const std::wstring cancelAll = LoadStringResource(nullptr, IDS_FILEOPS_BTN_CANCEL_ALL);
    const std::wstring close     = LoadStringResource(nullptr, IDS_FILEOPS_BTN_CLOSE);
    _failureText                 = CreateWindowExW(0,
                                   L"STATIC",
                                   L"",
                                   WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL,
                                   0,
                                   0,
                                   0,
                                   0,
                                   hwnd,
                                   reinterpret_cast<HMENU>(kFailureSurfaceTextId),
                                   instance,
                                   nullptr);
    _failureCancelAll            = CreateWindowExW(0,
                                        L"BUTTON",
                                        cancelAll.c_str(),
                                        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                                        0,
                                        0,
                                        0,
                                        0,
                                        hwnd,
                                        reinterpret_cast<HMENU>(kFailureSurfaceCancelAllId),
                                        instance,
                                        nullptr);
    _failureClose                = CreateWindowExW(0,
                                    L"BUTTON",
                                    close.c_str(),
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                                    0,
                                    0,
                                    0,
                                    0,
                                    hwnd,
                                    reinterpret_cast<HMENU>(kFailureSurfaceCloseId),
                                    instance,
                                    nullptr);
    if (! _failureText || ! _failureCancelAll || ! _failureClose)
    {
        // The existing host error surface carries the explanation; Cancel all stays reachable
        // from the main window's File Operations command.
        Debug::Error(L"File operations popup could not create its failure surface.");
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
        return false;
    }
    for (const HWND button : {_failureCancelAll, _failureClose})
    {
        SetWindowLongPtrW(button, GWLP_USERDATA, GetWindowLongPtrW(button, GWLP_WNDPROC));
        static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(button, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(&FailureSurfaceButtonProc)));
    }
    LayoutFailureSurface(hwnd);
    RefreshFailureSurfaceText();
    if (GetActiveWindow() == hwnd)
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
        _failureFontDpi   = _dpi;
        const HFONT font  = _failureFont ? _failureFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
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
    std::wstring text                               = LoadStringResource(nullptr, IDS_FILEOPS_FAILURE_SURFACE_TEXT);
    const std::wstring summaryText                  = FormatGlobalStatusSummaryText(summary);
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
        text.append(L"\r\n").append(LoadStringResource(nullptr, operationId)).append(L": ").append(StatusTextForTask(task, ResolveTaskStatusKind(task), nowTick));
        ++listed;
    }
    if (text != _failureTextValue)
    {
        _failureTextValue = std::move(text);
        SetWindowTextW(_failureText, _failureTextValue.c_str());
    }
}

bool FileOperationsPopupInternal::FileOperationsPopupState::HandleFailureSurfaceKeyDown(HWND hwnd, HWND focused, UINT virtualKey) noexcept
{
    if (! _failureSurface || ! _failureCancelAll || ! _failureClose)
    {
        return false;
    }
    if (virtualKey == VK_ESCAPE)
    {
        static_cast<void>(OnClose(hwnd));
        return true;
    }
    if (virtualKey == VK_TAB)
    {
        SetFocus(focused == _failureClose ? _failureCancelAll : _failureClose);
        return true;
    }
    if (virtualKey == VK_RETURN)
    {
        SendMessageW(focused == _failureCancelAll ? _failureCancelAll : _failureClose, BM_CLICK, 0, 0);
        return true;
    }
    return false;
}

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    RefreshLocalizedFooterText();

    if (folderWindow && hostLifetime.lock())
    {
        const AppTheme& theme = folderWindow->GetTheme();
        const auto palette    = MakeAppThemeDxPalette(theme, theme.windowBackground);
        _reducedMotion        = palette.reducedMotion;
        ApplyWindowChromeTheme(hwnd, theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd);
        RedSalamander::DxUi::WindowHost::AttachOptions options{};
        options.presentationMode = RedSalamander::DxUi::WindowHost::PresentationMode::CompositionSwapChain;
        bool forceAttachFailure  = false;
#ifdef ENABLE_TESTS
        unsigned int remainingForcedFailures = g_fileOperationsDxUiHostAttachForcedFailures.load(std::memory_order_acquire);
        while (remainingForcedFailures > 0u)
        {
            if (g_fileOperationsDxUiHostAttachForcedFailures.compare_exchange_weak(
                    remainingForcedFailures, remainingForcedFailures - 1u, std::memory_order_acq_rel, std::memory_order_acquire))
            {
                forceAttachFailure = true;
                break;
            }
        }
#endif
        if (! forceAttachFailure && _controlsHost.Attach(hwnd, options))
        {
            _controlsHost.SetTheme(palette);
            auto root     = std::make_unique<RedSalamander::DxUi::Panel>();
            _controlsRoot = root.get();
            _controlsHost.SetRoot(std::move(root));
            _controlsHost.SetOnEscape([this, hwnd]
            {
                if (SubmitPublishedConflictEscapeAction(hwnd))
                {
                    return true;
                }
                static_cast<void>(OnClose(hwnd));
                return true;
            });
        }
        else
        {
            if (! forceAttachFailure)
            {
                Debug::Error(L"File operations popup could not attach its DxUi control host.");
            }
            static_cast<void>(EnterFailureSurface(hwnd));
        }
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
        const auto palette    = MakeAppThemeDxPalette(theme, theme.windowBackground);
        _reducedMotion        = palette.reducedMotion;
        _controlsHost.SetTheme(palette);
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
    _controlsHost.Detach();
    _controlsRoot     = nullptr;
    _failureText      = nullptr;
    _failureCancelAll = nullptr;
    _failureClose     = nullptr;
    _failureFont.reset();
    _hostedButtons.clear();
    _hostedButtonHits.clear();
    _hostedProgressBars.clear();
    _hostedProgressAnimations.clear();
    _hostedGraphs.clear();
    _hostedTooltipRegions.clear();
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

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnSize(HWND hwnd, WPARAM sizeType, UINT width, UINT height) noexcept
{
    _clientSize.cx = static_cast<LONG>(width);
    _clientSize.cy = static_cast<LONG>(height);
    if (_failureSurface)
    {
        LayoutFailureSurface(hwnd);
    }

    if (_target && width > 0u && height > 0u)
    {
        _target->Resize(D2D1::SizeU(width, height));
    }

    bool hostedHandled = false;
    static_cast<void>(_controlsHost.HandleMessage(hwnd, WM_SIZE, sizeType, MAKELPARAM(width, height), hostedHandled));

    UpdateLastPopupRect(hwnd);
    Invalidate(hwnd);
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
    _graphTrailingLabelFormat.reset();
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
        const ULONGLONG nowTick = GetTickCount64();
        if (_uiMotionTimerDueTick != 0u && nowTick >= _uiMotionTimerDueTick)
        {
            _uiMotionTimerDueTick = 0u;
            SetTimer(hwnd, kFileOperationsPopupTimerId, kFileOperationsPopupTimerIntervalMs, nullptr);
        }

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
        if (_failureSurface)
        {
            UpdateTaskbarProgress(hwnd);
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

LRESULT FileOperationsPopupInternal::FileOperationsPopupState::OnActivatedHit(HWND hwnd, const PopupHitTest& hit) noexcept
{
    if (! hostLifetime.lock())
    {
        DestroyWindow(hwnd);
        return 0;
    }
    if (_failureSurface && hit.kind != PopupHitTest::Kind::FooterCancelAll)
    {
        // C4: the failure surface offers Cancel all and Close only; no other action is reachable.
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
        _completedGroupExpanded     = ! _completedGroupExpanded;
        _maxAutoSizedWindowHeight   = 0;
        _lastAutoSizedContentHeight = -1.0f;
        ArmUiMotionTimer(hwnd);
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
        static_cast<void>(SubmitConflictPrimaryAction(hwnd, hit.taskId, hit.data));
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
        return SubmitConflictOverflowAction(hwnd, payload->taskId, payload->data) ? 1 : 0;
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
        payload->kind == PopupHitTest::Kind::TaskToggleCollapse || payload->kind == PopupHitTest::Kind::TaskStartNow ||
        payload->kind == PopupHitTest::Kind::TaskQueueMoveUp || payload->kind == PopupHitTest::Kind::TaskQueueMoveDown)
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
    result.hostedActionControlCount               = _hostedButtons.size();
    result.hostedProgressControlCount             = _hostedProgressBars.size();
    result.hostedGraphControlCount                = _hostedGraphs.size();
    result.hostedTooltipRegionCount               = _hostedTooltipRegions.size();
    result.completedGroupBadgeTooltipsInformative = ! _hostedTooltipDescriptors.empty() && _hostedTooltipDescriptors.size() == _hostedTooltipRegions.size();
    for (size_t index = 0u; result.completedGroupBadgeTooltipsInformative && index < _hostedTooltipDescriptors.size(); ++index)
    {
        const HostedTooltipDescriptor& descriptor = _hostedTooltipDescriptors[index];
        const RedSalamander::DxUi::Label* region  = _hostedTooltipRegions[index];
        result.completedGroupBadgeTooltipsInformative =
            region && ! descriptor.text.empty() && descriptor.text.find_first_not_of(L"0123456789 \t") != std::wstring::npos &&
            region->GetTooltipText() == descriptor.text && region->GetAccessibleName() == descriptor.text && region->GetAccessibleHelpText() == descriptor.text;
    }
    result.hostedAccessibilityNotificationCount = _hostedAccessibilityNotificationCount;
    if (const auto* focusedControl = _controlsHost.GetFocusControl())
    {
        result.hostedFocusedAutomationId = focusedControl->GetAccessibleAutomationId();
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
        result.taskId = taskIt->taskId;
        result.found  = true;
        if (const auto animation = _taskHeightAnimations.find(taskIt->taskId); animation != _taskHeightAnimations.end())
        {
            result.selectedTaskHeightMotionActive = scalarAnimationActive(animation->second);
            result.layoutMotionActive             = result.layoutMotionActive || result.selectedTaskHeightMotionActive;
        }
        result.taskDiscoveryAheadActive                              = taskIt->discoveryAheadActive;
        const WholeTaskProgressPresentation taskProgressPresentation = ResolveWholeTaskProgressPresentation(*taskIt);
        result.taskDiscoveryIndicatorVisible                         = taskProgressPresentation.showDiscoveryActivity;
        result.taskTransferProgressVisible                           = taskProgressPresentation.showTransferProgress;
        result.taskTransferProgressDeterminate                       = result.taskTransferProgressVisible && taskProgressPresentation.determinate;
        result.taskTransferProgressProvisional                       = result.taskTransferProgressDeterminate && taskProgressPresentation.provisional;
        result.taskTransferProgressMarqueeVisible                    = taskProgressPresentation.showTransferMarquee;
        const TaskStatusKind status       = taskIt->statusKind != TaskStatusKind::None ? taskIt->statusKind : ResolveTaskStatusKind(*taskIt);
        result.taskStatusKind             = status;
        result.taskStatusActiveStateCount = SurfacedTaskStatusCount(status);
        bool graphShowAnimation           = false;
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
                    promptBottom = std::max(promptBottom, descriptor.bounds.bottom);
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
        std::array<TerminalAxisBadge, 3u> terminalBadges{};
        const size_t terminalBadgeCount = BuildTerminalAxisBadges(*taskIt, terminalBadges);
        for (size_t badgeIndex = 0u; badgeIndex < terminalBadgeCount; ++badgeIndex)
        {
            result.taskVerificationBadgeVisible   = result.taskVerificationBadgeVisible || terminalBadges[badgeIndex].automationSuffix == L"Verification";
            result.taskRetainedSourceBadgeVisible = result.taskRetainedSourceBadgeVisible || terminalBadges[badgeIndex].automationSuffix == L"SourceRetained";
            result.taskUnknownSourceBadgeVisible  = result.taskUnknownSourceBadgeVisible || terminalBadges[badgeIndex].automationSuffix == L"SourceUnknown";
        }
        result.taskTerminalAxisBadgesColorBlindSafe =
            terminalBadgeCount == 0u || std::ranges::all_of(terminalBadges.begin(),
                                                            terminalBadges.begin() + static_cast<std::ptrdiff_t>(terminalBadgeCount),
                                                            [](const TerminalAxisBadge& badge) noexcept { return ! badge.text.empty(); });
        result.taskStatusGlyphSignalVisible = StatusIsOk(status) || StatusIsWarning(status) || StatusIsError(status);
        result.taskStatusTextSignalVisible  = ! statusChipText.empty() || ! StatusTextForTask(*taskIt, status, GetTickCount64()).empty();
        result.taskStatusColorBlindSafeEncoding =
            statusTone == PopupStatusVisualTone::None || result.taskStatusGlyphSignalVisible || result.taskStatusTextSignalVisible;
        result.completedFailedItemsActionVisible       = taskIt->finished && (taskIt->warningCount > 0 || taskIt->errorCount > 0);
        result.completedOpenDestinationActionVisible   = CompletedTaskCanUseDestinationActions(*taskIt);
        result.completedRevealDestinationActionVisible = ResolveCompletedTaskRevealLocation(*taskIt).has_value();
        result.taskHiddenByCompletedGroup              = completedGroupVisible && ! _completedGroupExpanded && IsCompletedGroupTask(*taskIt);
        result.taskCollapsed                           = IsTaskCollapsed(taskIt->taskId);
        result.taskCompactRow                          = IsTaskCollapsedForDisplay(taskIt->taskId, result.popupCompactDensity);
        result.taskAutoCollapsedOnCompletion           = taskIt->finished && result.taskCollapsed;
        result.taskCompactProgressVisible =
            result.taskCompactRow && taskIt->kind == TaskSnapshot::Kind::FileOperation &&
            ((taskIt->finished && TaskHasKnownCompactProgress(*taskIt)) ||
             (taskProgressPresentation.showTransferProgress && (TaskHasKnownCompactProgress(*taskIt) || taskProgressPresentation.showTransferMarquee)));
        result.taskInlineSpeedRowVisible = false;
        result.taskInlineEtaRowVisible   = false;
        result.taskExpandedBaseHeightDip = taskIt->finished ? (TaskHasTerminalAxisBadges(*taskIt) ? 202.0f : 178.0f)
                                           : (taskIt->operation == FILESYSTEM_COPY || taskIt->operation == FILESYSTEM_MOVE)
                                               ? kFileOperationsExpandedTransferCardHeightDip
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
        }
        const auto graphIt = std::find_if(_hostedGraphDescriptors.begin(),
                                          _hostedGraphDescriptors.end(),
                                          [&result](const HostedGraphDescriptor& descriptor) noexcept { return descriptor.taskId == result.taskId; });
        if (graphIt != _hostedGraphDescriptors.end())
        {
            result.graphCurrentBandwidthLabelVisible = ! graphIt->currentBandwidthText.empty();
            result.graphEtaLabelVisible              = ! graphIt->currentBandwidthTrailingText.empty();
            result.graphEtaLabelRightAligned         = result.graphEtaLabelVisible;
        }

        result.failureSurfaceActive     = _failureSurface;
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

    result.visibleButtonCount = _buttons.size();
    for (size_t i = 0; i < _buttons.size(); ++i)
    {
        const PopupButton& button = _buttons[i];
        switch (button.hit.kind)
        {
            case PopupHitTest::Kind::FooterCancelAll:
            case PopupHitTest::Kind::FooterPauseResumeAll:
            case PopupHitTest::Kind::FooterAutoDismiss:
            case PopupHitTest::Kind::FooterDensity:
            case PopupHitTest::Kind::FooterOptions: ++result.footerVisibleButtonCount; break;
            case PopupHitTest::Kind::FooterToggleDetails:
                ++result.footerVisibleButtonCount;
                result.footerDetailsToggleUsesDisclosureChrome =
                    i < _hostedButtons.size() && _hostedButtons[i] && _hostedButtons[i]->GetVariant() == RedSalamander::DxUi::ButtonVariant::Disclosure;
                break;
            case PopupHitTest::Kind::FooterQueueMode:
                ++result.footerVisibleButtonCount;
                result.footerQueueModeUsesSelectorChrome =
                    i < _hostedButtons.size() && _hostedButtons[i] && _hostedButtons[i]->GetVariant() == RedSalamander::DxUi::ButtonVariant::Selector;
                break;
            case PopupHitTest::Kind::CompletedGroupToggle: result.completedGroupToggleVisible = true; break;
            case PopupHitTest::Kind::CompletedGroupClear:
                result.completedGroupClearVisible      = true;
                result.completedGroupClearTooltipEmpty = i < _hostedButtons.size() && _hostedButtons[i] && _hostedButtons[i]->GetTooltipText().empty();
                break;
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
                    result.taskSpeedLimitUsesSelectorChrome =
                        i < _hostedButtons.size() && _hostedButtons[i] && _hostedButtons[i]->GetVariant() == RedSalamander::DxUi::ButtonVariant::Selector;
                    result.taskSpeedLimitShowsCurrentValue =
                        i < _hostedButtons.size() && _hostedButtons[i] && taskIt != taskSnapshot.end() &&
                        _hostedButtons[i]->GetText() == FormatSpeedLimitSelectorText(taskIt->desiredSpeedLimitBytesPerSecond);
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

    if ((msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN) && HandleHostedKeyDown(hwnd, static_cast<UINT>(wp)))
    {
        return 0;
    }

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
            bool hostedHandled = false;
            static_cast<void>(_controlsHost.HandleMessage(hwnd, msg, wp, lp, hostedHandled));
            if (capturePerf)
            {
                Debug::Perf::Emit(L"FileOps.Popup.WmPaintUs", L"", PerfElapsedUs(startedUs), 0u, 0u, S_OK);
            }
            return 0;
        }
        case WM_SIZE: return OnSize(hwnd, wp, LOWORD(lp), HIWORD(lp));
        case WM_MOVE: return OnMove(hwnd);
        case WM_GETMINMAXINFO: return OnGetMinMaxInfo(hwnd, reinterpret_cast<MINMAXINFO*>(lp));
        case WM_ENTERSIZEMOVE: return OnEnterSizeMove(hwnd);
        case WM_EXITSIZEMOVE: return OnExitSizeMove(hwnd);
        case WM_TIMER: return OnTimer(hwnd, static_cast<UINT_PTR>(wp));
        case WM_VSCROLL: return OnVScroll(hwnd, static_cast<UINT>(LOWORD(wp)));
        case WM_MOUSEWHEEL: return OnMouseWheel(hwnd, GET_WHEEL_DELTA_WPARAM(wp));
        case WM_COMMAND:
            if (_failureSurface && HIWORD(wp) == BN_CLICKED)
            {
                const HWND control = reinterpret_cast<HWND>(lp);
                if (control == _failureCancelAll)
                {
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
        case WM_KEYDOWN:
            if (_failureSurface && HandleFailureSurfaceKeyDown(hwnd, GetFocus(), static_cast<UINT>(wp)))
            {
                return 0;
            }
            break;
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

RedSalamander::DxUi::ButtonVariant DebugResolveFileOperationsHostedButtonVariant(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept
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
        AppendResampledRateSamples(history, kRateSampleBucketMs, 400.0, 0.0, -1.0f);
    }

    out.taskId = 1;
    out.found  = true;
    PopulateGraphHueDebugSummary(history, out);
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
        const D2D1_COLOR_F graphColor = RedSalamander::DxUi::ThroughputGraphColorFromHue(graphStream.assignedHue, darkTheme.dark);
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
    uint8_t slotFor1 = RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot;
    uint8_t slotFor2 = RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot;
    uint8_t slotFor3 = RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot;
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

bool DebugValidateFileOperationsVerificationPresentation() noexcept
{
    TaskSnapshot task{};
    task.taskId                     = 1u;
    task.operation                  = FILESYSTEM_COPY;
    task.started                    = true;
    task.discoveryClosed            = true;
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

    task.discoveryClosed = false;
    if (TaskHasKnownCompactProgress(task))
    {
        return false;
    }
    task.discoveryClosed = true;

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
    streamingTask.discoveryClosed                             = true;
    streamingTask.totalBytes                                  = 240u;
    const WholeTaskProgressPresentation closedProgress        = ResolveWholeTaskProgressPresentation(streamingTask);
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

    RateSnapshot providerProofOnly{};
    providerProofOnly.verificationCompletedBytes = 100u;
    providerProofOnly.verificationReadBytes      = 0u;
    RateHistory providerProofHistory{};
    const RateByteDeltas providerProofDeltas = ComputeRateByteDeltas(providerProofOnly, providerProofHistory);
    providerProofOnly.verificationReadBytes  = 40u;
    const RateByteDeltas hostReadbackDeltas  = ComputeRateByteDeltas(providerProofOnly, providerProofHistory);
    if (providerProofDeltas.verificationRead != 0u || hostReadbackDeltas.verificationRead != 40u)
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
