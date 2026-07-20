#include "FolderWindow.FileOperations.State.Private.h"

#include "ConnectionProfileUtils.h"
#include "FileSystemPathIdentity.h"
#include "FolderWindow.FileOperations.IssuesPane.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "SessionState.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "SettingsStore.h"

#include <algorithm>
#include <array>
#include <cassert>
#include <chrono>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <deque>
#include <functional>
#include <iterator>
#include <set>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <utility>

#include <psapi.h>
#include <shellapi.h>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

using namespace FolderWindowFileOperationsStateInternal;

bool FolderWindow::FileOperationState::OpenDiagnosticsLogForTask(uint64_t taskId) noexcept
{
    FlushDiagnostics(true);

    std::filesystem::path logPath;
    {
        std::scoped_lock lock(_mutex);
        for (const auto& summary : _completedTasks)
        {
            if (summary.taskId != taskId)
            {
                continue;
            }

            logPath = summary.diagnosticsLogPath;
            break;
        }
    }

    std::error_code ec;
    if (! logPath.empty() && ! std::filesystem::exists(logPath, ec))
    {
        logPath.clear();
    }

    if (logPath.empty())
    {
        std::scoped_lock lock(_mutex);
        logPath = GetLatestDiagnosticsLogPathUnlocked();
    }

    ec.clear();
    if (logPath.empty() || ! std::filesystem::exists(logPath, ec))
    {
        return false;
    }

    HINSTANCE hinst = ShellExecuteW(_owner.GetHwnd(), L"open", logPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(hinst) > 32;
}

bool FolderWindow::FileOperationState::ExportTaskIssuesReport(uint64_t taskId, std::filesystem::path* reportPathOut, bool openAfterExport) noexcept
{
    if (reportPathOut)
    {
        reportPathOut->clear();
    }

    FlushDiagnostics(true);

    CompletedTaskSummary summary{};
    bool found = false;
    {
        std::scoped_lock lock(_mutex);
        for (const auto& candidate : _completedTasks)
        {
            if (candidate.taskId != taskId)
            {
                continue;
            }

            summary = candidate;
            found   = true;
            break;
        }
    }

    if (! found)
    {
        return false;
    }

    if (summary.issueDiagnostics.empty() && summary.warningCount == 0 && summary.errorCount == 0)
    {
        return false;
    }

    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
    if (logsDir.empty())
    {
        return false;
    }

    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);

    wchar_t fileName[128]{};
    constexpr size_t fileNameMax                            = (sizeof(fileName) / sizeof(fileName[0])) - 1;
    const auto r                                            = std::format_to_n(fileName,
                                                                               fileNameMax,
                                                                               L"{}Task{}-{:04}{:02}{:02}-{:02}{:02}{:02}{:03}{}",
                                                                               kDiagnosticsIssueReportPrefix,
                                                                               static_cast<unsigned long long>(taskId),
                                                                               static_cast<unsigned>(localNow.wYear),
                                                                               static_cast<unsigned>(localNow.wMonth),
                                                                               static_cast<unsigned>(localNow.wDay),
                                                                               static_cast<unsigned>(localNow.wHour),
                                                                               static_cast<unsigned>(localNow.wMinute),
                                                                               static_cast<unsigned>(localNow.wSecond),
                                                                               static_cast<unsigned>(localNow.wMilliseconds),
                                                                               kDiagnosticsIssueReportExtension);
    fileName[(r.size < fileNameMax) ? r.size : fileNameMax] = L'\0';
    if (r.size > fileNameMax)
    {
        return false;
    }

    const std::filesystem::path reportPath = logsDir / fileName;

    std::error_code ec;
    std::filesystem::create_directories(logsDir, ec);
    if (ec)
    {
        return false;
    }

    wil::unique_handle file(CreateFileW(
        reportPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    constexpr wchar_t bom = 0xFEFF;
    DWORD written         = 0;
    if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr) || written != sizeof(bom))
    {
        return false;
    }

    const std::wstring header =
        std::format(L"Task {:#x} ({})\r\nResult: 0x{:08X}\r\nWarnings: {:L}  Errors: {:L}\r\nCompleted items: {:L}/{:L}\r\nCompleted bytes: {:L}/{:L}\r\nFrom: "
                    L"{}\r\nTo: {}\r\n\r\nTime\tSeverity\tHRESULT\tStatus text\tCategory\tMessage\tSource\tDestination\r\n",
                    static_cast<unsigned long long>(summary.taskId),
                    OperationToString(summary.operation),
                    static_cast<unsigned long>(summary.resultHr),
                    summary.warningCount,
                    summary.errorCount,
                    summary.completedItems,
                    summary.totalItems,
                    summary.completedBytes,
                    summary.totalBytes,
                    summary.sourcePath.empty() ? L"-" : summary.sourcePath,
                    summary.destinationPath.empty() ? L"-" : summary.destinationPath);
    const size_t headerBytes = header.size() * sizeof(wchar_t);
    if (headerBytes > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
    {
        return false;
    }
    if (! WriteFile(file.get(), header.data(), static_cast<DWORD>(headerBytes), &written, nullptr) || written != static_cast<DWORD>(headerBytes))
    {
        return false;
    }

    for (const TaskDiagnosticEntry& issue : summary.issueDiagnostics)
    {
        const std::wstring statusText = EscapeDiagnosticField(FormatDiagnosticStatusText(issue.status));
        const std::wstring line       = std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}\t{}\t0x{:08X}\t{}\t{}\t{}\t{}\t{}\r\n",
                                                    static_cast<unsigned>(issue.localTime.wYear),
                                                    static_cast<unsigned>(issue.localTime.wMonth),
                                                    static_cast<unsigned>(issue.localTime.wDay),
                                                    static_cast<unsigned>(issue.localTime.wHour),
                                                    static_cast<unsigned>(issue.localTime.wMinute),
                                                    static_cast<unsigned>(issue.localTime.wSecond),
                                                    static_cast<unsigned>(issue.localTime.wMilliseconds),
                                                    DiagnosticSeverityToString(issue.severity),
                                                    static_cast<unsigned long>(issue.status),
                                                    statusText,
                                                    EscapeDiagnosticField(issue.category),
                                                    EscapeDiagnosticField(issue.message),
                                                    EscapeDiagnosticField(issue.sourcePath),
                                                    EscapeDiagnosticField(issue.destinationPath));

        const size_t bytesToWrite = line.size() * sizeof(wchar_t);
        if (bytesToWrite > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
        {
            continue;
        }

        if (! WriteFile(file.get(), line.data(), static_cast<DWORD>(bytesToWrite), &written, nullptr) || written != static_cast<DWORD>(bytesToWrite))
        {
            return false;
        }
    }

    if (reportPathOut)
    {
        *reportPathOut = reportPath;
    }

    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    CleanupDiagnosticsFilesInDirectory(
        logsDir, kDiagnosticsIssueReportPrefix, kDiagnosticsIssueReportExtension, diagnosticsSettings.maxDiagnosticsIssueReportFiles);

    if (! openAfterExport)
    {
        return true;
    }

    HINSTANCE hinst = ShellExecuteW(_owner.GetHwnd(), L"open", reportPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(hinst) > 32;
}

void FolderWindow::FileOperationState::ToggleIssuesPane() noexcept
{
    HWND pane = nullptr;
    {
        std::scoped_lock lock(_mutex);
        pane = _issuesPane.get();
        if (pane && IsWindow(pane) == FALSE)
        {
            static_cast<void>(_issuesPane.release());
            pane = nullptr;
        }
    }

    if (pane)
    {
        if (IsWindowVisible(pane))
        {
            static_cast<void>(SendMessageW(pane, WM_CLOSE, 0, 0));
        }
        else
        {
            ShowWindow(pane, SW_SHOW);
            SetForegroundWindow(pane);
            PostMessageW(pane, WM_THEMECHANGED, 0, 0);
        }
        return;
    }

    HWND ownerWindow = _owner.GetHwnd();
    if (ownerWindow)
    {
        HWND rootWindow = GetAncestor(ownerWindow, GA_ROOT);
        if (rootWindow)
        {
            ownerWindow = rootWindow;
        }
    }

    if (! ownerWindow)
    {
        ownerWindow = GetParent(_owner.GetHwnd());
        if (! ownerWindow)
        {
            ownerWindow = _owner.GetHwnd();
        }
    }

    std::weak_ptr<void> uiLifetime;
    {
        std::scoped_lock lock(_mutex);
        uiLifetime = _uiLifetime;
    }

    HWND createdPane = FileOperationsIssuesPane::Create(this, &_owner, ownerWindow, std::move(uiLifetime));
    if (! createdPane)
    {
        return;
    }

    {
        std::scoped_lock lock(_mutex);
        _issuesPane.reset(createdPane);
    }
}

bool FolderWindow::FileOperationState::IsIssuesPaneVisible() noexcept
{
    std::scoped_lock lock(_mutex);
    return _issuesPane && IsWindowVisible(_issuesPane.get()) != FALSE;
}

bool FolderWindow::FileOperationState::TryGetIssuesPanePlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept
{
    outRect      = RECT{};
    outMaximized = false;

    if (! _owner._settings)
    {
        return false;
    }

    const std::wstring windowId(kFileOpsIssuesPaneWindowId);
    const auto it = _owner._settings->windows.find(windowId);
    if (it == _owner._settings->windows.end())
    {
        return false;
    }

    const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, currentDpi);
    outRect.left                                       = normalized.bounds.x;
    outRect.top                                        = normalized.bounds.y;
    outRect.right                                      = normalized.bounds.x + std::max(1, normalized.bounds.width);
    outRect.bottom                                     = normalized.bounds.y + std::max(1, normalized.bounds.height);
    outMaximized                                       = normalized.state == Common::Settings::WindowState::Maximized;
    return true;
}

void FolderWindow::FileOperationState::SaveIssuesPanePlacement(HWND hwnd) noexcept
{
    if (! hwnd || ! _owner._settings || IsIconic(hwnd))
    {
        return;
    }

    WINDOWPLACEMENT placement{};
    placement.length = sizeof(placement);
    if (! GetWindowPlacement(hwnd, &placement))
    {
        return;
    }

    Common::Settings::WindowPlacement saved{};
    saved.state         = placement.showCmd == SW_SHOWMAXIMIZED ? Common::Settings::WindowState::Maximized : Common::Settings::WindowState::Normal;
    saved.bounds.x      = placement.rcNormalPosition.left;
    saved.bounds.y      = placement.rcNormalPosition.top;
    saved.bounds.width  = std::max(1, static_cast<int>(placement.rcNormalPosition.right - placement.rcNormalPosition.left));
    saved.bounds.height = std::max(1, static_cast<int>(placement.rcNormalPosition.bottom - placement.rcNormalPosition.top));
    saved.dpi           = GetDpiForWindow(hwnd);

    _owner._settings->windows[std::wstring(kFileOpsIssuesPaneWindowId)] = std::move(saved);
}

bool FolderWindow::FileOperationState::TryGetIssuesPaneViewState(std::wstring& outSortColumnId,
                                                                 bool& outSortDescending,
                                                                 std::vector<Common::Settings::GridColumnLayoutEntry>& outGridLayout) const noexcept
{
    outSortColumnId.clear();
    outSortDescending = false;
    outGridLayout.clear();

    if (! _owner._settings || ! _owner._settings->fileOperations.has_value())
    {
        return false;
    }

    const auto& fileOperations = _owner._settings->fileOperations.value();
    outSortColumnId            = fileOperations.issuesPaneSortColumnId;
    outSortDescending          = fileOperations.issuesPaneSortDescending;
    outGridLayout              = fileOperations.issuesPaneGridLayout;
    return ! outSortColumnId.empty() || ! outGridLayout.empty();
}

void FolderWindow::FileOperationState::SaveIssuesPaneViewState(std::wstring_view sortColumnId,
                                                               bool sortDescending,
                                                               const std::vector<Common::Settings::GridColumnLayoutEntry>& gridLayout) noexcept
{
    if (! _owner._settings)
    {
        return;
    }

    if (! _owner._settings->fileOperations.has_value())
    {
        _owner._settings->fileOperations.emplace();
    }

    Common::Settings::FileOperationsSettings& fileOperations = _owner._settings->fileOperations.value();
    fileOperations.issuesPaneSortColumnId.assign(sortColumnId);
    fileOperations.issuesPaneSortDescending = ! fileOperations.issuesPaneSortColumnId.empty() && sortDescending;
    fileOperations.issuesPaneGridLayout     = gridLayout;
}

namespace
{
[[nodiscard]] bool TryGetFileOperationsWindowPlacement(const Common::Settings::Settings* settings,
                                                       std::wstring_view windowId,
                                                       RECT& outRect,
                                                       bool& outMaximized,
                                                       UINT currentDpi) noexcept
{
    outRect      = RECT{};
    outMaximized = false;

    if (! settings)
    {
        return false;
    }

    const auto it = settings->windows.find(std::wstring(windowId));
    if (it == settings->windows.end())
    {
        return false;
    }

    const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, currentDpi);
    outRect.left                                       = normalized.bounds.x;
    outRect.top                                        = normalized.bounds.y;
    outRect.right                                      = normalized.bounds.x + std::max(1, normalized.bounds.width);
    outRect.bottom                                     = normalized.bounds.y + std::max(1, normalized.bounds.height);
    outMaximized                                       = normalized.state == Common::Settings::WindowState::Maximized;
    return true;
}

void SaveFileOperationsWindowPlacement(Common::Settings::Settings* settings, std::wstring_view windowId, HWND hwnd) noexcept
{
    if (! settings || windowId.empty() || ! hwnd || IsIconic(hwnd))
    {
        return;
    }

    WINDOWPLACEMENT placement{};
    placement.length = sizeof(placement);
    if (! GetWindowPlacement(hwnd, &placement))
    {
        return;
    }

    Common::Settings::WindowPlacement saved{};
    saved.state         = placement.showCmd == SW_SHOWMAXIMIZED ? Common::Settings::WindowState::Maximized : Common::Settings::WindowState::Normal;
    saved.bounds.x      = placement.rcNormalPosition.left;
    saved.bounds.y      = placement.rcNormalPosition.top;
    saved.bounds.width  = std::max(1, static_cast<int>(placement.rcNormalPosition.right - placement.rcNormalPosition.left));
    saved.bounds.height = std::max(1, static_cast<int>(placement.rcNormalPosition.bottom - placement.rcNormalPosition.top));
    saved.dpi           = GetDpiForWindow(hwnd);

    settings->windows[std::wstring(windowId)] = std::move(saved);
}
} // namespace

bool FolderWindow::FileOperationState::TryGetPopupPlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept
{
    return TryGetFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupWindowId, outRect, outMaximized, currentDpi);
}

bool FolderWindow::FileOperationState::TryGetPopupExpandedPlacement(RECT& outRect, UINT currentDpi) const noexcept
{
    bool maximized = false;
    return TryGetFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupExpandedWindowId, outRect, maximized, currentDpi);
}

void FolderWindow::FileOperationState::SavePopupPlacement(HWND hwnd) noexcept
{
    SaveFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupWindowId, hwnd);
    if (! GetPopupFooterOnly())
    {
        SaveFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupExpandedWindowId, hwnd);
    }
}

void FolderWindow::FileOperationState::SavePopupExpandedPlacement(HWND hwnd) noexcept
{
    SaveFileOperationsWindowPlacement(_owner._settings, kFileOpsPopupExpandedWindowId, hwnd);
}

void FolderWindow::FileOperationState::OnPopupDestroyed(HWND hwnd) noexcept
{
    if (hwnd && _owner._settings)
    {
        SavePopupPlacement(hwnd);
        QueueSettingsSave(L"popup destroy");
    }

    std::scoped_lock lock(_mutex);
    if (_popup.get() == hwnd)
    {
        static_cast<void>(_popup.release());
    }
}

void FolderWindow::FileOperationState::OnIssuesPaneDestroyed(HWND hwnd) noexcept
{
    if (hwnd && _owner._settings)
    {
        SaveIssuesPanePlacement(hwnd);
        QueueSettingsSave(L"issues pane destroy");
    }

    std::scoped_lock lock(_mutex);
    if (_issuesPane.get() == hwnd)
    {
        static_cast<void>(_issuesPane.release());
    }
}

void FolderWindow::FileOperationState::UpdateLastPopupRect(const RECT& rect) noexcept
{
    std::scoped_lock lock(_mutex);
    _lastPopupRect = rect;
}

std::optional<RECT> FolderWindow::FileOperationState::GetLastPopupRect() noexcept
{
    std::scoped_lock lock(_mutex);
    return _lastPopupRect;
}

std::filesystem::path FolderWindow::FileOperationState::GetDiagnosticsLogDirectory() noexcept
{
    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kFileOpsAppId);
    if (settingsPath.empty())
    {
        return {};
    }

    const std::filesystem::path settingsDir = settingsPath.parent_path();
    if (settingsDir.empty())
    {
        return {};
    }

    const std::filesystem::path appRootDir = settingsDir.parent_path();
    if (appRootDir.empty())
    {
        return settingsDir / L"Logs";
    }

    // Keep diagnostics logs as a top-level sibling of Settings/Crashes.
    return appRootDir / L"Logs";
}

std::filesystem::path FolderWindow::FileOperationState::GetDiagnosticsLogPathForDate(const SYSTEMTIME& localTime) noexcept
{
    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
    if (logsDir.empty())
    {
        return {};
    }

    wchar_t fileName[64]{};
    constexpr size_t fileNameMax                            = (sizeof(fileName) / sizeof(fileName[0])) - 1;
    const auto r                                            = std::format_to_n(fileName,
                                                                               fileNameMax,
                                                                               L"{}{:04}{:02}{:02}{}",
                                                                               kDiagnosticsLogPrefix,
                                                                               static_cast<unsigned>(localTime.wYear),
                                                                               static_cast<unsigned>(localTime.wMonth),
                                                                               static_cast<unsigned>(localTime.wDay),
                                                                               kDiagnosticsLogExtension);
    fileName[(r.size < fileNameMax) ? r.size : fileNameMax] = L'\0';
    if (r.size > fileNameMax)
    {
        return {};
    }

    return logsDir / fileName;
}

std::filesystem::path FolderWindow::FileOperationState::GetLatestDiagnosticsLogPathUnlocked() const noexcept
{
    const std::filesystem::path logsDir = GetDiagnosticsLogDirectory();
    if (logsDir.empty())
    {
        return {};
    }

    std::error_code ec;
    std::filesystem::path newestPath;
    for (std::filesystem::directory_iterator it(logsDir, ec), end; ! ec && it != end; it.increment(ec))
    {
        const auto& de = *it;
        if (! de.is_regular_file(ec))
        {
            continue;
        }

        const std::wstring fileNameText = de.path().filename().wstring();
        if (fileNameText.size() < (kDiagnosticsLogPrefix.size() + kDiagnosticsLogExtension.size()))
        {
            continue;
        }
        if (fileNameText.rfind(kDiagnosticsLogPrefix.data(), 0) != 0)
        {
            continue;
        }
        if (de.path().extension().wstring() != kDiagnosticsLogExtension)
        {
            continue;
        }

        if (newestPath.empty() || de.path().filename().wstring() > newestPath.filename().wstring())
        {
            newestPath = de.path();
        }
    }

    return newestPath;
}

void FolderWindow::FileOperationState::FlushDiagnostics(bool force) noexcept
{
    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    std::vector<TaskDiagnosticEntry> pending;
    ULONGLONG nowTick = GetTickCount64();

    {
        std::scoped_lock lock(_diagnosticsMutex);

        const bool flushIntervalReached =
            _lastDiagnosticsFlushTick == 0 ||
            (nowTick >= _lastDiagnosticsFlushTick && (nowTick - _lastDiagnosticsFlushTick) >= diagnosticsSettings.diagnosticsFlushIntervalMs);

        if (! force && ! flushIntervalReached && _diagnosticsPendingFlush.size() < diagnosticsSettings.maxDiagnosticsPerFlush)
        {
            return;
        }

        if (_diagnosticsPendingFlush.empty())
        {
            return;
        }

        pending.assign(std::make_move_iterator(_diagnosticsPendingFlush.begin()), std::make_move_iterator(_diagnosticsPendingFlush.end()));
        _diagnosticsPendingFlush.clear();
        _lastDiagnosticsFlushTick = nowTick;
    }

    const auto requeuePending = [&](size_t startIndex) noexcept
    {
        if (startIndex >= pending.size())
        {
            return;
        }

        std::scoped_lock lock(_diagnosticsMutex);
        const auto startIt = std::next(pending.begin(), static_cast<std::ptrdiff_t>(startIndex));
        _diagnosticsPendingFlush.insert(_diagnosticsPendingFlush.begin(), std::make_move_iterator(startIt), std::make_move_iterator(pending.end()));
    };

    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);
    const std::filesystem::path logPath = GetDiagnosticsLogPathForDate(localNow);
    if (logPath.empty())
    {
        requeuePending(0);
        return;
    }

    const std::filesystem::path logsDir = logPath.parent_path();

    std::error_code ec;
    std::filesystem::create_directories(logsDir, ec);
    if (ec)
    {
        requeuePending(0);
        return;
    }

    wil::unique_handle file(CreateFileW(
        logPath.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        requeuePending(0);
        return;
    }

    LARGE_INTEGER fileSize{};
    bool shouldWriteBom = false;
    if (GetFileSizeEx(file.get(), &fileSize) && fileSize.QuadPart == 0)
    {
        shouldWriteBom = true;
    }

    if (SetFilePointer(file.get(), 0, nullptr, FILE_END) == INVALID_SET_FILE_POINTER && GetLastError() != NO_ERROR)
    {
        requeuePending(0);
        return;
    }

    if (shouldWriteBom)
    {
        constexpr wchar_t bom = 0xFEFF;
        DWORD written         = 0;
        if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr) || written != sizeof(bom))
        {
            requeuePending(0);
            return;
        }
    }

    for (size_t index = 0; index < pending.size(); ++index)
    {
        const TaskDiagnosticEntry& entry = pending[index];
        const wchar_t* categoryText      = entry.category.empty() ? L"general" : entry.category.c_str();
        const unsigned long hrU32        = static_cast<unsigned long>(entry.status);

        const std::wstring timeText = std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}",
                                                  static_cast<unsigned>(entry.localTime.wYear),
                                                  static_cast<unsigned>(entry.localTime.wMonth),
                                                  static_cast<unsigned>(entry.localTime.wDay),
                                                  static_cast<unsigned>(entry.localTime.wHour),
                                                  static_cast<unsigned>(entry.localTime.wMinute),
                                                  static_cast<unsigned>(entry.localTime.wSecond),
                                                  static_cast<unsigned>(entry.localTime.wMilliseconds));
        const std::wstring hrHex    = std::format(L"0x{:08X}", hrU32);

        const std::wstring escapedCategory        = EscapeDiagnosticJsonString(categoryText);
        const std::wstring escapedMessage         = EscapeDiagnosticJsonString(entry.message);
        const std::wstring escapedSource          = EscapeDiagnosticJsonString(entry.sourcePath);
        const std::wstring escapedDest            = EscapeDiagnosticJsonString(entry.destinationPath);
        const std::wstring escapedHrName          = EscapeDiagnosticJsonString(FormatDiagnosticHresultName(entry.status));
        const std::wstring escapedConcurrencyMode = EscapeDiagnosticJsonString(entry.concurrencyMode);
        const std::wstring escapedStorageType     = EscapeDiagnosticJsonString(entry.storageType);
        const std::wstring escapedDestStorageType = EscapeDiagnosticJsonString(entry.destinationStorageType);

        std::wstring escapedHrText;
        if (entry.status != S_OK)
        {
            escapedHrText = EscapeDiagnosticJsonString(FormatDiagnosticStatusText(entry.status));
        }

        std::wstring line;
        line.reserve(256u + timeText.size() + hrHex.size() + escapedHrName.size() + escapedCategory.size() + escapedMessage.size() + escapedSource.size() +
                     escapedDest.size() + escapedHrText.size());

        line.append(L"{\"ts\":\"");
        line.append(timeText);
        line.append(L"\",\"level\":\"");
        line.append(DiagnosticSeverityToString(entry.severity));
        line.append(L"\",\"task\":");
        line.append(std::to_wstring(static_cast<unsigned long long>(entry.taskId)));
        line.append(L",\"op\":\"");
        line.append(OperationToString(entry.operation));
        line.append(L"\",\"category\":\"");
        line.append(escapedCategory);
        line.append(L"\",\"hr\":\"");
        line.append(hrHex);
        line.append(L"\",\"hrName\":\"");
        line.append(escapedHrName);
        line.append(L"\"");
        if (! escapedHrText.empty())
        {
            line.append(L",\"hrText\":\"");
            line.append(escapedHrText);
            line.append(L"\"");
        }
        if (entry.processWorkingSetBytes != 0 || entry.processPrivateBytes != 0)
        {
            line.append(L",\"memWorkingSetBytes\":");
            line.append(std::to_wstring(entry.processWorkingSetBytes));
            line.append(L",\"memPrivateBytes\":");
            line.append(std::to_wstring(entry.processPrivateBytes));
        }
        if (! entry.concurrencyMode.empty())
        {
            line.append(L",\"concurrencyMode\":\"");
            line.append(escapedConcurrencyMode);
            line.append(L"\"");
        }
        if (! entry.storageType.empty())
        {
            line.append(L",\"storageType\":\"");
            line.append(escapedStorageType);
            line.append(L"\"");
        }
        if (! entry.destinationStorageType.empty())
        {
            line.append(L",\"destinationStorageType\":\"");
            line.append(escapedDestStorageType);
            line.append(L"\"");
        }
        if (entry.autoTunedConcurrency != 0)
        {
            line.append(L",\"autoTunedConcurrency\":");
            line.append(std::to_wstring(entry.autoTunedConcurrency));
        }
        if (entry.effectiveConcurrencyBudget != 0)
        {
            line.append(L",\"effectiveConcurrencyBudget\":");
            line.append(std::to_wstring(entry.effectiveConcurrencyBudget));
        }
        line.append(L",\"message\":\"");
        line.append(escapedMessage);
        line.append(L"\"");

        if (! entry.sourcePath.empty())
        {
            line.append(L",\"src\":\"");
            line.append(escapedSource);
            line.append(L"\"");

            const std::wstring_view leaf = GetPathLeaf(entry.sourcePath);
            if (! leaf.empty())
            {
                line.append(L",\"srcLeaf\":\"");
                line.append(EscapeDiagnosticJsonString(leaf));
                line.append(L"\"");
            }
        }
        else
        {
            line.append(L",\"src\":null");
        }

        if (! entry.destinationPath.empty())
        {
            line.append(L",\"dst\":\"");
            line.append(escapedDest);
            line.append(L"\"");

            const std::wstring_view leaf = GetPathLeaf(entry.destinationPath);
            if (! leaf.empty())
            {
                line.append(L",\"dstLeaf\":\"");
                line.append(EscapeDiagnosticJsonString(leaf));
                line.append(L"\"");
            }
        }
        else
        {
            line.append(L",\"dst\":null");
        }

        line.append(L"}\r\n");

        const size_t bytesToWrite = line.size() * sizeof(wchar_t);
        if (bytesToWrite > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
        {
            continue;
        }

        DWORD written = 0;
        if (! WriteFile(file.get(), line.data(), static_cast<DWORD>(bytesToWrite), &written, nullptr) || written != static_cast<DWORD>(bytesToWrite))
        {
            requeuePending(index);
            return;
        }
    }

    bool runCleanup = force;
    {
        std::scoped_lock lock(_diagnosticsMutex);
        if (! runCleanup)
        {
            runCleanup = _lastDiagnosticsCleanupTick == 0 || (nowTick >= _lastDiagnosticsCleanupTick &&
                                                              (nowTick - _lastDiagnosticsCleanupTick) >= diagnosticsSettings.diagnosticsCleanupIntervalMs);
        }
        if (runCleanup)
        {
            _lastDiagnosticsCleanupTick = nowTick;
        }
    }

    if (! runCleanup)
    {
        return;
    }

    CleanupDiagnosticsFilesInDirectory(logsDir, kDiagnosticsLogPrefix, kDiagnosticsLogExtension, diagnosticsSettings.maxDiagnosticsLogFiles);
    CleanupDiagnosticsFilesInDirectory(
        logsDir, kDiagnosticsIssueReportPrefix, kDiagnosticsIssueReportExtension, diagnosticsSettings.maxDiagnosticsIssueReportFiles);
}

void FolderWindow::FileOperationState::RecordTaskDiagnostic(uint64_t taskId,
                                                            FileSystemOperation operation,
                                                            DiagnosticSeverity severity,
                                                            HRESULT status,
                                                            std::wstring_view category,
                                                            std::wstring_view message,
                                                            std::wstring_view sourcePath,
                                                            std::wstring_view destinationPath) noexcept
{
    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    if (severity == DiagnosticSeverity::Info && ! diagnosticsSettings.infoEnabled)
    {
        return;
    }
    if (severity == DiagnosticSeverity::Debug && ! diagnosticsSettings.debugEnabled)
    {
        return;
    }

    TaskDiagnosticEntry entry{};
    GetLocalTime(&entry.localTime);
    entry.taskId    = taskId;
    entry.operation = operation;
    entry.severity  = severity;
    entry.status    = status;
    if (severity == DiagnosticSeverity::Debug || severity == DiagnosticSeverity::Error)
    {
        const ProcessMemorySnapshot snapshot = CaptureProcessMemorySnapshot();
        entry.processWorkingSetBytes         = snapshot.workingSetBytes;
        entry.processPrivateBytes            = snapshot.privateBytes;
    }
    entry.category        = std::wstring(category);
    entry.message         = std::wstring(message);
    entry.sourcePath      = std::wstring(sourcePath);
    entry.destinationPath = std::wstring(destinationPath);
    EnqueueTaskDiagnostic(std::move(entry));
}

void FolderWindow::FileOperationState::EnqueueTaskDiagnostic(TaskDiagnosticEntry entry) noexcept
{
    const DiagnosticsSettings diagnosticsSettings = GetDiagnosticsSettingsFromSettings(_owner._settings);
    if (entry.severity == DiagnosticSeverity::Info && ! diagnosticsSettings.infoEnabled)
    {
        return;
    }
    if (entry.severity == DiagnosticSeverity::Debug && ! diagnosticsSettings.debugEnabled)
    {
        return;
    }

    if (entry.localTime.wYear == 0)
    {
        GetLocalTime(&entry.localTime);
    }

    if ((entry.severity == DiagnosticSeverity::Debug || entry.severity == DiagnosticSeverity::Error) && entry.processWorkingSetBytes == 0 &&
        entry.processPrivateBytes == 0)
    {
        const ProcessMemorySnapshot snapshot = CaptureProcessMemorySnapshot();
        entry.processWorkingSetBytes         = snapshot.workingSetBytes;
        entry.processPrivateBytes            = snapshot.privateBytes;
    }

    const ULONGLONG nowTick = GetTickCount64();
    bool shouldFlush        = false;
    {
        std::scoped_lock lock(_diagnosticsMutex);

        _diagnosticsInMemory.push_back(entry);
        while (_diagnosticsInMemory.size() > diagnosticsSettings.maxDiagnosticsInMemory)
        {
            _diagnosticsInMemory.pop_front();
        }

        _diagnosticsPendingFlush.push_back(entry);
        const size_t maxPendingFlush = std::max(diagnosticsSettings.maxDiagnosticsInMemory, diagnosticsSettings.maxDiagnosticsPerFlush);
        if (_diagnosticsPendingFlush.size() > maxPendingFlush)
        {
            const size_t overflow = _diagnosticsPendingFlush.size() - maxPendingFlush;
            _diagnosticsPendingFlush.erase(_diagnosticsPendingFlush.begin(), _diagnosticsPendingFlush.begin() + static_cast<std::ptrdiff_t>(overflow));
        }

        if (entry.severity == DiagnosticSeverity::Warning || entry.severity == DiagnosticSeverity::Error)
        {
            auto& counts = _taskDiagnosticCounts[entry.taskId];
            if (entry.severity == DiagnosticSeverity::Warning)
            {
                ++counts.first;
            }
            else
            {
                ++counts.second;
            }

            if (! entry.message.empty())
            {
                _taskLastDiagnosticMessage[entry.taskId] = entry.message;
            }

            auto& issues = _taskIssueDiagnostics[entry.taskId];
            issues.push_back(entry);
            while (issues.size() > kMaxTaskIssueDiagnostics)
            {
                issues.pop_front();
            }
        }

        const bool flushIntervalReached =
            _lastDiagnosticsFlushTick == 0 ||
            (nowTick >= _lastDiagnosticsFlushTick && (nowTick - _lastDiagnosticsFlushTick) >= diagnosticsSettings.diagnosticsFlushIntervalMs);
        shouldFlush = flushIntervalReached || _diagnosticsPendingFlush.size() >= diagnosticsSettings.maxDiagnosticsPerFlush;
    }

    if (shouldFlush)
    {
        FlushDiagnostics(false);
    }
}

FolderWindow::FileOperationState::CompletedTaskSummary FolderWindow::FileOperationState::RecordCompletedTask(Task& task) noexcept
{
    CompletedTaskSummary summary{};
    SYSTEMTIME localNow{};
    GetLocalTime(&localNow);
    summary.taskId                                   = task._taskId;
    summary.operation                                = task._operation;
    summary.sourcePane                               = task._sourcePane;
    summary.destinationPane                          = task._destinationPane;
    summary.destinationPluginId                      = task._destinationPluginId;
    summary.destinationPluginShortId                 = task._destinationPluginShortId;
    summary.destinationInstanceContext               = task._destinationInstanceContext;
    summary.destinationFolder                        = task.GetDestinationFolder();
    summary.diagnosticsLogPath                       = GetDiagnosticsLogPathForDate(localNow);
    summary.resultHr                                 = task.GetResult();
    summary.completedTick                            = GetTickCount64();
    const PublishedProgressSnapshot progressSnapshot = LoadPublishedProgressSnapshot(task);
    summary.totalItems                               = progressSnapshot.totalItems;
    summary.completedItems                           = progressSnapshot.completedItems;
    summary.totalBytes                               = progressSnapshot.totalBytes;
    summary.completedBytes                           = progressSnapshot.completedBytes;
    summary.preCalcSkipped                           = task._preCalcSkipped.load(std::memory_order_acquire);
    summary.completedFiles                           = progressSnapshot.completedFiles;
    summary.completedFolders                         = progressSnapshot.completedFolders;
    summary.autoConcurrencyUsed                      = task._autoConcurrencyUsed.load(std::memory_order_acquire);
    summary.autoConcurrencyStorageKind               = task._autoConcurrencyStorageKind.load(std::memory_order_acquire);
    summary.autoConcurrencyDestinationStorageKind    = task._autoConcurrencyDestinationStorageKind.load(std::memory_order_acquire);
    summary.autoTunedConcurrency                     = task._autoTunedConcurrency.load(std::memory_order_acquire);
    summary.effectiveConcurrencyBudget               = task._effectiveConcurrencyBudget.load(std::memory_order_acquire);

    {
        std::scoped_lock lock(task._progressPathMutex);
        summary.lastProgressCallbackTick = task._lastProgressCallbackTick;
        summary.sourcePath               = task._progressSourcePath;
        summary.destinationPath          = task._progressDestinationPath;
    }

    if ((summary.operation == FILESYSTEM_COPY || summary.operation == FILESYSTEM_MOVE) && summary.destinationPane.has_value() &&
        ! summary.destinationFolder.empty() && task._sourcePaths.size() == 1u)
    {
        const bool destinationPathOnlyNamesFolder =
            summary.destinationPath.empty() || NavigationLocation::EqualsNoCase(summary.destinationPath, summary.destinationFolder.native());
        if (destinationPathOnlyNamesFolder)
        {
            if (task._resolvedItems.size() == task._sourcePaths.size() && ! task._resolvedItems.front().destinationPath.empty())
            {
                summary.destinationPath = task._resolvedItems.front().destinationPath.native();
            }
            else
            {
                summary.destinationPath = JoinFolderAndLeaf(summary.destinationFolder.native(), GetPathLeaf(task._sourcePaths.front().native()));
            }
        }
    }

    {
        std::scoped_lock lock(_diagnosticsMutex);
        const auto countsIt = _taskDiagnosticCounts.find(summary.taskId);
        if (countsIt != _taskDiagnosticCounts.end())
        {
            summary.warningCount = countsIt->second.first;
            summary.errorCount   = countsIt->second.second;
            _taskDiagnosticCounts.erase(countsIt);
        }

        const auto messageIt = _taskLastDiagnosticMessage.find(summary.taskId);
        if (messageIt != _taskLastDiagnosticMessage.end())
        {
            summary.lastDiagnosticMessage = messageIt->second;
            _taskLastDiagnosticMessage.erase(messageIt);
        }

        const auto issuesIt = _taskIssueDiagnostics.find(summary.taskId);
        if (issuesIt != _taskIssueDiagnostics.end())
        {
            summary.issueDiagnostics.assign(issuesIt->second.begin(), issuesIt->second.end());
            _taskIssueDiagnostics.erase(issuesIt);
        }
    }

    if (FAILED(summary.resultHr) && summary.warningCount == 0 && summary.errorCount == 0)
    {
        const HRESULT partialHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        if (summary.resultHr == partialHr)
        {
            summary.warningCount = 1;
            if (summary.lastDiagnosticMessage.empty())
            {
                summary.lastDiagnosticMessage = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_SKIPPED_ITEMS);
            }
        }
        else if (! IsCancellationStatus(summary.resultHr))
        {
            summary.errorCount = 1;
            if (summary.lastDiagnosticMessage.empty())
            {
                summary.lastDiagnosticMessage =
                    FormatStringResource(nullptr, IDS_FMT_FILEOPS_DIAG_FAILED_NO_DETAILS, static_cast<unsigned long>(summary.resultHr));
            }
        }
    }

    if ((summary.warningCount > 0 || summary.errorCount > 0) && summary.issueDiagnostics.empty())
    {
        TaskDiagnosticEntry synthetic{};
        synthetic.localTime       = localNow;
        synthetic.taskId          = summary.taskId;
        synthetic.operation       = summary.operation;
        synthetic.severity        = summary.errorCount > 0 ? DiagnosticSeverity::Error : DiagnosticSeverity::Warning;
        synthetic.status          = summary.resultHr;
        synthetic.category        = L"task.summary";
        synthetic.message         = summary.lastDiagnosticMessage.empty() ? L"Task completed with diagnostics." : summary.lastDiagnosticMessage;
        synthetic.sourcePath      = summary.sourcePath;
        synthetic.destinationPath = summary.destinationPath;
        summary.issueDiagnostics.push_back(std::move(synthetic));
    }

    std::wstring completedStatus = L"success";
    if (summary.resultHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
    {
        completedStatus = L"partial";
    }
    else if (IsCancellationStatus(summary.resultHr))
    {
        completedStatus = L"canceled";
    }
    else if (FAILED(summary.resultHr))
    {
        completedStatus = L"failed";
    }

    const std::wstring completedMessage = std::format(L"Task completed: status={}, op={}, hr=0x{:08X}, items={:L}/{:L}, bytes={:L}/{:L}.",
                                                      completedStatus,
                                                      OperationToString(summary.operation),
                                                      static_cast<unsigned long>(summary.resultHr),
                                                      summary.completedItems,
                                                      summary.totalItems,
                                                      summary.completedBytes,
                                                      summary.totalBytes);
    RecordTaskDiagnostic(summary.taskId,
                         summary.operation,
                         DiagnosticSeverity::Info,
                         summary.resultHr,
                         L"task.completed",
                         completedMessage,
                         summary.sourcePath,
                         summary.destinationPath);

    const bool autoConcurrencyUsed                = task._autoConcurrencyUsed.load(std::memory_order_acquire);
    const uint32_t autoStorageKind                = task._autoConcurrencyStorageKind.load(std::memory_order_acquire);
    const uint32_t destinationStorageKind         = task._autoConcurrencyDestinationStorageKind.load(std::memory_order_acquire);
    const unsigned int autoTunedConcurrency       = task._autoTunedConcurrency.load(std::memory_order_acquire);
    const unsigned int effectiveConcurrencyBudget = task._effectiveConcurrencyBudget.load(std::memory_order_acquire);

    if (autoConcurrencyUsed && autoTunedConcurrency > 0u)
    {
        TaskDiagnosticEntry autoEntry{};
        autoEntry.taskId               = summary.taskId;
        autoEntry.operation            = summary.operation;
        autoEntry.severity             = DiagnosticSeverity::Info;
        autoEntry.status               = summary.resultHr;
        autoEntry.category             = L"task.autoConcurrency";
        autoEntry.message              = std::format(L"Resolved {} concurrency: storageType={}, autoTunedConcurrency={:L}, effectiveConcurrencyBudget={:L}.",
                                                     ConcurrencyModeToString(FileSystemConcurrencyMode::Auto),
                                                     StorageKindToString(autoStorageKind),
                                                     static_cast<unsigned long>(autoTunedConcurrency),
                                                     static_cast<unsigned long>(effectiveConcurrencyBudget));
        autoEntry.sourcePath           = summary.sourcePath;
        autoEntry.destinationPath      = summary.destinationPath;
        autoEntry.concurrencyMode      = ConcurrencyModeToString(FileSystemConcurrencyMode::Auto);
        autoEntry.storageType          = StorageKindToString(autoStorageKind);
        autoEntry.autoTunedConcurrency = autoTunedConcurrency;
        autoEntry.effectiveConcurrencyBudget = effectiveConcurrencyBudget;
        if (destinationStorageKind != FILESYSTEM_STORAGE_UNKNOWN)
        {
            autoEntry.destinationStorageType = StorageKindToString(destinationStorageKind);
        }
        EnqueueTaskDiagnostic(std::move(autoEntry));
    }

    {
        std::scoped_lock lock(_mutex);
        _completedTasks.push_front(summary);
        while (_completedTasks.size() > kMaxCompletedTaskSummaries)
        {
            _completedTasks.pop_back();
        }
    }

    FlushDiagnostics(true);
    return summary;
}
