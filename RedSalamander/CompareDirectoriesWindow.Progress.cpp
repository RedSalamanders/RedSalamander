#include "Framework.h"

#include "CompareDirectoriesWindow.Internal.h"
#include "D2DHdcPaint.h"

namespace CompareDirectoriesWindowInternal
{
using SteadyClock                                       = std::chrono::steady_clock;
constexpr ULONGLONG kCompareTaskCardUpdateMinIntervalMs = 50;

[[nodiscard]] uint64_t ElapsedUsSince(SteadyClock::time_point startedAt) noexcept
{
    if (startedAt == SteadyClock::time_point{})
    {
        return 0u;
    }

    const auto elapsed = std::chrono::steady_clock::now() - startedAt;
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count());
}

struct ScanProgressPayload
{
    uint64_t runId                      = 0;
    uint32_t activeScans                = 0;
    uint64_t folderCount                = 0;
    uint64_t entryCount                 = 0;
    uint64_t contentCandidateFileCount  = 0;
    uint64_t contentCandidateTotalBytes = 0;
    std::filesystem::path relativeFolder;
    std::wstring entryName;
    SteadyClock::time_point enqueuedAt{};
};

struct ContentProgressPayload
{
    uint64_t runId                    = 0;
    uint32_t workerIndex              = std::numeric_limits<uint32_t>::max();
    uint64_t pendingContentCompares   = 0;
    uint64_t fileTotalBytes           = 0;
    uint64_t fileCompletedBytes       = 0;
    uint64_t overallTotalBytes        = 0;
    uint64_t overallCompletedBytes    = 0;
    uint64_t totalContentCompares     = 0;
    uint64_t completedContentCompares = 0;
    std::filesystem::path relativeFolder;
    std::wstring entryName;
    SteadyClock::time_point enqueuedAt{};
};

[[nodiscard]] bool TryStartCompareTimer(HWND hwnd, UINT_PTR timerId, UINT delayMs, bool& activeFlag, const wchar_t* failureMessage) noexcept
{
    activeFlag = hwnd && SetTimer(hwnd, timerId, delayMs, nullptr) != 0;
    if (! activeFlag && failureMessage)
    {
        Debug::ErrorWithLastError(failureMessage);
    }
    return activeFlag;
}

[[nodiscard]] std::wstring FormatDurationHmsNoexcept(uint64_t seconds) noexcept
{
    const uint64_t hours64   = seconds / 3600u;
    const uint64_t minutes64 = (seconds % 3600u) / 60u;
    const uint64_t seconds64 = seconds % 60u;

    const unsigned long long hours = static_cast<unsigned long long>(hours64);
    const unsigned int minutes     = static_cast<unsigned int>(minutes64);
    const unsigned int secs        = static_cast<unsigned int>(seconds64);

    try
    {
        if (hours > 0ull)
        {
            return std::format(L"{0}:{1:02d}:{2:02d}", hours, minutes, secs);
        }

        return std::format(L"{0:02d}:{1:02d}", minutes, secs);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::format_error&)
    {
        return {};
    }
}

void CompareDirectoriesWindow::SetSessionCallbacksForRun(uint64_t runId) noexcept
{
    if (! _session || ! _hWnd)
    {
        return;
    }

    const HWND hwnd = _hWnd.get();
    _session->SetScanProgressCallback([hwnd, runId](const std::filesystem::path& relativeFolder,
                                                    std::wstring_view currentEntryName,
                                                    uint64_t scannedFolders,
                                                    uint64_t scannedEntries,
                                                    uint32_t activeScans,
                                                    uint64_t contentCandidateFileCount,
                                                    uint64_t contentCandidateTotalBytes) noexcept
    {
        if (! hwnd)
        {
            return;
        }

        auto payload                        = std::make_unique<ScanProgressPayload>();
        payload->runId                      = runId;
        payload->activeScans                = activeScans;
        payload->folderCount                = scannedFolders;
        payload->entryCount                 = scannedEntries;
        payload->contentCandidateFileCount  = contentCandidateFileCount;
        payload->contentCandidateTotalBytes = contentCandidateTotalBytes;
        payload->relativeFolder             = relativeFolder;
        payload->entryName                  = std::wstring(currentEntryName);
        payload->enqueuedAt                 = SteadyClock::now();
        const WPARAM operationKey           = static_cast<WPARAM>(payload->runId);
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kCompareDirectoriesScanProgress, operationKey, std::move(payload)));
    });

    _session->SetContentProgressCallback([hwnd, runId](uint32_t workerIndex,
                                                       const std::filesystem::path& relativeFolder,
                                                       std::wstring_view entryName,
                                                       uint64_t fileTotalBytes,
                                                       uint64_t fileCompletedBytes,
                                                       uint64_t overallTotalBytes,
                                                       uint64_t overallCompletedBytes,
                                                       uint64_t pendingContentCompares,
                                                       uint64_t totalContentCompares,
                                                       uint64_t completedContentCompares) noexcept
    {
        if (! hwnd)
        {
            return;
        }

        auto payload                      = std::make_unique<ContentProgressPayload>();
        payload->runId                    = runId;
        payload->workerIndex              = workerIndex;
        payload->pendingContentCompares   = pendingContentCompares;
        payload->fileTotalBytes           = fileTotalBytes;
        payload->fileCompletedBytes       = fileCompletedBytes;
        payload->overallTotalBytes        = overallTotalBytes;
        payload->overallCompletedBytes    = overallCompletedBytes;
        payload->totalContentCompares     = totalContentCompares;
        payload->completedContentCompares = completedContentCompares;
        payload->relativeFolder           = relativeFolder;
        payload->entryName                = std::wstring(entryName);
        payload->enqueuedAt               = SteadyClock::now();
        const WPARAM operationKey         = static_cast<WPARAM>(payload->runId);
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kCompareDirectoriesContentProgress, operationKey, std::move(payload)));
    });

    _session->SetDecisionUpdatedCallback([hwnd, runId]() noexcept
    {
        if (! hwnd || IsWindow(hwnd) == 0)
        {
            return;
        }

        PostMessageW(hwnd, WndMsg::kCompareDirectoriesDecisionUpdated, static_cast<WPARAM>(runId), 0);
    });
}

void CompareDirectoriesWindow::ScheduleDecisionRefresh() noexcept
{
    _decisionRefreshPending = true;

    if (_decisionRefreshTimerActive || _decisionRefreshFallbackPending || ! _hWnd)
    {
        return;
    }

    if (TryStartCompareTimer(_hWnd.get(),
                             kCompareDecisionRefreshTimerId,
                             kCompareDecisionRefreshTimerIntervalMs,
                             _decisionRefreshTimerActive,
                             L"CompareDirectories: failed to start decision refresh timer."))
    {
        return;
    }

    _decisionRefreshFallbackPending = true;
    if (PostMessageW(_hWnd.get(), WndMsg::kCompareDirectoriesDecisionRefreshNow, static_cast<WPARAM>(_compareRunId), 0) == 0)
    {
        _decisionRefreshFallbackPending = false;
        Debug::ErrorWithLastError(L"CompareDirectories: failed to post decision refresh fallback.");
        OnDecisionRefreshTimer();
    }
}

void CompareDirectoriesWindow::OnDecisionRefreshTimer() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    KillTimer(_hWnd.get(), kCompareDecisionRefreshTimerId);
    _decisionRefreshTimerActive = false;

    if (! _decisionRefreshPending)
    {
        return;
    }
    _decisionRefreshPending = false;

    if (! _compareStarted || ! _compareActive || ! _session)
    {
        return;
    }

    constexpr size_t kMaxFlushContentFoldersPerTick = 8;
    constexpr size_t kMaxFlushSubdirKeysPerTick     = 16;

    const bool moreContentPending = _session->FlushPendingContentCompareUpdatesBudgeted(kMaxFlushContentFoldersPerTick);
    const bool moreSubdirPending  = _session->FlushPendingSubdirUpdatesBudgeted(kMaxFlushSubdirKeysPerTick);
    const bool morePending        = moreContentPending || moreSubdirPending;

    const Common::Settings::CompareDirectoriesSettings s = GetEffectiveCompareSettings();
    if (s.showIdenticalItems)
    {
        _folderWindow.RefreshPaneDetailsText(FolderWindow::Pane::Left);
        _folderWindow.RefreshPaneDetailsText(FolderWindow::Pane::Right);

        if (const auto leftPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Left))
        {
            ApplySelectionForFolder(ComparePane::Left, leftPath.value());
            UpdateEmptyStateForFolder(ComparePane::Left, leftPath.value());
        }
        if (const auto rightPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Right))
        {
            ApplySelectionForFolder(ComparePane::Right, rightPath.value());
            UpdateEmptyStateForFolder(ComparePane::Right, rightPath.value());
        }

        if (morePending)
        {
            ScheduleDecisionRefresh();
        }
        return;
    }

    struct RefreshDecisionSnapshot
    {
        std::optional<std::filesystem::path> relativeFolder;
        std::shared_ptr<const CompareDirectoriesFolderDecision> decision;
    };

    const auto snapshotForPane = [&](ComparePane pane, FolderWindow::Pane fwPane) noexcept -> RefreshDecisionSnapshot
    {
        RefreshDecisionSnapshot snap{};
        if (! _session)
        {
            return snap;
        }

        const auto currentPathOpt = _folderWindow.GetCurrentPluginPath(fwPane);
        if (! currentPathOpt.has_value())
        {
            return snap;
        }

        const std::filesystem::path absolute = currentPathOpt.value();
        const auto relativeOpt               = _session->TryMakeRelative(pane, absolute);
        if (! relativeOpt.has_value())
        {
            return snap;
        }

        snap.relativeFolder = relativeOpt.value();
        snap.decision       = _session->TryGetCachedDecision(snap.relativeFolder.value());
        return snap;
    };

    const auto changedSinceLast = [](const RefreshDecisionSnapshot& snap,
                                     const std::optional<std::filesystem::path>& lastFolder,
                                     const std::shared_ptr<const CompareDirectoriesFolderDecision>& lastDecision) noexcept -> bool
    { return snap.relativeFolder != lastFolder || snap.decision != lastDecision; };

    const RefreshDecisionSnapshot leftSnapshot  = snapshotForPane(ComparePane::Left, FolderWindow::Pane::Left);
    const RefreshDecisionSnapshot rightSnapshot = snapshotForPane(ComparePane::Right, FolderWindow::Pane::Right);

    const bool shouldRefresh = changedSinceLast(leftSnapshot, _lastRefreshedLeftRelativeFolder, _lastRefreshedLeftDecision) ||
                               changedSinceLast(rightSnapshot, _lastRefreshedRightRelativeFolder, _lastRefreshedRightDecision);

    if (shouldRefresh)
    {
        RefreshBothPanes();
        _lastRefreshedLeftRelativeFolder  = leftSnapshot.relativeFolder;
        _lastRefreshedLeftDecision        = leftSnapshot.decision;
        _lastRefreshedRightRelativeFolder = rightSnapshot.relativeFolder;
        _lastRefreshedRightDecision       = rightSnapshot.decision;
    }
    if (morePending)
    {
        ScheduleDecisionRefresh();
    }
}

LRESULT CompareDirectoriesWindow::OnScanProgress(WPARAM operationKey, LPARAM lp) noexcept
{
    auto drained = TakeAndCoalesceContiguousPostedPayloads<ScanProgressPayload>(
        _hWnd.get(), WndMsg::kCompareDirectoriesScanProgress, operationKey, lp, [](const ScanProgressPayload&, uint64_t) noexcept {
        return true;
    }, [](std::unique_ptr<ScanProgressPayload>& current, std::unique_ptr<ScanProgressPayload> newer) noexcept { current = std::move(newer); });
    auto payload = std::move(drained.payload);
    if (! payload)
    {
        return 0;
    }
    const uint64_t drainedPayloadCount = drained.drainedPayloadCount;

    Debug::Perf::EmitCounter(L"compare.ui.scan_progress_message_count");
    if (drainedPayloadCount != 0u)
    {
        Debug::Perf::EmitCounter(L"compare.ui.scan_progress_messages_coalesced_count");
        Debug::Perf::EmitCounter(L"compare.ui.scan_progress_messages_drained", drainedPayloadCount);
    }
    if (payload->enqueuedAt != SteadyClock::time_point{})
    {
        Debug::Perf::Emit(L"compare.ui.scan_progress_to_visible_latency_ms",
                          L"",
                          ElapsedUsSince(payload->enqueuedAt),
                          payload->activeScans,
                          static_cast<uint64_t>(drainedPayloadCount),
                          S_OK);
    }

    if (! _compareActive || payload->runId != _compareRunId || static_cast<WPARAM>(payload->runId) != operationKey)
    {
        return 0;
    }

    _progress.compareRunSawScanProgress = true;

    _progress.banner.scanActiveScans                = payload->activeScans;
    _progress.banner.scanFolderCount                = payload->folderCount;
    _progress.banner.scanEntryCount                 = payload->entryCount;
    _progress.banner.scanContentCandidateFileCount  = payload->contentCandidateFileCount;
    _progress.banner.scanContentCandidateTotalBytes = payload->contentCandidateTotalBytes;
    _progress.banner.scanRelativeFolder             = std::move(payload->relativeFolder);
    _progress.banner.scanEntryName                  = std::move(payload->entryName);

    if (_progress.banner.scanActiveScans == 0u)
    {
        _progress.banner.scanRelativeFolder.clear();
        _progress.banner.scanEntryName.clear();
    }

    UpdateRescanButtonText();
    UpdateProgressControls();
    if (_compareRunPending)
    {
        UpdateCompareTaskCard(false);
    }
    MaybeCompleteCompareRun();
    return 0;
}

LRESULT CompareDirectoriesWindow::OnContentProgress(WPARAM operationKey, LPARAM lp) noexcept
{
    auto drained = TakeAndCoalesceContiguousPostedPayloads<ContentProgressPayload>(
        _hWnd.get(), WndMsg::kCompareDirectoriesContentProgress, operationKey, lp, [](const ContentProgressPayload&, uint64_t) noexcept {
        return true;
    }, [](std::unique_ptr<ContentProgressPayload>& current, std::unique_ptr<ContentProgressPayload> newer) noexcept { current = std::move(newer); });
    auto payload = std::move(drained.payload);
    if (! payload)
    {
        return 0;
    }
    const uint64_t drainedPayloadCount = drained.drainedPayloadCount;

    Debug::Perf::EmitCounter(L"compare.ui.content_progress_message_count");
    if (drainedPayloadCount != 0u)
    {
        Debug::Perf::EmitCounter(L"compare.ui.content_progress_messages_coalesced_count");
        Debug::Perf::EmitCounter(L"compare.ui.content_progress_messages_drained", drainedPayloadCount);
    }
    if (payload->enqueuedAt != SteadyClock::time_point{})
    {
        Debug::Perf::Emit(L"compare.ui.content_progress_to_visible_latency_ms",
                          L"",
                          ElapsedUsSince(payload->enqueuedAt),
                          payload->pendingContentCompares,
                          static_cast<uint64_t>(drainedPayloadCount),
                          S_OK);
    }

    if (! _compareActive || payload->runId != _compareRunId || static_cast<WPARAM>(payload->runId) != operationKey)
    {
        return 0;
    }

    const ULONGLONG nowTick = GetTickCount64();

    _progress.banner.contentPendingCompares       = payload->pendingContentCompares;
    _progress.banner.contentTotalCompares         = payload->totalContentCompares;
    _progress.banner.contentCompletedCompares     = payload->completedContentCompares;
    _progress.banner.contentOverallTotalBytes     = payload->overallTotalBytes;
    _progress.banner.contentOverallCompletedBytes = payload->overallCompletedBytes;
    _progress.banner.contentFileTotalBytes        = payload->fileTotalBytes;
    _progress.banner.contentFileCompletedBytes    = payload->fileCompletedBytes;

    if (_progress.banner.contentPendingCompares > 0u)
    {
        std::filesystem::path fileRel = payload->relativeFolder;
        if (! payload->entryName.empty())
        {
            fileRel /= std::filesystem::path(payload->entryName);
        }

        if (! fileRel.empty())
        {
            const uint32_t slotIndex = payload->workerIndex;
            if (slotIndex < _progress.banner.contentInFlight.size())
            {
                auto& slot          = _progress.banner.contentInFlight[slotIndex];
                slot.relativePath   = std::move(fileRel);
                slot.totalBytes     = payload->fileTotalBytes;
                slot.completedBytes = payload->fileCompletedBytes;
                slot.lastUpdateTick = nowTick;
            }
        }
    }
    else
    {
        for (auto& slot : _progress.banner.contentInFlight)
        {
            slot = {};
        }
    }

    _progress.banner.contentRelativeFolder = std::move(payload->relativeFolder);
    _progress.banner.contentEntryName      = std::move(payload->entryName);

    if (_progress.banner.contentPendingCompares > 0u)
    {
        const uint64_t completed = _progress.banner.contentOverallCompletedBytes;
        const uint64_t total     = _progress.banner.contentOverallTotalBytes;

        if (_progress.contentEtaLastTickMs != 0u && nowTick > _progress.contentEtaLastTickMs && completed >= _progress.contentEtaLastCompletedBytes)
        {
            const uint64_t deltaBytes = completed - _progress.contentEtaLastCompletedBytes;
            const double deltaSeconds = static_cast<double>(nowTick - _progress.contentEtaLastTickMs) / 1000.0;
            if (deltaBytes > 0u && deltaSeconds >= 0.2)
            {
                const double rate = static_cast<double>(deltaBytes) / deltaSeconds;
                if (_progress.contentEtaSmoothedBytesPerSec <= 1.0)
                {
                    _progress.contentEtaSmoothedBytesPerSec = rate;
                }
                else
                {
                    constexpr double kAlpha                 = 0.15;
                    _progress.contentEtaSmoothedBytesPerSec = (_progress.contentEtaSmoothedBytesPerSec * (1.0 - kAlpha)) + (rate * kAlpha);
                }
            }
        }

        _progress.contentEtaLastTickMs         = nowTick;
        _progress.contentEtaLastCompletedBytes = completed;

        _progress.contentEtaSeconds.reset();
        constexpr double kMinUsefulContentEtaBytesPerSec = 1024.0;
        constexpr double kMaxDisplayedContentEtaSeconds  = 7.0 * 24.0 * 60.0 * 60.0;
        if (total > 0u && completed <= total && std::isfinite(_progress.contentEtaSmoothedBytesPerSec) &&
            _progress.contentEtaSmoothedBytesPerSec >= kMinUsefulContentEtaBytesPerSec)
        {
            const uint64_t remaining = total - completed;
            const double secondsD    = static_cast<double>(remaining) / _progress.contentEtaSmoothedBytesPerSec;
            if (std::isfinite(secondsD) && secondsD >= 0.0 && secondsD <= kMaxDisplayedContentEtaSeconds)
            {
                _progress.contentEtaSeconds = static_cast<uint64_t>(std::ceil(secondsD));
            }
        }
    }
    else
    {
        _progress.contentEtaLastTickMs          = 0;
        _progress.contentEtaLastCompletedBytes  = 0;
        _progress.contentEtaSmoothedBytesPerSec = 0.0;
        _progress.contentEtaSeconds.reset();
    }

    if (_progress.banner.contentPendingCompares == 0u)
    {
        _progress.banner.contentFileTotalBytes     = 0;
        _progress.banner.contentFileCompletedBytes = 0;
        _progress.banner.contentRelativeFolder.clear();
        _progress.banner.contentEntryName.clear();
    }

    UpdateRescanButtonText();
    UpdateProgressControls();
    if (_compareRunPending)
    {
        UpdateCompareTaskCard(false);
    }
    MaybeCompleteCompareRun();
    return 0;
}

void CompareDirectoriesWindow::UpdateProgressControls() noexcept
{
    Debug::Perf::EmitCounter(L"compare.ui.progress_controls_update_count");
    const SteadyClock::time_point startedAt = SteadyClock::now();

    if (! _progress.scanProgressText && ! _progress.scanProgressTextHostHwnd && ! _progress.scanProgressBar)
    {
        UpdateCompareWatermark();
        if (_progress.watermarkState == CompareWatermarkState::InProgress)
        {
            EnsureProgressPulseTimer();
        }
        else
        {
            StopProgressPulseTimerIfIdle();
        }
        Debug::Perf::Emit(L"compare.ui.progress_controls_update_us", L"", ElapsedUsSince(startedAt), 0u, 0u, S_OK);
        return;
    }

    const bool show = (_compareActive && _compareRunPending) || _progress.banner.scanActiveScans > 0u || _progress.banner.contentPendingCompares > 0u;
    const bool useDxProgressText = _chrome.usesBannerText && _progress.scanProgressTextHostHwnd && _progress.scanProgressTextLabel;
    const bool wasVisible        = (useDxProgressText && IsWindowVisible(_progress.scanProgressTextHostHwnd.get()) != 0) ||
                                   (_progress.scanProgressText && IsWindowVisible(_progress.scanProgressText.get()) != 0) ||
                                   (_progress.scanProgressBar && IsWindowVisible(_progress.scanProgressBar.get()) != 0);

    if (! show)
    {
        if (_progress.scanProgressBar)
        {
            if (_progress.controlsVisible || IsWindowVisible(_progress.scanProgressBar.get()) != 0)
            {
                ShowWindow(_progress.scanProgressBar.get(), SW_HIDE);
            }
        }
        if (_progress.scanProgressText)
        {
            if (! useDxProgressText && ! _progress.lastMessage.empty())
            {
                SetWindowTextW(_progress.scanProgressText.get(), L"");
                Debug::Perf::EmitCounter(L"compare.ui.progress_controls_text_applied_count");
            }
            if (_progress.controlsVisible || IsWindowVisible(_progress.scanProgressText.get()) != 0)
            {
                ShowWindow(_progress.scanProgressText.get(), SW_HIDE);
            }
        }
        if (useDxProgressText)
        {
            if (! _progress.lastMessage.empty())
            {
                _progress.scanProgressTextLabel->SetText(std::wstring{});
                _progress.scanProgressTextHost.Invalidate();
                Debug::Perf::EmitCounter(L"compare.ui.progress_controls_text_applied_count");
            }
            if (_progress.controlsVisible || IsWindowVisible(_progress.scanProgressTextHostHwnd.get()) != 0)
            {
                ShowWindow(_progress.scanProgressTextHostHwnd.get(), SW_HIDE);
            }
        }
        _progress.controlsVisible = false;
        _progress.lastMessage.clear();
        if (wasVisible)
        {
            Debug::Perf::EmitCounter(L"compare.ui.progress_controls_hide_count");
            Layout();
        }
        else
        {
            Debug::Perf::EmitCounter(L"compare.ui.progress_controls_skipped_count");
        }
        UpdateCompareWatermark();
        StopProgressPulseTimerIfIdle();
        Debug::Perf::Emit(L"compare.ui.progress_controls_update_us", L"", ElapsedUsSince(startedAt), 0u, 0u, S_OK);
        return;
    }

    std::wstring scanText;
    if (_progress.banner.scanActiveScans > 0u || (_compareActive && _compareRunPending && _progress.banner.contentPendingCompares == 0u))
    {
        std::filesystem::path displayPath = _progress.banner.scanRelativeFolder;
        if (! _progress.banner.scanEntryName.empty())
        {
            displayPath /= std::filesystem::path(_progress.banner.scanEntryName);
        }

        std::wstring pathText;
        if (displayPath.empty())
        {
            pathText = L".";
        }
        else
        {
            pathText = displayPath.wstring();
        }

        scanText = FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, pathText, _progress.banner.scanFolderCount, _progress.banner.scanEntryCount);
        if (_progress.scanStartTickMs != 0)
        {
            const uint64_t elapsedSec   = (GetTickCount64() - _progress.scanStartTickMs) / 1000u;
            const std::wstring duration = FormatDurationHmsNoexcept(elapsedSec);
            if (! duration.empty())
            {
                const std::wstring elapsedText = FormatStringResource(nullptr, IDS_FMT_COMPARE_ELAPSED, duration);
                if (! elapsedText.empty())
                {
                    scanText.append(L" \u2022 ");
                    scanText.append(elapsedText);
                }
            }
        }
    }

    std::wstring contentText;
    if (_progress.banner.contentPendingCompares > 0u && ! _progress.banner.contentEntryName.empty())
    {
        std::filesystem::path displayPath = _progress.banner.contentRelativeFolder;
        if (! _progress.banner.contentEntryName.empty())
        {
            displayPath /= std::filesystem::path(_progress.banner.contentEntryName);
        }

        std::wstring pathText;
        if (displayPath.empty())
        {
            pathText = L".";
        }
        else
        {
            pathText = displayPath.wstring();
        }

        const std::wstring completedText = FormatBytesCompact(_progress.banner.contentFileCompletedBytes);
        if (_progress.banner.contentFileTotalBytes > 0u)
        {
            const std::wstring totalText = FormatBytesCompact(_progress.banner.contentFileTotalBytes);
            contentText                  = FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_STATUS, pathText, completedText, totalText);
        }
        else
        {
            contentText = FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_STATUS_UNKNOWN, pathText, completedText);
        }

        if (_progress.contentEtaSeconds.has_value())
        {
            const std::wstring duration = FormatDurationHmsNoexcept(_progress.contentEtaSeconds.value());
            if (! duration.empty())
            {
                const std::wstring etaText = FormatStringResource(nullptr, IDS_FMT_COMPARE_ETA, duration);
                if (! etaText.empty())
                {
                    contentText.append(L" \u2022 ");
                    contentText.append(etaText);
                }
            }
        }
    }

    std::wstring message;
    if (! scanText.empty())
    {
        message = std::move(scanText);
    }
    if (! contentText.empty())
    {
        if (! message.empty())
        {
            message.append(L" \u2022 ");
        }
        message.append(contentText);
    }

    bool textApplied = false;
    if (useDxProgressText)
    {
        if (message != _progress.lastMessage)
        {
            _progress.scanProgressTextLabel->SetText(message);
            _progress.lastMessage = message;
            textApplied           = true;
            _progress.scanProgressTextHost.Invalidate();
            Debug::Perf::EmitCounter(L"compare.ui.progress_controls_text_applied_count");
        }
    }
    else if (_progress.scanProgressText)
    {
        if (message != _progress.lastMessage)
        {
            SetWindowTextW(_progress.scanProgressText.get(), message.c_str());
            _progress.lastMessage = message;
            textApplied           = true;
            Debug::Perf::EmitCounter(L"compare.ui.progress_controls_text_applied_count");
        }
    }

    bool uiChanged = false;
    if (useDxProgressText)
    {
        if (_progress.scanProgressText && IsWindowVisible(_progress.scanProgressText.get()) != 0)
        {
            ShowWindow(_progress.scanProgressText.get(), SW_HIDE);
        }
        if (! _progress.controlsVisible || IsWindowVisible(_progress.scanProgressTextHostHwnd.get()) == 0)
        {
            ShowWindow(_progress.scanProgressTextHostHwnd.get(), SW_SHOWNA);
            uiChanged = true;
        }
    }
    else if (_progress.scanProgressText && (! _progress.controlsVisible || IsWindowVisible(_progress.scanProgressText.get()) == 0))
    {
        ShowWindow(_progress.scanProgressText.get(), SW_SHOW);
        uiChanged = true;
    }
    if (_progress.scanProgressBar && (! _progress.controlsVisible || IsWindowVisible(_progress.scanProgressBar.get()) == 0))
    {
        ShowWindow(_progress.scanProgressBar.get(), SW_SHOW);
        Debug::Perf::EmitCounter(L"compare.ui.progress_controls_show_count");
        uiChanged = true;
    }
    if (_progress.scanProgressBar && (! _progress.controlsVisible || ! _progress.pulseTimerActive))
    {
        InvalidateRect(_progress.scanProgressBar.get(), nullptr, FALSE);
        Debug::Perf::EmitCounter(L"compare.ui.progress_controls_progressbar_invalidate_count");
    }
    if (! wasVisible)
    {
        Layout();
    }
    else if (! uiChanged && ! textApplied)
    {
        Debug::Perf::EmitCounter(L"compare.ui.progress_controls_skipped_count");
    }
    _progress.controlsVisible = true;
    UpdateCompareWatermark();
    EnsureProgressPulseTimer();
    Debug::Perf::Emit(L"compare.ui.progress_controls_update_us", L"", ElapsedUsSince(startedAt), 0u, 0u, S_OK);
}

void CompareDirectoriesWindow::EnsureProgressPulseTimer() noexcept
{
    if (! _hWnd || _progress.pulseTimerActive)
    {
        return;
    }

    const bool spinnerVisible = _progress.scanProgressBar && IsWindowVisible(_progress.scanProgressBar.get()) != 0;
    if (! spinnerVisible && _progress.watermarkState != CompareWatermarkState::InProgress)
    {
        return;
    }

    _progress.spinnerAngleDeg   = 0.0f;
    _progress.spinnerLastTickMs = GetTickCount64();
    static_cast<void>(TryStartCompareTimer(_hWnd.get(),
                                           kCompareProgressPulseTimerId,
                                           kCompareProgressPulseTimerIntervalMs,
                                           _progress.pulseTimerActive,
                                           L"CompareDirectories: failed to start progress pulse timer."));
}

void CompareDirectoriesWindow::StopProgressPulseTimerIfIdle() noexcept
{
    if (! _progress.pulseTimerActive)
    {
        return;
    }

    const bool spinnerVisible = _progress.scanProgressBar && IsWindowVisible(_progress.scanProgressBar.get()) != 0;
    if (spinnerVisible || _progress.watermarkState == CompareWatermarkState::InProgress)
    {
        return;
    }

    if (_hWnd)
    {
        KillTimer(_hWnd.get(), kCompareProgressPulseTimerId);
    }
    _progress.pulseTimerActive = false;
}

void CompareDirectoriesWindow::InvalidateCompareWatermarkPanesIfDue(ULONGLONG now) noexcept
{
    if (_progress.watermarkState != CompareWatermarkState::InProgress)
    {
        return;
    }

    constexpr ULONGLONG kInvalidateIntervalMs = 100;
    if (_progress.paneWatermarkLastInvalidateTickMs != 0 && (now - _progress.paneWatermarkLastInvalidateTickMs) < kInvalidateIntervalMs)
    {
        return;
    }

    _progress.paneWatermarkLastInvalidateTickMs = now;

    if (const HWND left = _folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left))
    {
        InvalidateRect(left, nullptr, FALSE);
    }
    if (const HWND right = _folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right))
    {
        InvalidateRect(right, nullptr, FALSE);
    }
}

void CompareDirectoriesWindow::OnProgressPulseTimer() noexcept
{
    if (! _hWnd || ! _progress.pulseTimerActive)
    {
        return;
    }

    const ULONGLONG now         = GetTickCount64();
    const bool spinnerVisible   = _progress.scanProgressBar && IsWindowVisible(_progress.scanProgressBar.get()) != 0;
    const ULONGLONG lastSpinner = _progress.spinnerLastTickMs;
    _progress.spinnerLastTickMs = now;

    if (spinnerVisible)
    {
        double deltaSec = 0.0;
        if (now > lastSpinner)
        {
            deltaSec = static_cast<double>(now - lastSpinner) / 1000.0;
        }

        constexpr float kSpinnerDegPerSec = 180.0f;
        _progress.spinnerAngleDeg += static_cast<float>(deltaSec * static_cast<double>(kSpinnerDegPerSec));
        while (_progress.spinnerAngleDeg >= 360.0f)
        {
            _progress.spinnerAngleDeg -= 360.0f;
        }

        InvalidateRect(_progress.scanProgressBar.get(), nullptr, FALSE);
    }

    InvalidateCompareWatermarkPanesIfDue(now);
    StopProgressPulseTimerIfIdle();
}

void CompareDirectoriesWindow::DrawProgressSpinner(HDC hdc, const RECT& bounds) noexcept
{
    if (! hdc)
    {
        return;
    }

    RECT rc = bounds;
    if (rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    HBRUSH bgBrush = _backgroundBrush ? _backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH));
    FillRect(hdc, &rc, bgBrush);

    const float width  = static_cast<float>(std::max(0L, rc.right - rc.left));
    const float height = static_cast<float>(std::max(0L, rc.bottom - rc.top));
    const float minDim = std::min(width, height);
    if (minDim <= 2.0f)
    {
        return;
    }

    // Keep a small safety margin so wide pen strokes never clip at the edges (which can cause visible artifacts).
    const int stroke   = std::clamp(static_cast<int>(std::lround(minDim * 0.08f)), 1, 2);
    const float radius = std::max(1.0f, (minDim * 0.5f) - (static_cast<float>(stroke) * 0.5f + 1.0f));
    const float innerR = radius * 0.55f;
    const float outerR = radius;

    const float cx = static_cast<float>(rc.left) + width * 0.5f;
    const float cy = static_cast<float>(rc.top) + height * 0.5f;

    const COLORREF bg     = _theme.windowBackground;
    const COLORREF accent = _theme.menu.selectionBg;

    const bool rainbowSpinner = _theme.menu.rainbowMode && ! _theme.highContrast;
    uint32_t rainbowSeedHash  = 0u;
    if (rainbowSpinner)
    {
        const std::wstring_view seed =
            _leftContext.rootPluginPath.empty() ? std::wstring_view(L"compare") : std::wstring_view(_leftContext.rootPluginPath.native());
        rainbowSeedHash = StableHash32(seed);
    }

    constexpr float kPi = 3.14159265358979323846f;
    const float baseRad = (_progress.spinnerAngleDeg - 90.0f) * (kPi / 180.0f);
    float rainbowHue    = 0.0f;
    float rainbowSat    = 0.0f;
    float rainbowVal    = 0.0f;
    if (rainbowSpinner)
    {
        rainbowHue = static_cast<float>(rainbowSeedHash % 360u);
        rainbowSat = _theme.menu.darkBase ? 0.70f : 0.55f;
        rainbowVal = _theme.menu.darkBase ? 0.95f : 0.85f;
    }

    D2DHdcPaint::Session paint;
    if (! paint.Begin(hdc, rc))
    {
        return;
    }

    for (int i = 0; i < kProgressSpinnerSegments; ++i)
    {
        const float t     = static_cast<float>(i) / static_cast<float>(kProgressSpinnerSegments);
        const float angle = baseRad + t * (2.0f * kPi);
        const float s     = std::sin(angle);
        const float c     = std::cos(angle);

        const int x1 = static_cast<int>(std::lround(cx + c * innerR));
        const int y1 = static_cast<int>(std::lround(cy + s * innerR));
        const int x2 = static_cast<int>(std::lround(cx + c * outerR));
        const int y2 = static_cast<int>(std::lround(cy + s * outerR));

        COLORREF segmentBase = accent;
        if (rainbowSpinner)
        {
            const float hueStep    = 360.0f / static_cast<float>(kProgressSpinnerSegments);
            const float hueDegrees = rainbowHue + static_cast<float>(i) * hueStep;
            segmentBase            = ColorToCOLORREF(ColorFromHSV(hueDegrees, rainbowSat, rainbowVal));
        }

        const float alpha       = 0.15f + 0.85f * (1.0f - t);
        const int overlayWeight = static_cast<int>(std::lround(std::clamp(alpha, 0.0f, 1.0f) * 255.0f));
        const COLORREF color    = UiMetrics::BlendColorRefWeightedTruncate(bg, segmentBase, overlayWeight, 255);

        paint.DrawLine(static_cast<float>(x1), static_cast<float>(y1), static_cast<float>(x2), static_cast<float>(y2), color, static_cast<float>(stroke));
    }
}

void CompareDirectoriesWindow::UpdateCompareWatermark() noexcept
{
    const bool optionsVisible = _optionsPanel.dlg && IsWindowVisible(_optionsPanel.dlg.get()) != 0;
    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);

    CompareWatermarkState desired = CompareWatermarkState::Hidden;
    if (_compareStarted && ! optionsVisible && _compareActive)
    {
        if (_progress.compareRunResultHr == cancelledHr)
        {
            desired = CompareWatermarkState::Cancelled;
        }
        else
        {
            const bool runBusy = _compareRunPending || _progress.banner.scanActiveScans > 0u || _progress.banner.contentPendingCompares > 0u;
            if (runBusy)
            {
                desired = CompareWatermarkState::InProgress;
            }
        }
    }

    if (desired == _progress.watermarkState)
    {
        return;
    }

    _progress.watermarkState = desired;
    if (desired == CompareWatermarkState::Hidden)
    {
        _folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Left, {}, false);
        _folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Right, {}, false);
        return;
    }

    const UINT textId = (desired == CompareWatermarkState::InProgress) ? IDS_COMPARE_WATERMARK_IN_PROGRESS : IDS_COMPARE_WATERMARK_CANCELLED;

    const std::wstring text = LoadStringResource(nullptr, textId);
    const bool animated     = desired == CompareWatermarkState::InProgress;

    _progress.paneWatermarkLastInvalidateTickMs = 0;
    _folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Left, text, animated);
    _folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Right, text, animated);
}

void CompareDirectoriesWindow::UpdateRescanButtonText() noexcept
{
    const bool runBusy          = _compareRunPending || _progress.banner.scanActiveScans > 0u || _progress.banner.contentPendingCompares > 0u;
    const bool shouldShowCancel = _compareActive && runBusy;
    if (shouldShowCancel == _chrome.rescanIsCancel)
    {
        return;
    }

    _chrome.rescanIsCancel  = shouldShowCancel;
    const UINT textId       = shouldShowCancel ? IDS_COMPARE_BANNER_CANCEL : IDS_COMPARE_BANNER_RESCAN;
    const std::wstring text = LoadStringResource(nullptr, textId);
    if (_chrome.bannerRescanButton)
    {
        SetWindowTextW(_chrome.bannerRescanButton.get(), text.c_str());
    }
    SyncDxBannerButtons();
    Layout();
    if (_chrome.bannerRescanButton)
    {
        InvalidateRect(_chrome.bannerRescanButton.get(), nullptr, TRUE);
    }
}

void CompareDirectoriesWindow::UpdateCompareTaskCard(bool finished, bool force) noexcept
{
    if (! finished && ! force && _progress.compareTaskId != 0)
    {
        const ULONGLONG nowTick = GetTickCount64();
        if (_progress.lastTaskCardUpdateTickMs != 0 && nowTick >= _progress.lastTaskCardUpdateTickMs &&
            (nowTick - _progress.lastTaskCardUpdateTickMs) < kCompareTaskCardUpdateMinIntervalMs)
        {
            const ULONGLONG elapsed = nowTick - _progress.lastTaskCardUpdateTickMs;
            const ULONGLONG delayMs = std::max<ULONGLONG>(1, kCompareTaskCardUpdateMinIntervalMs - elapsed);
            ScheduleCompareTaskCardTrailingFlush(delayMs);
            Debug::Perf::EmitCounter(L"compare.ui.taskcard_update_throttled_count");
            return;
        }
    }

    CancelCompareTaskCardTrailingFlush();

    Debug::Perf::EmitCounter(L"compare.ui.taskcard_update_count");
    const SteadyClock::time_point startedAt      = SteadyClock::now();
    const SteadyClock::time_point buildStartedAt = startedAt;

    FolderWindow::InformationalTaskUpdate update{};
    update.kind      = FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories;
    update.taskId    = _progress.compareTaskId;
    update.title     = LoadStringResource(nullptr, IDS_COMPARE_BANNER_TITLE);
    update.leftRoot  = _leftContext.rootPluginPath;
    update.rightRoot = _rightContext.rootPluginPath;

    update.scanActive = _compareRunPending && (_progress.banner.scanActiveScans > 0u || ! _progress.compareRunSawScanProgress);
    if (_progress.banner.scanActiveScans > 0u)
    {
        std::filesystem::path current = _progress.banner.scanRelativeFolder;
        if (! _progress.banner.scanEntryName.empty())
        {
            current /= std::filesystem::path(_progress.banner.scanEntryName);
        }
        update.scanCurrentRelative = std::move(current);
    }
    update.scanFolderCount         = _progress.banner.scanFolderCount;
    update.scanEntryCount          = _progress.banner.scanEntryCount;
    update.scanCandidateFileCount  = _progress.banner.scanContentCandidateFileCount;
    update.scanCandidateTotalBytes = static_cast<uint64_t>(_progress.banner.scanContentCandidateTotalBytes);
    if (update.scanActive && _progress.scanStartTickMs != 0)
    {
        update.scanElapsedSeconds = (GetTickCount64() - _progress.scanStartTickMs) / 1000u;
    }

    update.contentActive = _progress.banner.contentPendingCompares > 0u;
    if (update.contentActive)
    {
        std::filesystem::path current = _progress.banner.contentRelativeFolder;
        if (! _progress.banner.contentEntryName.empty())
        {
            current /= std::filesystem::path(_progress.banner.contentEntryName);
        }
        update.contentCurrentRelative = std::move(current);
    }
    update.contentCurrentTotalBytes     = static_cast<uint64_t>(_progress.banner.contentFileTotalBytes);
    update.contentCurrentCompletedBytes = static_cast<uint64_t>(_progress.banner.contentFileCompletedBytes);
    update.contentTotalBytes            = static_cast<uint64_t>(_progress.banner.contentOverallTotalBytes);
    update.contentCompletedBytes        = static_cast<uint64_t>(_progress.banner.contentOverallCompletedBytes);
    update.contentPendingCount          = _progress.banner.contentPendingCompares;
    update.contentCompletedCount        = _progress.banner.contentCompletedCompares;
    if (update.contentActive && _progress.contentEtaSeconds.has_value())
    {
        update.contentEtaSeconds = _progress.contentEtaSeconds;
    }

    for (const auto& slot : _progress.banner.contentInFlight)
    {
        if (update.contentInFlightCount >= update.contentInFlight.size())
        {
            break;
        }
        if (slot.lastUpdateTick == 0 || slot.relativePath.empty())
        {
            continue;
        }

        FolderWindow::InformationalTaskUpdate::ContentInFlightFile entry{};
        entry.relativePath                                  = slot.relativePath;
        entry.totalBytes                                    = slot.totalBytes;
        entry.completedBytes                                = slot.completedBytes;
        entry.lastUpdateTick                                = slot.lastUpdateTick;
        update.contentInFlight[update.contentInFlightCount] = std::move(entry);
        ++update.contentInFlightCount;
    }

    update.finished = finished;
    if (finished)
    {
        update.resultHr = _progress.compareRunResultHr;

        if (_progress.banner.contentTotalCompares > 0u)
        {
            update.doneSummary = FormatStringResource(nullptr,
                                                      IDS_FMT_COMPARE_DONE_SUMMARY,
                                                      _progress.banner.scanFolderCount,
                                                      _progress.banner.scanEntryCount,
                                                      _progress.banner.contentCompletedCompares,
                                                      _progress.banner.contentTotalCompares);
        }
        else
        {
            update.doneSummary =
                FormatStringResource(nullptr, IDS_FMT_COMPARE_DONE_SUMMARY_SCAN_ONLY, _progress.banner.scanFolderCount, _progress.banner.scanEntryCount);
        }
    }

    const uint64_t buildUs = ElapsedUsSince(buildStartedAt);
    Debug::Perf::Emit(L"compare.ui.taskcard_build_us", L"", buildUs, update.contentPendingCount, update.scanEntryCount, S_OK);

    const SteadyClock::time_point applyStartedAt = SteadyClock::now();
    _progress.compareTaskId                      = _folderWindow.CreateOrUpdateInformationalTask(update);
    const uint64_t applyUs                       = ElapsedUsSince(applyStartedAt);
    _progress.lastTaskCardUpdateTickMs           = GetTickCount64();
    Debug::Perf::Emit(L"compare.ui.taskcard_apply_us", L"", applyUs, update.contentPendingCount, update.scanEntryCount, S_OK);
    Debug::Perf::Emit(L"compare.ui.taskcard_update_us", L"", ElapsedUsSince(startedAt), update.contentPendingCount, update.scanEntryCount, S_OK);
}

void CompareDirectoriesWindow::ScheduleCompareTaskCardTrailingFlush(ULONGLONG delayMs) noexcept
{
    if (! _hWnd || _progress.taskCardTrailingFlushPending)
    {
        return;
    }

    _progress.taskCardTrailingFlushPending = true;
    const UINT delay                       = static_cast<UINT>(std::clamp<ULONGLONG>(delayMs, 1, std::numeric_limits<UINT>::max()));
    bool timerActive                       = false;
    if (TryStartCompareTimer(
            _hWnd.get(), kCompareTaskCardTrailingFlushTimerId, delay, timerActive, L"CompareDirectories: failed to start task-card trailing flush timer."))
    {
        return;
    }

    _progress.taskCardTrailingFlushPending = false;
    UpdateCompareTaskCard(false, true);
}

void CompareDirectoriesWindow::CancelCompareTaskCardTrailingFlush() noexcept
{
    if (! _progress.taskCardTrailingFlushPending)
    {
        return;
    }

    if (_hWnd)
    {
        KillTimer(_hWnd.get(), kCompareTaskCardTrailingFlushTimerId);
    }
    _progress.taskCardTrailingFlushPending = false;
}

void CompareDirectoriesWindow::OnCompareTaskCardTrailingFlushTimer() noexcept
{
    if (_hWnd)
    {
        KillTimer(_hWnd.get(), kCompareTaskCardTrailingFlushTimerId);
    }
    _progress.taskCardTrailingFlushPending = false;

    if (_progress.compareTaskId != 0 && _compareActive && _compareRunPending)
    {
        UpdateCompareTaskCard(false, true);
    }
}

void CompareDirectoriesWindow::MaybeCompleteCompareRun() noexcept
{
    if (! _compareActive || ! _compareRunPending)
    {
        return;
    }

    if (_progress.banner.scanActiveScans != 0u || _progress.banner.contentPendingCompares != 0u)
    {
        return;
    }

    // Content progress resets (e.g. SetRoots/Invalidate) can post "idle" updates before any scan begins.
    // Don't mark the run complete until we see scan progress (or the run was canceled/failed).
    if (! _progress.compareRunSawScanProgress && _progress.compareRunResultHr == S_OK)
    {
        return;
    }

    if (_session)
    {
        LogComparePerfStats(L"done", _session, _progress.compareRunResultHr);
    }

    _compareRunPending = false;
    UpdateRescanButtonText();

    UpdateCompareTaskCard(true);
    if (_hWnd)
    {
        SetTimer(_hWnd.get(), kCompareTaskAutoDismissTimerId, kCompareTaskAutoDismissDelayMs, nullptr);
    }

    UpdateProgressControls();

    if (const auto leftPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Left))
    {
        UpdateEmptyStateForFolder(ComparePane::Left, leftPath.value());
    }
    if (const auto rightPath = _folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Right))
    {
        UpdateEmptyStateForFolder(ComparePane::Right, rightPath.value());
    }
}

void CompareDirectoriesWindow::DismissCompareTaskCard() noexcept
{
    CancelCompareTaskCardTrailingFlush();

    if (_progress.compareTaskId == 0)
    {
        _progress.lastTaskCardUpdateTickMs = 0;
        return;
    }

    _folderWindow.DismissInformationalTask(_progress.compareTaskId);
    _progress.compareTaskId            = 0;
    _progress.lastTaskCardUpdateTickMs = 0;
}

} // namespace CompareDirectoriesWindowInternal
