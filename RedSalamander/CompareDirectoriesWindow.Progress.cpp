#include "Framework.h"

#include "CompareDirectoriesWindow.Internal.h"

namespace CompareDirectoriesWindowInternal
{
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
};

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
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kCompareDirectoriesScanProgress, 0, std::move(payload)));
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
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kCompareDirectoriesContentProgress, 0, std::move(payload)));
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

    if (_decisionRefreshTimerActive || ! _hWnd)
    {
        return;
    }

    _decisionRefreshTimerActive = SetTimer(_hWnd.get(), kCompareDecisionRefreshTimerId, kCompareDecisionRefreshTimerIntervalMs, nullptr) != 0;
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

        if (const auto leftPath = _folderWindow.GetCurrentPath(FolderWindow::Pane::Left))
        {
            ApplySelectionForFolder(ComparePane::Left, leftPath.value());
            UpdateEmptyStateForFolder(ComparePane::Left, leftPath.value());
        }
        if (const auto rightPath = _folderWindow.GetCurrentPath(FolderWindow::Pane::Right))
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

        const auto currentPathOpt = _folderWindow.GetCurrentPath(fwPane);
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

LRESULT CompareDirectoriesWindow::OnScanProgress(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<ScanProgressPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    if (! _compareActive || payload->runId != _compareRunId)
    {
        return 0;
    }

    _compareRunSawScanProgress = true;

    _progress.scanActiveScans                = payload->activeScans;
    _progress.scanFolderCount                = payload->folderCount;
    _progress.scanEntryCount                 = payload->entryCount;
    _progress.scanContentCandidateFileCount  = payload->contentCandidateFileCount;
    _progress.scanContentCandidateTotalBytes = payload->contentCandidateTotalBytes;
    _progress.scanRelativeFolder             = std::move(payload->relativeFolder);
    _progress.scanEntryName                  = std::move(payload->entryName);

    if (_progress.scanActiveScans == 0u)
    {
        _progress.scanRelativeFolder.clear();
        _progress.scanEntryName.clear();
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

LRESULT CompareDirectoriesWindow::OnContentProgress(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<ContentProgressPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    if (! _compareActive || payload->runId != _compareRunId)
    {
        return 0;
    }

    const ULONGLONG nowTick = GetTickCount64();

    _progress.contentPendingCompares       = payload->pendingContentCompares;
    _progress.contentTotalCompares         = payload->totalContentCompares;
    _progress.contentCompletedCompares     = payload->completedContentCompares;
    _progress.contentOverallTotalBytes     = payload->overallTotalBytes;
    _progress.contentOverallCompletedBytes = payload->overallCompletedBytes;
    _progress.contentFileTotalBytes        = payload->fileTotalBytes;
    _progress.contentFileCompletedBytes    = payload->fileCompletedBytes;

    if (_progress.contentPendingCompares > 0u)
    {
        std::filesystem::path fileRel = payload->relativeFolder;
        if (! payload->entryName.empty())
        {
            fileRel /= std::filesystem::path(payload->entryName);
        }

        if (! fileRel.empty())
        {
            const uint32_t slotIndex = payload->workerIndex;
            if (slotIndex < _progress.contentInFlight.size())
            {
                auto& slot          = _progress.contentInFlight[slotIndex];
                slot.relativePath   = std::move(fileRel);
                slot.totalBytes     = payload->fileTotalBytes;
                slot.completedBytes = payload->fileCompletedBytes;
                slot.lastUpdateTick = nowTick;
            }
        }
    }
    else
    {
        for (auto& slot : _progress.contentInFlight)
        {
            slot = {};
        }
    }

    _progress.contentRelativeFolder = std::move(payload->relativeFolder);
    _progress.contentEntryName      = std::move(payload->entryName);

    if (_progress.contentPendingCompares > 0u)
    {
        const uint64_t completed = _progress.contentOverallCompletedBytes;
        const uint64_t total     = _progress.contentOverallTotalBytes;

        if (_contentEtaLastTickMs != 0u && nowTick > _contentEtaLastTickMs && completed >= _contentEtaLastCompletedBytes)
        {
            const uint64_t deltaBytes = completed - _contentEtaLastCompletedBytes;
            const double deltaSeconds = static_cast<double>(nowTick - _contentEtaLastTickMs) / 1000.0;
            if (deltaBytes > 0u && deltaSeconds >= 0.2)
            {
                const double rate = static_cast<double>(deltaBytes) / deltaSeconds;
                if (_contentEtaSmoothedBytesPerSec <= 1.0)
                {
                    _contentEtaSmoothedBytesPerSec = rate;
                }
                else
                {
                    constexpr double kAlpha        = 0.15;
                    _contentEtaSmoothedBytesPerSec = (_contentEtaSmoothedBytesPerSec * (1.0 - kAlpha)) + (rate * kAlpha);
                }
            }
        }

        _contentEtaLastTickMs         = nowTick;
        _contentEtaLastCompletedBytes = completed;

        _contentEtaSeconds.reset();
        if (total > 0u && completed <= total && _contentEtaSmoothedBytesPerSec > 1.0)
        {
            const uint64_t remaining = total - completed;
            const double secondsD    = static_cast<double>(remaining) / _contentEtaSmoothedBytesPerSec;
            _contentEtaSeconds       = static_cast<uint64_t>(std::ceil(std::max(0.0, secondsD)));
        }
    }
    else
    {
        _contentEtaLastTickMs          = 0;
        _contentEtaLastCompletedBytes  = 0;
        _contentEtaSmoothedBytesPerSec = 0.0;
        _contentEtaSeconds.reset();
    }

    if (_progress.contentPendingCompares == 0u)
    {
        _progress.contentFileTotalBytes     = 0;
        _progress.contentFileCompletedBytes = 0;
        _progress.contentRelativeFolder.clear();
        _progress.contentEntryName.clear();
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
    if (! _scanProgressText && ! _scanProgressBar)
    {
        UpdateCompareWatermark();
        return;
    }

    const bool show = (_compareActive && _compareRunPending) || _progress.scanActiveScans > 0u || _progress.contentPendingCompares > 0u;
    const bool wasVisible =
        (_scanProgressText && IsWindowVisible(_scanProgressText.get()) != 0) || (_scanProgressBar && IsWindowVisible(_scanProgressBar.get()) != 0);

    if (! show)
    {
        if (_progressSpinnerTimerActive && _hWnd)
        {
            KillTimer(_hWnd.get(), kCompareBannerSpinnerTimerId);
            _progressSpinnerTimerActive = false;
        }

        if (_scanProgressBar)
        {
            ShowWindow(_scanProgressBar.get(), SW_HIDE);
        }
        if (_scanProgressText)
        {
            SetWindowTextW(_scanProgressText.get(), L"");
            ShowWindow(_scanProgressText.get(), SW_HIDE);
        }
        if (wasVisible)
        {
            Layout();
        }
        UpdateCompareWatermark();
        return;
    }

    std::wstring scanText;
    if (_progress.scanActiveScans > 0u || (_compareActive && _compareRunPending && _progress.contentPendingCompares == 0u))
    {
        std::filesystem::path displayPath = _progress.scanRelativeFolder;
        if (! _progress.scanEntryName.empty())
        {
            displayPath /= std::filesystem::path(_progress.scanEntryName);
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

        scanText = FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, pathText, _progress.scanFolderCount, _progress.scanEntryCount);
        if (_scanStartTickMs != 0)
        {
            const uint64_t elapsedSec   = (GetTickCount64() - _scanStartTickMs) / 1000u;
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
    if (_progress.contentPendingCompares > 0u && ! _progress.contentEntryName.empty())
    {
        std::filesystem::path displayPath = _progress.contentRelativeFolder;
        if (! _progress.contentEntryName.empty())
        {
            displayPath /= std::filesystem::path(_progress.contentEntryName);
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

        const std::wstring completedText = FormatBytesCompact(_progress.contentFileCompletedBytes);
        if (_progress.contentFileTotalBytes > 0u)
        {
            const std::wstring totalText = FormatBytesCompact(_progress.contentFileTotalBytes);
            contentText                  = FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_STATUS, pathText, completedText, totalText);
        }
        else
        {
            contentText = FormatStringResource(nullptr, IDS_FMT_COMPARE_CONTENT_STATUS_UNKNOWN, pathText, completedText);
        }

        if (_contentEtaSeconds.has_value())
        {
            const std::wstring duration = FormatDurationHmsNoexcept(_contentEtaSeconds.value());
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

    if (_scanProgressText)
    {
        SetWindowTextW(_scanProgressText.get(), message.c_str());
    }

    if (_scanProgressText)
    {
        ShowWindow(_scanProgressText.get(), SW_SHOW);
    }
    if (_scanProgressBar)
    {
        ShowWindow(_scanProgressBar.get(), SW_SHOW);
        InvalidateRect(_scanProgressBar.get(), nullptr, FALSE);
    }
    if (! _progressSpinnerTimerActive && _hWnd && _scanProgressBar)
    {
        _progressSpinnerAngleDeg    = 0.0f;
        _progressSpinnerLastTickMs  = GetTickCount64();
        _progressSpinnerTimerActive = SetTimer(_hWnd.get(), kCompareBannerSpinnerTimerId, kCompareBannerSpinnerTimerIntervalMs, nullptr) != 0;
    }
    if (! wasVisible)
    {
        Layout();
    }
    UpdateCompareWatermark();
}

void CompareDirectoriesWindow::OnProgressSpinnerTimer() noexcept
{
    if (! _hWnd || ! _scanProgressBar || ! _progressSpinnerTimerActive)
    {
        return;
    }

    if (IsWindowVisible(_scanProgressBar.get()) == 0)
    {
        return;
    }

    const ULONGLONG now        = GetTickCount64();
    const ULONGLONG last       = _progressSpinnerLastTickMs;
    _progressSpinnerLastTickMs = now;

    double deltaSec = 0.0;
    if (now > last)
    {
        deltaSec = static_cast<double>(now - last) / 1000.0;
    }

    constexpr float kSpinnerDegPerSec = 180.0f;
    _progressSpinnerAngleDeg += static_cast<float>(deltaSec * static_cast<double>(kSpinnerDegPerSec));
    while (_progressSpinnerAngleDeg >= 360.0f)
    {
        _progressSpinnerAngleDeg -= 360.0f;
    }

    InvalidateRect(_scanProgressBar.get(), nullptr, FALSE);
    if (_watermarkState == CompareWatermarkState::InProgress)
    {
        constexpr ULONGLONG kInvalidateIntervalMs = 100;
        if (_paneWatermarkLastInvalidateTickMs == 0 || (now - _paneWatermarkLastInvalidateTickMs) >= kInvalidateIntervalMs)
        {
            _paneWatermarkLastInvalidateTickMs = now;

            if (const HWND left = _folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left))
            {
                InvalidateRect(left, nullptr, FALSE);
            }
            if (const HWND right = _folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right))
            {
                InvalidateRect(right, nullptr, FALSE);
            }
        }
    }
}

void CompareDirectoriesWindow::InvalidateSpinnerPens() noexcept
{
    _progressSpinnerPenKeyValid = false;
    for (auto& pen : _progressSpinnerPens)
    {
        pen.reset();
    }
}

void CompareDirectoriesWindow::EnsureSpinnerPens(COLORREF background, COLORREF accent, bool rainbowSpinner, uint32_t rainbowSeedHash, int stroke) noexcept
{
    if (stroke <= 0)
    {
        InvalidateSpinnerPens();
        return;
    }

    ProgressSpinnerPenKey key{};
    key.background    = background;
    key.accent        = accent;
    key.rainbow       = rainbowSpinner;
    key.darkBase      = _theme.menu.darkBase;
    key.seedHash      = rainbowSpinner ? rainbowSeedHash : 0u;
    key.strokeWidthPx = stroke;

    if (_progressSpinnerPenKeyValid && _progressSpinnerPenKey == key)
    {
        bool allValid = true;
        for (const auto& pen : _progressSpinnerPens)
        {
            allValid = allValid && static_cast<bool>(pen);
        }
        if (allValid)
        {
            return;
        }
    }

    _progressSpinnerPenKey      = key;
    _progressSpinnerPenKeyValid = true;
    for (auto& pen : _progressSpinnerPens)
    {
        pen.reset();
    }

    float rainbowHue = 0.0f;
    float rainbowSat = 0.0f;
    float rainbowVal = 0.0f;
    if (rainbowSpinner)
    {
        rainbowHue = static_cast<float>(rainbowSeedHash % 360u);
        rainbowSat = key.darkBase ? 0.70f : 0.55f;
        rainbowVal = key.darkBase ? 0.95f : 0.85f;
    }

    for (int i = 0; i < kProgressSpinnerSegments; ++i)
    {
        const float t     = static_cast<float>(i) / static_cast<float>(kProgressSpinnerSegments);
        const float alpha = 0.15f + 0.85f * (1.0f - t);

        COLORREF segmentBase = accent;
        if (rainbowSpinner)
        {
            const float hueStep    = 360.0f / static_cast<float>(kProgressSpinnerSegments);
            const float hueDegrees = rainbowHue + static_cast<float>(i) * hueStep;
            segmentBase            = ColorToCOLORREF(ColorFromHSV(hueDegrees, rainbowSat, rainbowVal));
        }

        const int overlayWeight = static_cast<int>(std::lround(std::clamp(alpha, 0.0f, 1.0f) * 255.0f));
        const COLORREF color    = ThemedControls::BlendColor(background, segmentBase, overlayWeight, 255);

        _progressSpinnerPens[static_cast<size_t>(i)].reset(CreatePen(PS_SOLID, stroke, color));
    }
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
    const float baseRad = (_progressSpinnerAngleDeg - 90.0f) * (kPi / 180.0f);

    EnsureSpinnerPens(bg, accent, rainbowSpinner, rainbowSeedHash, stroke);

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

        wil::unique_hpen penFallback;
        HPEN pen = _progressSpinnerPens[static_cast<size_t>(i)].get();
        if (! pen)
        {
            penFallback.reset(CreatePen(PS_SOLID, stroke, accent));
            pen = penFallback.get();
        }

        if (! pen)
        {
            continue;
        }

        [[maybe_unused]] auto oldPen = wil::SelectObject(hdc, pen);
        MoveToEx(hdc, x1, y1, nullptr);
        LineTo(hdc, x2, y2);
    }
}

void CompareDirectoriesWindow::UpdateCompareWatermark() noexcept
{
    const bool optionsVisible = _optionsDlg && IsWindowVisible(_optionsDlg.get()) != 0;
    const HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);

    CompareWatermarkState desired = CompareWatermarkState::Hidden;
    if (_compareStarted && ! optionsVisible && _compareActive)
    {
        if (_compareRunResultHr == cancelledHr)
        {
            desired = CompareWatermarkState::Cancelled;
        }
        else
        {
            const bool runBusy = _compareRunPending || _progress.scanActiveScans > 0u || _progress.contentPendingCompares > 0u;
            if (runBusy)
            {
                desired = CompareWatermarkState::InProgress;
            }
        }
    }

    if (desired == _watermarkState)
    {
        return;
    }

    _watermarkState = desired;
    if (desired == CompareWatermarkState::Hidden)
    {
        _folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Left, {}, false);
        _folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Right, {}, false);
        return;
    }

    const UINT textId = (desired == CompareWatermarkState::InProgress) ? IDS_COMPARE_WATERMARK_IN_PROGRESS : IDS_COMPARE_WATERMARK_CANCELLED;

    const std::wstring text = LoadStringResource(nullptr, textId);
    const bool animated     = desired == CompareWatermarkState::InProgress;

    _paneWatermarkLastInvalidateTickMs = 0;
    _folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Left, text, animated);
    _folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Right, text, animated);
}

void CompareDirectoriesWindow::UpdateRescanButtonText() noexcept
{
    if (! _bannerRescanButton)
    {
        return;
    }

    const bool runBusy          = _compareRunPending || _progress.scanActiveScans > 0u || _progress.contentPendingCompares > 0u;
    const bool shouldShowCancel = _compareActive && runBusy;
    if (shouldShowCancel == _bannerRescanIsCancel)
    {
        return;
    }

    _bannerRescanIsCancel   = shouldShowCancel;
    const UINT textId       = shouldShowCancel ? IDS_COMPARE_BANNER_CANCEL : IDS_COMPARE_BANNER_RESCAN;
    const std::wstring text = LoadStringResource(nullptr, textId);
    SetWindowTextW(_bannerRescanButton.get(), text.c_str());
    Layout();
    InvalidateRect(_bannerRescanButton.get(), nullptr, TRUE);
}

void CompareDirectoriesWindow::UpdateCompareTaskCard(bool finished) noexcept
{
    FolderWindow::InformationalTaskUpdate update{};
    update.kind      = FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories;
    update.taskId    = _compareTaskId;
    update.title     = LoadStringResource(nullptr, IDS_COMPARE_BANNER_TITLE);
    update.leftRoot  = _leftContext.rootPluginPath;
    update.rightRoot = _rightContext.rootPluginPath;

    update.scanActive = _compareRunPending && (_progress.scanActiveScans > 0u || ! _compareRunSawScanProgress);
    if (_progress.scanActiveScans > 0u)
    {
        std::filesystem::path current = _progress.scanRelativeFolder;
        if (! _progress.scanEntryName.empty())
        {
            current /= std::filesystem::path(_progress.scanEntryName);
        }
        update.scanCurrentRelative = std::move(current);
    }
    update.scanFolderCount         = _progress.scanFolderCount;
    update.scanEntryCount          = _progress.scanEntryCount;
    update.scanCandidateFileCount  = _progress.scanContentCandidateFileCount;
    update.scanCandidateTotalBytes = static_cast<uint64_t>(_progress.scanContentCandidateTotalBytes);
    if (update.scanActive && _scanStartTickMs != 0)
    {
        update.scanElapsedSeconds = (GetTickCount64() - _scanStartTickMs) / 1000u;
    }

    update.contentActive = _progress.contentPendingCompares > 0u;
    if (update.contentActive)
    {
        std::filesystem::path current = _progress.contentRelativeFolder;
        if (! _progress.contentEntryName.empty())
        {
            current /= std::filesystem::path(_progress.contentEntryName);
        }
        update.contentCurrentRelative = std::move(current);
    }
    update.contentCurrentTotalBytes     = static_cast<uint64_t>(_progress.contentFileTotalBytes);
    update.contentCurrentCompletedBytes = static_cast<uint64_t>(_progress.contentFileCompletedBytes);
    update.contentTotalBytes            = static_cast<uint64_t>(_progress.contentOverallTotalBytes);
    update.contentCompletedBytes        = static_cast<uint64_t>(_progress.contentOverallCompletedBytes);
    update.contentPendingCount          = _progress.contentPendingCompares;
    update.contentCompletedCount        = _progress.contentCompletedCompares;
    if (update.contentActive && _contentEtaSeconds.has_value())
    {
        update.contentEtaSeconds = _contentEtaSeconds;
    }

    for (const auto& slot : _progress.contentInFlight)
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
        update.resultHr = _compareRunResultHr;

        if (_progress.contentTotalCompares > 0u)
        {
            update.doneSummary = FormatStringResource(nullptr,
                                                      IDS_FMT_COMPARE_DONE_SUMMARY,
                                                      _progress.scanFolderCount,
                                                      _progress.scanEntryCount,
                                                      _progress.contentCompletedCompares,
                                                      _progress.contentTotalCompares);
        }
        else
        {
            update.doneSummary = FormatStringResource(nullptr, IDS_FMT_COMPARE_DONE_SUMMARY_SCAN_ONLY, _progress.scanFolderCount, _progress.scanEntryCount);
        }
    }

    _compareTaskId = _folderWindow.CreateOrUpdateInformationalTask(update);
}

void CompareDirectoriesWindow::MaybeCompleteCompareRun() noexcept
{
    if (! _compareActive || ! _compareRunPending)
    {
        return;
    }

    if (_progress.scanActiveScans != 0u || _progress.contentPendingCompares != 0u)
    {
        return;
    }

    // Content progress resets (e.g. SetRoots/Invalidate) can post "idle" updates before any scan begins.
    // Don't mark the run complete until we see scan progress (or the run was canceled/failed).
    if (! _compareRunSawScanProgress && _compareRunResultHr == S_OK)
    {
        return;
    }

    if (_session)
    {
        LogComparePerfStats(L"done", _session, _compareRunResultHr);
    }

    _compareRunPending = false;
    UpdateRescanButtonText();

    UpdateCompareTaskCard(true);
    if (_hWnd)
    {
        SetTimer(_hWnd.get(), kCompareTaskAutoDismissTimerId, kCompareTaskAutoDismissDelayMs, nullptr);
    }

    UpdateProgressControls();

    if (const auto leftPath = _folderWindow.GetCurrentPath(FolderWindow::Pane::Left))
    {
        UpdateEmptyStateForFolder(ComparePane::Left, leftPath.value());
    }
    if (const auto rightPath = _folderWindow.GetCurrentPath(FolderWindow::Pane::Right))
    {
        UpdateEmptyStateForFolder(ComparePane::Right, rightPath.value());
    }
}

void CompareDirectoriesWindow::DismissCompareTaskCard() noexcept
{
    if (_compareTaskId == 0)
    {
        return;
    }

    _folderWindow.DismissInformationalTask(_compareTaskId);
    _compareTaskId = 0;
}

} // namespace CompareDirectoriesWindowInternal
