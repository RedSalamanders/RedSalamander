// Opt-in FileOps fixture. Include from FolderWindow.FileOperations.Tests.Runner.cpp under ENABLE_TESTS.
// Install identical bytes on the current I26 product before and after its DxUi pin change.
[[nodiscard]] bool RunFileOpsIntegratedPresentationPerf(SelfTestState& state) noexcept
{
    using namespace FileOperationsPopupInternal;
    using Clock       = std::chrono::steady_clock;
    using Diagnostic  = FolderWindow::FileOperationState::TaskDiagnosticEntry;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps integrated presentation perf: {}", reason));
        return false;
    };
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread())
        return reject(L"no-activation guard is absent");
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = folder ? TryGetFileOps(folder) : nullptr;
    if (! folder || ! operations || operations->HasActiveOperations() || operations->GetPopupHwndForSelfTest() || operations->GetIssuesPaneHwndForSelfTest())
        return reject(L"File Operations must be idle with no retained windows");
    std::vector<Diagnostic> diagnostics;
    std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
    operations->CollectDiagnostics(diagnostics);
    operations->CollectCompletedTasks(completed);
    if (! diagnostics.empty() || ! completed.empty())
        return reject(L"pre-existing history would change the bounded workload");

    const auto previousTheme      = folder->GetTheme();
    const bool previousFooter     = operations->GetPopupFooterOnly();
    const bool previousDensity    = operations->GetPopupCompactDensity();
    const bool previousThroughput = operations->GetPopupThroughputExpanded();
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui->language.empty() && g_settings.ui->language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui->language};
    bool restored      = false;
    const auto restore = [&]() noexcept
    {
        folder->ApplyTheme(previousTheme);
        operations->SetPopupFooterOnly(previousFooter);
        operations->SetPopupCompactDensity(previousDensity);
        operations->SetPopupThroughputExpanded(previousThroughput);
        const HRESULT languageHr = Localization::ApplyLanguagePreference(previousLanguage);
        const bool saveSettled   = operations->DebugFlushPendingSettingsSaveForSelfTest(10000u);
        return SUCCEEDED(languageHr) && saveSettled;
    };
    auto restoreSettings = wil::scope_exit([&]() noexcept
    {
        if (! restored)
            static_cast<void>(restore());
    });
    if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, L"fr-FR"})))
        return reject(L"French resources unavailable");
    auto theme                  = ResolveAppTheme(ThemeMode::Dark, L"FileOpsIntegratedPresentationPerf");
    theme.reducedMotionOverride = true;
    folder->ApplyTheme(theme);
    operations->SetPopupFooterOnly(false);
    operations->SetPopupCompactDensity(false);
    operations->SetPopupThroughputExpanded(true);
    if (! operations->DebugFlushPendingSettingsSaveForSelfTest(10000u))
        return reject(L"fixture settings did not settle");

    struct ResourceSample
    {
        const wchar_t* phase  = L"";
        uint64_t privateBytes = 0u, workingSetBytes = 0u, processCpuUs = 0u, uiCpuUs = 0u;
        uint64_t handles = 0u, gdi = 0u, user = 0u, hosts = 0u, attachments = 0u;
    };
    std::array<ResourceSample, 40u> resources{};
    size_t resourceCount   = 0u;
    using SurfaceResources = FileOperationsSelfTest::SurfaceResources;
    struct SurfaceSample
    {
        const wchar_t* phase = L"";
        SurfaceResources value{};
    };
    std::array<SurfaceSample, 32u> surfaces{};
    size_t surfaceCount      = 0u;
    const auto sampleSurface = [&](const wchar_t* phase, const SurfaceResources& value) noexcept
    {
        if (surfaceCount == surfaces.size())
            return false;
        surfaces[surfaceCount++] = {phase, value};
        return value.logicalBytesAvailable && value.width != 0u && value.height != 0u && value.renderCount != 0u;
    };
    const auto sameSurface = [](const SurfaceResources& before, const SurfaceResources& after) noexcept
    {
        return before.logicalBytesAvailable && after.logicalBytesAvailable && before.width == after.width && before.height == after.height &&
               before.format == after.format && before.singleTargetLogicalBytes == after.singleTargetLogicalBytes &&
               before.brushEntries == after.brushEntries && before.textFormatEntries == after.textFormatEntries && before.resizeCount == after.resizeCount;
    };
    const auto matchesClient = [](const SurfaceResources& value, const RECT& client) noexcept
    {
        return client.right > client.left && client.bottom > client.top && value.width == static_cast<uint32_t>(client.right - client.left) &&
               value.height == static_cast<uint32_t>(client.bottom - client.top);
    };
    std::array<uint64_t, 5u> graphOpen8Us{}, graphCollapsed8Us{}, issuesDetail8Us{};
    std::array<uint64_t, 32u> churnUpdate8Us{};
    size_t openCount = 0u, collapsedCount = 0u, churnCount = 0u, issuesCount = 0u;
    constexpr std::array<const wchar_t*, 4u> scanKindNames{{L"compare", L"change-case", L"change-attributes", L"file-list"}};
    std::array<std::array<uint64_t, 5u>, 4u> scanClean8Us{}, scanChanging8Us{};
    std::array<size_t, 4u> scanCleanCount{}, scanChangingCount{};
    uint64_t scanDigest     = 14695981039346656037ull;
    const auto hashScanWord = [&](uint64_t value) noexcept
    {
        for (unsigned int byte = 0u; byte < 8u; ++byte)
        {
            scanDigest ^= (value >> (byte * 8u)) & 0xFFu;
            scanDigest *= 1099511628211ull;
        }
    };
    const auto hashScanText = [&](std::wstring_view value) noexcept
    {
        hashScanWord(value.size());
        for (const wchar_t codeUnit : value)
            hashScanWord(static_cast<uint16_t>(codeUnit));
    };
    RECT popupClient{}, issuesClient{};
    UINT popupDpi = 0u, issuesDpi = 0u;
    uint64_t digest     = 14695981039346656037ull;
    const auto hashWord = [&](uint64_t value) noexcept
    {
        for (unsigned int byte = 0u; byte < 8u; ++byte)
        {
            digest ^= (value >> (byte * 8u)) & 0xFFu;
            digest *= 1099511628211ull;
        }
    };
    const auto hashText = [&](std::wstring_view value) noexcept
    {
        hashWord(value.size());
        for (const wchar_t codeUnit : value)
            hashWord(static_cast<uint16_t>(codeUnit));
    };
    hashText(L"FileOps.IntegratedPresentation.v2;fr-FR;dark;reduced-motion;default-density;fixed-760x480-DIP;exact-eight-frames");
    for (const uint64_t value : {8u, 127u, 1u, 32u, 8u, 4u * 1024u * 1024u})
        hashWord(value);
    const auto fileTimeUs       = [](FILETIME value) noexcept { return ((static_cast<uint64_t>(value.dwHighDateTime) << 32u) | value.dwLowDateTime) / 10u; };
    const DWORD uiThreadId      = GetCurrentThreadId();
    const uint64_t initialHosts = DxUi::DebugGetAttachedWindowHostCount();
    const uint64_t initialAttachments = DxUi::DebugGetSharedWindowHostAttachmentCountForThread(uiThreadId);
    const auto sample                 = [&](const wchar_t* phase) noexcept -> bool
    {
        if (resourceCount == resources.size())
            return false;
        const auto memory = FolderWindowFileOperationsStateInternal::CaptureProcessMemorySnapshot();
        DWORD handles     = 0u;
        FILETIME created{}, exited{}, kernel{}, user{};
        if (memory.privateBytes == 0u || memory.workingSetBytes == 0u || GetProcessHandleCount(GetCurrentProcess(), &handles) == FALSE ||
            GetProcessTimes(GetCurrentProcess(), &created, &exited, &kernel, &user) == FALSE)
            return false;
        ResourceSample value{};
        value.phase           = phase;
        value.privateBytes    = memory.privateBytes;
        value.workingSetBytes = memory.workingSetBytes;
        value.processCpuUs    = fileTimeUs(kernel) + fileTimeUs(user);
        if (GetThreadTimes(GetCurrentThread(), &created, &exited, &kernel, &user) == FALSE)
            return false;
        value.uiCpuUs              = fileTimeUs(kernel) + fileTimeUs(user);
        value.handles              = handles;
        value.gdi                  = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS);
        value.user                 = GetGuiResources(GetCurrentProcess(), GR_USEROBJECTS);
        value.hosts                = DxUi::DebugGetAttachedWindowHostCount();
        value.attachments          = DxUi::DebugGetSharedWindowHostAttachmentCountForThread(uiThreadId);
        resources[resourceCount++] = value;
        return true;
    };
    auto emitSamples = wil::scope_exit([&]() noexcept
    {
        for (size_t i = 0u; i < resourceCount; ++i)
        {
            const auto& s = resources[i];
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.Memory", s.phase, 0u, s.privateBytes, s.workingSetBytes, S_OK);
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.Cpu", s.phase, 0u, s.uiCpuUs, s.processCpuUs, S_OK);
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.Handles", s.phase, 0u, s.handles, s.hosts, S_OK);
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.GuiAndAttachments", s.phase, s.attachments, s.gdi, s.user, S_OK);
        }
        for (size_t i = 0u; i < openCount; ++i)
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.GraphOpenRedraw8Us", L"first-task-open", graphOpen8Us[i], i, 8u, S_OK);
        for (size_t i = 0u; i < collapsedCount; ++i)
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.GraphCollapsedRedraw8Us", L"first-task-collapsed", graphCollapsed8Us[i], i, 8u, S_OK);
        for (size_t i = 0u; i < churnCount; ++i)
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.ChurnUpdate8Us", L"eight-task-open", churnUpdate8Us[i], i, 8u, S_OK);
        for (size_t i = 0u; i < issuesCount; ++i)
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.IssuesDetailRedraw8Us", L"selected-typed-detail", issuesDetail8Us[i], i, 8u, S_OK);
        for (size_t kind = 0u; kind < scanKindNames.size(); ++kind)
        {
            for (size_t round = 0u; round < scanCleanCount[kind]; ++round)
                Debug::Perf::Emit(L"FileOps.IntegratedPresentation.ScanStatusCleanRedraw8Us", scanKindNames[kind], scanClean8Us[kind][round], round, 8u, S_OK);
            for (size_t round = 0u; round < scanChangingCount[kind]; ++round)
                Debug::Perf::Emit(
                    L"FileOps.IntegratedPresentation.ScanStatusChangingRedraw8Us", scanKindNames[kind], scanChanging8Us[kind][round], round, 8u, S_OK);
        }
        if (scanCleanCount[3u] == 5u && scanChangingCount[3u] == 5u)
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.ScanStatusWorkloadFNV64",
                              L"v2-four-informational-fixed-480x720dip-warm-range-exact-eight",
                              scanDigest,
                              4u,
                              40u,
                              S_OK);
        if (popupDpi != 0u)
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.ClientPixels",
                              L"popup",
                              popupDpi,
                              static_cast<uint64_t>(popupClient.right - popupClient.left),
                              static_cast<uint64_t>(popupClient.bottom - popupClient.top),
                              S_OK);
        if (issuesDpi != 0u)
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.ClientPixels",
                              L"issues",
                              issuesDpi,
                              static_cast<uint64_t>(issuesClient.right - issuesClient.left),
                              static_cast<uint64_t>(issuesClient.bottom - issuesClient.top),
                              S_OK);
        Debug::Perf::Emit(
            L"FileOps.IntegratedPresentation.WorkloadFNV64", L"v2-eight-running-fixed-760x480dip-exact-eight-127-warnings-one-typed", digest, 8u, 128u, S_OK);
        for (size_t i = 0u; i < surfaceCount; ++i)
        {
            const auto& s = surfaces[i];
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.SingleTargetLogicalBytes",
                              s.phase,
                              s.value.singleTargetLogicalBytes,
                              s.value.width,
                              s.value.height,
                              s.value.logicalBytesAvailable ? S_OK : E_NOTIMPL);
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.HostCacheEntries",
                              s.phase,
                              static_cast<uint64_t>(s.value.format),
                              s.value.brushEntries,
                              s.value.textFormatEntries,
                              S_OK);
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.HostRenderResize", s.phase, 0u, s.value.renderCount, s.value.resizeCount, S_OK);
        }
        // Current target logical bytes/cache entries are observable; actual driver allocations/peaks are not.
        Debug::Perf::Emit(L"FileOps.IntegratedPresentation.SurfaceCounters", L"driver-allocation-and-peak-unavailable", 0u, 0u, 0u, E_NOTIMPL);
    });
    if (! sample(L"cold"))
        return reject(L"cold sample unavailable");

    std::vector<TaskSnapshot> tasks;
    tasks.reserve(8u);
    const ULONGLONG now = GetTickCount64();
    for (uint64_t i = 0u; i < 8u; ++i)
    {
        TaskSnapshot task{};
        task.taskId         = 0xF330000000000000ull + i;
        task.operation      = (i % 2u) == 0u ? FILESYSTEM_COPY : FILESYSTEM_MOVE;
        task.lifecyclePhase = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
        task.started = task.discoveryClosed = task.hasProgressCallbacks = true;
        task.lastProgressCallbackTick                                   = now;
        task.operationStartTick                                         = now > 25'000u ? now - 25'000u : 0u;
        task.totalBytes                                                 = 8ull * 1024ull * 1024ull * 1024ull;
        task.completedBytes                                             = (i + 1u) * 256ull * 1024ull * 1024ull;
        task.totalItems                                                 = 128u;
        task.completedItems                                             = static_cast<unsigned long>(i + 1u);
        task.itemTotalBytes                                             = 512ull * 1024ull * 1024ull;
        task.itemCompletedBytes                                         = 192ull * 1024ull * 1024ull;
        task.effectiveConcurrencyBudget                                 = 4u;
        task.currentSourcePath                                          = std::format(L"C:\\Photos\\IntegratedPerf\\source-{:02}.raw", i);
        task.currentDestinationPath                                     = std::format(L"D:\\Archives\\IntegratedPerf\\source-{:02}.raw", i);
        hashWord(task.taskId);
        hashWord(static_cast<uint64_t>(task.operation));
        hashWord(task.lifecyclePhase);
        hashWord(task.totalBytes);
        hashWord(task.completedBytes);
        hashWord(task.totalItems);
        hashWord(task.completedItems);
        hashWord(task.itemTotalBytes);
        hashWord(task.itemCompletedBytes);
        hashText(task.currentSourcePath);
        hashText(task.currentDestinationPath);
        tasks.push_back(std::move(task));
    }
    const auto pumpBatch = []() noexcept -> bool
    {
        MSG message{};
        for (unsigned int drained = 0u; drained < 16u && PeekMessageW(&message, nullptr, 0u, 0u, PM_REMOVE) != FALSE; ++drained)
        {
            if (message.message == WM_QUIT)
            {
                PostQuitMessage(static_cast<int>(message.wParam));
                return false;
            }
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
        return true;
    };
    const auto prepareVisibleRedrawBatch = [&](HWND hwnd) noexcept
    { return pumpBatch() && IsWindow(hwnd) != FALSE && IsWindowVisible(hwnd) != FALSE && IsIconic(hwnd) == FALSE; };
    const auto visibleNonIconic = [](HWND hwnd) noexcept { return IsWindow(hwnd) != FALSE && IsWindowVisible(hwnd) != FALSE && IsIconic(hwnd) == FALSE; };
    const auto pumpHidden       = [&]() noexcept -> bool
    {
        const ULONGLONG started = GetTickCount64();
        while (GetTickCount64() - started < 350u)
        {
            MsgWaitForMultipleObjectsEx(0u, nullptr, 20u, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
            if (! pumpBatch())
                return false;
        }
        return true;
    };
    // This fixture owns its synthetic viewport. Derive the nonclient extent from
    // the actual host rather than assuming its title/frame matches DefWindowProc.
    // All bounded settlement is outside measurement; exact client/frame guards remain.
    const auto sizeWindow = [](HWND hwnd, int widthDip = 760, int heightDip = 480) noexcept
    {
        for (unsigned int attempt = 0u; attempt < 3u; ++attempt)
        {
            const UINT dpi = GetDpiForWindow(hwnd);
            RECT outer{}, client{};
            if (dpi == 0u || ! GetWindowRect(hwnd, &outer) || ! GetClientRect(hwnd, &client))
                return false;
            const int width  = MulDiv(widthDip, static_cast<int>(dpi), 96);
            const int height = MulDiv(heightDip, static_cast<int>(dpi), 96);
            if (client.right - client.left == width && client.bottom - client.top == height)
                return true;
            const int outerWidth  = outer.right - outer.left + width - (client.right - client.left);
            const int outerHeight = outer.bottom - outer.top + height - (client.bottom - client.top);
            if (outerWidth <= 0 || outerHeight <= 0 || ! SetWindowPos(hwnd, nullptr, 0, 0, outerWidth, outerHeight, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE))
                return false;
        }
        RECT client{};
        const UINT dpi = GetDpiForWindow(hwnd);
        return dpi != 0u && GetClientRect(hwnd, &client) && client.right - client.left == MulDiv(widthDip, static_cast<int>(dpi), 96) &&
               client.bottom - client.top == MulDiv(heightDip, static_cast<int>(dpi), 96);
    };
    auto lifetime = std::make_shared<int>(0);
    {
        wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
        if (! popup)
            return reject(L"popup creation failed");
        auto* popupState = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA));
        if (! popupState)
            return reject(L"popup state unavailable");
        popupState->debugFixedGeometry = true;
        popupState->DebugSetVisualScenario(tasks, false, true);
        if (! sizeWindow(popup.get()))
            return reject(L"popup geometry unavailable");
        ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
        // Native showing can finish a mixed-DPI creation transition. Size the
        // declared viewport again after that queue work, before warm-up/timing.
        if (! prepareVisibleRedrawBatch(popup.get()) || ! sizeWindow(popup.get()))
            return reject(L"popup fixed geometry after native show unavailable");
        if (GetClientRect(popup.get(), &popupClient) == FALSE)
            return reject(L"popup client unavailable");
        popupDpi          = GetDpiForWindow(popup.get());
        const auto redraw = [&]() noexcept
        {
            for (unsigned int n = 0u; n < 8u; ++n)
                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        };
        redraw();
        if (GetClientRect(popup.get(), &popupClient) == FALSE ||
            popupClient.right - popupClient.left != MulDiv(760, static_cast<int>(GetDpiForWindow(popup.get())), 96) ||
            popupClient.bottom - popupClient.top != MulDiv(480, static_cast<int>(GetDpiForWindow(popup.get())), 96))
        {
            AppendLog(std::format(L"Integrated v2 fixed client mismatch: dpi={} observed={}x{} expected={}x{}",
                                  GetDpiForWindow(popup.get()),
                                  popupClient.right - popupClient.left,
                                  popupClient.bottom - popupClient.top,
                                  MulDiv(760, static_cast<int>(GetDpiForWindow(popup.get())), 96),
                                  MulDiv(480, static_cast<int>(GetDpiForWindow(popup.get())), 96)));
            return reject(L"settled popup client unavailable");
        }
        hashWord(popupDpi);
        hashWord(static_cast<uint64_t>(popupClient.right - popupClient.left));
        hashWord(static_cast<uint64_t>(popupClient.bottom - popupClient.top));
        PopupLayoutDebugSnapshot open{};
        open.taskId = tasks.front().taskId;
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), open) || ! open.taskThroughputExpanded || ! open.taskThroughputDisclosureVisible ||
            ! open.taskCancelVisible || ! open.taskCancelHitTargetActive || open.hostedGraphControlCount == 0u)
            return reject(L"first task lacks open graph or Cancel action");
        const uint64_t openGraphCount = open.hostedGraphControlCount;
        hashWord(openGraphCount);
        if (! sample(L"popup-open-before") || ! sampleSurface(L"popup-open-before", open.surfaceResources) ||
            ! matchesClient(open.surfaceResources, popupClient))
            return reject(L"open sample unavailable");
        for (size_t i = 0u; i < graphOpen8Us.size(); ++i)
        {
            if (! prepareVisibleRedrawBatch(popup.get()))
                return reject(L"popup was not visible/noniconic before open redraw batch");
            const uint64_t beforeFrames = popupState->DebugHostedRenderCount();
            const auto start            = Clock::now();
            redraw();
            graphOpen8Us[openCount++]  = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(Clock::now() - start).count());
            const uint64_t afterFrames = popupState->DebugHostedRenderCount();
            if (afterFrames < beforeFrames || afterFrames - beforeFrames != 8u || ! visibleNonIconic(popup.get()))
                return reject(L"open redraw batch did not render exactly eight frames while visible");
        }
        PopupLayoutDebugSnapshot openAfter{};
        if (! sample(L"popup-open-after") || ! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), openAfter) ||
            ! sampleSurface(L"popup-open-after", openAfter.surfaceResources) || ! sameSurface(open.surfaceResources, openAfter.surfaceResources) ||
            openAfter.surfaceResources.renderCount <= open.surfaceResources.renderCount)
            return reject(L"post-open target/cache growth, missing redraw or unavailable sample");
        if (! DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleThroughput, tasks.front().taskId, 0u}))
            return reject(L"throughput disclosure did not collapse");
        redraw();
        PopupLayoutDebugSnapshot collapsed{};
        collapsed.taskId = tasks.front().taskId;
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), collapsed) || collapsed.taskThroughputExpanded ||
            collapsed.taskRateRect.right <= collapsed.taskRateRect.left || collapsed.taskEtaRect.right <= collapsed.taskEtaRect.left ||
            ! collapsed.taskCancelVisible || ! collapsed.taskCancelHitTargetActive || collapsed.hostedGraphControlCount >= openGraphCount ||
            ! operations->DebugFlushPendingSettingsSaveForSelfTest(10000u))
            return reject(L"collapsed graph, Cancel action or settings save did not settle");
        hashWord(collapsed.hostedGraphControlCount);
        if (! sample(L"popup-collapsed-before") || ! sampleSurface(L"popup-collapsed-before", collapsed.surfaceResources))
            return reject(L"collapsed sample unavailable");
        for (size_t i = 0u; i < graphCollapsed8Us.size(); ++i)
        {
            if (! prepareVisibleRedrawBatch(popup.get()))
                return reject(L"popup was not visible/noniconic before collapsed redraw batch");
            const uint64_t beforeFrames = popupState->DebugHostedRenderCount();
            const auto start            = Clock::now();
            redraw();
            graphCollapsed8Us[collapsedCount++] = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(Clock::now() - start).count());
            const uint64_t afterFrames          = popupState->DebugHostedRenderCount();
            if (afterFrames < beforeFrames || afterFrames - beforeFrames != 8u || ! visibleNonIconic(popup.get()))
                return reject(L"collapsed redraw batch did not render exactly eight frames while visible");
        }
        PopupLayoutDebugSnapshot collapsedAfter{};
        if (! sample(L"popup-collapsed-after") || ! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), collapsedAfter) ||
            ! sampleSurface(L"popup-collapsed-after", collapsedAfter.surfaceResources) ||
            ! sameSurface(collapsed.surfaceResources, collapsedAfter.surfaceResources) ||
            collapsedAfter.surfaceResources.renderCount <= collapsed.surfaceResources.renderCount)
            return reject(L"post-collapsed target/cache growth, missing redraw or unavailable sample");
        if (! DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleThroughput, tasks.front().taskId, 0u}))
            return reject(L"throughput disclosure did not reopen");
        redraw();
        PopupLayoutDebugSnapshot reopened{};
        reopened.taskId = tasks.front().taskId;
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), reopened) || ! reopened.taskThroughputExpanded || ! reopened.taskCancelVisible ||
            ! reopened.taskCancelHitTargetActive || reopened.hostedGraphControlCount != openGraphCount ||
            ! operations->DebugFlushPendingSettingsSaveForSelfTest(10000u))
            return reject(L"reopened graph or settings save did not settle");
        for (size_t batch = 0u; batch < churnUpdate8Us.size(); ++batch)
        {
            if (! prepareVisibleRedrawBatch(popup.get()))
                return reject(L"popup was not visible/noniconic before churn redraw batch");
            const uint64_t beforeFrames = popupState->DebugHostedRenderCount();
            const auto start            = Clock::now();
            for (size_t step = 0u; step < 8u; ++step)
            {
                for (auto& task : tasks)
                    task.completedBytes += 4ull * 1024ull * 1024ull;
                popupState->DebugSetVisualScenario(tasks, false, true);
                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            }
            churnUpdate8Us[churnCount++] = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(Clock::now() - start).count());
            const uint64_t afterFrames   = popupState->DebugHostedRenderCount();
            if (afterFrames < beforeFrames || afterFrames - beforeFrames != 8u || ! visibleNonIconic(popup.get()))
                return reject(L"churn redraw batch did not render exactly eight frames while visible");
        }
        if (! sample(L"popup-after-churn"))
            return reject(L"churn sample unavailable");
        PopupLayoutDebugSnapshot afterChurn{};
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), afterChurn) || ! sampleSurface(L"popup-after-churn", afterChurn.surfaceResources) ||
            ! sameSurface(reopened.surfaceResources, afterChurn.surfaceResources))
            return reject(L"churn target/cache grew or snapshot unavailable");
        ShowWindow(popup.get(), SW_HIDE);
        const uint64_t hiddenFrames = popupState->DebugHostedRenderCount();
        PopupLayoutDebugSnapshot hiddenStart{};
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), hiddenStart) || ! sampleSurface(L"popup-hidden-start", hiddenStart.surfaceResources))
            return reject(L"popup hidden surface unavailable");
        if (! sample(L"popup-hidden-start"))
            return reject(L"popup hidden-start sample unavailable");
        if (! pumpHidden())
            return reject(L"WM_QUIT interrupted hidden popup interval");
        if (popupState->DebugHostedRenderCount() != hiddenFrames || ! sample(L"popup-hidden-end"))
            return reject(L"hidden popup rendered or lost sample");
        PopupLayoutDebugSnapshot hiddenEnd{};
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), hiddenEnd) || ! sampleSurface(L"popup-hidden-end", hiddenEnd.surfaceResources) ||
            ! sameSurface(hiddenStart.surfaceResources, hiddenEnd.surfaceResources))
            return reject(L"hidden popup target/cache changed");
        ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
        redraw();
        if (popupState->DebugHostedRenderCount() <= hiddenFrames || ! sample(L"popup-reshown"))
            return reject(L"reshown popup failed to render or sample");
        PopupLayoutDebugSnapshot reshown{};
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), reshown) || ! sampleSurface(L"popup-reshown", reshown.surfaceResources))
            return reject(L"reshown popup surface unavailable");
    }
    if (DxUi::DebugGetAttachedWindowHostCount() != initialHosts || DxUi::DebugGetSharedWindowHostAttachmentCountForThread(uiThreadId) != initialAttachments ||
        ! sample(L"popup-destroyed"))
        return reject(L"popup host/attachments did not return");

    constexpr uint64_t issueBase = 0xF340000000000000ull;
    const uint64_t typedId       = issueBase + 127u;
    bool issuesRemoved           = false;
    auto removeIssues            = wil::scope_exit([&]() noexcept
    {
        if (! issuesRemoved)
            for (uint64_t i = 0u; i < 128u; ++i)
                operations->DebugRemoveDiagnosticsForTask(issueBase + i);
    });
    for (uint64_t i = 0u; i < 127u; ++i)
    {
        const std::wstring message     = std::format(L"Un élément portant ce nom existe déjà ({:03}).", i);
        const std::wstring source      = std::format(L"C:\\Photos\\IntegratedPerf\\issue-{:03}.raw", i);
        const std::wstring destination = std::format(L"D:\\Archives\\IntegratedPerf\\issue-{:03}.raw", i);
        hashWord(issueBase + i);
        hashText(message);
        hashText(source);
        hashText(destination);
        Diagnostic warning{};
        warning.taskId          = issueBase + i;
        warning.operation       = FILESYSTEM_COPY;
        warning.localTime       = SYSTEMTIME{2026, 9, 1, 21, 12, 0, 0, 0};
        warning.severity        = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
        warning.status          = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        warning.category        = L"selftest.integratedPresentation";
        warning.message         = message;
        warning.sourcePath      = source;
        warning.destinationPath = destination;
        operations->EnqueueTaskDiagnostic(std::move(warning));
    }
    // Same typed producer record shape as cmd_pane_fileops_issues_pane_typed_decision_projection.
    Diagnostic problem{};
    problem.taskId          = typedId;
    problem.operation       = FILESYSTEM_COPY;
    problem.localTime       = SYSTEMTIME{2026, 9, 1, 21, 12, 0, 0, 0};
    problem.severity        = FolderWindow::FileOperationState::DiagnosticSeverity::Warning;
    problem.status          = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    problem.category        = L"item.conflict.prompt";
    problem.message         = L"Une photographie portant ce nom existe déjà dans le dossier de destination.";
    problem.sourcePath      = L"C:\\Photographies\\IntegratedPerf\\Photographie originale.raw";
    problem.destinationPath = L"D:\\Sauvegardes\\IntegratedPerf\\Photographie originale.raw";
    problem.conflictPrompt  = FileOperations::ConflictPromptObservation{10u, FileOperations::ConflictClass::RegularFileExists, 0u};
    hashWord(typedId);
    hashWord(10u);
    hashText(problem.category);
    hashText(problem.message);
    hashText(problem.sourcePath);
    hashText(problem.destinationPath);
    operations->EnqueueTaskDiagnostic(problem);
    auto decision                        = problem;
    decision.severity                    = FolderWindow::FileOperationState::DiagnosticSeverity::Info;
    decision.category                    = L"item.conflict.decision";
    decision.acceptedDecision            = FileOperations::AcceptedConflictDecision{};
    auto& accepted                       = decision.acceptedDecision.value();
    accepted.taskId                      = typedId;
    accepted.promptGeneration            = 10u;
    accepted.sourceItemIndex             = 0u;
    accepted.action                      = FileOperations::ConflictAction::Overwrite;
    accepted.scope.bucket                = FileOperations::ConflictClass::RegularFileExists;
    accepted.applyToAll                  = true;
    accepted.originatingPromptGeneration = 10u;
    hashText(decision.category);
    hashWord(static_cast<uint64_t>(accepted.action));
    hashWord(static_cast<uint64_t>(accepted.scope.bucket));
    hashWord(accepted.applyToAll);
    hashWord(accepted.originatingPromptGeneration);
    operations->EnqueueTaskDiagnostic(decision);
    auto result = decision;
    result.acceptedDecision.reset();
    result.conflictPrompt.reset();
    result.category       = L"item.conflict.result";
    result.status         = S_OK;
    result.conflictResult = FileOperations::ConflictResultObservation{{typedId, 10u},
                                                                      FileOperations::PublicationState::Published,
                                                                      FileOperations::VerificationState::NotRequested,
                                                                      FileOperations::SourceDisposition::Retained,
                                                                      FileOperations::ItemCompletion::Completed,
                                                                      FileOperations::OwnedStageDisposition::NotApplicable,
                                                                      S_OK};
    hashText(result.category);
    hashWord(static_cast<uint64_t>(result.status));
    hashWord(static_cast<uint64_t>(result.conflictResult->publication));
    hashWord(static_cast<uint64_t>(result.conflictResult->verification));
    hashWord(static_cast<uint64_t>(result.conflictResult->sourceDisposition));
    hashWord(static_cast<uint64_t>(result.conflictResult->completion));
    hashWord(static_cast<uint64_t>(result.conflictResult->ownedStageDisposition));
    operations->EnqueueTaskDiagnostic(result);
    diagnostics.clear();
    operations->CollectDiagnostics(diagnostics);
    if (diagnostics.size() != 130u)
        return reject(L"typed diagnostic stream is not 127 plus three records");
    std::vector<Diagnostic>{}.swap(diagnostics);
    {
        wil::unique_hwnd pane(FileOperationsIssuesPane::Create(operations, folder, state.mainWindow, lifetime));
        if (! pane || ! sizeWindow(pane.get()))
            return reject(L"Issues pane or geometry unavailable");
        ShowWindow(pane.get(), SW_SHOWNOACTIVATE);
        if (GetClientRect(pane.get(), &issuesClient) == FALSE)
            return reject(L"Issues client unavailable");
        issuesDpi = GetDpiForWindow(pane.get());
        FileOperationsIssuesPane::SelfTestSnapshot selected{};
        bool ready = false;
        for (unsigned int attempt = 0u; attempt < 100u && ! ready; ++attempt)
        {
            if (! pumpBatch())
                return reject(L"WM_QUIT interrupted Issues settle");
            static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane.get(), false));
            ready = FileOperationsIssuesPane::SelfTestSelectTask(pane.get(), typedId) &&
                    FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane.get(), selected, true) && selected.rowCount == 128u &&
                    selected.primarySelectedTaskId == typedId && selected.detailPresent && selected.primaryHasOriginalProblem &&
                    selected.primaryHasAcceptedDecision && selected.primaryHasConflictResult && ! selected.primaryResultIsAmbiguous &&
                    ! selected.detailCopyText.empty() && selected.dxRenderCount != 0u;
            if (! ready)
                MsgWaitForMultipleObjectsEx(0u, nullptr, 20u, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
        }
        if (! ready)
            return reject(L"128 rows and selected typed detail did not settle");
        if (GetClientRect(pane.get(), &issuesClient) == FALSE)
            return reject(L"settled Issues client unavailable");
        hashWord(issuesDpi);
        hashWord(static_cast<uint64_t>(issuesClient.right - issuesClient.left));
        hashWord(static_cast<uint64_t>(issuesClient.bottom - issuesClient.top));
        if (! sample(L"issues-detail-before"))
            return reject(L"Issues pre-detail sample unavailable");
        for (unsigned int n = 0u; n < 5u; ++n)
            RedrawWindow(pane.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane.get(), selected, true) ||
            ! sampleSurface(L"issues-detail-before", selected.surfaceResources) || ! matchesClient(selected.surfaceResources, issuesClient))
            return reject(L"Issues warm target/cache unavailable");
        const auto detailBefore = selected.surfaceResources;
        for (size_t i = 0u; i < issuesDetail8Us.size(); ++i)
        {
            if (! prepareVisibleRedrawBatch(pane.get()) || ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane.get(), selected, true))
                return reject(L"Issues pane was not visible/noniconic before detail redraw batch");
            const uint64_t beforeFrames = selected.surfaceResources.renderCount;
            const auto start            = Clock::now();
            for (unsigned int n = 0u; n < 8u; ++n)
                RedrawWindow(pane.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            issuesDetail8Us[issuesCount++] = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(Clock::now() - start).count());
            if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane.get(), selected, true) || selected.surfaceResources.renderCount < beforeFrames ||
                selected.surfaceResources.renderCount - beforeFrames != 8u || ! visibleNonIconic(pane.get()))
                return reject(L"Issues detail redraw batch did not render exactly eight frames while visible");
        }
        if (! sample(L"issues-detail-after") || ! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane.get(), selected, true) ||
            ! sampleSurface(L"issues-detail-after", selected.surfaceResources) || ! sameSurface(detailBefore, selected.surfaceResources) ||
            selected.surfaceResources.renderCount <= detailBefore.renderCount)
            return reject(L"Issues target/cache growth, missing redraw or unavailable sample");
        ShowWindow(pane.get(), SW_HIDE);
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane.get(), selected, true))
            return reject(L"Issues hidden counter unavailable");
        const uint64_t hiddenFrames = selected.dxRenderCount;
        const auto hiddenSurface    = selected.surfaceResources;
        if (! sample(L"issues-hidden-start") || ! sampleSurface(L"issues-hidden-start", hiddenSurface))
            return reject(L"Issues hidden-start sample unavailable");
        if (! pumpHidden())
            return reject(L"WM_QUIT interrupted hidden Issues interval");
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane.get(), selected, true) || selected.dxRenderCount != hiddenFrames ||
            ! sample(L"issues-hidden-end") || ! sampleSurface(L"issues-hidden-end", selected.surfaceResources) ||
            ! sameSurface(hiddenSurface, selected.surfaceResources))
            return reject(L"hidden Issues pane rendered or lost sample");
        ShowWindow(pane.get(), SW_SHOWNOACTIVATE);
        RedrawWindow(pane.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane.get(), selected, true) || selected.dxRenderCount <= hiddenFrames ||
            selected.primarySelectedTaskId != typedId || ! sample(L"issues-reshown") || ! sampleSurface(L"issues-reshown", selected.surfaceResources))
            return reject(L"Issues pane failed to restore selected detail");
    }
    if (DxUi::DebugGetAttachedWindowHostCount() != initialHosts || DxUi::DebugGetSharedWindowHostAttachmentCountForThread(uiThreadId) != initialAttachments ||
        ! sample(L"issues-destroyed"))
        return reject(L"Issues host/attachments did not return");
    for (uint64_t i = 0u; i < 128u; ++i)
        operations->DebugRemoveDiagnosticsForTask(issueBase + i);
    diagnostics.clear();
    operations->CollectDiagnostics(diagnostics);
    if (! diagnostics.empty())
        return reject(L"fixture diagnostic history did not clear");
    issuesRemoved = true;
    if (! sample(L"issues-cleaned"))
        return reject(L"Issues cleanup sample unavailable");

    // Additional workload, identical on both product revisions. Keep the original
    // eight-running-task digest and all its sample phases unchanged.
    constexpr std::array<const wchar_t*, 4u> scanBeforePhases{
        {L"scan-compare-before", L"scan-change-case-before", L"scan-change-attributes-before", L"scan-file-list-before"}};
    constexpr std::array<const wchar_t*, 4u> scanAfterPhases{
        {L"scan-compare-after", L"scan-change-case-after", L"scan-change-attributes-after", L"scan-file-list-after"}};
    constexpr std::array<const wchar_t*, 4u> scanClosedPhases{
        {L"scan-compare-closed", L"scan-change-case-closed", L"scan-change-attributes-closed", L"scan-file-list-closed"}};
    constexpr std::array<const wchar_t*, 4u> scanTitles{
        {L"Comparer les dossiers", L"Modifier les majuscules et minuscules", L"Modifier les attributs", L"Créer une liste de fichiers"}};
    const std::filesystem::path scanDirectory =
        L"C:\\Photographies\\Archives familiales de plusieurs générations\\Sélection définitive 📷\\re\u0301union familiale de septembre 2026";
    std::array<std::filesystem::path, 40u> scanPaths{};
    for (size_t step = 0u; step < scanPaths.size(); ++step)
        scanPaths[step] = scanDirectory / std::format(L"Photographie originale {:03}.raw", step);
    hashScanText(L"FileOps.IntegratedPresentation.ScanStatus.v2;fr-FR;dark;reduced-motion;fixed-480x720-DIP;exact-eight-frames;warm-path-range");
    hashScanText(scanDirectory.native());
    for (const auto& path : scanPaths)
        hashScanText(path.native());

    const auto sizeScanWindow = [&](HWND hwnd) noexcept { return sizeWindow(hwnd, 480, 720); };
    for (size_t kind = 0u; kind < scanKindNames.size(); ++kind)
    {
        TaskSnapshot scan{};
        scan.kind                           = TaskSnapshot::Kind::Informational;
        scan.taskId                         = 0xF350000000000000ull + kind;
        auto& info                          = scan.informational;
        info.taskId                         = scan.taskId;
        info.kind                           = static_cast<FolderWindow::InformationalTaskUpdate::Kind>(kind);
        info.title                          = scanTitles[kind];
        info.leftRoot                       = L"C:\\Photographies";
        info.rightRoot                      = L"D:\\Sauvegardes";
        info.scanActive                     = kind == 0u;
        info.scanFolderCount                = 12u;
        info.scanEntryCount                 = 128u;
        info.scanCurrentRelative            = scanPaths.front().lexically_relative(info.leftRoot);
        info.changeCaseEnumerating          = kind == 1u;
        info.changeCaseScannedFolders       = 12u;
        info.changeCaseScannedEntries       = 128u;
        info.changeCaseCurrentPath          = scanPaths.front();
        info.changeAttributesEnumerating    = kind == 2u;
        info.changeAttributesScannedFolders = 12u;
        info.changeAttributesScannedEntries = 128u;
        info.changeAttributesCurrentPath    = scanPaths.front();
        info.makeFileListCollecting         = kind == 3u;
        info.makeFileListScannedFolders     = 12u;
        info.makeFileListScannedEntries     = 128u;
        info.makeFileListCurrentPath        = scanPaths.front();
        hashScanWord(kind);
        hashScanWord(scan.taskId);
        hashScanText(info.title);
        hashScanWord(12u);
        hashScanWord(128u);
        for (size_t step = 0u; step < scanPaths.size(); ++step)
        {
            hashScanWord(12u + step);
            hashScanWord(128u + step);
        }

        {
            wil::unique_hwnd scanPopup(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
            if (! scanPopup)
                return reject(L"scan-status popup creation failed");
            auto* scanState = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(scanPopup.get(), GWLP_USERDATA));
            if (! scanState)
                return reject(L"scan-status popup state unavailable");
            scanState->debugFixedGeometry = true;
            scanState->DebugSetVisualScenario(std::vector<TaskSnapshot>{scan}, false, true);
            if (! sizeScanWindow(scanPopup.get()))
                return reject(L"scan-status popup geometry unavailable");
            ShowWindow(scanPopup.get(), SW_SHOWNOACTIVATE);
            if (! prepareVisibleRedrawBatch(scanPopup.get()) || ! sizeScanWindow(scanPopup.get()))
                return reject(L"scan-status fixed geometry after native show unavailable");
            const auto setScanStep = [&](size_t step) noexcept
            {
                const uint64_t folders = 12u + step;
                const uint64_t entries = 128u + step;
                switch (kind)
                {
                    case 0u:
                        info.scanCurrentRelative = scanPaths[step].lexically_relative(info.leftRoot);
                        info.scanFolderCount     = folders;
                        info.scanEntryCount      = entries;
                        break;
                    case 1u:
                        info.changeCaseCurrentPath    = scanPaths[step];
                        info.changeCaseScannedFolders = folders;
                        info.changeCaseScannedEntries = entries;
                        break;
                    case 2u:
                        info.changeAttributesCurrentPath    = scanPaths[step];
                        info.changeAttributesScannedFolders = folders;
                        info.changeAttributesScannedEntries = entries;
                        break;
                    default:
                        info.makeFileListCurrentPath    = scanPaths[step];
                        info.makeFileListScannedFolders = folders;
                        info.makeFileListScannedEntries = entries;
                        break;
                }
            };
            // Warm the complete declared changing-path range before comparing
            // retained caches. First use of a new layout/hover palette belongs
            // to cold preparation, not an unbounded-growth assertion.
            const auto initialInformation = info;
            for (size_t step = 0u; step < scanPaths.size(); ++step)
            {
                if (! prepareVisibleRedrawBatch(scanPopup.get()))
                    return reject(L"scan-status changing-path warm-up interrupted");
                setScanStep(step);
                scanState->DebugSetVisualScenario(std::vector<TaskSnapshot>{scan}, false, true);
                RedrawWindow(scanPopup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            }
            info = initialInformation;
            // The first paint can consume the creation-time throughput geometry
            // preservation flag. Refresh on each warm paint, just as in the
            // changing workload, so the following paint applies natural sizing
            // before the first sample instead of during the measured updates.
            for (unsigned int warm = 0u; warm < 8u; ++warm)
            {
                scanState->DebugSetVisualScenario(std::vector<TaskSnapshot>{scan}, false, true);
                RedrawWindow(scanPopup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            }
            RECT scanClient{};
            if (GetClientRect(scanPopup.get(), &scanClient) == FALSE ||
                scanClient.right - scanClient.left != MulDiv(480, static_cast<int>(GetDpiForWindow(scanPopup.get())), 96) ||
                scanClient.bottom - scanClient.top != MulDiv(720, static_cast<int>(GetDpiForWindow(scanPopup.get())), 96))
                return reject(L"scan-status fixed 480x720-DIP client unavailable");
            const UINT scanDpi = GetDpiForWindow(scanPopup.get());
            hashScanWord(scanDpi);
            hashScanWord(static_cast<uint64_t>(scanClient.right - scanClient.left));
            hashScanWord(static_cast<uint64_t>(scanClient.bottom - scanClient.top));
            Debug::Perf::Emit(L"FileOps.IntegratedPresentation.ScanStatusClientPixels",
                              scanKindNames[kind],
                              scanDpi,
                              static_cast<uint64_t>(scanClient.right - scanClient.left),
                              static_cast<uint64_t>(scanClient.bottom - scanClient.top),
                              S_OK);
            // V41's public task-layout helper only reports FileOperation tasks.
            // Its diagnostic message still fills the host-wide surface observation
            // for informational cards before that task projection. Validate the
            // observed resource fields below rather than interpreting task-found.
            const auto readScanSurface = [&](PopupLayoutDebugSnapshot& snapshot) noexcept
            {
                static_cast<void>(SendMessageW(scanPopup.get(), WndMsg::kFileOpsPopupLayoutSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)));
                return snapshot.surfaceResources.logicalBytesAvailable && snapshot.surfaceResources.renderCount != 0u;
            };
            PopupLayoutDebugSnapshot before{};
            before.taskId = scan.taskId;
            if (! readScanSurface(before) || ! sample(scanBeforePhases[kind]) || ! sampleSurface(scanBeforePhases[kind], before.surfaceResources) ||
                ! matchesClient(before.surfaceResources, scanClient))
                return reject(L"scan-status warm surface or resource sample unavailable");
            const uint64_t beforeFrames = scanState->DebugHostedRenderCount();
            for (size_t round = 0u; round < scanClean8Us[kind].size(); ++round)
            {
                if (! prepareVisibleRedrawBatch(scanPopup.get()))
                    return reject(L"scan-status popup was not visible/noniconic before clean redraw batch");
                const uint64_t batchBeforeFrames = scanState->DebugHostedRenderCount();
                PopupLayoutDebugSnapshot batchBefore{};
                batchBefore.taskId = scan.taskId;
                if (! readScanSurface(batchBefore))
                    return reject(L"scan-status clean redraw pre-counter unavailable");
                const auto started = Clock::now();
                for (unsigned int redraw = 0u; redraw < 8u; ++redraw)
                    RedrawWindow(scanPopup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                scanClean8Us[kind][scanCleanCount[kind]++] =
                    static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(Clock::now() - started).count());
                PopupLayoutDebugSnapshot batchAfter{};
                batchAfter.taskId = scan.taskId;
                if (! readScanSurface(batchAfter) || batchAfter.surfaceResources.renderCount < batchBefore.surfaceResources.renderCount ||
                    batchAfter.surfaceResources.renderCount - batchBefore.surfaceResources.renderCount != 8u ||
                    scanState->DebugHostedRenderCount() < batchBeforeFrames || scanState->DebugHostedRenderCount() - batchBeforeFrames != 8u ||
                    ! visibleNonIconic(scanPopup.get()))
                    return reject(L"scan-status clean redraw batch did not render exactly eight frames while visible");
            }
            for (size_t round = 0u; round < scanChanging8Us[kind].size(); ++round)
            {
                if (! prepareVisibleRedrawBatch(scanPopup.get()))
                    return reject(L"scan-status popup was not visible/noniconic before changing redraw batch");
                const uint64_t batchBeforeFrames = scanState->DebugHostedRenderCount();
                PopupLayoutDebugSnapshot batchBefore{};
                batchBefore.taskId = scan.taskId;
                if (! readScanSurface(batchBefore))
                    return reject(L"scan-status changing redraw pre-counter unavailable");
                const auto started = Clock::now();
                for (size_t redraw = 0u; redraw < 8u; ++redraw)
                {
                    setScanStep(round * 8u + redraw);
                    scanState->DebugSetVisualScenario(std::vector<TaskSnapshot>{scan}, false, true);
                    RedrawWindow(scanPopup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                }
                scanChanging8Us[kind][scanChangingCount[kind]++] =
                    static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(Clock::now() - started).count());
                PopupLayoutDebugSnapshot batchAfter{};
                batchAfter.taskId = scan.taskId;
                if (! readScanSurface(batchAfter) || batchAfter.surfaceResources.renderCount < batchBefore.surfaceResources.renderCount ||
                    batchAfter.surfaceResources.renderCount - batchBefore.surfaceResources.renderCount != 8u ||
                    scanState->DebugHostedRenderCount() < batchBeforeFrames || scanState->DebugHostedRenderCount() - batchBeforeFrames != 8u ||
                    ! visibleNonIconic(scanPopup.get()))
                    return reject(L"scan-status changing redraw batch did not render exactly eight frames while visible");
            }
            PopupLayoutDebugSnapshot after{};
            after.taskId = scan.taskId;
            if (! sample(scanAfterPhases[kind]) || ! readScanSurface(after) || ! sampleSurface(scanAfterPhases[kind], after.surfaceResources) ||
                ! sameSurface(before.surfaceResources, after.surfaceResources) || scanState->DebugHostedRenderCount() <= beforeFrames ||
                after.surfaceResources.renderCount <= before.surfaceResources.renderCount)
                return reject(L"scan-status redraw or bounded surface/cache evidence failed");
        }
        if (DxUi::DebugGetAttachedWindowHostCount() != initialHosts ||
            DxUi::DebugGetSharedWindowHostAttachmentCountForThread(uiThreadId) != initialAttachments || ! sample(scanClosedPhases[kind]))
            return reject(L"scan-status popup host/attachment teardown failed");
    }

    if (! restore())
        return reject(L"settings restoration failed");
    restored = true;
    return true;
}
