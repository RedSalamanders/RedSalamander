// Included only by the opt-in documentation gallery. No operation is admitted.
struct FileOpsGalleryMonitor final
{
    RECT work{};
    std::wstring device;
};

[[nodiscard]] std::vector<FileOpsGalleryMonitor> GetFileOpsGalleryMonitors()
{
    std::vector<FileOpsGalleryMonitor> monitors;
    EnumDisplayMonitors(nullptr,
                        nullptr,
                        [](HMONITOR monitor, HDC, LPRECT, LPARAM context) noexcept -> BOOL
    {
        MONITORINFOEXW info{};
        info.cbSize = sizeof(info);
        if (GetMonitorInfoW(monitor, &info))
            reinterpret_cast<std::vector<FileOpsGalleryMonitor>*>(context)->push_back({info.rcWork, info.szDevice});
        return TRUE;
    },
                        reinterpret_cast<LPARAM>(&monitors));
    return monitors;
}

void PlaceFileOpsGalleryWindow(HWND hwnd, const FileOpsGalleryMonitor& monitor, int widthDip, int heightDip)
{
    SetWindowPos(hwnd, HWND_BOTTOM, monitor.work.left + 16, monitor.work.top + 16, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE);
    const UINT dpi = GetDpiForWindow(hwnd);
    RECT rect{0, 0, MulDiv(widthDip, static_cast<int>(dpi), 96), MulDiv(heightDip, static_cast<int>(dpi), 96)};
    AdjustWindowRectExForDpi(
        &rect, static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE)), FALSE, static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE)), dpi);
    SetWindowPos(hwnd,
                 HWND_BOTTOM,
                 monitor.work.left + 16,
                 monitor.work.top + 16,
                 std::min(rect.right - rect.left, monitor.work.right - monitor.work.left - 32),
                 std::min(rect.bottom - rect.top, monitor.work.bottom - monitor.work.top - 32),
                 SWP_NOACTIVATE);
    ShowWindow(hwnd, SW_SHOWNOACTIVATE);
    RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
}

[[nodiscard]] bool SaveFileOpsGallerySurface(
    HWND hwnd, const std::filesystem::path& output, std::ofstream& manifest, std::wstring_view name, size_t display, std::wstring_view fixture)
{
    RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    const UINT dpi  = GetDpiForWindow(hwnd);
    const auto file = std::format(L"{}__dark__fr-FR__display{}__{}dpi.png", name, display + 1u, dpi);
    if (FAILED(RedSalamander::TestSupport::SaveWindowScreenshot(hwnd, output / file)))
        return false;
    manifest << Common::Strings::Utf8FromUtf16ReplacingInvalid(file) << '\t' << Common::Strings::Utf8FromUtf16ReplacingInvalid(name) << "\tdark\t0\t" << dpi
             << '\t' << Common::Strings::Utf8FromUtf16ReplacingInvalid(fixture) << '\n';
    manifest.flush();
    AppendLog(std::format(L"VisualGallery: {}", file));
    return manifest.good();
}

struct FileOpsGalleryConfirmationCapture final
{
    std::filesystem::path output;
    std::ofstream* manifest = nullptr;
    std::wstring name;
    size_t display    = 0u;
    bool succeeded    = false;
    unsigned attempts = 0u;
};
FileOpsGalleryConfirmationCapture* g_fileOpsGalleryConfirmation = nullptr;

void CALLBACK CaptureFileOpsGalleryConfirmation(HWND, UINT, UINT_PTR timerId, DWORD) noexcept
{
    auto* capture = g_fileOpsGalleryConfirmation;
    if (! capture)
        return;
    HWND overlay = nullptr;
    EnumThreadWindows(GetCurrentThreadId(),
                      [](HWND candidate, LPARAM context) noexcept -> BOOL
    {
        wchar_t name[128]{};
        GetClassNameW(candidate, name, static_cast<int>(std::size(name)));
        if (std::wstring_view(name) == L"RedSalamander.AlertOverlayWindow" && IsWindowVisible(candidate))
            *reinterpret_cast<HWND*>(context) = candidate;
        return TRUE;
    },
                      reinterpret_cast<LPARAM>(&overlay));
    RedSalamander::Ui::AlertOverlayWindowDebugSnapshot snapshot{};
    if (! overlay || ! RedSalamander::Ui::DebugGetAlertOverlayWindowSnapshot(overlay, snapshot) || ! snapshot.hasLayout || snapshot.paintCount == 0u)
    {
        if (++capture->attempts > 100u)
        {
            KillTimer(nullptr, timerId);
            HostBeginPromptShutdown(); // Bounded failure: leave the nested test prompt without accepting it.
        }
        return;
    }
    KillTimer(nullptr, timerId);
    capture->succeeded =
        SaveFileOpsGallerySurface(overlay, capture->output, *capture->manifest, capture->name, capture->display, L"production-confirmation-synthetic-request");
    SendMessageW(overlay, WM_KEYDOWN, VK_ESCAPE, 0);
}

[[nodiscard]] bool CaptureFileOpsGallerySupplemental(SelfTestState& state,
                                                     const std::filesystem::path& output,
                                                     std::ofstream& manifest,
                                                     const std::vector<FileOpsGalleryMonitor>& monitors)
{
    auto* folder                  = TryGetFolderWindow(state.mainWindow);
    auto* operations              = TryGetFileOps(folder);
    auto lifetime                 = std::make_shared<int>(0);
    const bool previousAutoAccept = HostGetAutoAcceptPrompts();
    const auto previousOverride   = HostGetTestPromptResultOverride();
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();
    auto restorePrompts = wil::scope_exit([&]
    {
        HostSetAutoAcceptPrompts(previousAutoAccept);
        HostSetTestPromptResultOverride(previousOverride);
        HostResetPromptShutdown();
    });
    for (size_t display = 0u; display < monitors.size(); ++display)
    {
        const auto& monitor = monitors[display];
        wil::unique_hwnd anchor(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
        if (! anchor)
            return false;
        auto* renderer = reinterpret_cast<FileOperationsPopupInternal::FileOperationsPopupState*>(GetWindowLongPtrW(anchor.get(), GWLP_USERDATA));
        renderer->DebugSetVisualScenario({}, false, true);
        PlaceFileOpsGalleryWindow(anchor.get(), monitor, 920, 620);
        for (uint32_t variant = 0u; variant < 10u; ++variant)
        {
            HostFileOperationPromptOptions options{};
            options.sizeBytes                    = sizeof(options);
            options.verifyAfterCopy              = variant % 2u;
            options.verificationAvailability     = static_cast<HostFileOperationVerificationAvailability>(variant % 4u);
            options.clipboardMoveConsumesCutList = variant >= 4u ? 1u : 0u;
            options.linkPolicy                   = variant % 2u ? HOST_FILE_OPERATION_LINK_SKIP : HOST_FILE_OPERATION_LINK_PRESERVE;
            options.executionMode                = variant % 2u ? HOST_FILE_OPERATION_EXECUTION_PARALLEL : HOST_FILE_OPERATION_EXECUTION_QUEUE;
            options.bandwidthLimitBytesPerSecond = variant % 2u ? 64ull * 1024ull * 1024ull : 0u;
            HostPromptRequest request{};
            request.sizeBytes     = sizeof(request);
            request.scope         = HOST_ALERT_SCOPE_WINDOW;
            request.targetWindow  = anchor.get();
            request.severity      = HOST_ALERT_WARNING;
            request.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
            request.defaultResult = HOST_PROMPT_RESULT_CANCEL;
            request.presentation  = variant < 4u   ? HOST_PROMPT_PRESENTATION_COPY
                                    : variant < 8u ? HOST_PROMPT_PRESENTATION_MOVE
                                                   : HOST_PROMPT_PRESENTATION_DELETE;
            request.message =
                L"128 fichiers sélectionnés dans C:\\Photographies\\Vacances en famille — septembre 2026\\Sélection définitive pour impression et archivage.\n"
                L"Destination : D:\\Sauvegardes\\Archives photographiques personnelles\\Exposition annuelle\\Photographies retouchées en très haute "
                L"résolution.\n"
                L"Veuillez vérifier les emplacements et les options avant de poursuivre cette opération.";
            request.fileOperationOptions = variant < 8u ? &options : nullptr;
            FileOpsGalleryConfirmationCapture capture{output, &manifest, std::format(L"102-confirmation-{}", variant), display};
            g_fileOpsGalleryConfirmation = &capture;
            const UINT_PTR timer         = SetTimer(nullptr, 0u, 100u, CaptureFileOpsGalleryConfirmation);
            if (! timer)
                return false;
            auto reset              = wil::scope_exit([&]
            {
                KillTimer(nullptr, timer);
                g_fileOpsGalleryConfirmation = nullptr;
            });
            HostPromptResult result = HOST_PROMPT_RESULT_NONE;
            const HRESULT hr        = HostShowPrompt(request, nullptr, &result);
            if (FAILED(hr) || ! capture.succeeded || result != HOST_PROMPT_RESULT_CANCEL)
            {
                Fail(std::format(L"Confirmation capture failed: variant {}, hr 0x{:08X}, result {}",
                                 variant,
                                 static_cast<unsigned long>(hr),
                                 static_cast<unsigned>(result)));
                return false;
            }
        }
        wil::unique_hwnd issues(FileOperationsIssuesPane::Create(operations, folder, state.mainWindow, lifetime));
        if (! issues)
            return false;
        PlaceFileOpsGalleryWindow(issues.get(), monitor, 1000, 620);
        FileOperationsIssuesPane::SelfTestRefresh(issues.get(), true);
        if (! SaveFileOpsGallerySurface(issues.get(), output, manifest, L"103-issues-empty", display, L"production-issues-pane"))
            return false;
        auto dismiss = wil::scope_exit([&]
        {
            operations->DismissCompletedTask(90001u);
            operations->DismissCompletedTask(90002u);
        });
        for (uint64_t id : {90001ull, 90002ull})
        {
            FolderWindow::FileOperationState::CompletedTaskSummary summary{};
            summary.taskId         = id;
            summary.operation      = id == 90001u ? FILESYSTEM_COPY : FILESYSTEM_MOVE;
            summary.resultHr       = E_ACCESSDENIED;
            summary.errorCount     = 14u;
            summary.totalItems     = 128u;
            summary.completedItems = 114u;
            summary.completedTick  = GetTickCount64();
            for (unsigned row = 0u; row < 14u; ++row)
            {
                FolderWindow::FileOperationState::TaskDiagnosticEntry issue{};
                GetLocalTime(&issue.localTime);
                issue.taskId    = id;
                issue.operation = summary.operation;
                issue.severity  = FolderWindow::FileOperationState::DiagnosticSeverity::Error;
                issue.status    = E_ACCESSDENIED;
                issue.category  = L"Vérification et conservation de la source";
                issue.message   = L"Impossible de terminer la vérification : le fichier source a été conservé. Vérifiez les autorisations du dossier de "
                                  L"destination avant de réessayer.";
                issue.sourcePath =
                    std::format(L"C:\\Photographies\\Vacances en famille — septembre 2026\\Photographie panoramique retouchée numéro {:02}.raw", row);
                issue.destinationPath = L"D:\\Sauvegardes\\Archives photographiques personnelles\\Exposition annuelle\\Sélection définitive";
                summary.issueDiagnostics.push_back(std::move(issue));
            }
            operations->DebugAppendCompletedTaskForSelfTest(std::move(summary));
        }
        FileOperationsIssuesPane::SelfTestRefresh(issues.get(), true);
        FileOperationsIssuesPane::SelfTestSnapshot issueSnapshot{};
        if (! FileOperationsIssuesPane::TryGetSelfTestSnapshot(issues.get(), issueSnapshot) || issueSnapshot.rowCount != 28u)
            return false;
        if (! SaveFileOpsGallerySurface(issues.get(), output, manifest, L"104-issues-populated", display, L"production-issues-synthetic-diagnostics"))
            return false;
        FileOperationsIssuesPane::SelfTestSelectTask(issues.get(), 90001u);
        if (! SaveFileOpsGallerySurface(issues.get(), output, manifest, L"105-issues-selected", display, L"production-issues-selection"))
            return false;
        FileOperationsIssuesPane::SelfTestScrollByWheelDetents(issues.get(), -8);
        if (! SaveFileOpsGallerySurface(issues.get(), output, manifest, L"106-issues-scrolled", display, L"production-issues-scroll"))
            return false;
    }
    return true;
}
