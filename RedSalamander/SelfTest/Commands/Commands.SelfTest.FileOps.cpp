// Commands.SelfTest.FileOps.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// FileOps test family: 19 test functions.

[[nodiscard]] auto ResetFileOperationsIssuesPaneViewStateForTest(FolderWindow::FileOperationState* fileOps) noexcept
{
    std::wstring savedSortColumnId;
    bool savedSortDescending = false;
    std::vector<Common::Settings::GridColumnLayoutEntry> savedGridLayout;
    const bool hadSavedViewState = fileOps && fileOps->TryGetIssuesPaneViewState(savedSortColumnId, savedSortDescending, savedGridLayout);
    if (fileOps)
    {
        fileOps->DebugResetIssuesPaneForSelfTest();
    }

    return wil::scope_exit([fileOps,
                            hadSavedViewState,
                            savedSortColumnId = std::move(savedSortColumnId),
                            savedSortDescending,
                            savedGridLayout = std::move(savedGridLayout)]() mutable noexcept
    {
        if (! fileOps)
        {
            return;
        }

        if (hadSavedViewState)
        {
            fileOps->SaveIssuesPaneViewState(savedSortColumnId, savedSortDescending, savedGridLayout);
            return;
        }

        fileOps->SaveIssuesPaneViewState(L"", false, {});
    });
}

[[nodiscard]] bool TestFileOperationsIssuesPaneUsesDxUiHostWithNoVisibleChildControls(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane DX host baseline validation.");
    if (! fileOps)
    {
        return false;
    }

    const AppTheme previousTheme = g_folderWindow.GetTheme();
    const auto restoreTheme     = wil::scope_exit([&]() noexcept { g_folderWindow.ApplyTheme(previousTheme); });
    const AppTheme backdropTheme =
        MakeWindowBackdropSelfTestTheme(Common::Settings::WindowBackdropMode::Acrylic, L"fileops-issues-backdrop-selftest");
    g_folderWindow.ApplyTheme(backdropTheme);
    const Common::WindowBackdrop::Kind expectedToolBackdropKind =
        Common::WindowBackdrop::Resolve(Common::Settings::WindowBackdropMode::Acrylic, Common::WindowBackdrop::Target::Tool, false);

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File-operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });

    constexpr size_t kTaskCount   = 2u;
    const uint64_t taskBase       = 0xF250000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        for (size_t index = 0; index < kTaskCount; ++index)
        {
            fileOps->DebugRemoveDiagnosticsForTask(taskBase + index);
        }

        if (const HWND pane = getVisibleIssuesPane(); pane && IsWindow(pane) != FALSE)
        {
            static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
        }
    });

    const std::array<std::wstring, kTaskCount> labels{
        L"issues-baseline-0",
        L"issues-baseline-1",
    };
    for (size_t index = 0; index < kTaskCount; ++index)
    {
        const uint64_t taskId = taskBase + index;
        fileOps->RecordTaskDiagnostic(taskId,
                                      (index % 2u) == 0u ? FILESYSTEM_COPY : FILESYSTEM_MOVE,
                                      (index % 2u) == 0u ? FolderWindow::FileOperationState::DiagnosticSeverity::Warning
                                                         : FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                      HRESULT_FROM_WIN32((index % 2u) == 0u ? ERROR_ACCESS_DENIED : ERROR_DISK_FULL),
                                      L"selftest.dxui.issues.baseline",
                                      labels[index],
                                      std::format(L"C:\\selftest-issues-baseline-source-{:02}.txt", index),
                                      std::format(L"D:\\selftest-issues-baseline-dest-{:02}.txt", index));
    }

    const auto validateIssuesPaneHostSurface = [&](std::wstring_view context) noexcept
    {
        state.Require(setIssuesVisible(true, context), std::format(L"Failed to show the file-operations issues pane during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
        state.Require(pane != nullptr && IsWindow(pane) != FALSE, std::format(L"File operations issues pane did not open during {}.", context));
        if (! pane || IsWindow(pane) == FALSE)
        {
            return false;
        }

        state.Require(WaitForAppliedBackdropKind(pane, expectedToolBackdropKind, L"File Operations issues pane", state),
                      std::format(L"File Operations issues pane did not apply the selected Acrylic tool-window backdrop during {}.", context));
        state.Require(WindowExposesUiaProvider(pane), std::format(L"File-operations issues pane should answer WM_GETOBJECT during {}.", context));
        const size_t visibleChildWindowCount = CountVisibleChildWindows(pane);
        state.Require(visibleChildWindowCount == 0u,
                      std::format(L"Issues pane should not expose visible child-control fallback during {}; got {} visible child window(s).",
                                  context,
                                  visibleChildWindowCount));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), std::format(L"Failed to request an issues-pane refresh during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto waitForSnapshot = [&](auto&& predicate, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
        };

        FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
        state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        { return value.rowCount >= kTaskCount && value.visibleWork.visibleRowCount > 0u; },
                                      snapshot),
                      std::format(L"File-operations issues pane did not repopulate visible DX rows during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.visibleWork.visibleColumnCount > 0u,
                      std::format(L"File-operations issues pane should expose visible DX columns during {}.", context));
        state.Require(snapshot.dxResizeFailureCount == 0u,
                      std::format(L"File-operations issues pane should not hit DX resize failures during {}; saw {}.", context, snapshot.dxResizeFailureCount));

        const uint64_t expectedTaskId = taskBase;
        state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, expectedTaskId),
                      std::format(L"Failed to select issues-pane task {} during {}.", expectedTaskId, context));
        state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        { return value.selectionCount == 1u && value.primarySelectedTaskId == expectedTaskId; },
                                      snapshot),
                      std::format(L"File-operations issues pane did not expose the selected task {} during {}.", expectedTaskId, context));
        const std::wstring expectedTaskIdText = std::to_wstring(static_cast<unsigned long long>(expectedTaskId));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(pane);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the issues pane during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"File-operations issues pane should expose visible UI Automation descendants during {}.", context));
        }

        const auto selectionState = CollectVisibleDescendantSelectionPatternState(pane, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(), std::format(L"Failed to collect UI Automation selection state for the issues pane during {}.", context));
        if (selectionState.has_value())
        {
            state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                          std::format(L"File-operations issues pane should expose a visible UI Automation DataGrid during {}.", context));
            state.Require(selectionState->hasSelectionPattern, std::format(L"File-operations issues pane should expose SelectionPattern during {}.", context));
            state.Require(selectionState->selectionCount == 1u,
                          std::format(L"File-operations issues pane should expose exactly one selected UIA row during {}; saw {}.",
                                      context,
                                      selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                          std::format(L"File-operations issues selected UIA row should expose the DataItem control type during {}.", context));
            state.Require(selectionState->selectedHasSelectionItemPattern,
                          std::format(L"File-operations issues selected UIA row should expose SelectionItemPattern during {}.", context));
            state.Require(! selectionState->selectedName.empty(),
                          std::format(L"File-operations issues selected UIA row should expose a non-empty accessible name during {}.", context));
            state.Require(selectionState->selectedName.find(expectedTaskIdText) != std::wstring::npos,
                          std::format(L"File-operations issues selected UIA row name '{}' should include the selected task id '{}' during {}.",
                                      selectionState->selectedName,
                                      expectedTaskIdText,
                                      context));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(setIssuesVisible(false, context), std::format(L"Failed to hide the file-operations issues pane during {}.", context));
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline && getVisibleIssuesPane() != nullptr)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);
        }

        state.Require(getVisibleIssuesPane() == nullptr, std::format(L"File operations issues pane did not close during {}.", context));
        return state.failure.empty();
    };

    if (! validateIssuesPaneHostSurface(L"the initial issues-pane DX host baseline probe"))
    {
        return false;
    }

    if (! validateIssuesPaneHostSurface(L"the reopened issues-pane DX host baseline probe"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneDiffRefreshPreservesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane diff-refresh validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible        = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto restoreVisibility = wil::scope_exit([&] noexcept
    {
        const bool visibleNow = g_folderWindow.IsFileOperationsIssuesPaneVisible();
        if (visibleNow != wasVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline && g_folderWindow.IsFileOperationsIssuesPaneVisible() != wasVisible)
            {
                PumpPendingMessages();
                std::this_thread::sleep_for(20ms);
            }
        }
    });

    if (! wasVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for diff-refresh validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const uint64_t baseTaskId     = 0xF000000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const uint64_t taskA          = baseTaskId + 1u;
    const uint64_t taskB          = baseTaskId + 2u;
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskA);
        fileOps->DebugRemoveDiagnosticsForTask(taskB);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot initialSnapshot{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, initialSnapshot), L"Failed to capture initial issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t initialGeneration = initialSnapshot.refreshGeneration;

    fileOps->RecordTaskDiagnostic(taskA,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                                  L"selftest.dxui.issues",
                                  std::format(L"selection-preserve-a-{}", taskA),
                                  L"C:\\selftest-a.txt",
                                  L"D:\\selftest-a.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, false), L"Failed to request first issues-pane refresh.");

    FileOperationsIssuesPane::SelfTestSnapshot afterA{};
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& snapshot) noexcept
    { return snapshot.refreshGeneration > initialGeneration; },
                                  SelfTest::Scale(3000ms),
                                  afterA),
                  L"Issues pane did not observe the first diagnostic-driven refresh.");
    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskA), L"Failed to select the first injected issues-pane task.");

    FileOperationsIssuesPane::SelfTestSnapshot selectedA{};
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& snapshot) noexcept
    { return snapshot.selectionCount == 1u && snapshot.primarySelectedTaskId == taskA; },
                                  SelfTest::Scale(3000ms),
                                  selectedA),
                  L"Issues pane did not report the expected selection after selecting the first injected task.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, false), L"Failed to request no-op issues-pane refresh.");

    FileOperationsIssuesPane::SelfTestSnapshot afterNoOp{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, afterNoOp), L"Failed to capture issues-pane snapshot after no-op refresh.");
    state.Require(afterNoOp.refreshGeneration == selectedA.refreshGeneration, L"No-op issues-pane refresh should not advance the applied refresh generation.");
    state.Require(afterNoOp.selectionCount == 1u && afterNoOp.primarySelectedTaskId == taskA, L"No-op issues-pane refresh should preserve the selected task.");

    fileOps->RecordTaskDiagnostic(taskB,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                  HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION),
                                  L"selftest.dxui.issues",
                                  std::format(L"selection-preserve-b-{}", taskB),
                                  L"C:\\selftest-b.txt",
                                  L"D:\\selftest-b.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, false), L"Failed to request second issues-pane refresh.");

    FileOperationsIssuesPane::SelfTestSnapshot afterB{};
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& snapshot) noexcept
    { return snapshot.refreshGeneration > afterNoOp.refreshGeneration && snapshot.selectionCount == 1u && snapshot.primarySelectedTaskId == taskA; },
                                  SelfTest::Scale(3000ms),
                                  afterB),
                  L"Changed issues-pane refresh did not preserve the first selected task.");
    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskB), L"Failed to select the second injected issues-pane task.");

    FileOperationsIssuesPane::SelfTestSnapshot selectedB{};
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& snapshot) noexcept
    { return snapshot.selectionCount == 1u && snapshot.primarySelectedTaskId == taskB; },
                                  SelfTest::Scale(3000ms),
                                  selectedB),
                  L"Issues pane did not expose the second injected task after diff refresh.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneHeaderClickSortsResults(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane header-sort validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane header-sort baseline reset"),
                  L"Failed to hide the file-operations issues pane before header-sort validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane header-sort validation"), L"Failed to show the file-operations issues pane.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for header-sort validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const uint64_t baseTaskId     = 0xF100000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const uint64_t taskHigh       = baseTaskId + 9u;
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskHigh);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration = baseline.refreshGeneration;
    fileOps->RecordTaskDiagnostic(taskHigh,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_DISK_FULL),
                                  L"selftest.dxui.issues.sort",
                                  L"zzzz-selftest-issues-sort-high",
                                  L"C:\\selftest-issues-sort-high.txt",
                                  L"D:\\selftest-issues-sort-high.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for header-sort validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows and the task header for header-sort validation.");
    state.Require(snapshot.visibleWork.visibleColumnCount > 0u, L"Issues pane should expose visible DX columns for header-sort validation.");
    state.Require(snapshot.dxResizeFailureCount == 0u,
                  std::format(L"Issues pane should not hit DX resize failures during header-sort validation; saw {}.", snapshot.dxResizeFailureCount));

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskHigh),
                  std::format(L"Failed to select task {} before issues-pane header-sort validation.", taskHigh));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before header-sort validation.", taskHigh));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto taskHeaderClickPoint = [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept -> LPARAM
    {
        return DipPointToClientLParam(
            pane, (value.taskHeaderRect.left + value.taskHeaderRect.right) * 0.5f, (value.taskHeaderRect.top + value.taskHeaderRect.bottom) * 0.5f);
    };

    const auto hitTestTaskHeaderCenter = [&](const FileOperationsIssuesPane::SelfTestSnapshot& value,
                                             FileOperationsIssuesPane::SelfTestGridHit& hit) noexcept -> bool
    {
        return FileOperationsIssuesPane::SelfTestHitTestGridPoint(pane,
                                                                  (value.taskHeaderRect.left + value.taskHeaderRect.right) * 0.5f,
                                                                  (value.taskHeaderRect.top + value.taskHeaderRect.bottom) * 0.5f,
                                                                  hit) &&
               static_cast<unsigned>(hit.zone) == 1u && hit.columnIndex == 1u && ! hit.isHeaderResize;
    };

    const LPARAM clickPoint = taskHeaderClickPoint(snapshot);
    FileOperationsIssuesPane::SelfTestGridHit preClickHit{};
    state.Require(hitTestTaskHeaderCenter(snapshot, preClickHit), L"Failed to hit-test the issues-pane Task header before header-sort validation.");
    state.Require(static_cast<unsigned>(preClickHit.zone) == 1u && preClickHit.columnIndex == 1u && ! preClickHit.isHeaderResize,
                  std::format(L"Issues-pane Task header center did not resolve to a sortable header hit before header-sort validation. "
                              L"zone={}, column={}, resize={}, rect=({}, {}, {}, {})",
                              static_cast<unsigned>(preClickHit.zone),
                              preClickHit.columnIndex,
                              preClickHit.isHeaderResize,
                              preClickHit.rectDip.left,
                              preClickHit.rectDip.top,
                              preClickHit.rectDip.right,
                              preClickHit.rectDip.bottom));
    if (! state.failure.empty())
    {
        return false;
    }
    SendMouseClickToResolvedPointWindow(pane, clickPoint);

    FileOperationsIssuesPane::SelfTestSnapshot immediateAfterClick{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, immediateAfterClick),
                  L"Failed to capture issues-pane snapshot immediately after the first header click.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsIssuesPane::SelfTestSnapshot ascending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.hasActiveSort && value.sortColumnIndex == 1u && ! value.sortDescending && value.selectionCount == 1u &&
               value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount > 0u &&
               value.taskHeaderRect.right > value.taskHeaderRect.left && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      ascending),
                  std::format(L"Issues pane did not switch to ascending task sort after the first header click. "
                              L"immediateSortActive={}, immediateColumn={}, immediateDescending={}, immediateRowCount={}, immediateFirstTask={}, "
                              L"immediateSelectedTask={}, sortActive={}, column={}, descending={}, firstTask={}, selectedTask={}",
                              immediateAfterClick.hasActiveSort,
                              immediateAfterClick.sortColumnIndex,
                              immediateAfterClick.sortDescending,
                              immediateAfterClick.rowCount,
                              immediateAfterClick.firstVisibleTaskId,
                              immediateAfterClick.primarySelectedTaskId,
                              ascending.hasActiveSort,
                              ascending.sortColumnIndex,
                              ascending.sortDescending,
                              ascending.firstVisibleTaskId,
                              ascending.primarySelectedTaskId));

    FileOperationsIssuesPane::SelfTestGridHit secondPreClickHit{};
    state.Require(hitTestTaskHeaderCenter(ascending, secondPreClickHit), L"Failed to hit-test the issues-pane Task header before the second sort click.");
    if (! state.failure.empty())
    {
        return false;
    }
    SendMouseClickToResolvedPointWindow(pane, taskHeaderClickPoint(ascending));

    FileOperationsIssuesPane::SelfTestSnapshot immediateAfterSecondClick{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, immediateAfterSecondClick),
                  L"Failed to capture issues-pane snapshot immediately after the second header click.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsIssuesPane::SelfTestSnapshot descending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.hasActiveSort && value.sortColumnIndex == 1u && value.sortDescending && value.firstVisibleTaskId == taskHigh &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      descending),
                  std::format(L"Issues pane did not switch to descending task sort after the second header click. "
                              L"immediateSortActive={}, immediateColumn={}, immediateDescending={}, immediateFirstTask={}, immediateSelectedTask={}, "
                              L"sortActive={}, column={}, descending={}, firstTask={}, selectedTask={}",
                              immediateAfterSecondClick.hasActiveSort,
                              immediateAfterSecondClick.sortColumnIndex,
                              immediateAfterSecondClick.sortDescending,
                              immediateAfterSecondClick.firstVisibleTaskId,
                              immediateAfterSecondClick.primarySelectedTaskId,
                              descending.hasActiveSort,
                              descending.sortColumnIndex,
                              descending.sortDescending,
                              descending.firstVisibleTaskId,
                              descending.primarySelectedTaskId));
    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneHeaderResizeChangesVisibleWidth(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane header-resize validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane header-resize baseline reset"),
                  L"Failed to hide the file-operations issues pane before header-resize validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane header-resize validation"), L"Failed to show the file-operations issues pane.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for header-resize validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const uint64_t taskId         = 0xF300000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskId);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration = baseline.refreshGeneration;
    fileOps->RecordTaskDiagnostic(taskId,
                                  FILESYSTEM_MOVE,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                  HRESULT_FROM_WIN32(ERROR_DISK_FULL),
                                  L"selftest.dxui.issues.resize",
                                  L"issues-resize-target",
                                  L"C:\\selftest-issues-resize-source.txt",
                                  L"D:\\selftest-issues-resize-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for header-resize validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for header-resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId),
                  std::format(L"Failed to select task {} before issues-pane header-resize validation.", taskId));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskId; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before header-resize validation.", taskId));
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineTaskHeaderWidth     = std::max(0.0f, snapshot.taskHeaderRect.right - snapshot.taskHeaderRect.left);
    const float baselineOperationHeaderLeft = snapshot.operationHeaderRect.left;
    const uint64_t baselineRenderCount      = snapshot.dxRenderCount;
    const uint64_t baselineResizeCount      = snapshot.dxResizeCount;
    const uint64_t baselineVisibleRows      = snapshot.visibleWork.visibleRowCount;
    const size_t baselineVisibleColumns     = snapshot.visibleWork.visibleColumnCount;

    state.Require(SendIssuesPaneTaskHeaderResizeDrag(pane, snapshot, 56.0f), L"Failed to drive a real issues-pane task-header resize drag.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsIssuesPane::SelfTestSnapshot resized{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.dxRenderCount > baselineRenderCount && currentTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f &&
               value.operationHeaderRect.left > baselineOperationHeaderLeft + 10.0f && value.selectionCount == 1u && value.primarySelectedTaskId == taskId &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleRowCount <= baselineVisibleRows + 1u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumns && value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      resized),
                  std::format(L"Issues pane did not widen the task header after a real pointer resize drag. "
                              L"baselineWidth={:.2f} currentWidth={:.2f} baselineOperationLeft={:.2f} currentOperationLeft={:.2f} "
                              L"visibleRows={} visibleColumns={} resizeCount={} resizeFailures={} selectedTask={}",
                              baselineTaskHeaderWidth,
                              std::max(0.0f, resized.taskHeaderRect.right - resized.taskHeaderRect.left),
                              baselineOperationHeaderLeft,
                              resized.operationHeaderRect.left,
                              resized.visibleWork.visibleRowCount,
                              resized.visibleWork.visibleColumnCount,
                              resized.dxResizeCount,
                              resized.dxResizeFailureCount,
                              resized.primarySelectedTaskId));

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneHeaderDragReordersColumnsWithoutSort(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane header-reorder validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane header-reorder baseline reset"),
                  L"Failed to hide the file-operations issues pane before header-reorder validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane header-reorder validation"), L"Failed to show the file-operations issues pane.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for header-reorder validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const uint64_t taskId         = 0xF400000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskId);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration = baseline.refreshGeneration;
    fileOps->RecordTaskDiagnostic(taskId,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_DISK_FULL),
                                  L"selftest.dxui.issues.reorder",
                                  L"issues-reorder-target",
                                  L"C:\\selftest-issues-reorder-source.txt",
                                  L"D:\\selftest-issues-reorder-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for header-reorder validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for header-reorder validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! snapshot.hasActiveSort, L"Issues pane should start unsorted before header-reorder validation.");
    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId),
                  std::format(L"Failed to select task {} before issues-pane header-reorder validation.", taskId));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskId; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before header-reorder validation.", taskId));
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineTaskLeft        = snapshot.taskHeaderRect.left;
    const float baselineOperationLeft   = snapshot.operationHeaderRect.left;
    const uint64_t baselineRenderCount  = snapshot.dxRenderCount;
    const size_t baselineVisibleColumns = snapshot.visibleWork.visibleColumnCount;

    const float dragStartXDip  = (snapshot.operationHeaderRect.left + snapshot.operationHeaderRect.right) * 0.5f;
    const float dragStartYDip  = (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f;
    const float dragTargetXDip = snapshot.taskHeaderRect.left + 12.0f;
    const LPARAM dragStart     = DipPointToClientLParam(pane, dragStartXDip, dragStartYDip);
    const LPARAM dragTarget    = DipPointToClientLParam(pane, dragTargetXDip, dragStartYDip);
    SendMouseDragToResolvedPointWindow(pane, dragStart, dragTarget);

    FileOperationsIssuesPane::SelfTestSnapshot reordered{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > baselineRenderCount && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               value.operationHeaderRect.left + 4.0f < baselineOperationLeft && value.taskHeaderRect.left > baselineTaskLeft + 10.0f && ! value.hasActiveSort &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskId && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumns && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      reordered),
                  std::format(L"Issues pane did not reorder the visible Operation header ahead of Task after a real pointer drag. "
                              L"baselineTaskLeft={:.2f} baselineOperationLeft={:.2f} currentTaskLeft={:.2f} currentOperationLeft={:.2f} "
                              L"sortActive={} selectedTask={} visibleColumns={} resizeFailures={}",
                              baselineTaskLeft,
                              baselineOperationLeft,
                              reordered.taskHeaderRect.left,
                              reordered.operationHeaderRect.left,
                              reordered.hasActiveSort,
                              reordered.primarySelectedTaskId,
                              reordered.visibleWork.visibleColumnCount,
                              reordered.dxResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneCopyFollowsReorderedColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane reordered copy validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };

    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane reordered copy baseline reset"),
                  L"Failed to hide the file-operations issues pane before reordered copy validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane reordered copy validation"), L"Failed to show issues pane for reordered copy validation.");
    const HWND pane = getVisibleIssuesPane();
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"Issues pane handle unavailable for reordered copy validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    const uint64_t taskId         = 0xF500000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskId);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration    = baseline.refreshGeneration;
    const std::wstring expectedOperation = LoadStringResource(nullptr, IDS_CMD_COPY);

    fileOps->RecordTaskDiagnostic(taskId,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                                  L"selftest.dxui.issues.copy",
                                  L"issues-copy-target",
                                  L"C:\\selftest-issues-copy-source.txt",
                                  L"D:\\selftest-issues-copy-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for reordered copy validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for reordered copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId),
                  std::format(L"Failed to select task {} before issues-pane reordered copy validation.", taskId));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskId; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before reordered copy validation.", taskId));
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRenderCount = snapshot.dxRenderCount;
    const LPARAM dragStart             = DipPointToClientLParam(pane,
                                                                (snapshot.operationHeaderRect.left + snapshot.operationHeaderRect.right) * 0.5f,
                                                                (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f);
    const LPARAM dragTarget =
        DipPointToClientLParam(pane, snapshot.taskHeaderRect.left + 12.0f, (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f);
    SendMouseDragToResolvedPointWindow(pane, dragStart, dragTarget);

    FileOperationsIssuesPane::SelfTestSnapshot reordered{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > baselineRenderCount && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && ! value.hasActiveSort &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskId && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      reordered),
                  L"Issues pane did not expose reordered visible headers before clipboard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestFocusGrid(pane), L"Failed to focus the issues-pane grid before reordered copy validation.");
    ClearClipboardContents(pane);
    SendMessageW(pane, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(pane, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(pane, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(pane, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(pane);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Issues pane Ctrl+C should copy the reordered visible row content to the clipboard.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t lineBreak       = copiedSelection.find_first_of(L"\r\n");
    const std::wstring firstLine = lineBreak == std::wstring::npos ? copiedSelection : copiedSelection.substr(0u, lineBreak);
    std::vector<std::wstring> columns;
    size_t start = 0u;
    while (start <= firstLine.size())
    {
        const size_t tab = firstLine.find(L'\t', start);
        if (tab == std::wstring::npos)
        {
            columns.push_back(firstLine.substr(start));
            break;
        }
        columns.push_back(firstLine.substr(start, tab - start));
        start = tab + 1u;
    }

    const std::wstring expectedTask = std::to_wstring(static_cast<unsigned long long>(taskId));
    state.Require(columns.size() >= 3u,
                  std::format(L"Issues pane clipboard copy should export at least three visible columns after reorder, got {}.", columns.size()));
    state.Require(columns.size() >= 2u && columns[1] == expectedOperation,
                  std::format(L"Issues pane clipboard copy should place the reordered Operation column before Task. Expected second column '{}', got '{}'.",
                              expectedOperation,
                              columns.size() >= 2u ? columns[1] : std::wstring{}));
    state.Require(columns.size() >= 3u && columns[2] == expectedTask,
                  std::format(L"Issues pane clipboard copy should keep the selected task id '{}' immediately after the reordered Operation column. Got '{}'.",
                              expectedTask,
                              columns.size() >= 3u ? columns[2] : std::wstring{}));

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneReorderedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane reordered-sort validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };

    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane reordered-sort baseline reset"),
                  L"Failed to hide the file-operations issues pane before reordered-sort validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane reordered-sort validation"), L"Failed to show the file-operations issues pane.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for reordered-sort validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const uint64_t baseTaskId     = 0xF500000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const uint64_t taskHigh       = baseTaskId + 9u;
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskHigh);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration = baseline.refreshGeneration;
    fileOps->RecordTaskDiagnostic(taskHigh,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_DISK_FULL),
                                  L"selftest.dxui.issues.reorderedSort",
                                  L"issues-reordered-sort-target",
                                  L"C:\\selftest-issues-reordered-sort-source.txt",
                                  L"D:\\selftest-issues-reordered-sort-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for reordered-sort validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for reordered-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskHigh),
                  std::format(L"Failed to select task {} before issues-pane reordered-sort validation.", taskHigh));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before reordered-sort validation.", taskHigh));
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineTaskLeft            = snapshot.taskHeaderRect.left;
    const float baselineOperationLeft       = snapshot.operationHeaderRect.left;
    const size_t baselineVisibleColumnCount = snapshot.visibleWork.visibleColumnCount;
    const uint64_t baselineRenderCount      = snapshot.dxRenderCount;

    const float dragStartXDip  = (snapshot.operationHeaderRect.left + snapshot.operationHeaderRect.right) * 0.5f;
    const float dragStartYDip  = (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f;
    const float dragTargetXDip = snapshot.taskHeaderRect.left + 12.0f;
    const LPARAM dragStart     = DipPointToClientLParam(pane, dragStartXDip, dragStartYDip);
    const LPARAM dragTarget    = DipPointToClientLParam(pane, dragTargetXDip, dragStartYDip);
    SendMouseDragToResolvedPointWindow(pane, dragStart, dragTarget);

    FileOperationsIssuesPane::SelfTestSnapshot reordered{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > baselineRenderCount && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               value.operationHeaderRect.left + 4.0f < baselineOperationLeft && value.taskHeaderRect.left > baselineTaskLeft + 10.0f && ! value.hasActiveSort &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      reordered),
                  std::format(L"Issues pane did not keep the reordered Operation header ahead of Task before sort validation. "
                              L"baselineTaskLeft={:.2f} baselineOperationLeft={:.2f} currentTaskLeft={:.2f} currentOperationLeft={:.2f} "
                              L"sortActive={} selectedTask={} visibleColumns={} resizeFailures={}",
                              baselineTaskLeft,
                              baselineOperationLeft,
                              reordered.taskHeaderRect.left,
                              reordered.operationHeaderRect.left,
                              reordered.hasActiveSort,
                              reordered.primarySelectedTaskId,
                              reordered.visibleWork.visibleColumnCount,
                              reordered.dxResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto clickTaskHeader = [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const LPARAM clickPoint = DipPointToClientLParam(
            pane, (value.taskHeaderRect.left + value.taskHeaderRect.right) * 0.5f, (value.taskHeaderRect.top + value.taskHeaderRect.bottom) * 0.5f);
        SendMessageW(pane, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(pane, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(pane, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    clickTaskHeader(reordered);

    FileOperationsIssuesPane::SelfTestSnapshot ascending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > reordered.dxRenderCount && value.hasActiveSort && value.sortColumnIndex == 1u && ! value.sortDescending &&
               value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      ascending),
                  std::format(L"Issues pane did not keep reordered columns intact after the first task-header sort click. "
                              L"sortActive={} column={} descending={} taskLeft={:.2f} operationLeft={:.2f} selectedTask={} visibleColumns={}",
                              ascending.hasActiveSort,
                              ascending.sortColumnIndex,
                              ascending.sortDescending,
                              ascending.taskHeaderRect.left,
                              ascending.operationHeaderRect.left,
                              ascending.primarySelectedTaskId,
                              ascending.visibleWork.visibleColumnCount));

    clickTaskHeader(ascending);

    FileOperationsIssuesPane::SelfTestSnapshot descending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > ascending.dxRenderCount && value.hasActiveSort && value.sortColumnIndex == 1u && value.sortDescending &&
               value.firstVisibleTaskId == taskHigh && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && value.selectionCount == 1u &&
               value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      descending),
                  std::format(L"Issues pane did not keep reordered columns intact after the second task-header sort click. "
                              L"sortActive={} column={} descending={} firstTask={} taskLeft={:.2f} operationLeft={:.2f} selectedTask={} visibleColumns={}",
                              descending.hasActiveSort,
                              descending.sortColumnIndex,
                              descending.sortDescending,
                              descending.firstVisibleTaskId,
                              descending.taskHeaderRect.left,
                              descending.operationHeaderRect.left,
                              descending.primarySelectedTaskId,
                              descending.visibleWork.visibleColumnCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneReorderedCopyFollowsVisibleColumnsAfterSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane reordered-copy-after-sort validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };

    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane reordered-copy-after-sort baseline reset"),
                  L"Failed to hide the file-operations issues pane before reordered copy after sort validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane reordered-copy-after-sort validation"),
                  L"Failed to show issues pane for reordered copy after sort validation.");
    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"Issues pane handle unavailable for reordered copy after sort validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    const uint64_t taskId         = 0xF600000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskId);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration    = baseline.refreshGeneration;
    const std::wstring expectedOperation = LoadStringResource(nullptr, IDS_CMD_COPY);

    fileOps->RecordTaskDiagnostic(taskId,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                                  L"selftest.dxui.issues.copyAfterSort",
                                  L"issues-copy-after-sort-target",
                                  L"C:\\selftest-issues-copy-after-sort-source.txt",
                                  L"D:\\selftest-issues-copy-after-sort-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for reordered copy after sort validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for reordered copy after sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId),
                  std::format(L"Failed to select task {} before issues-pane reordered copy after sort validation.", taskId));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskId; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before reordered copy after sort validation.", taskId));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleColumnCount = snapshot.visibleWork.visibleColumnCount;
    const uint64_t baselineRenderCount      = snapshot.dxRenderCount;
    const LPARAM dragStart                  = DipPointToClientLParam(pane,
                                                                     (snapshot.operationHeaderRect.left + snapshot.operationHeaderRect.right) * 0.5f,
                                                                     (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f);
    const LPARAM dragTarget =
        DipPointToClientLParam(pane, snapshot.taskHeaderRect.left + 12.0f, (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f);
    SendMouseDragToResolvedPointWindow(pane, dragStart, dragTarget);

    FileOperationsIssuesPane::SelfTestSnapshot reordered{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > baselineRenderCount && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && ! value.hasActiveSort &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskId && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      reordered),
                  L"Issues pane did not expose reordered visible headers before reordered copy after sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto clickTaskHeader = [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const LPARAM clickPoint = DipPointToClientLParam(
            pane, (value.taskHeaderRect.left + value.taskHeaderRect.right) * 0.5f, (value.taskHeaderRect.top + value.taskHeaderRect.bottom) * 0.5f);
        SendMessageW(pane, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(pane, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(pane, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    clickTaskHeader(reordered);

    FileOperationsIssuesPane::SelfTestSnapshot ascending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > reordered.dxRenderCount && value.hasActiveSort && value.sortColumnIndex == 1u && ! value.sortDescending &&
               value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && value.selectionCount == 1u && value.primarySelectedTaskId == taskId &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      ascending),
                  L"Issues pane did not stay reordered after the first sort click during reordered copy after sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    clickTaskHeader(ascending);

    FileOperationsIssuesPane::SelfTestSnapshot descending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > ascending.dxRenderCount && value.hasActiveSort && value.sortColumnIndex == 1u && value.sortDescending &&
               value.firstVisibleTaskId == taskId && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && value.selectionCount == 1u &&
               value.primarySelectedTaskId == taskId && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      descending),
                  L"Issues pane did not stay reordered after the second sort click during reordered copy after sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestFocusGrid(pane), L"Failed to focus the issues-pane grid before reordered copy after sort validation.");
    ClearClipboardContents(pane);
    SendMessageW(pane, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(pane, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(pane, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(pane, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(pane);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Issues pane Ctrl+C should copy the reordered visible row content after sort churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t lineBreak       = copiedSelection.find_first_of(L"\r\n");
    const std::wstring firstLine = lineBreak == std::wstring::npos ? copiedSelection : copiedSelection.substr(0u, lineBreak);
    std::vector<std::wstring> columns;
    size_t start = 0u;
    while (start <= firstLine.size())
    {
        const size_t tab = firstLine.find(L'\t', start);
        if (tab == std::wstring::npos)
        {
            columns.push_back(firstLine.substr(start));
            break;
        }
        columns.push_back(firstLine.substr(start, tab - start));
        start = tab + 1u;
    }

    const std::wstring expectedTask = std::to_wstring(static_cast<unsigned long long>(taskId));
    state.Require(columns.size() >= 3u,
                  std::format(L"Issues pane clipboard copy should export at least three visible columns after reordered sort churn, got {}.", columns.size()));
    state.Require(columns.size() >= 2u && columns[1] == expectedOperation,
                  std::format(L"Issues pane clipboard copy should still place the reordered Operation column before Task after sort churn. Expected second "
                              L"column '{}', got '{}'.",
                              expectedOperation,
                              columns.size() >= 2u ? columns[1] : std::wstring{}));
    state.Require(
        columns.size() >= 3u && columns[2] == expectedTask,
        std::format(
            L"Issues pane clipboard copy should keep selected task id '{}' immediately after the reordered Operation column after sort churn. Got '{}'.",
            expectedTask,
            columns.size() >= 3u ? columns[2] : std::wstring{}));

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneResizedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane resized-sort validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane resized-sort baseline reset"),
                  L"Failed to hide the file-operations issues pane before resized-sort validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane resized-sort validation"), L"Failed to show the file-operations issues pane.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for resized-sort validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const uint64_t taskId         = 0xF700000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskId);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration = baseline.refreshGeneration;
    fileOps->RecordTaskDiagnostic(taskId,
                                  FILESYSTEM_MOVE,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                  HRESULT_FROM_WIN32(ERROR_DISK_FULL),
                                  L"selftest.dxui.issues.resizeSort",
                                  L"issues-resize-sort-target",
                                  L"C:\\selftest-issues-resize-sort-source.txt",
                                  L"D:\\selftest-issues-resize-sort-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for resized-sort validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for resized-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId),
                  std::format(L"Failed to select task {} before issues-pane resized-sort validation.", taskId));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskId; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before resized-sort validation.", taskId));
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineTaskHeaderWidth     = std::max(0.0f, snapshot.taskHeaderRect.right - snapshot.taskHeaderRect.left);
    const float baselineOperationHeaderLeft = snapshot.operationHeaderRect.left;
    const size_t baselineVisibleColumnCount = snapshot.visibleWork.visibleColumnCount;
    const uint64_t baselineRenderCount      = snapshot.dxRenderCount;
    const uint64_t baselineResizeCount      = snapshot.dxResizeCount;
    const uint64_t baselineVisibleRows      = snapshot.visibleWork.visibleRowCount;

    state.Require(SendIssuesPaneTaskHeaderResizeDrag(pane, snapshot, 56.0f),
                  L"Failed to drive a real issues-pane task-header resize drag before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsIssuesPane::SelfTestSnapshot resized{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.dxRenderCount > baselineRenderCount && currentTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f &&
               value.operationHeaderRect.left > baselineOperationHeaderLeft + 10.0f && value.selectionCount == 1u && value.primarySelectedTaskId == taskId &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleRowCount <= baselineVisibleRows + 1u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeCount == baselineResizeCount &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      resized),
                  L"Issues pane did not expose the widened Task header before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto clickTaskHeader = [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const LPARAM clickPoint = DipPointToClientLParam(
            pane, (value.taskHeaderRect.left + value.taskHeaderRect.right) * 0.5f, (value.taskHeaderRect.top + value.taskHeaderRect.bottom) * 0.5f);
        SendMessageW(pane, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(pane, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(pane, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    clickTaskHeader(resized);

    FileOperationsIssuesPane::SelfTestSnapshot ascending{};
    state.Require(
        waitForSnapshot(
            [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.dxRenderCount > resized.dxRenderCount && value.hasActiveSort && value.sortColumnIndex == 1u && ! value.sortDescending &&
               currentTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f && value.operationHeaderRect.left > baselineOperationHeaderLeft + 10.0f &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskId && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeCount == baselineResizeCount &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            ascending),
        std::format(
            L"Issues pane did not keep the widened Task header after the first sort click. "
            L"sortActive={} column={} descending={} width={:.2f} operationLeft={:.2f} selectedTask={} visibleColumns={} resizeCount={} resizeFailures={}",
            ascending.hasActiveSort,
            ascending.sortColumnIndex,
            ascending.sortDescending,
            std::max(0.0f, ascending.taskHeaderRect.right - ascending.taskHeaderRect.left),
            ascending.operationHeaderRect.left,
            ascending.primarySelectedTaskId,
            ascending.visibleWork.visibleColumnCount,
            ascending.dxResizeCount,
            ascending.dxResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    clickTaskHeader(ascending);

    FileOperationsIssuesPane::SelfTestSnapshot descending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.dxRenderCount > ascending.dxRenderCount && value.hasActiveSort && value.sortColumnIndex == 1u && value.sortDescending &&
               value.firstVisibleTaskId == taskId && currentTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f &&
               value.operationHeaderRect.left > baselineOperationHeaderLeft + 10.0f && value.selectionCount == 1u && value.primarySelectedTaskId == taskId &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      descending),
                  std::format(L"Issues pane did not keep the widened Task header after the second sort click. "
                              L"sortActive={} column={} descending={} firstTask={} width={:.2f} operationLeft={:.2f} selectedTask={} visibleColumns={} "
                              L"resizeCount={} resizeFailures={}",
                              descending.hasActiveSort,
                              descending.sortColumnIndex,
                              descending.sortDescending,
                              descending.firstVisibleTaskId,
                              std::max(0.0f, descending.taskHeaderRect.right - descending.taskHeaderRect.left),
                              descending.operationHeaderRect.left,
                              descending.primarySelectedTaskId,
                              descending.visibleWork.visibleColumnCount,
                              descending.dxResizeCount,
                              descending.dxResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneReorderedResizedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane reordered-resized-sort validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };

    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane reordered-resized-sort baseline reset"),
                  L"Failed to hide the file-operations issues pane before reordered-resized-sort validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane reordered-resized-sort validation"), L"Failed to show the file-operations issues pane.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for reordered-resized-sort validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const uint64_t baseTaskId     = 0xF800000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const uint64_t taskHigh       = baseTaskId + 9u;
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskHigh);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration = baseline.refreshGeneration;
    fileOps->RecordTaskDiagnostic(taskHigh,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_DISK_FULL),
                                  L"selftest.dxui.issues.reorderedResizeSort",
                                  L"issues-reordered-resize-sort-target",
                                  L"C:\\selftest-issues-reordered-resize-sort-source.txt",
                                  L"D:\\selftest-issues-reordered-resize-sort-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for reordered-resized-sort validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for reordered-resized-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskHigh),
                  std::format(L"Failed to select task {} before issues-pane reordered-resized-sort validation.", taskHigh));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before reordered-resized-sort validation.", taskHigh));
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineTaskLeft            = snapshot.taskHeaderRect.left;
    const float baselineOperationLeft       = snapshot.operationHeaderRect.left;
    const float baselineTaskHeaderWidth     = std::max(0.0f, snapshot.taskHeaderRect.right - snapshot.taskHeaderRect.left);
    const size_t baselineVisibleColumnCount = snapshot.visibleWork.visibleColumnCount;
    const uint64_t baselineRenderCount      = snapshot.dxRenderCount;
    const uint64_t baselineResizeCount      = snapshot.dxResizeCount;

    const float dragStartXDip  = (snapshot.operationHeaderRect.left + snapshot.operationHeaderRect.right) * 0.5f;
    const float dragStartYDip  = (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f;
    const float dragTargetXDip = snapshot.taskHeaderRect.left + 12.0f;
    const LPARAM dragStart     = DipPointToClientLParam(pane, dragStartXDip, dragStartYDip);
    const LPARAM dragTarget    = DipPointToClientLParam(pane, dragTargetXDip, dragStartYDip);
    SendMouseDragToResolvedPointWindow(pane, dragStart, dragTarget);

    FileOperationsIssuesPane::SelfTestSnapshot reordered{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > baselineRenderCount && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               value.operationHeaderRect.left + 4.0f < baselineOperationLeft && value.taskHeaderRect.left > baselineTaskLeft + 10.0f && ! value.hasActiveSort &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      reordered),
                  std::format(L"Issues pane did not expose reordered columns before resized-sort validation. "
                              L"baselineTaskLeft={:.2f} baselineOperationLeft={:.2f} currentTaskLeft={:.2f} currentOperationLeft={:.2f}",
                              baselineTaskLeft,
                              baselineOperationLeft,
                              reordered.taskHeaderRect.left,
                              reordered.operationHeaderRect.left));
    if (! state.failure.empty())
    {
        return false;
    }

    const float reorderedTaskHeaderWidth = std::max(0.0f, reordered.taskHeaderRect.right - reordered.taskHeaderRect.left);
    state.Require(SendIssuesPaneTaskHeaderResizeDrag(pane, reordered, 56.0f), L"Failed to drive a real issues-pane task-header resize drag after reorder.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsIssuesPane::SelfTestSnapshot resized{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.dxRenderCount > reordered.dxRenderCount && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               currentTaskHeaderWidth > reorderedTaskHeaderWidth + 20.0f && value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      resized),
                  L"Issues pane did not expose the reordered widened Task header before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedTaskHeaderWidth = std::max(0.0f, resized.taskHeaderRect.right - resized.taskHeaderRect.left);
    state.Require(widenedTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f, L"Task header width did not grow enough after reorder+resize.");
    const auto clickTaskHeader = [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const LPARAM clickPoint = DipPointToClientLParam(
            pane, (value.taskHeaderRect.left + value.taskHeaderRect.right) * 0.5f, (value.taskHeaderRect.top + value.taskHeaderRect.bottom) * 0.5f);
        SendMessageW(pane, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(pane, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(pane, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    clickTaskHeader(resized);

    FileOperationsIssuesPane::SelfTestSnapshot ascending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.sortColumnIndex == 1u && ! value.sortDescending && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               currentTaskHeaderWidth > reorderedTaskHeaderWidth + 20.0f && value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      ascending),
                  L"Issues pane did not keep reordered+widened columns stable through ascending sort.");
    if (! state.failure.empty())
    {
        return false;
    }

    clickTaskHeader(ascending);

    FileOperationsIssuesPane::SelfTestSnapshot descending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.sortColumnIndex == 1u && value.sortDescending && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               currentTaskHeaderWidth > reorderedTaskHeaderWidth + 20.0f && value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh &&
               value.firstVisibleTaskId == taskHigh && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeCount == baselineResizeCount &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      descending),
                  L"Issues pane did not keep reordered+widened columns stable through descending sort.");
    if (! state.failure.empty())
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneReorderedResizedCopyFollowsVisibleColumnsAfterSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane reordered-resized-copy-after-sort validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };

    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane reordered-resized-copy-after-sort baseline reset"),
                  L"Failed to hide the file-operations issues pane before reordered-resized copy after sort validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    state.Require(setIssuesVisible(true, L"issues-pane reordered-resized-copy-after-sort validation"),
                  L"Failed to show issues pane for reordered-resized copy after sort validation.");
    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"Issues pane handle unavailable for reordered-resized copy after sort validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    const uint64_t taskId         = 0xF900000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskId);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration    = baseline.refreshGeneration;
    const std::wstring expectedOperation = LoadStringResource(nullptr, IDS_CMD_COPY);

    fileOps->RecordTaskDiagnostic(taskId,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                                  L"selftest.dxui.issues.copyAfterReorderedResizeSort",
                                  L"issues-copy-after-reordered-resize-sort-target",
                                  L"C:\\selftest-issues-copy-after-reordered-resize-sort-source.txt",
                                  L"D:\\selftest-issues-copy-after-reordered-resize-sort-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true),
                  L"Failed to request issues-pane refresh for reordered-resized copy after sort validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount > 0u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for reordered-resized copy after sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId),
                  std::format(L"Failed to select task {} before issues-pane reordered-resized copy after sort validation.", taskId));
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskId; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before reordered-resized copy after sort validation.", taskId));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleColumnCount = snapshot.visibleWork.visibleColumnCount;
    const uint64_t baselineRenderCount      = snapshot.dxRenderCount;
    const uint64_t baselineResizeCount      = snapshot.dxResizeCount;
    const LPARAM dragStart                  = DipPointToClientLParam(pane,
                                                                     (snapshot.operationHeaderRect.left + snapshot.operationHeaderRect.right) * 0.5f,
                                                                     (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f);
    const LPARAM dragTarget =
        DipPointToClientLParam(pane, snapshot.taskHeaderRect.left + 12.0f, (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f);
    SendMouseDragToResolvedPointWindow(pane, dragStart, dragTarget);

    FileOperationsIssuesPane::SelfTestSnapshot reordered{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.dxRenderCount > baselineRenderCount && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && ! value.hasActiveSort &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskId && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      reordered),
                  L"Issues pane did not expose reordered visible headers before reordered-resized copy after sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float reorderedTaskHeaderWidth = std::max(0.0f, reordered.taskHeaderRect.right - reordered.taskHeaderRect.left);
    state.Require(SendIssuesPaneTaskHeaderResizeDrag(pane, reordered, 56.0f),
                  L"Failed to drive a real issues-pane task-header resize drag before reordered-resized copy after sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsIssuesPane::SelfTestSnapshot resized{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.dxRenderCount > reordered.dxRenderCount && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               currentTaskHeaderWidth > reorderedTaskHeaderWidth + 20.0f && value.selectionCount == 1u && value.primarySelectedTaskId == taskId &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      resized),
                  L"Issues pane did not expose the reordered widened Task header before reordered-resized copy after sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto clickTaskHeader = [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const LPARAM clickPoint = DipPointToClientLParam(
            pane, (value.taskHeaderRect.left + value.taskHeaderRect.right) * 0.5f, (value.taskHeaderRect.top + value.taskHeaderRect.bottom) * 0.5f);
        SendMessageW(pane, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(pane, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(pane, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    clickTaskHeader(resized);

    FileOperationsIssuesPane::SelfTestSnapshot ascending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.sortColumnIndex == 1u && ! value.sortDescending && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               currentTaskHeaderWidth > reorderedTaskHeaderWidth + 20.0f && value.selectionCount == 1u && value.primarySelectedTaskId == taskId &&
               value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      ascending),
                  L"Issues pane did not keep reordered+widened columns stable through ascending sort before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    clickTaskHeader(ascending);

    FileOperationsIssuesPane::SelfTestSnapshot descending{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.sortColumnIndex == 1u && value.sortDescending && value.firstVisibleTaskId == taskId &&
               value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && currentTaskHeaderWidth > reorderedTaskHeaderWidth + 20.0f &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskId && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeCount == baselineResizeCount &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      descending),
                  L"Issues pane did not keep reordered+widened columns stable through descending sort before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestFocusGrid(pane),
                  L"Failed to focus the issues-pane grid before reordered-resized copy after sort validation.");
    ClearClipboardContents(pane);
    SendMessageW(pane, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(pane, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(pane, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(pane, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(pane);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Issues pane Ctrl+C should copy the reordered-resized visible row content after sort churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t lineBreak       = copiedSelection.find_first_of(L"\r\n");
    const std::wstring firstLine = lineBreak == std::wstring::npos ? copiedSelection : copiedSelection.substr(0u, lineBreak);
    std::vector<std::wstring> columns;
    size_t start = 0u;
    while (start <= firstLine.size())
    {
        const size_t tab = firstLine.find(L'\t', start);
        if (tab == std::wstring::npos)
        {
            columns.push_back(firstLine.substr(start));
            break;
        }
        columns.push_back(firstLine.substr(start, tab - start));
        start = tab + 1u;
    }

    const std::wstring expectedTask = std::to_wstring(static_cast<unsigned long long>(taskId));
    state.Require(
        columns.size() >= 3u,
        std::format(L"Issues pane clipboard copy should export at least three visible columns after reordered-resized sort churn, got {}.", columns.size()));
    state.Require(columns.size() >= 2u && columns[1] == expectedOperation,
                  std::format(L"Issues pane clipboard copy should still place the reordered Operation column before Task after reorder+resize sort churn. "
                              L"Expected second column '{}', got '{}'.",
                              expectedOperation,
                              columns.size() >= 2u ? columns[1] : std::wstring{}));
    state.Require(columns.size() >= 3u && columns[2] == expectedTask,
                  std::format(L"Issues pane clipboard copy should keep selected task id '{}' immediately after the reordered Operation column after "
                              L"reorder+resize sort churn. Got '{}'.",
                              expectedTask,
                              columns.size() >= 3u ? columns[2] : std::wstring{}));

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneRestoresCombinedViewStateAfterRecreate(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane persisted-view-state validation.");
    if (! fileOps)
    {
        return false;
    }

    const auto getAnyIssuesPane = [&]() noexcept -> HWND
    {
        const HWND pane = fileOps->GetIssuesPaneHwndForSelfTest();
        return (pane && IsWindow(pane) != FALSE) ? pane : nullptr;
    };
    const auto getVisibleIssuesPane = [&]() noexcept -> HWND
    {
        const HWND pane = fileOps->GetIssuesPaneHwndForSelfTest();
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible                                                                = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const std::optional<Common::Settings::FileOperationsSettings> previousFileOperations = g_settings.fileOperations;
    const auto restoreSettings = wil::scope_exit([&] noexcept { g_settings.fileOperations = previousFileOperations; });

    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });
    state.Require(setIssuesVisible(false, L"issues-pane combined-state baseline reset"),
                  L"Failed to hide the file-operations issues pane before combined-state recreate validation.");
    const auto restoreViewState = ResetFileOperationsIssuesPaneViewStateForTest(fileOps);

    Common::Settings::FileOperationsSettings isolatedSettings = previousFileOperations.value_or(Common::Settings::FileOperationsSettings{});
    isolatedSettings.autoDismissSuccess                       = true;
    isolatedSettings.preCalcEnabled                           = false;
    isolatedSettings.preCalcMaxWorkers                        = 7u;
    isolatedSettings.crossFsBridgeBufferSizeKB                = 2048u;
    isolatedSettings.defaultBandwidthLimitBytesPerSecond      = 987654u;
    isolatedSettings.maxDiagnosticsLogFiles                   = 23u;
    isolatedSettings.issuesPaneSortColumnId.clear();
    isolatedSettings.issuesPaneSortDescending = false;
    isolatedSettings.issuesPaneGridLayout.clear();
    g_settings.fileOperations = isolatedSettings;

    const auto requireUnrelatedFileOperationsSettingsIntact = [&](std::wstring_view label) noexcept
    {
        state.Require(g_settings.fileOperations.has_value(), std::format(L"File-operations settings should remain present {}.", label));
        if (! g_settings.fileOperations.has_value())
        {
            return false;
        }

        const auto& current = g_settings.fileOperations.value();
        state.Require(current.autoDismissSuccess == isolatedSettings.autoDismissSuccess,
                      std::format(L"Issues-pane view-state save should preserve autoDismissSuccess {}.", label));
        state.Require(current.preCalcEnabled == isolatedSettings.preCalcEnabled,
                      std::format(L"Issues-pane view-state save should preserve preCalcEnabled {}.", label));
        state.Require(current.preCalcMaxWorkers == isolatedSettings.preCalcMaxWorkers,
                      std::format(L"Issues-pane view-state save should preserve preCalcMaxWorkers {}.", label));
        state.Require(current.crossFsBridgeBufferSizeKB == isolatedSettings.crossFsBridgeBufferSizeKB,
                      std::format(L"Issues-pane view-state save should preserve crossFsBridgeBufferSizeKB {}.", label));
        state.Require(current.defaultBandwidthLimitBytesPerSecond == isolatedSettings.defaultBandwidthLimitBytesPerSecond,
                      std::format(L"Issues-pane view-state save should preserve defaultBandwidthLimitBytesPerSecond {}.", label));
        state.Require(current.maxDiagnosticsLogFiles == isolatedSettings.maxDiagnosticsLogFiles,
                      std::format(L"Issues-pane view-state save should preserve maxDiagnosticsLogFiles {}.", label));
        return state.failure.empty();
    };

    if (const HWND existingPane = getAnyIssuesPane(); existingPane && IsWindow(existingPane) != FALSE)
    {
        DestroyWindow(existingPane);
        state.Require(WaitForWindowClosed(existingPane, SelfTest::Scale(3000ms)),
                      L"Failed to destroy the existing issues pane before persisted-view-state validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(setIssuesVisible(true, L"issues-pane persisted-view-state validation"), L"Failed to show issues pane for persisted-view-state validation.");
    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"Issues pane handle unavailable for persisted-view-state validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot =
        [&](HWND targetPane, auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(targetPane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(targetPane, outSnapshot) && predicate(outSnapshot);
    };
    const auto waitForSnapshotReady = [&](HWND targetPane, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        return waitForSnapshot(targetPane,
                               [](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        {
            return value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
                   value.operationHeaderRect.right > value.operationHeaderRect.left;
        },
                               timeout,
                               outSnapshot);
    };

    const uint64_t taskHigh              = 0xFB00000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const uint64_t taskLow               = taskHigh - 1u;
    const std::wstring expectedOperation = LoadStringResource(nullptr, IDS_CMD_MOVE);
    const auto cleanupDiagnostics        = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskLow);
        fileOps->DebugRemoveDiagnosticsForTask(taskHigh);
        if (const HWND currentPane = getAnyIssuesPane(); currentPane && IsWindow(currentPane) != FALSE)
        {
            static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(currentPane, true));
        }
    });

    FileOperationsIssuesPane::SelfTestSnapshot baseline{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, baseline), L"Failed to capture baseline issues-pane snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineGeneration = baseline.refreshGeneration;
    fileOps->RecordTaskDiagnostic(taskLow,
                                  FILESYSTEM_MOVE,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                                  L"selftest.dxui.issues.persistedViewState",
                                  L"issues-persisted-view-state-low",
                                  L"C:\\selftest-issues-persisted-low-source.txt",
                                  L"D:\\selftest-issues-persisted-low-dest.txt");
    fileOps->RecordTaskDiagnostic(taskHigh,
                                  FILESYSTEM_MOVE,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_DISK_FULL),
                                  L"selftest.dxui.issues.persistedViewState",
                                  L"issues-persisted-view-state-high",
                                  L"C:\\selftest-issues-persisted-high-source.txt",
                                  L"D:\\selftest-issues-persisted-high-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh for persisted-view-state validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(pane,
                                  [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > baselineGeneration && value.rowCount >= 2u && value.visibleWork.visibleRowCount > 0u &&
               value.visibleWork.visibleColumnCount > 0u && value.taskHeaderRect.right > value.taskHeaderRect.left &&
               value.operationHeaderRect.right > value.operationHeaderRect.left;
    },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  L"Issues pane did not expose visible rows plus adjacent visible headers for persisted-view-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskHigh),
                  std::format(L"Failed to select task {} before persisted-view-state validation.", taskHigh));
    state.Require(waitForSnapshot(pane,
                                  [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not expose task {} as selected before persisted-view-state validation.", taskHigh));
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineTaskHeaderWidth     = std::max(0.0f, snapshot.taskHeaderRect.right - snapshot.taskHeaderRect.left);
    const size_t baselineVisibleColumnCount = snapshot.visibleWork.visibleColumnCount;
    const uint64_t baselineResizeCount      = snapshot.dxResizeCount;

    const LPARAM dragStart = DipPointToClientLParam(pane,
                                                    (snapshot.operationHeaderRect.left + snapshot.operationHeaderRect.right) * 0.5f,
                                                    (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f);
    const LPARAM dragTarget =
        DipPointToClientLParam(pane, snapshot.taskHeaderRect.left + 12.0f, (snapshot.operationHeaderRect.top + snapshot.operationHeaderRect.bottom) * 0.5f);
    SendMouseDragToResolvedPointWindow(pane, dragStart, dragTarget);

    FileOperationsIssuesPane::SelfTestSnapshot reordered{};
    state.Require(waitForSnapshot(pane,
                                  [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && ! value.hasActiveSort && value.selectionCount == 1u &&
               value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeFailureCount == 0u;
    },
                                  SelfTest::Scale(3000ms),
                                  reordered),
                  L"Issues pane did not expose reordered visible headers before persisted-view-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SendIssuesPaneTaskHeaderResizeDrag(pane, reordered, 56.0f),
                  L"Failed to drive a real issues-pane task-header resize drag before persisted-view-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsIssuesPane::SelfTestSnapshot resized{};
    state.Require(waitForSnapshot(pane,
                                  [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && currentTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                                  SelfTest::Scale(3000ms),
                                  resized),
                  L"Issues pane did not expose reordered+widened columns before persisted-view-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto clickTaskHeader = [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const LPARAM clickPoint = DipPointToClientLParam(
            pane, (value.taskHeaderRect.left + value.taskHeaderRect.right) * 0.5f, (value.taskHeaderRect.top + value.taskHeaderRect.bottom) * 0.5f);
        SendMessageW(pane, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(pane, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(pane, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    clickTaskHeader(resized);

    FileOperationsIssuesPane::SelfTestSnapshot ascending{};
    state.Require(waitForSnapshot(pane,
                                  [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.hasActiveSort && value.sortColumnIndex == 1u && ! value.sortDescending &&
               value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && currentTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                                  SelfTest::Scale(3000ms),
                                  ascending),
                  L"Issues pane did not keep reordered+widened columns stable through ascending sort before recreate.");
    if (! state.failure.empty())
    {
        return false;
    }

    clickTaskHeader(ascending);

    FileOperationsIssuesPane::SelfTestSnapshot descending{};
    state.Require(waitForSnapshot(pane,
                                  [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.hasActiveSort && value.sortColumnIndex == 1u && value.sortDescending && value.firstVisibleTaskId == taskHigh &&
               value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && currentTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f &&
               value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
               value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u;
    },
                                  SelfTest::Scale(3000ms),
                                  descending),
                  L"Issues pane did not keep reordered+widened columns stable through descending sort before recreate.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DestroyWindow(pane) != FALSE, L"Failed to destroy the issues pane before recreate validation.");
    state.Require(WaitForWindowClosed(pane, SelfTest::Scale(3000ms)), L"Issues pane did not close before recreate validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring persistedSortColumnId;
    bool persistedSortDescending = false;
    std::vector<Common::Settings::GridColumnLayoutEntry> persistedGridLayout;
    state.Require(fileOps->TryGetIssuesPaneViewState(persistedSortColumnId, persistedSortDescending, persistedGridLayout),
                  L"Issues pane did not persist any saved view state before recreate validation.");
    state.Require(requireUnrelatedFileOperationsSettingsIntact(L"after saving the issues-pane combined view state"),
                  L"Saving issues-pane view state should not clobber unrelated file-operations settings.");
    state.Require(persistedSortColumnId == L"task" && persistedSortDescending,
                  std::format(L"Issues pane persisted unexpected sort state before recreate validation. column='{}' descending={}",
                              persistedSortColumnId,
                              persistedSortDescending));
    const auto persistedTaskLayout = std::find_if(persistedGridLayout.begin(),
                                                  persistedGridLayout.end(),
                                                  [](const Common::Settings::GridColumnLayoutEntry& entry) noexcept { return entry.columnId == L"task"; });
    const auto persistedOperationLayout =
        std::find_if(persistedGridLayout.begin(), persistedGridLayout.end(), [](const Common::Settings::GridColumnLayoutEntry& entry) noexcept {
        return entry.columnId == L"operation";
    });
    state.Require(persistedTaskLayout != persistedGridLayout.end() && persistedOperationLayout != persistedGridLayout.end(),
                  L"Issues pane persisted view state did not include task/operation layout entries before recreate validation.");
    if (persistedTaskLayout != persistedGridLayout.end() && persistedOperationLayout != persistedGridLayout.end())
    {
        state.Require(persistedOperationLayout->displayIndex < persistedTaskLayout->displayIndex,
                      std::format(L"Issues pane persisted unexpected reordered layout before recreate validation. operationIndex={} taskIndex={}",
                                  persistedOperationLayout->displayIndex,
                                  persistedTaskLayout->displayIndex));
        state.Require(persistedTaskLayout->widthDip > baselineTaskHeaderWidth + 20.0f,
                      std::format(L"Issues pane persisted unexpected task width before recreate validation. savedWidth={:.2f} baselineWidth={:.2f}",
                                  persistedTaskLayout->widthDip,
                                  baselineTaskHeaderWidth));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(setIssuesVisible(true, L"issues-pane persisted-view-state recreate validation"),
                  L"Failed to reopen the issues pane for persisted-view-state validation.");
    const HWND reopenedPane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(reopenedPane != nullptr && IsWindow(reopenedPane) != FALSE, L"Issues pane did not reopen for persisted-view-state validation.");
    if (! reopenedPane || IsWindow(reopenedPane) == FALSE)
    {
        return false;
    }
    FileOperationsIssuesPane::SelfTestSnapshot reopenedReady{};
    state.Require(waitForSnapshotReady(reopenedPane, SelfTest::Scale(3000ms), reopenedReady),
                  std::format(L"Reopened issues pane did not expose a settled selftest snapshot before restore validation. "
                              L"handleChanged={} rowCount={} sortActive={} column={} descending={} visibleColumns={} taskWidth={:.2f} operationWidth={:.2f}",
                              reopenedPane != pane,
                              reopenedReady.rowCount,
                              reopenedReady.hasActiveSort,
                              reopenedReady.sortColumnIndex,
                              reopenedReady.sortDescending,
                              reopenedReady.visibleWork.visibleColumnCount,
                              std::max(0.0f, reopenedReady.taskHeaderRect.right - reopenedReady.taskHeaderRect.left),
                              std::max(0.0f, reopenedReady.operationHeaderRect.right - reopenedReady.operationHeaderRect.left)));
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(reopenedPane, true),
                  L"Failed to request reopened issues-pane refresh for persisted-view-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsIssuesPane::SelfTestSnapshot reopened{};
    state.Require(waitForSnapshot(reopenedPane,
                                  [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        const float currentTaskHeaderWidth = std::max(0.0f, value.taskHeaderRect.right - value.taskHeaderRect.left);
        return value.rowCount >= 2u && value.hasActiveSort && value.sortColumnIndex == 1u && value.sortDescending && value.firstVisibleTaskId == taskHigh &&
               value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left && currentTaskHeaderWidth > baselineTaskHeaderWidth + 20.0f &&
               value.visibleWork.visibleColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u;
    },
                                  SelfTest::Scale(3000ms),
                                  reopened),
                  std::format(L"Reopened issues pane did not restore the combined reordered, resized, and sorted view state. "
                              L"handleChanged={} rowCount={} refreshGeneration={} sortActive={} column={} descending={} firstTask={} "
                              L"taskWidth={:.2f} taskLeft={:.2f} operationLeft={:.2f} visibleColumns={} resizeFailures={}",
                              reopenedPane != pane,
                              reopened.rowCount,
                              reopened.refreshGeneration,
                              reopened.hasActiveSort,
                              reopened.sortColumnIndex,
                              reopened.sortDescending,
                              reopened.firstVisibleTaskId,
                              std::max(0.0f, reopened.taskHeaderRect.right - reopened.taskHeaderRect.left),
                              reopened.taskHeaderRect.left,
                              reopened.operationHeaderRect.left,
                              reopened.visibleWork.visibleColumnCount,
                              reopened.dxResizeFailureCount));

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(reopenedPane, taskHigh),
                  std::format(L"Failed to reselect task {} after issues-pane recreate.", taskHigh));
    state.Require(waitForSnapshot(reopenedPane,
                                  [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.selectionCount == 1u && value.primarySelectedTaskId == taskHigh && value.operationHeaderRect.left + 4.0f < value.taskHeaderRect.left &&
               value.hasActiveSort && value.sortColumnIndex == 1u && value.sortDescending && value.dxResizeFailureCount == 0u;
    },
                                  SelfTest::Scale(3000ms),
                                  reopened),
                  std::format(L"Reopened issues pane did not reselect task {} on the restored combined view state.", taskHigh));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestFocusGrid(reopenedPane),
                  L"Failed to focus the reopened issues-pane grid before restored combined-view-state copy validation.");
    ClearClipboardContents(reopenedPane);
    SendMessageW(reopenedPane, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(reopenedPane, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(reopenedPane, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(reopenedPane, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(reopenedPane);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(),
                  L"Reopened issues pane Ctrl+C should copy the restored combined reordered-resized-sorted visible row content after recreate.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t lineBreak       = copiedSelection.find_first_of(L"\r\n");
    const std::wstring firstLine = lineBreak == std::wstring::npos ? copiedSelection : copiedSelection.substr(0u, lineBreak);
    std::vector<std::wstring> columns;
    size_t start = 0u;
    while (start <= firstLine.size())
    {
        const size_t tab = firstLine.find(L'\t', start);
        if (tab == std::wstring::npos)
        {
            columns.push_back(firstLine.substr(start));
            break;
        }
        columns.push_back(firstLine.substr(start, tab - start));
        start = tab + 1u;
    }

    const std::wstring expectedTask = std::to_wstring(static_cast<unsigned long long>(taskHigh));
    state.Require(columns.size() >= 3u,
                  std::format(L"Reopened issues pane clipboard copy should export at least three visible columns after recreate, got {}.", columns.size()));
    state.Require(columns.size() >= 2u && columns[1] == expectedOperation,
                  std::format(L"Reopened issues pane clipboard copy should still place the restored reordered Operation column before Task after recreate. "
                              L"Expected second column '{}', got '{}'.",
                              expectedOperation,
                              columns.size() >= 2u ? columns[1] : std::wstring{}));
    state.Require(columns.size() >= 3u && columns[2] == expectedTask,
                  std::format(L"Reopened issues pane clipboard copy should keep selected task id '{}' immediately after the restored reordered Operation "
                              L"column after recreate. Got '{}'.",
                              expectedTask,
                              columns.size() >= 3u ? columns[2] : std::wstring{}));
    state.Require(requireUnrelatedFileOperationsSettingsIntact(L"after recreating the issues pane"),
                  L"Recreating the issues pane should not clobber unrelated file-operations settings.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneExposesLiveUiaSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane UIA validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File-operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });

    state.Require(setIssuesVisible(true, L"initial live UIA validation open"), L"Failed to show the file-operations issues pane for live UIA validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t taskA          = 0xF100000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const uint64_t taskB          = taskA + 1u;
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskA);
        fileOps->DebugRemoveDiagnosticsForTask(taskB);

        if (const HWND pane = getVisibleIssuesPane(); pane && IsWindow(pane) != FALSE)
        {
            static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
        }
    });

    const auto waitForSnapshot =
        [&](HWND pane, auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    fileOps->RecordTaskDiagnostic(taskA,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                                  L"selftest.dxui.issues",
                                  std::format(L"uia-selection-{}", taskA),
                                  L"C:\\selftest-uia-a.txt",
                                  L"D:\\selftest-uia-a.txt");
    fileOps->RecordTaskDiagnostic(taskB,
                                  FILESYSTEM_MOVE,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                  HRESULT_FROM_WIN32(ERROR_DISK_FULL),
                                  L"selftest.dxui.issues",
                                  std::format(L"uia-selection-{}", taskB),
                                  L"C:\\selftest-uia-b.txt",
                                  L"D:\\selftest-uia-b.txt");

    const auto runLiveUiaCycle = [&](std::wstring_view cycleLabel) noexcept -> bool
    {
        const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
        state.Require(pane != nullptr && IsWindow(pane) != FALSE,
                      std::format(L"File operations issues pane did not open for live UIA validation during {}.", cycleLabel));
        if (! pane || IsWindow(pane) == FALSE)
        {
            return false;
        }

        FileOperationsIssuesPane::SelfTestSnapshot initialSnapshot{};
        state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, initialSnapshot),
                      std::format(L"Failed to capture initial issues-pane snapshot for UIA validation during {}.", cycleLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        const uint64_t initialGeneration = initialSnapshot.refreshGeneration;
        state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, false),
                      std::format(L"Failed to request issues-pane refresh for UIA validation during {}.", cycleLabel));

        FileOperationsIssuesPane::SelfTestSnapshot refreshed{};
        state.Require(waitForSnapshot(pane,
                                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& snapshot) noexcept
        { return snapshot.refreshGeneration > initialGeneration && snapshot.rowCount >= 2u; },
                                      SelfTest::Scale(3000ms),
                                      refreshed),
                      std::format(L"Issues pane did not observe the diagnostic-driven refresh for UIA validation during {}.", cycleLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto requireIssuesSelection = [&](const uint64_t taskId, const std::wstring_view label) noexcept
        {
            state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId),
                          std::format(L"Failed to select issues-pane task {} during {} in {}.", taskId, label, cycleLabel));
            if (! state.failure.empty())
            {
                return false;
            }

            FileOperationsIssuesPane::SelfTestSnapshot selected{};
            state.Require(waitForSnapshot(pane,
                                          [&](const FileOperationsIssuesPane::SelfTestSnapshot& snapshot) noexcept
            { return snapshot.selectionCount == 1u && snapshot.primarySelectedTaskId == taskId; },
                                          SelfTest::Scale(3000ms),
                                          selected),
                          std::format(L"Issues pane did not report the expected selection for {} in {}.", label, cycleLabel));
            if (! state.failure.empty())
            {
                return false;
            }

            const auto selectionState = CollectVisibleDescendantSelectionPatternState(pane, UIA_DataGridControlTypeId);
            state.Require(
                selectionState.has_value(),
                std::format(L"Failed to collect live UI Automation selection state for the file-operations issues grid after {} in {}.", label, cycleLabel));
            if (! selectionState.has_value())
            {
                return false;
            }

            state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                          L"File-operations issues host should expose a UI Automation DataGrid control.");
            state.Require(selectionState->hasSelectionPattern, L"File-operations issues grid should expose SelectionPattern.");
            state.Require(selectionState->selectionCount == 1u,
                          std::format(L"File-operations issues grid should expose exactly one selected UIA row after {} in {}; saw {}.",
                                      label,
                                      cycleLabel,
                                      selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                          L"File-operations issues selected UIA row should expose the DataItem control type.");
            state.Require(selectionState->selectedHasSelectionItemPattern, L"File-operations issues selected UIA row should expose SelectionItemPattern.");
            state.Require(! selectionState->selectedName.empty(), L"File-operations issues selected UIA row should expose a non-empty accessible name.");
            state.Require(selectionState->selectedName.find(std::to_wstring(static_cast<unsigned long long>(taskId))) != std::wstring::npos,
                          std::format(L"File-operations issues selected UIA row name '{}' should include the selected task id '{}' after {} in {}.",
                                      selectionState->selectedName,
                                      std::to_wstring(static_cast<unsigned long long>(taskId)),
                                      label,
                                      cycleLabel));
            return state.failure.empty();
        };

        state.Require(requireIssuesSelection(taskA, L"selecting the first issues task"),
                      std::format(L"File-operations issues grid should keep UIA selection synchronized for the first injected task during {}.", cycleLabel));
        state.Require(requireIssuesSelection(taskB, L"switching to the second issues task"),
                      std::format(L"File-operations issues grid should keep UIA selection synchronized when switching to the second injected task during {}.",
                                  cycleLabel));
        state.Require(
            requireIssuesSelection(taskA, L"switching back to the first issues task"),
            std::format(L"File-operations issues grid should keep UIA selection synchronized when switching back to the first injected task during {}.",
                        cycleLabel));
        return state.failure.empty();
    };

    state.Require(runLiveUiaCycle(L"the initial issues-pane UIA pass"),
                  L"File-operations issues pane should expose the expected live UIA selection state on the initial pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(setIssuesVisible(false, L"issues-pane live UIA close between passes"),
                  L"Failed to hide the file-operations issues pane between live UIA validation passes.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < closeDeadline)
    {
        PumpPendingMessages();
        if (getVisibleIssuesPane() == nullptr)
        {
            break;
        }

        std::this_thread::sleep_for(20ms);
    }

    state.Require(getVisibleIssuesPane() == nullptr, L"File-operations issues pane should fully close between live UIA validation passes.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(setIssuesVisible(true, L"reopened live UIA validation pass"),
                  L"Failed to reopen the file-operations issues pane for the second live UIA validation pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runLiveUiaCycle(L"the reopened issues-pane UIA pass"),
                  L"File-operations issues pane should expose the same live UIA selection state after reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneTabTraversalKeepsGridFocus(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane keyboard-focus validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File-operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });

    const uint64_t taskA          = 0xF300000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const uint64_t taskB          = taskA + 1u;
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskA);
        fileOps->DebugRemoveDiagnosticsForTask(taskB);
        if (const HWND pane = getVisibleIssuesPane(); pane && IsWindow(pane) != FALSE)
        {
            static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
        }
    });

    fileOps->RecordTaskDiagnostic(taskA,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Warning,
                                  HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                                  L"selftest.dxui.issues.tab",
                                  std::format(L"tab-focus-a-{}", taskA),
                                  L"C:\\selftest-issues-tab-a.txt",
                                  L"D:\\selftest-issues-tab-a.txt");
    fileOps->RecordTaskDiagnostic(taskB,
                                  FILESYSTEM_MOVE,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                  HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION),
                                  L"selftest.dxui.issues.tab",
                                  std::format(L"tab-focus-b-{}", taskB),
                                  L"C:\\selftest-issues-tab-b.txt",
                                  L"D:\\selftest-issues-tab-b.txt");

    state.Require(setIssuesVisible(true, L"issues-pane keyboard-focus validation open"),
                  L"Failed to show the file-operations issues pane for keyboard-focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File-operations issues pane did not open for keyboard-focus validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    state.Require(WindowExposesUiaProvider(pane), L"File-operations issues pane should answer WM_GETOBJECT during keyboard-focus validation.");
    state.Require(CountVisibleChildWindows(pane) == 0u,
                  L"File-operations issues pane should not expose visible child fallback during keyboard-focus validation.");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, true), L"Failed to request issues-pane refresh before keyboard-focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForSnapshot = [&](auto&& predicate, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.rowCount >= 2u && value.visibleWork.visibleRowCount > 0u && value.visibleWork.visibleColumnCount > 0u &&
               value.visibleWork.visibleCellCount > 0u && value.dxResizeFailureCount == 0u;
    },
                      snapshot),
                  L"File-operations issues pane did not settle onto a visible DX grid before keyboard-focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskA),
                  std::format(L"Failed to select issues-pane task {} before keyboard-focus validation.", taskA));
    state.Require(FileOperationsIssuesPane::SelfTestFocusGrid(pane), L"Failed to focus the issues-pane DX grid before keyboard-focus validation.");
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.selectionCount == 1u && value.primarySelectedTaskId == taskA && value.gridFocused && value.dxResizeFailureCount == 0u; },
                                  snapshot),
                  L"File-operations issues pane DX grid did not take focus with the expected selection before keyboard-focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineSelectedRowId    = snapshot.primarySelectedRowId;
    const uint64_t baselineResizeCount      = snapshot.dxResizeCount;
    const size_t baselineVisibleRowCount    = snapshot.visibleWork.visibleRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleWork.visibleColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleWork.visibleCellCount;

    const auto sendTab = [&](const bool reverse, std::wstring_view label) noexcept
    {
        if (reverse)
        {
            SendMessageW(pane, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(pane, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(pane, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(pane, WM_KEYUP, VK_SHIFT, 0);
        }

        state.Require(waitForSnapshot(
                          [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        {
            return value.selectionCount == 1u && value.primarySelectedTaskId == taskA && value.primarySelectedRowId == baselineSelectedRowId &&
                   value.gridFocused && value.dxResizeCount == baselineResizeCount && value.dxResizeFailureCount == 0u &&
                   value.visibleWork.visibleRowCount == baselineVisibleRowCount && value.visibleWork.visibleColumnCount == baselineVisibleColumnCount &&
                   value.visibleWork.visibleCellCount == baselineVisibleCellCount;
        },
                          snapshot),
                      std::format(L"File-operations issues pane did not keep grid focus and selection stable after {}.", label));
    };

    sendTab(false, L"forward Tab");
    sendTab(true, L"reverse Shift+Tab");
    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneLongRunScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane long-run scrolling validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible        = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto restoreVisibility = wil::scope_exit([&] noexcept
    {
        const bool visibleNow = g_folderWindow.IsFileOperationsIssuesPaneVisible();
        if (visibleNow != wasVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline && g_folderWindow.IsFileOperationsIssuesPaneVisible() != wasVisible)
            {
                PumpPendingMessages();
                std::this_thread::sleep_for(20ms);
            }
        }
    });

    if (! wasVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for long-run scrolling validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    constexpr size_t kTaskCount   = 48u;
    const uint64_t taskBase       = 0xF200000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        for (size_t index = 0; index < kTaskCount; ++index)
        {
            fileOps->DebugRemoveDiagnosticsForTask(taskBase + index);
        }
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    FileOperationsIssuesPane::SelfTestSnapshot initialSnapshot{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, initialSnapshot),
                  L"Failed to capture initial issues-pane snapshot for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t initialGeneration = initialSnapshot.refreshGeneration;
    const size_t initialRowCount     = initialSnapshot.rowCount;
    for (size_t index = 0; index < kTaskCount; ++index)
    {
        const uint64_t taskId = taskBase + index;
        fileOps->RecordTaskDiagnostic(taskId,
                                      (index % 2u) == 0u ? FILESYSTEM_COPY : FILESYSTEM_MOVE,
                                      (index % 3u) == 0u ? FolderWindow::FileOperationState::DiagnosticSeverity::Warning
                                                         : FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                      HRESULT_FROM_WIN32((index % 2u) == 0u ? ERROR_ACCESS_DENIED : ERROR_DISK_FULL),
                                      L"selftest.dxui.issues.longrun",
                                      std::format(L"long-run-scroll-{}", taskId),
                                      std::format(L"C:\\selftest-longrun-source-{:02}.txt", index),
                                      std::format(L"D:\\selftest-longrun-dest-{:02}.txt", index));
    }

    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, false), L"Failed to request issues-pane refresh for long-run scrolling validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.refreshGeneration > initialGeneration && value.rowCount >= initialRowCount + kTaskCount; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  L"Issues pane did not populate enough rows for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.visibleWork.visibleRowCount > 0u, L"Issues pane should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.visibleWork.visibleColumnCount > 0u, L"Issues pane should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.visibleWork.visibleRowCount < snapshot.rowCount,
                  std::format(L"Issues pane should stay virtualized during long-run scrolling validation; visible rows={} total rows={}.",
                              snapshot.visibleWork.visibleRowCount,
                              snapshot.rowCount));
    state.Require(snapshot.visibleWork.hasVerticalScrollbar, L"Issues pane should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.dxResizeFailureCount == 0u,
                  std::format(L"Issues pane should start with zero DX resize failures; saw {}.", snapshot.dxResizeFailureCount));

    const uint64_t initialVisibleRows  = snapshot.visibleWork.visibleRowCount;
    const size_t initialVisibleColumns = snapshot.visibleWork.visibleColumnCount;
    const uint64_t initialResizeCount  = snapshot.dxResizeCount;
    state.Require(FileOperationsIssuesPane::SelfTestScrollByWheelDetents(pane, 120),
                  L"Issues pane did not accept top reset before long-run scrolling validation.");
    RedrawWindow(pane, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    state.Require(waitForSnapshot([](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    { return value.visibleWork.verticalScrollDip <= 0.5f; },
                                  SelfTest::Scale(3000ms),
                                  snapshot),
                  std::format(L"Issues pane did not reset to the top before long-run scrolling validation; scrollDip={}.",
                              snapshot.visibleWork.verticalScrollDip));
    if (! state.failure.empty())
    {
        return false;
    }

    uint64_t previousRenderCount = snapshot.dxRenderCount;

    for (size_t chunk = 0; chunk < 8u; ++chunk)
    {
        const int detents              = (chunk % 2u) == 0u ? -6 : 6;
        const float previousScrollDip  = snapshot.visibleWork.verticalScrollDip;
        const uint64_t beforeRender    = previousRenderCount;
        const bool expectedScrollDown  = detents < 0;
        const auto scrollMovedExpected = [expectedScrollDown, previousScrollDip](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        {
            const float currentScrollDip = value.visibleWork.verticalScrollDip;
            return expectedScrollDown ? currentScrollDip > previousScrollDip + 0.5f : currentScrollDip + 0.5f < previousScrollDip;
        };

        state.Require(FileOperationsIssuesPane::SelfTestScrollByWheelDetents(pane, detents),
                      std::format(L"Issues pane did not accept long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }
        RedrawWindow(pane, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);

        state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        { return scrollMovedExpected(value) && value.dxRenderCount > beforeRender; },
                                      SelfTest::Scale(3000ms),
                                      snapshot),
                      std::format(L"Issues pane did not move and repaint after long-run scroll chunk {}; detents={}, beforeScrollDip={}, afterScrollDip={}, "
                                  L"beforeRender={}, afterRender={}, rows={}, visibleRows={}, verticalScrollbar={}.",
                                  chunk,
                                  detents,
                                  previousScrollDip,
                                  snapshot.visibleWork.verticalScrollDip,
                                  beforeRender,
                                  snapshot.dxRenderCount,
                                  snapshot.rowCount,
                                  snapshot.visibleWork.visibleRowCount,
                                  snapshot.visibleWork.hasVerticalScrollbar));
        if (! state.failure.empty())
        {
            return false;
        }

        previousRenderCount = snapshot.dxRenderCount;
        state.Require(snapshot.rowCount >= initialRowCount + kTaskCount,
                      std::format(L"Issues pane lost rows during long-run scroll chunk {}; saw {}.", chunk, snapshot.rowCount));
        state.Require(snapshot.visibleWork.visibleRowCount > 0u && snapshot.visibleWork.visibleRowCount <= initialVisibleRows + 1u,
                      std::format(L"Issues pane visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleWork.visibleRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.visibleWork.visibleColumnCount == initialVisibleColumns,
                      std::format(L"Issues pane visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleWork.visibleColumnCount,
                                  initialVisibleColumns));
        state.Require(snapshot.visibleWork.visibleCellCount <= snapshot.visibleWork.visibleRowCount * snapshot.visibleWork.visibleColumnCount,
                      std::format(L"Issues pane visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                                  chunk,
                                  snapshot.visibleWork.visibleCellCount,
                                  snapshot.visibleWork.visibleRowCount,
                                  snapshot.visibleWork.visibleColumnCount));
        state.Require(snapshot.visibleWork.hasVerticalScrollbar,
                      std::format(L"Issues pane lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(
            snapshot.dxResizeCount == initialResizeCount,
            std::format(
                L"Issues pane churned DX host resizes during chunk {}; resize count moved from {} to {}.", chunk, initialResizeCount, snapshot.dxResizeCount));
        state.Require(snapshot.dxResizeFailureCount == 0u,
                      std::format(L"Issues pane hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.dxResizeFailureCount));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane long-run open/close validation.");
    if (! fileOps)
    {
        return false;
    }

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible       = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto setIssuesVisible = [&](bool visible, std::wstring_view label) noexcept -> bool
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
        {
            return true;
        }

        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.IsFileOperationsIssuesPaneVisible() == visible)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        state.Require(false, std::format(L"File-operations issues pane did not switch to {} during {}.", visible ? L"visible" : L"hidden", label));
        return false;
    };
    const auto restoreVisibility = wil::scope_exit([&] noexcept { static_cast<void>(setIssuesVisible(wasVisible, L"restore")); });

    constexpr size_t kTaskCount   = 4u;
    const uint64_t taskBase       = 0xF240000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        for (size_t index = 0; index < kTaskCount; ++index)
        {
            fileOps->DebugRemoveDiagnosticsForTask(taskBase + index);
        }

        if (const HWND pane = getVisibleIssuesPane(); pane && IsWindow(pane) != FALSE)
        {
            static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
        }
    });

    const std::array<std::wstring, kTaskCount> labels{
        L"issues-open-close-0",
        L"issues-open-close-1",
        L"issues-open-close-2",
        L"issues-open-close-3",
    };

    for (size_t index = 0; index < kTaskCount; ++index)
    {
        const uint64_t taskId = taskBase + index;
        fileOps->RecordTaskDiagnostic(taskId,
                                      (index % 2u) == 0u ? FILESYSTEM_COPY : FILESYSTEM_MOVE,
                                      (index % 3u) == 0u ? FolderWindow::FileOperationState::DiagnosticSeverity::Warning
                                                         : FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                      HRESULT_FROM_WIN32((index % 2u) == 0u ? ERROR_ACCESS_DENIED : ERROR_DISK_FULL),
                                      L"selftest.dxui.issues.openclose",
                                      labels[index],
                                      std::format(L"C:\\selftest-openclose-source-{:02}.txt", index),
                                      std::format(L"D:\\selftest-openclose-dest-{:02}.txt", index));
    }

    state.Require(setIssuesVisible(false, L"initial hide"), L"Failed to hide issues pane before open/close churn validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const uint64_t expectedTaskId         = taskBase + (cycle % kTaskCount);
        const std::wstring expectedTaskIdText = std::to_wstring(static_cast<unsigned long long>(expectedTaskId));

        state.Require(setIssuesVisible(true, std::format(L"open cycle {}", cycle)), std::format(L"Failed to show the issues pane during cycle {}.", cycle));
        const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
        state.Require(pane != nullptr && IsWindow(pane) != FALSE, std::format(L"File-operations issues pane did not appear during cycle {}.", cycle));
        if (! pane || IsWindow(pane) == FALSE)
        {
            return false;
        }

        state.Require(WindowExposesUiaProvider(pane), std::format(L"File-operations issues pane should answer WM_GETOBJECT during cycle {}.", cycle));
        state.Require(CountVisibleChildWindows(pane) == 0u,
                      std::format(L"File-operations issues pane should not expose visible child-control fallback during cycle {}.", cycle));
        state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, false), std::format(L"Failed to refresh the issues pane during cycle {}.", cycle));

        const auto waitForSnapshot = [&](auto&& predicate, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
        };

        FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
        state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        { return value.rowCount >= kTaskCount && value.visibleWork.visibleRowCount > 0u; },
                                      snapshot),
                      std::format(L"File-operations issues pane did not repopulate rows during cycle {}.", cycle));
        state.Require(snapshot.visibleWork.visibleColumnCount > 0u,
                      std::format(L"File-operations issues pane should expose visible columns during cycle {}.", cycle));
        state.Require(snapshot.dxResizeFailureCount == 0u,
                      std::format(L"File-operations issues pane hit DX resize failures during cycle {}; saw {}.", cycle, snapshot.dxResizeFailureCount));

        state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, expectedTaskId),
                      std::format(L"Failed to select issues-pane task {} during cycle {}.", expectedTaskId, cycle));
        state.Require(waitForSnapshot([&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        { return value.selectionCount == 1u && value.primarySelectedTaskId == expectedTaskId; },
                                      snapshot),
                      std::format(L"File-operations issues pane did not expose the selected task during cycle {}.", cycle));

        const auto selectionState = CollectVisibleDescendantSelectionPatternState(pane, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(), std::format(L"Failed to collect UI Automation selection state for the issues pane during cycle {}.", cycle));
        if (selectionState.has_value())
        {
            state.Require(selectionState->hasSelectionPattern,
                          std::format(L"File-operations issues pane should expose SelectionPattern during cycle {}.", cycle));
            state.Require(selectionState->selectionCount == 1u,
                          std::format(L"File-operations issues pane should keep exactly one selected UIA row during cycle {}; saw {}.",
                                      cycle,
                                      selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                          std::format(L"File-operations issues pane selected UIA row should remain a DataItem during cycle {}.", cycle));
            state.Require(selectionState->selectedHasSelectionItemPattern,
                          std::format(L"File-operations issues pane selected UIA row should keep SelectionItemPattern during cycle {}.", cycle));
            state.Require(selectionState->selectedName.find(expectedTaskIdText) != std::wstring::npos,
                          std::format(L"File-operations issues pane selected UIA row name should track task {} during cycle {}; saw '{}'.",
                                      expectedTaskIdText,
                                      cycle,
                                      selectionState->selectedName));
        }

        state.Require(setIssuesVisible(false, std::format(L"close cycle {}", cycle)), std::format(L"Failed to hide the issues pane during cycle {}.", cycle));
        const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < closeDeadline)
        {
            PumpPendingMessages();
            if (getVisibleIssuesPane() == nullptr)
            {
                break;
            }
            std::this_thread::sleep_for(20ms);
        }
        state.Require(getVisibleIssuesPane() == nullptr,
                      std::format(L"File-operations issues pane should not remain visible after close during cycle {}.", cycle));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsIssuesPaneThemeCycleKeepsGridLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for issues-pane theme-cycle validation.");
    if (! fileOps)
    {
        return false;
    }

    const AppTheme originalTheme = g_folderWindow.GetTheme();
    const auto restoreTheme      = wil::scope_exit([&] noexcept { g_folderWindow.ApplyTheme(originalTheme); });

    constexpr wchar_t kIssuesPaneClassName[] = L"RedSalamander.FileOperationsIssuesPane";
    const auto getVisibleIssuesPane          = [className = kIssuesPaneClassName]() noexcept -> HWND
    {
        const HWND pane = FindWindowW(className, nullptr);
        return (pane && IsActuallyVisibleChildWindow(pane)) ? pane : nullptr;
    };

    const bool wasVisible        = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto restoreVisibility = wil::scope_exit([&] noexcept
    {
        const bool visibleNow = g_folderWindow.IsFileOperationsIssuesPaneVisible();
        if (visibleNow != wasVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline && g_folderWindow.IsFileOperationsIssuesPaneVisible() != wasVisible)
            {
                PumpPendingMessages();
                std::this_thread::sleep_for(20ms);
            }
        }
    });

    if (! wasVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    }

    const HWND pane = WaitForWindow(getVisibleIssuesPane, SelfTest::Scale(5000ms));
    state.Require(pane != nullptr && IsWindow(pane) != FALSE, L"File operations issues pane did not open for theme-cycle validation.");
    if (! pane || IsWindow(pane) == FALSE)
    {
        return false;
    }

    const uint64_t taskId         = 0xF300000000000000ull + static_cast<uint64_t>(GetTickCount64() & 0x0000FFFFFFFFFFFFull);
    const auto cleanupDiagnostics = wil::scope_exit([&] noexcept
    {
        fileOps->DebugRemoveDiagnosticsForTask(taskId);
        static_cast<void>(FileOperationsIssuesPane::SelfTestRefresh(pane, true));
    });

    const auto waitForSnapshot = [&](auto&& predicate, std::chrono::milliseconds timeout, FileOperationsIssuesPane::SelfTestSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, outSnapshot) && predicate(outSnapshot);
    };

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"fileops-issues-selftest-theme-cycle-initial");
    g_folderWindow.ApplyTheme(initialTheme);

    FileOperationsIssuesPane::SelfTestSnapshot initialSnapshot{};
    state.Require(FileOperationsIssuesPane::TryGetSelfTestSnapshot(pane, initialSnapshot),
                  L"Failed to capture initial issues-pane snapshot for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t initialGeneration = initialSnapshot.refreshGeneration;
    fileOps->RecordTaskDiagnostic(taskId,
                                  FILESYSTEM_COPY,
                                  FolderWindow::FileOperationState::DiagnosticSeverity::Error,
                                  HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                                  L"selftest.dxui.issues",
                                  std::format(L"theme-cycle-{}", taskId),
                                  L"C:\\selftest-theme-cycle-source.txt",
                                  L"D:\\selftest-theme-cycle-dest.txt");
    state.Require(FileOperationsIssuesPane::SelfTestRefresh(pane, false), L"Failed to request issues-pane refresh for theme-cycle validation.");

    FileOperationsIssuesPane::SelfTestSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.refreshGeneration > initialGeneration && value.rowCount >= 1u && value.themeDark == initialTheme.dark &&
               value.themeHighContrast == initialTheme.highContrast && value.themeRainbow == initialTheme.menu.rainbowMode;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not observe the diagnostic-driven refresh for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FileOperationsIssuesPane::SelfTestSelectTask(pane, taskId),
                  std::format(L"Failed to select issues-pane task {} for theme-cycle validation.", taskId));
    state.Require(waitForSnapshot(
                      [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
    {
        return value.selectionCount == 1u && value.primarySelectedTaskId == taskId && value.selectedIssueRowFillArgb != 0u &&
               value.selectedIssueRowTextArgb != 0u;
    },
                      SelfTest::Scale(3000ms),
                      snapshot),
                  L"Issues pane did not expose a selected row for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto unpackColor = [](uint32_t argb) noexcept
    {
        return D2D1::ColorF(static_cast<float>((argb >> 16u) & 0xFFu) / 255.0f,
                            static_cast<float>((argb >> 8u) & 0xFFu) / 255.0f,
                            static_cast<float>(argb & 0xFFu) / 255.0f,
                            static_cast<float>((argb >> 24u) & 0xFFu) / 255.0f);
    };
    const auto luminance = [&](uint32_t argb) noexcept
    {
        const D2D1_COLOR_F color = unpackColor(argb);
        const auto linearize = [](float channel) noexcept { return (channel <= 0.03928f) ? (channel / 12.92f) : std::pow((channel + 0.055f) / 1.055f, 2.4f); };

        return (0.2126f * linearize(color.r)) + (0.7152f * linearize(color.g)) + (0.0722f * linearize(color.b));
    };
    const auto contrastRatio = [&](uint32_t a, uint32_t b) noexcept
    {
        const float lumA    = luminance(a);
        const float lumB    = luminance(b);
        const float lighter = (std::max)(lumA, lumB);
        const float darker  = (std::min)(lumA, lumB);
        return (lighter + 0.05f) / (darker + 0.05f);
    };

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        const uint64_t previousRenderCount = snapshot.dxRenderCount;
        g_folderWindow.ApplyTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const FileOperationsIssuesPane::SelfTestSnapshot& value) noexcept
        {
            return value.selectionCount == 1u && value.primarySelectedTaskId == taskId && value.themeDark == theme.dark &&
                   value.themeHighContrast == theme.highContrast && value.themeRainbow == theme.menu.rainbowMode && value.selectedIssueRowFillArgb != 0u &&
                   value.selectedIssueRowTextArgb != 0u && value.dxRenderCount >= previousRenderCount;
        },
                          SelfTest::Scale(3000ms),
                          snapshot),
                      std::format(L"Issues pane did not settle after switching to {} theme.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(IsWindow(pane) != FALSE, std::format(L"Issues pane did not survive the {} theme update.", label));
        state.Require(snapshot.selectedIssueRowUsesRainbow == expectRainbow,
                      std::format(L"Issues pane selected-row rainbow state mismatch after {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast, std::format(L"Issues pane high-contrast state mismatch after {} theme update.", label));
        state.Require(snapshot.selectedIssueRowFillArgb != snapshot.selectedIssueRowTextArgb,
                      std::format(L"Issues pane selected-row colors collapsed to the same value after {} theme update.", label));

        const float minimumContrast = expectHighContrast ? 4.5f : 3.0f;
        state.Require(contrastRatio(snapshot.selectedIssueRowFillArgb, snapshot.selectedIssueRowTextArgb) >= minimumContrast,
                      std::format(L"Issues pane selected-row text contrast dropped below {:.1f}:1 after {} theme update.", minimumContrast, label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"fileops-issues-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"fileops-issues-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"fileops-issues-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"fileops-issues-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

template <typename WorkerFunc>
bool RunFileOperationsSpeedLimitPromptModalCycle(HWND popup,
                                                 const FileOperationsPopupInternal::PopupSelfTestInvoke& openPrompt,
                                                 WorkerFunc&& workerFunc) noexcept
{
    using namespace std::chrono_literals;

    std::atomic<bool> workerDone   = false;
    std::atomic<HWND> promptHandle = nullptr;

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow([]() noexcept { return GetFileOperationsSpeedLimitPromptHandle(); }, SelfTest::Scale(5000ms));
        promptHandle.store(prompt, std::memory_order_release);
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                if (! DebugCancelFileOperationsSpeedLimitPrompt())
                {
                    PostMessageW(prompt, WM_CLOSE, 0, 0);
                }
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
            }
            workerDone.store(true, std::memory_order_release);
        });

        workerFunc(prompt);
    });

    const bool invoked = DebugInvokeFileOperationsPopup(popup, openPrompt);
    if (! invoked)
    {
        worker.join();
        return false;
    }

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(15000ms);
    while (! workerDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }

    if (! workerDone.load(std::memory_order_acquire))
    {
        if (const HWND prompt = promptHandle.load(std::memory_order_acquire); prompt && IsWindow(prompt) != FALSE)
        {
            if (! DebugCancelFileOperationsSpeedLimitPrompt())
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
            }
        }
    }

    worker.join();
    return invoked;
}

[[nodiscard]] bool CloseFileOperationsPopupForSelfTest(FolderWindow::FileOperationState* fileOps) noexcept
{
    using namespace std::chrono_literals;

    if (! fileOps)
    {
        return true;
    }

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        fileOps->CancelAll();

        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completed;
        fileOps->CollectCompletedTasks(completed);
        for (const auto& summary : completed)
        {
            fileOps->DismissCompletedTask(summary.taskId);
        }

        std::vector<FolderWindow::FileOperationState::Task*> tasks;
        fileOps->CollectTasks(tasks);
        const bool hasRemainingTasks = std::ranges::any_of(tasks, [](const auto* task) noexcept { return task != nullptr; });

        const HWND popup = fileOps->GetPopupHwndForSelfTest();
        if ((! popup || IsWindow(popup) == FALSE) && ! hasRemainingTasks)
        {
            return true;
        }

        if (popup && IsWindow(popup) != FALSE && ! hasRemainingTasks && ! fileOps->HasActiveOperations())
        {
            PostMessageW(popup, WM_CLOSE, 0, 0);
        }

        std::this_thread::sleep_for(20ms);
    }

    const HWND popup = fileOps->GetPopupHwndForSelfTest();
    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    fileOps->CollectTasks(tasks);
    const bool hasRemainingTasks = std::ranges::any_of(tasks, [](const auto* task) noexcept { return task != nullptr; });
    return (popup == nullptr || IsWindow(popup) == FALSE) && ! hasRemainingTasks;
}

template <typename Duration>
[[nodiscard]] std::optional<uint64_t> ResolveNewFileOperationsTaskIdForSelfTest(FolderWindow::FileOperationState* fileOps,
                                                                                const std::unordered_set<uint64_t>& existingTaskIds,
                                                                                Duration timeout) noexcept
{
    using namespace std::chrono_literals;

    if (! fileOps)
    {
        return std::nullopt;
    }

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        std::vector<FolderWindow::FileOperationState::Task*> tasks;
        fileOps->CollectTasks(tasks);
        for (auto* task : tasks)
        {
            if (task && ! existingTaskIds.contains(task->GetId()))
            {
                return task->GetId();
            }
        }

        if (const HWND popup = fileOps->GetPopupHwndForSelfTest(); popup && IsWindow(popup) != FALSE)
        {
            FileOperationsPopupInternal::TaskSnapshot popupTaskSnapshot{};
            if (DebugGetFileOperationsPopupTaskSnapshot(popup, 0, popupTaskSnapshot) &&
                popupTaskSnapshot.kind == FileOperationsPopupInternal::TaskSnapshot::Kind::FileOperation && popupTaskSnapshot.taskId != 0 &&
                ! popupTaskSnapshot.finished && ! existingTaskIds.contains(popupTaskSnapshot.taskId))
            {
                return popupTaskSnapshot.taskId;
            }
        }

        std::this_thread::sleep_for(20ms);
    }

    return std::nullopt;
}

[[nodiscard]] bool RefreshTrackedFileOperationsTaskForSelfTest(FolderWindow::FileOperationState* fileOps,
                                                               HWND popup,
                                                               std::optional<uint64_t>& taskId,
                                                               FileOperationsPopupInternal::TaskSnapshot* snapshotOut = nullptr) noexcept
{
    using namespace std::chrono_literals;

    if (! fileOps)
    {
        return false;
    }

    const auto tryResolveLiveTask = [&](const uint64_t requestedTaskId, FileOperationsPopupInternal::TaskSnapshot* resolvedSnapshot = nullptr) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
        do
        {
            if (auto* task = fileOps->FindTask(requestedTaskId); task)
            {
                taskId = requestedTaskId;
                if (resolvedSnapshot)
                {
                    *resolvedSnapshot = {};
                    static_cast<void>(DebugGetFileOperationsPopupTaskSnapshot(popup, requestedTaskId, *resolvedSnapshot));
                }
                return true;
            }

            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);
        } while (std::chrono::steady_clock::now() < deadline);

        return false;
    };

    if (taskId.has_value() && tryResolveLiveTask(taskId.value(), snapshotOut))
    {
        return true;
    }

    auto trySnapshot = [&](const uint64_t requestedTaskId) noexcept
    {
        FileOperationsPopupInternal::TaskSnapshot snapshot{};
        if (! popup || IsWindow(popup) == FALSE || ! DebugGetFileOperationsPopupTaskSnapshot(popup, requestedTaskId, snapshot) ||
            snapshot.kind != FileOperationsPopupInternal::TaskSnapshot::Kind::FileOperation || snapshot.taskId == 0 || snapshot.finished)
        {
            return false;
        }

        if (! tryResolveLiveTask(snapshot.taskId))
        {
            return false;
        }

        taskId = snapshot.taskId;
        if (snapshotOut)
        {
            *snapshotOut = snapshot;
        }
        return true;
    };

    if (trySnapshot(taskId.value_or(0)))
    {
        return true;
    }

    if (taskId.has_value() && trySnapshot(0))
    {
        return true;
    }

    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    fileOps->CollectTasks(tasks);
    const auto it = std::find_if(tasks.begin(), tasks.end(), [](const auto* task) noexcept { return task != nullptr; });
    if (it == tasks.end())
    {
        return false;
    }

    taskId = (*it)->GetId();
    if (snapshotOut)
    {
        *snapshotOut = {};
        static_cast<void>(DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value(), *snapshotOut));
    }
    return true;
}

[[nodiscard]] std::optional<uint64_t> ResolveAndPauseNewFileOperationsTaskForSelfTest(FolderWindow::FileOperationState* fileOps,
                                                                                      const std::unordered_set<uint64_t>& existingTaskIds,
                                                                                      uint64_t desiredSpeedLimitBytesPerSecond) noexcept
{
    if (! fileOps)
    {
        return std::nullopt;
    }

    std::vector<FolderWindow::FileOperationState::Task*> tasks;
    fileOps->CollectTasks(tasks);
    for (auto* task : tasks)
    {
        if (! task || existingTaskIds.contains(task->GetId()))
        {
            continue;
        }

        task->SetDesiredSpeedLimit(desiredSpeedLimitBytesPerSecond);
        if (! task->IsPaused())
        {
            task->TogglePause();
        }
        return task->GetId();
    }

    return std::nullopt;
}

constexpr std::wstring_view kBuiltinDummyFileSystemIdForFileOpsPrompt = L"builtin/file-system-dummy";

[[nodiscard]] bool CreateFileSystemIoForFileOpsPrompt(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemIO>& outIo) noexcept
{
    outIo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemIO), outIo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outIo);
}

[[nodiscard]] bool SetPluginConfigurationForFileOpsPrompt(IInformations* info, std::string_view configUtf8) noexcept
{
    if (! info)
    {
        return false;
    }

    std::string owned(configUtf8);
    owned.push_back('\0');
    return SUCCEEDED(info->SetConfiguration(owned.c_str()));
}

[[nodiscard]] bool BackupPluginConfigurationForFileOpsPrompt(IInformations* info, std::string& outConfigUtf8) noexcept
{
    if (! info)
    {
        return false;
    }

    const char* config = nullptr;
    const HRESULT hr   = info->GetConfiguration(&config);
    if (FAILED(hr) || ! config)
    {
        return false;
    }

    outConfigUtf8 = config;
    return true;
}

[[nodiscard]] bool EnsureDummyFolderExistsForFileOpsPrompt(IFileSystem* fs, std::wstring_view destinationFolder) noexcept
{
    if (! fs || destinationFolder.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());
    if (FAILED(hr) || ! dirOps)
    {
        return false;
    }

    const std::filesystem::path normalized = std::filesystem::path(destinationFolder).lexically_normal();
    std::filesystem::path current          = normalized.root_path();
    const std::filesystem::path relative   = normalized.relative_path();
    if (relative.empty())
    {
        return true;
    }

    for (const auto& part : relative)
    {
        if (part.empty() || part == L".")
        {
            continue;
        }

        current /= part;
        const std::wstring currentText = current.generic_wstring();
        const HRESULT createHr         = dirOps->CreateDirectory(currentText.c_str());
        if (FAILED(createHr) && createHr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] std::wstring ToPluginPathTextForFileOpsPrompt(const std::filesystem::path& path) noexcept
{
    return path.generic_wstring();
}

[[nodiscard]] std::wstring SanitizeDummyPathSegmentForFileOpsPrompt(std::wstring text) noexcept
{
    std::erase_if(text, [](wchar_t ch) noexcept { return ! (std::iswalnum(ch) != 0 || ch == L'-' || ch == L'_'); });
    return text;
}

[[nodiscard]] bool WritePatternFileFsIoForFileOpsPrompt(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, uint64_t sizeBytes) noexcept
{
    if (! io)
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const std::wstring pathText = ToPluginPathTextForFileOpsPrompt(path);
    const HRESULT createHr      = io->CreateFileWriter(pathText.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    std::array<unsigned char, 256 * 1024> buffer{};
    for (size_t i = 0; i < buffer.size(); ++i)
    {
        buffer[i] = static_cast<unsigned char>((i * 131u) ^ 0x5Au);
    }

    uint64_t remaining = sizeBytes;
    while (remaining > 0)
    {
        const unsigned long chunk = static_cast<unsigned long>((std::min<uint64_t>)(remaining, static_cast<uint64_t>(buffer.size())));
        unsigned long written     = 0;
        const HRESULT writeHr     = writer->Write(buffer.data(), chunk, &written);
        if (FAILED(writeHr) || written != chunk)
        {
            return false;
        }

        remaining -= chunk;
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool TestFileOperationsSpeedLimitPromptUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr uint64_t kInitialLimitBytesPerSecond = 64ull * 1024ull;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for speed-limit prompt validation.");
    if (! fileOps)
    {
        return false;
    }

    const AppTheme previousTheme = g_folderWindow.GetTheme();
    const auto restoreTheme     = wil::scope_exit([&]() noexcept { g_folderWindow.ApplyTheme(previousTheme); });
    const AppTheme backdropTheme =
        MakeWindowBackdropSelfTestTheme(Common::Settings::WindowBackdropMode::Acrylic, L"fileops-speedlimit-backdrop-selftest");
    g_folderWindow.ApplyTheme(backdropTheme);
    const Common::WindowBackdrop::Kind expectedToolBackdropKind =
        Common::WindowBackdrop::Resolve(Common::Settings::WindowBackdropMode::Acrylic, Common::WindowBackdrop::Target::Tool, false);

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root           = suiteRoot / L"work" / (L"fileops_speed_limit_" + NewGuidText());
    const std::filesystem::path sourceDir      = root / L"src";
    const std::filesystem::path destDir        = root / L"dst";
    const std::filesystem::path dummyRoot      = std::filesystem::path(L"/cmd-speedlimit-surface") / SanitizeDummyPathSegmentForFileOpsPrompt(NewGuidText());
    const std::filesystem::path dummySourceDir = dummyRoot / L"src";
    const std::filesystem::path dummyDestDir   = dummyRoot / L"dst";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create speed-limit test source directory.");
    state.Require(SelfTest::EnsureDirectory(destDir), L"Failed to create speed-limit test destination directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    static_cast<void>(pluginManager.EnablePlugin(kBuiltinDummyFileSystemIdForFileOpsPrompt.data(), g_settings));
    const FileSystemPluginManager::PluginEntry* dummyEntry = FindFileSystemPluginById(kBuiltinDummyFileSystemIdForFileOpsPrompt);
    state.Require(dummyEntry && dummyEntry->fileSystem && dummyEntry->informations, L"Loaded dummy file system unavailable for speed-limit prompt validation.");
    const wil::com_ptr<IFileSystem> dummyFileSystem = dummyEntry ? dummyEntry->fileSystem : nullptr;
    const wil::com_ptr<IInformations> dummyInfo     = dummyEntry ? dummyEntry->informations : nullptr;
    wil::com_ptr<IFileSystemIO> dummyIo;
    state.Require(CreateFileSystemIoForFileOpsPrompt(dummyFileSystem, dummyIo),
                  L"Loaded dummy file system missing IFileSystemIO for speed-limit prompt validation.");

    constexpr std::string_view kSeedDummyConfig =
        R"json({"maxChildrenPerDirectory":42,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"1048576"})json";
    constexpr std::string_view kSlowDummyConfig =
        R"json({"maxChildrenPerDirectory":42,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":1000,"virtualSpeedLimit":"1048576"})json";
    constexpr uint64_t kDummyPayloadBytes = 8ull * 1024ull * 1024ull;

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::FileOperationsSettings> previousFileOperations = g_settings.fileOperations;
    std::string previousDummyConfig;
    state.Require(BackupPluginConfigurationForFileOpsPrompt(dummyInfo.get(), previousDummyConfig),
                  L"Failed to snapshot dummy configuration for speed-limit prompt validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    const auto restoreSettings = wil::scope_exit([&]() noexcept
    {
        g_settings.fileOperations = previousFileOperations;
        static_cast<void>(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), previousDummyConfig));
    });

    state.Require(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), kSeedDummyConfig),
                  L"Failed to apply deterministic dummy seed configuration for speed-limit prompt validation.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummyRoot)),
                  L"Failed to create dummy root folder for speed-limit prompt validation.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummySourceDir)),
                  L"Failed to create dummy source folder for speed-limit prompt validation.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummyDestDir)),
                  L"Failed to create dummy destination folder for speed-limit prompt validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<std::filesystem::path> sourceFiles;
    sourceFiles.reserve(16u);
    for (size_t index = 0; index < 16u; ++index)
    {
        const std::filesystem::path filePath = index == 0u ? (dummySourceDir / L"payload.bin") : (dummySourceDir / std::format(L"payload_{:02}.bin", index));
        state.Require(WritePatternFileFsIoForFileOpsPrompt(dummyIo, filePath, kDummyPayloadBytes),
                      std::format(L"Failed to seed dummy speed-limit prompt payload {}.", filePath.filename().native()));
        if (! state.failure.empty())
        {
            return false;
        }
        sourceFiles.push_back(filePath);
    }

    state.Require(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), kSlowDummyConfig),
                  L"Failed to apply deterministic dummy latency for speed-limit prompt validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closePrompt = []() noexcept
    {
        if (const HWND prompt = GetFileOperationsSpeedLimitPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            if (! DebugCancelFileOperationsSpeedLimitPrompt())
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
            }
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        closePrompt();
        static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));

        if (! leftPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        }
        if (! rightPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        }
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    closePrompt();
    static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));
    {
        Common::Settings::FileOperationsSettings fileOperations = previousFileOperations.value_or(Common::Settings::FileOperationsSettings{});
        fileOperations.defaultBandwidthLimitBytesPerSecond      = kInitialLimitBytesPerSecond;
        g_settings.fileOperations                               = fileOperations;
    }

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for speed-limit prompt test (left pane).");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for speed-limit prompt test (right pane).");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, sourceDir);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, destDir);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sourceDir, SelfTest::Scale(3000ms)), L"Failed to set left pane path for speed-limit prompt test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, destDir, SelfTest::Scale(3000ms)), L"Failed to set right pane path for speed-limit prompt test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<FolderWindow::FileOperationState::Task*> beforeTasks;
    fileOps->CollectTasks(beforeTasks);
    std::unordered_set<uint64_t> existingTaskIds;
    existingTaskIds.reserve(beforeTasks.size());
    for (auto* task : beforeTasks)
    {
        if (task)
        {
            existingTaskIds.insert(task->GetId());
        }
    }

    const wil::com_ptr<IFileSystem> leftFileSystem  = dummyFileSystem;
    const wil::com_ptr<IFileSystem> rightFileSystem = dummyFileSystem;
    state.Require(leftFileSystem && rightFileSystem, L"Failed to resolve dummy file-system interfaces for speed-limit prompt test.");
    if (! leftFileSystem || ! rightFileSystem)
    {
        return false;
    }

    const HRESULT startHr = fileOps->StartOperation(FILESYSTEM_COPY,
                                                    FolderWindow::Pane::Left,
                                                    FolderWindow::Pane::Right,
                                                    leftFileSystem,
                                                    sourceFiles,
                                                    dummyDestDir,
                                                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                    false,
                                                    kInitialLimitBytesPerSecond,
                                                    FolderWindow::FileOperationState::ExecutionMode::BulkItems,
                                                    false,
                                                    rightFileSystem);
    state.Require(SUCCEEDED(startHr),
                  std::format(L"Failed to start file-operations speed-limit prompt test copy (hr=0x{:08X}).", static_cast<unsigned long>(startHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::optional<uint64_t> taskId = ResolveAndPauseNewFileOperationsTaskForSelfTest(fileOps, existingTaskIds, kInitialLimitBytesPerSecond);
    if (! taskId.has_value())
    {
        taskId = ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(5000ms));
    }
    state.Require(taskId.has_value(), L"Failed to identify the new file-operations task for speed-limit prompt test.");
    if (! taskId.has_value())
    {
        return false;
    }

    const HWND popup = WaitForWindow([&]() noexcept { return fileOps->GetPopupHwndForSelfTest(); }, SelfTest::Scale(5000ms));
    state.Require(popup != nullptr && IsWindow(popup) != FALSE, L"File-operations popup did not open for speed-limit prompt test.");
    if (! popup || IsWindow(popup) == FALSE)
    {
        return false;
    }
    Trace(std::format(L"SpeedLimitPromptTest: popup=0x{:X}", reinterpret_cast<uintptr_t>(popup)));

    state.Require(WaitForAppliedBackdropKind(popup, expectedToolBackdropKind, L"File Operations progress popup", state),
                  L"File Operations progress popup did not apply the selected Acrylic tool-window backdrop.");

    FileOperationsPopupInternal::CaptionGlyphDebugSnapshot glyphSnapshot{};
    state.Require(DebugGetFileOperationsPopupCaptionGlyphSnapshot(popup, glyphSnapshot),
                  L"Failed to capture file-operations popup caption glyph renderer snapshot.");
    state.Require(glyphSnapshot.usesDirectWriteGlyphRendering, L"File-operations popup caption glyphs should render through DirectWrite.");
    state.Require(! glyphSnapshot.usesGdiTextFallback, L"File-operations popup caption glyphs should not use the legacy text fallback.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(taskId.has_value(), L"Failed to identify the file-operations task for speed-limit prompt validation.");
    if (! taskId.has_value())
    {
        return false;
    }

    FileOperationsPopupInternal::TaskSnapshot pausedTaskSnapshot{};
    state.Require(RefreshTrackedFileOperationsTaskForSelfTest(fileOps, popup, taskId, &pausedTaskSnapshot),
                  L"Failed to resolve the live speed-limit prompt task before pausing it.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (auto* task = fileOps->FindTask(taskId.value()))
    {
        task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
        if (! task->IsPaused())
        {
            task->TogglePause();
        }
    }

    const auto holdDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    bool taskHeld           = false;
    while (std::chrono::steady_clock::now() < holdDeadline)
    {
        PumpPendingMessages();
        if (auto* task = fileOps->FindTask(taskId.value());
            task && (task->IsPaused() || task->IsQueuePaused() || task->IsWaitingForOthers() || task->IsWaitingInQueue() || ! task->HasStarted()))
        {
            taskHeld = true;
            break;
        }

        std::this_thread::sleep_for(20ms);
    }
    state.Require(taskHeld, L"File-operations speed-limit prompt test task did not stay held before prompt validation.");
    if (! taskHeld)
    {
        return false;
    }
    Trace(std::format(L"SpeedLimitPromptTest: task {} resolved, held, and speed limit set", taskId.value()));

    FileOperationsPopupInternal::PopupSelfTestInvoke openPrompt{};
    openPrompt.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskSpeedLimit;
    openPrompt.taskId = taskId.value();
    openPrompt.data   = 1u;

    const auto requireDxSurface = [&](HWND prompt, std::wstring_view expectedText, std::wstring_view failureContext) noexcept
    {
        FileOperationsSpeedLimitPromptDebugSnapshot snapshot{};
        state.Require(DebugGetFileOperationsSpeedLimitPromptSnapshot(snapshot),
                      std::format(L"Failed to capture custom speed-limit prompt snapshot during {}.", failureContext));
        state.Require(snapshot.usesDxUiHost, std::format(L"Custom speed-limit prompt should use a DxUi host during {}.", failureContext));
        state.Require(snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Custom speed-limit prompt should not expose more than one visible child window during {}; saw {}.",
                                  failureContext,
                                  snapshot.visibleChildWindowCount));
        state.Require(snapshot.text == expectedText,
                      std::format(L"Custom speed-limit prompt should expose '{}' through the live snapshot during {}.", expectedText, failureContext));
        state.Require(WaitForAppliedBackdropKind(prompt, expectedToolBackdropKind, L"File Operations speed-limit prompt", state),
                      std::format(L"File Operations speed-limit prompt did not apply the selected Acrylic tool-window backdrop during {}.", failureContext));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(prompt);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation stats for custom speed-limit prompt during {}.", failureContext));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Custom speed-limit prompt should expose visible UI Automation descendants during {}.", failureContext));
            state.Require(uiaPatternStats->editControlCount > 0u,
                          std::format(L"Custom speed-limit prompt should expose a visible editable DX field during {}.", failureContext));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Custom speed-limit prompt should expose ValuePattern for the speed-limit field during {}.", failureContext));
            state.Require(uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Custom speed-limit prompt should expose visible UI Automation command buttons during {}.", failureContext));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Custom speed-limit prompt should expose InvokePattern for its visible DX command button during {}.", failureContext));
        }

        const auto valueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        state.Require(valueState.has_value(), std::format(L"Custom speed-limit prompt should expose a visible editable DX field during {}.", failureContext));
        if (valueState.has_value())
        {
            state.Require(! valueState->name.empty(),
                          std::format(L"Custom speed-limit prompt DX edit should expose a stable accessible name during {}.", failureContext));
            state.Require(! valueState->isReadOnly, std::format(L"Custom speed-limit prompt DX edit should remain editable during {}.", failureContext));
            state.Require(
                valueState->value == snapshot.text,
                std::format(L"Custom speed-limit prompt field should match the live prompt snapshot during {}; saw '{}'.", failureContext, valueState->value));
        }

        const auto buttonState = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
        state.Require(buttonState.has_value(), std::format(L"Custom speed-limit prompt should expose a visible DX command button during {}.", failureContext));
        if (buttonState.has_value())
        {
            state.Require(
                ! buttonState->name.empty(),
                std::format(L"Custom speed-limit prompt visible DX command button should expose a stable accessible name during {}.", failureContext));
        }
    };

    const auto setPromptTextForConfirm = [&](HWND prompt, std::wstring_view requestedText, std::wstring_view failureContext) noexcept
    {
        const auto initialValueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        state.Require(initialValueState.has_value(),
                      std::format(L"Custom speed-limit prompt should expose a visible editable DX field before mutation during {}.", failureContext));
        if (! initialValueState.has_value())
        {
            return false;
        }

        const std::wstring editName = initialValueState->name;
        state.Require(DebugSetFileOperationsSpeedLimitPromptText(requestedText),
                      std::format(L"Failed to set custom speed-limit prompt text '{}' during {}.", requestedText, failureContext));

        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        bool updated        = false;
        std::wstring lastValueText;
        std::wstring lastSnapshotText;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            FileOperationsSpeedLimitPromptDebugSnapshot snapshot{};
            const auto valueState = CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, editName);
            if (valueState.has_value())
            {
                lastValueText = valueState->value;
            }
            if (DebugGetFileOperationsSpeedLimitPromptSnapshot(snapshot))
            {
                lastSnapshotText = snapshot.text;
            }
            if (snapshot.text == requestedText)
            {
                updated = true;
                break;
            }
            std::this_thread::sleep_for(20ms);
        }
        state.Require(updated,
                      std::format(L"Custom speed-limit prompt text did not update after the debug setter during {}; requested='{}' value='{}' snapshot='{}'.",
                                  failureContext,
                                  requestedText,
                                  lastValueText,
                                  lastSnapshotText));

        FileOperationsSpeedLimitPromptDebugSnapshot editedSnapshot{};
        state.Require(DebugGetFileOperationsSpeedLimitPromptSnapshot(editedSnapshot),
                      std::format(L"Failed to recapture custom speed-limit prompt snapshot during {}.", failureContext));
        state.Require(editedSnapshot.text == requestedText,
                      std::format(L"Custom speed-limit prompt snapshot should show '{}' after the debug setter during {}; saw '{}'.",
                                  requestedText,
                                  failureContext,
                                  editedSnapshot.text));
        return state.failure.empty();
    };

    FileOperationsSpeedLimitPromptDebugSnapshot initialSnapshot{};
    HWND confirmPrompt        = nullptr;
    const bool confirmInvoked = RunFileOperationsSpeedLimitPromptModalCycle(popup,
                                                                            openPrompt,
                                                                            [&](const HWND prompt) noexcept
    {
        confirmPrompt = prompt;
        state.Require(prompt != nullptr && IsWindow(prompt) != FALSE, L"Custom speed-limit prompt did not open during initial confirm pass.");
        if (! prompt || IsWindow(prompt) == FALSE)
        {
            return;
        }

        state.Require(IsOwnedBy(prompt, popup), L"Custom speed-limit prompt should be owned by the file-operations popup during initial confirm pass.");
        state.Require(WindowExposesUiaProvider(prompt), L"Custom speed-limit prompt should answer WM_GETOBJECT during initial confirm pass.");
        state.Require(DebugGetFileOperationsSpeedLimitPromptSnapshot(initialSnapshot),
                      L"Failed to capture initial custom speed-limit prompt snapshot before baseline confirm interaction.");
        if (! state.failure.empty())
        {
            return;
        }

        requireDxSurface(prompt, initialSnapshot.text, L"initial confirm pass");
        state.Require(setPromptTextForConfirm(prompt, L"128KB", L"initial confirm pass"),
                      L"Custom speed-limit prompt baseline text update failed before confirm.");
        state.Require(DebugConfirmFileOperationsSpeedLimitPrompt(), L"Failed to confirm custom speed-limit prompt during initial confirm pass.");
        state.Require(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)),
                      L"Custom speed-limit prompt did not close after confirm during initial confirm pass.");
    });
    Trace(std::format(L"SpeedLimitPromptTest: invoke result={} task={} context=initial confirm pass", confirmInvoked, taskId.value()));
    state.Require(confirmInvoked, L"Failed to request the custom speed-limit prompt through the file-operations popup during initial confirm pass.");
    state.Require(confirmPrompt != nullptr && IsWindow(confirmPrompt) == FALSE,
                  L"Custom speed-limit prompt should not remain open after the initial confirm pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsSpeedLimitPromptLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr uint64_t kInitialLimitBytesPerSecond = 64ull * 1024ull;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for speed-limit prompt live interaction validation.");
    if (! fileOps)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root       = suiteRoot / L"work" / (L"fileops_speed_limit_live_dx_" + NewGuidText());
    const std::filesystem::path sourceDir  = root / L"src";
    const std::filesystem::path destDir    = root / L"dst";
    const std::filesystem::path sourceFile = sourceDir / L"payload.bin";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create speed-limit live interaction source directory.");
    state.Require(SelfTest::EnsureDirectory(destDir), L"Failed to create speed-limit live interaction destination directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::string payload(64u * 1024u * 1024u, 'v');
    state.Require(SelfTest::WriteTextFile(sourceFile, payload), L"Failed to create speed-limit live interaction payload.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::FileOperationsSettings> previousFileOperations = g_settings.fileOperations;
    const auto restoreSettings = wil::scope_exit([&]() noexcept { g_settings.fileOperations = previousFileOperations; });

    const auto closePrompt = []() noexcept
    {
        if (const HWND prompt = GetFileOperationsSpeedLimitPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            if (! DebugCancelFileOperationsSpeedLimitPrompt())
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
            }
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        closePrompt();
        static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));

        if (! leftPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        }
        if (! rightPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        }
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    closePrompt();
    static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));
    {
        Common::Settings::FileOperationsSettings fileOperations = previousFileOperations.value_or(Common::Settings::FileOperationsSettings{});
        fileOperations.defaultBandwidthLimitBytesPerSecond      = kInitialLimitBytesPerSecond;
        g_settings.fileOperations                               = fileOperations;
    }

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for speed-limit live interaction test (left pane).");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for speed-limit live interaction test (right pane).");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, sourceDir);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, destDir);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sourceDir, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for speed-limit live interaction test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, destDir, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for speed-limit live interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<FolderWindow::FileOperationState::Task*> beforeTasks;
    fileOps->CollectTasks(beforeTasks);
    std::unordered_set<uint64_t> existingTaskIds;
    existingTaskIds.reserve(beforeTasks.size());
    for (auto* task : beforeTasks)
    {
        if (task)
        {
            existingTaskIds.insert(task->GetId());
        }
    }

    const wil::com_ptr<IFileSystem> leftFileSystem  = g_folderWindow.GetFileSystem(FolderWindow::Pane::Left);
    const wil::com_ptr<IFileSystem> rightFileSystem = g_folderWindow.GetFileSystem(FolderWindow::Pane::Right);
    state.Require(leftFileSystem && rightFileSystem, L"Failed to resolve local file-system interfaces for speed-limit live interaction test.");
    if (! leftFileSystem || ! rightFileSystem)
    {
        return false;
    }

    const HRESULT startHr = fileOps->StartOperation(FILESYSTEM_COPY,
                                                    FolderWindow::Pane::Left,
                                                    FolderWindow::Pane::Right,
                                                    leftFileSystem,
                                                    {sourceFile},
                                                    destDir,
                                                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                    false,
                                                    kInitialLimitBytesPerSecond,
                                                    FolderWindow::FileOperationState::ExecutionMode::BulkItems,
                                                    false,
                                                    rightFileSystem);
    state.Require(SUCCEEDED(startHr),
                  std::format(L"Failed to start file-operations speed-limit live interaction test copy (hr=0x{:08X}).", static_cast<unsigned long>(startHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::optional<uint64_t> taskId = ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(5000ms));
    state.Require(taskId.has_value(), L"Failed to identify the new file-operations task for speed-limit live interaction test.");
    if (! taskId.has_value())
    {
        return false;
    }

    const HWND popup = WaitForWindow([&]() noexcept { return fileOps->GetPopupHwndForSelfTest(); }, SelfTest::Scale(5000ms));
    state.Require(popup != nullptr && IsWindow(popup) != FALSE, L"File-operations popup did not open for speed-limit live interaction test.");
    if (! popup || IsWindow(popup) == FALSE)
    {
        return false;
    }

    const auto startedDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    bool taskStarted           = false;
    while (std::chrono::steady_clock::now() < startedDeadline)
    {
        PumpPendingMessages();
        auto* task = fileOps->FindTask(taskId.value());
        FileOperationsPopupInternal::TaskSnapshot popupTaskSnapshot{};
        const bool havePopupTaskSnapshot = DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value_or(0), popupTaskSnapshot);
        if (havePopupTaskSnapshot && popupTaskSnapshot.taskId != 0)
        {
            taskId = popupTaskSnapshot.taskId;
            task   = fileOps->FindTask(taskId.value());
        }

        if (task && task->HasStarted())
        {
            taskStarted = true;
            break;
        }

        if (havePopupTaskSnapshot)
        {
            taskStarted = true;
            break;
        }

        std::this_thread::sleep_for(20ms);
    }
    state.Require(taskStarted, L"File-operations speed-limit live interaction task did not start in time.");
    if (! taskStarted)
    {
        return false;
    }
    if (auto* task = fileOps->FindTask(taskId.value()))
    {
        task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
    }

    FileOperationsPopupInternal::PopupSelfTestInvoke openPrompt{};
    openPrompt.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskSpeedLimit;
    openPrompt.taskId = taskId.value();
    openPrompt.data   = 1u;

    auto* trackedTask = fileOps->FindTask(taskId.value());
    state.Require(trackedTask != nullptr, L"Speed-limit live interaction task disappeared before prompt open.");
    if (! trackedTask)
    {
        return false;
    }

    const auto mutateSpeedLimitPromptText = [&](const HWND prompt, std::wstring_view expectedValue, std::wstring_view context) noexcept
    {
        FileOperationsSpeedLimitPromptDebugSnapshot snapshot{};
        state.Require(DebugGetFileOperationsSpeedLimitPromptSnapshot(snapshot),
                      std::format(L"Failed to capture custom speed-limit prompt snapshot during {}.", context));
        state.Require(snapshot.usesDxUiHost, std::format(L"Custom speed-limit prompt should use a DxUi host during {}.", context));
        state.Require(snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Custom speed-limit prompt should not expose more than one visible child window during {}; saw {}.",
                                  context,
                                  snapshot.visibleChildWindowCount));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto initialValueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        state.Require(initialValueState.has_value(), std::format(L"Custom speed-limit prompt should expose a visible editable DX field during {}.", context));
        if (! initialValueState.has_value() || ! state.failure.empty())
        {
            return false;
        }

        state.Require(! initialValueState->name.empty(),
                      std::format(L"Custom speed-limit prompt DX edit should expose a stable accessible name during {}.", context));
        state.Require(! initialValueState->isReadOnly, std::format(L"Custom speed-limit prompt DX edit should remain editable during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring editName = initialValueState->name;
        state.Require(SetVisibleDescendantValueByName(prompt, UIA_EditControlTypeId, editName, expectedValue),
                      std::format(L"Failed to set custom speed-limit prompt DX edit '{}' during {}.", editName, context));

        const auto waitForEditValue = [&](std::wstring_view value) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, editName);
                if (valueState.has_value() && valueState->value == value)
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            const auto valueState = CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, editName);
            return valueState.has_value() && valueState->value == value;
        };

        state.Require(waitForEditValue(expectedValue),
                      std::format(L"Custom speed-limit prompt DX edit did not update after live UIA interaction during {}.", context));

        FileOperationsSpeedLimitPromptDebugSnapshot editedSnapshot{};
        state.Require(DebugGetFileOperationsSpeedLimitPromptSnapshot(editedSnapshot),
                      std::format(L"Failed to recapture custom speed-limit prompt snapshot during {}.", context));
        state.Require(
            ! editedSnapshot.text.empty(),
            std::format(L"Custom speed-limit prompt should keep non-empty normalized text after the live DX edit during {}; saw empty text.", context));
        return state.failure.empty();
    };

    const std::wstring okButtonText = LoadStringResource(nullptr, IDS_BTN_OK);
    state.Require(! okButtonText.empty(), L"Failed to resolve OK button caption for custom speed-limit prompt.");
    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(), L"Failed to resolve Cancel button caption for custom speed-limit prompt.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileOperationsSpeedLimitPromptDebugSnapshot initialSnapshot{};
    HWND confirmPrompt        = nullptr;
    const bool confirmInvoked = RunFileOperationsSpeedLimitPromptModalCycle(popup,
                                                                            openPrompt,
                                                                            [&](const HWND prompt) noexcept
    {
        confirmPrompt = prompt;
        state.Require(prompt != nullptr && IsWindow(prompt) != FALSE, L"Custom speed-limit prompt did not open for live confirm interaction.");
        if (! prompt || IsWindow(prompt) == FALSE)
        {
            return;
        }

        state.Require(IsOwnedBy(prompt, popup), L"Custom speed-limit prompt should be owned by the file-operations popup during live confirm interaction.");
        state.Require(DebugGetFileOperationsSpeedLimitPromptSnapshot(initialSnapshot),
                      L"Failed to capture initial custom speed-limit prompt snapshot before live confirm interaction.");
        state.Require(initialSnapshot.usesDxUiHost, L"Custom speed-limit prompt should use a DxUi host before live confirm interaction.");
        state.Require(initialSnapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Custom speed-limit prompt should not expose more than one visible child window before live confirm interaction; saw {}.",
                                  initialSnapshot.visibleChildWindowCount));
        const auto initialValueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        state.Require(initialValueState.has_value(), L"Custom speed-limit prompt should expose a visible editable DX field before live confirm interaction.");
        if (initialValueState.has_value())
        {
            state.Require(initialValueState->value == initialSnapshot.text,
                          L"Custom speed-limit prompt field should match the live prompt snapshot before live confirm interaction.");
        }
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugSetFileOperationsSpeedLimitPromptText(L"128KB"), L"Failed to set custom speed-limit prompt text before live confirm interaction.");
        state.Require(InvokeVisibleDescendantByName(prompt, UIA_ButtonControlTypeId, okButtonText),
                      L"Custom speed-limit prompt visible DX OK action did not expose live UIA InvokePattern interaction.");
        state.Require(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)),
                      L"Custom speed-limit prompt did not close after live UIA InvokePattern confirmation.");
    });
    state.Require(confirmInvoked, L"Failed to request the custom speed-limit prompt for live confirm interaction.");
    state.Require(confirmPrompt != nullptr && IsWindow(confirmPrompt) == FALSE,
                  L"Custom speed-limit prompt should not remain open after live confirm interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto limitDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < limitDeadline)
    {
        PumpPendingMessages();
        if (trackedTask->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire) == 128ull * 1024ull)
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }
    state.Require(trackedTask->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire) == 128ull * 1024ull,
                  L"Live confirm interaction did not update the task speed limit.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsSpeedLimitPromptLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr uint64_t kInitialLimitBytesPerSecond = 64ull * 1024ull;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for speed-limit prompt churn validation.");
    if (! fileOps)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root           = suiteRoot / L"work" / (L"fileops_speed_limit_churn_" + NewGuidText());
    const std::filesystem::path sourceDir      = root / L"src";
    const std::filesystem::path destDir        = root / L"dst";
    const std::filesystem::path sourceFile     = sourceDir / L"payload.bin";
    const std::filesystem::path dummyRoot      = std::filesystem::path(L"/cmd-speedlimit-churn") / SanitizeDummyPathSegmentForFileOpsPrompt(NewGuidText());
    const std::filesystem::path dummySourceDir = dummyRoot / L"src";
    const std::filesystem::path dummyDestDir   = dummyRoot / L"dst";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create speed-limit churn source directory.");
    state.Require(SelfTest::EnsureDirectory(destDir), L"Failed to create speed-limit churn destination directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SelfTest::WriteTextFile(sourceFile, "payload"), L"Failed to create visible churn shell payload.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    static_cast<void>(pluginManager.EnablePlugin(kBuiltinDummyFileSystemIdForFileOpsPrompt.data(), g_settings));
    const FileSystemPluginManager::PluginEntry* dummyEntry = FindFileSystemPluginById(kBuiltinDummyFileSystemIdForFileOpsPrompt);
    state.Require(dummyEntry && dummyEntry->fileSystem && dummyEntry->informations, L"Loaded dummy file system unavailable for speed-limit churn validation.");
    const wil::com_ptr<IFileSystem> dummyFileSystem = dummyEntry ? dummyEntry->fileSystem : nullptr;
    const wil::com_ptr<IInformations> dummyInfo     = dummyEntry ? dummyEntry->informations : nullptr;
    wil::com_ptr<IFileSystemIO> dummyIo;
    state.Require(CreateFileSystemIoForFileOpsPrompt(dummyFileSystem, dummyIo),
                  L"Loaded dummy file system missing IFileSystemIO for speed-limit churn validation.");

    constexpr std::string_view kSeedDummyConfig =
        R"json({"maxChildrenPerDirectory":42,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"1048576"})json";
    constexpr std::string_view kSlowDummyConfig =
        R"json({"maxChildrenPerDirectory":42,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":1000,"virtualSpeedLimit":"1048576"})json";
    constexpr uint64_t kDummyPayloadBytes                  = 8ull * 1024ull * 1024ull;
    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::FileOperationsSettings> previousFileOperations = g_settings.fileOperations;
    std::string previousDummyConfig;
    state.Require(BackupPluginConfigurationForFileOpsPrompt(dummyInfo.get(), previousDummyConfig),
                  L"Failed to snapshot dummy configuration for speed-limit churn validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    const auto restoreSettings = wil::scope_exit([&]() noexcept
    {
        g_settings.fileOperations = previousFileOperations;
        static_cast<void>(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), previousDummyConfig));
    });

    state.Require(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), kSeedDummyConfig),
                  L"Failed to apply deterministic dummy seed configuration for speed-limit churn validation.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummyRoot)),
                  L"Failed to create dummy root folder for speed-limit churn validation.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummySourceDir)),
                  L"Failed to create dummy source folder for speed-limit churn validation.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummyDestDir)),
                  L"Failed to create dummy destination folder for speed-limit churn validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<std::filesystem::path> sourceFiles;
    sourceFiles.reserve(16u);
    for (size_t index = 0; index < 16u; ++index)
    {
        const std::filesystem::path filePath = index == 0u ? (dummySourceDir / L"payload.bin") : (dummySourceDir / std::format(L"payload_{:02}.bin", index));
        state.Require(WritePatternFileFsIoForFileOpsPrompt(dummyIo, filePath, kDummyPayloadBytes),
                      std::format(L"Failed to seed dummy churn payload {}.", filePath.filename().native()));
        if (! state.failure.empty())
        {
            return false;
        }
        sourceFiles.push_back(filePath);
    }

    state.Require(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), kSlowDummyConfig),
                  L"Failed to apply deterministic dummy latency for speed-limit churn validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closePrompt = []() noexcept
    {
        if (const HWND prompt = GetFileOperationsSpeedLimitPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            if (! DebugCancelFileOperationsSpeedLimitPrompt())
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
            }
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        closePrompt();
        static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));

        if (! leftPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        }
        if (! rightPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        }
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    closePrompt();
    static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));
    {
        Common::Settings::FileOperationsSettings fileOperations = previousFileOperations.value_or(Common::Settings::FileOperationsSettings{});
        fileOperations.defaultBandwidthLimitBytesPerSecond      = kInitialLimitBytesPerSecond;
        g_settings.fileOperations                               = fileOperations;
    }

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for speed-limit churn test (left pane).");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for speed-limit churn test (right pane).");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, sourceDir);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, destDir);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sourceDir, SelfTest::Scale(3000ms)), L"Failed to set left pane path for speed-limit churn test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, destDir, SelfTest::Scale(3000ms)), L"Failed to set right pane path for speed-limit churn test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const wil::com_ptr<IFileSystem> leftFileSystem  = dummyFileSystem;
    const wil::com_ptr<IFileSystem> rightFileSystem = dummyFileSystem;
    state.Require(leftFileSystem && rightFileSystem, L"Failed to resolve dummy file-system interfaces for speed-limit churn test.");
    if (! leftFileSystem || ! rightFileSystem)
    {
        return false;
    }

    std::vector<FolderWindow::FileOperationState::Task*> beforeTasks;
    fileOps->CollectTasks(beforeTasks);
    std::unordered_set<uint64_t> existingTaskIds;
    existingTaskIds.reserve(beforeTasks.size());
    for (auto* task : beforeTasks)
    {
        if (task)
        {
            existingTaskIds.insert(task->GetId());
        }
    }

    const HRESULT startHr = fileOps->StartOperation(FILESYSTEM_COPY,
                                                    FolderWindow::Pane::Left,
                                                    FolderWindow::Pane::Right,
                                                    leftFileSystem,
                                                    sourceFiles,
                                                    dummyDestDir,
                                                    static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                    false,
                                                    kInitialLimitBytesPerSecond,
                                                    FolderWindow::FileOperationState::ExecutionMode::BulkItems,
                                                    false,
                                                    rightFileSystem);
    state.Require(SUCCEEDED(startHr),
                  std::format(L"Failed to start file-operations speed-limit churn test copy (hr=0x{:08X}).", static_cast<unsigned long>(startHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::optional<uint64_t> taskId = ResolveAndPauseNewFileOperationsTaskForSelfTest(fileOps, existingTaskIds, kInitialLimitBytesPerSecond);

    HWND popup = WaitForWindow([&]() noexcept { return fileOps->GetPopupHwndForSelfTest(); }, SelfTest::Scale(5000ms));
    state.Require(popup != nullptr && IsWindow(popup) != FALSE, L"File-operations popup did not open for speed-limit churn validation.");
    if (! popup || IsWindow(popup) == FALSE)
    {
        return false;
    }

    const auto readyDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    if (! taskId.has_value())
    {
        taskId = ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(250ms));
    }
    bool taskReady = false;
    if (auto* task = taskId.has_value() ? fileOps->FindTask(taskId.value()) : nullptr)
    {
        task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
    }
    while (std::chrono::steady_clock::now() < readyDeadline)
    {
        PumpPendingMessages();

        auto* task = taskId.has_value() ? fileOps->FindTask(taskId.value()) : nullptr;
        FileOperationsPopupInternal::TaskSnapshot popupTaskSnapshot{};
        const bool havePopupTaskSnapshot = DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value_or(0), popupTaskSnapshot);
        if (havePopupTaskSnapshot && popupTaskSnapshot.taskId != 0)
        {
            taskId = popupTaskSnapshot.taskId;
            task   = fileOps->FindTask(taskId.value());
            if (task)
            {
                task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
            }
        }

        if (task)
        {
            taskReady = true;
            task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
            break;
        }

        if (havePopupTaskSnapshot && popupTaskSnapshot.taskId != 0 && ! popupTaskSnapshot.finished)
        {
            taskReady = true;
            if (task)
            {
                task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
            }
            break;
        }

        std::this_thread::sleep_for(20ms);
    }
    state.Require(taskReady, L"File-operations speed-limit churn task did not become available in time.");
    if (! taskReady)
    {
        return false;
    }
    state.Require(taskId.has_value(), L"Failed to identify the file-operations task for speed-limit churn validation.");
    if (! taskId.has_value())
    {
        return false;
    }

    FileOperationsPopupInternal::TaskSnapshot readyTaskSnapshot{};
    state.Require(RefreshTrackedFileOperationsTaskForSelfTest(fileOps, popup, taskId, &readyTaskSnapshot),
                  L"Failed to resolve the queued speed-limit churn task before prompt cycling.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (auto* task = fileOps->FindTask(taskId.value()))
    {
        task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
        if (! task->IsPaused())
        {
            task->TogglePause();
        }
    }
    const auto holdDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    bool taskHeld           = false;
    while (std::chrono::steady_clock::now() < holdDeadline)
    {
        PumpPendingMessages();
        if (auto* task = fileOps->FindTask(taskId.value()); task)
        {
            taskHeld = true;
            break;
        }
        std::this_thread::sleep_for(20ms);
    }
    state.Require(taskHeld, L"Speed-limit churn task did not remain available before prompt cycling.");
    if (! taskHeld)
    {
        return false;
    }

    FileOperationsPopupInternal::PopupSelfTestInvoke openPrompt{};
    openPrompt.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskSpeedLimit;
    openPrompt.taskId = taskId.value();
    openPrompt.data   = 1u;

    auto ensureChurnTaskForCycle = [&](const size_t cycle, const uint64_t expectedLimit, FileOperationsPopupInternal::TaskSnapshot& cycleSnapshot) -> bool
    {
        if (popup && IsWindow(popup) != FALSE && RefreshTrackedFileOperationsTaskForSelfTest(fileOps, popup, taskId, &cycleSnapshot))
        {
            return true;
        }

        std::vector<FolderWindow::FileOperationState::Task*> beforeRestartTasks;
        fileOps->CollectTasks(beforeRestartTasks);
        std::unordered_set<uint64_t> restartExistingTaskIds;
        restartExistingTaskIds.reserve(beforeRestartTasks.size());
        for (auto* task : beforeRestartTasks)
        {
            if (task)
            {
                restartExistingTaskIds.insert(task->GetId());
            }
        }

        const HRESULT restartHr = fileOps->StartOperation(FILESYSTEM_COPY,
                                                          FolderWindow::Pane::Left,
                                                          FolderWindow::Pane::Right,
                                                          leftFileSystem,
                                                          sourceFiles,
                                                          dummyDestDir,
                                                          static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                          false,
                                                          expectedLimit,
                                                          FolderWindow::FileOperationState::ExecutionMode::BulkItems,
                                                          false,
                                                          rightFileSystem);
        state.Require(SUCCEEDED(restartHr),
                      std::format(L"Failed to restart file-operations speed-limit churn copy before cycle {} (hr=0x{:08X}).",
                                  cycle,
                                  static_cast<unsigned long>(restartHr)));
        if (! state.failure.empty())
        {
            return false;
        }

        taskId = ResolveAndPauseNewFileOperationsTaskForSelfTest(fileOps, restartExistingTaskIds, expectedLimit);

        popup = WaitForWindow([&]() noexcept { return fileOps->GetPopupHwndForSelfTest(); }, SelfTest::Scale(5000ms));
        state.Require(popup != nullptr && IsWindow(popup) != FALSE,
                      std::format(L"File-operations popup did not reopen for speed-limit churn cycle {}.", cycle));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return false;
        }

        if (! taskId.has_value())
        {
            taskId = ResolveNewFileOperationsTaskIdForSelfTest(fileOps, restartExistingTaskIds, SelfTest::Scale(5000ms));
        }
        state.Require(taskId.has_value(), std::format(L"Failed to identify restarted speed-limit churn task before cycle {}.", cycle));
        if (! taskId.has_value())
        {
            return false;
        }

        if (auto* task = fileOps->FindTask(taskId.value()))
        {
            task->SetDesiredSpeedLimit(expectedLimit);
            if (! task->IsPaused())
            {
                task->TogglePause();
            }
        }

        return RefreshTrackedFileOperationsTaskForSelfTest(fileOps, popup, taskId, &cycleSnapshot);
    };

    uint64_t expectedLimit   = kInitialLimitBytesPerSecond;
    constexpr size_t kCycles = 3u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        closePrompt();
        const auto promptCloseDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < promptCloseDeadline)
        {
            PumpPendingMessages();
            const HWND lingeringPrompt = GetFileOperationsSpeedLimitPromptHandle();
            if (! lingeringPrompt || IsWindow(lingeringPrompt) == FALSE)
            {
                break;
            }
            std::this_thread::sleep_for(20ms);
        }
        state.Require(GetFileOperationsSpeedLimitPromptHandle() == nullptr,
                      std::format(L"Custom speed-limit prompt should close before churn cycle {} starts.", cycle));
        FileOperationsPopupInternal::TaskSnapshot cycleSnapshot{};
        const bool haveChurnTask = ensureChurnTaskForCycle(cycle, expectedLimit, cycleSnapshot);
        state.Require(haveChurnTask && taskId.has_value(), std::format(L"Failed to resolve the queued speed-limit churn task before cycle {}.", cycle));
        if (! haveChurnTask || ! taskId.has_value())
        {
            closePrompt();
            return false;
        }

        if (taskId.has_value())
        {
            openPrompt.taskId = taskId.value();
        }
        if (auto* task = fileOps->FindTask(taskId.value()))
        {
            task->SetDesiredSpeedLimit(expectedLimit);
            if (! task->IsPaused())
            {
                task->TogglePause();
            }
        }
        state.Require(fileOps->GetPopupHwndForSelfTest() == popup,
                      std::format(L"File-operations popup should stay open for speed-limit churn cycle {}.", cycle));
        if (! state.failure.empty())
        {
            return false;
        }

        const bool accept                = false;
        const uint64_t requestedLimit    = (96ull + 32ull * static_cast<uint64_t>(cycle)) * 1024ull;
        const std::wstring requestedText = std::format(L"{}KB", requestedLimit / 1024ull);

        HWND promptHandle  = nullptr;
        const bool invoked = RunFileOperationsSpeedLimitPromptModalCycle(popup,
                                                                         openPrompt,
                                                                         [&](const HWND prompt) noexcept
        {
            promptHandle = prompt;
            state.Require(prompt != nullptr && IsWindow(prompt) != FALSE, std::format(L"Custom speed-limit prompt did not open during cycle {}.", cycle));
            if (! prompt || IsWindow(prompt) == FALSE)
            {
                return;
            }

            state.Require(IsOwnedBy(prompt, popup), std::format(L"Custom speed-limit prompt should be owned by the popup during cycle {}.", cycle));
            state.Require(WindowExposesUiaProvider(prompt), std::format(L"Custom speed-limit prompt should answer WM_GETOBJECT during cycle {}.", cycle));

            FileOperationsSpeedLimitPromptDebugSnapshot snapshot{};
            state.Require(DebugGetFileOperationsSpeedLimitPromptSnapshot(snapshot),
                          std::format(L"Failed to capture custom speed-limit prompt snapshot during cycle {}.", cycle));
            state.Require(snapshot.usesDxUiHost, std::format(L"Custom speed-limit prompt should use a DxUi host during cycle {}.", cycle));
            state.Require(snapshot.visibleChildWindowCount <= 1u,
                          std::format(L"Custom speed-limit prompt should not expose more than one visible child window during cycle {}; saw {}.",
                                      cycle,
                                      snapshot.visibleChildWindowCount));
            state.Require(snapshot.initialLimitBytesPerSecond == expectedLimit,
                          std::format(L"Custom speed-limit prompt initial limit mismatch during cycle {}.", cycle));
            state.Require(snapshot.text == FormatBytesCompact(expectedLimit),
                          std::format(L"Custom speed-limit prompt should show the live limit text during cycle {}.", cycle));

            const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(prompt);
            state.Require(uiaPatternStats.has_value(),
                          std::format(L"Failed to collect UI Automation stats for custom speed-limit prompt during cycle {}.", cycle));
            if (uiaPatternStats.has_value())
            {
                state.Require(uiaPatternStats->visibleElementCount > 0u,
                              std::format(L"Custom speed-limit prompt should expose visible UI Automation descendants during cycle {}.", cycle));
                state.Require(uiaPatternStats->editControlCount > 0u,
                              std::format(L"Custom speed-limit prompt should expose a visible edit descendant during cycle {}.", cycle));
                state.Require(uiaPatternStats->valuePatternCount > 0u,
                              std::format(L"Custom speed-limit prompt should expose ValuePattern during cycle {}.", cycle));
                state.Require(uiaPatternStats->buttonControlCount > 0u,
                              std::format(L"Custom speed-limit prompt should expose visible command buttons during cycle {}.", cycle));
            }

            const auto valueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
            state.Require(valueState.has_value(),
                          std::format(L"Custom speed-limit prompt should expose a visible editable ValuePattern field during cycle {}.", cycle));
            if (valueState.has_value())
            {
                state.Require(! valueState->isReadOnly, std::format(L"Custom speed-limit prompt field should remain editable during cycle {}.", cycle));
                state.Require(valueState->value == snapshot.text,
                              std::format(L"Custom speed-limit prompt ValuePattern should start with '{}' during cycle {}.", snapshot.text, cycle));
                state.Require(! valueState->name.empty(),
                              std::format(L"Custom speed-limit prompt field should expose a stable accessible name during cycle {}.", cycle));
            }

            const auto buttonState = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
            state.Require(buttonState.has_value(), std::format(L"Custom speed-limit prompt should expose a visible DX command button during cycle {}.", cycle));
            if (buttonState.has_value())
            {
                state.Require(! buttonState->name.empty(),
                              std::format(L"Custom speed-limit prompt command button should expose a stable accessible name during cycle {}.", cycle));
            }

            state.Require(DebugSetFileOperationsSpeedLimitPromptText(requestedText),
                          std::format(L"Failed to set custom speed-limit prompt text during cycle {}.", cycle));
            state.Require(accept ? DebugConfirmFileOperationsSpeedLimitPrompt() : DebugCancelFileOperationsSpeedLimitPrompt(),
                          std::format(L"Failed to close the custom speed-limit prompt during cycle {}.", cycle));
            state.Require(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)), std::format(L"Custom speed-limit prompt did not close after cycle {}.", cycle));
        });
        state.Require(invoked, std::format(L"Failed to request the custom speed-limit prompt during cycle {}.", cycle));
        state.Require(promptHandle != nullptr && IsWindow(promptHandle) == FALSE,
                      std::format(L"Custom speed-limit prompt should not remain open after cycle {}.", cycle));

        if (accept)
        {
            expectedLimit = requestedLimit;
        }

        const auto limitDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < limitDeadline)
        {
            PumpPendingMessages();
            if (auto* task = fileOps->FindTask(taskId.value()); task && task->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire) == expectedLimit)
            {
                break;
            }
            std::this_thread::sleep_for(20ms);
        }

        FileOperationsPopupInternal::TaskSnapshot postCycleSnapshot{};
        const bool havePostCycleSnapshot = DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value_or(0), postCycleSnapshot);
        if (taskId.has_value())
        {
            openPrompt.taskId = taskId.value();
        }
        auto* trackedTask = taskId.has_value() ? fileOps->FindTask(taskId.value()) : nullptr;
        if (trackedTask != nullptr || (havePostCycleSnapshot && ! postCycleSnapshot.finished))
        {
            state.Require((trackedTask != nullptr && trackedTask->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire) == expectedLimit) ||
                              (havePostCycleSnapshot && ! postCycleSnapshot.finished && postCycleSnapshot.desiredSpeedLimitBytesPerSecond == expectedLimit),
                          std::format(L"Speed-limit churn cycle {} did not leave the expected live limit (taskId={}, tracked={}, trackedLimit={}, "
                                      L"snapshotFound={}, snapshotFinished={}, snapshotTaskId={}, snapshotLimit={}, expected={}).",
                                      cycle,
                                      taskId.value_or(0),
                                      trackedTask ? L"yes" : L"no",
                                      trackedTask ? trackedTask->_desiredSpeedLimitBytesPerSecond.load(std::memory_order_acquire) : 0ull,
                                      havePostCycleSnapshot ? L"yes" : L"no",
                                      havePostCycleSnapshot && postCycleSnapshot.finished ? L"yes" : L"no",
                                      havePostCycleSnapshot ? postCycleSnapshot.taskId : 0ull,
                                      havePostCycleSnapshot ? postCycleSnapshot.desiredSpeedLimitBytesPerSecond : 0ull,
                                      expectedLimit));
        }

        const HWND lingeringPrompt = GetFileOperationsSpeedLimitPromptHandle();
        state.Require(lingeringPrompt == nullptr || IsWindow(lingeringPrompt) == FALSE,
                      std::format(L"Custom speed-limit prompt should not remain open after cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closePrompt();
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFileOperationsSpeedLimitPromptKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    Trace(L"SpeedLimitNavShell: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    auto* fileOps = g_folderWindow.DebugGetFileOperationState();
    state.Require(fileOps != nullptr, L"File-operations state unavailable for speed-limit shell-stability validation.");
    if (! fileOps)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root           = suiteRoot / L"work" / (L"fileops_speed_limit_nav_shell_" + NewGuidText());
    const std::filesystem::path sourceDir      = root / L"src";
    const std::filesystem::path destDir        = root / L"dst";
    const std::filesystem::path sourceFile     = sourceDir / L"payload.bin";
    const std::filesystem::path dummyRoot      = std::filesystem::path(L"/cmd-speedlimit-nav-shell") / SanitizeDummyPathSegmentForFileOpsPrompt(NewGuidText());
    const std::filesystem::path dummySourceDir = dummyRoot / L"src";
    const std::filesystem::path dummyDestDir   = dummyRoot / L"dst";
    constexpr uint64_t kInitialLimitBytesPerSecond = 64ull * 1024ull;

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sourceDir), L"Failed to create speed-limit shell-stability source directory.");
    state.Require(SelfTest::EnsureDirectory(destDir), L"Failed to create speed-limit shell-stability destination directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SelfTest::WriteTextFile(sourceFile, "payload"), L"Failed to create visible shell-stability payload.");
    if (! state.failure.empty())
    {
        return false;
    }

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    static_cast<void>(pluginManager.EnablePlugin(kBuiltinDummyFileSystemIdForFileOpsPrompt.data(), g_settings));
    const FileSystemPluginManager::PluginEntry* dummyEntry = FindFileSystemPluginById(kBuiltinDummyFileSystemIdForFileOpsPrompt);
    state.Require(dummyEntry && dummyEntry->fileSystem && dummyEntry->informations, L"Loaded dummy file system unavailable for speed-limit shell stability.");
    const wil::com_ptr<IFileSystem> dummyFileSystem = dummyEntry ? dummyEntry->fileSystem : nullptr;
    const wil::com_ptr<IInformations> dummyInfo     = dummyEntry ? dummyEntry->informations : nullptr;
    wil::com_ptr<IFileSystemIO> dummyIo;
    state.Require(CreateFileSystemIoForFileOpsPrompt(dummyFileSystem, dummyIo),
                  L"Loaded dummy file system missing IFileSystemIO for speed-limit shell stability.");

    constexpr std::string_view kSeedDummyConfig =
        R"json({"maxChildrenPerDirectory":42,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":0,"virtualSpeedLimit":"1048576"})json";
    constexpr std::string_view kSlowDummyConfig =
        R"json({"maxChildrenPerDirectory":42,"maxDepth":10,"seed":42,"latencyMs":0,"streamChunkLatencyMs":1000,"virtualSpeedLimit":"1048576"})json";
    constexpr uint64_t kDummyPayloadBytes                  = 8ull * 1024ull * 1024ull;
    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::FileOperationsSettings> previousFileOperations = g_settings.fileOperations;
    std::string previousDummyConfig;
    state.Require(BackupPluginConfigurationForFileOpsPrompt(dummyInfo.get(), previousDummyConfig),
                  L"Failed to snapshot dummy configuration for speed-limit shell stability.");
    if (! state.failure.empty())
    {
        return false;
    }
    const auto restoreSettings = wil::scope_exit([&] noexcept
    {
        g_settings.fileOperations = previousFileOperations;
        static_cast<void>(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), previousDummyConfig));
    });

    state.Require(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), kSeedDummyConfig),
                  L"Failed to apply deterministic dummy seed configuration for speed-limit shell stability.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummyRoot)),
                  L"Failed to create dummy root folder for speed-limit shell stability.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummySourceDir)),
                  L"Failed to create dummy source folder for speed-limit shell stability.");
    state.Require(EnsureDummyFolderExistsForFileOpsPrompt(dummyFileSystem.get(), ToPluginPathTextForFileOpsPrompt(dummyDestDir)),
                  L"Failed to create dummy destination folder for speed-limit shell stability.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::vector<std::filesystem::path> sourceFiles;
    sourceFiles.reserve(16u);
    for (size_t index = 0; index < 16u; ++index)
    {
        const std::filesystem::path filePath = index == 0u ? (dummySourceDir / L"payload.bin") : (dummySourceDir / std::format(L"payload_{:02}.bin", index));
        state.Require(WritePatternFileFsIoForFileOpsPrompt(dummyIo, filePath, kDummyPayloadBytes),
                      std::format(L"Failed to seed dummy shell-stability payload {}.", filePath.filename().native()));
        if (! state.failure.empty())
        {
            return false;
        }
        sourceFiles.push_back(filePath);
    }

    state.Require(SetPluginConfigurationForFileOpsPrompt(dummyInfo.get(), kSlowDummyConfig),
                  L"Failed to apply deterministic dummy latency for speed-limit shell stability.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closePrompt = []() noexcept
    {
        if (const HWND prompt = GetFileOperationsSpeedLimitPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            if (! DebugCancelFileOperationsSpeedLimitPrompt())
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
            }
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        closePrompt();
        static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));

        if (! leftPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        }
        if (! rightPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        }
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    closePrompt();
    static_cast<void>(CloseFileOperationsPopupForSelfTest(fileOps));
    {
        Common::Settings::FileOperationsSettings fileOperations = previousFileOperations.value_or(Common::Settings::FileOperationsSettings{});
        fileOperations.defaultBandwidthLimitBytesPerSecond      = kInitialLimitBytesPerSecond;
        g_settings.fileOperations                               = fileOperations;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for speed-limit shell-stability test (left pane).");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for speed-limit shell-stability test (right pane).");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, sourceDir);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, destDir);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sourceDir, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for speed-limit shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, destDir, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for speed-limit shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"payload.bin"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for speed-limit shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"SpeedLimitNavShell: panes ready");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"payload.bin"),
                  L"Failed to focus payload.bin before speed-limit shell-stability validation.");
    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder view handle unavailable for speed-limit shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"SpeedLimitNavShell: baseline navigation snapshots captured");

    NavigationViewDebugSnapshot leftBaselineSnapshot{};
    NavigationViewDebugSnapshot rightBaselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == sourceDir.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &leftBaselineSnapshot),
                  L"Failed to capture the baseline left navigation shell before speed-limit validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == destDir.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &rightBaselineSnapshot),
                  L"Failed to capture the baseline right navigation shell before speed-limit validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineLeftItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineLeftSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const size_t baselineRightSelectedCount = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right);

    const auto requireStableNavigationShells = [&](std::wstring_view context, const bool expectFocusRestore) noexcept
    {
        NavigationViewDebugSnapshot leftSnapshot{};
        NavigationViewDebugSnapshot rightSnapshot{};
        std::optional<std::filesystem::path> leftPath;
        std::optional<std::filesystem::path> rightPath;
        bool shellRestored  = false;
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();

            leftPath  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
            rightPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
            static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, leftSnapshot));
            static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Right, rightSnapshot));

            const bool leftPathStable  = leftPath.has_value() && OrdinalString::EqualsNoCasePath(leftPath.value(), sourceDir);
            const bool rightPathStable = rightPath.has_value() && OrdinalString::EqualsNoCasePath(rightPath.value(), destDir);
            const bool leftShellQuiet  = leftSnapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! leftSnapshot.editMode &&
                                         ! leftSnapshot.historyDropdownVisible && ! leftSnapshot.editSuggestPopupVisible &&
                                         ! leftSnapshot.fullPathPopupVisible && ! leftSnapshot.fullPathPopupEditMode &&
                                         leftSnapshot.visibleChildWindowCount == 0u;
            const bool rightShellQuiet = rightSnapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! rightSnapshot.editMode &&
                                         ! rightSnapshot.historyDropdownVisible && ! rightSnapshot.editSuggestPopupVisible &&
                                         ! rightSnapshot.fullPathPopupVisible && ! rightSnapshot.fullPathPopupEditMode &&
                                         rightSnapshot.visibleChildWindowCount == 0u;
            const bool selectionStable = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineLeftSelectedCount &&
                                         g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == baselineRightSelectedCount;
            const bool leftItemsStable = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineLeftItemCount;
            const bool leftFocusStable = ! expectFocusRestore || (g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView &&
                                                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"payload.bin");
            if (leftPathStable && rightPathStable && leftShellQuiet && rightShellQuiet && selectionStable && leftItemsStable && leftFocusStable)
            {
                shellRestored = true;
                break;
            }

            Sleep(20);
        }

        state.Require(
            shellRestored,
            std::format(L"Navigation shells did not stay stable after {}; leftPath='{}', rightPath='{}', leftFocusTarget={}, rightFocusTarget={}, "
                        L"leftEditMode={}, rightEditMode={}, leftHistoryVisible={}, rightHistoryVisible={}, leftSuggestVisible={}, rightSuggestVisible={}, "
                        L"leftPopupVisible={}, rightPopupVisible={}, leftChildren={}, rightChildren={}, leftItemCount={}, leftSelectedCount={}, "
                        L"rightSelectedCount={}, focusedItem='{}', focusedFolderView=0x{:X}.",
                        context,
                        leftPath.has_value() ? leftPath->wstring() : std::wstring{},
                        rightPath.has_value() ? rightPath->wstring() : std::wstring{},
                        static_cast<unsigned>(leftSnapshot.focusTarget),
                        static_cast<unsigned>(rightSnapshot.focusTarget),
                        leftSnapshot.editMode ? L"yes" : L"no",
                        rightSnapshot.editMode ? L"yes" : L"no",
                        leftSnapshot.historyDropdownVisible ? L"yes" : L"no",
                        rightSnapshot.historyDropdownVisible ? L"yes" : L"no",
                        leftSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                        rightSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                        leftSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                        rightSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                        leftSnapshot.visibleChildWindowCount,
                        rightSnapshot.visibleChildWindowCount,
                        g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                        g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                        g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right),
                        g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                        static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd()))));
    };

    const auto requireAnchoredNavigationShells = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot leftSnapshot{};
        NavigationViewDebugSnapshot rightSnapshot{};
        std::optional<std::filesystem::path> leftPath;
        std::optional<std::filesystem::path> rightPath;
        bool shellAnchored  = false;
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(750ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();

            leftPath  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
            rightPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
            static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, leftSnapshot));
            static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Right, rightSnapshot));

            const bool leftPathStable  = leftPath.has_value() && OrdinalString::EqualsNoCasePath(leftPath.value(), sourceDir);
            const bool rightPathStable = rightPath.has_value() && OrdinalString::EqualsNoCasePath(rightPath.value(), destDir);
            const bool leftShellQuiet  = ! leftSnapshot.editMode && ! leftSnapshot.historyDropdownVisible && ! leftSnapshot.editSuggestPopupVisible &&
                                         ! leftSnapshot.fullPathPopupVisible && ! leftSnapshot.fullPathPopupEditMode;
            const bool rightShellQuiet = ! rightSnapshot.editMode && ! rightSnapshot.historyDropdownVisible && ! rightSnapshot.editSuggestPopupVisible &&
                                         ! rightSnapshot.fullPathPopupVisible && ! rightSnapshot.fullPathPopupEditMode;
            if (leftPathStable && rightPathStable && leftShellQuiet && rightShellQuiet)
            {
                shellAnchored = true;
                break;
            }

            Sleep(20);
        }

        state.Require(shellAnchored,
                      std::format(L"Navigation shells did not stay anchored after {}; leftPath='{}', rightPath='{}', leftEditMode={}, rightEditMode={}, "
                                  L"leftHistoryVisible={}, rightHistoryVisible={}, leftSuggestVisible={}, rightSuggestVisible={}, "
                                  L"leftPopupVisible={}, rightPopupVisible={}.",
                                  context,
                                  leftPath.has_value() ? leftPath->wstring() : std::wstring{},
                                  rightPath.has_value() ? rightPath->wstring() : std::wstring{},
                                  leftSnapshot.editMode ? L"yes" : L"no",
                                  rightSnapshot.editMode ? L"yes" : L"no",
                                  leftSnapshot.historyDropdownVisible ? L"yes" : L"no",
                                  rightSnapshot.historyDropdownVisible ? L"yes" : L"no",
                                  leftSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  rightSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  leftSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  rightSnapshot.fullPathPopupVisible ? L"yes" : L"no"));
    };

    const wil::com_ptr<IFileSystem> leftFileSystem  = dummyFileSystem;
    const wil::com_ptr<IFileSystem> rightFileSystem = dummyFileSystem;
    state.Require(leftFileSystem && rightFileSystem, L"Failed to resolve dummy file-system interfaces for speed-limit shell-stability test.");
    if (! leftFileSystem || ! rightFileSystem)
    {
        return false;
    }
    HWND popup = nullptr;
    std::optional<uint64_t> taskId;
    FileOperationsPopupInternal::PopupSelfTestInvoke openPrompt{};
    const auto startTrackedSpeedLimitTask = [&](std::wstring_view context) noexcept -> bool
    {
        std::vector<FolderWindow::FileOperationState::Task*> beforeTasks;
        fileOps->CollectTasks(beforeTasks);
        std::unordered_set<uint64_t> existingTaskIds;
        existingTaskIds.reserve(beforeTasks.size());
        for (auto* task : beforeTasks)
        {
            if (task)
            {
                existingTaskIds.insert(task->GetId());
            }
        }

        const HRESULT startHr = fileOps->StartOperation(FILESYSTEM_COPY,
                                                        FolderWindow::Pane::Left,
                                                        FolderWindow::Pane::Right,
                                                        leftFileSystem,
                                                        sourceFiles,
                                                        dummyDestDir,
                                                        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE),
                                                        false,
                                                        kInitialLimitBytesPerSecond,
                                                        // Keep this observable between items; bulk dummy copies can finish before
                                                        // the full Commands sweep reaches the prompt readiness checks.
                                                        FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                        false,
                                                        rightFileSystem);
        state.Require(SUCCEEDED(startHr),
                      std::format(L"Failed to start file-operations speed-limit shell-stability test copy for {} (hr=0x{:08X}).",
                                  context,
                                  static_cast<unsigned long>(startHr)));
        if (! state.failure.empty())
        {
            return false;
        }
        Trace(std::format(L"SpeedLimitNavShell: copy started for {}", context));

        taskId = ResolveAndPauseNewFileOperationsTaskForSelfTest(fileOps, existingTaskIds, kInitialLimitBytesPerSecond);
        if (taskId.has_value())
        {
            Trace(std::format(L"SpeedLimitNavShell: task {} resolved and paused early for {}", taskId.value(), context));
        }

        popup = WaitForWindow([&]() noexcept { return fileOps->GetPopupHwndForSelfTest(); }, SelfTest::Scale(5000ms));
        state.Require(popup != nullptr && IsWindow(popup) != FALSE,
                      std::format(L"File-operations popup did not open for speed-limit shell-stability test during {}.", context));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return false;
        }
        Trace(std::format(L"SpeedLimitNavShell: popup=0x{:X} for {}", reinterpret_cast<uintptr_t>(popup), context));

        bool taskReady = false;
        std::wstring readinessDebug;
        const auto readyDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        if (! taskId.has_value())
        {
            taskId = ResolveNewFileOperationsTaskIdForSelfTest(fileOps, existingTaskIds, SelfTest::Scale(250ms));
        }
        while (std::chrono::steady_clock::now() < readyDeadline)
        {
            PumpPendingMessages();

            auto* task = taskId.has_value() ? fileOps->FindTask(taskId.value()) : nullptr;
            FileOperationsPopupInternal::TaskSnapshot popupTaskSnapshot{};
            const bool havePopupTaskSnapshot = DebugGetFileOperationsPopupTaskSnapshot(popup, taskId.value_or(0), popupTaskSnapshot);
            std::vector<FolderWindow::FileOperationState::Task*> liveTasks;
            std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> completedTasks;
            fileOps->CollectTasks(liveTasks);
            fileOps->CollectCompletedTasks(completedTasks);
            readinessDebug = std::format(L"trackedTaskId={}, liveTask={}, liveTaskCount={}, completedTaskCount={}, popupSnapshot={}, popupTaskId={}, "
                                         L"popupFinished={}, popupStarted={}, popupPaused={}, popupWaitingInQueue={}, popupQueuePaused={}, popupLimit={}.",
                                         taskId.value_or(0),
                                         task ? L"yes" : L"no",
                                         liveTasks.size(),
                                         completedTasks.size(),
                                         havePopupTaskSnapshot ? L"yes" : L"no",
                                         havePopupTaskSnapshot ? popupTaskSnapshot.taskId : 0ull,
                                         havePopupTaskSnapshot && popupTaskSnapshot.finished ? L"yes" : L"no",
                                         havePopupTaskSnapshot && popupTaskSnapshot.started ? L"yes" : L"no",
                                         havePopupTaskSnapshot && popupTaskSnapshot.paused ? L"yes" : L"no",
                                         havePopupTaskSnapshot && popupTaskSnapshot.waitingInQueue ? L"yes" : L"no",
                                         havePopupTaskSnapshot && popupTaskSnapshot.queuePaused ? L"yes" : L"no",
                                         havePopupTaskSnapshot ? popupTaskSnapshot.desiredSpeedLimitBytesPerSecond : 0ull);
            if (havePopupTaskSnapshot && popupTaskSnapshot.taskId != 0)
            {
                taskId = popupTaskSnapshot.taskId;
                task   = fileOps->FindTask(taskId.value());
            }

            if (task)
            {
                taskReady = true;
                task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
                break;
            }

            if (havePopupTaskSnapshot && popupTaskSnapshot.taskId != 0 && ! popupTaskSnapshot.finished)
            {
                taskReady = true;
                if (task)
                {
                    task->SetDesiredSpeedLimit(kInitialLimitBytesPerSecond);
                }
                break;
            }

            std::this_thread::sleep_for(20ms);
        }
        state.Require(taskReady,
                      std::format(L"File-operations speed-limit shell-stability task did not become available in time for {}. {}", context, readinessDebug));
        state.Require(taskId.has_value(),
                      std::format(L"Failed to identify the new file-operations task for speed-limit shell-stability test during {}.", context));
        if (! taskReady || ! taskId.has_value())
        {
            return false;
        }
        Trace(std::format(L"SpeedLimitNavShell: task {} ready for {}", taskId.value(), context));

        FileOperationsPopupInternal::TaskSnapshot heldTaskSnapshot{};
        if (RefreshTrackedFileOperationsTaskForSelfTest(fileOps, popup, taskId, &heldTaskSnapshot))
        {
            if (auto* task = fileOps->FindTask(taskId.value()); task && task->HasStarted() && ! task->IsPaused())
            {
                task->TogglePause();
            }

            const auto holdDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < holdDeadline)
            {
                PumpPendingMessages();
                if (auto* task = fileOps->FindTask(taskId.value());
                    task && (task->IsWaitingForOthers() || task->IsWaitingInQueue() || task->IsPaused() || task->IsQueuePaused()))
                {
                    Trace(std::format(L"SpeedLimitNavShell: task held for {}", context));
                    break;
                }
                std::this_thread::sleep_for(20ms);
            }
        }

        openPrompt        = {};
        openPrompt.kind   = FileOperationsPopupInternal::PopupHitTest::Kind::TaskSpeedLimit;
        openPrompt.taskId = taskId.value();
        openPrompt.data   = 1u;
        return true;
    };

    if (! startTrackedSpeedLimitTask(L"cancel pass"))
    {
        return false;
    }

    const auto runOwnedPromptCycle = [&](std::wstring_view context, FileOperationsSpeedLimitPromptDebugSnapshot& snapshot, auto&& workerFunc) noexcept -> bool
    {
        const bool requestSpecificTaskId = openPrompt.taskId != 0;
        FileOperationsPopupInternal::TaskSnapshot liveTaskSnapshot{};
        static_cast<void>(RefreshTrackedFileOperationsTaskForSelfTest(fileOps, popup, taskId, &liveTaskSnapshot));
        if (requestSpecificTaskId && taskId.has_value())
        {
            openPrompt.taskId = taskId.value();
        }

        Trace(std::format(L"SpeedLimitNavShell: invoking prompt for {}", context));
        return RunFileOperationsSpeedLimitPromptModalCycle(popup,
                                                           openPrompt,
                                                           [&](const HWND prompt) noexcept
        {
            FileOperationsPopupInternal::TaskSnapshot taskSnapshot{};
            const bool haveTaskSnapshot = DebugGetFileOperationsPopupTaskSnapshot(popup, openPrompt.taskId, taskSnapshot);
            const bool haveTrackedTask  = taskId.has_value() && fileOps->FindTask(taskId.value()) != nullptr;
            state.Require(prompt != nullptr && IsWindow(prompt) != FALSE,
                          std::format(L"Custom speed-limit prompt did not open during {} (invokeTaskId={}, trackedTaskId={}, trackedLive={}, "
                                      L"popupTaskFound={}, popupTaskId={}, popupTaskFinished={}, popupTaskLimit={}, popupCurrent=0x{:X}, livePopup=0x{:X}).",
                                      context,
                                      openPrompt.taskId,
                                      taskId.value_or(0),
                                      haveTrackedTask ? L"yes" : L"no",
                                      haveTaskSnapshot ? L"yes" : L"no",
                                      haveTaskSnapshot ? taskSnapshot.taskId : 0ull,
                                      haveTaskSnapshot && taskSnapshot.finished ? L"yes" : L"no",
                                      haveTaskSnapshot ? taskSnapshot.desiredSpeedLimitBytesPerSecond : 0ull,
                                      static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(popup)),
                                      static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(fileOps->GetPopupHwndForSelfTest()))));
            state.Require(prompt == nullptr || IsOwnedBy(prompt, popup),
                          std::format(L"Custom speed-limit prompt should be owned by the file-operations popup during {}.", context));
            if (! prompt || IsWindow(prompt) == FALSE || ! state.failure.empty())
            {
                return;
            }

            Trace(std::format(L"SpeedLimitNavShell: prompt=0x{:X} opened for {}", reinterpret_cast<uintptr_t>(prompt), context));

            bool settled              = false;
            const auto promptDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < promptDeadline)
            {
                snapshot = {};
                if (DebugGetFileOperationsSpeedLimitPromptSnapshot(snapshot) && snapshot.usesDxUiHost && snapshot.visibleChildWindowCount <= 1u &&
                    ! snapshot.text.empty())
                {
                    settled = true;
                    break;
                }
                Sleep(10);
            }

            state.Require(settled,
                          std::format(L"Custom speed-limit prompt did not settle during {}; usesDxUiHost={}, visibleChildren={}, text='{}'.",
                                      context,
                                      snapshot.usesDxUiHost ? L"yes" : L"no",
                                      snapshot.visibleChildWindowCount,
                                      snapshot.text));
            if (! settled || ! state.failure.empty())
            {
                return;
            }

            Trace(std::format(L"SpeedLimitNavShell: prompt settled for {} text='{}'", context, snapshot.text));
            workerFunc(prompt);
        });
    };

    FileOperationsSpeedLimitPromptDebugSnapshot cancelSnapshot{};
    const bool cancelInvoked = runOwnedPromptCycle(L"cancel pass",
                                                   cancelSnapshot,
                                                   [&](const HWND prompt) noexcept
    {
        state.Require(DebugCancelFileOperationsSpeedLimitPrompt(), L"Failed to cancel the custom speed-limit prompt during shell-stability validation.");
        state.Require(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)),
                      L"Custom speed-limit prompt did not close after cancel during shell-stability validation.");
    });
    state.Require(cancelInvoked, L"Failed to request the custom speed-limit prompt during the cancel shell-stability pass.");
    if (! cancelInvoked || ! state.failure.empty())
    {
        return false;
    }
    Trace(L"SpeedLimitNavShell: cancel pass closed");

    requireAnchoredNavigationShells(L"speed-limit cancel");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"SpeedLimitNavShell: cancel pass shell stable");

    state.Require(CloseFileOperationsPopupForSelfTest(fileOps), L"File-operations popup did not close after cancel shell-stability validation.");
    state.Require(GetFileOperationsSpeedLimitPromptHandle() == nullptr,
                  L"Custom speed-limit prompt should not remain open after cancel shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"SpeedLimitNavShell: popup cleanup closed");

    requireStableNavigationShells(L"speed-limit cleanup", false);
    if (state.failure.empty())
    {
        Trace(L"SpeedLimitNavShell: done");
    }
    return state.failure.empty();
}

} // namespace (tests)

void RunFileOpsCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_uses_dxui_host_without_visible_child_controls", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneUsesDxUiHostWithNoVisibleChildControls(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_diff_refresh_preserves_selection", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneDiffRefreshPreservesSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_header_click_sorts_results", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneHeaderClickSortsResults(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_header_resize_changes_visible_width", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneHeaderResizeChangesVisibleWidth(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_header_drag_reorders_columns_without_sort", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneHeaderDragReordersColumnsWithoutSort(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_copy_follows_reordered_columns", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneCopyFollowsReorderedColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_reordered_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneReorderedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_reordered_copy_follows_visible_columns_after_sort_cycles", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneReorderedCopyFollowsVisibleColumnsAfterSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_resized_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneResizedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_reordered_resized_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneReorderedResizedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_fileops_issues_pane_reordered_resized_copy_follows_visible_columns_after_sort_cycles", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneReorderedResizedCopyFollowsVisibleColumnsAfterSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_restores_combined_view_state_after_recreate", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneRestoresCombinedViewStateAfterRecreate(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_exposes_live_uia_selection", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneExposesLiveUiaSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_tab_keeps_grid_focus", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneTabTraversalKeepsGridFocus(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_long_run_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneLongRunScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_issues_pane_theme_cycle_keeps_grid_legible", [=](CaseState& state) noexcept {
        return TestFileOperationsIssuesPaneThemeCycleKeepsGridLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestFileOperationsSpeedLimitPromptUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_speedLimit_prompt_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestFileOperationsSpeedLimitPromptLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestFileOperationsSpeedLimitPromptLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_fileops_speedLimit_prompt_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestFileOperationsSpeedLimitPromptKeepsNavigationShellStable(mainWindow, state);
    });
}

namespace
{
