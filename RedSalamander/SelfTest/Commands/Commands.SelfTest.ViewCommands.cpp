// Commands.SelfTest.ViewCommands.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// ViewCommands test family: 28 test functions.

[[nodiscard]] bool WaitForMainMenuBarVisibility(HWND mainWindow, bool expectedVisible, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        if (DebugIsMainMenuBarSurfaceVisible(mainWindow) == expectedVisible)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return DebugIsMainMenuBarSurfaceVisible(mainWindow) == expectedVisible;
}

[[nodiscard]] std::wstring DescribeWindowHandleForSelfTest(HWND hwnd)
{
    if (! hwnd)
    {
        return L"null";
    }

    std::array<wchar_t, 128> className{};
    static_cast<void>(GetClassNameW(hwnd, className.data(), static_cast<int>(className.size())));
    return std::format(L"0x{:X} class='{}' visible={}",
                       reinterpret_cast<uintptr_t>(hwnd),
                       className.data(),
                       IsWindowVisible(hwnd) != FALSE ? L"yes" : L"no");
}

[[nodiscard]] std::wstring DescribeWindowHandleAndRectForSelfTest(HWND hwnd)
{
    if (! hwnd)
    {
        return L"null";
    }

    std::array<wchar_t, 128> className{};
    static_cast<void>(GetClassNameW(hwnd, className.data(), static_cast<int>(className.size())));
    RECT rect{};
    const bool isWindow = IsWindow(hwnd) != FALSE;
    const bool gotRect  = isWindow && GetWindowRect(hwnd, &rect) != FALSE;
    return std::format(L"0x{:X} isWindow={} class='{}' visible={} rect=({},{}-{},{} )",
                       reinterpret_cast<uintptr_t>(hwnd),
                       isWindow ? L"yes" : L"no",
                       className.data(),
                       IsWindowVisible(hwnd) != FALSE ? L"yes" : L"no",
                       gotRect ? rect.left : 0,
                       gotRect ? rect.top : 0,
                       gotRect ? rect.right : 0,
                       gotRect ? rect.bottom : 0);
}

[[nodiscard]] std::wstring FormatRectForSelfTest(const RECT& rect)
{
    return std::format(L"({},{}-{},{} )", rect.left, rect.top, rect.right, rect.bottom);
}

[[nodiscard]] std::wstring DescribeNavigationViewVisibilityForSelfTest(FolderWindow::Pane pane, HWND expectedNavigationView)
{
    NavigationViewDebugSnapshot snapshot{};
    const bool snapshotOk = g_folderWindow.DebugGetNavigationViewSnapshot(pane, snapshot);
    const std::optional<std::filesystem::path> panePath = g_folderWindow.GetCurrentPath(pane);
    const HWND currentNavigationView                    = g_folderWindow.DebugGetNavigationViewHwnd(pane);
    const std::optional<FolderWindow::Pane> zoomedPane  = g_folderWindow.GetZoomedPane();
    RECT mainClient{};
    const bool mainClientOk = GetClientRect(g_folderWindow.GetHwnd(), &mainClient) != FALSE;
    FolderWindow::PaneViewOptionsDebugSnapshot paneOptions{};
    const bool paneOptionsOk = g_folderWindow.DebugGetPaneViewOptionsSnapshot(pane, paneOptions);
    FolderWindow::PreviewPaneDebugSnapshot preview{};
    const bool previewOk = g_folderWindow.DebugGetPreviewPaneSnapshot(preview);
    FolderWindow::CommandLineDebugSnapshot commandLine{};
    const bool commandLineOk = g_folderWindow.DebugGetCommandLineSnapshot(commandLine);
    FolderWindow::FolderWindowFunctionBarDebugSnapshot functionBar{};
    const bool functionBarOk = g_folderWindow.DebugGetFunctionBarSnapshot(functionBar);
    const HWND paneFolderView = g_folderWindow.GetFolderViewHwnd(pane);

    return std::format(L"pane={}, storedVisible={}, currentNav={}, expectedNav={}, activePane={}, focusedWindow=0x{:X}, "
                       L"focusedFolderView=0x{:X}, leftNavVisible={}, rightNavVisible={}, panePath='{}', snapshotOk={}, "
                       L"snapshotPath='{}', focusTarget={}, editMode={}, visibleChildren={}, zoomedPane={}, mainClientOk={}, mainClient={}, "
                       L"paneFolderView={}, paneOptionsOk={}, paneOptionsNavWindowVisible={}, paneOptionsFilterVisible={}, previewOk={}, "
                       L"previewActive={}, previewSource={}, previewHost={}, previewTabSelected={}, folderTabSelected={}, folderViewVisible={}, "
                       L"previewClient={}, previewContent={}, previewFunctionBar={}, commandLineOk={}, commandLineVisible={}, commandLinePane={}, "
                       L"functionBarOk={}, functionBarVisible={}, functionBarWindowVisible={}, functionBarRect={}",
                       pane == FolderWindow::Pane::Left ? L"left" : L"right",
                       g_folderWindow.GetNavigationBarVisible(pane) ? L"yes" : L"no",
                       DescribeWindowHandleAndRectForSelfTest(currentNavigationView),
                       DescribeWindowHandleAndRectForSelfTest(expectedNavigationView),
                       g_folderWindow.GetActivePane() == FolderWindow::Pane::Left ? L"left" : L"right",
                       reinterpret_cast<uintptr_t>(GetFocus()),
                       reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd()),
                       g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left) ? L"yes" : L"no",
                       g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Right) ? L"yes" : L"no",
                       panePath.has_value() ? panePath.value().wstring() : L"",
                       snapshotOk ? L"yes" : L"no",
                       snapshotOk ? snapshot.currentPathText : L"",
                       snapshotOk ? static_cast<int>(snapshot.focusTarget) : -1,
                       snapshotOk && snapshot.editMode ? L"yes" : L"no",
                       snapshotOk ? snapshot.visibleChildWindowCount : 0u,
                       zoomedPane.has_value() ? (zoomedPane.value() == FolderWindow::Pane::Left ? L"left" : L"right") : L"none",
                       mainClientOk ? L"yes" : L"no",
                       FormatRectForSelfTest(mainClient),
                       DescribeWindowHandleAndRectForSelfTest(paneFolderView),
                       paneOptionsOk ? L"yes" : L"no",
                       paneOptionsOk && paneOptions.navigationViewWindowVisible ? L"yes" : L"no",
                       paneOptionsOk && paneOptions.filterBarVisible ? L"yes" : L"no",
                       previewOk ? L"yes" : L"no",
                       previewOk && preview.active ? L"yes" : L"no",
                       previewOk ? (preview.sourcePane == FolderWindow::Pane::Left ? L"left" : L"right") : L"unknown",
                       previewOk ? (preview.hostPane == FolderWindow::Pane::Left ? L"left" : L"right") : L"unknown",
                       previewOk && preview.previewTabSelected ? L"yes" : L"no",
                       previewOk && preview.folderTabSelected ? L"yes" : L"no",
                       previewOk && preview.folderViewVisible ? L"yes" : L"no",
                       previewOk ? FormatRectForSelfTest(preview.clientRect) : L"",
                       previewOk ? FormatRectForSelfTest(preview.contentRect) : L"",
                       previewOk ? FormatRectForSelfTest(preview.functionBarRect) : L"",
                       commandLineOk ? L"yes" : L"no",
                       commandLineOk && commandLine.visible ? L"yes" : L"no",
                       commandLineOk ? (commandLine.pane == FolderWindow::Pane::Left ? L"left" : L"right") : L"unknown",
                       functionBarOk ? L"yes" : L"no",
                       functionBarOk && functionBar.visible ? L"yes" : L"no",
                       functionBarOk && functionBar.windowVisible ? L"yes" : L"no",
                       functionBarOk ? FormatRectForSelfTest(functionBar.rect) : L"");
}

[[nodiscard]] bool ColorNear(COLORREF actual, COLORREF expected, int tolerance) noexcept
{
    return std::abs(static_cast<int>(GetRValue(actual)) - static_cast<int>(GetRValue(expected))) <= tolerance &&
           std::abs(static_cast<int>(GetGValue(actual)) - static_cast<int>(GetGValue(expected))) <= tolerance &&
           std::abs(static_cast<int>(GetBValue(actual)) - static_cast<int>(GetBValue(expected))) <= tolerance;
}

struct WindowColorScan final
{
    bool isWindow           = false;
    bool visible            = false;
    bool gotClient          = false;
    int width               = 0;
    int height              = 0;
    int matchedX            = -1;
    int matchedY            = -1;
    int differentX          = -1;
    int differentY          = -1;
    COLORREF topLeft        = CLR_INVALID;
    COLORREF center         = CLR_INVALID;
    COLORREF firstDifferent = CLR_INVALID;
};

[[nodiscard]] std::wstring DescribeColor(COLORREF color)
{
    if (color == CLR_INVALID)
    {
        return L"invalid";
    }

    return std::format(
        L"#{:02X}{:02X}{:02X}", static_cast<unsigned>(GetRValue(color)), static_cast<unsigned>(GetGValue(color)), static_cast<unsigned>(GetBValue(color)));
}

#pragma warning(push)
#pragma warning(disable : 4625 4626)
[[nodiscard]] bool WindowContainsApproxColor(HWND hwnd, COLORREF expectedColor, WindowColorScan* outScan = nullptr, int tolerance = 3) noexcept
{
    if (outScan)
    {
        outScan->isWindow = hwnd && IsWindow(hwnd) != FALSE;
        outScan->visible  = outScan->isWindow && IsWindowVisible(hwnd) != FALSE;
    }

    if (! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
    {
        return false;
    }

    RECT client{};
    if (! GetClientRect(hwnd, &client))
    {
        return false;
    }
    if (outScan)
    {
        outScan->gotClient = true;
    }

    const int width  = std::max(0L, client.right - client.left);
    const int height = std::max(0L, client.bottom - client.top);
    if (outScan)
    {
        outScan->width  = width;
        outScan->height = height;
    }
    if (width <= 0 || height <= 0)
    {
        return false;
    }

    if (const HWND root = GetAncestor(hwnd, GA_ROOT))
    {
        ShowWindow(root, SW_SHOWNORMAL);
        BringWindowToTop(root);
        SetForegroundWindow(root);
        SetWindowPos(root, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        SetWindowPos(root, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        UpdateWindow(root);
    }

    RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PumpPendingMessages();

    HDC hdc = GetDC(hwnd);
    if (! hdc)
    {
        return false;
    }
    const auto releaseDc = wil::scope_exit([&]() noexcept { ReleaseDC(hwnd, hdc); });

    if (outScan)
    {
        outScan->topLeft = GetPixel(hdc, 0, 0);
        outScan->center  = GetPixel(hdc, width / 2, height / 2);
    }

    for (int y = 0; y < height; ++y)
    {
        for (int x = 0; x < width; ++x)
        {
            const COLORREF px = GetPixel(hdc, x, y);
            if (px != CLR_INVALID && ColorNear(px, expectedColor, tolerance))
            {
                if (outScan)
                {
                    outScan->matchedX = x;
                    outScan->matchedY = y;
                }
                return true;
            }
            if (outScan && outScan->firstDifferent == CLR_INVALID && px != CLR_INVALID && ! ColorNear(px, expectedColor, tolerance))
            {
                outScan->firstDifferent = px;
                outScan->differentX     = x;
                outScan->differentY     = y;
            }
        }
    }

    return false;
}

[[nodiscard]] bool WindowContainsPixelDifferentFromColor(HWND hwnd, COLORREF backgroundColor) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
    {
        return false;
    }

    RECT client{};
    if (! GetClientRect(hwnd, &client))
    {
        return false;
    }

    const int width  = std::max(0L, client.right - client.left);
    const int height = std::max(0L, client.bottom - client.top);
    if (width <= 0 || height <= 0)
    {
        return false;
    }

    if (const HWND root = GetAncestor(hwnd, GA_ROOT))
    {
        ShowWindow(root, SW_SHOWNORMAL);
        BringWindowToTop(root);
        SetForegroundWindow(root);
        SetWindowPos(root, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        SetWindowPos(root, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        UpdateWindow(root);
    }

    RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PumpPendingMessages();

    HDC hdc = GetDC(hwnd);
    if (! hdc)
    {
        return false;
    }
    const auto releaseDc = wil::scope_exit([&]() noexcept { ReleaseDC(hwnd, hdc); });

    constexpr int kBackgroundTolerance = 10;
    for (int y = 0; y < height; ++y)
    {
        for (int x = 0; x < width; ++x)
        {
            const COLORREF px = GetPixel(hdc, x, y);
            if (px != CLR_INVALID && ! ColorNear(px, backgroundColor, kBackgroundTolerance))
            {
                return true;
            }
        }
    }

    return false;
}
#pragma warning(pop)

[[nodiscard]] HWND FindVisibleDxUiContextMenuWindow() noexcept
{
    const HWND hwnd = FindWindowW(L"DxUi_ContextMenu", nullptr);
    return (hwnd && IsWindowVisible(hwnd) != FALSE) ? hwnd : nullptr;
}

[[nodiscard]] std::vector<HWND> FindVisibleOwnedDxUiContextMenuWindows(HWND ownerHwnd) noexcept
{
    std::vector<HWND> windows;
    const HWND rootOwner = ownerHwnd ? GetAncestor(ownerHwnd, GA_ROOT) : nullptr;
    for (HWND popup = FindWindowW(L"DxUi_ContextMenu", nullptr); popup != nullptr; popup = FindWindowExW(nullptr, popup, L"DxUi_ContextMenu", nullptr))
    {
        if (IsWindowVisible(popup) == FALSE)
        {
            continue;
        }

        if (ownerHwnd)
        {
            const HWND popupOwner = GetWindow(popup, GW_OWNER);
            if (popupOwner != ownerHwnd && popupOwner != rootOwner)
            {
                continue;
            }
        }

        windows.push_back(popup);
    }

    return windows;
}

[[nodiscard]] HWND FindVisibleOwnedDxUiContextMenuWindow(HWND ownerHwnd) noexcept
{
    std::vector<HWND> windows = FindVisibleOwnedDxUiContextMenuWindows(ownerHwnd);
    return windows.empty() ? nullptr : windows.front();
}

[[nodiscard]] bool WaitForOwnedDxUiContextMenuWindowCount(HWND ownerHwnd,
                                                          size_t expectedCount,
                                                          std::chrono::milliseconds timeout,
                                                          std::vector<HWND>* outWindows = nullptr) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        std::vector<HWND> windows = FindVisibleOwnedDxUiContextMenuWindows(ownerHwnd);
        if (windows.size() >= expectedCount)
        {
            if (outWindows)
            {
                *outWindows = std::move(windows);
            }
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    std::vector<HWND> windows = FindVisibleOwnedDxUiContextMenuWindows(ownerHwnd);
    if (outWindows)
    {
        *outWindows = std::move(windows);
    }
    return outWindows ? outWindows->size() >= expectedCount : windows.size() >= expectedCount;
}

struct UiMenuRoundTripResult
{
    bool opened = false;
    bool closed = false;
};

[[nodiscard]] UiMenuRoundTripResult TriggerAndDismissKeyboardActivatedUiMenu(HWND targetHwnd,
                                                                             WPARAM virtualKey,
                                                                             DWORD uiThreadId,
                                                                             HWND ownerHwnd,
                                                                             std::chrono::milliseconds openTimeout,
                                                                             std::chrono::milliseconds closeTimeout) noexcept
{
    using namespace std::chrono_literals;

    std::atomic<bool> menuOpened{false};
    std::atomic<bool> menuClosed{false};

    {
        std::jthread closer([&](std::stop_token stopToken) noexcept
        {
            const auto openDeadline = std::chrono::steady_clock::now() + openTimeout;
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < openDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
                const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(ownerHwnd);
                if (inMenuMode || popup != nullptr)
                {
                    menuOpened.store(true, std::memory_order_release);
                    break;
                }

                std::this_thread::sleep_for(10ms);
            }

            if (! menuOpened.load(std::memory_order_acquire))
            {
                return;
            }

            const auto closeDeadline = std::chrono::steady_clock::now() + closeTimeout;
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
                const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(ownerHwnd);
                if (! inMenuMode && popup == nullptr)
                {
                    menuClosed.store(true, std::memory_order_release);
                    return;
                }

                const HWND dismissTarget =
                    popup != nullptr ? popup
                                     : (hasGuiInfo && gti.hwndMenuOwner ? gti.hwndMenuOwner : (hasGuiInfo && gti.hwndActive ? gti.hwndActive : ownerHwnd));
                if (dismissTarget != nullptr)
                {
                    PostMessageW(dismissTarget, WM_KEYDOWN, VK_ESCAPE, 0);
                    PostMessageW(dismissTarget, WM_KEYUP, VK_ESCAPE, 0);
                }

                std::this_thread::sleep_for(30ms);
            }

            GUITHREADINFO gti{};
            gti.cbSize            = sizeof(gti);
            const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
            const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
            const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(ownerHwnd);
            if (! inMenuMode && popup == nullptr)
            {
                menuClosed.store(true, std::memory_order_release);
            }
        });

        SendMessageW(targetHwnd, WM_KEYDOWN, virtualKey, 0);
        SendMessageW(targetHwnd, WM_KEYUP, virtualKey, 0);
        PumpPendingMessages();
    }

    return UiMenuRoundTripResult{
        .opened = menuOpened.load(std::memory_order_acquire),
        .closed = menuClosed.load(std::memory_order_acquire),
    };
}

[[nodiscard]] UiMenuRoundTripResult TriggerAndDismissCommandActivatedUiMenu(
    HWND targetHwnd, UINT commandId, DWORD uiThreadId, HWND ownerHwnd, std::chrono::milliseconds openTimeout, std::chrono::milliseconds closeTimeout) noexcept
{
    using namespace std::chrono_literals;

    std::atomic<bool> menuOpened{false};
    std::atomic<bool> menuClosed{false};

    {
        std::jthread closer([&](std::stop_token stopToken) noexcept
        {
            const auto openDeadline = std::chrono::steady_clock::now() + openTimeout;
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < openDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
                const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(ownerHwnd);
                if (inMenuMode || popup != nullptr)
                {
                    menuOpened.store(true, std::memory_order_release);
                    break;
                }

                std::this_thread::sleep_for(10ms);
            }

            if (! menuOpened.load(std::memory_order_acquire))
            {
                return;
            }

            const auto closeDeadline = std::chrono::steady_clock::now() + closeTimeout;
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
                const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(ownerHwnd);
                if (! inMenuMode && popup == nullptr)
                {
                    menuClosed.store(true, std::memory_order_release);
                    return;
                }

                const HWND dismissTarget =
                    popup != nullptr ? popup
                                     : (hasGuiInfo && gti.hwndMenuOwner ? gti.hwndMenuOwner : (hasGuiInfo && gti.hwndActive ? gti.hwndActive : ownerHwnd));
                if (dismissTarget != nullptr)
                {
                    PostMessageW(dismissTarget, WM_KEYDOWN, VK_ESCAPE, 0);
                    PostMessageW(dismissTarget, WM_KEYUP, VK_ESCAPE, 0);
                }

                std::this_thread::sleep_for(30ms);
            }

            GUITHREADINFO gti{};
            gti.cbSize            = sizeof(gti);
            const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
            const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
            const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(ownerHwnd);
            if (! inMenuMode && popup == nullptr)
            {
                menuClosed.store(true, std::memory_order_release);
            }
        });

        SendMessageW(targetHwnd, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
        PumpPendingMessages();
    }

    return UiMenuRoundTripResult{
        .opened = menuOpened.load(std::memory_order_acquire),
        .closed = menuClosed.load(std::memory_order_acquire),
    };
}

[[nodiscard]] bool WaitForContextMenuKeyboardIndex(HWND popupHwnd, size_t expectedIndex, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
        if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popupHwnd, popupState) && popupState.keyboardIndex.has_value() &&
            popupState.keyboardIndex.value() == expectedIndex)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
    return RedSalamander::DxUi::DebugGetContextMenuPopupState(popupHwnd, popupState) && popupState.keyboardIndex.has_value() &&
           popupState.keyboardIndex.value() == expectedIndex;
}

[[nodiscard]] std::optional<wchar_t> FindFirstTopLevelMenuMnemonic(HMENU mainMenu) noexcept
{
    if (! mainMenu)
    {
        return std::nullopt;
    }

    const int itemCount = GetMenuItemCount(mainMenu);
    for (int itemIndex = 0; itemIndex < itemCount; ++itemIndex)
    {
        const std::wstring itemText = GetMenuItemTextByPosition(mainMenu, itemIndex);
        for (size_t textIndex = 0; textIndex + 1u < itemText.size(); ++textIndex)
        {
            if (itemText[textIndex] != L'&')
            {
                continue;
            }

            const wchar_t next = itemText[textIndex + 1u];
            if (next == L'&')
            {
                ++textIndex;
                continue;
            }

            return static_cast<wchar_t>(std::towupper(next));
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<wchar_t> FindTopLevelMenuMnemonic(HMENU mainMenu, size_t itemIndex) noexcept
{
    if (! mainMenu)
    {
        return std::nullopt;
    }

    const int itemCount = GetMenuItemCount(mainMenu);
    if (itemIndex >= static_cast<size_t>((std::max)(itemCount, 0)))
    {
        return std::nullopt;
    }

    const std::wstring itemText = GetMenuItemTextByPosition(mainMenu, static_cast<int>(itemIndex));
    for (size_t textIndex = 0; textIndex + 1u < itemText.size(); ++textIndex)
    {
        if (itemText[textIndex] != L'&')
        {
            continue;
        }

        const wchar_t next = itemText[textIndex + 1u];
        if (next == L'&')
        {
            ++textIndex;
            continue;
        }

        return static_cast<wchar_t>(std::towupper(next));
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<std::pair<size_t, size_t>> FindFirstTwoEnabledTopLevelMenuIndices(HMENU mainMenu) noexcept
{
    if (! mainMenu)
    {
        return std::nullopt;
    }

    std::optional<size_t> firstIndex;
    const int itemCount = GetMenuItemCount(mainMenu);
    for (int itemIndex = 0; itemIndex < itemCount; ++itemIndex)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(mainMenu, static_cast<UINT>(itemIndex), TRUE, &itemInfo))
        {
            continue;
        }

        if ((itemInfo.fType & MFT_SEPARATOR) != 0 || itemInfo.hSubMenu == nullptr)
        {
            continue;
        }

        if ((itemInfo.fState & MFS_GRAYED) != 0)
        {
            continue;
        }

        if (! firstIndex.has_value())
        {
            firstIndex = static_cast<size_t>(itemIndex);
            continue;
        }

        return std::pair<size_t, size_t>{firstIndex.value(), static_cast<size_t>(itemIndex)};
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<size_t> FindFirstEnabledTopLevelMenuIndex(HMENU mainMenu) noexcept
{
    if (! mainMenu)
    {
        return std::nullopt;
    }

    const int itemCount = GetMenuItemCount(mainMenu);
    for (int itemIndex = 0; itemIndex < itemCount; ++itemIndex)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(mainMenu, static_cast<UINT>(itemIndex), TRUE, &itemInfo))
        {
            continue;
        }

        if ((itemInfo.fType & MFT_SEPARATOR) != 0 || itemInfo.hSubMenu == nullptr || (itemInfo.fState & MFS_GRAYED) != 0)
        {
            continue;
        }

        return static_cast<size_t>(itemIndex);
    }

    return std::nullopt;
}

[[nodiscard]] std::wstring NormalizeMenuItemLabel(std::wstring_view rawText)
{
    std::wstring normalized;
    normalized.reserve(rawText.size());

    const size_t shortcutSeparator = rawText.find(L'\t');
    const std::wstring_view label  = rawText.substr(0, shortcutSeparator);
    for (size_t index = 0; index < label.size(); ++index)
    {
        const wchar_t ch = label[index];
        if (ch != L'&')
        {
            normalized.push_back(ch);
            continue;
        }

        if (index + 1u < label.size() && label[index + 1u] == L'&')
        {
            normalized.push_back(L'&');
            ++index;
        }
    }

    return normalized;
}

[[nodiscard]] std::optional<size_t> FindVisualTopLevelMenuIndexByLabel(std::wstring_view label) noexcept
{
    for (size_t visualIndex = 0;; ++visualIndex)
    {
        std::wstring visualLabel;
        if (! DebugGetMainMenuBarItemLabel(visualIndex, visualLabel))
        {
            break;
        }

        if (visualLabel == label)
        {
            return visualIndex;
        }
    }

    return std::nullopt;
}

[[nodiscard]] size_t CountOwnerDrawMenuItems(HMENU menu) noexcept
{
    if (! menu)
    {
        return 0u;
    }

    size_t ownerDrawCount = 0u;
    const int itemCount   = GetMenuItemCount(menu);
    for (int itemIndex = 0; itemIndex < itemCount; ++itemIndex)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE;
        if (! GetMenuItemInfoW(menu, static_cast<UINT>(itemIndex), TRUE, &itemInfo))
        {
            continue;
        }

        if ((itemInfo.fType & MFT_OWNERDRAW) != 0)
        {
            ++ownerDrawCount;
        }
    }

    return ownerDrawCount;
}

struct TopLevelMenuMapping
{
    size_t rawIndex    = 0;
    size_t visualIndex = 0;
    std::wstring label;
};

[[nodiscard]] std::vector<TopLevelMenuMapping> CollectEnabledTopLevelMenuMappings(HMENU mainMenu) noexcept
{
    std::vector<TopLevelMenuMapping> mappings;
    if (! mainMenu)
    {
        return mappings;
    }

    std::unordered_set<size_t> usedVisualIndices;
    const int itemCount = GetMenuItemCount(mainMenu);
    for (int itemIndex = 0; itemIndex < itemCount; ++itemIndex)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(mainMenu, static_cast<UINT>(itemIndex), TRUE, &itemInfo))
        {
            continue;
        }

        if ((itemInfo.fType & MFT_SEPARATOR) != 0 || itemInfo.hSubMenu == nullptr || (itemInfo.fState & MFS_GRAYED) != 0)
        {
            continue;
        }

        const std::wstring label = NormalizeMenuItemLabel(GetMenuItemTextByPosition(mainMenu, itemIndex));
        if (label.empty())
        {
            continue;
        }

        const std::optional<size_t> visualIndex = FindVisualTopLevelMenuIndexByLabel(label);
        if (! visualIndex.has_value() || ! usedVisualIndices.insert(visualIndex.value()).second)
        {
            continue;
        }

        mappings.push_back(TopLevelMenuMapping{
            .rawIndex    = static_cast<size_t>(itemIndex),
            .visualIndex = visualIndex.value(),
            .label       = label,
        });
    }

    return mappings;
}

[[nodiscard]] std::optional<TopLevelMenuMapping> FindPaneTopLevelMenuMapping(HMENU mainMenu) noexcept
{
    std::optional<TopLevelMenuMapping> fallback;
    for (const TopLevelMenuMapping& mapping : CollectEnabledTopLevelMenuMappings(mainMenu))
    {
        if (mapping.label == L"Right")
        {
            return mapping;
        }

        if (! fallback.has_value() && mapping.label == L"Left")
        {
            fallback = mapping;
        }
    }

    return fallback;
}

[[nodiscard]] std::optional<TopLevelMenuMapping> FindFirstRightJustifiedTopLevelMenuMapping(HMENU mainMenu) noexcept
{
    if (! mainMenu)
    {
        return std::nullopt;
    }

    for (const TopLevelMenuMapping& mapping : CollectEnabledTopLevelMenuMappings(mainMenu))
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(mainMenu, static_cast<UINT>(mapping.rawIndex), TRUE, &itemInfo))
        {
            continue;
        }

        if ((itemInfo.fType & MFT_RIGHTJUSTIFY) != 0)
        {
            return mapping;
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<TopLevelMenuMapping> FindTopLevelMenuMappingByLabel(HMENU mainMenu, std::wstring_view label) noexcept
{
    for (const TopLevelMenuMapping& mapping : CollectEnabledTopLevelMenuMappings(mainMenu))
    {
        if (mapping.label == label)
        {
            return mapping;
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<std::pair<size_t, size_t>> FindFirstCascadeMenuPath(HMENU mainMenu) noexcept
{
    if (! mainMenu)
    {
        return std::nullopt;
    }

    const int topLevelCount = GetMenuItemCount(mainMenu);
    for (int topLevelIndex = 0; topLevelIndex < topLevelCount; ++topLevelIndex)
    {
        const HMENU popupMenu = GetSubMenu(mainMenu, topLevelIndex);
        if (! popupMenu)
        {
            continue;
        }

        const int childCount = GetMenuItemCount(popupMenu);
        for (int childIndex = 0; childIndex < childCount; ++childIndex)
        {
            MENUITEMINFOW itemInfo{};
            itemInfo.cbSize = sizeof(itemInfo);
            itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
            if (! GetMenuItemInfoW(popupMenu, static_cast<UINT>(childIndex), TRUE, &itemInfo))
            {
                continue;
            }

            if ((itemInfo.fType & MFT_SEPARATOR) != 0 || itemInfo.hSubMenu == nullptr || (itemInfo.fState & MFS_GRAYED) != 0)
            {
                continue;
            }

            return std::pair<size_t, size_t>{static_cast<size_t>(topLevelIndex), static_cast<size_t>(childIndex)};
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool WaitForMainMenuBarSelectedIndex(std::optional<size_t> expectedIndex, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const int expectedValue = expectedIndex.has_value() ? static_cast<int>(expectedIndex.value()) : -1;
    const auto deadline     = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (DebugGetMainMenuBarSelectedIndex() == expectedValue)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return DebugGetMainMenuBarSelectedIndex() == expectedValue;
}

[[nodiscard]] bool WaitForMainMenuBarVisualHighlightCount(int expectedCount, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (DebugGetMainMenuBarVisualHighlightCount() == expectedCount)
        {
            return true;
        }
        std::this_thread::sleep_for(10ms);
    }

    return DebugGetMainMenuBarVisualHighlightCount() == expectedCount;
}

[[nodiscard]] bool WaitForMainMenuBarVisualHighlightIndex(std::optional<size_t> expectedIndex, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const int expectedValue = expectedIndex.has_value() ? static_cast<int>(expectedIndex.value()) : -1;
    const auto deadline     = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (DebugGetMainMenuBarVisualHighlightIndex() == expectedValue)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return DebugGetMainMenuBarVisualHighlightIndex() == expectedValue;
}

[[nodiscard]] bool WaitForMainMenuBarRenderCountAtLeast(uint64_t minimumRenderCount, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (DebugGetMainMenuBarRenderCount() >= minimumRenderCount)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return DebugGetMainMenuBarRenderCount() >= minimumRenderCount;
}

[[nodiscard]] HWND FindMainMenuBarWindow(HWND mainWindow) noexcept
{
    return FindWindowExW(mainWindow, nullptr, L"RedSalamander.DxMainMenuBar", nullptr);
}

[[nodiscard]] bool WaitForFocusedFolderViewForMainMenu(HWND expectedFolderView, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.GetFocusedFolderViewHwnd() == expectedFolderView)
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    return g_folderWindow.GetFocusedFolderViewHwnd() == expectedFolderView;
}

[[nodiscard]] HWND WaitForReplacementDxUiContextMenuWindow(HWND previousPopup, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        const HWND currentPopup = FindVisibleDxUiContextMenuWindow();
        if (currentPopup && currentPopup != previousPopup)
        {
            return currentPopup;
        }

        std::this_thread::sleep_for(10ms);
    }

    const HWND currentPopup = FindVisibleDxUiContextMenuWindow();
    return (currentPopup && currentPopup != previousPopup) ? currentPopup : nullptr;
}

[[nodiscard]] bool TestLoadSelectionMenuLinksToRestoreSelection(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    DWORD processId        = 0;
    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, &processId);
    state.Require(uiThreadId != 0u, L"Failed to resolve the UI thread for app drive-menu shell-stability validation.");
    if (uiThreadId == 0u)
    {
        return false;
    }

    state.Require(FindCommandInfo(L"cmd/pane/loadSelection") == nullptr, L"cmd/pane/loadSelection should not be registered.");
    state.Require(FindCommandInfo(L"cmd/pane/selection/restore") != nullptr, L"cmd/pane/selection/restore should be registered.");

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const HMENU advancedMenu = FindMenuContainingCommandId(mainMenu, IDM_PANE_SAVE_SELECTION);
    state.Require(advancedMenu != nullptr, L"Failed to find Edit > Advanced menu.");
    if (! advancedMenu)
    {
        return false;
    }

    state.Require(MenuContainsCommandId(advancedMenu, IDM_PANE_SELECTION_RESTORE), L"Edit > Advanced menu should contain selection restore command.");
    state.Require(! MenuContainsCommandId(advancedMenu, IDM_PANE_LOAD_SELECTION), L"Edit > Advanced menu should not contain IDM_PANE_LOAD_SELECTION.");

    return state.failure.empty();
}

[[nodiscard]] bool TestCopyTextCommandsMenuContract(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const HMENU editMenu = FindMenuContainingCommandId(mainMenu, IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT);
    state.Require(editMenu != nullptr, L"Failed to find Edit menu for copy-text command contract.");
    if (! editMenu)
    {
        return false;
    }

    const int firstCopyTextPos = FindMenuItemPosById(editMenu, IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT);
    state.Require(firstCopyTextPos >= 0, L"Copy Path + Name as Text menu entry not found.");
    if (firstCopyTextPos < 0)
    {
        return false;
    }

    state.Require(firstCopyTextPos > 0 && IsMenuSeparatorAt(editMenu, firstCopyTextPos - 1), L"Copy-text menu group should be preceded by a separator.");

    using MenuContractExpectation                                   = std::pair<UINT, std::wstring_view>;
    constexpr std::array<MenuContractExpectation, 4> kExpectedItems = {
        MenuContractExpectation{IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT, std::wstring_view{L"Copy Path + Name as Text"}},
        MenuContractExpectation{IDM_PANE_COPY_NAME_AS_TEXT, std::wstring_view{L"Copy Name as Text"}},
        MenuContractExpectation{IDM_PANE_COPY_PATH_AS_TEXT, std::wstring_view{L"Copy Path as Text"}},
        MenuContractExpectation{IDM_PANE_COPY_PATH_AND_FILE_NAME, std::wstring_view{L"Copy UNC Path + Name as Text"}},
    };

    const int itemCount = GetMenuItemCount(editMenu);
    state.Require(itemCount >= (firstCopyTextPos + static_cast<int>(kExpectedItems.size())), L"Copy-text menu group truncated.");
    if (itemCount < (firstCopyTextPos + static_cast<int>(kExpectedItems.size())))
    {
        return false;
    }

    for (int index = 0; index < static_cast<int>(kExpectedItems.size()); ++index)
    {
        const auto& [expectedId, expectedText] = kExpectedItems[static_cast<size_t>(index)];
        const int pos                          = firstCopyTextPos + index;
        const UINT actualId                    = GetMenuItemID(editMenu, pos);
        const std::wstring actualText          = GetMenuItemTextByPosition(editMenu, pos);

        state.Require(actualId == expectedId, std::format(L"Copy-text menu item {} expected command id {}.", index, expectedId));
        state.Require(actualText == expectedText, std::format(L"Copy-text menu item {} expected label '{}'.", index, expectedText));
        state.Require(DebugGetMainMenuIconGlyph(expectedId) == 0, std::format(L"{} should remain a text-only menu entry.", expectedText));
    }

    state.Require((firstCopyTextPos + static_cast<int>(kExpectedItems.size())) < itemCount &&
                      IsMenuSeparatorAt(editMenu, firstCopyTextPos + static_cast<int>(kExpectedItems.size())),
                  L"Copy-text menu group should be followed by a separator.");

    return state.failure.empty();
}

[[nodiscard]] bool TestDispatchAllCommandsSmoke(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const DWORD processId             = GetCurrentProcessId();
    const DWORD uiThreadId            = GetWindowThreadProcessId(mainWindow, nullptr);
    const auto baseline               = SnapshotTopLevelWindowsForProcess(processId);
    const bool baselineMenuBarVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);

    const auto restoreSmokeUiState = [&]() noexcept
    {
        if (g_folderWindow.DebugIsViewWidthAdjustActive())
        {
            static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_ESCAPE));
        }

        const LONG_PTR style = GetWindowLongPtrW(mainWindow, GWL_STYLE);
        if ((style & WS_POPUP) != 0 && (style & WS_CAPTION) == 0)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_FULL_SCREEN, 0), 0);
        }

        if (DebugIsMainMenuBarSurfaceVisible(mainWindow) != baselineMenuBarVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, baselineMenuBarVisible, SelfTest::Scale(std::chrono::milliseconds{1000})));
        }

        PumpPendingMessages();
    };

    std::jthread autoCloser;
    try
    {
        autoCloser = std::jthread([&](std::stop_token stopToken) noexcept { AutoCloseTransientUi(stopToken, uiThreadId, processId, baseline, mainWindow); });
    }
    catch (const std::system_error&)
    {
        state.Require(false, L"Dispatch smoke: failed to start UI auto-closer thread.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / L"dispatch_smoke";
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create dispatch_smoke left folder.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create dispatch_smoke right folder.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Dispatch smoke: failed to set left pane path.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Dispatch smoke: failed to set right pane path.");

    const std::unordered_set<std::wstring_view> skipIds = {
        L"cmd/app/exit",
        L"cmd/app/externalHelp",
        L"cmd/app/openFileExplorerKnownFolder",
        L"cmd/pane/openCommandShell",
        L"cmd/pane/openCurrentFolder",
        // Avoid starting real file operations in the command-dispatch smoke test (covered by FileOperations suite).
        L"cmd/pane/copyToOtherPane",
        L"cmd/pane/moveToOtherPane",
        L"cmd/pane/moveToRecycleBin",
        L"cmd/pane/delete",
        L"cmd/pane/permanentDelete",
        L"cmd/pane/rename",
        L"cmd/pane/createDirectory",
    };

    const auto commands = GetAllCommands();
    state.Require(! commands.empty(), L"Dispatch smoke: GetAllCommands returned empty.");
    if (commands.empty())
    {
        return false;
    }

    for (const CommandInfo& cmd : commands)
    {
        if (skipIds.contains(cmd.id))
        {
            continue;
        }

        PumpPendingMessages();

        Trace(std::format(L"Dispatch: {}", cmd.id));

        if (cmd.wmCommandId != 0)
        {
            const WPARAM wp = MAKEWPARAM(static_cast<WORD>(cmd.wmCommandId), 0);
            SendMessageW(mainWindow, WM_COMMAND, wp, 0);
        }
        else
        {
            static_cast<void>(DebugDispatchShortcutCommand(mainWindow, cmd.id));
        }

        std::this_thread::sleep_for(10ms);

        state.Require(EnsureUiNotInMenuMode(uiThreadId, mainWindow, SelfTest::Scale(std::chrono::milliseconds{500})),
                      std::format(L"Dispatch smoke: {} left UI in menu mode.", cmd.id));
        state.Require(WaitForNoNonBaselineWindows(processId, baseline, mainWindow, SelfTest::Scale(std::chrono::milliseconds{500})),
                      std::format(L"Dispatch smoke: {} left windows open.", cmd.id));
        restoreSmokeUiState();
        if (! state.failure.empty())
        {
            return false;
        }
    }

    restoreSmokeUiState();
    state.Require(EnsureUiNotInMenuMode(uiThreadId, mainWindow, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Dispatch smoke: cleanup left UI in menu mode.");
    state.Require(WaitForNoNonBaselineWindows(processId, baseline, mainWindow, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Dispatch smoke: cleanup left windows open.");
    return state.failure.empty();
}

[[nodiscard]] bool TestModelessWindowOwnership(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePaths                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = GetPreferencesDialogHandle();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open.");
    if (prefs)
    {
        state.Require(! IsOwnedBy(prefs, mainWindow), L"Preferences window should be an independent top-level window.");
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Preferences window did not close.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CONNECTION_MANAGER, 0), 0);
    const HWND connMgr = GetConnectionManagerDialogHandle();
    state.Require(connMgr != nullptr && IsWindow(connMgr) != FALSE, L"Connection Manager window did not open.");
    if (connMgr)
    {
        state.Require(! IsOwnedBy(connMgr, mainWindow), L"Connection Manager window should be an independent top-level window.");
        PostMessageW(connMgr, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(connMgr, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Connection Manager window did not close.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);
    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open.");
    if (shortcuts)
    {
        state.Require(! IsOwnedBy(shortcuts, mainWindow), L"Shortcuts window should be an independent top-level window.");
        PostMessageW(shortcuts, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(shortcuts, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Shortcuts window did not close.");
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");

    const std::filesystem::path compareRoot = suiteRoot / L"work" / L"compare_modeless";
    const std::filesystem::path leftFolder  = compareRoot / L"left";
    const std::filesystem::path rightFolder = compareRoot / L"right";

    std::error_code ec;
    std::filesystem::remove_all(compareRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftFolder), L"Failed to create compare_modeless left folder.");
    state.Require(SelfTest::EnsureDirectory(rightFolder), L"Failed to create compare_modeless right folder.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftFolder);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightFolder);

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);
    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"Compare window did not open.");
    if (compare)
    {
        state.Require(! IsOwnedBy(compare, mainWindow), L"Compare window should be an independent top-level window.");
        PostMessageW(compare, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(compare, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Compare window did not close.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_FIND, 0), 0);
    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open.");
    if (findWindow)
    {
        state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window.");
        const AppTheme darkTheme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-dark");
        UpdateFindFilesWindowsTheme(darkTheme);
        PumpPendingMessages();
        state.Require(IsWindow(findWindow) != FALSE, L"Find window did not survive runtime theme update to dark mode.");

        const AppTheme lightTheme = ResolveAppTheme(ThemeMode::Light, L"find-selftest-light");
        UpdateFindFilesWindowsTheme(lightTheme);
        PumpPendingMessages();
        state.Require(IsWindow(findWindow) != FALSE, L"Find window did not survive runtime theme update to light mode.");

        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(std::chrono::milliseconds{2000})), L"Find window did not close.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFullScreenToggle(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const LONG_PTR styleBefore = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exBefore    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_FULL_SCREEN, 0), 0);

    const LONG_PTR styleFull = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exFull    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);

    state.Require((styleFull & WS_POPUP) != 0, L"Fullscreen expected WS_POPUP.");
    state.Require((styleFull & WS_CAPTION) == 0, L"Fullscreen expected no WS_CAPTION.");
    state.Require((exFull & WS_EX_TOPMOST) != 0, L"Fullscreen expected WS_EX_TOPMOST.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_FULL_SCREEN, 0), 0);

    const LONG_PTR styleAfter = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exAfter    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);

    state.Require(styleAfter == styleBefore, L"Fullscreen toggle did not restore original style.");
    state.Require(exAfter == exBefore, L"Fullscreen toggle did not restore original ex-style.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFullScreenKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    DWORD processId        = 0;
    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, &processId);
    state.Require(uiThreadId != 0u, L"Failed to resolve the UI thread for app drive-menu shell-stability validation.");
    if (uiThreadId == 0u)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"fullscreen_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for fullscreen shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for fullscreen shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for fullscreen shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for fullscreen shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const LONG_PTR styleBefore                             = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exBefore                                = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);
    const auto restoreState                                = wil::scope_exit([&]
    {
        const LONG_PTR styleNow = GetWindowLongPtrW(mainWindow, GWL_STYLE);
        if ((styleNow & WS_POPUP) != 0)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_FULL_SCREEN, 0), 0);
        }

        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }

        const LONG_PTR finalStyle = GetWindowLongPtrW(mainWindow, GWL_STYLE);
        const LONG_PTR finalEx    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);
        if (finalStyle != styleBefore || finalEx != exBefore)
        {
            SetWindowLongPtrW(mainWindow, GWL_STYLE, styleBefore);
            SetWindowLongPtrW(mainWindow, GWL_EXSTYLE, exBefore);
            SetWindowPos(mainWindow, nullptr, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during fullscreen shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during fullscreen shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for fullscreen shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for fullscreen shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for fullscreen shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for fullscreen shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before fullscreen shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const size_t baselineLeftItemCount  = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before fullscreen shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before fullscreen shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  std::filesystem::path expectedPath,
                                                  size_t expectedItemCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == 0u &&
                   ! g_folderWindow.DebugIsNameFilterActive(pane);
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}.",
                                  context,
                                  static_cast<unsigned>(pane),
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(pane),
                                  g_folderWindow.DebugGetSelectedCount(pane),
                                  g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no"));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_FULL_SCREEN, 0), 0);

    const LONG_PTR styleFull = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exFull    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);
    state.Require((styleFull & WS_POPUP) != 0, L"Fullscreen shell-stability validation expected WS_POPUP.");
    state.Require((styleFull & WS_CAPTION) == 0, L"Fullscreen shell-stability validation expected no WS_CAPTION.");
    state.Require((exFull & WS_EX_TOPMOST) != 0, L"Fullscreen shell-stability validation expected WS_EX_TOPMOST.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Full Screen on");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Full Screen on");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_FULL_SCREEN, 0), 0);

    const LONG_PTR styleAfter = GetWindowLongPtrW(mainWindow, GWL_STYLE);
    const LONG_PTR exAfter    = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);
    state.Require(styleAfter == styleBefore, L"Fullscreen shell-stability validation did not restore original style.");
    state.Require(exAfter == exBefore, L"Fullscreen shell-stability validation did not restore original ex-style.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Full Screen off");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Full Screen off");

    return state.failure.empty();
}

[[nodiscard]] bool TestDriveMenuCommands(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    DWORD processId        = 0;
    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, &processId);
    state.Require(uiThreadId != 0u, L"Failed to resolve the UI thread for drive-menu DxUI validation.");
    if (uiThreadId == 0u)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for drive-menu DxUI validation.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"drive_menu_dxui_popup_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for drive-menu DxUI validation.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for drive-menu DxUI validation.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for drive-menu DxUI validation.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for drive-menu DxUI validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<FolderWindow::Pane> zoomBefore     = g_folderWindow.GetZoomedPane();
    const std::optional<float> zoomRestoreBefore           = g_folderWindow.GetZoomRestoreSplitRatio();
    const auto restoreZoom                                 = wil::scope_exit([&] { g_folderWindow.SetZoomState(zoomBefore, zoomRestoreBefore); });
    const auto restorePanes                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.SetZoomState(std::nullopt, std::nullopt);
    PumpPendingMessages();
    state.Require(! g_folderWindow.GetZoomedPane().has_value(),
                  L"Unfocused-pane navigation click validation requires both panes visible and could not clear the prior zoom state.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set the local file-system plugin for the left pane during drive-menu DxUI validation.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set the local file-system plugin for the right pane during drive-menu DxUI validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)), L"Failed to set left pane path for drive-menu DxUI validation.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)), L"Failed to set right pane path for drive-menu DxUI validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for drive-menu DxUI validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for drive-menu DxUI validation.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before drive-menu DxUI validation.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Right, L"right.txt"),
                  L"Failed to focus right.txt before drive-menu DxUI validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND leftFolderView      = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND rightFolderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right);
    const HWND leftNavigationView  = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    const HWND rightNavigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Right);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE, L"Left folder view handle unavailable for drive-menu DxUI validation.");
    state.Require(rightFolderView != nullptr && IsWindow(rightFolderView) != FALSE, L"Right folder view handle unavailable for drive-menu DxUI validation.");
    state.Require(leftNavigationView != nullptr && IsWindow(leftNavigationView) != FALSE,
                  L"Left navigation-view handle unavailable for drive-menu DxUI validation.");
    state.Require(rightNavigationView != nullptr && IsWindow(rightNavigationView) != FALSE,
                  L"Right navigation-view handle unavailable for drive-menu DxUI validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto captureBaseline =
        [&](const FolderWindow::Pane pane, const std::filesystem::path& expectedPath, NavigationViewDebugSnapshot& outSnapshot) noexcept
    {
        return WaitForNavigationViewSnapshot(pane,
                                             [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.showMenuSection && value.currentPathText == expectedPath.wstring();
        },
                                             SelfTest::Scale(3000ms),
                                             &outSnapshot);
    };

    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(captureBaseline(FolderWindow::Pane::Left, left, baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before drive-menu DxUI validation.");
    state.Require(captureBaseline(FolderWindow::Pane::Right, right, baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before drive-menu DxUI validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto openAndDismissDxDriveMenu = [&](const FolderWindow::Pane pane,
                                               const UINT commandId,
                                               const HWND folderView,
                                               const HWND navigationView,
                                               const std::filesystem::path& expectedPath,
                                               const NavigationViewDebugSnapshot& baselineSnapshot,
                                               std::wstring_view expectedFocusedItem,
                                               std::wstring_view label) noexcept
    {
        const size_t baselineItemCount     = g_folderWindow.DebugGetItemCount(pane);
        const size_t baselineSelectedCount = g_folderWindow.DebugGetSelectedCount(pane);

        FocusFolderViewPane(pane);

        const UiMenuRoundTripResult roundTrip =
            TriggerAndDismissCommandActivatedUiMenu(mainWindow, commandId, uiThreadId, navigationView, SelfTest::Scale(3000ms), SelfTest::Scale(3000ms));
        PumpPendingMessages();

        state.Require(roundTrip.opened, std::format(L"{} did not open a live menu session.", label));
        state.Require(roundTrip.closed, std::format(L"{} popup did not dismiss after Escape.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        NavigationViewDebugSnapshot closedSnapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == baselineItemCount && g_folderWindow.DebugGetSelectedCount(pane) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &closedSnapshot),
                      std::format(L"{} did not settle back to the baseline navigation shell after Escape.", label));
        state.Require(GetFocus() == folderView || GetFocus() == navigationView,
                      std::format(L"{} should return focus to the pane shell or folder view.", label));
        state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(pane) == expectedFocusedItem,
                      std::format(L"{} should preserve the focused item on the pane.", label));
        return state.failure.empty();
    };

    state.Require(openAndDismissDxDriveMenu(FolderWindow::Pane::Left,
                                            IDM_LEFT_CHANGE_DRIVE,
                                            leftFolderView,
                                            leftNavigationView,
                                            left,
                                            baselineLeftSnapshot,
                                            L"left.txt",
                                            L"Open Left Drive Menu"),
                  L"Left drive-menu command did not use the stable DxUI popup path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(openAndDismissDxDriveMenu(FolderWindow::Pane::Right,
                                            IDM_RIGHT_CHANGE_DRIVE,
                                            rightFolderView,
                                            rightNavigationView,
                                            right,
                                            baselineRightSnapshot,
                                            L"right.txt",
                                            L"Open Right Drive Menu"),
                  L"Right drive-menu command did not use the stable DxUI popup path.");
    if (! state.failure.empty())
    {
        return false;
    }

    return state.failure.empty();
}

class ScopedNavigationBarVisibilityRestore final
{
public:
    ScopedNavigationBarVisibilityRestore() noexcept
        : _leftVisible(g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left)),
          _rightVisible(g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Right))
    {
    }

    ScopedNavigationBarVisibilityRestore(const ScopedNavigationBarVisibilityRestore&)            = delete;
    ScopedNavigationBarVisibilityRestore& operator=(const ScopedNavigationBarVisibilityRestore&) = delete;

    ~ScopedNavigationBarVisibilityRestore() noexcept
    {
        g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Left, _leftVisible);
        g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Right, _rightVisible);
        PumpPendingMessages();
    }

private:
    bool _leftVisible{};
    bool _rightVisible{};
};

[[nodiscard]] bool EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane pane, HWND navigationView, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    g_folderWindow.SetNavigationBarVisible(pane, true);
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.GetNavigationBarVisible(pane) && navigationView != nullptr && IsWindow(navigationView) != FALSE &&
            IsWindowVisible(navigationView) != FALSE)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    return g_folderWindow.GetNavigationBarVisible(pane) && navigationView != nullptr && IsWindow(navigationView) != FALSE &&
           IsWindowVisible(navigationView) != FALSE;
}

[[nodiscard]] bool TestPaneFocusAddressBarTabTraversal(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view traversal test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"navigation_view_tab_traversal_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create navigation-view traversal test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for navigation-view traversal test.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for navigation-view traversal test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view traversal test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for navigation-view traversal test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for navigation-view traversal test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for navigation-view traversal test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for navigation-view traversal test.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE, L"Navigation view handle unavailable for navigation-view traversal test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const bool originalLeftNavigationBarVisible  = g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left);
    const bool originalRightNavigationBarVisible = g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Right);
    const auto restoreNavigationBars             = wil::scope_exit([&]() noexcept
    {
        g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Left, originalLeftNavigationBarVisible);
        g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Right, originalRightNavigationBarVisible);
    });

    const std::wstring expectedFocusedItem = L"alpha.txt";
    const uint64_t baselineRefreshCount    = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount         = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount     = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    const auto waitForFolderFocus = [&](std::wstring_view /*context*/) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        return g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    };

    const auto focusAddressBarAndWait = [&](std::wstring_view context, NavigationViewDebugSnapshot& snapshot) noexcept
    {
        if (const HWND rootWindow = GetAncestor(mainWindow, GA_ROOT); rootWindow && GetActiveWindow() != rootWindow)
        {
            SetActiveWindow(rootWindow);
        }
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/focusAddressBar"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/focusAddressBar during {}.", context));
        const bool editModeReady = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                                 [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left) &&
                   value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && ! value.currentEditText.empty() &&
                   value.visibleChildWindowCount == 1u && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode &&
                   value.currentPathText == root.wstring();
        },
                                                                 SelfTest::Scale(3000ms),
                                                                 &snapshot);
        state.Require(editModeReady,
                      std::format(L"Navigation view did not enter address-bar edit mode during {}; focusTarget={}, editMode={}, editText='{}', "
                                  L"pathText='{}', visibleChildren={}, navigationBarVisible={}, fullPathPopupVisible={}, fullPathPopupEditMode={}, "
                                  L"focusedPane={}, focusedView={:#x}, leftView={:#x}.",
                                  context,
                                  static_cast<int>(snapshot.focusTarget),
                                  snapshot.editMode ? 1 : 0,
                                  snapshot.currentEditText,
                                  snapshot.currentPathText,
                                  snapshot.visibleChildWindowCount,
                                  g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left) ? 1 : 0,
                                  snapshot.fullPathPopupVisible ? 1 : 0,
                                  snapshot.fullPathPopupEditMode ? 1 : 0,
                                  static_cast<int>(g_folderWindow.GetFocusedPane()),
                                  reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd()),
                                  reinterpret_cast<uintptr_t>(folderView)));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto findAddressBarEditBridge = [&](HWND parentHwnd) noexcept
        {
            const auto isEditLikeClass = [](HWND hwnd) noexcept
            {
                if (! hwnd || IsWindow(hwnd) == FALSE)
                {
                    return false;
                }

                std::array<wchar_t, 128> className{};
                const int classLength = GetClassNameW(hwnd, className.data(), static_cast<int>(className.size()));
                if (classLength <= 0)
                {
                    return false;
                }

                const std::wstring_view actualClassName(className.data(), static_cast<size_t>(classLength));
                return actualClassName == L"Edit" || (actualClassName.size() >= 8u && _wcsnicmp(actualClassName.data(), L"RICHEDIT", 8) == 0);
            };

            if (snapshot.currentEditBridgeHwnd && IsWindow(snapshot.currentEditBridgeHwnd) != FALSE)
            {
                return snapshot.currentEditBridgeHwnd;
            }
            if (snapshot.currentEditHostHwnd && IsWindow(snapshot.currentEditHostHwnd) != FALSE && isEditLikeClass(snapshot.currentEditHostHwnd))
            {
                return snapshot.currentEditHostHwnd;
            }
            if (const HWND focused = GetFocus(); focused && (focused == parentHwnd || IsChild(parentHwnd, focused) != FALSE) && isEditLikeClass(focused))
            {
                return focused;
            }

            struct DescendantSearchContext
            {
                HWND found = nullptr;
            } searchContext{};

            static_cast<void>(EnumChildWindows(parentHwnd,
                                               [](HWND child, LPARAM lParam) noexcept -> BOOL
            {
                auto& contextRef = *reinterpret_cast<DescendantSearchContext*>(lParam);
                std::array<wchar_t, 128> className{};
                const int classLength = GetClassNameW(child, className.data(), static_cast<int>(className.size()));
                if (classLength <= 0)
                {
                    return TRUE;
                }

                const std::wstring_view actualClassName(className.data(), static_cast<size_t>(classLength));
                if (actualClassName != L"Edit" && (actualClassName.size() < 8u || _wcsnicmp(actualClassName.data(), L"RICHEDIT", 8) != 0))
                {
                    return TRUE;
                }

                contextRef.found = child;
                return FALSE;
            },
                                               reinterpret_cast<LPARAM>(&searchContext)));

            return searchContext.found;
        };

        const auto readEditWindowText = [](HWND editHwnd) noexcept -> std::wstring
        {
            const int textLength = GetWindowTextLengthW(editHwnd);
            if (textLength < 0)
            {
                return {};
            }

            std::wstring windowText(static_cast<size_t>(textLength) + 1u, L'\0');
            if (textLength > 0)
            {
                GetWindowTextW(editHwnd, windowText.data(), textLength + 1);
            }
            windowText.resize(wcsnlen(windowText.c_str(), static_cast<size_t>(textLength)));
            return windowText;
        };

        std::optional<UiaValuePatternState> valueState;
        const auto uiaDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < uiaDeadline)
        {
            PumpPendingMessages();
            const HWND editHwnd = findAddressBarEditBridge(navigationView);
            if (editHwnd && IsWindow(editHwnd) != FALSE)
            {
                wil::com_ptr<IUIAutomationElement> editElement;
                if (TryGetUiAutomationRootElement(editHwnd, editElement) && editElement)
                {
                    valueState.emplace();
                    static_cast<void>(editElement->get_CurrentControlType(&valueState->controlType));

                    wil::unique_bstr name;
                    if (SUCCEEDED(editElement->get_CurrentName(&name)))
                    {
                        valueState->name.assign(name.get() ? name.get() : L"");
                    }

                    wil::com_ptr<IValueProvider> valueProvider;
                    if (SUCCEEDED(editElement->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IValueProvider), valueProvider.put_void())) && valueProvider)
                    {
                        wil::unique_bstr value;
                        if (SUCCEEDED(valueProvider->get_Value(&value)))
                        {
                            valueState->value.assign(value.get() ? value.get() : L"");
                        }

                        BOOL readOnly = FALSE;
                        if (SUCCEEDED(valueProvider->get_IsReadOnly(&readOnly)))
                        {
                            valueState->isReadOnly = readOnly != FALSE;
                        }
                    }

                    if (valueState->value.empty())
                    {
                        VARIANT propertyValue{};
                        if (SUCCEEDED(editElement->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &propertyValue)))
                        {
                            const auto clearPropertyValue = wil::scope_exit([&propertyValue]() noexcept { VariantClear(&propertyValue); });
                            if (propertyValue.vt == VT_BSTR && propertyValue.bstrVal)
                            {
                                valueState->value.assign(propertyValue.bstrVal);
                            }
                        }
                    }

                    if (valueState->controlType == UIA_EditControlTypeId && ! valueState->value.empty() &&
                        valueState->value.find(root.wstring()) != std::wstring::npos)
                    {
                        break;
                    }

                    if (valueState->value.empty())
                    {
                        valueState->value = readEditWindowText(editHwnd);
                    }

                    if (valueState->controlType == UIA_EditControlTypeId && ! valueState->value.empty() &&
                        valueState->value.find(root.wstring()) != std::wstring::npos)
                    {
                        break;
                    }
                }

                if (! valueState.has_value())
                {
                    valueState.emplace();
                    valueState->controlType = UIA_EditControlTypeId;
                    valueState->value       = readEditWindowText(editHwnd);
                    valueState->isReadOnly  = (GetWindowLongPtrW(editHwnd, GWL_STYLE) & ES_READONLY) != 0;
                }

                if (valueState->controlType == UIA_EditControlTypeId && ! valueState->value.empty() &&
                    valueState->value.find(root.wstring()) != std::wstring::npos)
                {
                    break;
                }
            }
            std::this_thread::sleep_for(20ms);
        }

        if (! valueState.has_value() && ! snapshot.currentEditText.empty())
        {
            valueState.emplace();
            valueState->controlType = UIA_EditControlTypeId;
            valueState->value       = snapshot.currentEditText;
            valueState->isReadOnly  = false;
        }

        state.Require(valueState.has_value(), std::format(L"Failed to collect the navigation-view address-bar edit UIA value state during {}.", context));
        if (valueState.has_value())
        {
            state.Require(valueState->controlType == UIA_EditControlTypeId || ! valueState->value.empty(),
                          std::format(L"Navigation view address-bar child should expose an Edit control or readable value during {}; controlType={} name='{}'.",
                                      context,
                                      valueState->controlType,
                                      valueState->name));
            state.Require(! valueState->value.empty(), std::format(L"Navigation view edit text should not be empty during {}.", context));
            state.Require(valueState->value.find(root.wstring()) != std::wstring::npos,
                          std::format(L"Navigation view edit text should include the current folder during {}; saw '{}'.", context, valueState->value));
        }

        return state.failure.empty();
    };

    NavigationViewDebugSnapshot snapshot{};
    const auto sendTabFromFocusedWindow = [&](const bool reverse, std::wstring_view context) noexcept
    {
        const auto describeFocusState = [&](const NavigationViewDebugSnapshot& snapshot) noexcept
        {
            return std::format(L"focus={}, active=0x{:X}, host={}, bridge={}, focusTarget={}, editMode={}, editText='{}', "
                               L"visibleChildren={}, folderFocus=0x{:X}, expectedFolder=0x{:X}.",
                               DescribeWindowHandleForSelfTest(GetFocus()),
                               reinterpret_cast<uintptr_t>(GetActiveWindow()),
                               DescribeWindowHandleForSelfTest(snapshot.currentEditHostHwnd),
                               DescribeWindowHandleForSelfTest(snapshot.currentEditBridgeHwnd),
                               static_cast<int>(snapshot.focusTarget),
                               snapshot.editMode ? 1 : 0,
                               snapshot.currentEditText,
                               snapshot.visibleChildWindowCount,
                               reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd()),
                               reinterpret_cast<uintptr_t>(folderView));
        };

        const auto focusBelongsToEditSurface = [](HWND focus, const NavigationViewDebugSnapshot& snapshot) noexcept
        {
            return focus && IsWindow(focus) != FALSE &&
                   ((snapshot.currentEditHostHwnd && focus == snapshot.currentEditHostHwnd) ||
                    (snapshot.currentEditBridgeHwnd && focus == snapshot.currentEditBridgeHwnd) ||
                    (snapshot.currentEditHostHwnd && IsChild(snapshot.currentEditHostHwnd, focus) != FALSE));
        };

        NavigationViewDebugSnapshot tabSnapshot = snapshot;
        HWND focused              = nullptr;
        const auto focusDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
        while (std::chrono::steady_clock::now() < focusDeadline)
        {
            focused = GetFocus();
            if (focusBelongsToEditSurface(focused, tabSnapshot))
            {
                break;
            }

            const bool editSurfaceReady = tabSnapshot.editMode && tabSnapshot.focusTarget == NavigationViewDebugFocusTarget::PathEdit &&
                                          tabSnapshot.visibleChildWindowCount == 1u && ! tabSnapshot.currentEditText.empty() &&
                                          tabSnapshot.currentEditText.find(root.wstring()) != std::wstring::npos;
            if (editSurfaceReady)
            {
                const HWND preferredTarget =
                    (tabSnapshot.currentEditBridgeHwnd && IsWindow(tabSnapshot.currentEditBridgeHwnd) != FALSE) ? tabSnapshot.currentEditBridgeHwnd
                                                                                                               : tabSnapshot.currentEditHostHwnd;
                if (preferredTarget && IsWindow(preferredTarget) != FALSE)
                {
                    if (const HWND rootWindow = GetAncestor(mainWindow, GA_ROOT); rootWindow && GetActiveWindow() != rootWindow)
                    {
                        SetActiveWindow(rootWindow);
                    }
                    SetFocus(preferredTarget);
                    focused = GetFocus();
                    if (focusBelongsToEditSurface(focused, tabSnapshot))
                    {
                        break;
                    }
                }
            }

            PumpPendingMessages();

            NavigationViewDebugSnapshot currentSnapshot{};
            if (g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, currentSnapshot))
            {
                tabSnapshot = currentSnapshot;
            }

            std::this_thread::sleep_for(10ms);
        }

        state.Require(focusBelongsToEditSurface(focused, tabSnapshot),
                      std::format(L"Focused edit window unavailable before {}; {}.", context, describeFocusState(tabSnapshot)));
        if (! state.failure.empty())
        {
            return false;
        }

        BYTE keyboardState[256]{};
        bool keyboardStateCaptured      = false;
        const auto restoreKeyboardState = wil::scope_exit([&]() noexcept
        {
            if (keyboardStateCaptured)
            {
                SetKeyboardState(keyboardState);
            }
        });
        if (reverse)
        {
            keyboardStateCaptured = GetKeyboardState(keyboardState) != FALSE;
            if (keyboardStateCaptured)
            {
                BYTE shiftedState[256]{};
                std::copy(std::begin(keyboardState), std::end(keyboardState), std::begin(shiftedState));
                shiftedState[VK_SHIFT] = 0x80;
                SetKeyboardState(shiftedState);
            }
        }

        if (reverse)
        {
            SendMessageW(focused, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(focused, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(focused, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(focused, WM_KEYUP, VK_SHIFT, 0);
        }

        state.Require(waitForFolderFocus(context), std::format(L"Pane folder view did not regain focus with stable selection after {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        NavigationViewDebugSnapshot postSnapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [](const NavigationViewDebugSnapshot& value) noexcept
        {
            return ! value.editMode && value.focusTarget == NavigationViewDebugFocusTarget::None && value.visibleChildWindowCount == 0u &&
                   ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &postSnapshot),
                      std::format(L"Navigation view should tear down the address-bar edit surface after {}.", context));
        return state.failure.empty();
    };

    state.Require(waitForFolderFocus(L"initial address-bar focus setup"),
                  L"Pane folder view did not hold stable left focus before opening the address-bar edit surface.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Left, false);
    PumpPendingMessages();
    state.Require(! g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left),
                  L"Navigation-view traversal test could not hide the left navigation bar before command reveal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(focusAddressBarAndWait(L"forward tab handoff", snapshot),
                  L"Navigation view did not expose the live address-bar edit surface before forward tab handoff.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(sendTabFromFocusedWindow(false, L"forward address-bar tab handoff"),
                  L"Forward tab from the address bar should return focus to the left folder view.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(focusAddressBarAndWait(L"reverse tab handoff", snapshot),
                  L"Navigation view did not expose the live address-bar edit surface before reverse tab handoff.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(sendTabFromFocusedWindow(true, L"reverse address-bar shift-tab handoff"),
                  L"Reverse Shift+Tab from the address bar should return focus to the left folder view.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewPathDoubleClickEntersEditMode(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view path double-click test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / (L"nav_path_dblclk_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create navigation-view path double-click root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for navigation-view path double-click test.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for navigation-view path double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view path double-click test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for navigation-view path double-click test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for navigation-view path double-click test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for navigation-view path double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for navigation-view path double-click test.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for navigation-view path double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before navigation-view path double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring expectedFocusedItem = L"alpha.txt";
    const uint64_t baselineRefreshCount    = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount         = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount     = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.pathRegionRect.right > value.pathRegionRect.left &&
               value.pathRegionRect.bottom > value.pathRegionRect.top && value.pathLastSegmentVisible &&
               value.pathLastSegmentRect.right > value.pathLastSegmentRect.left &&
               value.pathLastSegmentRect.bottom > value.pathLastSegmentRect.top;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture baseline navigation-view snapshot for path double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG pathWidth = baselineSnapshot.pathRegionRect.right - baselineSnapshot.pathRegionRect.left;
    state.Require(pathWidth > 40, std::format(L"Navigation-view path region should be wide enough for the path double-click test; saw {} px.", pathWidth));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto makePathClickPoint = [](const NavigationViewDebugSnapshot& snapshot) noexcept
    {
        const RECT& target = snapshot.pathLastSegmentVisible ? snapshot.pathLastSegmentRect : snapshot.pathRegionRect;
        POINT point{};
        point.x = (target.left + target.right) / 2;
        point.y = (target.top + target.bottom) / 2;
        return point;
    };

    const POINT singleClickPoint   = makePathClickPoint(baselineSnapshot);
    const LPARAM singleClickLParam = MAKELPARAM(singleClickPoint.x, singleClickPoint.y);
    SendMouseClickToResolvedPointWindow(navigationView, singleClickLParam);

    NavigationViewDebugSnapshot afterSingleSnapshot{};
    const bool singleClickStable = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                                 [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.pathRegionRect.right > value.pathRegionRect.left &&
               value.pathRegionRect.bottom > value.pathRegionRect.top &&
               value.pathLastSegmentVisible && value.pathLastSegmentRect.right > value.pathLastSegmentRect.left &&
               value.pathLastSegmentRect.bottom > value.pathLastSegmentRect.top &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                                 SelfTest::Scale(3000ms),
                                                                 &afterSingleSnapshot);
    const std::optional<std::filesystem::path> panePathAfterSingle = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    state.Require(singleClickStable,
                  std::format(L"Single-clicking the navigation-view path region should keep the shell stable before the edit-mode double-click; "
                              L"focusTarget={}, editMode={}, path='{}', childWindows={}, pathRect=({},{}-{},{}), lastVisible={}, "
                              L"lastRect=({},{}-{},{}), panePath='{}', "
                              L"click=({},{}), refresh={}/{}, items={}/{}, selected={}/{}.",
                              static_cast<int>(afterSingleSnapshot.focusTarget),
                              afterSingleSnapshot.editMode ? 1 : 0,
                              afterSingleSnapshot.currentPathText,
                              afterSingleSnapshot.visibleChildWindowCount,
                              afterSingleSnapshot.pathRegionRect.left,
                              afterSingleSnapshot.pathRegionRect.top,
                              afterSingleSnapshot.pathRegionRect.right,
                              afterSingleSnapshot.pathRegionRect.bottom,
                              afterSingleSnapshot.pathLastSegmentVisible ? 1 : 0,
                              afterSingleSnapshot.pathLastSegmentRect.left,
                              afterSingleSnapshot.pathLastSegmentRect.top,
                              afterSingleSnapshot.pathLastSegmentRect.right,
                              afterSingleSnapshot.pathLastSegmentRect.bottom,
                              panePathAfterSingle.has_value() ? panePathAfterSingle->wstring() : std::wstring(),
                              singleClickPoint.x,
                              singleClickPoint.y,
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              baselineRefreshCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              baselineItemCount,
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              baselineSelectedCount));
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot preDoubleSnapshot{};
    const bool preDoubleStable = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                               [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == root.wstring() && value.pathRegionRect.right > value.pathRegionRect.left &&
               value.pathRegionRect.bottom > value.pathRegionRect.top &&
               value.pathLastSegmentVisible && value.pathLastSegmentRect.right > value.pathLastSegmentRect.left &&
               value.pathLastSegmentRect.bottom > value.pathLastSegmentRect.top &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                               SelfTest::Scale(1000ms),
                                                               &preDoubleSnapshot);
    const std::optional<std::filesystem::path> panePathBeforeDouble = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    state.Require(preDoubleStable,
                  std::format(L"Navigation-view path region should be stable immediately before the edit-mode double-click; "
                              L"focusTarget={}, editMode={}, path='{}', childWindows={}, pathRect=({},{}-{},{}), lastVisible={}, "
                              L"lastRect=({},{}-{},{}), panePath='{}', "
                              L"refresh={}/{}, items={}/{}, selected={}/{}.",
                              static_cast<int>(preDoubleSnapshot.focusTarget),
                              preDoubleSnapshot.editMode ? 1 : 0,
                              preDoubleSnapshot.currentPathText,
                              preDoubleSnapshot.visibleChildWindowCount,
                              preDoubleSnapshot.pathRegionRect.left,
                              preDoubleSnapshot.pathRegionRect.top,
                              preDoubleSnapshot.pathRegionRect.right,
                              preDoubleSnapshot.pathRegionRect.bottom,
                              preDoubleSnapshot.pathLastSegmentVisible ? 1 : 0,
                              preDoubleSnapshot.pathLastSegmentRect.left,
                              preDoubleSnapshot.pathLastSegmentRect.top,
                              preDoubleSnapshot.pathLastSegmentRect.right,
                              preDoubleSnapshot.pathLastSegmentRect.bottom,
                              panePathBeforeDouble.has_value() ? panePathBeforeDouble->wstring() : std::wstring(),
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              baselineRefreshCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              baselineItemCount,
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              baselineSelectedCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const POINT clickPoint        = makePathClickPoint(preDoubleSnapshot);
    const LPARAM clickPointLParam = MAKELPARAM(clickPoint.x, clickPoint.y);
    const HWND clickTarget        = ResolveMouseInputWindowForHostPoint(navigationView, clickPointLParam);
    const HWND messageTarget      = (clickTarget != nullptr && IsWindow(clickTarget) != FALSE) ? clickTarget : navigationView;
    const LPARAM targetClickPoint =
        (clickTarget != nullptr && IsWindow(clickTarget) != FALSE) ? MapClientPointLParam(navigationView, clickTarget, clickPointLParam) : clickPointLParam;

    SendMessageW(messageTarget, WM_MOUSEMOVE, 0, targetClickPoint);
    SendMessageW(messageTarget, WM_LBUTTONDOWN, MK_LBUTTON, targetClickPoint);
    SendMessageW(messageTarget, WM_LBUTTONUP, 0, targetClickPoint);
    NavigationViewDebugSnapshot afterLeadingClickSnapshot{};
    const bool afterLeadingClickSnapshotAvailable = g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, afterLeadingClickSnapshot);
    SendMessageW(messageTarget, WM_LBUTTONDBLCLK, MK_LBUTTON, targetClickPoint);
    SendMessageW(messageTarget, WM_LBUTTONUP, 0, targetClickPoint);
    PumpPendingMessages();

    NavigationViewDebugSnapshot editSnapshot{};
    const bool reachedEditMode = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                               [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && ! value.currentEditText.empty() &&
               value.visibleChildWindowCount == 1u && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.currentEditHostHwnd != nullptr &&
               value.currentEditBridgeHwnd != nullptr && value.currentEditHasSelection && value.currentEditSelectionStart == 0u &&
               value.currentEditSelectionEnd == value.currentEditText.size() && value.currentPathText == root.wstring() &&
               value.currentEditText.find(root.wstring()) != std::wstring::npos &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                               SelfTest::Scale(3000ms),
                                                               &editSnapshot);
    const std::optional<std::filesystem::path> panePathAfterDouble = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    state.Require(reachedEditMode,
                  std::format(L"Navigation view did not enter address-bar edit mode after path-region double-click; focusTarget={}, editMode={}, "
                              L"path='{}', edit='{}', childWindows={}, editHost=0x{:X}, editBridge=0x{:X}, hasSelection={}, selection=[{},{}), "
                              L"fullPathPopup={}, fullPathPopupEdit={}, suggestPopup={}, navVisible={}, navWindowVisible={}, focusedWindow=0x{:X}, "
                              L"navigationView=0x{:X}, folderView=0x{:X}, click=({},{}), clickTarget=0x{:X}, mappedClick=({},{}), "
                              L"pathRect=({},{}-{},{}), ellipsisVisible={}, ellipsisRect=({},{}-{},{}), prePath='{}', preRect=({},{}-{},{}), "
                              L"lastVisible={}, lastRect=({},{}-{},{}), leadingSnapshot={}, leadingPath='{}', leadingEditMode={}, "
                              L"leadingRect=({},{}-{},{}), leadingLastVisible={}, leadingLastRect=({},{}-{},{}), leadingChildren={}, "
                              L"panePath='{}', refresh={}/{}, items={}/{}, selected={}/{}.",
                              static_cast<int>(editSnapshot.focusTarget),
                              editSnapshot.editMode ? 1 : 0,
                              editSnapshot.currentPathText,
                              editSnapshot.currentEditText,
                              editSnapshot.visibleChildWindowCount,
                              reinterpret_cast<uintptr_t>(editSnapshot.currentEditHostHwnd),
                              reinterpret_cast<uintptr_t>(editSnapshot.currentEditBridgeHwnd),
                              editSnapshot.currentEditHasSelection ? 1 : 0,
                              editSnapshot.currentEditSelectionStart,
                              editSnapshot.currentEditSelectionEnd,
                              editSnapshot.fullPathPopupVisible ? 1 : 0,
                              editSnapshot.fullPathPopupEditMode ? 1 : 0,
                              editSnapshot.editSuggestPopupVisible ? 1 : 0,
                              g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left) ? 1 : 0,
                              IsWindowVisible(navigationView) != FALSE ? 1 : 0,
                              reinterpret_cast<uintptr_t>(GetFocus()),
                              reinterpret_cast<uintptr_t>(navigationView),
                              reinterpret_cast<uintptr_t>(folderView),
                              clickPoint.x,
                              clickPoint.y,
                              reinterpret_cast<uintptr_t>(clickTarget),
                              GET_X_LPARAM(targetClickPoint),
                              GET_Y_LPARAM(targetClickPoint),
                              editSnapshot.pathRegionRect.left,
                              editSnapshot.pathRegionRect.top,
                              editSnapshot.pathRegionRect.right,
                              editSnapshot.pathRegionRect.bottom,
                              editSnapshot.pathEllipsisVisible ? 1 : 0,
                              editSnapshot.pathEllipsisRect.left,
                              editSnapshot.pathEllipsisRect.top,
                              editSnapshot.pathEllipsisRect.right,
                              editSnapshot.pathEllipsisRect.bottom,
                              preDoubleSnapshot.currentPathText,
                              preDoubleSnapshot.pathRegionRect.left,
                              preDoubleSnapshot.pathRegionRect.top,
                              preDoubleSnapshot.pathRegionRect.right,
                              preDoubleSnapshot.pathRegionRect.bottom,
                              editSnapshot.pathLastSegmentVisible ? 1 : 0,
                              editSnapshot.pathLastSegmentRect.left,
                              editSnapshot.pathLastSegmentRect.top,
                              editSnapshot.pathLastSegmentRect.right,
                              editSnapshot.pathLastSegmentRect.bottom,
                              afterLeadingClickSnapshotAvailable ? 1 : 0,
                              afterLeadingClickSnapshot.currentPathText,
                              afterLeadingClickSnapshot.editMode ? 1 : 0,
                              afterLeadingClickSnapshot.pathRegionRect.left,
                              afterLeadingClickSnapshot.pathRegionRect.top,
                              afterLeadingClickSnapshot.pathRegionRect.right,
                              afterLeadingClickSnapshot.pathRegionRect.bottom,
                              afterLeadingClickSnapshot.pathLastSegmentVisible ? 1 : 0,
                              afterLeadingClickSnapshot.pathLastSegmentRect.left,
                              afterLeadingClickSnapshot.pathLastSegmentRect.top,
                              afterLeadingClickSnapshot.pathLastSegmentRect.right,
                              afterLeadingClickSnapshot.pathLastSegmentRect.bottom,
                              afterLeadingClickSnapshot.visibleChildWindowCount,
                              panePathAfterDouble.has_value() ? panePathAfterDouble->wstring() : std::wstring(),
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              baselineRefreshCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              baselineItemCount,
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              baselineSelectedCount));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND editBridge = editSnapshot.currentEditBridgeHwnd;
    state.Require(editBridge != nullptr && IsWindow(editBridge) != FALSE, L"Navigation view edit bridge unavailable after path-region double-click.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int editTextLength = static_cast<int>(editSnapshot.currentEditText.size());
    state.Require(editTextLength > 0, L"Navigation view edit bridge should contain the current path text after path-region double-click.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(editSnapshot.currentEditHasSelection && editSnapshot.currentEditSelectionStart == 0u &&
                      editSnapshot.currentEditSelectionEnd == static_cast<size_t>(editTextLength),
                  std::format(L"Navigation view path edit should start with the whole path selected after the opening double-click; saw selection [{}, {}).",
                              editSnapshot.currentEditSelectionStart,
                              editSnapshot.currentEditSelectionEnd));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND editHost = editSnapshot.currentEditHostHwnd;
    state.Require(editHost != nullptr && IsWindow(editHost) != FALSE, L"Navigation view edit host unavailable for in-edit double-click selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT editHostClient{};
    state.Require(GetClientRect(editHost, &editHostClient) != FALSE,
                  L"Failed to query the navigation-view edit host bounds for double-click selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG editDoubleClickX = std::clamp<LONG>((editHostClient.right * 2) / 5, 8, std::max<LONG>(8, editHostClient.right - 40));
    const LONG editDoubleClickY = std::clamp<LONG>((editHostClient.bottom - editHostClient.top) / 2, 4, std::max<LONG>(4, editHostClient.bottom - 4));
    SendMouseDoubleClickToResolvedPointWindow(editHost, MAKELPARAM(editDoubleClickX, editDoubleClickY));

    NavigationViewDebugSnapshot wordSelectionSnapshot{};
    const bool selectionChanged = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && value.currentEditHasSelection &&
               value.currentEditSelectionEnd > value.currentEditSelectionStart &&
               ! (value.currentEditSelectionStart == 0u && value.currentEditSelectionEnd == static_cast<size_t>(editTextLength));
    },
                                                                SelfTest::Scale(3000ms),
                                                                &wordSelectionSnapshot);

    state.Require(
        selectionChanged,
        std::format(L"In-edit double-click should replace the initial select-all range with a smaller word selection; final selection was [{}, {}) of {}.",
                    wordSelectionSnapshot.currentEditSelectionStart,
                    wordSelectionSnapshot.currentEditSelectionEnd,
                    editTextLength));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForFolderFocus = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.fullPathPopupVisible &&
                   ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Pane folder view did not regain focus with stable state after {}.", context));
        return state.failure.empty();
    };

    const HWND focused = GetFocus();
    state.Require(focused != nullptr && IsWindow(focused) != FALSE, L"Focused window unavailable before navigation-view path edit tab handoff.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(focused, WM_KEYDOWN, VK_TAB, 0);
    SendMessageW(focused, WM_KEYUP, VK_TAB, 0);
    PumpPendingMessages();

    state.Require(waitForFolderFocus(L"path double-click edit-mode tab handoff"),
                  L"Tab from the navigation-view path edit should return focus to the left folder view.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewPathRegionKeyboardActivationEntersEditMode(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view path-region keyboard activation test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / (L"nav_path_key_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create navigation-view path keyboard root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for navigation-view path keyboard test.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for navigation-view path keyboard test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view path keyboard test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for navigation-view path keyboard test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for navigation-view path keyboard test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before navigation-view path keyboard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for navigation-view path keyboard validation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for navigation-view path keyboard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before navigation-view path keyboard validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount    = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount         = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount     = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem = L"alpha.txt";

    const auto waitForPathRegionFocus = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::PathRegion && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation view did not focus the path shell region during {}.", context));
        return state.failure.empty();
    };

    const auto waitForFolderFocus = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == root.wstring() && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Pane folder view did not regain focus with stable state after {}.", context));
        return state.failure.empty();
    };

    const auto activatePathRegionAndExit = [&](const WPARAM virtualKey, std::wstring_view label) noexcept
    {
        state.Require(g_folderWindow.DebugFocusNavigationViewRegion(FolderWindow::Pane::Left, NavigationView::FocusRegion::Path),
                      std::format(L"Failed to focus the navigation-view path region before {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForPathRegionFocus(std::format(L"{} focus settle", label)),
                      std::format(L"Navigation view path region did not settle before {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        SendMessageW(navigationView, WM_KEYDOWN, virtualKey, 0);
        SendMessageW(navigationView, WM_KEYUP, virtualKey, 0);
        PumpPendingMessages();

        NavigationViewDebugSnapshot editSnapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && ! value.currentEditText.empty() &&
                   value.visibleChildWindowCount == 1u && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
                   ! value.fullPathPopupEditMode && value.currentPathText == root.wstring() &&
                   value.currentEditText.find(root.wstring()) != std::wstring::npos &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &editSnapshot),
                      std::format(L"Navigation view did not enter path edit mode after {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        const HWND focused = GetFocus();
        state.Require(focused != nullptr && IsWindow(focused) != FALSE, std::format(L"Focused window unavailable before {} tab handoff.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        SendMessageW(focused, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(focused, WM_KEYUP, VK_TAB, 0);
        PumpPendingMessages();

        state.Require(waitForFolderFocus(std::format(L"{} tab handoff", label)),
                      std::format(L"Tab from the navigation-view path edit should return focus to the folder view after {}.", label));
        return state.failure.empty();
    };

    state.Require(activatePathRegionAndExit(VK_RETURN, L"Enter activation"),
                  L"Navigation view path region did not round-trip cleanly through Enter activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(activatePathRegionAndExit(VK_SPACE, L"Space activation"),
                  L"Navigation view path region did not round-trip cleanly through Space activation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewPathAncestorClickNavigatesToAncestor(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view breadcrumb ancestor-click test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"navigation_view_breadcrumb_ancestor_" + NewGuidText()) / L"alpha" / L"beta" / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root.parent_path().parent_path().parent_path(), ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create breadcrumb ancestor-click root.");
    state.Require(SelfTest::WriteTextFile(root / L"anchor.txt", "anchor"), L"Failed to create anchor.txt for breadcrumb ancestor-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for breadcrumb ancestor-click test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for breadcrumb ancestor-click test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"anchor.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for breadcrumb ancestor-click test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"anchor.txt"),
                  L"Failed to focus anchor.txt before breadcrumb ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for breadcrumb ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before breadcrumb ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.currentPathText == root.wstring() &&
               value.pathAncestorSegmentVisible && value.pathAncestorSegmentRect.right > value.pathAncestorSegmentRect.left &&
               value.pathAncestorSegmentRect.bottom > value.pathAncestorSegmentRect.top && ! value.pathAncestorTargetText.empty();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture a visible breadcrumb ancestor segment before ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path expectedAncestorPath = std::filesystem::path(baselineSnapshot.pathAncestorTargetText);
    state.Require(! expectedAncestorPath.empty() && expectedAncestorPath != root,
                  L"Breadcrumb ancestor-click target should resolve to a non-current ancestor path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX = (baselineSnapshot.pathAncestorSegmentRect.left + baselineSnapshot.pathAncestorSegmentRect.right) / 2;
    const LONG clickY = (baselineSnapshot.pathAncestorSegmentRect.top + baselineSnapshot.pathAncestorSegmentRect.bottom) / 2;
    SendMouseClickToResolvedPointWindow(navigationView, MAKELPARAM(clickX, clickY));

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, expectedAncestorPath, SelfTest::Scale(3000ms)),
                  std::format(L"Breadcrumb ancestor click did not navigate to '{}'.", expectedAncestorPath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot settledSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.currentPathText == expectedAncestorPath.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &settledSnapshot),
                  L"Navigation view did not settle cleanly after breadcrumb ancestor navigation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewClickInUnfocusedPaneFocusesTargetPane(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for unfocused-pane navigation click test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path testRoot = suiteRoot / L"work" / (L"navigation_view_unfocused_pane_click_" + NewGuidText());
    const std::filesystem::path left     = testRoot / L"left";
    const std::filesystem::path right    = testRoot / L"right" / L"alpha" / L"beta" / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(testRoot, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left pane root for unfocused-pane navigation click test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right pane root for unfocused-pane navigation click test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for unfocused-pane navigation click test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for unfocused-pane navigation click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<FolderWindow::Pane> zoomBefore     = g_folderWindow.GetZoomedPane();
    const std::optional<float> zoomRestoreBefore           = g_folderWindow.GetZoomRestoreSplitRatio();
    const auto restoreZoom                                 = wil::scope_exit([&] { g_folderWindow.SetZoomState(zoomBefore, zoomRestoreBefore); });
    const auto restorePanes                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.SetZoomState(std::nullopt, std::nullopt);
    PumpPendingMessages();
    state.Require(! g_folderWindow.GetZoomedPane().has_value(),
                  L"Unfocused-pane navigation click validation requires both panes visible and could not clear the prior zoom state.");
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::PreviewPaneDebugSnapshot inheritedPreview{};
    if (g_folderWindow.DebugGetPreviewPaneSnapshot(inheritedPreview) && inheritedPreview.active)
    {
        g_folderWindow.TogglePreviewPane(inheritedPreview.sourcePane);
        const auto previewCloseDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        bool previewClosed              = false;
        do
        {
            PumpPendingMessages();
            FolderWindow::PreviewPaneDebugSnapshot currentPreview{};
            previewClosed = g_folderWindow.DebugGetPreviewPaneSnapshot(currentPreview) && ! currentPreview.active;
            if (previewClosed)
            {
                break;
            }
            std::this_thread::sleep_for(20ms);
        } while (std::chrono::steady_clock::now() < previewCloseDeadline);

        state.Require(previewClosed,
                      std::format(L"Unfocused-pane navigation click validation could not close inherited preview pane state; source={} host={} previewTab={}.",
                                  inheritedPreview.sourcePane == FolderWindow::Pane::Left ? L"left" : L"right",
                                  inheritedPreview.hostPane == FolderWindow::Pane::Left ? L"left" : L"right",
                                  inheritedPreview.previewTabSelected ? L"yes" : L"no"));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for the left pane during unfocused-pane navigation click test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for the right pane during unfocused-pane navigation click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for unfocused-pane navigation click test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for unfocused-pane navigation click test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for unfocused-pane navigation click test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for unfocused-pane navigation click test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before unfocused-pane navigation click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView      = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND rightFolderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right);
    const HWND rightNavigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Right);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder view handle unavailable for unfocused-pane navigation click validation.");
    state.Require(rightFolderView != nullptr && IsWindow(rightFolderView) != FALSE,
                  L"Right folder view handle unavailable for unfocused-pane navigation click validation.");
    state.Require(rightNavigationView != nullptr && IsWindow(rightNavigationView) != FALSE,
                  L"Right navigation-view handle unavailable for unfocused-pane navigation click validation.");
    state.Require(WaitForFocusedFolderViewForMainMenu(leftFolderView, SelfTest::Scale(2000ms)),
                  L"Left folder view did not take baseline focus before unfocused-pane navigation click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Right, rightNavigationView, SelfTest::Scale(3000ms)),
                  std::format(L"Right navigation bar did not become visible before unfocused-pane navigation click validation; {}.",
                              DescribeNavigationViewVisibilityForSelfTest(FolderWindow::Pane::Right, rightNavigationView)));
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.currentPathText == right.wstring() &&
               value.pathAncestorSegmentVisible && value.pathAncestorSegmentRect.right > value.pathAncestorSegmentRect.left &&
               value.pathAncestorSegmentRect.bottom > value.pathAncestorSegmentRect.top && ! value.pathAncestorTargetText.empty();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture a visible right-pane breadcrumb ancestor segment before unfocused-pane navigation click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path expectedAncestorPath = std::filesystem::path(baselineSnapshot.pathAncestorTargetText);
    state.Require(! expectedAncestorPath.empty() && expectedAncestorPath != right,
                  L"Unfocused-pane breadcrumb click target should resolve to a non-current ancestor path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX = (baselineSnapshot.pathAncestorSegmentRect.left + baselineSnapshot.pathAncestorSegmentRect.right) / 2;
    const LONG clickY = (baselineSnapshot.pathAncestorSegmentRect.top + baselineSnapshot.pathAncestorSegmentRect.bottom) / 2;
    SendMouseClickToResolvedPointWindow(rightNavigationView, MAKELPARAM(clickX, clickY));

    state.Require(WaitForPanePath(FolderWindow::Pane::Right, expectedAncestorPath, SelfTest::Scale(3000ms)),
                  std::format(L"Right-pane breadcrumb click did not navigate to '{}'.", expectedAncestorPath.wstring()));
    state.Require(WaitForFocusedFolderViewForMainMenu(rightFolderView, SelfTest::Scale(2000ms)),
                  L"Clicking navigation in the unfocused right pane should move keyboard focus to the right folder view.");
    state.Require(g_folderWindow.GetActivePane() == FolderWindow::Pane::Right,
                  L"Clicking navigation in the unfocused right pane should make the right pane active.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewFullPathPopupEditRoute(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view full-path popup test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path popupRoot = suiteRoot / L"w" / (L"nvp_" + NewGuidText()) / L"alpha_segment_name_for_popup" / L"beta_segment_name_for_popup" /
                                            L"gamma_segment_name_for_popup" / L"delta_segment_name_for_popup";
    std::error_code ec;
    std::filesystem::remove_all(popupRoot, ec);
    state.Require(SelfTest::EnsureDirectory(popupRoot), L"Failed to create navigation-view full-path popup root.");
    state.Require(SelfTest::WriteTextFile(popupRoot / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for navigation-view full-path popup test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view full-path popup test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, popupRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, popupRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for navigation-view full-path popup test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for navigation-view full-path popup test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for navigation-view full-path popup test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for navigation-view full-path popup test.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for navigation-view full-path popup test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before navigation-view full-path popup test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const UINT originalMainDpi = GetDpiForWindow(mainWindow);
    const UINT popupTestDpi    = originalMainDpi >= 144u ? 192u : 144u;
    RECT originalMainRect{};
    state.Require(GetWindowRect(mainWindow, &originalMainRect) != FALSE, L"Main window rect unavailable for navigation-view full-path popup DPI validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto scaleMainRectForDpi = [&](const UINT targetDpi) noexcept
    {
        RECT rect           = originalMainRect;
        const LONG widthPx  = std::max<LONG>(1, rect.right - rect.left);
        const LONG heightPx = std::max<LONG>(1, rect.bottom - rect.top);
        rect.right          = rect.left + MulDiv(widthPx, static_cast<int>(targetDpi), static_cast<int>(originalMainDpi));
        rect.bottom         = rect.top + MulDiv(heightPx, static_cast<int>(targetDpi), static_cast<int>(originalMainDpi));
        return rect;
    };

    const auto restoreMainWindowDpi = wil::scope_exit([&]() noexcept
    {
        SendMessageW(mainWindow,
                     WM_DPICHANGED,
                     static_cast<WPARAM>(static_cast<DWORD>(MAKELONG(static_cast<WORD>(originalMainDpi), static_cast<WORD>(originalMainDpi)))),
                     reinterpret_cast<LPARAM>(&originalMainRect));
        PumpPendingMessages();
    });

    const DWORD processId                  = GetCurrentProcessId();
    const auto baselineWindows             = SnapshotTopLevelWindowsForProcess(processId);
    const std::wstring expectedFocusedItem = L"alpha.txt";
    const uint64_t baselineRefreshCount    = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount         = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount     = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode &&
               value.pathEllipsisVisible && value.visibleChildWindowCount == 0u && value.currentPathText == popupRoot.wstring() &&
               value.pathEllipsisRect.right > value.pathEllipsisRect.left && value.pathEllipsisRect.bottom > value.pathEllipsisRect.top;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture baseline navigation-view ellipsis state for full-path popup test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForPopupWindow = [&](HWND& popupHwnd, std::wstring_view context) noexcept
    {
        constexpr std::wstring_view popupClassName = L"RedSalamander.FullPathPopup";

        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();

            const auto currentWindows = SnapshotTopLevelWindowsForProcess(processId);
            for (const uintptr_t raw : currentWindows)
            {
                if (baselineWindows.contains(raw))
                {
                    continue;
                }

                const HWND hwnd = reinterpret_cast<HWND>(raw);
                if (! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
                {
                    continue;
                }

                wchar_t className[128]{};
                const int classLength = GetClassNameW(hwnd, className, static_cast<int>(std::size(className)));
                if (classLength <= 0)
                {
                    continue;
                }

                if (std::wstring_view(className, static_cast<size_t>(classLength)) == popupClassName)
                {
                    popupHwnd = hwnd;
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        popupHwnd = nullptr;
        state.Require(false, std::format(L"Navigation-view full-path popup window did not appear during {}.", context));
        return false;
    };

    const LONG clickX = (baselineSnapshot.pathEllipsisRect.left + baselineSnapshot.pathEllipsisRect.right) / 2;
    const LONG clickY = (baselineSnapshot.pathEllipsisRect.top + baselineSnapshot.pathEllipsisRect.bottom) / 2;
    SendMouseClickToResolvedPointWindow(navigationView, MAKELPARAM(clickX, clickY));

    NavigationViewDebugSnapshot popupSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.pathEllipsisVisible &&
               value.visibleChildWindowCount == 0u && value.currentPathText == popupRoot.wstring() &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &popupSnapshot),
                  L"Navigation view did not open the full-path popup after clicking the visible ellipsis.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND popupHwnd = nullptr;
    state.Require(waitForPopupWindow(popupHwnd, L"full-path popup open"), L"Navigation-view full-path popup handle unavailable after ellipsis click.");
    if (! state.failure.empty())
    {
        return false;
    }

    const SIZE popupClientSizeBeforeDpi = popupSnapshot.fullPathPopupClientSize;
    const RECT popupScaledRect          = scaleMainRectForDpi(popupTestDpi);
    SendMessageW(mainWindow,
                 WM_DPICHANGED,
                 static_cast<WPARAM>(static_cast<DWORD>(MAKELONG(static_cast<WORD>(popupTestDpi), static_cast<WORD>(popupTestDpi)))),
                 reinterpret_cast<LPARAM>(&popupScaledRect));
    PumpPendingMessages();

    NavigationViewDebugSnapshot dpiPopupSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.currentPathText == popupRoot.wstring() &&
               value.dpi == popupTestDpi &&
               (value.fullPathPopupClientSize.cx > popupClientSizeBeforeDpi.cx || value.fullPathPopupClientSize.cy > popupClientSizeBeforeDpi.cy) &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &dpiPopupSnapshot),
                  std::format(L"Navigation view full-path popup did not reflow across the synthetic DPI transition. popupVisible={} popupEditMode={} dpi={} "
                              L"client={}x{} baseline={}x{} currentPath='{}'.",
                              popupSnapshot.fullPathPopupVisible ? L"yes" : L"no",
                              popupSnapshot.fullPathPopupEditMode ? L"yes" : L"no",
                              dpiPopupSnapshot.dpi,
                              dpiPopupSnapshot.fullPathPopupClientSize.cx,
                              dpiPopupSnapshot.fullPathPopupClientSize.cy,
                              popupClientSizeBeforeDpi.cx,
                              popupClientSizeBeforeDpi.cy,
                              dpiPopupSnapshot.currentPathText));
    state.Require(IsWindow(popupHwnd) != FALSE && IsWindowVisible(popupHwnd) != FALSE,
                  L"Navigation-view full-path popup should stay alive and visible across the synthetic DPI transition.");
    if (! state.failure.empty())
    {
        return false;
    }

    static_cast<void>(SetForegroundWindow(popupHwnd));
    state.Require(SetFocus(popupHwnd) == popupHwnd, L"Navigation-view full-path popup should accept Win32 focus before keyboard edit routing.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(popupHwnd, WM_KEYDOWN, VK_F4, 0);
    SendMessageW(popupHwnd, WM_KEYUP, VK_F4, 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot editSnapshot{};
    const bool reachedEditMode = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                               [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && value.fullPathPopupVisible && value.fullPathPopupEditMode &&
               value.focusTarget == NavigationViewDebugFocusTarget::FullPathPopupEdit && ! value.currentEditText.empty() &&
               value.currentEditText.find(popupRoot.wstring()) != std::wstring::npos && value.visibleChildWindowCount == 0u &&
               value.currentPathText == popupRoot.wstring() && g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                               SelfTest::Scale(3000ms),
                                                               &editSnapshot);
    if (! reachedEditMode)
    {
        NavigationViewDebugSnapshot currentSnapshot{};
        if (g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, currentSnapshot))
        {
            editSnapshot = currentSnapshot;
        }

        const HWND focused = GetFocus();
        wchar_t focusedClass[128]{};
        if (focused)
        {
            static_cast<void>(GetClassNameW(focused, focusedClass, static_cast<int>(std::size(focusedClass))));
        }

        state.Require(false,
                      std::format(L"Navigation view did not enter full-path popup edit mode after the popup keyboard edit route. "
                                  L"popupVisible={} popupEditMode={} focusTarget={} visibleChildren={} focusedHwnd=0x{:X} focusedClass='{}' "
                                  L"currentEdit='{}' currentPath='{}' refreshCount={} itemCount={} selectedCount={}.",
                                  editSnapshot.fullPathPopupVisible ? 1 : 0,
                                  editSnapshot.fullPathPopupEditMode ? 1 : 0,
                                  static_cast<unsigned int>(editSnapshot.focusTarget),
                                  editSnapshot.visibleChildWindowCount,
                                  reinterpret_cast<uintptr_t>(focused),
                                  focusedClass,
                                  editSnapshot.currentEditText,
                                  editSnapshot.currentPathText,
                                  g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left)));
        return false;
    }

    const HWND popupEdit = editSnapshot.currentEditBridgeHwnd;
    state.Require(popupEdit != nullptr && IsWindow(popupEdit) != FALSE, L"Navigation-view full-path popup edit bridge should exist before escape cleanup.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! editSnapshot.currentEditText.empty(), L"Navigation-view full-path popup edit text should not be empty.");
    state.Require(
        editSnapshot.currentEditText.find(popupRoot.wstring()) != std::wstring::npos,
        std::format(L"Navigation-view full-path popup edit text should include '{}' but saw '{}'.", popupRoot.wstring(), editSnapshot.currentEditText));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(popupEdit, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(popupEdit, WM_KEYUP, VK_ESCAPE, 0);
    PumpPendingMessages();

    NavigationViewDebugSnapshot cleanupSnapshot{};
    const bool cleanupReached = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                              [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == popupRoot.wstring() && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                              SelfTest::Scale(3000ms),
                                                              &cleanupSnapshot);
    if (! cleanupReached)
    {
        NavigationViewDebugSnapshot currentSnapshot{};
        static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, currentSnapshot));

        const HWND focusedWindow = GetFocus();
        wchar_t focusedClassName[128]{};
        if (focusedWindow && IsWindow(focusedWindow) != FALSE)
        {
            static_cast<void>(GetClassNameW(focusedWindow, focusedClassName, static_cast<int>(std::size(focusedClassName))));
        }

        state.Require(false,
                      std::format(L"Escape from the navigation-view full-path popup edit should close the popup and return focus to the left folder view. "
                                  L"After timeout: popupVisible={} popupEditMode={} focusTarget={} visibleChildren={} focusedHwnd=0x{:X} focusedClass='{}' "
                                  L"focusedFolderView=0x{:X} expectedFolderView=0x{:X} focusedItem='{}' refreshCount={} itemCount={} selectedCount={}.",
                                  currentSnapshot.fullPathPopupVisible ? 1 : 0,
                                  currentSnapshot.fullPathPopupEditMode ? 1 : 0,
                                  static_cast<unsigned int>(currentSnapshot.focusTarget),
                                  currentSnapshot.visibleChildWindowCount,
                                  static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(focusedWindow)),
                                  focusedClassName,
                                  static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd())),
                                  static_cast<unsigned long long>(reinterpret_cast<uintptr_t>(folderView)),
                                  g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                                  g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left)));
    }
    state.Require(WaitForNoNonBaselineWindows(processId, baselineWindows, mainWindow, SelfTest::Scale(3000ms)),
                  L"Navigation-view full-path popup should not leave non-baseline top-level windows behind after escape cleanup.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewFullPathPopupAncestorClickNavigatesToAncestor(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view full-path ancestor-click test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path popupRoot = suiteRoot / L"w" / (L"nvp_ancestor_" + NewGuidText()) / L"alpha_segment_name_for_popup" /
                                            L"beta_segment_name_for_popup" / L"gamma_segment_name_for_popup" / L"delta_segment_name_for_popup";
    std::error_code ec;
    std::filesystem::remove_all(popupRoot.parent_path().parent_path().parent_path().parent_path(), ec);
    state.Require(SelfTest::EnsureDirectory(popupRoot), L"Failed to create full-path popup ancestor-click root.");
    state.Require(SelfTest::WriteTextFile(popupRoot / L"anchor.txt", "anchor"), L"Failed to create anchor.txt for full-path popup ancestor-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for full-path popup ancestor-click test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, popupRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, popupRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for full-path popup ancestor-click test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"anchor.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for full-path popup ancestor-click test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"anchor.txt"),
                  L"Failed to focus anchor.txt before full-path popup ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for full-path popup ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before full-path popup ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD processId      = GetCurrentProcessId();
    const auto baselineWindows = SnapshotTopLevelWindowsForProcess(processId);
    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode &&
               value.pathEllipsisVisible && value.visibleChildWindowCount == 0u && value.currentPathText == popupRoot.wstring() &&
               value.pathEllipsisRect.right > value.pathEllipsisRect.left && value.pathEllipsisRect.bottom > value.pathEllipsisRect.top;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture baseline ellipsis state before full-path popup ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForPopupWindow = [&](HWND& popupHwnd, std::wstring_view context) noexcept
    {
        constexpr std::wstring_view popupClassName = L"RedSalamander.FullPathPopup";

        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();

            const auto currentWindows = SnapshotTopLevelWindowsForProcess(processId);
            for (const uintptr_t raw : currentWindows)
            {
                if (baselineWindows.contains(raw))
                {
                    continue;
                }

                const HWND hwnd = reinterpret_cast<HWND>(raw);
                if (! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
                {
                    continue;
                }

                wchar_t className[128]{};
                const int classLength = GetClassNameW(hwnd, className, static_cast<int>(std::size(className)));
                if (classLength <= 0)
                {
                    continue;
                }

                if (std::wstring_view(className, static_cast<size_t>(classLength)) == popupClassName)
                {
                    popupHwnd = hwnd;
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        popupHwnd = nullptr;
        state.Require(false, std::format(L"Navigation-view full-path popup window did not appear during {}.", context));
        return false;
    };

    const LONG ellipsisClickX = (baselineSnapshot.pathEllipsisRect.left + baselineSnapshot.pathEllipsisRect.right) / 2;
    const LONG ellipsisClickY = (baselineSnapshot.pathEllipsisRect.top + baselineSnapshot.pathEllipsisRect.bottom) / 2;
    SendMouseClickToResolvedPointWindow(navigationView, MAKELPARAM(ellipsisClickX, ellipsisClickY));

    NavigationViewDebugSnapshot popupSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.pathEllipsisVisible &&
               value.visibleChildWindowCount == 0u && value.currentPathText == popupRoot.wstring() && value.fullPathPopupAncestorSegmentVisible &&
               value.fullPathPopupAncestorSegmentRect.right > value.fullPathPopupAncestorSegmentRect.left &&
               value.fullPathPopupAncestorSegmentRect.bottom > value.fullPathPopupAncestorSegmentRect.top && ! value.fullPathPopupAncestorTargetText.empty();
    },
                                                SelfTest::Scale(3000ms),
                                                &popupSnapshot),
                  L"Navigation view did not expose a clickable full-path popup ancestor segment.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND popupHwnd = nullptr;
    state.Require(waitForPopupWindow(popupHwnd, L"full-path popup ancestor-click open"),
                  L"Navigation-view full-path popup handle unavailable before ancestor-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path expectedAncestorPath = std::filesystem::path(popupSnapshot.fullPathPopupAncestorTargetText);
    state.Require(! expectedAncestorPath.empty() && expectedAncestorPath != popupRoot,
                  L"Full-path popup ancestor-click target should resolve to a non-current ancestor path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG ancestorClickX = (popupSnapshot.fullPathPopupAncestorSegmentRect.left + popupSnapshot.fullPathPopupAncestorSegmentRect.right) / 2;
    const LONG ancestorClickY = (popupSnapshot.fullPathPopupAncestorSegmentRect.top + popupSnapshot.fullPathPopupAncestorSegmentRect.bottom) / 2;
    SendMouseClickToResolvedPointWindow(popupHwnd, MAKELPARAM(ancestorClickX, ancestorClickY));

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, expectedAncestorPath, SelfTest::Scale(3000ms)),
                  std::format(L"Full-path popup ancestor click did not navigate to '{}'.", expectedAncestorPath.wstring()));
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot settledSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.currentPathText == expectedAncestorPath.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &settledSnapshot),
                  L"Navigation view did not settle cleanly after full-path popup ancestor navigation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForNoNonBaselineWindows(processId, baselineWindows, mainWindow, SelfTest::Scale(3000ms)),
                  L"Full-path popup ancestor navigation should not leave non-baseline top-level windows behind.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewRegionTraversal(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view region traversal test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"navigation_view_region_traversal_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create navigation-view region traversal root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for navigation-view region traversal test.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for navigation-view region traversal test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view region traversal test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for navigation-view region traversal test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for navigation-view region traversal test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for navigation-view region traversal test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for navigation-view region traversal test.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for navigation-view region traversal test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before navigation-view region traversal test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount    = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount         = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount     = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem = L"alpha.txt";

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    { return ! value.editMode && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.currentPathText == root.wstring(); },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture baseline navigation-view snapshot for region traversal test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForNavRegion = [&](const NavigationViewDebugFocusTarget expectedTarget, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == expectedTarget && ! value.editMode && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode &&
                   value.visibleChildWindowCount == 0u && value.currentPathText == root.wstring() &&
                   value.showMenuSection == baselineSnapshot.showMenuSection && value.showDiskInfoSection == baselineSnapshot.showDiskInfoSection &&
                   value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation view did not focus the expected shell region during {}.", context));
        return state.failure.empty();
    };

    const auto waitForFolderFocus = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.fullPathPopupVisible &&
                   ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Pane folder view did not regain focus with stable state during {}.", context));
        return state.failure.empty();
    };

    const auto sendHostTab = [&](const bool reverse, std::wstring_view context) noexcept
    {
        const HWND focused = GetFocus();
        state.Require(focused == navigationView, std::format(L"Navigation view host should own focus before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        BYTE keyboardState[256]{};
        bool keyboardStateCaptured      = false;
        const auto restoreKeyboardState = wil::scope_exit([&]() noexcept
        {
            if (keyboardStateCaptured)
            {
                SetKeyboardState(keyboardState);
            }
        });
        if (reverse)
        {
            keyboardStateCaptured = GetKeyboardState(keyboardState) != FALSE;
            if (keyboardStateCaptured)
            {
                BYTE shiftedState[256]{};
                std::copy(std::begin(keyboardState), std::end(keyboardState), std::begin(shiftedState));
                shiftedState[VK_SHIFT] = 0x80;
                SetKeyboardState(shiftedState);
            }
        }

        if (reverse)
        {
            SendMessageW(navigationView, WM_KEYDOWN, VK_SHIFT, 0);
        }
        SendMessageW(navigationView, WM_KEYDOWN, VK_TAB, 0);
        SendMessageW(navigationView, WM_KEYUP, VK_TAB, 0);
        if (reverse)
        {
            SendMessageW(navigationView, WM_KEYUP, VK_SHIFT, 0);
        }
        return true;
    };

    const auto focusRegion = [&](NavigationView::FocusRegion region, NavigationViewDebugFocusTarget expectedTarget, std::wstring_view context) noexcept
    {
        state.Require(g_folderWindow.DebugFocusNavigationViewRegion(FolderWindow::Pane::Left, region),
                      std::format(L"Failed to focus navigation-view region during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        return waitForNavRegion(expectedTarget, context);
    };

    std::vector<std::pair<NavigationView::FocusRegion, NavigationViewDebugFocusTarget>> order;
    if (baselineSnapshot.showMenuSection)
    {
        order.emplace_back(NavigationView::FocusRegion::Menu, NavigationViewDebugFocusTarget::MenuRegion);
    }
    order.emplace_back(NavigationView::FocusRegion::Path, NavigationViewDebugFocusTarget::PathRegion);
    order.emplace_back(NavigationView::FocusRegion::History, NavigationViewDebugFocusTarget::HistoryRegion);
    if (baselineSnapshot.showDiskInfoSection)
    {
        order.emplace_back(NavigationView::FocusRegion::DiskInfo, NavigationViewDebugFocusTarget::DiskInfoRegion);
    }

    state.Require(order.size() >= 2u, L"Navigation view should expose at least path and history regions.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(focusRegion(order.front().first, order.front().second, L"initial forward region focus"),
                  L"Failed to focus the first navigation-view shell region.");
    if (! state.failure.empty())
    {
        return false;
    }

    for (size_t i = 1; i < order.size(); ++i)
    {
        state.Require(sendHostTab(false, std::format(L"forward navigation region step {}", i)), L"Failed to send forward Tab to the navigation-view host.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForNavRegion(order[i].second, std::format(L"forward navigation region step {}", i)),
                      L"Forward Tab should advance to the next visible navigation-view shell region.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(sendHostTab(false, L"forward navigation handoff to folder view"), L"Failed to send forward Tab from the last navigation-view shell region.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForFolderFocus(L"forward navigation handoff to folder view"),
                  L"Forward Tab from the last navigation-view shell region should return focus to the pane folder view.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(focusRegion(order.back().first, order.back().second, L"initial reverse region focus"),
                  L"Failed to focus the last navigation-view shell region.");
    if (! state.failure.empty())
    {
        return false;
    }

    for (size_t i = order.size(); i-- > 1u;)
    {
        state.Require(sendHostTab(true, std::format(L"reverse navigation region step {}", i - 1u)),
                      L"Failed to send reverse Shift+Tab to the navigation-view host.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForNavRegion(order[i - 1u].second, std::format(L"reverse navigation region step {}", i - 1u)),
                      L"Reverse Shift+Tab should move to the previous visible navigation-view shell region.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(sendHostTab(true, L"reverse navigation handoff to folder view"),
                  L"Failed to send reverse Shift+Tab from the first navigation-view shell region.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForFolderFocus(L"reverse navigation handoff to folder view"),
                  L"Reverse Shift+Tab from the first navigation-view shell region should return focus to the pane folder view.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewHistoryDropdownKeyboardNavigation(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view history-dropdown test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"navigation_view_history_dropdown_" + NewGuidText());
    const std::filesystem::path rootA = root / L"A";
    const std::filesystem::path rootB = root / L"B";
    const std::filesystem::path rootC = root / L"C";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(rootA), L"Failed to create history path A.");
    state.Require(SelfTest::EnsureDirectory(rootB), L"Failed to create history path B.");
    state.Require(SelfTest::EnsureDirectory(rootC), L"Failed to create history path C.");
    state.Require(SelfTest::WriteTextFile(rootA / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history path A.");
    state.Require(SelfTest::WriteTextFile(rootB / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history path B.");
    state.Require(SelfTest::WriteTextFile(rootC / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history path C.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view history-dropdown test.");

    const auto navigatePane = [&](const std::filesystem::path& path, std::wstring_view context) noexcept
    {
        g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, path);
        state.Require(WaitForPanePath(FolderWindow::Pane::Left, path, SelfTest::Scale(3000ms)),
                      std::format(L"Failed to set left pane path to '{}' during {}.", path.wstring(), context));
        state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                      std::format(L"Pane contents not ready at '{}' during {}.", path.wstring(), context));
        return state.failure.empty();
    };

    state.Require(navigatePane(rootA, L"initial history seeding"), L"Navigation-view history seeding failed at A.");
    state.Require(navigatePane(rootB, L"initial history seeding"), L"Navigation-view history seeding failed at B.");
    state.Require(navigatePane(rootC, L"initial history seeding"), L"Navigation-view history seeding failed at C.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before navigation-view history-dropdown validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE, L"Navigation view handle unavailable for history-dropdown validation.");
    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before history-dropdown validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugFocusNavigationViewRegion(FolderWindow::Pane::Left, NavigationView::FocusRegion::History),
                  L"Failed to focus the NavigationView history region before keyboard dropdown validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::HistoryRegion && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == rootC.wstring() && value.historyCount >= 2u;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the focused NavigationView history-region baseline before keyboard dropdown validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct HistoryPopupKeyboardResult
    {
        bool popupObserved      = false;
        bool popupOwnerValid    = false;
        bool popupStateCaptured = false;
        bool selectionMoved     = false;
        bool popupClosed        = false;
        int selectedIndex       = -1;
        int targetIndex         = -1;
    } popupResult{};

    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(3000ms));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return;
        }

        popupResult.popupObserved   = true;
        const HWND popupOwner       = GetWindow(popup, GW_OWNER);
        popupResult.popupOwnerValid = popupOwner == navigationView || popupOwner == GetAncestor(navigationView, GA_ROOT);

        RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
        if (! RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState) || ! popupState.keyboardIndex.has_value())
        {
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
            popupResult.popupClosed = WaitForWindowClosed(popup, SelfTest::Scale(2000ms));
            return;
        }

        popupResult.popupStateCaptured = true;
        popupResult.selectedIndex      = static_cast<int>(popupState.keyboardIndex.value());
        const int itemCount            = static_cast<int>(baselineSnapshot.historyCount);
        if (itemCount < 2)
        {
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
            popupResult.popupClosed = WaitForWindowClosed(popup, SelfTest::Scale(2000ms));
            return;
        }

        popupResult.targetIndex = popupResult.selectedIndex > 0 ? (popupResult.selectedIndex - 1) : (popupResult.selectedIndex + 1);
        if (popupResult.targetIndex < 0 || popupResult.targetIndex >= itemCount)
        {
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
            popupResult.popupClosed = WaitForWindowClosed(popup, SelfTest::Scale(2000ms));
            return;
        }

        const UINT directionKey = popupResult.targetIndex < popupResult.selectedIndex ? VK_UP : VK_DOWN;
        PostMessageW(popup, WM_KEYDOWN, directionKey, 0);
        PostMessageW(popup, WM_KEYUP, directionKey, 0);

        const auto selectionDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < selectionDeadline)
        {
            RedSalamander::DxUi::ContextMenuPopupDebugState currentState{};
            if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, currentState) && currentState.keyboardIndex.has_value() &&
                static_cast<int>(currentState.keyboardIndex.value()) == popupResult.targetIndex)
            {
                popupResult.selectionMoved = true;
                break;
            }

            std::this_thread::sleep_for(10ms);
        }

        if (! popupResult.selectionMoved)
        {
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
            popupResult.popupClosed = WaitForWindowClosed(popup, SelfTest::Scale(2000ms));
            return;
        }

        PostMessageW(popup, WM_KEYDOWN, VK_RETURN, 0);
        PostMessageW(popup, WM_KEYUP, VK_RETURN, 0);
        popupResult.popupClosed = WaitForWindowClosed(popup, SelfTest::Scale(2000ms));
    });

    SendMessageW(navigationView, WM_KEYDOWN, VK_SPACE, 0);
    SendMessageW(navigationView, WM_KEYUP, VK_SPACE, 0);
    PumpPendingMessages();
    popupDriver.join();

    state.Require(popupResult.popupObserved, L"NavigationView history region did not open the live DxUI dropdown through keyboard activation.");
    state.Require(popupResult.popupOwnerValid, L"Focused DxUI history popup should remain owned by the active pane window hierarchy.");
    state.Require(popupResult.popupStateCaptured, L"Failed to query the live DxUI history popup state after keyboard activation.");
    state.Require(popupResult.selectedIndex >= 0 && popupResult.selectedIndex < static_cast<int>(baselineSnapshot.historyCount),
                  std::format(L"NavigationView history dropdown should expose a valid selected index; saw {} of {}.",
                              popupResult.selectedIndex,
                              baselineSnapshot.historyCount));
    state.Require(popupResult.targetIndex >= 0 && popupResult.targetIndex < static_cast<int>(baselineSnapshot.historyCount),
                  L"NavigationView history dropdown did not expose an adjacent keyboard-selection target.");
    state.Require(popupResult.selectionMoved, L"NavigationView history dropdown did not expose the adjacent keyboard-selection target in the DxUI popup.");
    state.Require(popupResult.popupClosed, L"NavigationView history dropdown should close after accepting the DxUI popup keyboard selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path expectedTarget = rootB;
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, expectedTarget, SelfTest::Scale(3000ms)),
                  std::format(L"NavigationView history dropdown keyboard selection did not navigate to '{}'.", expectedTarget.wstring()));
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents did not settle after NavigationView history dropdown keyboard selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot closedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode &&
               value.visibleChildWindowCount == 0u && value.currentPathText == expectedTarget.wstring() && value.historyCount >= baselineSnapshot.historyCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &closedSnapshot),
                  L"NavigationView history dropdown should close cleanly after keyboard path selection.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewMenuRegionKeyboardActivationOpensMenu(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    DWORD processId        = 0;
    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, &processId);
    state.Require(uiThreadId != 0u, L"Failed to resolve the UI thread for navigation-view menu-region keyboard activation.");
    if (uiThreadId == 0u)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view menu-region keyboard activation.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"navigation_view_menu_region_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create the navigation-view menu-region keyboard-activation root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for menu-region keyboard activation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for menu-region keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set the local file-system plugin for navigation-view menu-region keyboard activation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set the left pane path for navigation-view menu-region keyboard activation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for navigation-view menu-region keyboard activation.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before navigation-view menu-region keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for navigation-view menu-region keyboard activation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for navigation-view menu-region keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before navigation-view menu-region keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left));

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.showMenuSection && value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view snapshot before menu-region keyboard activation.");
    state.Require(baselineSnapshot.showMenuSection, L"Navigation view should expose the menu region before menu-region keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForMenuRegionFocus = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::MenuRegion && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.showMenuSection && value.currentPathText == root.wstring() &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation view did not focus the menu shell region during {}.", context));
        return state.failure.empty();
    };

    const auto waitForClosedShellState = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
                   ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.showMenuSection && value.currentPathText == root.wstring() &&
                   value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation view did not settle back to its shell baseline during {}.", context));
        return state.failure.empty();
    };

    const auto activateAndDismissMenu = [&](const WPARAM virtualKey, std::wstring_view label) noexcept
    {
        state.Require(g_folderWindow.DebugFocusNavigationViewRegion(FolderWindow::Pane::Left, NavigationView::FocusRegion::Menu),
                      std::format(L"Failed to focus the navigation-view menu region before {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForMenuRegionFocus(std::format(L"{} focus settle", label)),
                      std::format(L"Navigation view menu region did not settle before {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        const UiMenuRoundTripResult roundTrip =
            TriggerAndDismissKeyboardActivatedUiMenu(navigationView, virtualKey, uiThreadId, navigationView, SelfTest::Scale(2000ms), SelfTest::Scale(3000ms));

        state.Require(roundTrip.opened,
                      std::format(L"Navigation view menu-region keyboard activation did not enter menu mode or materialize a DxUI popup during {}.", label));
        state.Require(roundTrip.closed,
                      std::format(L"Navigation view menu-region keyboard activation did not leave menu mode or dismiss the DxUI popup during {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForClosedShellState(std::format(L"{} post-close settle", label)),
                      std::format(L"Navigation view shell state did not settle after {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }
        return state.failure.empty();
    };

    state.Require(activateAndDismissMenu(VK_RETURN, L"Enter activation"), L"Navigation view menu region did not round-trip cleanly through Enter activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(activateAndDismissMenu(VK_SPACE, L"Space activation"), L"Navigation view menu region did not round-trip cleanly through Space activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView || GetFocus() == navigationView,
                  L"Navigation view menu-region keyboard activation should return focus to the pane shell or folder view after dismissal.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewDiskInfoRegionKeyboardActivationOpensMenu(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    DWORD processId        = 0;
    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, &processId);
    state.Require(uiThreadId != 0u, L"Failed to resolve the UI thread for navigation-view disk-info keyboard activation.");
    if (uiThreadId == 0u)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view disk-info keyboard activation.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"navigation_view_disk_info_region_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create the navigation-view disk-info keyboard-activation root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for disk-info keyboard activation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta"), L"Failed to create beta.txt for disk-info keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set the local file-system plugin for navigation-view disk-info keyboard activation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set the left pane path for navigation-view disk-info keyboard activation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for navigation-view disk-info keyboard activation.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before navigation-view disk-info keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for navigation-view disk-info keyboard activation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for navigation-view disk-info keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before navigation-view disk-info keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left));

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.showDiskInfoSection && value.currentPathText == root.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view snapshot before disk-info keyboard activation.");
    state.Require(baselineSnapshot.showDiskInfoSection, L"Navigation view should expose the disk-info region before disk-info keyboard activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForDiskInfoFocus = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::DiskInfoRegion && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.showDiskInfoSection && value.currentPathText == root.wstring() &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation view did not focus the disk-info shell region during {}.", context));
        return state.failure.empty();
    };

    const auto waitForClosedShellState = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
                   ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u && value.showDiskInfoSection &&
                   value.currentPathText == root.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation view did not settle back to its shell baseline during {}.", context));
        return state.failure.empty();
    };

    const auto activateAndDismissMenu = [&](const WPARAM virtualKey, std::wstring_view label) noexcept
    {
        state.Require(g_folderWindow.DebugFocusNavigationViewRegion(FolderWindow::Pane::Left, NavigationView::FocusRegion::DiskInfo),
                      std::format(L"Failed to focus the navigation-view disk-info region before {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForDiskInfoFocus(std::format(L"{} focus settle", label)),
                      std::format(L"Navigation view disk-info region did not settle before {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        const UiMenuRoundTripResult roundTrip =
            TriggerAndDismissKeyboardActivatedUiMenu(navigationView, virtualKey, uiThreadId, navigationView, SelfTest::Scale(2000ms), SelfTest::Scale(3000ms));

        state.Require(roundTrip.opened,
                      std::format(L"Navigation view disk-info keyboard activation did not enter menu mode or materialize a DxUI popup during {}.", label));
        state.Require(roundTrip.closed,
                      std::format(L"Navigation view disk-info keyboard activation did not leave menu mode or dismiss the DxUI popup during {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForClosedShellState(std::format(L"{} post-close settle", label)),
                      std::format(L"Navigation view shell state did not settle after {}.", label));
        if (! state.failure.empty())
        {
            return false;
        }
        return state.failure.empty();
    };

    state.Require(activateAndDismissMenu(VK_RETURN, L"Enter activation"),
                  L"Navigation view disk-info region did not round-trip cleanly through Enter activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(activateAndDismissMenu(VK_SPACE, L"Space activation"),
                  L"Navigation view disk-info region did not round-trip cleanly through Space activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView || GetFocus() == navigationView,
                  L"Navigation view disk-info keyboard activation should return focus to the pane shell or folder view after dismissal.");
    return state.failure.empty();
}

[[maybe_unused]] [[nodiscard]] bool TestPaneNavigationViewPointerClickDropdownsStayOpen(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view pointer-click dropdown validation.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"navigation_view_pointer_click_dropdowns_" + NewGuidText());
    const std::filesystem::path rootA = root / L"A";
    const std::filesystem::path rootB = root / L"B";
    const std::filesystem::path rootC = root / L"C";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(rootA), L"Failed to create pointer-click history path A.");
    state.Require(SelfTest::EnsureDirectory(rootB), L"Failed to create pointer-click history path B.");
    state.Require(SelfTest::EnsureDirectory(rootC), L"Failed to create pointer-click history path C.");
    state.Require(SelfTest::WriteTextFile(rootA / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in pointer-click history path A.");
    state.Require(SelfTest::WriteTextFile(rootB / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in pointer-click history path B.");
    state.Require(SelfTest::WriteTextFile(rootC / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in pointer-click history path C.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view pointer-click dropdown validation.");

    const auto navigatePane = [&](const std::filesystem::path& path, std::wstring_view context) noexcept
    {
        g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, path);
        state.Require(WaitForPanePath(FolderWindow::Pane::Left, path, SelfTest::Scale(3000ms)),
                      std::format(L"Failed to set left pane path to '{}' during {}.", path.wstring(), context));
        state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                      std::format(L"Pane contents not ready at '{}' during {}.", path.wstring(), context));
        return state.failure.empty();
    };

    state.Require(navigatePane(rootA, L"pointer-click history seeding"), L"Navigation-view pointer-click seeding failed at A.");
    state.Require(navigatePane(rootB, L"pointer-click history seeding"), L"Navigation-view pointer-click seeding failed at B.");
    state.Require(navigatePane(rootC, L"pointer-click history seeding"), L"Navigation-view pointer-click seeding failed at C.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before navigation-view pointer-click dropdown validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for navigation-view pointer-click dropdown validation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for navigation-view pointer-click dropdown validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before navigation-view pointer-click dropdown validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left));

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.showMenuSection && value.showDiskInfoSection && value.currentPathText == rootC.wstring() && value.historyCount >= 2u;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before pointer-click dropdown validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForClosedShellState = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.showMenuSection && value.showDiskInfoSection && value.currentPathText == rootC.wstring() &&
                   value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation view did not settle back to its baseline shell state during {}.", context));
        return state.failure.empty();
    };

    const auto findOwnedPopup = [&]() noexcept -> HWND
    {
        const HWND rootOwner = GetAncestor(navigationView, GA_ROOT);
        for (HWND popup = FindWindowW(L"DxUi_ContextMenu", nullptr); popup != nullptr; popup = FindWindowExW(nullptr, popup, L"DxUi_ContextMenu", nullptr))
        {
            if (IsWindowVisible(popup) == FALSE)
            {
                continue;
            }

            const HWND owner = GetWindow(popup, GW_OWNER);
            if (owner == navigationView || owner == rootOwner)
            {
                return popup;
            }
        }

        return nullptr;
    };

    const auto clickRegionAndDismiss = [&](const RECT region, const NavigationViewDebugDropdownKind expectedKind, std::wstring_view label) noexcept
    {
        const LONG clickX       = region.left + ((region.right - region.left) / 2);
        const LONG clickY       = region.top + ((region.bottom - region.top) / 2);
        const LPARAM clickPoint = MAKELPARAM(clickX, clickY);

        PostMessageW(navigationView, WM_MOUSEMOVE, 0, clickPoint);
        PostMessageW(navigationView, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        PostMessageW(navigationView, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();

        NavigationViewDebugSnapshot popupSnapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return ! value.editMode && value.historyDropdownVisible && value.dropdownKind == expectedKind && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == rootC.wstring() && g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &popupSnapshot),
                      std::format(L"Navigation view {} click did not open the expected DxUI dropdown.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        const HWND popup = WaitForWindow(findOwnedPopup, SelfTest::Scale(3000ms));
        state.Require(popup != nullptr && IsWindow(popup) != FALSE,
                      std::format(L"Navigation view {} click did not materialize a live DxUI popup window.", label));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return false;
        }

        std::this_thread::sleep_for(SelfTest::Scale(150ms));
        PumpPendingMessages();
        state.Require(IsWindowVisible(popup) != FALSE,
                      std::format(L"Navigation view {} popup dismissed immediately after the opening click/release path.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::format(L"Navigation view {} popup did not dismiss after Escape.", label));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(waitForClosedShellState(std::format(L"{} post-dismiss settle", label)),
                      std::format(L"Navigation view shell state did not settle after {} dismissal.", label));
        return state.failure.empty();
    };

    state.Require(clickRegionAndDismiss(baselineSnapshot.menuRegionRect, NavigationViewDebugDropdownKind::Menu, L"menu region"),
                  L"Navigation view menu-region click path did not keep the DxUI popup open.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(clickRegionAndDismiss(baselineSnapshot.historyRegionRect, NavigationViewDebugDropdownKind::History, L"history region"),
                  L"Navigation view history-region click path did not keep the DxUI popup open.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(clickRegionAndDismiss(baselineSnapshot.diskInfoRegionRect, NavigationViewDebugDropdownKind::DiskInfo, L"disk-info region"),
                  L"Navigation view disk-info click path did not keep the DxUI popup open.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView || GetFocus() == navigationView,
                  L"Navigation view pointer-click dropdown dismissal should return focus to the pane shell or folder view.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewHistoryDropdownEscapeReturnsFocusToFolderView(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view history-dropdown escape test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"navigation_view_history_dropdown_escape_" + NewGuidText());
    const std::filesystem::path rootA = root / L"A";
    const std::filesystem::path rootB = root / L"B";
    const std::filesystem::path rootC = root / L"C";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(rootA), L"Failed to create history escape path A.");
    state.Require(SelfTest::EnsureDirectory(rootB), L"Failed to create history escape path B.");
    state.Require(SelfTest::EnsureDirectory(rootC), L"Failed to create history escape path C.");
    state.Require(SelfTest::WriteTextFile(rootA / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history escape path A.");
    state.Require(SelfTest::WriteTextFile(rootB / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history escape path B.");
    state.Require(SelfTest::WriteTextFile(rootC / L"alpha.txt", "alpha"), L"Failed to create alpha.txt in history escape path C.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view history-dropdown escape test.");

    const auto navigatePane = [&](const std::filesystem::path& path, std::wstring_view context) noexcept
    {
        g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, path);
        state.Require(WaitForPanePath(FolderWindow::Pane::Left, path, SelfTest::Scale(3000ms)),
                      std::format(L"Failed to set left pane path to '{}' during {}.", path.wstring(), context));
        state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                      std::format(L"Pane contents not ready at '{}' during {}.", path.wstring(), context));
        return state.failure.empty();
    };

    state.Require(navigatePane(rootA, L"initial history escape seeding"), L"Navigation-view history escape seeding failed at A.");
    state.Require(navigatePane(rootB, L"initial history escape seeding"), L"Navigation-view history escape seeding failed at B.");
    state.Require(navigatePane(rootC, L"initial history escape seeding"), L"Navigation-view history escape seeding failed at C.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt before navigation-view history-dropdown escape validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for history-dropdown escape validation.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE,
                  L"Navigation view handle unavailable for history-dropdown escape validation.");
    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before history-dropdown escape validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugFocusNavigationViewRegion(FolderWindow::Pane::Left, NavigationView::FocusRegion::History),
                  L"Failed to focus the NavigationView history region before history-dropdown escape validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left));

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::HistoryRegion && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == rootC.wstring() && value.historyCount >= 2u;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the focused NavigationView history-region baseline before history-dropdown escape validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct HistoryPopupEscapeResult
    {
        bool popupObserved   = false;
        bool popupOwnerValid = false;
        bool popupClosed     = false;
    } popupResult{};

    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(3000ms));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return;
        }

        popupResult.popupObserved   = true;
        const HWND popupOwner       = GetWindow(popup, GW_OWNER);
        popupResult.popupOwnerValid = popupOwner == navigationView || popupOwner == GetAncestor(navigationView, GA_ROOT);
        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        popupResult.popupClosed = WaitForWindowClosed(popup, SelfTest::Scale(2000ms));
    });

    SendMessageW(navigationView, WM_KEYDOWN, VK_SPACE, 0);
    SendMessageW(navigationView, WM_KEYUP, VK_SPACE, 0);
    PumpPendingMessages();
    popupDriver.join();

    state.Require(popupResult.popupObserved, L"NavigationView history region did not open the live DxUI dropdown before escape validation.");
    state.Require(popupResult.popupOwnerValid, L"Focused DxUI history popup should remain owned by the active pane window hierarchy before escape validation.");
    state.Require(popupResult.popupClosed, L"DxUI history popup did not close after Escape.");
    PumpPendingMessages();

    NavigationViewDebugSnapshot closedSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == rootC.wstring() && value.historyCount == baselineSnapshot.historyCount &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount &&
               g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem;
    },
                                                SelfTest::Scale(3000ms),
                                                &closedSnapshot),
                  L"Escape should close the NavigationView history dropdown and return focus to the pane folder view without pane churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView, L"Folder view should reclaim Win32 focus after closing the NavigationView history dropdown with Escape.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneNavigationViewEditSuggestKeyboardRouting(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for navigation-view edit-suggest test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"navigation_view_edit_suggest_" + NewGuidText());
    const std::wstring alphaName     = L"alpha_suggest_target";
    const std::wstring betaName      = L"beta_suggest_other";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root / alphaName), L"Failed to create alpha_suggest_target.");
    state.Require(SelfTest::EnsureDirectory(root / betaName), L"Failed to create beta_suggest_other.");
    state.Require(SelfTest::WriteTextFile(root / L"anchor.txt", "anchor"), L"Failed to create anchor.txt for navigation-view edit-suggest test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for navigation-view edit-suggest test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for navigation-view edit-suggest test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {alphaName, betaName, L"anchor.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for navigation-view edit-suggest test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"anchor.txt"),
                  L"Failed to focus anchor.txt before edit-suggest validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND navigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Folder view handle unavailable for navigation-view edit-suggest test.");
    state.Require(navigationView != nullptr && IsWindow(navigationView) != FALSE, L"Navigation view handle unavailable for navigation-view edit-suggest test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const ScopedNavigationBarVisibilityRestore restoreNavigationBars;
    state.Require(EnsureNavigationViewVisibleForSelfTest(FolderWindow::Pane::Left, navigationView, SelfTest::Scale(3000ms)),
                  L"Left navigation bar did not become visible before navigation-view edit-suggest test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const UINT originalMainDpi = GetDpiForWindow(mainWindow);
    const UINT suggestTestDpi  = originalMainDpi >= 144u ? 192u : 144u;
    RECT originalMainRect{};
    state.Require(GetWindowRect(mainWindow, &originalMainRect) != FALSE, L"Main window rect unavailable for navigation-view edit-suggest DPI validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto scaleMainRectForDpi = [&](const UINT targetDpi) noexcept
    {
        RECT rect           = originalMainRect;
        const LONG widthPx  = std::max<LONG>(1, rect.right - rect.left);
        const LONG heightPx = std::max<LONG>(1, rect.bottom - rect.top);
        rect.right          = rect.left + MulDiv(widthPx, static_cast<int>(targetDpi), static_cast<int>(originalMainDpi));
        rect.bottom         = rect.top + MulDiv(heightPx, static_cast<int>(targetDpi), static_cast<int>(originalMainDpi));
        return rect;
    };

    const auto restoreMainWindowDpi = wil::scope_exit([&]() noexcept
    {
        SendMessageW(mainWindow,
                     WM_DPICHANGED,
                     static_cast<WPARAM>(static_cast<DWORD>(MAKELONG(static_cast<WORD>(originalMainDpi), static_cast<WORD>(originalMainDpi)))),
                     reinterpret_cast<LPARAM>(&originalMainRect));
        PumpPendingMessages();
    });

    const uint64_t baselineRefreshCount    = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount         = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount     = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const std::wstring expectedFocusedItem = L"anchor.txt";
    const std::wstring rootText            = root.wstring();
    const auto separator                   = static_cast<wchar_t>(std::filesystem::path::preferred_separator);
    const std::wstring queryOne            = rootText + separator + L"a";
    const std::wstring queryTwo            = rootText + separator + L"al";
    const std::wstring expectedAppliedText = (root / alphaName).wstring() + separator;

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return ! value.editMode && ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible &&
               ! value.fullPathPopupEditMode && value.currentPathText == rootText;
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture baseline navigation-view state for edit-suggest validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugFocusNavigationViewRegion(FolderWindow::Pane::Left, NavigationView::FocusRegion::Path),
                  L"Failed to focus the NavigationView path region before edit-suggest validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot pathRegionSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::PathRegion && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.currentPathText == rootText &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &pathRegionSnapshot),
                  L"NavigationView path region did not settle before edit-suggest validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(navigationView, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(navigationView, WM_KEYUP, VK_RETURN, 0);

    NavigationViewDebugSnapshot editSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && ! value.currentEditText.empty() &&
               ! value.historyDropdownVisible && ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode &&
               value.visibleChildWindowCount == 1u && value.currentPathText == rootText &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &editSnapshot),
                  L"Navigation view did not enter path-edit mode before edit-suggest validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND pathEdit = editSnapshot.currentEditBridgeHwnd;
    state.Require(pathEdit != nullptr && IsWindow(pathEdit) != FALSE, L"Navigation view path edit bridge unavailable for edit-suggest validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForFolderFocus = [&](std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == rootText && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                   g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedFocusedItem &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Pane folder view did not reclaim focus with stable state during {}.", context));
        return state.failure.empty();
    };

    const auto setEditTextAndMoveCaret = [&](std::wstring_view text, std::wstring_view context) noexcept
    {
        const std::wstring queryText(text);
        state.Require(queryText.starts_with(rootText),
                      std::format(L"Navigation-view edit-suggest query did not share the expected root prefix during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(SetWindowTextW(pathEdit, rootText.c_str()) != FALSE,
                      std::format(L"Navigation-view path edit did not reset to the current root path during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto rootCaret = static_cast<LPARAM>(std::min<size_t>(rootText.size(), static_cast<size_t>(std::numeric_limits<WPARAM>::max())));
        SendMessageW(pathEdit, EM_SETSEL, static_cast<WPARAM>(rootCaret), rootCaret);
        SetFocus(pathEdit);

        const std::wstring suffix = queryText.substr(rootText.size());
        for (const wchar_t ch : suffix)
        {
            SendMessageW(pathEdit, WM_CHAR, static_cast<WPARAM>(ch), 1);
        }
        PumpPendingMessages();

        const int length = GetWindowTextLengthW(pathEdit);
        state.Require(length == static_cast<int>(queryText.size()),
                      std::format(L"Navigation-view path edit did not report the expected typed query length during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        std::wstring currentText;
        currentText.resize(static_cast<size_t>(length) + 1u);
        GetWindowTextW(pathEdit, currentText.data(), static_cast<int>(currentText.size()));
        currentText.resize(wcsnlen(currentText.c_str(), currentText.size()));
        state.Require(currentText == queryText, std::format(L"Navigation-view path edit text did not match the seeded query during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto caret = static_cast<LPARAM>(std::min<size_t>(queryText.size(), static_cast<size_t>(std::numeric_limits<WPARAM>::max())));
        SendMessageW(pathEdit, EM_SETSEL, static_cast<WPARAM>(caret), caret);
        SetFocus(pathEdit);
        PumpPendingMessages();
        return true;
    };

    const auto waitForSuggestPopup = [&](std::wstring_view expectedText, int expectedSelectedIndex, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        const bool reached = WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                           [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && value.editSuggestPopupVisible &&
                   value.editSuggestItemCount >= 1u && value.editSuggestSelectedIndex == expectedSelectedIndex && ! value.historyDropdownVisible &&
                   ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 1u && value.currentPathText == rootText &&
                   value.currentEditText == expectedText && g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                           SelfTest::Scale(3000ms),
                                                           &snapshot);

        if (! reached)
        {
            NavigationViewDebugSnapshot currentSnapshot{};
            if (g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, currentSnapshot))
            {
                snapshot = currentSnapshot;
            }

            state.Require(false,
                          std::format(L"Navigation view did not expose the edit-suggest popup during {}. focusTarget={} editMode={} popupVisible={} "
                                      L"popupItems={} selectedIndex={} visibleChildren={} currentEdit='{}' currentPath='{}'",
                                      context,
                                      static_cast<int>(snapshot.focusTarget),
                                      snapshot.editMode,
                                      snapshot.editSuggestPopupVisible,
                                      snapshot.editSuggestItemCount,
                                      snapshot.editSuggestSelectedIndex,
                                      snapshot.visibleChildWindowCount,
                                      snapshot.currentEditText,
                                      snapshot.currentPathText));
        }

        return reached && state.failure.empty();
    };

    const auto waitForEditWithoutPopup = [&](std::wstring_view expectedText, std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && ! value.editSuggestPopupVisible &&
                   ! value.historyDropdownVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 1u &&
                   value.currentPathText == rootText && value.currentEditText == expectedText &&
                   g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
                   g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
                   g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation view edit state did not settle without the suggest popup during {}.", context));
        return state.failure.empty();
    };

    state.Require(setEditTextAndMoveCaret(queryOne, L"initial query"), L"Failed to seed the first edit-suggest query.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSuggestPopup(queryOne, -1, L"initial query"), L"Navigation view did not open the live suggest popup for the first query.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot suggestPopupSnapshot{};
    state.Require(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, suggestPopupSnapshot),
                  L"Navigation view snapshot unavailable after opening the live edit-suggest popup.");
    if (! state.failure.empty())
    {
        return false;
    }

    const SIZE suggestPopupClientSizeBeforeDpi = suggestPopupSnapshot.editSuggestPopupClientSize;
    const RECT suggestScaledRect               = scaleMainRectForDpi(suggestTestDpi);
    SendMessageW(mainWindow,
                 WM_DPICHANGED,
                 static_cast<WPARAM>(static_cast<DWORD>(MAKELONG(static_cast<WORD>(suggestTestDpi), static_cast<WORD>(suggestTestDpi)))),
                 reinterpret_cast<LPARAM>(&suggestScaledRect));
    PumpPendingMessages();

    NavigationViewDebugSnapshot dpiSuggestSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && value.editMode && value.editSuggestPopupVisible &&
               value.editSuggestItemCount >= 1u && value.currentEditText == queryOne && value.currentPathText == rootText && value.dpi == suggestTestDpi &&
               (value.editSuggestPopupClientSize.cx > suggestPopupClientSizeBeforeDpi.cx ||
                value.editSuggestPopupClientSize.cy > suggestPopupClientSizeBeforeDpi.cy) &&
               g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
               g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
               g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount;
    },
                                                SelfTest::Scale(3000ms),
                                                &dpiSuggestSnapshot),
                  std::format(L"Navigation view edit-suggest popup did not reflow across the synthetic DPI transition. popupVisible={} dpi={} client={}x{} "
                              L"baseline={}x{} currentEdit='{}' currentPath='{}'.",
                              dpiSuggestSnapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              dpiSuggestSnapshot.dpi,
                              dpiSuggestSnapshot.editSuggestPopupClientSize.cx,
                              dpiSuggestSnapshot.editSuggestPopupClientSize.cy,
                              suggestPopupClientSizeBeforeDpi.cx,
                              suggestPopupClientSizeBeforeDpi.cy,
                              dpiSuggestSnapshot.currentEditText,
                              dpiSuggestSnapshot.currentPathText));
    const auto hasPathEditFocus = [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        const HWND focused = GetFocus();
        return value.focusTarget == NavigationViewDebugFocusTarget::PathEdit && focused != nullptr &&
               (focused == value.currentEditBridgeHwnd || focused == value.currentEditHostHwnd ||
                (value.currentEditHostHwnd && IsChild(value.currentEditHostHwnd, focused) != FALSE));
    };

    state.Require(hasPathEditFocus(dpiSuggestSnapshot),
                  L"Path edit should retain Win32 focus while the live edit-suggest popup reflows across the synthetic DPI transition.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(pathEdit, WM_KEYDOWN, VK_DOWN, 0);
    SendMessageW(pathEdit, WM_KEYUP, VK_DOWN, 0);
    state.Require(waitForSuggestPopup(queryOne, 0, L"first suggestion selection"), L"VK_DOWN should select the first live edit-suggest row.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(pathEdit, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(pathEdit, WM_KEYUP, VK_ESCAPE, 0);
    state.Require(waitForEditWithoutPopup(queryOne, L"popup escape close"),
                  L"Escape should close the edit-suggest popup while keeping the live path edit focused.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot escapeSnapshot{};
    state.Require(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, escapeSnapshot),
                  L"Navigation view snapshot unavailable after closing the live edit-suggest popup with Escape.");
    state.Require(hasPathEditFocus(escapeSnapshot), L"Path edit should retain Win32 focus after closing the edit-suggest popup with Escape.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(setEditTextAndMoveCaret(queryTwo, L"second query"), L"Failed to seed the second edit-suggest query.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForSuggestPopup(queryTwo, -1, L"second query"), L"Navigation view did not reopen the live suggest popup for the second query.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(pathEdit, WM_KEYDOWN, VK_DOWN, 0);
    SendMessageW(pathEdit, WM_KEYUP, VK_DOWN, 0);
    state.Require(waitForSuggestPopup(queryTwo, 0, L"second suggestion selection"),
                  L"VK_DOWN should select the first live edit-suggest row for the second query.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(pathEdit, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(pathEdit, WM_KEYUP, VK_RETURN, 0);
    state.Require(waitForEditWithoutPopup(expectedAppliedText, L"applying selected suggestion"),
                  L"Enter should apply the selected live edit-suggest item without navigating the pane yet.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(expectedAppliedText.ends_with(L"\\") || expectedAppliedText.ends_with(L"/"),
                  L"Expected applied suggestion text should preserve a trailing directory separator.");
    state.Require(GetFocus() == pathEdit, L"Path edit should retain Win32 focus after applying a live edit-suggest item.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(pathEdit, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(pathEdit, WM_KEYUP, VK_ESCAPE, 0);
    state.Require(waitForFolderFocus(L"final edit escape cleanup"),
                  L"Escape from the live path edit should return focus to the pane folder view after suggestion validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(GetFocus() == folderView, L"Folder view should reclaim Win32 focus after the final edit-suggest cleanup.");
    return state.failure.empty();
}

[[nodiscard]] bool TestViewWidthAdjust(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const float ratio0 = g_folderWindow.GetSplitRatio();

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_VIEW_WIDTH, 0), 0);
    state.Require(g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth mode did not activate.");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_RIGHT));
    const float ratio1 = g_folderWindow.GetSplitRatio();
    state.Require(ratio1 > ratio0, L"ViewWidth VK_RIGHT did not increase split ratio.");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_ESCAPE));
    const float ratio2 = g_folderWindow.GetSplitRatio();
    state.Require(! g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth mode did not cancel on VK_ESCAPE.");
    state.Require(std::abs(ratio2 - ratio0) < 1e-5f, L"ViewWidth cancel did not restore split ratio.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_VIEW_WIDTH, 0), 0);
    state.Require(g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth mode did not activate (second run).");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_LEFT));
    const float ratio3 = g_folderWindow.GetSplitRatio();
    state.Require(ratio3 < ratio0, L"ViewWidth VK_LEFT did not decrease split ratio.");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_RETURN));
    state.Require(! g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth mode did not commit on VK_RETURN.");
    return state.failure.empty();
}

[[nodiscard]] bool TestViewWidthKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"view_width_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for view-width shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for view-width shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for view-width shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for view-width shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const float ratioBefore                                = g_folderWindow.GetSplitRatio();
    const auto restoreState                                = wil::scope_exit([&]
    {
        if (g_folderWindow.DebugIsViewWidthAdjustActive())
        {
            static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_ESCAPE));
        }

        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during view-width shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during view-width shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for view-width shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for view-width shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for view-width shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for view-width shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before view-width shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const size_t baselineLeftItemCount  = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before view-width shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before view-width shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  std::filesystem::path expectedPath,
                                                  size_t expectedItemCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == 0u &&
                   ! g_folderWindow.DebugIsNameFilterActive(pane);
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}.",
                                  context,
                                  static_cast<unsigned>(pane),
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(pane),
                                  g_folderWindow.DebugGetSelectedCount(pane),
                                  g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no"));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_VIEW_WIDTH, 0), 0);
    state.Require(g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth shell-stability validation did not activate ViewWidth mode.");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_RIGHT));
    const float ratioRight = g_folderWindow.GetSplitRatio();
    state.Require(ratioRight > ratioBefore, L"ViewWidth shell-stability validation expected VK_RIGHT to increase split ratio.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"View Width right adjust");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"View Width right adjust");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_ESCAPE));
    state.Require(! g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth shell-stability validation did not cancel ViewWidth mode.");
    const float ratioAfterCancel = g_folderWindow.GetSplitRatio();
    state.Require(std::abs(ratioAfterCancel - ratioBefore) < 1e-5f, L"ViewWidth shell-stability validation did not restore split ratio after cancel.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"View Width cancel");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"View Width cancel");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_VIEW_WIDTH, 0), 0);
    state.Require(g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth shell-stability validation did not reactivate ViewWidth mode.");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_LEFT));
    const float ratioLeft = g_folderWindow.GetSplitRatio();
    state.Require(ratioLeft < ratioBefore, L"ViewWidth shell-stability validation expected VK_LEFT to decrease split ratio.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"View Width left adjust");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"View Width left adjust");

    static_cast<void>(g_folderWindow.HandleViewWidthAdjustKey(VK_RETURN));
    state.Require(! g_folderWindow.DebugIsViewWidthAdjustActive(), L"ViewWidth shell-stability validation did not commit ViewWidth mode.");
    const float ratioAfterCommit = g_folderWindow.GetSplitRatio();
    state.Require(ratioAfterCommit < ratioBefore, L"ViewWidth shell-stability validation expected committed split ratio to stay narrowed.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"View Width commit");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"View Width commit");

    return state.failure.empty();
}

[[nodiscard]] bool TestAppOpenDriveMenusKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    DWORD processId        = 0;
    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, &processId);
    state.Require(uiThreadId != 0u, L"Failed to resolve the UI thread for app drive-menu shell-stability validation.");
    if (uiThreadId == 0u)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"app_open_drive_menus_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for app drive-menu shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for app drive-menu shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for app drive-menu shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for app drive-menu shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during app drive-menu shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during app drive-menu shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for app drive-menu shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for app drive-menu shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for app drive-menu shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for app drive-menu shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before app drive-menu shell-stability validation.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Right, L"right.txt"),
                  L"Failed to focus right.txt before app drive-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND leftFolderView      = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    const HWND rightFolderView     = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right);
    const HWND leftNavigationView  = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Left);
    const HWND rightNavigationView = g_folderWindow.DebugGetNavigationViewHwnd(FolderWindow::Pane::Right);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder view handle unavailable for app drive-menu shell-stability validation.");
    state.Require(rightFolderView != nullptr && IsWindow(rightFolderView) != FALSE,
                  L"Right folder view handle unavailable for app drive-menu shell-stability validation.");
    state.Require(leftNavigationView != nullptr && IsWindow(leftNavigationView) != FALSE,
                  L"Left navigation-view handle unavailable for app drive-menu shell-stability validation.");
    state.Require(rightNavigationView != nullptr && IsWindow(rightNavigationView) != FALSE,
                  L"Right navigation-view handle unavailable for app drive-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineLeftItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount     = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    const size_t baselineLeftSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);
    const size_t baselineRightSelectedCount = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.showMenuSection && value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before app drive-menu shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.showMenuSection && value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before app drive-menu shell-stability validation.");
    state.Require(baselineLeftSnapshot.showMenuSection,
                  L"Left navigation view should expose the menu region before app drive-menu shell-stability validation.");
    state.Require(baselineRightSnapshot.showMenuSection,
                  L"Right navigation view should expose the menu region before app drive-menu shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  const std::filesystem::path& expectedPath,
                                                  const size_t expectedItemCount,
                                                  const size_t expectedSelectedCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.showMenuSection && value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == expectedSelectedCount &&
                   ! g_folderWindow.DebugIsNameFilterActive(pane);
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}.",
                                  context,
                                  static_cast<unsigned>(pane),
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(pane),
                                  g_folderWindow.DebugGetSelectedCount(pane),
                                  g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no"));
    };

    const auto openAndCloseDriveMenu = [&](FolderWindow::Pane pane,
                                           UINT commandId,
                                           HWND folderView,
                                           HWND navigationView,
                                           const std::filesystem::path& expectedPath,
                                           const size_t expectedItemCount,
                                           const size_t expectedSelectedCount,
                                           const NavigationViewDebugSnapshot& baselineSnapshot,
                                           std::wstring_view expectedFocusedItem,
                                           std::wstring_view label) noexcept
    {
        FocusFolderViewPane(pane);

        const UiMenuRoundTripResult roundTrip =
            TriggerAndDismissCommandActivatedUiMenu(mainWindow, commandId, uiThreadId, navigationView, SelfTest::Scale(3000ms), SelfTest::Scale(3000ms));
        state.Require(roundTrip.opened, std::format(L"{} did not open the expected menu popup.", label));
        state.Require(roundTrip.closed, std::format(L"{} popup did not dismiss after Escape.", label));
        if (! state.failure.empty())
        {
            return;
        }

        requireStableNavigationShell(pane, expectedPath, expectedItemCount, expectedSelectedCount, baselineSnapshot, label);

        const FolderWindow::Pane otherPane = pane == FolderWindow::Pane::Left ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
        requireStableNavigationShell(otherPane,
                                     otherPane == FolderWindow::Pane::Left ? left : right,
                                     otherPane == FolderWindow::Pane::Left ? baselineLeftItemCount : baselineRightItemCount,
                                     otherPane == FolderWindow::Pane::Left ? baselineLeftSelectedCount : baselineRightSelectedCount,
                                     otherPane == FolderWindow::Pane::Left ? baselineLeftSnapshot : baselineRightSnapshot,
                                     std::format(L"{} (other pane)", label));
        state.Require(GetFocus() == folderView || GetFocus() == navigationView,
                      std::format(L"{} should return focus to the target pane shell or folder view.", label));
        state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(pane) == expectedFocusedItem,
                      std::format(L"{} should preserve the focused item on the target pane.", label));
    };

    openAndCloseDriveMenu(FolderWindow::Pane::Left,
                          IDM_LEFT_CHANGE_DRIVE,
                          leftFolderView,
                          leftNavigationView,
                          left,
                          baselineLeftItemCount,
                          baselineLeftSelectedCount,
                          baselineLeftSnapshot,
                          L"left.txt",
                          L"App Open Left Drive Menu");
    if (! state.failure.empty())
    {
        return false;
    }

    openAndCloseDriveMenu(FolderWindow::Pane::Right,
                          IDM_RIGHT_CHANGE_DRIVE,
                          rightFolderView,
                          rightNavigationView,
                          right,
                          baselineRightItemCount,
                          baselineRightSelectedCount,
                          baselineRightSnapshot,
                          L"right.txt",
                          L"App Open Right Drive Menu");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRefresh(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const uint64_t before = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_REFRESH, 0), 0);
    const uint64_t after = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    state.Require(after == before + 1u, L"Left refresh did not call FolderView::ForceRefresh.");
    return state.failure.empty();
}

[[nodiscard]] bool TestShortcutFunctionBarDispatchRefresh(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const uint64_t before = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WndMsg::kFunctionBarInvoke, VK_F9, static_cast<LPARAM>(ShortcutManager::kModCtrl));
    const uint64_t after = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);

    state.Require(after == before + 1u, L"Ctrl+F9 shortcut dispatch did not call FolderView::ForceRefresh.");
    return state.failure.empty();
}

[[nodiscard]] bool TestViewSpace(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / std::format(L"view_space_{}", GetTickCount64());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create view-space test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"nested"), L"Failed to create nested folder for view-space test.");
    state.Require(SelfTest::WriteTextFile(root / L"root.txt", "root"), L"Failed to create root file for view-space test.");
    state.Require(SelfTest::WriteTextFile(root / L"nested" / L"child.txt", "child"), L"Failed to create nested file for view-space test.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for view-space test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for view-space test.");

    FocusFolderViewPane(FolderWindow::Pane::Left);

    const size_t before = g_folderWindow.DebugGetViewerInstanceCount();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_VIEW_SPACE, 0), 0);
    const size_t after = g_folderWindow.DebugGetViewerInstanceCount();

    state.Require(after == before + 1u, L"Calculate Occupied Space did not open a viewer instance.");
    state.Require(g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-space"), L"Space Viewer instance missing after command.");

    g_folderWindow.CloseAllViewers();
    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == 0u, L"CloseAllViewers did not close all viewers.");
    return state.failure.empty();
}

[[nodiscard]] bool TestToggleUiChrome(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const bool menuBarBefore = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
    const bool menuBarAfter = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    state.Require(menuBarBefore != menuBarAfter, L"ToggleMenuBar did not change the visible DxUI menu-bar surface.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
    state.Require(DebugIsMainMenuBarSurfaceVisible(mainWindow) == menuBarBefore, L"ToggleMenuBar did not restore the visible DxUI menu-bar surface.");

    const auto toggleFunctionBarCommand = [&]() noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FUNCTIONBAR, 0), 0);
        PumpPendingMessages();
    };
    const bool funcBefore = g_folderWindow.GetFunctionBarVisible();
    const auto restoreFunctionBar = wil::scope_exit([&]() noexcept
    {
        for (int attempt = 0; attempt < 3 && g_folderWindow.GetFunctionBarVisible() != funcBefore; ++attempt)
        {
            toggleFunctionBarCommand();
        }
        if (g_folderWindow.GetFunctionBarVisible() != funcBefore)
        {
            g_folderWindow.SetFunctionBarVisible(funcBefore);
        }
    });
    for (int attempt = 0; attempt < 2 && ! g_folderWindow.GetFunctionBarVisible(); ++attempt)
    {
        toggleFunctionBarCommand();
    }
    FolderWindow::FolderWindowFunctionBarDebugSnapshot functionBarSnapshot{};
    state.Require(g_folderWindow.DebugGetFunctionBarSnapshot(functionBarSnapshot) && functionBarSnapshot.visible && functionBarSnapshot.windowVisible &&
                      functionBarSnapshot.rect.bottom > functionBarSnapshot.rect.top,
                  L"Function bar should have a visible child window and non-empty layout when enabled.");
    state.Require(functionBarSnapshot.usesDirectWriteTextMetrics,
                  L"Function bar should measure key and modifier text through DirectWrite, not the native font bridge.");
    const HWND functionBarHwnd = FindWindowExW(g_folderWindow.GetHwnd(), nullptr, L"RedSalamander.FunctionBar", nullptr);
    state.Require(functionBarHwnd != nullptr && IsWindow(functionBarHwnd) != FALSE, L"Visible Function bar HWND should be discoverable.");
    WindowColorScan functionBarScan{};
    const COLORREF functionBarBackground        = g_folderWindow.GetTheme().menu.background;
    constexpr int kFunctionBarMaterialTolerance = 20;
    const bool functionBarHasBackground = WindowContainsApproxColor(functionBarHwnd, functionBarBackground, &functionBarScan, kFunctionBarMaterialTolerance);
    state.Require(functionBarHasBackground,
                  std::format(L"Visible Function bar should paint the themed background instead of leaving a blank child strip. expected={} hwndOk={} "
                              L"visible={} gotClient={} client={}x{} topLeft={} center={} firstDifferent={}@{},{}",
                              DescribeColor(functionBarBackground),
                              functionBarScan.isWindow ? 1 : 0,
                              functionBarScan.visible ? 1 : 0,
                              functionBarScan.gotClient ? 1 : 0,
                              functionBarScan.width,
                              functionBarScan.height,
                              DescribeColor(functionBarScan.topLeft),
                              DescribeColor(functionBarScan.center),
                              DescribeColor(functionBarScan.firstDifferent),
                              functionBarScan.differentX,
                              functionBarScan.differentY));
    const COLORREF observedFunctionBarBackground = functionBarScan.center == CLR_INVALID ? functionBarBackground : functionBarScan.center;
    state.Require(WindowContainsPixelDifferentFromColor(functionBarHwnd, observedFunctionBarBackground),
                  L"Visible Function bar should draw key glyphs, separators, or labels instead of painting only an empty background.");

    toggleFunctionBarCommand();
    bool funcAfter = g_folderWindow.GetFunctionBarVisible();
    if (funcAfter)
    {
        toggleFunctionBarCommand();
        funcAfter = g_folderWindow.GetFunctionBarVisible();
    }
    state.Require(! funcAfter, L"ToggleFunctionBar did not hide the visible FolderWindow function bar.");
    toggleFunctionBarCommand();
    state.Require(g_folderWindow.GetFunctionBarVisible(), L"ToggleFunctionBar did not restore FolderWindow function bar visibility.");
    state.Require(g_folderWindow.DebugGetFunctionBarSnapshot(functionBarSnapshot) && functionBarSnapshot.windowVisible,
                  L"Restored Function bar should be visible at the HWND level.");

    const std::optional<FolderWindow::Pane> zoomBefore = g_folderWindow.GetZoomedPane();
    const std::optional<float> zoomRestoreBefore       = g_folderWindow.GetZoomRestoreSplitRatio();
    const auto restoreZoom                             = wil::scope_exit([&] { g_folderWindow.SetZoomState(zoomBefore, zoomRestoreBefore); });
    g_folderWindow.SetZoomState(std::nullopt, std::nullopt);

    FolderWindow::FolderWindowSplitterDebugSnapshot splitterSnapshot{};
    state.Require(g_folderWindow.DebugGetSplitterSnapshot(splitterSnapshot) && splitterSnapshot.leftArrowRect.bottom > splitterSnapshot.leftArrowRect.top &&
                      splitterSnapshot.rightArrowRect.bottom > splitterSnapshot.rightArrowRect.top,
                  L"Folder splitter should expose clickable navigation-height arrow zones at both ends.");
    state.Require(splitterSnapshot.leftArrowTargetPane == FolderWindow::Pane::Left && splitterSnapshot.leftArrowGlyph == L'>' &&
                      splitterSnapshot.rightArrowTargetPane == FolderWindow::Pane::Right && splitterSnapshot.rightArrowGlyph == L'<',
                  L"Restored splitter arrows should point in the direction the splitter will move to maximize the target pane.");
    state.Require(splitterSnapshot.arrowColor == splitterSnapshot.gripColor, L"Splitter arrows should use the same color as the centered grip dots.");
    state.Require(splitterSnapshot.arrowChevronSizePx > splitterSnapshot.gripDotSizePx &&
                      splitterSnapshot.arrowChevronSizePx <= (splitterSnapshot.gripDotSizePx * 2),
                  L"Splitter arrows should remain compact and close to the grip-dot scale.");
    state.Require(g_folderWindow.DebugHoverSplitterArrow(FolderWindow::Pane::Left) && g_folderWindow.DebugGetSplitterSnapshot(splitterSnapshot) &&
                      splitterSnapshot.hoveredArrowPane == FolderWindow::Pane::Left && splitterSnapshot.leftArrowCursorHand,
                  L"Left splitter arrow zone should hover and use the hand cursor.");
    state.Require(g_folderWindow.DebugClickSplitterArrow(FolderWindow::Pane::Left), L"Left splitter arrow click helper failed.");
    state.Require(g_folderWindow.GetZoomedPane() == FolderWindow::Pane::Left, L"Left splitter arrow should maximize the left pane.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left),
                  L"Left splitter arrow should focus the maximized left folder view.");
    state.Require(g_folderWindow.DebugGetSplitterSnapshot(splitterSnapshot) && splitterSnapshot.leftArrowTargetPane == FolderWindow::Pane::Right &&
                      splitterSnapshot.rightArrowTargetPane == FolderWindow::Pane::Right && splitterSnapshot.leftArrowGlyph == L'<' &&
                      splitterSnapshot.rightArrowGlyph == L'<',
                  L"When the left pane is maximized, both splitter arrows should point left to switch to the hidden right pane.");
    state.Require(g_folderWindow.DebugClickSplitterArrow(FolderWindow::Pane::Left), L"Left splitter arrow click helper failed while left pane was maximized.");
    state.Require(g_folderWindow.GetZoomedPane() == FolderWindow::Pane::Right,
                  L"When the left pane is maximized, either splitter arrow should maximize the hidden right pane instead of restoring left.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Right),
                  L"Splitter arrow should focus the newly maximized right folder view.");
    state.Require(g_folderWindow.DebugGetSplitterSnapshot(splitterSnapshot) && splitterSnapshot.leftArrowTargetPane == FolderWindow::Pane::Left &&
                      splitterSnapshot.rightArrowTargetPane == FolderWindow::Pane::Left && splitterSnapshot.leftArrowGlyph == L'>' &&
                      splitterSnapshot.rightArrowGlyph == L'>',
                  L"When the right pane is maximized, both splitter arrows should point right to switch to the hidden left pane.");
    state.Require(g_folderWindow.DebugHoverSplitterArrow(FolderWindow::Pane::Right) && g_folderWindow.DebugGetSplitterSnapshot(splitterSnapshot) &&
                      splitterSnapshot.hoveredArrowPane == FolderWindow::Pane::Right && splitterSnapshot.rightArrowCursorHand,
                  L"Right splitter arrow zone should hover and use the hand cursor.");
    state.Require(g_folderWindow.DebugClickSplitterArrow(FolderWindow::Pane::Right), L"Right splitter arrow click helper failed.");
    state.Require(g_folderWindow.GetZoomedPane() == FolderWindow::Pane::Left,
                  L"When the right pane is maximized, either splitter arrow should maximize the hidden left pane instead of restoring right.");
    state.Require(g_folderWindow.GetFocusedFolderViewHwnd() == g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left),
                  L"Splitter arrow should focus the newly maximized left folder view.");

    const bool issuesBefore = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    const bool issuesAfter = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    state.Require(issuesAfter != issuesBefore, L"ToggleFileOperationsFailedItems did not change issues pane visibility.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    state.Require(g_folderWindow.IsFileOperationsIssuesPaneVisible() == issuesBefore,
                  L"ToggleFileOperationsFailedItems did not restore issues pane visibility.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuSystemKeyShowsTemporaryDxMenuBar(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                      L"Failed to hide persistent menu bar before SC_KEYMENU activation test.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_SYSCOMMAND, SC_KEYMENU, 0);
    state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)), L"SC_KEYMENU did not show the temporary DxUI menu-bar surface.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                  L"Temporary DxUI menu-bar surface did not dismiss when focus returned to FolderView.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPersistentMainMenuBarRendersNonEmptyLabels(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (! menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                      L"Failed to show the persistent DxUI menu-bar surface for startup-render validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(WaitForMainMenuBarRenderCountAtLeast(1u, SelfTest::Scale(2000ms)), L"Persistent DxUI menu-bar surface did not render after startup/show.");

    std::wstring firstLabel;
    state.Require(DebugGetMainMenuBarItemLabel(0u, firstLabel), L"Failed to query the first top-level menu-bar label.");
    state.Require(! firstLabel.empty(), L"First top-level menu-bar label was empty.");

    RECT firstRect{};
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, 0u, firstRect), L"Failed to query the first top-level menu-bar item screen rect.");
    state.Require(firstRect.right > firstRect.left && firstRect.bottom > firstRect.top, L"First top-level menu-bar item rect was empty.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuMnemonicOpensDxContextMenu(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const std::optional<wchar_t> mnemonic = FindFirstTopLevelMenuMnemonic(mainMenu);
    state.Require(mnemonic.has_value(), L"Failed to resolve a top-level menu mnemonic.");
    if (! mnemonic.has_value())
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before mnemonic activation test.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                      L"Failed to hide persistent menu bar before mnemonic activation test.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    std::atomic<bool> popupObserved{false};
    std::atomic<bool> popupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(3000ms));
        if (! popup)
        {
            return;
        }

        popupObserved.store(true, std::memory_order_release);
        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        popupClosed.store(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    SendMessageW(mainWindow, WM_SYSCOMMAND, SC_KEYMENU, static_cast<LPARAM>(mnemonic.value()));
    popupDriver.join();
    state.Require(popupObserved.load(std::memory_order_acquire), L"Top-level menu mnemonic did not open a DxUI context-menu popup.");
    state.Require(popupClosed.load(std::memory_order_acquire), L"DxUI context-menu popup did not dismiss after Escape.");

    state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                  L"Temporary DxUI menu-bar surface did not dismiss after mnemonic popup closed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuTopLevelHighlightFollowsKeyboardOpenedRoot(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const std::optional<TopLevelMenuMapping> editMapping  = FindTopLevelMenuMappingByLabel(mainMenu, L"Edit");
    const std::optional<TopLevelMenuMapping> rightMapping = FindTopLevelMenuMappingByLabel(mainMenu, L"Right");
    state.Require(editMapping.has_value(), L"Failed to resolve the Edit top-level menu mapping.");
    state.Require(rightMapping.has_value(), L"Failed to resolve the Right top-level menu mapping.");
    if (! editMapping.has_value() || ! rightMapping.has_value())
    {
        return false;
    }

    const std::optional<wchar_t> editMnemonic = FindTopLevelMenuMnemonic(mainMenu, editMapping->rawIndex);
    state.Require(editMnemonic.has_value(), L"Failed to resolve the Edit top-level menu mnemonic.");
    if (! editMnemonic.has_value())
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before top-level highlight validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (! menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                      L"Failed to show persistent menu bar before top-level highlight validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const HWND menuBarWindow = FindMainMenuBarWindow(mainWindow);
    state.Require(menuBarWindow != nullptr, L"Failed to find the persistent DxUI main menu-bar window.");
    if (! menuBarWindow)
    {
        return false;
    }

    RECT rightItemRect{};
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, rightMapping->visualIndex, rightItemRect),
                  L"Failed to query the Right top-level menu item screen rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const POINT rightScreenPoint{(rightItemRect.left + rightItemRect.right) / 2, (rightItemRect.top + rightItemRect.bottom) / 2};
    size_t hitIndex = 0u;
    state.Require(DebugHitTestMainMenuBarScreenPoint(mainWindow, rightScreenPoint, hitIndex) && hitIndex == rightMapping->visualIndex,
                  std::format(L"Right top-level menu center did not resolve through hit-test before keyboard-opening Edit (hit={}, expected={}).",
                              hitIndex,
                              rightMapping->visualIndex));
    if (! state.failure.empty())
    {
        return false;
    }

    POINT originalCursor{};
    const bool restoreCursor = GetCursorPos(&originalCursor) != FALSE;
    const auto restoreCursorPosition = wil::scope_exit([&]() noexcept
    {
        if (restoreCursor)
        {
            SetCursorPos(originalCursor.x, originalCursor.y);
        }
    });

    state.Require(SetCursorPos(rightScreenPoint.x, rightScreenPoint.y) != FALSE, L"Failed to position the cursor over Right before keyboard-opening Edit.");
    if (! state.failure.empty())
    {
        return false;
    }

    POINT rightClientPoint = rightScreenPoint;
    state.Require(ScreenToClient(menuBarWindow, &rightClientPoint) != FALSE, L"Failed to map the Right menu-bar point to client coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(rightClientPoint.x, rightClientPoint.y));

    std::atomic<bool> popupObserved{false};
    std::atomic<bool> highlightFollowedEdit{false};
    std::atomic<int> highlightIndexWhileOpen{-2};
    std::atomic<int> selectedIndexWhileOpen{-2};
    std::atomic<bool> popupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(3000ms));
        if (! popup)
        {
            return;
        }

        popupObserved.store(true, std::memory_order_release);
        SendMessageW(menuBarWindow, WM_KILLFOCUS, reinterpret_cast<WPARAM>(popup), 0);
        const bool highlightMatched = WaitForMainMenuBarVisualHighlightIndex(editMapping->visualIndex, SelfTest::Scale(1000ms));
        highlightIndexWhileOpen.store(DebugGetMainMenuBarVisualHighlightIndex(), std::memory_order_release);
        selectedIndexWhileOpen.store(DebugGetMainMenuBarSelectedIndex(), std::memory_order_release);
        highlightFollowedEdit.store(highlightMatched, std::memory_order_release);
        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        popupClosed.store(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    SendMessageW(mainWindow, WM_SYSCOMMAND, SC_KEYMENU, static_cast<LPARAM>(editMnemonic.value()));
    popupDriver.join();

    state.Require(popupObserved.load(std::memory_order_acquire), L"Keyboard-opening Edit did not show a DxUI context-menu popup.");
    state.Require(highlightFollowedEdit.load(std::memory_order_acquire),
                  std::format(L"Keyboard-opened Edit popup did not keep the Edit top-level highlight while a stale Right hover existed (highlight={}, selected={}, "
                              L"expectedEdit={}).",
                              highlightIndexWhileOpen.load(std::memory_order_acquire),
                              selectedIndexWhileOpen.load(std::memory_order_acquire),
                              editMapping->visualIndex));
    state.Require(popupClosed.load(std::memory_order_acquire), L"Edit DxUI popup did not dismiss after Escape.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuModelRemainsPlainHMenu(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const std::optional<size_t> firstIndex = FindFirstEnabledTopLevelMenuIndex(mainMenu);
    state.Require(firstIndex.has_value(), L"Need at least one enabled top-level menu to validate the hidden HMENU contract.");
    if (! firstIndex.has_value())
    {
        return false;
    }

    const HMENU popupMenu = GetSubMenu(mainMenu, static_cast<int>(firstIndex.value()));
    state.Require(popupMenu != nullptr, L"Failed to resolve the first enabled top-level popup menu.");
    if (! popupMenu)
    {
        return false;
    }

    const std::optional<wchar_t> mnemonic = FindTopLevelMenuMnemonic(mainMenu, firstIndex.value());
    state.Require(mnemonic.has_value(), L"Failed to resolve the first enabled top-level menu mnemonic.");
    if (! mnemonic.has_value())
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before hidden HMENU contract validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                      L"Failed to hide persistent menu bar before hidden HMENU contract validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    std::atomic<bool> popupObserved{false};
    std::atomic<bool> popupClosed{false};
    std::atomic<size_t> mainOwnerDrawCount{std::numeric_limits<size_t>::max()};
    std::atomic<size_t> popupOwnerDrawCount{std::numeric_limits<size_t>::max()};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(3000ms));
        if (! popup)
        {
            return;
        }

        popupObserved.store(true, std::memory_order_release);
        mainOwnerDrawCount.store(CountOwnerDrawMenuItems(mainMenu), std::memory_order_release);
        popupOwnerDrawCount.store(CountOwnerDrawMenuItems(popupMenu), std::memory_order_release);
        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        popupClosed.store(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    SendMessageW(mainWindow, WM_SYSCOMMAND, SC_KEYMENU, static_cast<LPARAM>(mnemonic.value()));
    popupDriver.join();

    state.Require(popupObserved.load(std::memory_order_acquire), L"Top-level menu mnemonic did not open a DxUI popup for hidden HMENU contract validation.");
    state.Require(mainOwnerDrawCount.load(std::memory_order_acquire) == 0u,
                  std::format(L"Hidden top-level HMENU should remain a plain model source, but {} owner-draw items were found.",
                              mainOwnerDrawCount.load(std::memory_order_acquire)));
    state.Require(popupOwnerDrawCount.load(std::memory_order_acquire) == 0u,
                  std::format(L"Hidden popup HMENU should remain a plain model source after initialization, but {} owner-draw items were found.",
                              popupOwnerDrawCount.load(std::memory_order_acquire)));
    state.Require(popupClosed.load(std::memory_order_acquire), L"DxUI popup did not dismiss after hidden HMENU contract validation.");
    state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                  L"Temporary DxUI menu-bar surface did not dismiss after hidden HMENU contract validation.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuArrowSwitchesTopLevelPopup(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const auto topLevelIndices = FindFirstTwoEnabledTopLevelMenuIndices(mainMenu);
    state.Require(topLevelIndices.has_value(), L"Need at least two enabled top-level menus to validate arrow switching.");
    if (! topLevelIndices.has_value())
    {
        return false;
    }

    const auto [firstIndex, secondIndex]  = topLevelIndices.value();
    const std::optional<wchar_t> mnemonic = FindTopLevelMenuMnemonic(mainMenu, firstIndex);
    state.Require(mnemonic.has_value(), L"Failed to resolve the first enabled top-level menu mnemonic.");
    if (! mnemonic.has_value())
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before arrow-switch validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                      L"Failed to hide persistent menu bar before arrow-switch validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    std::atomic<bool> menuBarVisible{false};
    std::atomic<bool> firstMenuSelected{false};
    std::atomic<bool> initialPopupObserved{false};
    std::atomic<bool> secondMenuSelected{false};
    std::atomic<bool> replacementPopupObserved{false};
    std::atomic<bool> initialPopupClosed{false};
    std::atomic<bool> replacementPopupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        menuBarVisible.store(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)), std::memory_order_release);
        firstMenuSelected.store(WaitForMainMenuBarSelectedIndex(firstIndex, SelfTest::Scale(2000ms)), std::memory_order_release);

        const HWND initialPopup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(2000ms));
        if (! initialPopup)
        {
            return;
        }

        initialPopupObserved.store(true, std::memory_order_release);
        PostMessageW(initialPopup, WM_KEYDOWN, VK_RIGHT, 0);
        secondMenuSelected.store(WaitForMainMenuBarSelectedIndex(secondIndex, SelfTest::Scale(2000ms)), std::memory_order_release);

        const HWND replacementPopup = WaitForReplacementDxUiContextMenuWindow(initialPopup, SelfTest::Scale(2000ms));
        if (! replacementPopup)
        {
            PostMessageW(initialPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
            return;
        }

        replacementPopupObserved.store(true, std::memory_order_release);
        initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
        PostMessageW(replacementPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        replacementPopupClosed.store(WaitForWindowClosed(replacementPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    SendMessageW(mainWindow, WM_SYSCOMMAND, SC_KEYMENU, static_cast<LPARAM>(mnemonic.value()));
    popupDriver.join();
    state.Require(menuBarVisible.load(std::memory_order_acquire),
                  L"Mnemonic activation did not show the temporary DxUI menu-bar surface for arrow-switch validation.");
    state.Require(firstMenuSelected.load(std::memory_order_acquire), L"Mnemonic activation did not select the expected first top-level menu.");
    state.Require(initialPopupObserved.load(std::memory_order_acquire), L"Failed to open the initial DxUI context-menu popup for arrow-switch validation.");
    state.Require(secondMenuSelected.load(std::memory_order_acquire), L"VK_RIGHT did not switch the selected top-level menu while the popup was open.");
    state.Require(replacementPopupObserved.load(std::memory_order_acquire), L"Top-level menu arrow switch did not replace the active DxUI popup.");
    state.Require(initialPopupClosed.load(std::memory_order_acquire), L"Initial DxUI popup did not close after top-level arrow switching.");
    state.Require(replacementPopupClosed.load(std::memory_order_acquire), L"Replacement DxUI popup did not dismiss after Escape.");

    state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                  L"Temporary DxUI menu-bar surface did not dismiss after arrow-switch validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuHoverSwitchesTopLevelPopup(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const auto topLevelIndices = FindFirstTwoEnabledTopLevelMenuIndices(mainMenu);
    state.Require(topLevelIndices.has_value(), L"Need at least two enabled top-level menus to validate hover switching.");
    if (! topLevelIndices.has_value())
    {
        return false;
    }

    const auto [firstIndex, secondIndex] = topLevelIndices.value();

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before hover-switch validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                      L"Failed to hide persistent menu bar before hover-switch validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    SendMessageW(mainWindow, WM_SYSCOMMAND, SC_KEYMENU, 0);
    state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                  L"SC_KEYMENU did not show the temporary DxUI menu-bar surface for hover-switch validation.");

    const HWND menuBarWindow = FindMainMenuBarWindow(mainWindow);
    state.Require(menuBarWindow != nullptr, L"Failed to find the visible DxUI main menu-bar window.");
    if (! menuBarWindow)
    {
        return false;
    }

    RECT firstItemRect{};
    RECT secondItemRect{};
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, firstIndex, firstItemRect), L"Failed to query the first top-level menu item screen rect.");
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, secondIndex, secondItemRect), L"Failed to query the second top-level menu item screen rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const POINT firstScreenPoint{(firstItemRect.left + firstItemRect.right) / 2, (firstItemRect.top + firstItemRect.bottom) / 2};
    const POINT secondScreenPoint{(secondItemRect.left + secondItemRect.right) / 2, (secondItemRect.top + secondItemRect.bottom) / 2};
    POINT firstPoint = firstScreenPoint;
    state.Require(ScreenToClient(menuBarWindow, &firstPoint) != FALSE, L"Failed to map the first menu-bar item to client coordinates.");
    POINT secondPoint = secondScreenPoint;
    state.Require(ScreenToClient(menuBarWindow, &secondPoint) != FALSE, L"Failed to map the second menu-bar item to client coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> popupDriverDone{false};
    std::atomic<bool> hoverMessageDelivered{false};
    std::atomic<bool> firstMenuSelected{false};
    std::atomic<bool> initialPopupObserved{false};
    std::atomic<bool> secondMenuSelected{false};
    std::atomic<bool> singleHighlightAfterSwitch{false};
    std::atomic<int> selectedIndexAfterSwitch{-2};
    std::atomic<int> visualHighlightCountAfterSwitch{-1};
    std::atomic<bool> replacementPopupObserved{false};
    std::atomic<bool> initialPopupClosed{false};
    std::atomic<bool> replacementPopupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const auto done = wil::scope_exit([&]() noexcept { popupDriverDone.store(true, std::memory_order_release); });

        PostMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(firstPoint.x, firstPoint.y));
        PostMessageW(menuBarWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(firstPoint.x, firstPoint.y));
        PostMessageW(menuBarWindow, WM_LBUTTONUP, 0, MAKELPARAM(firstPoint.x, firstPoint.y));

        firstMenuSelected.store(WaitForMainMenuBarSelectedIndex(firstIndex, SelfTest::Scale(2000ms)), std::memory_order_release);
        const HWND initialPopup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(2000ms));
        if (! initialPopup)
        {
            return;
        }

        initialPopupObserved.store(true, std::memory_order_release);
        if (IsWindow(menuBarWindow) != FALSE)
        {
            static_cast<void>(SendMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(secondPoint.x, secondPoint.y)));
            hoverMessageDelivered.store(true, std::memory_order_release);
        }
        if (! hoverMessageDelivered.load(std::memory_order_acquire))
        {
            PostMessageW(initialPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
            return;
        }

        const HWND replacementPopup = WaitForReplacementDxUiContextMenuWindow(initialPopup, SelfTest::Scale(2000ms));
        secondMenuSelected.store(replacementPopup != nullptr || WaitForMainMenuBarSelectedIndex(secondIndex, SelfTest::Scale(250ms)),
                                 std::memory_order_release);
        if (! replacementPopup)
        {
            PostMessageW(initialPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
            return;
        }

        replacementPopupObserved.store(true, std::memory_order_release);
        const bool singleHighlight = WaitForMainMenuBarVisualHighlightCount(1, SelfTest::Scale(500ms));
        selectedIndexAfterSwitch.store(DebugGetMainMenuBarSelectedIndex(), std::memory_order_release);
        visualHighlightCountAfterSwitch.store(DebugGetMainMenuBarVisualHighlightCount(), std::memory_order_release);
        singleHighlightAfterSwitch.store(singleHighlight, std::memory_order_release);
        initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
        PostMessageW(replacementPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        replacementPopupClosed.store(WaitForWindowClosed(replacementPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    const auto driverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (! popupDriverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < driverDeadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }
    popupDriver.join();

    state.Require(firstMenuSelected.load(std::memory_order_acquire), L"Menu-bar click did not select the expected first top-level menu.");
    state.Require(initialPopupObserved.load(std::memory_order_acquire), L"Menu-bar click did not open the initial DxUI popup.");
    state.Require(hoverMessageDelivered.load(std::memory_order_acquire), L"Failed to deliver the neighboring temporary menu-bar hover message.");
    state.Require(secondMenuSelected.load(std::memory_order_acquire),
                  L"Hovering a neighboring top-level menu did not activate the replacement top-level popup.");
    state.Require(replacementPopupObserved.load(std::memory_order_acquire), L"Hover-switching did not replace the active DxUI popup.");
    state.Require(singleHighlightAfterSwitch.load(std::memory_order_acquire),
                  std::format(L"Hover-switching did not settle to a single active top-level highlight (selectedIndex={} highlightCount={}).",
                              selectedIndexAfterSwitch.load(std::memory_order_acquire),
                              visualHighlightCountAfterSwitch.load(std::memory_order_acquire)));
    state.Require(initialPopupClosed.load(std::memory_order_acquire), L"Initial DxUI popup did not close after hover-switching.");
    state.Require(replacementPopupClosed.load(std::memory_order_acquire), L"Replacement DxUI popup did not dismiss after Escape.");

    state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                  L"Temporary DxUI menu-bar surface did not dismiss after hover-switch validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuPersistentDirectHoverSwitchesTopLevelPopup(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const auto topLevelIndices = FindFirstTwoEnabledTopLevelMenuIndices(mainMenu);
    state.Require(topLevelIndices.has_value(), L"Need at least two enabled top-level menus to validate persistent direct hover switching.");
    if (! topLevelIndices.has_value())
    {
        return false;
    }

    const auto [firstIndex, secondIndex] = topLevelIndices.value();

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before persistent direct hover-switch validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (! menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                      L"Failed to show persistent menu bar before persistent direct hover-switch validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const HWND menuBarWindow = FindMainMenuBarWindow(mainWindow);
    state.Require(menuBarWindow != nullptr, L"Failed to find the persistent DxUI main menu-bar window.");
    if (! menuBarWindow)
    {
        return false;
    }

    RECT firstItemRect{};
    RECT secondItemRect{};
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, firstIndex, firstItemRect), L"Failed to query the first top-level menu item screen rect.");
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, secondIndex, secondItemRect), L"Failed to query the second top-level menu item screen rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const POINT firstScreenPoint{(firstItemRect.left + firstItemRect.right) / 2, (firstItemRect.top + firstItemRect.bottom) / 2};
    const POINT secondScreenPoint{(secondItemRect.left + secondItemRect.right) / 2, (secondItemRect.top + secondItemRect.bottom) / 2};
    POINT firstClientPoint = firstScreenPoint;
    state.Require(ScreenToClient(menuBarWindow, &firstClientPoint) != FALSE, L"Failed to map the first menu-bar item to client coordinates.");
    RECT menuBarWindowRect{};
    state.Require(GetWindowRect(menuBarWindow, &menuBarWindowRect) != FALSE, L"Failed to query the persistent DxUI main menu-bar window bounds.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG lowerHoverY =
        secondItemRect.bottom < (menuBarWindowRect.bottom - 1)
            ? secondItemRect.bottom + std::max<LONG>(1, (menuBarWindowRect.bottom - secondItemRect.bottom) / 2)
            : secondScreenPoint.y;
    const POINT secondLowerEdgeScreenPoint{secondScreenPoint.x, std::min<LONG>(lowerHoverY, menuBarWindowRect.bottom - 1)};
    POINT secondLowerEdgeClientPoint = secondLowerEdgeScreenPoint;
    state.Require(ScreenToClient(menuBarWindow, &secondLowerEdgeClientPoint) != FALSE,
                  L"Failed to map the neighboring persistent menu-bar hover point to client coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> popupDriverDone{false};
    std::atomic<bool> clickInputSent{false};
    std::atomic<bool> hoverMessageDelivered{false};
    std::atomic<bool> hoverHitResolved{false};
    std::atomic<int> hoverHitIndex{-1};
    std::atomic<bool> firstMenuSelected{false};
    std::atomic<bool> initialPopupObserved{false};
    std::atomic<bool> secondMenuSelected{false};
    std::atomic<bool> replacementPopupObserved{false};
    std::atomic<bool> initialPopupClosed{false};
    std::atomic<bool> replacementPopupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const auto done = wil::scope_exit([&]() noexcept { popupDriverDone.store(true, std::memory_order_release); });

        BringWindowToTop(mainWindow);
        SetForegroundWindow(mainWindow);

        const bool postedOpen =
            PostMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(firstClientPoint.x, firstClientPoint.y)) != FALSE &&
            PostMessageW(menuBarWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(firstClientPoint.x, firstClientPoint.y)) != FALSE &&
            PostMessageW(menuBarWindow, WM_LBUTTONUP, 0, MAKELPARAM(firstClientPoint.x, firstClientPoint.y)) != FALSE;
        clickInputSent.store(postedOpen, std::memory_order_release);
        if (! clickInputSent.load(std::memory_order_acquire))
        {
            return;
        }

        firstMenuSelected.store(WaitForMainMenuBarSelectedIndex(firstIndex, SelfTest::Scale(2000ms)), std::memory_order_release);
        const HWND initialPopup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(2000ms));
        if (! initialPopup)
        {
            return;
        }

        initialPopupObserved.store(true, std::memory_order_release);

        size_t hitIndex = 0u;
        if (DebugHitTestMainMenuBarScreenPoint(mainWindow, secondLowerEdgeScreenPoint, hitIndex))
        {
            hoverHitResolved.store(true, std::memory_order_release);
            hoverHitIndex.store(static_cast<int>(hitIndex), std::memory_order_release);
        }

        if (IsWindow(menuBarWindow) != FALSE)
        {
            static_cast<void>(SendMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(secondLowerEdgeClientPoint.x, secondLowerEdgeClientPoint.y)));
            hoverMessageDelivered.store(true, std::memory_order_release);
        }
        if (! hoverMessageDelivered.load(std::memory_order_acquire))
        {
            PostMessageW(initialPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
            return;
        }

        const HWND replacementPopup = WaitForReplacementDxUiContextMenuWindow(initialPopup, SelfTest::Scale(2000ms));
        secondMenuSelected.store(replacementPopup != nullptr || WaitForMainMenuBarSelectedIndex(secondIndex, SelfTest::Scale(250ms)),
                                 std::memory_order_release);
        if (! replacementPopup)
        {
            PostMessageW(initialPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
            return;
        }

        replacementPopupObserved.store(true, std::memory_order_release);
        initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
        PostMessageW(replacementPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        replacementPopupClosed.store(WaitForWindowClosed(replacementPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    const auto driverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (! popupDriverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < driverDeadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }
    popupDriver.join();

    state.Require(clickInputSent.load(std::memory_order_acquire), L"Failed to post a mouse click on the persistent menu bar.");
    state.Require(hoverMessageDelivered.load(std::memory_order_acquire), L"Failed to deliver the neighboring persistent menu-bar hover message.");
    state.Require(hoverHitResolved.load(std::memory_order_acquire) && hoverHitIndex.load(std::memory_order_acquire) == static_cast<int>(secondIndex),
                  std::format(L"Neighboring persistent menu item lower-edge hover did not resolve through the menu-bar hit-test; resolved={} hitIndex={} expected={} point=({},{}).",
                              hoverHitResolved.load(std::memory_order_acquire),
                              hoverHitIndex.load(std::memory_order_acquire),
                              secondIndex,
                              secondLowerEdgeScreenPoint.x,
                              secondLowerEdgeScreenPoint.y));
    state.Require(firstMenuSelected.load(std::memory_order_acquire), L"Persistent menu-bar click did not select the expected first top-level menu.");
    state.Require(initialPopupObserved.load(std::memory_order_acquire), L"Persistent menu-bar click did not open the initial DxUI popup.");
    state.Require(secondMenuSelected.load(std::memory_order_acquire),
                  L"Directly hovering a neighboring persistent top-level menu did not select the replacement top-level menu.");
    state.Require(replacementPopupObserved.load(std::memory_order_acquire),
                  L"Direct persistent top-level hover did not replace the active DxUI popup before entering the first popup item.");
    state.Require(initialPopupClosed.load(std::memory_order_acquire), L"Initial DxUI popup did not close after persistent direct hover-switching.");
    state.Require(replacementPopupClosed.load(std::memory_order_acquire), L"Replacement DxUI popup did not dismiss after Escape.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuPersistentViewToPluginsHoverSwitchesPopup(HWND mainWindow, CaseState& state, bool hoverViaMouseLeaveRefresh) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before View-to-Plugins hover-switch validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (! menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                      L"Failed to show persistent menu bar before View-to-Plugins hover-switch validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::optional<TopLevelMenuMapping> pluginsMapping = FindTopLevelMenuMappingByLabel(mainMenu, L"Plugins");
    const std::optional<TopLevelMenuMapping> viewMapping    = FindTopLevelMenuMappingByLabel(mainMenu, L"View");
    state.Require(pluginsMapping.has_value(), L"Failed to resolve the Plugins top-level menu mapping.");
    state.Require(viewMapping.has_value(), L"Failed to resolve the View top-level menu mapping.");
    if (! pluginsMapping.has_value() || ! viewMapping.has_value())
    {
        return false;
    }

    state.Require(viewMapping->visualIndex > pluginsMapping->visualIndex,
                  std::format(L"Expected View to appear after Plugins for backward hover-switch validation (Plugins visual={}, View visual={}).",
                              pluginsMapping->visualIndex,
                              viewMapping->visualIndex));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND menuBarWindow = FindMainMenuBarWindow(mainWindow);
    state.Require(menuBarWindow != nullptr, L"Failed to find the persistent DxUI main menu-bar window.");
    if (! menuBarWindow)
    {
        return false;
    }

    RECT pluginsItemRect{};
    RECT viewItemRect{};
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, pluginsMapping->visualIndex, pluginsItemRect),
                  L"Failed to query the Plugins top-level menu item screen rect.");
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, viewMapping->visualIndex, viewItemRect),
                  L"Failed to query the View top-level menu item screen rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const POINT viewScreenPoint{(viewItemRect.left + viewItemRect.right) / 2, (viewItemRect.top + viewItemRect.bottom) / 2};
    const POINT pluginsScreenPoint{(pluginsItemRect.left + pluginsItemRect.right) / 2, (pluginsItemRect.top + pluginsItemRect.bottom) / 2};
    POINT viewClientPoint = viewScreenPoint;
    state.Require(ScreenToClient(menuBarWindow, &viewClientPoint) != FALSE, L"Failed to map the View menu-bar item to client coordinates.");

    size_t hitIndex = 0u;
    state.Require(DebugHitTestMainMenuBarScreenPoint(mainWindow, pluginsScreenPoint, hitIndex) && hitIndex == pluginsMapping->visualIndex,
                  std::format(L"Plugins menu-bar center should resolve through hit-test before opening View (hit={}, expected={}).",
                              hitIndex,
                              pluginsMapping->visualIndex));
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> popupDriverDone{false};
    std::atomic<bool> clickInputSent{false};
    std::atomic<bool> hoverMessageDelivered{false};
    std::atomic<bool> hoverCursorMoved{! hoverViaMouseLeaveRefresh};
    std::atomic<bool> viewMenuSelected{false};
    std::atomic<bool> initialPopupObserved{false};
    std::atomic<bool> pluginsMenuSelected{false};
    std::atomic<bool> replacementPopupObserved{false};
    std::atomic<bool> initialPopupClosed{false};
    std::atomic<bool> replacementPopupClosed{false};
    std::atomic<bool> hoverHitResolvedAfterMove{false};
    std::atomic<int> hoverHitIndexAfterMove{-1};
    std::atomic<int> cursorAfterMoveX{0};
    std::atomic<int> cursorAfterMoveY{0};
    std::atomic<int> selectedIndexAfterHoverWait{-2};
    std::atomic<unsigned long long> rootPointerSwitchCount{0};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const auto done = wil::scope_exit([&]() noexcept { popupDriverDone.store(true, std::memory_order_release); });

        BringWindowToTop(mainWindow);
        SetForegroundWindow(mainWindow);

        const bool postedOpen =
            PostMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(viewClientPoint.x, viewClientPoint.y)) != FALSE &&
            PostMessageW(menuBarWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(viewClientPoint.x, viewClientPoint.y)) != FALSE &&
            PostMessageW(menuBarWindow, WM_LBUTTONUP, 0, MAKELPARAM(viewClientPoint.x, viewClientPoint.y)) != FALSE;
        clickInputSent.store(postedOpen, std::memory_order_release);
        if (! clickInputSent.load(std::memory_order_acquire))
        {
            return;
        }

        viewMenuSelected.store(WaitForMainMenuBarSelectedIndex(viewMapping->visualIndex, SelfTest::Scale(2000ms)), std::memory_order_release);
        const HWND initialPopup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(2000ms));
        if (! initialPopup)
        {
            return;
        }

        initialPopupObserved.store(true, std::memory_order_release);

        RECT livePluginsItemRect{};
        POINT livePluginsScreenPoint = pluginsScreenPoint;
        if (DebugGetMainMenuBarItemScreenRect(mainWindow, pluginsMapping->visualIndex, livePluginsItemRect))
        {
            livePluginsScreenPoint = POINT{(livePluginsItemRect.left + livePluginsItemRect.right) / 2, (livePluginsItemRect.top + livePluginsItemRect.bottom) / 2};
        }

        cursorAfterMoveX.store(livePluginsScreenPoint.x, std::memory_order_release);
        cursorAfterMoveY.store(livePluginsScreenPoint.y, std::memory_order_release);
        size_t hoverHitIndex = 0u;
        if (DebugHitTestMainMenuBarScreenPoint(mainWindow, livePluginsScreenPoint, hoverHitIndex))
        {
            hoverHitResolvedAfterMove.store(true, std::memory_order_release);
            hoverHitIndexAfterMove.store(static_cast<int>(hoverHitIndex), std::memory_order_release);
        }

        if (hoverViaMouseLeaveRefresh)
        {
            POINT originalCursor{};
            const bool restoreCursor = GetCursorPos(&originalCursor) != FALSE;
            const auto restoreCursorPosition = wil::scope_exit([&]() noexcept
            {
                if (restoreCursor)
                {
                    SetCursorPos(originalCursor.x, originalCursor.y);
                }
            });

            hoverCursorMoved.store(SetCursorPos(livePluginsScreenPoint.x, livePluginsScreenPoint.y) != FALSE, std::memory_order_release);
            if (hoverCursorMoved.load(std::memory_order_acquire) && IsWindow(menuBarWindow) != FALSE)
            {
                static_cast<void>(SendMessageW(menuBarWindow, WM_MOUSELEAVE, 0, 0));
                hoverMessageDelivered.store(true, std::memory_order_release);
            }
        }
        else
        {
            POINT livePluginsClientPoint = livePluginsScreenPoint;
            if (ScreenToClient(menuBarWindow, &livePluginsClientPoint) != FALSE && IsWindow(menuBarWindow) != FALSE)
            {
                static_cast<void>(SendMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(livePluginsClientPoint.x, livePluginsClientPoint.y)));
                hoverMessageDelivered.store(true, std::memory_order_release);
            }
        }
        if (! hoverMessageDelivered.load(std::memory_order_acquire))
        {
            PostMessageW(initialPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
            return;
        }

        const HWND replacementPopup = WaitForReplacementDxUiContextMenuWindow(initialPopup, SelfTest::Scale(2000ms));
        RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
        if (RedSalamander::DxUi::DebugGetContextMenuPopupState(initialPopup, popupState))
        {
            rootPointerSwitchCount.store(popupState.rootPointerSwitchCount, std::memory_order_release);
        }
        selectedIndexAfterHoverWait.store(DebugGetMainMenuBarSelectedIndex(), std::memory_order_release);
        pluginsMenuSelected.store(replacementPopup != nullptr || WaitForMainMenuBarSelectedIndex(pluginsMapping->visualIndex, SelfTest::Scale(250ms)),
                                  std::memory_order_release);
        if (! replacementPopup)
        {
            PostMessageW(initialPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
            return;
        }

        replacementPopupObserved.store(true, std::memory_order_release);
        initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
        PostMessageW(replacementPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        replacementPopupClosed.store(WaitForWindowClosed(replacementPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    const auto driverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (! popupDriverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < driverDeadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }
    popupDriver.join();

    state.Require(clickInputSent.load(std::memory_order_acquire), L"Failed to post a mouse click on the View top-level menu.");
    state.Require(viewMenuSelected.load(std::memory_order_acquire), L"Clicking View did not select the View top-level menu.");
    state.Require(initialPopupObserved.load(std::memory_order_acquire), L"Clicking View did not open the initial DxUI popup.");
    state.Require(hoverCursorMoved.load(std::memory_order_acquire), L"Failed to move the real cursor from View to Plugins.");
    state.Require(hoverMessageDelivered.load(std::memory_order_acquire),
                  L"Failed to deliver the captured-session hover message to the persistent DxUI menu bar.");
    state.Require(pluginsMenuSelected.load(std::memory_order_acquire),
                  std::format(L"Delivering a Plugins hover while View is open did not select the Plugins top-level menu (mode={}, point=({},{}), hitResolved={}, "
                              L"hitIndex={}, selectedAfterWait={}, expectedPluginsIndex={}, pointerSwitches={}).",
                              hoverViaMouseLeaveRefresh ? L"mouse-leave-refresh" : L"mouse-move",
                              cursorAfterMoveX.load(std::memory_order_acquire),
                              cursorAfterMoveY.load(std::memory_order_acquire),
                              hoverHitResolvedAfterMove.load(std::memory_order_acquire) ? L"yes" : L"no",
                              hoverHitIndexAfterMove.load(std::memory_order_acquire),
                              selectedIndexAfterHoverWait.load(std::memory_order_acquire),
                              pluginsMapping->visualIndex,
                              rootPointerSwitchCount.load(std::memory_order_acquire)));
    state.Require(replacementPopupObserved.load(std::memory_order_acquire),
                  L"Hovering Plugins while View is open did not replace the active DxUI popup.");
    state.Require(initialPopupClosed.load(std::memory_order_acquire), L"View DxUI popup did not close after switching to Plugins.");
    state.Require(replacementPopupClosed.load(std::memory_order_acquire), L"Plugins DxUI popup did not dismiss after Escape.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuPersistentViewToFilesHoverHighlightFollowsPointer(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before View-to-Files highlight validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (! menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                      L"Failed to show persistent menu bar before View-to-Files highlight validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::optional<TopLevelMenuMapping> filesMapping = FindTopLevelMenuMappingByLabel(mainMenu, L"Files");
    const std::optional<TopLevelMenuMapping> viewMapping  = FindTopLevelMenuMappingByLabel(mainMenu, L"View");
    state.Require(filesMapping.has_value(), L"Failed to resolve the Files top-level menu mapping.");
    state.Require(viewMapping.has_value(), L"Failed to resolve the View top-level menu mapping.");
    if (! filesMapping.has_value() || ! viewMapping.has_value())
    {
        return false;
    }

    const HWND menuBarWindow = FindMainMenuBarWindow(mainWindow);
    state.Require(menuBarWindow != nullptr, L"Failed to find the persistent DxUI main menu-bar window.");
    if (! menuBarWindow)
    {
        return false;
    }

    RECT filesItemRect{};
    RECT viewItemRect{};
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, filesMapping->visualIndex, filesItemRect),
                  L"Failed to query the Files top-level menu item screen rect.");
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, viewMapping->visualIndex, viewItemRect),
                  L"Failed to query the View top-level menu item screen rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const POINT viewScreenPoint{(viewItemRect.left + viewItemRect.right) / 2, (viewItemRect.top + viewItemRect.bottom) / 2};
    const POINT filesScreenPoint{(filesItemRect.left + filesItemRect.right) / 2, (filesItemRect.top + filesItemRect.bottom) / 2};
    POINT viewClientPoint = viewScreenPoint;
    POINT filesClientPoint = filesScreenPoint;
    state.Require(ScreenToClient(menuBarWindow, &viewClientPoint) != FALSE, L"Failed to map the View menu-bar item to client coordinates.");
    state.Require(ScreenToClient(menuBarWindow, &filesClientPoint) != FALSE, L"Failed to map the Files menu-bar item to client coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    POINT originalCursor{};
    const bool restoreCursor = GetCursorPos(&originalCursor) != FALSE;
    const auto restoreCursorPosition = wil::scope_exit([&]() noexcept
    {
        if (restoreCursor)
        {
            SetCursorPos(originalCursor.x, originalCursor.y);
        }
    });

    std::atomic<bool> popupDriverDone{false};
    std::atomic<bool> clickInputSent{false};
    std::atomic<bool> viewMenuSelected{false};
    std::atomic<bool> initialPopupObserved{false};
    std::atomic<bool> hoverMessageDelivered{false};
    std::atomic<int> highlightIndexAfterHover{-2};
    std::atomic<int> selectedIndexAfterHover{-2};
    std::atomic<bool> filesMenuHighlighted{false};
    std::atomic<bool> replacementPopupObserved{false};
    std::atomic<bool> initialPopupClosed{false};
    std::atomic<bool> replacementPopupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const auto done = wil::scope_exit([&]() noexcept { popupDriverDone.store(true, std::memory_order_release); });

        BringWindowToTop(mainWindow);
        SetForegroundWindow(mainWindow);

        const bool postedOpen =
            PostMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(viewClientPoint.x, viewClientPoint.y)) != FALSE &&
            PostMessageW(menuBarWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(viewClientPoint.x, viewClientPoint.y)) != FALSE &&
            PostMessageW(menuBarWindow, WM_LBUTTONUP, 0, MAKELPARAM(viewClientPoint.x, viewClientPoint.y)) != FALSE;
        clickInputSent.store(postedOpen, std::memory_order_release);
        if (! clickInputSent.load(std::memory_order_acquire))
        {
            return;
        }

        viewMenuSelected.store(WaitForMainMenuBarSelectedIndex(viewMapping->visualIndex, SelfTest::Scale(2000ms)), std::memory_order_release);
        const HWND initialPopup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(2000ms));
        if (! initialPopup)
        {
            return;
        }

        initialPopupObserved.store(true, std::memory_order_release);
        static_cast<void>(SetCursorPos(filesScreenPoint.x, filesScreenPoint.y));
        if (IsWindow(menuBarWindow) != FALSE)
        {
            static_cast<void>(SendMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(filesClientPoint.x, filesClientPoint.y)));
            hoverMessageDelivered.store(true, std::memory_order_release);
        }

        highlightIndexAfterHover.store(DebugGetMainMenuBarVisualHighlightIndex(), std::memory_order_release);
        selectedIndexAfterHover.store(DebugGetMainMenuBarSelectedIndex(), std::memory_order_release);
        filesMenuHighlighted.store(highlightIndexAfterHover.load(std::memory_order_acquire) == static_cast<int>(filesMapping->visualIndex),
                                   std::memory_order_release);

        const HWND replacementPopup = WaitForReplacementDxUiContextMenuWindow(initialPopup, SelfTest::Scale(2000ms));
        if (! replacementPopup)
        {
            PostMessageW(initialPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
            return;
        }

        replacementPopupObserved.store(true, std::memory_order_release);
        initialPopupClosed.store(WaitForWindowClosed(initialPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
        PostMessageW(replacementPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        replacementPopupClosed.store(WaitForWindowClosed(replacementPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    const auto driverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (! popupDriverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < driverDeadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }
    popupDriver.join();

    state.Require(clickInputSent.load(std::memory_order_acquire), L"Failed to post a mouse click on the View top-level menu.");
    state.Require(viewMenuSelected.load(std::memory_order_acquire), L"Clicking View did not select the View top-level menu.");
    state.Require(initialPopupObserved.load(std::memory_order_acquire), L"Clicking View did not open the initial DxUI popup.");
    state.Require(hoverMessageDelivered.load(std::memory_order_acquire), L"Failed to deliver the Files hover message while the View popup was open.");
    state.Require(filesMenuHighlighted.load(std::memory_order_acquire),
                  std::format(L"Moving from View to Files left the top-level highlight on the previous menu (highlight={}, selected={}, expectedFiles={}).",
                              highlightIndexAfterHover.load(std::memory_order_acquire),
                              selectedIndexAfterHover.load(std::memory_order_acquire),
                              filesMapping->visualIndex));
    state.Require(replacementPopupObserved.load(std::memory_order_acquire), L"Hovering Files while View is open did not replace the active DxUI popup.");
    state.Require(initialPopupClosed.load(std::memory_order_acquire), L"View DxUI popup did not close after switching to Files.");
    state.Require(replacementPopupClosed.load(std::memory_order_acquire), L"Files DxUI popup did not dismiss after Escape.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuMouseOpenKeepsPopupSelectionClear(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before mouse-open selection validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (! menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                      L"Failed to show the persistent DxUI menu bar before mouse-open selection validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::optional<TopLevelMenuMapping> mapping = FindPaneTopLevelMenuMapping(mainMenu);
    state.Require(mapping.has_value(), L"Need the Left or Right top-level pane menu to validate mouse-open popup selection.");
    if (! mapping.has_value())
    {
        return false;
    }

    const HMENU popupMenu     = GetSubMenu(mainMenu, static_cast<int>(mapping->rawIndex));
    const int nativeItemCount = popupMenu ? GetMenuItemCount(popupMenu) : 0;
    state.Require(popupMenu != nullptr && nativeItemCount > 0, L"Pane top-level menu did not expose a non-empty popup menu.");
    if (! popupMenu || nativeItemCount <= 0)
    {
        return false;
    }

    const HWND menuBarWindow = FindMainMenuBarWindow(mainWindow);
    state.Require(menuBarWindow != nullptr && IsWindow(menuBarWindow) != FALSE,
                  L"Failed to find the visible DxUI main menu-bar window for mouse-open selection validation.");
    if (! menuBarWindow || IsWindow(menuBarWindow) == FALSE)
    {
        return false;
    }

    static_cast<void>(g_folderWindow.TryRestoreActivePaneFolderViewFocus());
    const HWND expectedFolderView = g_folderWindow.GetFocusedFolderViewHwnd();
    state.Require(expectedFolderView != nullptr,
                  L"Need active pane folder-view keyboard focus before mouse-opening the persistent DxUI menu bar.");
    if (! expectedFolderView)
    {
        return false;
    }

    SetFocus(menuBarWindow);
    state.Require(GetFocus() == menuBarWindow, L"Need keyboard focus on the persistent DxUI menu bar before mouse-opening the menu.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT itemRect{};
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, mapping->visualIndex, itemRect),
                  std::format(L"Failed to query the screen rect for top-level menu '{}'.", mapping->label));
    if (! state.failure.empty())
    {
        return false;
    }

    POINT clickPoint{itemRect.left + ((itemRect.right - itemRect.left) / 2), itemRect.top + ((itemRect.bottom - itemRect.top) / 2)};
    state.Require(ScreenToClient(menuBarWindow, &clickPoint) != FALSE, L"Failed to map the menu-bar click point to client coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> popupDriverDone{false};
    std::atomic<bool> popupObserved{false};
    std::atomic<bool> popupStateReadable{false};
    std::atomic<int> hoveredIndex{-2};
    std::atomic<int> keyboardIndex{-2};
    std::atomic<bool> itemPaintReadable{false};
    std::atomic<int> highlightedItemCount{-1};
    std::atomic<bool> popupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const auto done = wil::scope_exit([&]() noexcept { popupDriverDone.store(true, std::memory_order_release); });

        PostMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(clickPoint.x, clickPoint.y));
        PostMessageW(menuBarWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickPoint.x, clickPoint.y));
        PostMessageW(menuBarWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickPoint.x, clickPoint.y));

        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(2000ms));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return;
        }

        popupObserved.store(true, std::memory_order_release);

        RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
        if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState))
        {
            hoveredIndex.store(popupState.hoveredIndex.has_value() ? static_cast<int>(popupState.hoveredIndex.value()) : -1, std::memory_order_release);
            keyboardIndex.store(popupState.keyboardIndex.has_value() ? static_cast<int>(popupState.keyboardIndex.value()) : -1, std::memory_order_release);
            popupStateReadable.store(true, std::memory_order_release);
        }

        int readablePaintCount = 0;
        int highlightCount     = 0;
        for (int itemIndex = 0; itemIndex < nativeItemCount; ++itemIndex)
        {
            RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState paintState{};
            if (! RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popup, static_cast<size_t>(itemIndex), paintState))
            {
                continue;
            }

            ++readablePaintCount;
            if (paintState.usesHighlightFill)
            {
                ++highlightCount;
            }
        }

        itemPaintReadable.store(readablePaintCount > 0, std::memory_order_release);
        highlightedItemCount.store(highlightCount, std::memory_order_release);

        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        popupClosed.store(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    const auto driverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (! popupDriverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < driverDeadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }
    popupDriver.join();

    state.Require(popupObserved.load(std::memory_order_acquire), L"Mouse click on the pane top-level menu did not open a DxUI popup.");
    state.Require(popupStateReadable.load(std::memory_order_acquire), L"Mouse-opened DxUI popup did not expose debug state.");
    state.Require(
        keyboardIndex.load(std::memory_order_acquire) == -1,
        std::format(L"Mouse-opened popup should not synthesize a keyboard selection, but reported index {}.", keyboardIndex.load(std::memory_order_acquire)));
    state.Require(
        hoveredIndex.load(std::memory_order_acquire) == -1,
        std::format(L"Stationary mouse on the menu bar should not hover a popup item, but reported index {}.", hoveredIndex.load(std::memory_order_acquire)));
    state.Require(itemPaintReadable.load(std::memory_order_acquire), L"Mouse-opened DxUI popup did not expose item paint debug state.");
    state.Require(
        highlightedItemCount.load(std::memory_order_acquire) == 0,
        std::format(L"Mouse-opened popup should not draw a selected item fill until pointer movement or keyboard navigation, but {} item(s) were highlighted.",
                    highlightedItemCount.load(std::memory_order_acquire)));
    state.Require(popupClosed.load(std::memory_order_acquire), L"Mouse-opened DxUI popup did not dismiss after Escape.");
    const bool firstFocusRestored = WaitForFocusedFolderViewForMainMenu(expectedFolderView, SelfTest::Scale(2000ms));
    state.Require(firstFocusRestored,
                  std::format(L"Mouse-opened main menu popup should restore keyboard focus to the active pane folder view after closing; expected={}, actual={}.",
                              DescribeWindowHandleForSelfTest(expectedFolderView),
                              DescribeWindowHandleForSelfTest(GetFocus())));
    state.Require(WaitForMainMenuBarSelectedIndex(std::nullopt, SelfTest::Scale(2000ms)),
                  L"Mouse-opened main menu popup should clear the selected top-level menu after closing.");
    state.Require(WaitForMainMenuBarVisualHighlightCount(0, SelfTest::Scale(2000ms)),
                  L"Mouse-opened main menu popup should clear the menu-bar visual highlight after closing.");
    if (! state.failure.empty())
    {
        return false;
    }

    SetFocus(menuBarWindow);
    state.Require(GetFocus() == menuBarWindow, L"Need keyboard focus on the persistent DxUI menu bar before the second mouse-open pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> secondPopupDriverDone{false};
    std::atomic<bool> secondPopupObserved{false};
    std::atomic<bool> secondPopupClosed{false};
    std::jthread secondPopupDriver([&](std::stop_token) noexcept
    {
        const auto done = wil::scope_exit([&]() noexcept { secondPopupDriverDone.store(true, std::memory_order_release); });

        PostMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(clickPoint.x, clickPoint.y));
        PostMessageW(menuBarWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickPoint.x, clickPoint.y));
        PostMessageW(menuBarWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickPoint.x, clickPoint.y));

        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(2000ms));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return;
        }

        secondPopupObserved.store(true, std::memory_order_release);
        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        secondPopupClosed.store(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    const auto secondDriverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (! secondPopupDriverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < secondDriverDeadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }
    secondPopupDriver.join();

    state.Require(secondPopupObserved.load(std::memory_order_acquire), L"Second mouse click on the pane top-level menu did not open a DxUI popup.");
    state.Require(secondPopupClosed.load(std::memory_order_acquire), L"Second mouse-opened DxUI popup did not dismiss after Escape.");
    const bool secondFocusRestored = WaitForFocusedFolderViewForMainMenu(expectedFolderView, SelfTest::Scale(2000ms));
    state.Require(secondFocusRestored,
                  std::format(
                      L"Mouse-opened main menu popup should restore keyboard focus even when the menu bar owned focus before opening; expected={}, actual={}.",
                      DescribeWindowHandleForSelfTest(expectedFolderView),
                      DescribeWindowHandleForSelfTest(GetFocus())));
    state.Require(WaitForMainMenuBarSelectedIndex(std::nullopt, SelfTest::Scale(2000ms)),
                  L"Second mouse-opened main menu popup should clear the selected top-level menu after closing.");
    state.Require(WaitForMainMenuBarVisualHighlightCount(0, SelfTest::Scale(2000ms)),
                  L"Second mouse-opened main menu popup should clear the menu-bar visual highlight after closing.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuMouseOpenedPopupProcessesKeyboardBeforeMouseMove(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before mouse-open keyboard validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (! menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                      L"Failed to show the persistent DxUI menu bar before mouse-open keyboard validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const std::optional<TopLevelMenuMapping> mapping = FindPaneTopLevelMenuMapping(mainMenu);
    state.Require(mapping.has_value(), L"Need the Left or Right top-level pane menu to validate mouse-open keyboard routing.");
    if (! mapping.has_value())
    {
        return false;
    }

    const HWND menuBarWindow = FindMainMenuBarWindow(mainWindow);
    state.Require(menuBarWindow != nullptr && IsWindow(menuBarWindow) != FALSE,
                  L"Failed to find the visible DxUI main menu-bar window for mouse-open keyboard validation.");
    if (! menuBarWindow || IsWindow(menuBarWindow) == FALSE)
    {
        return false;
    }

    RECT itemRect{};
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, mapping->visualIndex, itemRect),
                  std::format(L"Failed to query the screen rect for top-level menu '{}'.", mapping->label));
    if (! state.failure.empty())
    {
        return false;
    }

    POINT clickPoint{itemRect.left + ((itemRect.right - itemRect.left) / 2), itemRect.top + ((itemRect.bottom - itemRect.top) / 2)};
    state.Require(ScreenToClient(menuBarWindow, &clickPoint) != FALSE, L"Failed to map the menu-bar click point to client coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> popupDriverDone{false};
    std::atomic<bool> clickInputSent{false};
    std::atomic<bool> popupObserved{false};
    std::atomic<bool> popupInitialStateReadable{false};
    std::atomic<int> initialKeyboardIndex{-2};
    std::atomic<unsigned long long> initialRenderCount{0};
    std::atomic<unsigned long long> popupHandle{0};
    std::atomic<bool> popupHadKeyboardFocusBeforeKey{false};
    std::atomic<bool> keyMessagePosted{false};
    std::atomic<unsigned long long> focusBeforeKey{0};
    std::atomic<unsigned long long> activeBeforeKey{0};
    std::atomic<bool> keyboardAppliedBeforeMouseMove{false};
    std::atomic<int> keyboardIndexAfterKey{-2};
    std::atomic<unsigned long long> renderCountAfterKey{0};
    std::atomic<bool> popupClosed{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const auto done = wil::scope_exit([&]() noexcept { popupDriverDone.store(true, std::memory_order_release); });

        BringWindowToTop(mainWindow);
        SetForegroundWindow(mainWindow);
        SetFocus(menuBarWindow);

        const bool postedOpen =
            PostMessageW(menuBarWindow, WM_MOUSEMOVE, 0, MAKELPARAM(clickPoint.x, clickPoint.y)) != FALSE &&
            PostMessageW(menuBarWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickPoint.x, clickPoint.y)) != FALSE &&
            PostMessageW(menuBarWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickPoint.x, clickPoint.y)) != FALSE;
        clickInputSent.store(postedOpen, std::memory_order_release);
        if (! clickInputSent.load(std::memory_order_acquire))
        {
            return;
        }

        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(2000ms));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return;
        }

        popupObserved.store(true, std::memory_order_release);
        popupHandle.store(reinterpret_cast<uintptr_t>(popup), std::memory_order_release);

        RedSalamander::DxUi::ContextMenuPopupDebugState initialPopupState{};
        if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, initialPopupState))
        {
            initialKeyboardIndex.store(initialPopupState.keyboardIndex.has_value() ? static_cast<int>(initialPopupState.keyboardIndex.value()) : -1,
                                       std::memory_order_release);
            initialRenderCount.store(initialPopupState.renderCount, std::memory_order_release);
            popupInitialStateReadable.store(true, std::memory_order_release);
        }

        const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
        GUITHREADINFO guiInfo{};
        guiInfo.cbSize = sizeof(guiInfo);
        if (uiThreadId != 0 && GetGUIThreadInfo(uiThreadId, &guiInfo) != FALSE)
        {
            focusBeforeKey.store(reinterpret_cast<uintptr_t>(guiInfo.hwndFocus), std::memory_order_release);
            activeBeforeKey.store(reinterpret_cast<uintptr_t>(guiInfo.hwndActive), std::memory_order_release);
        }

        popupHadKeyboardFocusBeforeKey.store(guiInfo.hwndFocus == popup, std::memory_order_release);
        const HWND keyTarget = guiInfo.hwndFocus ? guiInfo.hwndFocus : popup;
        keyMessagePosted.store(PostMessageW(keyTarget, WM_KEYDOWN, VK_DOWN, 0) != FALSE && PostMessageW(keyTarget, WM_KEYUP, VK_DOWN, 0) != FALSE,
                               std::memory_order_release);

        const auto keyDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1000ms);
        while (std::chrono::steady_clock::now() < keyDeadline)
        {
            RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
            if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState))
            {
                keyboardIndexAfterKey.store(popupState.keyboardIndex.has_value() ? static_cast<int>(popupState.keyboardIndex.value()) : -1,
                                            std::memory_order_release);
                renderCountAfterKey.store(popupState.renderCount, std::memory_order_release);
                if (popupState.keyboardIndex.has_value() && popupState.renderCount > initialRenderCount.load(std::memory_order_acquire))
                {
                    keyboardAppliedBeforeMouseMove.store(true, std::memory_order_release);
                    break;
                }
            }

            std::this_thread::sleep_for(10ms);
        }

        PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        popupClosed.store(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    const auto driverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (! popupDriverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < driverDeadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }
    popupDriver.join();

    state.Require(clickInputSent.load(std::memory_order_acquire), L"Failed to post a mouse click on the top-level menu.");
    state.Require(popupObserved.load(std::memory_order_acquire), L"Mouse click on the top-level menu did not open a DxUI popup.");
    state.Require(popupInitialStateReadable.load(std::memory_order_acquire), L"Mouse-opened DxUI popup did not expose initial debug state.");
    state.Require(initialKeyboardIndex.load(std::memory_order_acquire) == -1,
                  std::format(L"Mouse-opened popup should start without a keyboard item, but reported index {}.",
                              initialKeyboardIndex.load(std::memory_order_acquire)));
    state.Require(keyMessagePosted.load(std::memory_order_acquire), L"Failed to post VK_DOWN to the focused menu target.");
    state.Require(keyboardAppliedBeforeMouseMove.load(std::memory_order_acquire),
                  std::format(L"Mouse-opened popup did not process and repaint VK_DOWN before mouse movement (keyboardIndex={}, renderBefore={}, renderAfter={}, "
                              L"popup=0x{:X}, focusBeforeKey=0x{:X}, activeBeforeKey=0x{:X}, popupHadFocus={}).",
                              keyboardIndexAfterKey.load(std::memory_order_acquire),
                              initialRenderCount.load(std::memory_order_acquire),
                              renderCountAfterKey.load(std::memory_order_acquire),
                              popupHandle.load(std::memory_order_acquire),
                              focusBeforeKey.load(std::memory_order_acquire),
                              activeBeforeKey.load(std::memory_order_acquire),
                              popupHadKeyboardFocusBeforeKey.load(std::memory_order_acquire) ? L"yes" : L"no"));
    state.Require(popupClosed.load(std::memory_order_acquire), L"Mouse-opened DxUI popup did not dismiss after keyboard validation.");

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuTopLevelMappingMatchesRawMenu(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const std::vector<TopLevelMenuMapping> mappings = CollectEnabledTopLevelMenuMappings(mainMenu);
    state.Require(mappings.size() >= 2u, L"Need at least two enabled top-level menus to validate raw-to-visual mapping.");
    if (mappings.size() < 2u)
    {
        return false;
    }

    for (const TopLevelMenuMapping& mapping : mappings)
    {
        size_t actualSourceIndex = 0u;
        size_t hitTestIndex      = 0u;
        std::wstring actualLabel;
        state.Require(DebugGetMainMenuBarItemSourceIndex(mapping.visualIndex, actualSourceIndex),
                      std::format(L"Failed to query the raw source index for visual top-level menu '{}'.", mapping.label));
        state.Require(DebugGetMainMenuBarItemLabel(mapping.visualIndex, actualLabel),
                      std::format(L"Failed to query the retained label for visual top-level menu index {}.", mapping.visualIndex));
        RECT itemRectPx{};
        state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, mapping.visualIndex, itemRectPx),
                      std::format(L"Failed to query the screen rect for visual top-level menu '{}'.", mapping.label));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(actualSourceIndex == mapping.rawIndex,
                      std::format(L"Top-level menu '{}' should retain raw menu index {} but reported {}.", mapping.label, mapping.rawIndex, actualSourceIndex));
        state.Require(
            actualLabel == mapping.label,
            std::format(L"Top-level menu label mismatch for visual index {}: expected '{}' but got '{}'.", mapping.visualIndex, mapping.label, actualLabel));

        const POINT centerPoint{itemRectPx.left + ((itemRectPx.right - itemRectPx.left) / 2), itemRectPx.top + ((itemRectPx.bottom - itemRectPx.top) / 2)};
        state.Require(DebugHitTestMainMenuBarScreenPoint(mainWindow, centerPoint, hitTestIndex),
                      std::format(L"Failed to hit-test the screen center of top-level menu '{}'.", mapping.label));
        state.Require(
            hitTestIndex == mapping.visualIndex,
            std::format(
                L"Screen-point hit-testing for top-level menu '{}' returned visual index {} instead of {}.", mapping.label, hitTestIndex, mapping.visualIndex));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuRightJustifiedItemsAnchorToTrailingEdge(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const std::optional<TopLevelMenuMapping> mapping = FindFirstRightJustifiedTopLevelMenuMapping(mainMenu);
    state.Require(mapping.has_value(), L"Need a right-justified top-level menu item to validate trailing-edge layout.");
    if (! mapping.has_value())
    {
        return false;
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (! menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, true, SelfTest::Scale(2000ms)),
                      L"Failed to show the persistent DxUI menu bar before right-justified layout validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const HWND menuBarWindow = FindMainMenuBarWindow(mainWindow);
    state.Require(menuBarWindow != nullptr && IsWindow(menuBarWindow) != FALSE,
                  L"Failed to find the visible DxUI main menu-bar window for right-justified layout validation.");
    if (! menuBarWindow || IsWindow(menuBarWindow) == FALSE)
    {
        return false;
    }

    RECT menuBarRect{};
    RECT itemRect{};
    std::wstring label;
    state.Require(GetWindowRect(menuBarWindow, &menuBarRect) != FALSE,
                  L"Failed to query the visible DxUI menu-bar bounds for right-justified layout validation.");
    state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, mapping->visualIndex, itemRect),
                  std::format(L"Failed to query the right-justified top-level menu item rect for '{}'.", mapping->label));
    state.Require(DebugGetMainMenuBarItemLabel(mapping->visualIndex, label),
                  L"Failed to query the retained label for the right-justified top-level menu item.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int trailingGapPx = menuBarRect.right - itemRect.right;
    state.Require(trailingGapPx >= 0 && trailingGapPx <= 24,
                  std::format(L"Right-justified menu '{}' should anchor to the trailing edge; trailingGapPx={}.", label, trailingGapPx));
    state.Require(label == mapping->label, std::format(L"Right-justified top-level menu label mismatch: expected '{}' but saw '{}'.", mapping->label, label));
    if (mapping->visualIndex > 0u)
    {
        RECT previousRect{};
        state.Require(DebugGetMainMenuBarItemScreenRect(mainWindow, mapping->visualIndex - 1u, previousRect),
                      L"Failed to query the menu item immediately before the right-justified entry.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(previousRect.right <= itemRect.left, std::format(L"Right-justified menu '{}' should not overlap the preceding top-level item.", label));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuPopupHostMatchesFlyoutContract(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const std::optional<wchar_t> mnemonic = FindFirstTopLevelMenuMnemonic(mainMenu);
    state.Require(mnemonic.has_value(), L"Failed to resolve a top-level menu mnemonic.");
    if (! mnemonic.has_value())
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before popup host validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                      L"Failed to hide persistent menu bar before popup host validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    std::atomic<bool> popupObserved{false};
    std::atomic<bool> popupNotTopmost{false};
    std::atomic<bool> popupOwnedByMainWindow{false};
    std::atomic<bool> popupRegionClipped{false};
    std::atomic<bool> popupClosedOnDeactivate{false};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND popup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(3000ms));
        if (! popup || IsWindow(popup) == FALSE)
        {
            return;
        }

        popupObserved.store(true, std::memory_order_release);

        const LONG_PTR exStyle = GetWindowLongPtrW(popup, GWL_EXSTYLE);
        popupNotTopmost.store((exStyle & WS_EX_TOPMOST) == 0, std::memory_order_release);
        popupOwnedByMainWindow.store(GetWindow(popup, GW_OWNER) == mainWindow, std::memory_order_release);

        wil::unique_hrgn region(CreateRectRgn(0, 0, 0, 0));
        const int regionType = region ? GetWindowRgn(popup, region.get()) : ERROR;
        popupRegionClipped.store((regionType == SIMPLEREGION || regionType == COMPLEXREGION) && region && PtInRegion(region.get(), 0, 0) == FALSE,
                                 std::memory_order_release);

        PostMessageW(mainWindow, WM_ACTIVATEAPP, FALSE, 0);
        popupClosedOnDeactivate.store(WaitForWindowClosed(popup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    SendMessageW(mainWindow, WM_SYSCOMMAND, SC_KEYMENU, static_cast<LPARAM>(mnemonic.value()));
    popupDriver.join();

    PostMessageW(mainWindow, WM_ACTIVATEAPP, TRUE, 0);
    static_cast<void>(SetForegroundWindow(mainWindow));
    static_cast<void>(SetFocus(mainWindow));

    state.Require(popupObserved.load(std::memory_order_acquire), L"Top-level menu mnemonic did not open a DxUI context-menu popup.");
    state.Require(popupNotTopmost.load(std::memory_order_acquire), L"DxUI menu popup should be a transient owned flyout, not a topmost window.");
    state.Require(popupOwnedByMainWindow.load(std::memory_order_acquire), L"DxUI menu popup should remain owned by the main application window.");
    state.Require(popupRegionClipped.load(std::memory_order_acquire), L"DxUI menu popup should clip the host window corners to the rounded overlay shape.");
    state.Require(popupClosedOnDeactivate.load(std::memory_order_acquire), L"DxUI menu popup did not dismiss after simulated app deactivation.");

    state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                  L"Temporary DxUI menu-bar surface did not dismiss after popup host deactivation validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestMainMenuSubmenuPlacementMatchesSpec(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const HMENU mainMenu = DebugGetMainMenuModelHandle();
    state.Require(mainMenu != nullptr, L"Main menu handle not available.");
    if (! mainMenu)
    {
        return false;
    }

    const auto cascadePath = FindFirstCascadeMenuPath(mainMenu);
    state.Require(cascadePath.has_value(), L"Need a top-level menu item with a cascading submenu to validate submenu placement.");
    if (! cascadePath.has_value())
    {
        return false;
    }

    const auto [topLevelIndex, childIndex] = cascadePath.value();
    const HMENU popupMenu                  = GetSubMenu(mainMenu, static_cast<int>(topLevelIndex));
    state.Require(popupMenu != nullptr, L"Failed to resolve the top-level popup menu for submenu placement validation.");
    const std::optional<wchar_t> mnemonic = FindTopLevelMenuMnemonic(mainMenu, topLevelIndex);
    state.Require(mnemonic.has_value(), L"Failed to resolve the top-level mnemonic for submenu placement validation.");
    if (! popupMenu || ! mnemonic.has_value())
    {
        return false;
    }

    if (const HWND existingPopup = FindVisibleDxUiContextMenuWindow(); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Failed to dismiss pre-existing DxUI context menu before submenu placement validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const bool menuBarInitiallyVisible = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    const auto restoreMenuBar          = wil::scope_exit([&]
    {
        const bool visibleNow = DebugIsMainMenuBarSurfaceVisible(mainWindow);
        if (visibleNow != menuBarInitiallyVisible)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
            static_cast<void>(WaitForMainMenuBarVisibility(mainWindow, menuBarInitiallyVisible, SelfTest::Scale(2000ms)));
        }
    });

    if (menuBarInitiallyVisible)
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
        state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                      L"Failed to hide persistent menu bar before submenu placement validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    std::atomic<bool> rootPopupObserved{false};
    std::atomic<bool> submenuObserved{false};
    std::atomic<bool> submenuStayedOpenOnParentReturn{false};
    std::atomic<bool> submenuClosed{false};
    std::atomic<bool> rootPopupClosed{false};
    std::atomic<int> horizontalDeltaPx{-1};
    std::atomic<int> verticalDeltaPx{-1};
    std::atomic<int> rootKeyboardIndex{-2};
    std::atomic<int> rootHoveredIndex{-2};
    std::atomic<int> rootScrollOffsetPx{-1};
    std::atomic<int> parentTopPx{-1};
    std::atomic<int> parentBottomPx{-1};
    std::atomic<int> expectedSubmenuTopPx{-1};
    std::atomic<int> actualSubmenuTopPx{-1};
    std::atomic<bool> opensToRight{true};
    std::atomic<int> placementTolerancePx{0};
    std::jthread popupDriver([&](std::stop_token) noexcept
    {
        const HWND rootPopup = WaitForWindow([]() noexcept { return FindVisibleDxUiContextMenuWindow(); }, SelfTest::Scale(3000ms));
        if (! rootPopup || IsWindow(rootPopup) == FALSE)
        {
            return;
        }

        rootPopupObserved.store(true, std::memory_order_release);

        const size_t navigationBudget = static_cast<size_t>((std::max)(GetMenuItemCount(popupMenu), 1)) * 2u;
        for (size_t attempt = 0; attempt < navigationBudget; ++attempt)
        {
            if (WaitForContextMenuKeyboardIndex(rootPopup, childIndex, SelfTest::Scale(100ms)))
            {
                break;
            }

            PostMessageW(rootPopup, WM_KEYDOWN, VK_DOWN, 0);
            PostMessageW(rootPopup, WM_KEYUP, VK_DOWN, 0);
        }

        if (! WaitForContextMenuKeyboardIndex(rootPopup, childIndex, SelfTest::Scale(500ms)))
        {
            PostMessageW(rootPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(rootPopup, SelfTest::Scale(2000ms)));
            return;
        }

        RedSalamander::DxUi::ContextMenuPopupDebugState rootPopupState{};
        D2D1_RECT_F parentItemRectDip{};
        if (! RedSalamander::DxUi::DebugGetContextMenuPopupState(rootPopup, rootPopupState) ||
            ! RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(rootPopup, childIndex, parentItemRectDip))
        {
            PostMessageW(rootPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(rootPopup, SelfTest::Scale(2000ms)));
            return;
        }

        RECT rootPopupRect{};
        GetWindowRect(rootPopup, &rootPopupRect);
        const float scale = static_cast<float>(rootPopupState.dpi) / 96.0f;
        rootKeyboardIndex.store(rootPopupState.keyboardIndex.has_value() ? static_cast<int>(rootPopupState.keyboardIndex.value()) : -1,
                                std::memory_order_release);
        rootHoveredIndex.store(rootPopupState.hoveredIndex.has_value() ? static_cast<int>(rootPopupState.hoveredIndex.value()) : -1, std::memory_order_release);
        rootScrollOffsetPx.store(static_cast<int>(rootPopupState.scrollOffsetDip * scale + 0.5f), std::memory_order_release);
        const RECT parentItemRectPx{rootPopupRect.left + static_cast<LONG>(parentItemRectDip.left * scale + 0.5f),
                                    rootPopupRect.top + static_cast<LONG>(parentItemRectDip.top * scale + 0.5f),
                                    rootPopupRect.left + static_cast<LONG>(parentItemRectDip.right * scale + 0.5f),
                                    rootPopupRect.top + static_cast<LONG>(parentItemRectDip.bottom * scale + 0.5f)};
        parentTopPx.store(parentItemRectPx.top, std::memory_order_release);
        parentBottomPx.store(parentItemRectPx.bottom, std::memory_order_release);

        PostMessageW(rootPopup, WM_KEYDOWN, VK_RIGHT, 0);
        PostMessageW(rootPopup, WM_KEYUP, VK_RIGHT, 0);

        std::vector<HWND> popupWindows;
        if (! WaitForOwnedDxUiContextMenuWindowCount(mainWindow, 2u, SelfTest::Scale(2000ms), &popupWindows))
        {
            PostMessageW(rootPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(rootPopup, SelfTest::Scale(2000ms)));
            return;
        }

        HWND submenuPopup = nullptr;
        for (const HWND candidate : popupWindows)
        {
            if (candidate != rootPopup)
            {
                submenuPopup = candidate;
                break;
            }
        }

        if (! submenuPopup || IsWindow(submenuPopup) == FALSE)
        {
            PostMessageW(rootPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(rootPopup, SelfTest::Scale(2000ms)));
            return;
        }

        submenuObserved.store(true, std::memory_order_release);

        RECT submenuRect{};
        GetWindowRect(submenuPopup, &submenuRect);

        RedSalamander::DxUi::ContextMenuPopupDebugState submenuPopupState{};
        if (! RedSalamander::DxUi::DebugGetContextMenuPopupState(submenuPopup, submenuPopupState))
        {
            PostMessageW(submenuPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(submenuPopup, SelfTest::Scale(2000ms)));
            PostMessageW(rootPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(rootPopup, SelfTest::Scale(2000ms)));
            return;
        }

        const float submenuScale = static_cast<float>(submenuPopupState.dpi) / 96.0f;
        const RECT submenuSurfaceRectPx    = submenuPopupState.surfaceRectPx;

        RECT expectedSubmenuSurfaceRectPx{};
        const POINT submenuAnchor{parentItemRectPx.right, parentItemRectPx.top};
        if (! RedSalamander::DxUi::DebugComputeContextMenuPopupPosition(submenuAnchor,
                                                                        submenuPopupState.visibleWidthDip,
                                                                        submenuPopupState.visibleHeightDip,
                                                                        submenuPopupState.dpi,
                                                                        true,
                                                                        &rootPopupRect,
                                                                        &parentItemRectPx,
                                                                        RedSalamander::DxUi::ContextMenuRootHorizontalAlignment::Start,
                                                                        RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Below,
                                                                        expectedSubmenuSurfaceRectPx))
        {
            PostMessageW(submenuPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(submenuPopup, SelfTest::Scale(2000ms)));
            PostMessageW(rootPopup, WM_KEYDOWN, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(rootPopup, SelfTest::Scale(2000ms)));
            return;
        }

        const bool submenuOpensToRight = submenuSurfaceRectPx.left >= parentItemRectPx.left;
        const int submenuHorizontalDeltaPx =
            submenuOpensToRight ? std::abs(submenuSurfaceRectPx.left - parentItemRectPx.right) : std::abs(submenuSurfaceRectPx.right - parentItemRectPx.left);
        const int submenuVerticalDeltaPx = std::abs(submenuSurfaceRectPx.top - expectedSubmenuSurfaceRectPx.top);

        expectedSubmenuTopPx.store(expectedSubmenuSurfaceRectPx.top, std::memory_order_release);
        actualSubmenuTopPx.store(submenuSurfaceRectPx.top, std::memory_order_release);
        opensToRight.store(submenuOpensToRight, std::memory_order_release);
        horizontalDeltaPx.store(submenuHorizontalDeltaPx, std::memory_order_release);
        verticalDeltaPx.store(submenuVerticalDeltaPx, std::memory_order_release);
        placementTolerancePx.store((std::max)(6, static_cast<int>(4.0f * submenuScale + 0.5f)), std::memory_order_release);

        const LONG submenuClientX = static_cast<LONG>(((submenuSurfaceRectPx.left + submenuSurfaceRectPx.right) / 2) - submenuRect.left);
        const LONG submenuClientY = static_cast<LONG>(((submenuSurfaceRectPx.top + submenuSurfaceRectPx.bottom) / 2) - submenuRect.top);
        PostMessageW(submenuPopup, WM_MOUSEMOVE, 0, MAKELPARAM(submenuClientX, submenuClientY));
        std::this_thread::sleep_for(SelfTest::Scale(50ms));

        const LONG parentClientX = static_cast<LONG>(((parentItemRectPx.left + parentItemRectPx.right) / 2) - rootPopupRect.left);
        const LONG parentClientY = static_cast<LONG>(((parentItemRectPx.top + parentItemRectPx.bottom) / 2) - rootPopupRect.top);
        PostMessageW(rootPopup, WM_MOUSEMOVE, 0, MAKELPARAM(parentClientX, parentClientY));
        std::this_thread::sleep_for(SelfTest::Scale(650ms));

        std::vector<HWND> parentReturnPopupWindows;
        submenuStayedOpenOnParentReturn.store(WaitForOwnedDxUiContextMenuWindowCount(mainWindow, 2u, SelfTest::Scale(100ms), &parentReturnPopupWindows),
                                              std::memory_order_release);

        PostMessageW(submenuPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        submenuClosed.store(WaitForWindowClosed(submenuPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
        PostMessageW(rootPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        rootPopupClosed.store(WaitForWindowClosed(rootPopup, SelfTest::Scale(2000ms)), std::memory_order_release);
    });

    SendMessageW(mainWindow, WM_SYSCOMMAND, SC_KEYMENU, static_cast<LPARAM>(mnemonic.value()));
    popupDriver.join();

    state.Require(rootPopupObserved.load(std::memory_order_acquire), L"Top-level menu mnemonic did not open the root DxUI menu popup.");
    state.Require(submenuObserved.load(std::memory_order_acquire), L"Cascading submenu did not open a second DxUI popup window.");
    state.Require(horizontalDeltaPx.load(std::memory_order_acquire) >= 0, L"Failed to measure the submenu attachment geometry.");
    state.Require(verticalDeltaPx.load(std::memory_order_acquire) >= 0, L"Failed to measure the submenu vertical placement.");
    state.Require(horizontalDeltaPx.load(std::memory_order_acquire) <= placementTolerancePx.load(std::memory_order_acquire),
                  std::format(L"Cascading submenu should attach flush to the parent item edge (delta={} px, opensToRight={}).",
                              horizontalDeltaPx.load(std::memory_order_acquire),
                              opensToRight.load(std::memory_order_acquire) ? L"yes" : L"no"));
    state.Require(verticalDeltaPx.load(std::memory_order_acquire) <= placementTolerancePx.load(std::memory_order_acquire),
                  std::format(L"Cascading submenu should honor the 4 DIP vertical offset with work-area clamping (delta={} px, tolerance={} px, "
                              L"parentTop={}, parentBottom={}, expectedTop={}, actualTop={}, rootKeyboardIndex={}, rootHoveredIndex={}, rootScroll={} px).",
                              verticalDeltaPx.load(std::memory_order_acquire),
                              placementTolerancePx.load(std::memory_order_acquire),
                              parentTopPx.load(std::memory_order_acquire),
                              parentBottomPx.load(std::memory_order_acquire),
                              expectedSubmenuTopPx.load(std::memory_order_acquire),
                              actualSubmenuTopPx.load(std::memory_order_acquire),
                              rootKeyboardIndex.load(std::memory_order_acquire),
                              rootHoveredIndex.load(std::memory_order_acquire),
                              rootScrollOffsetPx.load(std::memory_order_acquire)));
    state.Require(submenuStayedOpenOnParentReturn.load(std::memory_order_acquire),
                  L"Cascading submenu should remain open when the pointer returns to the parent item that owns it.");
    state.Require(submenuClosed.load(std::memory_order_acquire), L"Cascading submenu did not close after Escape.");
    state.Require(rootPopupClosed.load(std::memory_order_acquire), L"Root DxUI popup did not close after submenu placement validation.");

    state.Require(WaitForMainMenuBarVisibility(mainWindow, false, SelfTest::Scale(2000ms)),
                  L"Temporary DxUI menu-bar surface did not dismiss after submenu placement validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestAppPreferencesKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"app_preferences_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for app Preferences shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for app Preferences shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for app Preferences shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for app Preferences shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Preferences window did not close before app Preferences shell-stability validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during app Preferences shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during app Preferences shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for app Preferences shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for app Preferences shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for app Preferences shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for app Preferences shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before app Preferences shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder-view handle unavailable for app Preferences shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD processId               = GetCurrentProcessId();
    const auto baselineWindows          = SnapshotTopLevelWindowsForProcess(processId);
    const size_t baselineLeftItemCount  = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before app Preferences shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before app Preferences shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  const std::filesystem::path& expectedPath,
                                                  const size_t expectedItemCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  const bool requireFocusedFolderView,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        const bool stable = WaitForNavigationViewSnapshot(pane,
                                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == 0u &&
                    ! g_folderWindow.DebugIsNameFilterActive(pane) &&
                    (! requireFocusedFolderView || g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView);
        },
                                                          SelfTest::Scale(3000ms),
                                                          &snapshot);
        state.Require(
            stable,
            std::format(
                L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, popupVisible={}, "
                L"childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}, focusedFolderViewMatches={}, "
                L"focusedWindow={}.",
                context,
                static_cast<unsigned>(pane),
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetItemCount(pane),
                g_folderWindow.DebugGetSelectedCount(pane),
                g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no",
                g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView ? L"yes" : L"no",
                DescribeWindowHandleForSelfTest(GetFocus())));
    };
    const auto waitForPreferencesSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(10ms);
        }
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };
    const auto waitForLeftFolderFocus = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView)
            {
                return true;
            }
            std::this_thread::sleep_for(10ms);
        }
        return g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView;
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open during app Preferences shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    PreferencesDebugSnapshot prefsSnapshot{};
    state.Require(
        waitForPreferencesSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.shellUsesDxUiHost && value.categoryTreeUsesDxUiHost && value.pageHostUsesDxUiHost && value.visibleLegacyTreeViewCount == 0u &&
               value.shellDxHostResizeFailureCount == 0u;
    },
            prefsSnapshot),
        std::format(
            L"Preferences window did not stabilize before close during app Preferences shell-stability validation; shellUsesDxUiHost={}, "
            L"categoryTreeUsesDxUiHost={}, pageHostUsesDxUiHost={}, visibleLegacyTreeViewCount={}, shellDxHostResizeFailureCount={}, currentCategory={}.",
            prefsSnapshot.shellUsesDxUiHost ? L"yes" : L"no",
            prefsSnapshot.categoryTreeUsesDxUiHost ? L"yes" : L"no",
            prefsSnapshot.pageHostUsesDxUiHost ? L"yes" : L"no",
            prefsSnapshot.visibleLegacyTreeViewCount,
            prefsSnapshot.shellDxHostResizeFailureCount,
            static_cast<unsigned>(prefsSnapshot.currentCategory)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! IsOwnedBy(prefs, mainWindow),
                  L"Preferences window should remain an independent top-level window during app Preferences shell-stability validation.");
    PostMessageW(prefs, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)), L"Preferences window did not close during app Preferences shell-stability validation.");
    state.Require(WaitForNoNonBaselineWindows(processId, baselineWindows, mainWindow, SelfTest::Scale(3000ms)),
                  L"App Preferences should not leave non-baseline top-level windows behind after close.");
    static_cast<void>(SetForegroundWindow(mainWindow));
    static_cast<void>(SetFocus(mainWindow));
    state.Require(waitForLeftFolderFocus(), L"App Preferences should restore folder-view focus to the active pane after close.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, true, L"App Preferences close");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, false, L"App Preferences close");

    return state.failure.empty();
}

[[nodiscard]] bool TestAppManagePluginsKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"app_manage_plugins_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for app Manage Plugins shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for app Manage Plugins shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for app Manage Plugins shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for app Manage Plugins shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Preferences window did not close before app Manage Plugins shell-stability validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during app Manage Plugins shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during app Manage Plugins shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for app Manage Plugins shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for app Manage Plugins shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for app Manage Plugins shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for app Manage Plugins shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before app Manage Plugins shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder-view handle unavailable for app Manage Plugins shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD processId               = GetCurrentProcessId();
    const auto baselineWindows          = SnapshotTopLevelWindowsForProcess(processId);
    const size_t baselineLeftItemCount  = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before app Manage Plugins shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before app Manage Plugins shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  const std::filesystem::path& expectedPath,
                                                  const size_t expectedItemCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  const bool requireFocusedFolderView,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(
            WaitForNavigationViewSnapshot(pane,
                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == 0u &&
                   ! g_folderWindow.DebugIsNameFilterActive(pane) &&
                   (! requireFocusedFolderView || g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView);
        },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
            std::format(
                L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, popupVisible={}, "
                L"childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}, focusedFolderViewMatches={}.",
                context,
                static_cast<unsigned>(pane),
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetItemCount(pane),
                g_folderWindow.DebugGetSelectedCount(pane),
                g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no",
                g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView ? L"yes" : L"no"));
    };
    const auto waitForPreferencesSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(10ms);
        }
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };
    const auto waitForLeftFolderFocus = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView)
            {
                return true;
            }
            std::this_thread::sleep_for(10ms);
        }
        return g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView;
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PLUGINS_MANAGE, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Manage Plugins should open the Preferences window during app shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    PreferencesDebugSnapshot prefsSnapshot{};
    state.Require(
        waitForPreferencesSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.shellUsesDxUiHost && value.categoryTreeUsesDxUiHost && value.pageHostUsesDxUiHost && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsPaneVisible && ! value.generalPaneVisible && value.pluginsMainListRowCount > 0u && value.visibleLegacyTreeViewCount == 0u &&
               value.shellDxHostResizeFailureCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
            prefsSnapshot),
        std::format(L"Manage Plugins did not stabilize on the Preferences Plugins page before close; shellUsesDxUiHost={}, categoryTreeUsesDxUiHost={}, "
                    L"pageHostUsesDxUiHost={}, currentCategory={}, pluginsPaneVisible={}, generalPaneVisible={}, pluginsMainListRowCount={}, "
                    L"visibleLegacyTreeViewCount={}, shellDxHostResizeFailureCount={}, currentPageDxHostResizeFailureCount={}.",
                    prefsSnapshot.shellUsesDxUiHost ? L"yes" : L"no",
                    prefsSnapshot.categoryTreeUsesDxUiHost ? L"yes" : L"no",
                    prefsSnapshot.pageHostUsesDxUiHost ? L"yes" : L"no",
                    static_cast<unsigned>(prefsSnapshot.currentCategory),
                    prefsSnapshot.pluginsPaneVisible ? L"yes" : L"no",
                    prefsSnapshot.generalPaneVisible ? L"yes" : L"no",
                    prefsSnapshot.pluginsMainListRowCount,
                    prefsSnapshot.visibleLegacyTreeViewCount,
                    prefsSnapshot.shellDxHostResizeFailureCount,
                    prefsSnapshot.currentPageDxHostResizeFailureCount));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! IsOwnedBy(prefs, mainWindow),
                  L"Manage Plugins should remain an independent top-level Preferences window during app shell-stability validation.");
    PostMessageW(prefs, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Manage Plugins Preferences window did not close during app shell-stability validation.");
    state.Require(WaitForNoNonBaselineWindows(processId, baselineWindows, mainWindow, SelfTest::Scale(3000ms)),
                  L"App Manage Plugins should not leave non-baseline top-level windows behind after close.");
    static_cast<void>(SetForegroundWindow(mainWindow));
    static_cast<void>(SetFocus(mainWindow));
    state.Require(waitForLeftFolderFocus(), L"App Manage Plugins should restore folder-view focus to the active pane after close.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, true, L"App Manage Plugins close");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, false, L"App Manage Plugins close");

    return state.failure.empty();
}

[[nodiscard]] bool TestAppAboutKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"app_about_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for app About shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for app About shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for app About shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for app About shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    if (const HWND existing = GetAboutDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing About dialog did not close before app About shell-stability validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during app About shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during app About shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for app About shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for app About shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for app About shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for app About shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before app About shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder-view handle unavailable for app About shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD processId               = GetCurrentProcessId();
    const auto baselineWindows          = SnapshotTopLevelWindowsForProcess(processId);
    const size_t baselineLeftItemCount  = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before app About shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before app About shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  const std::filesystem::path& expectedPath,
                                                  const size_t expectedItemCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  const bool requireFocusedFolderView,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(
            WaitForNavigationViewSnapshot(pane,
                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == 0u &&
                   ! g_folderWindow.DebugIsNameFilterActive(pane) &&
                   (! requireFocusedFolderView || g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView);
        },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
            std::format(
                L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, popupVisible={}, "
                L"childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}, focusedFolderViewMatches={}.",
                context,
                static_cast<unsigned>(pane),
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetItemCount(pane),
                g_folderWindow.DebugGetSelectedCount(pane),
                g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no",
                g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView ? L"yes" : L"no"));
    };

    struct AboutProbeResult final
    {
        HWND about                     = nullptr;
        bool sawDialog                 = false;
        bool independentTopLevel       = false;
        bool exposesUiaProvider        = false;
        size_t visibleChildWindowCount = 0u;
        bool closed                    = false;
    } aboutProbe{};

    std::jthread aboutWorker([&](std::stop_token) noexcept
    {
        const HWND about     = WaitForWindow([] noexcept { return GetAboutDialogHandle(); }, SelfTest::Scale(5000ms));
        aboutProbe.about     = about;
        aboutProbe.sawDialog = about != nullptr && IsWindow(about) != FALSE;
        if (! aboutProbe.sawDialog)
        {
            return;
        }

        aboutProbe.independentTopLevel     = ! IsOwnedBy(about, mainWindow);
        aboutProbe.exposesUiaProvider      = WindowExposesUiaProvider(about);
        aboutProbe.visibleChildWindowCount = CountVisibleChildWindows(about);
        PostMessageW(about, WM_CLOSE, 0, 0);
        aboutProbe.closed = WaitForWindowClosed(about, SelfTest::Scale(3000ms));
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_ABOUT, 0), 0);
    aboutWorker.join();

    state.Require(aboutProbe.sawDialog, L"About dialog did not open during app About shell-stability validation.");
    if (! aboutProbe.sawDialog)
    {
        return false;
    }

    state.Require(aboutProbe.independentTopLevel, L"About dialog should remain an independent top-level window during app About shell-stability validation.");
    state.Require(aboutProbe.exposesUiaProvider,
                  L"About dialog should answer WM_GETOBJECT with a UI Automation provider during app About shell-stability validation.");
    state.Require(aboutProbe.visibleChildWindowCount == 0u,
                  L"About dialog should not expose visible child-control fallback during app About shell-stability validation.");
    state.Require(aboutProbe.closed, L"About dialog did not close during app About shell-stability validation.");
    state.Require(WaitForNoNonBaselineWindows(processId, baselineWindows, mainWindow, SelfTest::Scale(3000ms)),
                  L"App About should not leave non-baseline top-level windows behind after close.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, true, L"App About close");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, false, L"App About close");

    return state.failure.empty();
}

[[nodiscard]] bool TestAppShowShortcutsKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"app_shortcuts_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for app Shortcuts shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for app Shortcuts shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for app Shortcuts shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for app Shortcuts shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    if (const HWND existing = GetShortcutsWindowHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing Shortcuts window did not close before app Shortcuts shell-stability validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during app Shortcuts shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during app Shortcuts shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for app Shortcuts shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for app Shortcuts shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for app Shortcuts shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for app Shortcuts shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before app Shortcuts shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder-view handle unavailable for app Shortcuts shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD processId               = GetCurrentProcessId();
    const auto baselineWindows          = SnapshotTopLevelWindowsForProcess(processId);
    const size_t baselineLeftItemCount  = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before app Shortcuts shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before app Shortcuts shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  const std::filesystem::path& expectedPath,
                                                  const size_t expectedItemCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  const bool requireFocusedFolderView,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(
            WaitForNavigationViewSnapshot(pane,
                                          [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == 0u &&
                   ! g_folderWindow.DebugIsNameFilterActive(pane) &&
                   (! requireFocusedFolderView || g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView);
        },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
            std::format(
                L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, popupVisible={}, "
                L"childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}, focusedFolderViewMatches={}.",
                context,
                static_cast<unsigned>(pane),
                static_cast<unsigned>(snapshot.focusTarget),
                snapshot.editMode ? L"yes" : L"no",
                snapshot.historyDropdownVisible ? L"yes" : L"no",
                snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                snapshot.fullPathPopupVisible ? L"yes" : L"no",
                snapshot.visibleChildWindowCount,
                snapshot.currentPathText,
                snapshot.historyCount,
                g_folderWindow.DebugGetItemCount(pane),
                g_folderWindow.DebugGetSelectedCount(pane),
                g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no",
                g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView ? L"yes" : L"no"));
    };

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SHOW_SHORTCUTS, 0), 0);
    const HWND shortcuts = WaitForWindow([] noexcept { return GetShortcutsWindowHandle(); }, SelfTest::Scale(3000ms));
    state.Require(shortcuts != nullptr && IsWindow(shortcuts) != FALSE, L"Shortcuts window did not open during app Shortcuts shell-stability validation.");
    if (! shortcuts || IsWindow(shortcuts) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(shortcuts, mainWindow),
                  L"Shortcuts window should remain an independent top-level window during app Shortcuts shell-stability validation.");
    state.Require(WindowExposesUiaProvider(shortcuts),
                  L"Shortcuts window should answer WM_GETOBJECT with a UI Automation provider during app Shortcuts shell-stability validation.");
    state.Require(CountVisibleChildWindows(shortcuts) == 0u,
                  L"Shortcuts window should not expose visible child-control fallback during app Shortcuts shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    PostMessageW(shortcuts, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(shortcuts, SelfTest::Scale(3000ms)), L"Shortcuts window did not close during app Shortcuts shell-stability validation.");
    state.Require(WaitForNoNonBaselineWindows(processId, baselineWindows, mainWindow, SelfTest::Scale(3000ms)),
                  L"App Shortcuts should not leave non-baseline top-level windows behind after close.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, true, L"App Shortcuts close");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, false, L"App Shortcuts close");

    return state.failure.empty();
}

[[nodiscard]] bool TestSwapPanesCommand(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePaths                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");

    const std::filesystem::path root  = suiteRoot / L"work" / L"swap_panes";
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create swap_panes left folder.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create swap_panes right folder.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Failed to set left pane path for swap test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"Failed to set right pane path for swap test.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SWAP_PANES, 0), 0);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, right, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"SwapPanes did not move right path into left pane.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, left, SelfTest::Scale(std::chrono::milliseconds{2000})),
                  L"SwapPanes did not move left path into right pane.");

    return state.failure.empty();
}

[[nodiscard]] bool TestAppCompareKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path leftRoot  = suiteRoot / L"work" / (L"app_compare_nav_left_" + NewGuidText());
    const std::filesystem::path rightRoot = suiteRoot / L"work" / (L"app_compare_nav_right_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create app Compare Directories left root.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create app Compare Directories right root.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"a.txt", "alpha"), L"Failed to create left a.txt for app Compare Directories shell-stability test.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"b.log", "bravo"), L"Failed to create left b.log for app Compare Directories shell-stability test.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"a.txt", "alpha"), L"Failed to create right a.txt for app Compare Directories shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePane                                 = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    if (const HWND existingCompare = GetCompareDirectoriesWindowHandle(); existingCompare && IsWindow(existingCompare) != FALSE)
    {
        SendMessageW(existingCompare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_CLOSE, 0), 0);
        state.Require(WaitForWindowClosed(existingCompare, SelfTest::Scale(3000ms)),
                      L"Existing Compare Directories window did not close before app shell-stability validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left app Compare Directories shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right app Compare Directories shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for app Compare Directories shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for app Compare Directories shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for app Compare Directories shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"a.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for app Compare Directories shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"Failed to focus b.log before app Compare Directories shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE,
                  L"Folder view handle unavailable for app Compare Directories shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint64_t baselineRefreshCount = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    const size_t baselineItemCount      = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineSelectedCount  = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

    NavigationViewDebugSnapshot baselineSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == leftRoot.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineSnapshot),
                  L"Failed to capture the baseline navigation-view state before app Compare Directories shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_COMPARE, 0), 0);

    const HWND compare = WaitForWindow([] noexcept { return GetCompareDirectoriesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(compare != nullptr && IsWindow(compare) != FALSE, L"App Compare Directories command did not open the window.");
    if (! compare || IsWindow(compare) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    CompareDirectoriesRunDebugSnapshot compareSnapshot{};
    bool compareSettled        = false;
    const auto compareDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < compareDeadline)
    {
        PumpPendingMessages();
        compareSnapshot = {};
        if (DebugGetCompareDirectoriesRunSnapshot(compareSnapshot) && compareSnapshot.windowVisible && ! compareSnapshot.compareRunPending)
        {
            compareSettled = true;
            break;
        }

        Sleep(10);
    }

    state.Require(compareSettled,
                  std::format(L"App Compare Directories did not settle after cmd/app/compare; windowVisible={}, optionsVisible={}, compareStarted={}, "
                              L"compareActive={}, pending={}.",
                              compareSnapshot.windowVisible ? L"yes" : L"no",
                              compareSnapshot.optionsDialogVisible ? L"yes" : L"no",
                              compareSnapshot.compareStarted ? L"yes" : L"no",
                              compareSnapshot.compareActive ? L"yes" : L"no",
                              compareSnapshot.compareRunPending ? L"yes" : L"no"));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(compare, WM_COMMAND, MAKEWPARAM(IDM_COMPARE_CLOSE, 0), 0);
    state.Require(WaitForWindowClosed(compare, SelfTest::Scale(3000ms)), L"App Compare Directories window did not close after shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot snapshot{};
    std::optional<std::filesystem::path> restoredPath;
    bool shellRestored         = false;
    const auto restoreDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < restoreDeadline)
    {
        PumpPendingMessages();

        restoredPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        static_cast<void>(g_folderWindow.DebugGetNavigationViewSnapshot(FolderWindow::Pane::Left, snapshot));

        const bool panePathStable = restoredPath.has_value() && OrdinalString::EqualsNoCasePath(restoredPath.value(), leftRoot);
        if (snapshot.focusTarget == NavigationViewDebugFocusTarget::None && ! snapshot.editMode && ! snapshot.historyDropdownVisible &&
            ! snapshot.editSuggestPopupVisible && ! snapshot.fullPathPopupVisible && ! snapshot.fullPathPopupEditMode &&
            snapshot.visibleChildWindowCount == 0u && panePathStable && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
            g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"b.log" &&
            g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) == baselineRefreshCount &&
            g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == baselineItemCount &&
            g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == baselineSelectedCount)
        {
            shellRestored = true;
            break;
        }

        Sleep(20);
    }

    state.Require(shellRestored,
                  std::format(L"Navigation shell did not restore cleanly after app Compare Directories close; focusTarget={}, editMode={}, historyVisible={}, "
                              L"suggestVisible={}, popupVisible={}, childWindows={}, currentPath='{}', panePath='{}', historyCount={}, refreshCount={}, "
                              L"itemCount={}, selectedCount={}, focusedItem='{}'.",
                              static_cast<unsigned>(snapshot.focusTarget),
                              snapshot.editMode ? L"yes" : L"no",
                              snapshot.historyDropdownVisible ? L"yes" : L"no",
                              snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                              snapshot.fullPathPopupVisible ? L"yes" : L"no",
                              snapshot.visibleChildWindowCount,
                              snapshot.currentPathText,
                              restoredPath.has_value() ? restoredPath->wstring() : std::wstring{},
                              snapshot.historyCount,
                              g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left),
                              g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left)));
    return state.failure.empty();
}

[[nodiscard]] bool TestSwapPanesKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"swap_panes_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create swap-panes shell-stability left folder.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create swap-panes shell-stability right folder.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for swap-panes shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for swap-panes shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePanes                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during swap-panes shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during swap-panes shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<uint32_t> leftEnumCount{0};
    std::atomic<uint32_t> rightEnumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, left) || OrdinalString::EqualsNoCasePath(folder, right))
        {
            leftEnumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, left) || OrdinalString::EqualsNoCasePath(folder, right))
        {
            rightEnumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallbacks = wil::scope_exit([&]
    {
        g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {});
        g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right, {});
    });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for swap-panes shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for swap-panes shell-stability test.");
    state.Require(WaitForAtomicAtLeast(leftEnumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Left pane enumeration did not complete for swap-panes shell-stability test.");
    state.Require(WaitForAtomicAtLeast(rightEnumCount, 1u, SelfTest::Scale(3000ms)),
                  L"Right pane enumeration did not complete for swap-panes shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for swap-panes shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for swap-panes shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before swap-panes shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder-view handle unavailable for swap-panes shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before swap-panes shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before swap-panes shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint32_t leftEnumBefore  = leftEnumCount.load(std::memory_order_acquire);
    const uint32_t rightEnumBefore = rightEnumCount.load(std::memory_order_acquire);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_APP_SWAP_PANES, 0), 0);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, right, SelfTest::Scale(3000ms)),
                  L"Swap Panes did not move the right path into the left pane during shell-stability validation.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, left, SelfTest::Scale(3000ms)),
                  L"Swap Panes did not move the left path into the right pane during shell-stability validation.");
    state.Require(WaitForAtomicAtLeast(leftEnumCount, leftEnumBefore + 1u, SelfTest::Scale(3000ms)),
                  L"Left pane enumeration did not refresh after Swap Panes during shell-stability validation.");
    state.Require(WaitForAtomicAtLeast(rightEnumCount, rightEnumBefore + 1u, SelfTest::Scale(3000ms)),
                  L"Right pane enumeration did not refresh after Swap Panes during shell-stability validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready after Swap Panes during shell-stability validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready after Swap Panes during shell-stability validation.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"right.txt"),
                  L"right.txt should be visible in the left pane after Swap Panes.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Right, L"left.txt"),
                  L"left.txt should be visible in the right pane after Swap Panes.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  std::filesystem::path expectedPath,
                                                  std::wstring_view expectedItemName,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  bool requireFocusedFolderView,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount >= baselineSnapshot.historyCount &&
                   (! requireFocusedFolderView || g_folderWindow.GetFocusedFolderViewHwnd() == leftFolderView) &&
                   g_folderWindow.DebugGetItemCount(pane) == 1u && ! g_folderWindow.DebugIsNameFilterActive(pane) &&
                   g_folderWindow.DebugHasItemDisplayName(pane, expectedItemName);
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, nameFilterActive={}.",
                                  context,
                                  static_cast<unsigned>(pane),
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(pane),
                                  g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no"));
    };

    requireStableNavigationShell(FolderWindow::Pane::Left, right, L"right.txt", baselineLeftSnapshot, true, L"Swap Panes (left pane after swap)");
    requireStableNavigationShell(FolderWindow::Pane::Right, left, L"left.txt", baselineRightSnapshot, false, L"Swap Panes (right pane after swap)");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 0u,
                  L"Swap Panes should not create left-pane selection while keeping the navigation shell quiet.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == 0u,
                  L"Swap Panes should not create right-pane selection while keeping the navigation shell quiet.");

    return state.failure.empty();
}

[[nodiscard]] bool TestToggleUiChromeKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"toggle_ui_chrome_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for toggle-ui-chrome shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for toggle-ui-chrome shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for toggle-ui-chrome shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for toggle-ui-chrome shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const bool issuesBefore                                = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto restoreState                                = wil::scope_exit([&]
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() != issuesBefore)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        }
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during toggle-ui-chrome shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during toggle-ui-chrome shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for toggle-ui-chrome shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for toggle-ui-chrome shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for toggle-ui-chrome shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for toggle-ui-chrome shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before toggle-ui-chrome shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const HWND leftFolderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(leftFolderView != nullptr && IsWindow(leftFolderView) != FALSE,
                  L"Left folder-view handle unavailable for toggle-ui-chrome shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineLeftItemCount  = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before toggle-ui-chrome shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before toggle-ui-chrome shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  std::filesystem::path expectedPath,
                                                  size_t expectedItemCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == 0u &&
                   ! g_folderWindow.DebugIsNameFilterActive(pane);
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}.",
                                  context,
                                  static_cast<unsigned>(pane),
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(pane),
                                  g_folderWindow.DebugGetSelectedCount(pane),
                                  g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no"));
    };

    const bool menuBarBefore = DebugIsMainMenuBarSurfaceVisible(mainWindow);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
    state.Require(DebugIsMainMenuBarSurfaceVisible(mainWindow) != menuBarBefore,
                  L"Toggle Menu Bar did not change the visible DxUI menu-bar surface during shell-stability validation.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Toggle Menu Bar off");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Toggle Menu Bar off");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_MENUBAR, 0), 0);
    state.Require(DebugIsMainMenuBarSurfaceVisible(mainWindow) == menuBarBefore,
                  L"Toggle Menu Bar did not restore the visible DxUI menu-bar surface during shell-stability validation.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Toggle Menu Bar on");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Toggle Menu Bar on");

    const bool functionBarBefore = g_folderWindow.GetFunctionBarVisible();
    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FUNCTIONBAR, 0), 0);
    state.Require(g_folderWindow.GetFunctionBarVisible() != functionBarBefore,
                  L"Toggle Function Bar did not change function-bar visibility during shell-stability validation.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Toggle Function Bar off");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Toggle Function Bar off");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FUNCTIONBAR, 0), 0);
    state.Require(g_folderWindow.GetFunctionBarVisible() == functionBarBefore,
                  L"Toggle Function Bar did not restore function-bar visibility during shell-stability validation.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Toggle Function Bar on");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Toggle Function Bar on");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    state.Require(g_folderWindow.IsFileOperationsIssuesPaneVisible() != issuesBefore,
                  L"Toggle File Operations Issues Pane did not change issues-pane visibility during shell-stability validation.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Toggle File Operations Issues Pane open");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Toggle File Operations Issues Pane open");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    state.Require(g_folderWindow.IsFileOperationsIssuesPaneVisible() == issuesBefore,
                  L"Toggle File Operations Issues Pane did not restore issues-pane visibility during shell-stability validation.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Toggle File Operations Issues Pane closed");
    requireStableNavigationShell(FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Toggle File Operations Issues Pane closed");

    return state.failure.empty();
}

[[nodiscard]] bool TestAppFileOperationsIssuesPaneKeepsNavigationShellStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root  = suiteRoot / L"work" / (L"app_fileops_issues_nav_shell_" + NewGuidText());
    const std::filesystem::path left  = root / L"left";
    const std::filesystem::path right = root / L"right";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(left), L"Failed to create left folder for file-operations issues-pane shell-stability test.");
    state.Require(SelfTest::EnsureDirectory(right), L"Failed to create right folder for file-operations issues-pane shell-stability test.");
    state.Require(SelfTest::WriteTextFile(left / L"left.txt", "left"), L"Failed to create left.txt for file-operations issues-pane shell-stability test.");
    state.Require(SelfTest::WriteTextFile(right / L"right.txt", "right"), L"Failed to create right.txt for file-operations issues-pane shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                    = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::wstring rightPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right));
    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const bool issuesBefore                                = g_folderWindow.IsFileOperationsIssuesPaneVisible();
    const auto restoreState                                = wil::scope_exit([&]
    {
        if (g_folderWindow.IsFileOperationsIssuesPaneVisible() != issuesBefore)
        {
            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
        }
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, rightPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for left pane during file-operations issues-pane shell-stability test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")) != FALSE,
                  L"Failed to set local file-system plugin for right pane during file-operations issues-pane shell-stability test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, left);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, right);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, left, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for file-operations issues-pane shell-stability test.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, right, SelfTest::Scale(3000ms)),
                  L"Failed to set right pane path for file-operations issues-pane shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"left.txt"}, SelfTest::Scale(3000ms)),
                  L"Left pane contents not ready for file-operations issues-pane shell-stability test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"right.txt"}, SelfTest::Scale(3000ms)),
                  L"Right pane contents not ready for file-operations issues-pane shell-stability test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"left.txt"),
                  L"Failed to focus left.txt before file-operations issues-pane shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    const size_t baselineLeftItemCount  = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left);
    const size_t baselineRightItemCount = g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Right);
    NavigationViewDebugSnapshot baselineLeftSnapshot{};
    NavigationViewDebugSnapshot baselineRightSnapshot{};
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Left,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == left.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineLeftSnapshot),
                  L"Failed to capture the baseline left navigation-view state before file-operations issues-pane shell-stability validation.");
    state.Require(WaitForNavigationViewSnapshot(FolderWindow::Pane::Right,
                                                [&](const NavigationViewDebugSnapshot& value) noexcept
    {
        return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
               ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
               value.currentPathText == right.wstring();
    },
                                                SelfTest::Scale(3000ms),
                                                &baselineRightSnapshot),
                  L"Failed to capture the baseline right navigation-view state before file-operations issues-pane shell-stability validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto requireStableNavigationShell = [&](FolderWindow::Pane pane,
                                                  const std::filesystem::path& expectedPath,
                                                  const size_t expectedItemCount,
                                                  const NavigationViewDebugSnapshot& baselineSnapshot,
                                                  std::wstring_view context) noexcept
    {
        NavigationViewDebugSnapshot snapshot{};
        state.Require(WaitForNavigationViewSnapshot(pane,
                                                    [&](const NavigationViewDebugSnapshot& value) noexcept
        {
            return value.focusTarget == NavigationViewDebugFocusTarget::None && ! value.editMode && ! value.historyDropdownVisible &&
                   ! value.editSuggestPopupVisible && ! value.fullPathPopupVisible && ! value.fullPathPopupEditMode && value.visibleChildWindowCount == 0u &&
                   value.currentPathText == expectedPath.wstring() && value.historyCount == baselineSnapshot.historyCount &&
                   g_folderWindow.DebugGetItemCount(pane) == expectedItemCount && g_folderWindow.DebugGetSelectedCount(pane) == 0u &&
                   ! g_folderWindow.DebugIsNameFilterActive(pane);
        },
                                                    SelfTest::Scale(3000ms),
                                                    &snapshot),
                      std::format(L"Navigation shell did not stay quiet during {}; pane={}, focusTarget={}, editMode={}, historyVisible={}, suggestVisible={}, "
                                  L"popupVisible={}, childWindows={}, currentPath='{}', historyCount={}, itemCount={}, selectedCount={}, nameFilterActive={}.",
                                  context,
                                  static_cast<unsigned>(pane),
                                  static_cast<unsigned>(snapshot.focusTarget),
                                  snapshot.editMode ? L"yes" : L"no",
                                  snapshot.historyDropdownVisible ? L"yes" : L"no",
                                  snapshot.editSuggestPopupVisible ? L"yes" : L"no",
                                  snapshot.fullPathPopupVisible ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.currentPathText,
                                  snapshot.historyCount,
                                  g_folderWindow.DebugGetItemCount(pane),
                                  g_folderWindow.DebugGetSelectedCount(pane),
                                  g_folderWindow.DebugIsNameFilterActive(pane) ? L"yes" : L"no"));
    };

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    state.Require(g_folderWindow.IsFileOperationsIssuesPaneVisible() != issuesBefore,
                  L"Standalone File Operations Issues Pane command did not change issues-pane visibility during shell-stability validation.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Standalone File Operations Issues Pane open");
    requireStableNavigationShell(
        FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Standalone File Operations Issues Pane open");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_FILEOPS_FAILED_ITEMS, 0), 0);
    state.Require(g_folderWindow.IsFileOperationsIssuesPaneVisible() == issuesBefore,
                  L"Standalone File Operations Issues Pane command did not restore issues-pane visibility during shell-stability validation.");
    requireStableNavigationShell(FolderWindow::Pane::Left, left, baselineLeftItemCount, baselineLeftSnapshot, L"Standalone File Operations Issues Pane closed");
    requireStableNavigationShell(
        FolderWindow::Pane::Right, right, baselineRightItemCount, baselineRightSnapshot, L"Standalone File Operations Issues Pane closed");

    return state.failure.empty();
}

[[nodiscard]] bool TestDisplayModeAndSortCommands(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FolderWindow::Pane pane = FolderWindow::Pane::Left;
    FocusFolderViewPane(pane);

    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(pane);
    const FolderView::SortBy sortBefore         = g_folderWindow.GetSortBy(pane);
    const FolderView::SortDirection dirBefore   = g_folderWindow.GetSortDirection(pane);
    const auto restore                          = wil::scope_exit([&]
    {
        g_folderWindow.SetActivePane(pane);
        g_folderWindow.SetDisplayMode(pane, displayBefore);
        g_folderWindow.SetSort(pane, sortBefore, dirBefore);
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_DISPLAY_DETAILED, 0), 0);
    state.Require(g_folderWindow.GetDisplayMode(pane) == FolderView::DisplayMode::Detailed, L"Display mode did not switch to Detailed.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_DISPLAY_EXTRA_DETAILED, 0), 0);
    state.Require(g_folderWindow.GetDisplayMode(pane) == FolderView::DisplayMode::ExtraDetailed, L"Display mode did not switch to ExtraDetailed.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_THUMBNAILS, 0), 0);
    state.Require(g_folderWindow.GetDisplayMode(pane) == FolderView::DisplayMode::Thumbnails,
                  L"Thumbnails command did not select the exclusive Thumbnails display mode.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_DISPLAY_BRIEF, 0), 0);
    state.Require(g_folderWindow.GetDisplayMode(pane) == FolderView::DisplayMode::Brief, L"Display mode did not switch to Brief.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SORT_NONE, 0), 0);
    state.Require(g_folderWindow.GetSortBy(pane) == FolderView::SortBy::None, L"Sort none did not set sort-by None.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SORT_NAME, 0), 0);
    state.Require(g_folderWindow.GetSortBy(pane) == FolderView::SortBy::Name, L"Sort by Name did not set sort-by Name.");

    const FolderView::SortDirection dir1 = g_folderWindow.GetSortDirection(pane);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SORT_NAME, 0), 0);
    state.Require(g_folderWindow.GetSortBy(pane) == FolderView::SortBy::Name, L"Second Sort by Name did not keep sort-by Name.");
    const FolderView::SortDirection dir2 = g_folderWindow.GetSortDirection(pane);
    state.Require(dir2 != dir1, L"Second Sort by Name did not change sort direction.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneStatusBarUsesOwnedWindowAndSortClickOpensMenu(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FolderWindow::Pane pane = FolderWindow::Pane::Left;
    const bool visibleBefore      = g_folderWindow.GetStatusBarVisible(pane);
    const auto restore            = wil::scope_exit([&]() noexcept
    {
        g_folderWindow.SetStatusBarVisible(pane, visibleBefore);
        const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
        static_cast<void>(EnsureUiNotInMenuMode(uiThreadId, mainWindow, SelfTest::Scale(2000ms)));
    });

    g_folderWindow.SetStatusBarVisible(pane, true);
    FocusFolderViewPane(pane);
    PumpPendingMessages();

    FolderWindow::FolderWindowPaneStatusBarDebugSnapshot snapshot{};
    state.Require(g_folderWindow.DebugGetPaneStatusBarSnapshot(pane, snapshot), L"Failed to capture pane status-bar snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.visible, L"Pane status bar should be visible.");
    state.Require(! snapshot.usesNativeStatusBarClass, L"Pane status bar should not use the native status-bar class.");
    state.Require(snapshot.usesDirectWriteTextRendering, L"Pane status bar should render text through the DirectWrite status-bar path.");
    state.Require(! snapshot.hasNativeFont, L"Pane status bar should not retain a native font assignment.");
    state.Require(! snapshot.sortText.empty(), L"Pane status bar should expose non-empty sort text.");

    constexpr int kLeftStatusBarChildId = 1005;

    const HWND folderWindow = g_folderWindow.GetHwnd();
    state.Require(folderWindow != nullptr && IsWindow(folderWindow) != FALSE, L"FolderWindow handle invalid.");
    const HWND statusBar = folderWindow ? GetDlgItem(folderWindow, kLeftStatusBarChildId) : nullptr;
    state.Require(statusBar != nullptr && IsWindow(statusBar) != FALSE, L"Pane status-bar HWND missing.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
    if (const HWND existingPopup = FindVisibleOwnedDxUiContextMenuWindow(mainWindow); existingPopup)
    {
        PostMessageW(existingPopup, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(existingPopup, WM_KEYUP, VK_ESCAPE, 0);
        state.Require(WaitForWindowClosed(existingPopup, SelfTest::Scale(2000ms)),
                      L"Existing DxUI context menu popup did not close before pane status-bar sort click validation.");
    }

    {
        std::atomic<bool> menuOpened{false};
        std::atomic<bool> menuClosed{false};
        std::atomic<bool> sawDxUiPopup{false};
        std::atomic<bool> popupDebugStateReadable{false};
        std::atomic<bool> popupSurfaceAboveStatusBar{false};

        std::jthread closer([&](std::stop_token stopToken) noexcept
        {
            const auto openDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < openDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
                const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(mainWindow);
                if (popup != nullptr && IsWindow(popup) != FALSE)
                {
                    sawDxUiPopup.store(true, std::memory_order_release);
                    RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
                    const bool debugReadable = RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState);
                    popupDebugStateReadable.store(debugReadable, std::memory_order_release);
                    RECT statusRect{};
                    if (debugReadable && GetWindowRect(statusBar, &statusRect) != FALSE)
                    {
                        popupSurfaceAboveStatusBar.store(popupState.surfaceRectPx.bottom <= statusRect.top, std::memory_order_release);
                    }
                }

                if (inMenuMode || popup != nullptr)
                {
                    menuOpened.store(true, std::memory_order_release);
                    break;
                }

                std::this_thread::sleep_for(10ms);
            }

            if (! menuOpened.load(std::memory_order_acquire))
            {
                return;
            }

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                GUITHREADINFO gti{};
                gti.cbSize            = sizeof(gti);
                const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
                const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
                const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(mainWindow);
                if (! inMenuMode && popup == nullptr)
                {
                    menuClosed.store(true, std::memory_order_release);
                    return;
                }

                const HWND dismissTarget =
                    popup != nullptr ? popup
                                     : (hasGuiInfo && gti.hwndMenuOwner ? gti.hwndMenuOwner : (hasGuiInfo && gti.hwndActive ? gti.hwndActive : statusBar));
                if (dismissTarget != nullptr)
                {
                    PostMessageW(dismissTarget, WM_KEYDOWN, VK_ESCAPE, 0);
                    PostMessageW(dismissTarget, WM_KEYUP, VK_ESCAPE, 0);
                }

                std::this_thread::sleep_for(30ms);
            }

            GUITHREADINFO gti{};
            gti.cbSize            = sizeof(gti);
            const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
            const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
            const HWND popup      = FindVisibleOwnedDxUiContextMenuWindow(mainWindow);
            if (! inMenuMode && popup == nullptr)
            {
                menuClosed.store(true, std::memory_order_release);
            }
        });

        state.Require(g_folderWindow.DebugClickPaneStatusBarSort(pane), L"Failed to post pane status-bar sort click.");
        if (state.failure.empty())
        {
            PumpPendingMessages();
        }
        closer.join();

        state.Require(menuOpened.load(std::memory_order_acquire), L"Pane status-bar sort click did not open the sort menu popup.");
        state.Require(menuClosed.load(std::memory_order_acquire), L"Pane status-bar sort click did not dismiss the sort menu popup.");
        if (sawDxUiPopup.load(std::memory_order_acquire))
        {
            state.Require(popupDebugStateReadable.load(std::memory_order_acquire),
                          L"Pane status-bar DxUI sort popup did not expose a readable popup debug state.");
            state.Require(popupSurfaceAboveStatusBar.load(std::memory_order_acquire),
                          L"Pane status-bar DxUI sort popup should open above the status bar instead of covering the bottom chrome.");
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestEmptyFolderState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"empty_folder_state_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create empty-folder-state test root.");

    const std::filesystem::path emptyChild = root / L"empty";
    state.Require(SelfTest::EnsureDirectory(emptyChild), L"Failed to create empty child folder.");

    {
        const auto metrics = FolderViewEmptyStateLayout::ResolvePlaceholderItemMetrics(FolderViewEmptyStateLayout::PlaceholderItemMetricsInput{
            .clientWidthDip          = 900.0f,
            .clientHeightDip         = 700.0f,
            .iconSizeDip             = 20.0f,
            .estimatedCharWidthDip   = 8.0f,
            .estimatedLabelHeightDip = 16.0f,
            .detailsLineHeightDip    = 12.0f,
            .metadataLineHeightDip   = 12.0f,
            .titleLength             = 12u,
            .includeDetailsLine      = false,
            .includeMetadataLine     = false,
        });
        state.Require(std::abs(metrics.tileWidthDip - 900.0f) <= 0.1f, L"Empty-folder pseudo item should span the pane item row width.");
        state.Require(std::abs(metrics.tileHeightDip - 28.0f) <= 0.1f,
                      std::format(L"Empty-folder pseudo item should be one row tall, not full-client. height={:.1f}", metrics.tileHeightDip));
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumEmpty{0};
    std::atomic<uint32_t> enumParent{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, emptyChild))
        {
            enumEmpty.fetch_add(1u, std::memory_order_release);
        }
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumParent.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, emptyChild);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, emptyChild, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for empty-folder-state test.");
    state.Require(WaitForAtomicAtLeast(enumEmpty, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for empty folder in empty-folder-state test.");

    state.Require(g_folderWindow.DebugIsEmptyFolderStateActive(FolderWindow::Pane::Left), L"Expected empty-folder state active for empty folder.");
    state.Require(! g_folderWindow.DebugGetEmptyFolderFunMessage(FolderWindow::Pane::Left).empty(),
                  L"Expected empty-folder fun message to be populated from resources.");
    {
        const std::wstring emptyTitle = LoadStringResource(nullptr, IDS_EMPTY_FOLDER_TITLE);
        const std::wstring parentRow  = LoadStringResource(nullptr, IDS_EMPTY_FOLDER_PARENT_ROW);
        state.Require(emptyTitle == L"Empty folder", L"Expected centered empty-folder title resource.");
        state.Require(parentRow == L"Go to parent", L"Expected empty-folder focused row to use the Go to parent resource.");
        state.Require(parentRow != emptyTitle, L"Empty-folder focused row label should be distinct from the centered title.");
    }
    {
        const FolderView::DebugEmptyFolderItemMetrics metrics = g_folderWindow.DebugGetEmptyFolderItemMetrics(FolderWindow::Pane::Left);
        state.Require(metrics.active, L"Expected empty-folder pseudo item metrics to be active.");
        state.Require(metrics.tileWidthDip >= metrics.clientWidthDip - 1.0f,
                      std::format(L"Empty-folder pseudo item should use current pane row width. width={:.1f} client={:.1f}",
                                  metrics.tileWidthDip,
                                  metrics.clientWidthDip));
        state.Require(metrics.tileHeightDip > 0.0f && metrics.tileHeightDip <= 96.0f,
                      std::format(L"Empty-folder pseudo item should be row-sized, not full-pane. height={:.1f}", metrics.tileHeightDip));
        if (metrics.clientHeightDip > 120.0f)
        {
            state.Require(metrics.tileHeightDip < metrics.clientHeightDip * 0.25f,
                          std::format(L"Empty-folder pseudo item height should not track full pane height. height={:.1f} client={:.1f}",
                                      metrics.tileHeightDip,
                                      metrics.clientHeightDip));
        }
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView && IsWindow(folderView), L"Left FolderView hwnd invalid.");

    RECT rc{};
    GetClientRect(folderView, &rc);
    const int x = (rc.left + rc.right) / 2;
    const int y = (rc.top + rc.bottom) / 2;

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(folderView, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(folderView, WM_KEYUP, VK_RETURN, 0);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enter in empty folder did not navigate up to parent.");
    state.Require(WaitForAtomicAtLeast(enumParent, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for parent folder after empty-folder Enter.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, emptyChild);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, emptyChild, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to restore empty child after empty-folder Enter.");
    state.Require(WaitForAtomicAtLeast(enumEmpty, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete after restoring empty child for empty-folder test.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(folderView, WM_LBUTTONDBLCLK, 0, MAKELPARAM(x, y));

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Double click in empty folder did not navigate up to parent.");
    state.Require(WaitForAtomicAtLeast(enumParent, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for parent folder after empty-folder double click.");

    return state.failure.empty();
}

[[nodiscard]] std::wstring FolderViewColumnDisplayModeName(FolderView::DisplayMode mode)
{
    switch (mode)
    {
        case FolderView::DisplayMode::Brief: return L"Brief";
        case FolderView::DisplayMode::Detailed: return L"Detailed";
        case FolderView::DisplayMode::ExtraDetailed: return L"ExtraDetailed";
        case FolderView::DisplayMode::Thumbnails: return L"Thumbnails";
    }

    return L"Unknown";
}

void AppendFolderViewColumnJsonString(std::wstring& out, std::wstring_view value)
{
    out.push_back(L'"');
    for (const wchar_t ch : value)
    {
        switch (ch)
        {
            case L'"': out.append(L"\\\""); break;
            case L'\\': out.append(L"\\\\"); break;
            case L'\b': out.append(L"\\b"); break;
            case L'\f': out.append(L"\\f"); break;
            case L'\n': out.append(L"\\n"); break;
            case L'\r': out.append(L"\\r"); break;
            case L'\t': out.append(L"\\t"); break;
            default:
                if (ch < 0x20)
                {
                    out.append(std::format(L"\\u{:04X}", static_cast<unsigned int>(ch)));
                }
                else
                {
                    out.push_back(ch);
                }
                break;
        }
    }
    out.push_back(L'"');
}

[[nodiscard]] std::wstring RepeatForFolderViewColumnName(wchar_t ch, size_t count)
{
    return std::wstring(count, ch);
}

struct FolderViewColumnAuditIntegrity
{
    uint32_t overlapCount                   = 0;
    uint32_t outOfOrderColumnCount          = 0;
    uint32_t negativeWidthCount             = 0;
    uint32_t visibleRangeMismatchCount      = 0;
    uint32_t hitTestSpacingFalsePositiveCount = 0;
};

[[nodiscard]] bool FolderViewColumnRectsOverlap(const D2D1_RECT_F& left, const D2D1_RECT_F& right) noexcept
{
    constexpr float kEpsilon = 0.5f;
    return left.left < right.right - kEpsilon && left.right > right.left + kEpsilon && left.top < right.bottom - kEpsilon &&
           left.bottom > right.top + kEpsilon;
}

[[nodiscard]] FolderViewColumnAuditIntegrity ComputeFolderViewColumnAuditIntegrity(
    const FolderView::DebugColumnLayoutSnapshot& snapshot) noexcept
{
    FolderViewColumnAuditIntegrity integrity{};
    integrity.hitTestSpacingFalsePositiveCount = snapshot.hitTestSpacingFalsePositiveCount;

    float previousRight = -std::numeric_limits<float>::infinity();
    for (const auto& column : snapshot.columns)
    {
        if (column.widthDip <= 0.0f || column.rightDip < column.leftDip)
        {
            ++integrity.negativeWidthCount;
        }
        if (column.leftDip < previousRight - 0.5f)
        {
            ++integrity.outOfOrderColumnCount;
        }
        previousRight = std::max(previousRight, column.rightDip);
    }

    for (size_t index = 0; index < snapshot.visibleItems.size(); ++index)
    {
        const auto& item = snapshot.visibleItems[index];
        if (item.bounds.right <= item.bounds.left || item.bounds.bottom <= item.bounds.top)
        {
            ++integrity.negativeWidthCount;
        }
        if (item.index < snapshot.firstVisibleIndex || item.index > snapshot.lastVisibleIndex)
        {
            ++integrity.visibleRangeMismatchCount;
        }
        for (size_t other = index + 1u; other < snapshot.visibleItems.size(); ++other)
        {
            if (FolderViewColumnRectsOverlap(item.bounds, snapshot.visibleItems[other].bounds))
            {
                ++integrity.overlapCount;
            }
        }
    }

    return integrity;
}

struct FolderViewColumnAuditSample
{
    std::wstring name;
    FolderView::DebugColumnLayoutSnapshot snapshot;
    FolderViewColumnAuditIntegrity integrity;
};

[[nodiscard]] bool WaitForFolderViewColumnAuditItemCount(size_t expectedCount, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedCount)
        {
            return true;
        }
        std::this_thread::sleep_for(20ms);
    }

    return g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedCount;
}

[[nodiscard]] bool CaptureFolderViewColumnAuditSample(HWND folderView,
                                                      std::wstring_view name,
                                                      std::vector<FolderViewColumnAuditSample>& samples) noexcept
{
    if (! folderView || IsWindow(folderView) == FALSE)
    {
        return false;
    }

    PumpPendingMessages();
    static_cast<void>(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left));
    PumpPendingMessages();

    FolderViewColumnAuditSample sample{};
    sample.name = std::wstring(name);
    if (! g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, sample.snapshot))
    {
        return false;
    }
    sample.integrity = ComputeFolderViewColumnAuditIntegrity(sample.snapshot);
    samples.push_back(std::move(sample));
    return true;
}

[[nodiscard]] bool DriveFolderViewColumnAuditScroll(HWND folderView,
                                                    UINT request,
                                                    std::wstring_view name,
                                                    std::vector<FolderViewColumnAuditSample>& samples) noexcept
{
    if (! folderView || IsWindow(folderView) == FALSE)
    {
        return false;
    }

    SendMessageW(folderView, WM_HSCROLL, static_cast<WPARAM>(request), 0);
    PumpPendingMessages();
    return CaptureFolderViewColumnAuditSample(folderView, name, samples);
}

void AppendFolderViewColumnAuditSampleJson(std::wstring& out,
                                           const FolderViewColumnAuditSample& sample,
                                           std::wstring_view indent)
{
    const auto& snapshot = sample.snapshot;
    out.append(indent);
    out.append(L"{\n");
    out.append(indent);
    out.append(L"  \"name\": ");
    AppendFolderViewColumnJsonString(out, sample.name);
    out.append(L",\n");
    out.append(std::format(L"{}  \"horizontalOffsetDip\": {:.3f},\n", indent, snapshot.horizontalOffsetDip));
    out.append(std::format(L"{}  \"firstVisibleIndex\": {},\n", indent, snapshot.firstVisibleIndex));
    out.append(std::format(L"{}  \"lastVisibleIndex\": {},\n", indent, snapshot.lastVisibleIndex));
    out.append(std::format(L"{}  \"visibleColumnCount\": {},\n", indent, snapshot.columns.size()));
    out.append(indent);
    out.append(L"  \"integrity\": {\n");
    out.append(std::format(L"{}    \"overlapCount\": {},\n", indent, sample.integrity.overlapCount));
    out.append(std::format(L"{}    \"outOfOrderColumnCount\": {},\n", indent, sample.integrity.outOfOrderColumnCount));
    out.append(std::format(L"{}    \"negativeWidthCount\": {},\n", indent, sample.integrity.negativeWidthCount));
    out.append(std::format(L"{}    \"visibleRangeMismatchCount\": {},\n", indent, sample.integrity.visibleRangeMismatchCount));
    out.append(std::format(L"{}    \"hitTestSpacingFalsePositiveCount\": {}\n", indent, sample.integrity.hitTestSpacingFalsePositiveCount));
    out.append(indent);
    out.append(L"  }\n");
    out.append(indent);
    out.append(L"}");
}

void AppendFolderViewColumnAuditRecordJson(std::wstring& out,
                                           bool& firstRecord,
                                           std::wstring_view scenario,
                                           std::wstring_view variant,
                                           FolderView::DisplayMode mode,
                                           size_t itemCount,
                                           const std::vector<FolderViewColumnAuditSample>& samples)
{
    if (samples.empty())
    {
        return;
    }

    const auto& base = samples.front().snapshot;
    FolderViewColumnAuditIntegrity aggregate{};
    for (const auto& sample : samples)
    {
        aggregate.overlapCount += sample.integrity.overlapCount;
        aggregate.outOfOrderColumnCount += sample.integrity.outOfOrderColumnCount;
        aggregate.negativeWidthCount += sample.integrity.negativeWidthCount;
        aggregate.visibleRangeMismatchCount += sample.integrity.visibleRangeMismatchCount;
        aggregate.hitTestSpacingFalsePositiveCount += sample.integrity.hitTestSpacingFalsePositiveCount;
    }

    if (! firstRecord)
    {
        out.append(L",\n");
    }
    firstRecord = false;

    out.append(L"  {\n");
    out.append(L"    \"case\": \"folderView_column_widths_audit\",\n");
    out.append(L"    \"scenario\": ");
    AppendFolderViewColumnJsonString(out, scenario);
    out.append(L",\n");
    out.append(L"    \"variant\": ");
    AppendFolderViewColumnJsonString(out, variant);
    out.append(L",\n");
    out.append(L"    \"displayMode\": ");
    AppendFolderViewColumnJsonString(out, FolderViewColumnDisplayModeName(mode));
    out.append(L",\n");
    out.append(std::format(L"    \"itemCount\": {},\n", itemCount));
    out.append(std::format(L"    \"clientWidthPx\": {},\n", static_cast<int>(std::lround(base.clientWidthDip))));
    out.append(std::format(L"    \"clientHeightPx\": {},\n", static_cast<int>(std::lround(base.clientHeightDip))));
    out.append(std::format(L"    \"rowsPerColumn\": {},\n", base.rowsPerColumn));
    out.append(std::format(L"    \"columnCount\": {},\n", base.columns.size()));
    out.append(std::format(L"    \"contentWidthDip\": {:.3f},\n", base.contentWidthDip));
    out.append(std::format(L"    \"maxHorizontalOffsetDip\": {:.3f},\n", base.maxHorizontalOffsetDip));
    out.append(L"    \"integrity\": {\n");
    out.append(std::format(L"      \"overlapCount\": {},\n", aggregate.overlapCount));
    out.append(std::format(L"      \"outOfOrderColumnCount\": {},\n", aggregate.outOfOrderColumnCount));
    out.append(std::format(L"      \"negativeWidthCount\": {},\n", aggregate.negativeWidthCount));
    out.append(std::format(L"      \"visibleRangeMismatchCount\": {},\n", aggregate.visibleRangeMismatchCount));
    out.append(std::format(L"      \"hitTestSpacingFalsePositiveCount\": {}\n", aggregate.hitTestSpacingFalsePositiveCount));
    out.append(L"    },\n");
    out.append(L"    \"scrollSamples\": [\n");
    for (size_t i = 0; i < samples.size(); ++i)
    {
        AppendFolderViewColumnAuditSampleJson(out, samples[i], L"      ");
        out.append(i + 1u < samples.size() ? L",\n" : L"\n");
    }
    out.append(L"    ],\n");
    out.append(L"    \"columns\": [\n");
    for (size_t i = 0; i < base.columns.size(); ++i)
    {
        const auto& column = base.columns[i];
        out.append(L"      {\n");
        out.append(std::format(L"        \"index\": {},\n", i));
        out.append(std::format(L"        \"startIndex\": {},\n", column.startIndex));
        out.append(std::format(L"        \"itemCount\": {},\n", column.itemCount));
        out.append(std::format(L"        \"leftDip\": {:.3f},\n", column.leftDip));
        out.append(std::format(L"        \"widthDip\": {:.3f},\n", column.widthDip));
        out.append(std::format(L"        \"rightDip\": {:.3f},\n", column.rightDip));
        out.append(std::format(L"        \"widestObservedWidthDip\": {:.3f},\n", column.widestObservedWidthDip));
        out.append(L"        \"widestDisplayName\": ");
        AppendFolderViewColumnJsonString(out, column.widestDisplayName);
        out.append(L"\n");
        out.append(i + 1u < base.columns.size() ? L"      },\n" : L"      }\n");
    }
    out.append(L"    ]\n");
    out.append(L"  }");
}

[[nodiscard]] bool CreateFolderViewColumnBlockScenario(CaseState& state,
                                                       const std::filesystem::path& root,
                                                       int rowsPerColumn,
                                                       std::wstring_view scenario,
                                                       std::span<const size_t> longNameByColumn,
                                                       std::vector<std::wstring>& outFocusNames) noexcept
{
    state.Require(SelfTest::EnsureDirectory(root), std::format(L"Failed to create {} scenario root.", scenario));
    if (! state.failure.empty())
    {
        return false;
    }

    outFocusNames.clear();
    const int rows = std::max(2, rowsPerColumn);
    for (size_t column = 0; column < longNameByColumn.size(); ++column)
    {
        for (int row = 0; row < rows; ++row)
        {
            const size_t repeat = std::min<size_t>(longNameByColumn[column], 48u);
            const std::wstring stem =
                std::format(L"c{:02}_r{:03}_{}", column, row, RepeatForFolderViewColumnName(static_cast<wchar_t>(L'a' + column), repeat));
            const std::wstring name = stem + L".txt";
            state.Require(SelfTest::WriteTextFile(root / name, "content"), std::format(L"Failed to create {}.", name));
            outFocusNames.push_back(name);
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool CreateFolderViewColumnUnicodeScenario(CaseState& state,
                                                         const std::filesystem::path& root,
                                                         int rowsPerColumn,
                                                         std::vector<std::wstring>& outFocusNames) noexcept
{
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create unicode width scenario root.");
    outFocusNames.clear();
    const int rows = std::max(2, rowsPerColumn);
    const std::array<std::wstring_view, 6> names = {{
        L"c00_ascii spaced name [one].txt",
        L"c01_cjk_\u4E2D\u6587\u30C6\u30B9\u30C8.txt",
        L"c02_combining_e\u0301_e\u0301_e\u0301.txt",
        L"c03_rtl_\u05E9\u05DC\u05D5\u05DD.txt",
        L"c04_many.dots.in.the.name.log",
        L"c05_case_MiXeD_Ext.TxT",
    }};

    for (int i = 0; i < rows * 6; ++i)
    {
        const std::wstring base(names[static_cast<size_t>(i) % names.size()]);
        const std::wstring fileName = std::format(L"{:04}_{}", i, base);
        state.Require(SelfTest::WriteTextFile(root / fileName, "unicode"), std::format(L"Failed to create {}.", fileName));
        outFocusNames.push_back(fileName);
    }
    return state.failure.empty();
}

[[nodiscard]] bool CreateFolderViewColumnNearMaxScenario(CaseState& state,
                                                         const std::filesystem::path& root,
                                                         int rowsPerColumn,
                                                         std::vector<std::wstring>& outFocusNames) noexcept
{
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create near-max filename scenario root.");
    outFocusNames.clear();
    const int rows = std::max(2, rowsPerColumn);
    for (int i = 0; i < rows * 4; ++i)
    {
        const std::wstring fileName = std::format(L"nearmax_{:04}_{}.txt", i, RepeatForFolderViewColumnName(L'x', 32u + static_cast<size_t>(i % 8)));
        state.Require(SelfTest::WriteTextFile(root / fileName, "near max"), std::format(L"Failed to create near max file {}.", i));
        outFocusNames.push_back(fileName);
    }
    return state.failure.empty();
}

[[nodiscard]] bool CreateFolderViewColumnManyScenario(CaseState& state,
                                                      const std::filesystem::path& root,
                                                      int rowsPerColumn,
                                                      std::vector<std::wstring>& outFocusNames) noexcept
{
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create many-columns scenario root.");
    outFocusNames.clear();
    for (size_t i = 0; i < 48u; ++i)
    {
        const std::filesystem::path dir = root / std::format(L"d_{:03}", i);
        state.Require(SelfTest::EnsureDirectory(dir), std::format(L"Failed to create {}.", dir.filename().native()));
    }

    const size_t fileCount = std::max<size_t>(2400u, static_cast<size_t>(std::max(2, rowsPerColumn)) * 90u);
    for (size_t i = 0; i < fileCount; ++i)
    {
        const bool longColumn = ((i / static_cast<size_t>(std::max(2, rowsPerColumn))) % 10u) == 5u;
        const std::wstring name =
            std::format(L"f_{:04}_{}{}", i, longColumn ? RepeatForFolderViewColumnName(L'm', 48u) : L"short", (i % 3u == 0u ? L".log" : L".txt"));
        state.Require(SelfTest::WriteTextFile(root / name, "many"), std::format(L"Failed to create many scenario file {}.", i));
        outFocusNames.push_back(name);
    }
    return state.failure.empty();
}

[[nodiscard]] bool CreateFolderViewColumnTinyScenario(CaseState& state,
                                                      const std::filesystem::path& root,
                                                      std::wstring_view variant,
                                                      int rowsPerColumn,
                                                      std::vector<std::wstring>& outFocusNames) noexcept
{
    state.Require(SelfTest::EnsureDirectory(root), std::format(L"Failed to create tiny scenario variant {}.", variant));
    outFocusNames.clear();
    size_t fileCount = 0;
    if (variant == L"one")
    {
        fileCount = 1u;
    }
    else if (variant == L"rows_minus_one")
    {
        fileCount = static_cast<size_t>(std::max(0, rowsPerColumn - 1));
    }
    else if (variant == L"rows_exact")
    {
        fileCount = static_cast<size_t>(std::max(1, rowsPerColumn));
    }

    for (size_t i = 0; i < fileCount; ++i)
    {
        const std::wstring name = std::format(L"tiny_{:03}.txt", i);
        state.Require(SelfTest::WriteTextFile(root / name, "tiny"), std::format(L"Failed to create {}.", name));
        outFocusNames.push_back(name);
    }
    return state.failure.empty();
}

[[nodiscard]] int ResolveFolderViewColumnAuditRowsPerColumn(CaseState& state, const std::filesystem::path& probeRoot) noexcept
{
    using namespace std::chrono_literals;

    state.Require(SelfTest::EnsureDirectory(probeRoot), L"Failed to create FolderView column audit probe root.");
    for (size_t i = 0; i < 96u; ++i)
    {
        const std::wstring name = std::format(L"probe_{:03}.txt", i);
        state.Require(SelfTest::WriteTextFile(probeRoot / name, "probe"), std::format(L"Failed to create {}.", name));
    }
    if (! state.failure.empty())
    {
        return 0;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, probeRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, probeRoot, SelfTest::Scale(4000ms)), L"FolderView column audit probe path did not load.");
    state.Require(WaitForFolderViewColumnAuditItemCount(96u, SelfTest::Scale(5000ms)), L"FolderView column audit probe item count did not settle.");
    static_cast<void>(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left));

    FolderView::DebugColumnLayoutSnapshot snapshot{};
    state.Require(g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, snapshot), L"Failed to capture FolderView column audit probe.");
    return std::max(2, snapshot.rowsPerColumn);
}

[[nodiscard]] bool RunFolderViewColumnAuditForPath(CaseState& state,
                                                   HWND folderView,
                                                   const std::filesystem::path& root,
                                                   std::wstring_view scenario,
                                                   std::wstring_view variant,
                                                   size_t expectedItemCount,
                                                   std::span<const std::wstring> focusNames,
                                                   std::wstring& json,
                                                   bool& firstRecord) noexcept
{
    using namespace std::chrono_literals;

    constexpr std::array<FolderView::DisplayMode, 4> kDisplayModes = {{
        FolderView::DisplayMode::Brief,
        FolderView::DisplayMode::Detailed,
        FolderView::DisplayMode::ExtraDetailed,
        FolderView::DisplayMode::Thumbnails,
    }};

    for (const FolderView::DisplayMode mode : kDisplayModes)
    {
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, mode);
        g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
        state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                      std::format(L"FolderView column audit path did not load for {} / {}.", scenario, FolderViewColumnDisplayModeName(mode)));
        state.Require(WaitForFolderViewColumnAuditItemCount(expectedItemCount, SelfTest::Scale(6000ms)),
                      std::format(L"FolderView column audit item count mismatch for {} / {}. expected={} actual={}",
                                  scenario,
                                  FolderViewColumnDisplayModeName(mode),
                                  expectedItemCount,
                                  g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left)));
        state.Require(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left),
                      std::format(L"FolderView column audit warm rendering failed for {} / {}.", scenario, FolderViewColumnDisplayModeName(mode)));
        if (! state.failure.empty())
        {
            return false;
        }

        std::vector<FolderViewColumnAuditSample> samples;
        state.Require(DriveFolderViewColumnAuditScroll(folderView, SB_LEFT, L"left", samples), L"Failed to capture audit left sample.");
        state.Require(DriveFolderViewColumnAuditScroll(folderView, SB_LINERIGHT, L"line-right", samples), L"Failed to capture audit line-right sample.");
        state.Require(DriveFolderViewColumnAuditScroll(folderView, SB_PAGERIGHT, L"page-right", samples), L"Failed to capture audit page-right sample.");
        state.Require(DriveFolderViewColumnAuditScroll(folderView, SB_RIGHT, L"right", samples), L"Failed to capture audit right sample.");
        state.Require(DriveFolderViewColumnAuditScroll(folderView, SB_PAGELEFT, L"page-left", samples), L"Failed to capture audit page-left sample.");
        state.Require(DriveFolderViewColumnAuditScroll(folderView, SB_LEFT, L"home", samples), L"Failed to capture audit home sample.");

        if (! focusNames.empty())
        {
            const auto focusAndCapture = [&](size_t index, std::wstring_view sampleName) noexcept
            {
                const size_t clampedIndex = std::min(index, focusNames.size() - 1u);
                state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, focusNames[clampedIndex]),
                              std::format(L"Failed to focus {} for {}.", focusNames[clampedIndex], scenario));
                return CaptureFolderViewColumnAuditSample(folderView, sampleName, samples);
            };
            state.Require(focusAndCapture(0u, L"focus-first"), L"Failed to capture focus-first sample.");
            state.Require(focusAndCapture(focusNames.size() / 2u, L"focus-middle"), L"Failed to capture focus-middle sample.");
            state.Require(focusAndCapture(focusNames.size() - 1u, L"focus-last"), L"Failed to capture focus-last sample.");
            state.Require(focusAndCapture(0u, L"focus-first-return"), L"Failed to capture focus-first-return sample.");
        }

        if (! state.failure.empty())
        {
            return false;
        }

        AppendFolderViewColumnAuditRecordJson(json, firstRecord, scenario, variant, mode, expectedItemCount, samples);
        Debug::Perf::Emit(L"render.folderViewColumnAudit.us",
                          scenario,
                          0,
                          static_cast<uint64_t>(expectedItemCount),
                          static_cast<uint64_t>(samples.size()),
                          S_OK);
    }

    return true;
}

[[nodiscard]] bool TestFolderViewColumnWidthsAudit(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid for FolderView column audit.");
        return false;
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView && IsWindow(folderView), L"Left FolderView hwnd invalid for FolderView column audit.");
    if (! folderView || IsWindow(folderView) == FALSE)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for FolderView column audit.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"folder_column_widths_audit_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create FolderView column audit root.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"folderView_column_widths_audit custom trace: entered case");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const FolderView::DisplayMode displayBefore     = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restoreDisplay                       = wil::scope_exit([&] { g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore); });
    const FolderView::SortBy sortBefore             = g_folderWindow.GetSortBy(FolderWindow::Pane::Left);
    const FolderView::SortDirection directionBefore = g_folderWindow.GetSortDirection(FolderWindow::Pane::Left);
    const auto restoreSort                          = wil::scope_exit([&] { g_folderWindow.SetSort(FolderWindow::Pane::Left, sortBefore, directionBefore); });

    RECT originalWindowRect{};
    const bool haveOriginalWindowRect = GetWindowRect(mainWindow, &originalWindowRect) != FALSE;
    const auto restoreWindowRect      = wil::scope_exit([&]
    {
        if (haveOriginalWindowRect)
        {
            SetWindowPos(mainWindow,
                         nullptr,
                         originalWindowRect.left,
                         originalWindowRect.top,
                         originalWindowRect.right - originalWindowRect.left,
                         originalWindowRect.bottom - originalWindowRect.top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
    });

    g_folderWindow.SetSort(FolderWindow::Pane::Left, FolderView::SortBy::Name, FolderView::SortDirection::Ascending);
    const int rowsPerColumn = ResolveFolderViewColumnAuditRowsPerColumn(state, root / L"probe");
    state.Require(rowsPerColumn >= 2, std::format(L"FolderView column audit rowsPerColumn should be at least 2, got {}.", rowsPerColumn));
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring json = L"[\n";
    bool firstRecord  = true;
    std::vector<std::wstring> focusNames;

    const auto runBlockScenario = [&](std::wstring_view scenario, std::span<const size_t> widths) noexcept
    {
        const std::filesystem::path scenarioRoot = root / std::wstring(scenario);
        state.Require(CreateFolderViewColumnBlockScenario(state, scenarioRoot, rowsPerColumn, scenario, widths, focusNames),
                      std::format(L"Failed to create {}.", scenario));
        const size_t expectedCount = (static_cast<size_t>(rowsPerColumn) * widths.size());
        return RunFolderViewColumnAuditForPath(state, folderView, scenarioRoot, scenario, L"default", expectedCount, focusNames, json, firstRecord);
    };

    constexpr std::array<size_t, 5> kSecondColumnPoison = {{8u, 96u, 10u, 12u, 9u}};
    constexpr std::array<size_t, 6> kEveryOtherPoison   = {{8u, 88u, 9u, 90u, 10u, 92u}};
    constexpr std::array<size_t, 5> kDetailsPoison      = {{12u, 14u, 96u, 16u, 18u}};
    state.Require(runBlockScenario(L"column_poison_second_column", kSecondColumnPoison), L"column_poison_second_column audit failed.");
    state.Require(runBlockScenario(L"column_poison_every_other_column", kEveryOtherPoison), L"column_poison_every_other_column audit failed.");
    state.Require(runBlockScenario(L"details_metadata_poison", kDetailsPoison), L"details_metadata_poison audit failed.");

    focusNames.clear();
    const std::filesystem::path unicodeRoot = root / L"unicode_width_mess";
    state.Require(CreateFolderViewColumnUnicodeScenario(state, unicodeRoot, rowsPerColumn, focusNames), L"unicode_width_mess creation failed.");
    state.Require(RunFolderViewColumnAuditForPath(state,
                                                  folderView,
                                                  unicodeRoot,
                                                  L"unicode_width_mess",
                                                  L"default",
                                                  static_cast<size_t>(rowsPerColumn * 6),
                                                  focusNames,
                                                  json,
                                                  firstRecord),
                  L"unicode_width_mess audit failed.");

    focusNames.clear();
    const std::filesystem::path nearMaxRoot = root / L"n";
    state.Require(CreateFolderViewColumnNearMaxScenario(state, nearMaxRoot, rowsPerColumn, focusNames), L"near_max_filename_lengths creation failed.");
    state.Require(RunFolderViewColumnAuditForPath(state,
                                                  folderView,
                                                  nearMaxRoot,
                                                  L"near_max_filename_lengths",
                                                  L"default",
                                                  static_cast<size_t>(rowsPerColumn * 4),
                                                  focusNames,
                                                  json,
                                                  firstRecord),
                  L"near_max_filename_lengths audit failed.");

    constexpr std::array<std::wstring_view, 4> kTinyVariants = {{L"empty", L"one", L"rows_minus_one", L"rows_exact"}};
    for (std::wstring_view variant : kTinyVariants)
    {
        focusNames.clear();
        const std::filesystem::path tinyRoot = root / L"tiny_and_empty" / std::wstring(variant);
        state.Require(CreateFolderViewColumnTinyScenario(state, tinyRoot, variant, rowsPerColumn, focusNames),
                      std::format(L"tiny_and_empty {} creation failed.", variant));
        const size_t expectedCount = focusNames.size();
        state.Require(RunFolderViewColumnAuditForPath(state, folderView, tinyRoot, L"tiny_and_empty", variant, expectedCount, focusNames, json, firstRecord),
                      std::format(L"tiny_and_empty {} audit failed.", variant));
    }

    focusNames.clear();
    const std::filesystem::path manyRoot = root / L"many_columns_large_folder";
    state.Require(CreateFolderViewColumnManyScenario(state, manyRoot, rowsPerColumn, focusNames), L"many_columns_large_folder creation failed.");
    state.Require(RunFolderViewColumnAuditForPath(state,
                                                  folderView,
                                                  manyRoot,
                                                  L"many_columns_large_folder",
                                                  L"default",
                                                  focusNames.size() + 48u,
                                                  focusNames,
                                                  json,
                                                  firstRecord),
                  L"many_columns_large_folder audit failed.");

    focusNames.clear();
    const std::filesystem::path resizeRoot = root / L"resize_changes_rows_per_column";
    constexpr std::array<size_t, 6> kResizeWidths = {{10u, 72u, 12u, 14u, 80u, 16u}};
    state.Require(CreateFolderViewColumnBlockScenario(state, resizeRoot, rowsPerColumn, L"resize_changes_rows_per_column", kResizeWidths, focusNames),
                  L"resize_changes_rows_per_column creation failed.");
    state.Require(RunFolderViewColumnAuditForPath(state,
                                                  folderView,
                                                  resizeRoot,
                                                  L"resize_changes_rows_per_column",
                                                  L"before_resize",
                                                  static_cast<size_t>(rowsPerColumn) * kResizeWidths.size(),
                                                  focusNames,
                                                  json,
                                                  firstRecord),
                  L"resize_changes_rows_per_column before-resize audit failed.");
    if (haveOriginalWindowRect)
    {
        const int width  = originalWindowRect.right - originalWindowRect.left;
        const int height = static_cast<int>(std::max<LONG>(360L, (originalWindowRect.bottom - originalWindowRect.top) - 160L));
        SetWindowPos(mainWindow, nullptr, originalWindowRect.left, originalWindowRect.top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();
        state.Require(RunFolderViewColumnAuditForPath(state,
                                                      folderView,
                                                      resizeRoot,
                                                      L"resize_changes_rows_per_column",
                                                      L"after_height_shrink",
                                                      static_cast<size_t>(rowsPerColumn) * kResizeWidths.size(),
                                                      focusNames,
                                                      json,
                                                      firstRecord),
                      L"resize_changes_rows_per_column after-resize audit failed.");
    }

    json.append(L"\n]\n");
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"folderView_column_widths_audit_metrics.json");
    const bool artifactWriteOk               = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, json);
    const bool artifactExists                = ! artifactPath.empty() && SelfTest::PathExists(artifactPath);
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                               std::format(L"folderView_column_widths_audit artifact path='{}' writeOk={} existsAfterWrite={}",
                                           artifactPath.native(),
                                           artifactWriteOk,
                                           artifactExists));
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"folderView_column_widths_audit custom trace: leaving case");
    state.Require(artifactWriteOk && artifactExists, L"Failed to write FolderView column widths audit metrics artifact.");
    state.Require(json.find(L"column_poison_second_column") != std::wstring::npos, L"Audit artifact missing column_poison_second_column.");
    state.Require(json.find(L"column_poison_every_other_column") != std::wstring::npos, L"Audit artifact missing column_poison_every_other_column.");
    state.Require(json.find(L"unicode_width_mess") != std::wstring::npos, L"Audit artifact missing unicode_width_mess.");
    state.Require(json.find(L"near_max_filename_lengths") != std::wstring::npos, L"Audit artifact missing near_max_filename_lengths.");
    state.Require(json.find(L"tiny_and_empty") != std::wstring::npos, L"Audit artifact missing tiny_and_empty.");
    state.Require(json.find(L"many_columns_large_folder") != std::wstring::npos, L"Audit artifact missing many_columns_large_folder.");
    state.Require(json.find(L"resize_changes_rows_per_column") != std::wstring::npos, L"Audit artifact missing resize_changes_rows_per_column.");
    state.Require(json.find(L"details_metadata_poison") != std::wstring::npos, L"Audit artifact missing details_metadata_poison.");

    return state.failure.empty();
}

[[nodiscard]] bool AssertFolderViewVisibleColumnWidths(CaseState& state,
                                                       const FolderView::DebugColumnLayoutSnapshot& snapshot,
                                                       std::wstring_view scenario,
                                                       FolderView::DisplayMode mode) noexcept
{
    state.Require(snapshot.columns.size() >= 4u,
                  std::format(L"{} / {} should expose at least four columns, got {}.",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              snapshot.columns.size()));
    if (snapshot.columns.size() < 4u)
    {
        return false;
    }

    const float firstWidth  = snapshot.columns[0].widthDip;
    const float poisonWidth = snapshot.columns[1].widthDip;
    const float thirdWidth  = snapshot.columns[2].widthDip;
    const float fourthWidth = snapshot.columns[3].widthDip;

    state.Require(poisonWidth > firstWidth + 30.0f,
                  std::format(L"{} / {} should size column 2 from its own long names. widths: c1={:.1f} c2={:.1f}",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              firstWidth,
                              poisonWidth));
    state.Require(poisonWidth > thirdWidth + 20.0f,
                  std::format(L"{} / {} should not let column 2's long names widen column 3. widths: c2={:.1f} c3={:.1f}",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              poisonWidth,
                              thirdWidth));
    state.Require(fourthWidth < poisonWidth - 20.0f,
                  std::format(L"{} / {} should keep column 4 narrower than the poisoned column. widths: c2={:.1f} c4={:.1f}",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              poisonWidth,
                              fourthWidth));

    for (size_t index = 1; index < snapshot.columns.size(); ++index)
    {
        state.Require(snapshot.columns[index].leftDip >= snapshot.columns[index - 1u].rightDip,
                      std::format(L"{} / {} column {} should start after the previous column. prevRight={:.1f} left={:.1f}",
                                  scenario,
                                  FolderViewColumnDisplayModeName(mode),
                                  index,
                                  snapshot.columns[index - 1u].rightDip,
                                  snapshot.columns[index].leftDip));
    }

    return state.failure.empty();
}

[[nodiscard]] bool AssertFolderViewLineScrollRestoresFirstGap(CaseState& state,
                                                              HWND folderView,
                                                              std::wstring_view scenario,
                                                              FolderView::DisplayMode mode) noexcept
{
    SendMessageW(folderView, WM_HSCROLL, SB_LEFT, 0);
    PumpPendingMessages();

    FolderView::DebugColumnLayoutSnapshot leftSnapshot{};
    state.Require(g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, leftSnapshot),
                  std::format(L"{} / {} failed to capture left-edge scroll snapshot.",
                              scenario,
                              FolderViewColumnDisplayModeName(mode)));
    state.Require(leftSnapshot.columns.size() >= 2u,
                  std::format(L"{} / {} needs at least two columns for line-scroll gap coverage, got {}.",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              leftSnapshot.columns.size()));
    if (leftSnapshot.columns.size() < 2u)
    {
        return false;
    }

    const float firstColumnLeft  = leftSnapshot.columns[0].leftDip;
    const float secondColumnLeft = leftSnapshot.columns[1].leftDip;
    const float expectedRight    = std::min(secondColumnLeft, leftSnapshot.maxHorizontalOffsetDip);
    constexpr float kScrollOffsetToleranceDip = 2.0f;

    state.Require(leftSnapshot.horizontalOffsetDip <= kScrollOffsetToleranceDip,
                  std::format(L"{} / {} should start at the canonical left stop. offset={:.1f}",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              leftSnapshot.horizontalOffsetDip));
    state.Require(expectedRight > firstColumnLeft + kScrollOffsetToleranceDip,
                  std::format(L"{} / {} fixture should have a meaningful second-column scroll stop. firstLeft={:.1f} expectedRight={:.1f}",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              firstColumnLeft,
                              expectedRight));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(folderView, WM_HSCROLL, SB_LINERIGHT, 0);
    PumpPendingMessages();

    FolderView::DebugColumnLayoutSnapshot rightSnapshot{};
    state.Require(g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, rightSnapshot),
                  std::format(L"{} / {} failed to capture first right-scroll snapshot.",
                              scenario,
                              FolderViewColumnDisplayModeName(mode)));
    state.Require(std::abs(rightSnapshot.horizontalOffsetDip - expectedRight) <= kScrollOffsetToleranceDip,
                  std::format(L"{} / {} first right line-scroll should move past the first column, not only remove the left gutter. offset={:.1f} expected={:.1f} firstLeft={:.1f}",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              rightSnapshot.horizontalOffsetDip,
                              expectedRight,
                              firstColumnLeft));

    SendMessageW(folderView, WM_HSCROLL, SB_LINELEFT, 0);
    PumpPendingMessages();

    FolderView::DebugColumnLayoutSnapshot restoredSnapshot{};
    state.Require(g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, restoredSnapshot),
                  std::format(L"{} / {} failed to capture restored left-scroll snapshot.",
                              scenario,
                              FolderViewColumnDisplayModeName(mode)));
    state.Require(restoredSnapshot.horizontalOffsetDip <= kScrollOffsetToleranceDip,
                  std::format(L"{} / {} first left line-scroll from column 2 should restore the left gutter in one step. offset={:.1f}",
                              scenario,
                              FolderViewColumnDisplayModeName(mode),
                              restoredSnapshot.horizontalOffsetDip));

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewVisibleColumnWidths(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid for FolderView visible column width test.");
        return false;
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView && IsWindow(folderView), L"Left FolderView hwnd invalid for visible column width test.");
    if (! folderView || IsWindow(folderView) == FALSE)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for visible column width test.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"folder_visible_column_widths_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create visible column width test root.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const FolderView::DisplayMode displayBefore     = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restoreDisplay                       = wil::scope_exit([&] { g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore); });
    const FolderView::SortBy sortBefore             = g_folderWindow.GetSortBy(FolderWindow::Pane::Left);
    const FolderView::SortDirection directionBefore = g_folderWindow.GetSortDirection(FolderWindow::Pane::Left);
    const auto restoreSort                          = wil::scope_exit([&] { g_folderWindow.SetSort(FolderWindow::Pane::Left, sortBefore, directionBefore); });

    g_folderWindow.SetSort(FolderWindow::Pane::Left, FolderView::SortBy::Name, FolderView::SortDirection::Ascending);
    constexpr std::array<size_t, 5> kSecondColumnPoison = {{8u, 48u, 10u, 12u, 9u}};

    constexpr std::array<FolderView::DisplayMode, 4> kDisplayModes = {{
        FolderView::DisplayMode::Brief,
        FolderView::DisplayMode::Detailed,
        FolderView::DisplayMode::ExtraDetailed,
        FolderView::DisplayMode::Thumbnails,
    }};

    for (const FolderView::DisplayMode mode : kDisplayModes)
    {
        std::vector<std::wstring> focusNames;
        const std::wstring modeName          = FolderViewColumnDisplayModeName(mode);
        const std::filesystem::path modeRoot = root / (L"column_poison_second_column_" + modeName);

        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, mode);
        int rowsPerColumn = ResolveFolderViewColumnAuditRowsPerColumn(state, root / (L"probe_" + modeName));
        state.Require(rowsPerColumn >= 2, std::format(L"Visible column width rowsPerColumn should be at least 2 for {}, got {}.", modeName, rowsPerColumn));
        if (! state.failure.empty())
        {
            return false;
        }

        auto loadScenario = [&](int rows) noexcept -> bool
        {
            std::error_code removeError;
            std::filesystem::remove_all(modeRoot, removeError);
            focusNames.clear();
            state.Require(CreateFolderViewColumnBlockScenario(state, modeRoot, rows, L"column_poison_second_column", kSecondColumnPoison, focusNames),
                          std::format(L"Failed to create visible column width poison scenario for {}.", modeName));
            const size_t expectedCount = static_cast<size_t>(rows) * kSecondColumnPoison.size();
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, modeRoot);
            state.Require(WaitForPanePath(FolderWindow::Pane::Left, modeRoot, SelfTest::Scale(5000ms)),
                          std::format(L"Visible column width path did not load for {}.", modeName));
            state.Require(WaitForFolderViewColumnAuditItemCount(expectedCount, SelfTest::Scale(6000ms)),
                          std::format(L"Visible column width item count mismatch for {}. expected={} actual={}",
                                      modeName,
                                      expectedCount,
                                      g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left)));
            state.Require(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left),
                          std::format(L"Visible column width warm rendering failed for {}.", modeName));
            return state.failure.empty();
        };

        if (! loadScenario(rowsPerColumn))
        {
            return false;
        }

        FolderView::DebugColumnLayoutSnapshot snapshot{};
        state.Require(g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, snapshot),
                      std::format(L"Failed to capture visible column width sizing snapshot for {}.", modeName));
        if (! state.failure.empty())
        {
            return false;
        }
        if (snapshot.rowsPerColumn >= 2 && snapshot.rowsPerColumn != rowsPerColumn)
        {
            rowsPerColumn = snapshot.rowsPerColumn;
            if (! loadScenario(rowsPerColumn))
            {
                return false;
            }
        }

        SendMessageW(folderView, WM_HSCROLL, SB_LEFT, 0);
        PumpPendingMessages();
        snapshot = {};
        state.Require(g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, snapshot),
                      std::format(L"Failed to capture visible column width snapshot for {}.", modeName));
        state.Require(AssertFolderViewVisibleColumnWidths(state, snapshot, L"column_poison_second_column", mode),
                      std::format(L"Visible column width assertion failed for {}.", modeName));
        state.Require(AssertFolderViewLineScrollRestoresFirstGap(state, folderView, L"column_poison_second_column", mode),
                      std::format(L"Visible column first-gap scroll assertion failed for {}.", modeName));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool LoadLeftPaneThumbnailFixture(CaseState& state,
                                                const std::filesystem::path& root,
                                                std::initializer_list<std::wstring_view> expectedNames,
                                                std::chrono::milliseconds timeout) noexcept
{
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for thumbnail fixture.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, timeout), L"Thumbnail fixture path did not load.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, expectedNames, timeout), L"Thumbnail fixture contents did not load.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailValidImagesShellFail(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"thumbnail_valid_shell_fail_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create valid-image thumbnail fixture.");
    state.Require(TestWriteTinyBmpFile(root / L"PXL_20260503_121254627.PORTRAIT.bmp"), L"Failed to write valid BMP fixture.");
    state.Require(TestWriteTinyBmpFile(root / L"landscape_0001.BMP"), L"Failed to write valid uppercase BMP fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceShellFailureAllowWic);
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    state.Require(LoadLeftPaneThumbnailFixture(state,
                                               root,
                                               {L"PXL_20260503_121254627.PORTRAIT.bmp", L"landscape_0001.BMP"},
                                               SelfTest::Scale(4000ms)),
                  L"Valid-image shell-failure fixture did not load.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
    FolderWindow::PaneViewOptionsDebugSnapshot settled{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailCompletedCount > 0u;
                                            },
                                            SelfTest::Scale(6000ms),
                                            &settled),
                  L"Valid-image shell-failure thumbnail work did not settle.");
    state.Require(settled.thumbnailWicSuccessCount >= 1u,
                  std::format(L"Valid shell-failed images should display through WIC fallback; wicSuccess={}, fallback={}, completed={}.",
                              settled.thumbnailWicSuccessCount,
                              settled.thumbnailFallbackCount,
                              settled.thumbnailCompletedCount));
    state.Require(settled.thumbnailFallbackCount == 0u,
                  std::format(L"Valid shell-failed images should not fall back to file icons; fallback={}, completed={}.",
                              settled.thumbnailFallbackCount,
                              settled.thumbnailCompletedCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailAspectRatio(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"thumbnail_aspect_ratio_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create thumbnail aspect-ratio fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"wide_synthetic.jpg", "synthetic thumbnail source"), L"Failed to write synthetic thumbnail item.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceSyntheticWideSuccess);
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    state.Require(LoadLeftPaneThumbnailFixture(state, root, {L"wide_synthetic.jpg"}, SelfTest::Scale(4000ms)),
                  L"Aspect-ratio fixture did not load.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
    FolderWindow::PaneViewOptionsDebugSnapshot settled{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailVisibleApplyCount > 0u;
                                            },
                                            SelfTest::Scale(6000ms),
                                            &settled),
                  L"Synthetic wide thumbnail did not apply.");
    state.Require(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left), L"Failed to warm-render synthetic wide thumbnail.");
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, settled),
                  L"Could not capture thumbnail draw snapshot.");

    const float sourceRatio = settled.thumbnailLastDrawSourceHeightPx == 0u
                                  ? 0.0f
                                  : static_cast<float>(settled.thumbnailLastDrawSourceWidthPx) /
                                        static_cast<float>(settled.thumbnailLastDrawSourceHeightPx);
    const float drawWidth = settled.thumbnailLastDrawRectDip.right - settled.thumbnailLastDrawRectDip.left;
    const float drawHeight = settled.thumbnailLastDrawRectDip.bottom - settled.thumbnailLastDrawRectDip.top;
    const float drawRatio = drawHeight <= 0.0f ? 0.0f : drawWidth / drawHeight;

    state.Require(settled.thumbnailLastDrawSawThumbnail, L"Warm render should capture a thumbnail draw.");
    state.Require(settled.thumbnailLastDrawSourceWidthPx > settled.thumbnailLastDrawSourceHeightPx,
                  L"Synthetic wide provider should produce a non-square source thumbnail.");
    state.Require(std::abs(drawRatio - sourceRatio) <= 0.05f,
                  std::format(L"Thumbnail draw must preserve image ratio; source={}x{} ratio={:.3f}, draw={:.1f}x{:.1f} ratio={:.3f}.",
                              settled.thumbnailLastDrawSourceWidthPx,
                              settled.thumbnailLastDrawSourceHeightPx,
                              sourceRatio,
                              drawWidth,
                              drawHeight,
                              drawRatio));
    state.Require(drawWidth <= (settled.thumbnailLastDrawSlotRectDip.right - settled.thumbnailLastDrawSlotRectDip.left) + 0.1f &&
                      drawHeight <= (settled.thumbnailLastDrawSlotRectDip.bottom - settled.thumbnailLastDrawSlotRectDip.top) + 0.1f,
                  L"Fitted thumbnail draw rect should stay inside the thumbnail slot.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailBadFilesFallback(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"thumbnail_bad_files_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create bad-file thumbnail fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"truncated_header.jpg", "\xFF\xD8\xFF"), L"Failed to write truncated JPEG fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"text_saved_as.PNG", "this is not a png"), L"Failed to write bad PNG fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"ordinary.txt", "not an image"), L"Failed to write non-image fixture.");
    {
        std::ofstream zero(root / L"zero_byte.gif", std::ios::binary | std::ios::trunc);
        state.Require(zero.good(), L"Failed to write zero-byte GIF fixture.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceShellFailureAllowWic);
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    state.Require(LoadLeftPaneThumbnailFixture(state,
                                               root,
                                               {L"truncated_header.jpg", L"text_saved_as.PNG", L"ordinary.txt", L"zero_byte.gif"},
                                               SelfTest::Scale(4000ms)),
                  L"Bad-file thumbnail fixture did not load.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
    FolderWindow::PaneViewOptionsDebugSnapshot settled{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailCompletedCount > 0u;
                                            },
                                            SelfTest::Scale(6000ms),
                                            &settled),
                  L"Bad-file thumbnail work did not settle.");
    state.Require(settled.thumbnailDecodeFailureCount >= 1u,
                  std::format(L"Bad image-looking files should count WIC decode failures; decodeFailures={}, fallback={}, completed={}.",
                              settled.thumbnailDecodeFailureCount,
                              settled.thumbnailFallbackCount,
                              settled.thumbnailCompletedCount));
    state.Require(settled.thumbnailFallbackCount == settled.thumbnailCompletedCount,
                  L"Bad image-looking files should complete as icon fallback without stuck pending work.");
    state.Require(settled.thumbnailPendingCount == 0u, L"Bad-file thumbnail pipeline should not leave pending UI bitmap work.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailScrollStress(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"thumbnail_scroll_stress_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create thumbnail scroll-stress fixture.");

    std::wstring firstName;
    std::wstring middleName;
    std::wstring lastName;
    constexpr int kItemCount = 640;
    for (int i = 0; i < kItemCount; ++i)
    {
        const int bucket = i % 8;
        std::wstring fileName;
        if (bucket == 0)
        {
            fileName = std::format(L"PXL_20260503_{:06}.PORTRAIT.bmp", i);
            state.Require(TestWriteTinyBmpFile(root / fileName), std::format(L"Failed to write {}.", fileName));
        }
        else if (bucket == 1)
        {
            fileName = std::format(L"bad_header_{:06}.jpg", i);
            state.Require(SelfTest::WriteTextFile(root / fileName, "\xFF\xD8\xFF"), std::format(L"Failed to write {}.", fileName));
        }
        else if (bucket == 2)
        {
            fileName = std::format(L"long_filename_near_budget_{:06}_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa.txt", i);
            state.Require(SelfTest::WriteTextFile(root / fileName, "long thumbnail row"), std::format(L"Failed to write {}.", fileName));
        }
        else
        {
            fileName = std::format(L"mixed_{:06}.txt", i);
            state.Require(SelfTest::WriteTextFile(root / fileName, "thumbnail scroll stress"), std::format(L"Failed to write {}.", fileName));
        }

        if (i == 0)
        {
            firstName = fileName;
        }
        if (i == kItemCount / 2)
        {
            middleName = fileName;
        }
        if (i == kItemCount - 1)
        {
            lastName = fileName;
        }
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceFallback);
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    state.Require(LoadLeftPaneThumbnailFixture(state, root, {firstName, middleName, lastName}, SelfTest::Scale(8000ms)),
                  L"Scroll-stress thumbnail fixture did not load.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto started = std::chrono::steady_clock::now();
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
    FolderWindow::PaneViewOptionsDebugSnapshot settled{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailCompletedCount > 0u;
                                            },
                                            SelfTest::Scale(8000ms),
                                            &settled),
                  L"Initial scroll-stress thumbnail work did not settle.");
    state.Require(settled.thumbnailQueuedCount <= 256u,
                  std::format(L"Initial thumbnail queue should stay bounded; queued={}.", settled.thumbnailQueuedCount));

    for (const std::wstring_view focusName : {std::wstring_view{lastName}, std::wstring_view{firstName}, std::wstring_view{middleName}, std::wstring_view{lastName}})
    {
        state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, focusName),
                      std::format(L"Failed to focus '{}' during thumbnail scroll stress.", focusName));
        state.Require(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left), L"Warm render failed during thumbnail scroll stress.");
        state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                                [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                    return snapshot.thumbnailPendingCount == 0u;
                                                },
                                                SelfTest::Scale(6000ms),
                                                &settled),
                      L"Thumbnail scroll-stress work left pending bitmap creates.");
        state.Require(settled.thumbnailQueuedCount <= 256u,
                      std::format(L"Scrolled thumbnail queue should stay bounded; queued={}.", settled.thumbnailQueuedCount));
    }

    const auto elapsedUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count();
    const std::wstring perfArtifactText =
        std::format(L"{{\n"
                    L"  \"case\": \"folderView_thumbnail_scroll_stress\",\n"
                    L"  \"itemCount\": {},\n"
                    L"  \"metrics\": {{\n"
                    L"    \"thumbnailScrollStress.elapsedUs\": {},\n"
                    L"    \"thumbnails.queued\": {},\n"
                    L"    \"thumbnails.completed\": {},\n"
                    L"    \"thumbnails.fallback\": {},\n"
                    L"    \"thumbnails.pending\": {},\n"
                    L"    \"thumbnails.staleDrops\": {}\n"
                    L"  }}\n"
                    L"}}\n",
                    kItemCount,
                    elapsedUs,
                    settled.thumbnailQueuedCount,
                    settled.thumbnailCompletedCount,
                    settled.thumbnailFallbackCount,
                    settled.thumbnailPendingCount,
                    settled.thumbnailStaleDropCount);
    const std::filesystem::path artifactPath = SelfTest::GetPerfArtifactPath(L"folderView_thumbnail_scroll_stress_metrics.json");
    const bool artifactWriteOk = ! artifactPath.empty() && SelfTest::WriteTextFile(artifactPath, perfArtifactText);
    state.Require(artifactWriteOk && SelfTest::PathExists(artifactPath), L"Failed to write thumbnail scroll-stress perf artifact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailScrollRequeuesVisible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"thumbnail_scroll_requeue_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create thumbnail scroll-requeue fixture.");

    constexpr int kItemCount = 96;
    for (int i = 0; i < kItemCount; ++i)
    {
        const std::wstring fileName = std::format(L"PXL_20260503_scroll_visible_{:03}.jpg", i);
        state.Require(SelfTest::WriteTextFile(root / fileName, "thumbnail horizontal scroll visible requeue"),
                      std::format(L"Failed to write {}.", fileName));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    RECT originalWindowRect{};
    const bool haveOriginalWindowRect = GetWindowRect(mainWindow, &originalWindowRect) != FALSE;
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (haveOriginalWindowRect)
        {
            SetWindowPos(mainWindow,
                         nullptr,
                         originalWindowRect.left,
                         originalWindowRect.top,
                         originalWindowRect.right - originalWindowRect.left,
                         originalWindowRect.bottom - originalWindowRect.top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
    });

    if (haveOriginalWindowRect)
    {
        SetWindowPos(mainWindow, nullptr, originalWindowRect.left, originalWindowRect.top, 760, 520, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView && IsWindow(folderView) != FALSE, L"Left FolderView handle unavailable for thumbnail scroll-requeue test.");

    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceSyntheticSuccess);
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    state.Require(LoadLeftPaneThumbnailFixture(state,
                                               root,
                                               {L"PXL_20260503_scroll_visible_000.jpg", L"PXL_20260503_scroll_visible_095.jpg"},
                                               SelfTest::Scale(5000ms)),
                  L"Thumbnail scroll-requeue fixture did not load.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
    FolderWindow::PaneViewOptionsDebugSnapshot initial{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailVisibleItemCount > 0u &&
                                                       snapshot.thumbnailVisibleThumbnailCount == snapshot.thumbnailVisibleItemCount;
                                            },
                                            SelfTest::Scale(6000ms),
                                            &initial),
                  std::format(L"Initial visible thumbnail set did not settle; visible={} visibleThumb={} pending={}.",
                              initial.thumbnailVisibleItemCount,
                              initial.thumbnailVisibleThumbnailCount,
                              initial.thumbnailPendingCount));

    FolderView::DebugColumnLayoutSnapshot beforeScroll{};
    state.Require(g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, beforeScroll),
                  L"Failed to capture thumbnail scroll-requeue layout before scroll.");
    state.Require(beforeScroll.maxHorizontalOffsetDip > beforeScroll.horizontalOffsetDip + 1.0f,
                  std::format(L"Thumbnail scroll-requeue fixture should have horizontal overflow; offset={:.1f} max={:.1f}.",
                              beforeScroll.horizontalOffsetDip,
                              beforeScroll.maxHorizontalOffsetDip));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(folderView, WM_HSCROLL, static_cast<WPARAM>(SB_LINERIGHT), 0);
    PumpPendingMessages();

    FolderView::DebugColumnLayoutSnapshot afterScrollLayout{};
    bool scrolled = false;
    const auto scrollDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(2000ms);
    while (std::chrono::steady_clock::now() < scrollDeadline)
    {
        PumpPendingMessages();
        if (g_folderWindow.DebugGetPaneColumnLayoutSnapshot(FolderWindow::Pane::Left, afterScrollLayout) &&
            afterScrollLayout.horizontalOffsetDip > beforeScroll.horizontalOffsetDip + 1.0f)
        {
            scrolled = true;
            break;
        }
        std::this_thread::sleep_for(20ms);
    }
    state.Require(scrolled,
                  std::format(L"Horizontal thumbnail scroll should move visible range; before={:.1f} after={:.1f}.",
                              beforeScroll.horizontalOffsetDip,
                              afterScrollLayout.horizontalOffsetDip));
    SelfTest::AppendSelfTestTrace(std::format(L"thumbnail_scroll_requeues_visible: beforeOffset={:.1f} afterOffset={:.1f} max={:.1f} first={} last={} columns={} visibleItems={}",
                                              beforeScroll.horizontalOffsetDip,
                                              afterScrollLayout.horizontalOffsetDip,
                                              afterScrollLayout.maxHorizontalOffsetDip,
                                              afterScrollLayout.firstVisibleIndex,
                                              afterScrollLayout.lastVisibleIndex,
                                              afterScrollLayout.columns.size(),
                                              afterScrollLayout.visibleItems.size()));
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::PaneViewOptionsDebugSnapshot immediateAfterScroll{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, immediateAfterScroll),
                  L"Failed to capture immediate thumbnail snapshot after horizontal scroll.");
    SelfTest::AppendSelfTestTrace(std::format(L"thumbnail_scroll_requeues_visible: immediate visible={} visibleThumb={} totalThumb={} queued={} completed={} pending={} thumbnailsVisible={}",
                                              immediateAfterScroll.thumbnailVisibleItemCount,
                                              immediateAfterScroll.thumbnailVisibleThumbnailCount,
                                              immediateAfterScroll.thumbnailTotalThumbnailCount,
                                              immediateAfterScroll.thumbnailQueuedCount,
                                              immediateAfterScroll.thumbnailCompletedCount,
                                              immediateAfterScroll.thumbnailPendingCount,
                                              immediateAfterScroll.thumbnailsVisible ? 1 : 0));

    FolderWindow::PaneViewOptionsDebugSnapshot afterScroll{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailVisibleItemCount > 0u &&
                                                       snapshot.thumbnailVisibleThumbnailCount == snapshot.thumbnailVisibleItemCount;
                                            },
                                            SelfTest::Scale(4000ms),
                                            &afterScroll),
                  std::format(L"Horizontal scroll should requeue thumbnails for newly visible columns; visible={} visibleThumb={} totalThumb={} queued={} pending={} offset={:.1f} max={:.1f} first={} last={} columns={} layoutVisible={}.",
                              afterScroll.thumbnailVisibleItemCount,
                              afterScroll.thumbnailVisibleThumbnailCount,
                              afterScroll.thumbnailTotalThumbnailCount,
                              afterScroll.thumbnailQueuedCount,
                              afterScroll.thumbnailPendingCount,
                              afterScrollLayout.horizontalOffsetDip,
                              afterScrollLayout.maxHorizontalOffsetDip,
                              afterScrollLayout.firstVisibleIndex,
                              afterScrollLayout.lastVisibleIndex,
                              afterScrollLayout.columns.size(),
                              afterScrollLayout.visibleItems.size()));

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailResizeRequeuesVisible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"thumbnail_resize_requeue_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create thumbnail resize-requeue fixture.");

    constexpr int kItemCount = 64;
    for (int i = 0; i < kItemCount; ++i)
    {
        const std::wstring fileName = std::format(L"PXL_20260503_resize_visible_{:03}.jpg", i);
        state.Require(SelfTest::WriteTextFile(root / fileName, "thumbnail resize visible requeue"),
                      std::format(L"Failed to write {}.", fileName));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    RECT originalWindowRect{};
    const bool haveOriginalWindowRect = GetWindowRect(mainWindow, &originalWindowRect) != FALSE;
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (haveOriginalWindowRect)
        {
            SetWindowPos(mainWindow,
                         nullptr,
                         originalWindowRect.left,
                         originalWindowRect.top,
                         originalWindowRect.right - originalWindowRect.left,
                         originalWindowRect.bottom - originalWindowRect.top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
    });

    if (haveOriginalWindowRect)
    {
        SetWindowPos(mainWindow, nullptr, originalWindowRect.left, originalWindowRect.top, 760, 520, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();
    }

    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceSyntheticSuccess);
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    state.Require(LoadLeftPaneThumbnailFixture(state,
                                               root,
                                               {L"PXL_20260503_resize_visible_000.jpg", L"PXL_20260503_resize_visible_063.jpg"},
                                               SelfTest::Scale(5000ms)),
                  L"Thumbnail resize-requeue fixture did not load.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
    FolderWindow::PaneViewOptionsDebugSnapshot initial{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailVisibleItemCount > 0u &&
                                                       snapshot.thumbnailVisibleThumbnailCount == snapshot.thumbnailVisibleItemCount;
                                            },
                                            SelfTest::Scale(6000ms),
                                            &initial),
                  std::format(L"Initial narrow thumbnail set did not settle; visible={} visibleThumb={} pending={}.",
                              initial.thumbnailVisibleItemCount,
                              initial.thumbnailVisibleThumbnailCount,
                              initial.thumbnailPendingCount));

    const uint64_t initialVisible = initial.thumbnailVisibleItemCount;
    state.Require(initialVisible > 0u, L"Initial narrow thumbnail resize-requeue view should expose visible items.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (haveOriginalWindowRect)
    {
        SetWindowPos(mainWindow, nullptr, originalWindowRect.left, originalWindowRect.top, 1700, 720, SWP_NOZORDER | SWP_NOACTIVATE);
        PumpPendingMessages();
    }

    FolderWindow::PaneViewOptionsDebugSnapshot afterResize{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [initialVisible](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailVisibleItemCount > initialVisible;
                                            },
                                            SelfTest::Scale(3000ms),
                                            &afterResize),
                  std::format(L"Resize should reveal more thumbnail items; before={} after={}.",
                              initialVisible,
                              afterResize.thumbnailVisibleItemCount));

    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailVisibleItemCount > 0u &&
                                                       snapshot.thumbnailVisibleThumbnailCount == snapshot.thumbnailVisibleItemCount;
                                            },
                                            SelfTest::Scale(3000ms),
                                            &afterResize),
                  std::format(L"Resize should requeue thumbnails for newly visible columns; visible={} visibleThumb={} totalThumb={} queued={} pending={}.",
                              afterResize.thumbnailVisibleItemCount,
                              afterResize.thumbnailVisibleThumbnailCount,
                              afterResize.thumbnailTotalThumbnailCount,
                              afterResize.thumbnailQueuedCount,
                              afterResize.thumbnailPendingCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailSizeChangeWhilePending(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"thumbnail_size_change_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create thumbnail size-change fixture.");
    for (int i = 0; i < 96; ++i)
    {
        const std::wstring fileName = std::format(L"size_change_{:03}.jpg", i);
        state.Require(SelfTest::WriteTextFile(root / fileName, "thumbnail size change"), std::format(L"Failed to write {}.", fileName));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceFallback);
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    state.Require(LoadLeftPaneThumbnailFixture(state, root, {L"size_change_000.jpg", L"size_change_095.jpg"}, SelfTest::Scale(5000ms)),
                  L"Thumbnail size-change fixture did not load.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
    FolderWindow::PaneViewOptionsDebugSnapshot beforeChange{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailPendingCount == 0u && snapshot.thumbnailCompletedCount > 0u;
                                            },
                                            SelfTest::Scale(5000ms),
                                            &beforeChange),
                  L"Initial thumbnail size-change work did not settle.");

    const uint64_t cancelBefore = beforeChange.thumbnailCancelCount;
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_THUMBNAIL_SIZE_LARGE, 0), 0);
    PumpPendingMessages();

    FolderWindow::PaneViewOptionsDebugSnapshot afterChange{};
    state.Require(WaitForPaneThumbnailStats(FolderWindow::Pane::Left,
                                            [](const FolderWindow::PaneViewOptionsDebugSnapshot& snapshot) noexcept {
                                                return snapshot.thumbnailTargetDip >= 95.0f && snapshot.thumbnailTargetDip <= 97.0f &&
                                                       snapshot.thumbnailPendingCount == 0u;
                                            },
                                            SelfTest::Scale(1500ms),
                                            &afterChange),
                  std::format(L"Thumbnail size command should switch the pane to 96 DIP and settle; target={:.1f}, pending={}.",
                              afterChange.thumbnailTargetDip,
                              afterChange.thumbnailPendingCount));
    state.Require(afterChange.thumbnailCancelCount > cancelBefore,
                  L"Changing thumbnail size while thumbnails are active should cancel stale queued thumbnail work.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailReturnToNormalIconSize(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"thumbnail_return_normal_icon_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create thumbnail-to-normal icon fixture.");
    state.Require(SelfTest::WriteTextFile(root / L"mode_switch_probe.zzziconcache", "icon cache mode switch probe"),
                  L"Failed to write thumbnail-to-normal icon fixture item.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::DisplayMode displayBefore = g_folderWindow.GetDisplayMode(FolderWindow::Pane::Left);
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::Shell);
        g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, displayBefore);
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    IconCache::GetInstance().Clear();
    g_folderWindow.DebugSetThumbnailProviderMode(FolderWindow::Pane::Left, FolderView::DebugThumbnailProviderMode::ForceFallback);
    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
    state.Require(LoadLeftPaneThumbnailFixture(state, root, {L"mode_switch_probe.zzziconcache"}, SelfTest::Scale(5000ms)),
                  L"Thumbnail-to-normal icon fixture did not load.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left), L"Failed to warm thumbnail icon fallback rendering.");
    FolderWindow::PaneViewOptionsDebugSnapshot thumbnailMode{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, thumbnailMode), L"Failed to capture thumbnail-mode icon snapshot.");
    state.Require(thumbnailMode.iconLastDrawSawIcon, L"Thumbnail mode with forced fallback should draw a fallback icon.");
    SelfTest::AppendSelfTestTrace(std::format(L"thumbnail_return_normal_icon_size: thumbnail source={}x{} drawWidth={:.1f}",
                                              thumbnailMode.iconLastDrawSourceWidthPx,
                                              thumbnailMode.iconLastDrawSourceHeightPx,
                                              thumbnailMode.iconLastDrawRectDip.right - thumbnailMode.iconLastDrawRectDip.left));
    state.Require(thumbnailMode.iconLastDrawSourceWidthPx > 64u,
                  std::format(L"Test setup should seed a large thumbnail-mode icon source; source={}x{}.",
                              thumbnailMode.iconLastDrawSourceWidthPx,
                              thumbnailMode.iconLastDrawSourceHeightPx));
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
    state.Require(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left), L"Failed to warm normal-mode rendering after thumbnail mode.");

    FolderWindow::PaneViewOptionsDebugSnapshot normalMode{};
    state.Require(g_folderWindow.DebugGetPaneViewOptionsSnapshot(FolderWindow::Pane::Left, normalMode), L"Failed to capture normal-mode icon snapshot.");
    state.Require(! normalMode.thumbnailsVisible, L"Detailed mode should not report thumbnails visible after returning from thumbnail mode.");
    state.Require(normalMode.iconLastDrawSawIcon, L"Normal mode should draw a file icon after returning from thumbnail mode.");
    SelfTest::AppendSelfTestTrace(std::format(L"thumbnail_return_normal_icon_size: normal source={}x{} drawWidth={:.1f}",
                                              normalMode.iconLastDrawSourceWidthPx,
                                              normalMode.iconLastDrawSourceHeightPx,
                                              normalMode.iconLastDrawRectDip.right - normalMode.iconLastDrawRectDip.left));
    state.Require(normalMode.iconLastDrawSourceWidthPx <= 64u && normalMode.iconLastDrawSourceHeightPx <= 64u,
                  std::format(L"Normal mode should not reuse thumbnail-mode jumbo icon bitmap; source={}x{}, drawWidth={:.1f}.",
                              normalMode.iconLastDrawSourceWidthPx,
                              normalMode.iconLastDrawSourceHeightPx,
                              normalMode.iconLastDrawRectDip.right - normalMode.iconLastDrawRectDip.left));

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewThumbnailSortPopupSlider(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FolderWindow::Pane pane = FolderWindow::Pane::Left;
    const bool statusBefore = g_folderWindow.GetStatusBarVisible(pane);
    const auto restoreState = wil::scope_exit([&]
    {
        g_folderWindow.SetStatusBarVisible(pane, statusBefore);
        const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
        static_cast<void>(EnsureUiNotInMenuMode(uiThreadId, mainWindow, SelfTest::Scale(2000ms)));
    });

    g_folderWindow.SetStatusBarVisible(pane, true);
    FocusFolderViewPane(pane);
    PumpPendingMessages();

    const DWORD uiThreadId = GetWindowThreadProcessId(mainWindow, nullptr);
    std::atomic<bool> sawPopup{false};
    std::atomic<bool> sawThumbnailSlider{false};
    std::atomic<bool> sawFourStops{false};
    std::atomic<bool> closedPopup{false};

    std::jthread closer([&](std::stop_token stopToken) noexcept
    {
        const auto openDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        HWND popup = nullptr;
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < openDeadline)
        {
            popup = FindVisibleOwnedDxUiContextMenuWindow(mainWindow);
            if (popup != nullptr && IsWindow(popup) != FALSE)
            {
                sawPopup.store(true, std::memory_order_release);
                RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
                if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState))
                {
                    for (size_t itemIndex = 0u; itemIndex < popupState.itemTexts.size(); ++itemIndex)
                    {
                        const std::wstring& text = popupState.itemTexts[itemIndex];
                        if (text.find(L"Thumbnail size") != std::wstring::npos && itemIndex < popupState.itemKinds.size() &&
                            popupState.itemKinds[itemIndex] == RedSalamander::DxUi::MenuItemKind::Slider)
                        {
                            sawThumbnailSlider.store(true, std::memory_order_release);
                            if (itemIndex < popupState.sliderStopCounts.size() && popupState.sliderStopCounts[itemIndex] == 4u)
                            {
                                sawFourStops.store(true, std::memory_order_release);
                            }
                            break;
                        }
                    }
                }
                break;
            }

            std::this_thread::sleep_for(10ms);
        }

        if (popup != nullptr)
        {
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        }

        const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(2000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
        {
            GUITHREADINFO gti{};
            gti.cbSize = sizeof(gti);
            const bool hasGuiInfo = GetGUIThreadInfo(uiThreadId, &gti) != FALSE;
            const bool inMenuMode = hasGuiInfo && (gti.flags & GUI_INMENUMODE) != 0;
            const HWND currentPopup = FindVisibleOwnedDxUiContextMenuWindow(mainWindow);
            if (! inMenuMode && currentPopup == nullptr)
            {
                closedPopup.store(true, std::memory_order_release);
                return;
            }
            std::this_thread::sleep_for(20ms);
        }
    });

    state.Require(g_folderWindow.DebugClickPaneStatusBarSort(pane), L"Failed to post pane status-bar sort click for thumbnail slider validation.");
    PumpPendingMessages();
    closer.join();

    state.Require(sawPopup.load(std::memory_order_acquire), L"Pane status-bar sort popup did not open for thumbnail slider validation.");
    state.Require(closedPopup.load(std::memory_order_acquire), L"Pane status-bar sort popup did not close after thumbnail slider validation.");
    state.Require(sawThumbnailSlider.load(std::memory_order_acquire),
                  L"Pane bottom-right sort popup should expose a Thumbnail size slider row.");
    state.Require(sawFourStops.load(std::memory_order_acquire), L"Thumbnail size slider should expose four discrete stops.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFolderViewPerfLargeFolderBaseline(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"folder_perf_baseline_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create folder-perf baseline root.");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"folderView_perf_large_folder_baseline custom trace: entered case");

    constexpr size_t kDirectoryCount          = 24u;
    constexpr size_t kFilesPerExtension       = 48u;
    constexpr size_t kExtensionlessCount      = 24u;
    constexpr size_t kExpectedItemCount       = kDirectoryCount + (kFilesPerExtension * 4u) + kExtensionlessCount + 1u;
    const std::wstring focusTopDisplayName    = L"dir_000";
    const std::wstring focusBottomDisplayName = L"zz_tail_focus_target.txt";

    for (size_t i = 0; i < kDirectoryCount; ++i)
    {
        const std::filesystem::path dir = root / std::format(L"dir_{:03}", i);
        state.Require(SelfTest::EnsureDirectory(dir), std::format(L"Failed to create {}.", dir.filename().native()));
        state.Require(SelfTest::WriteTextFile(dir / L"marker.txt", "folder"), std::format(L"Failed to create marker for {}.", dir.filename().native()));
    }

    constexpr std::array<std::wstring_view, 4> kExtensions = {{L".txt", L".log", L".jsonl", L".md"}};
    for (std::wstring_view extension : kExtensions)
    {
        for (size_t i = 0; i < kFilesPerExtension; ++i)
        {
            const std::filesystem::path file = root / std::format(L"file_{:03}{}", i, extension);
            state.Require(SelfTest::WriteTextFile(file, "content"), std::format(L"Failed to create {}.", file.filename().native()));
        }
    }

    for (size_t i = 0; i < kExtensionlessCount; ++i)
    {
        const std::filesystem::path file = root / std::format(L"noext_{:03}", i);
        state.Require(SelfTest::WriteTextFile(file, "content"), std::format(L"Failed to create {}.", file.filename().native()));
    }

    state.Require(SelfTest::WriteTextFile(root / focusBottomDisplayName, "tail"), L"Failed to create focus target.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const FolderView::SortBy sortBefore             = g_folderWindow.GetSortBy(FolderWindow::Pane::Left);
    const FolderView::SortDirection directionBefore = g_folderWindow.GetSortDirection(FolderWindow::Pane::Left);
    const auto restoreSort                          = wil::scope_exit([&] { g_folderWindow.SetSort(FolderWindow::Pane::Left, sortBefore, directionBefore); });

    g_folderWindow.SetSort(FolderWindow::Pane::Left, FolderView::SortBy::Name, FolderView::SortDirection::Ascending);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumerationCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumerationCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    const auto waitForItemCount = [&](size_t expectedCount, std::chrono::milliseconds timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left) == expectedCount)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }
        return false;
    };

    const auto waitForFocusedItem = [&](std::wstring_view expectedDisplayName, std::chrono::milliseconds timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedDisplayName)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }
        return false;
    };

    const auto settleUiFor = [&](std::chrono::milliseconds duration) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + duration;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);
        }
    };

    const auto waitForBitmapIconCount = [&](size_t minimumCount, std::chrono::milliseconds timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetPaneBitmapIconCount(FolderWindow::Pane::Left) >= minimumCount)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }
        return false;
    };

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(4000ms)), L"Failed to set left pane path for folder-perf baseline.");
    state.Require(WaitForAtomicAtLeast(enumerationCount, 1u, SelfTest::Scale(5000ms)), L"Enumeration did not complete for folder-perf baseline.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left,
                                   {focusTopDisplayName, L"file_000.txt", L"file_000.jsonl", L"noext_000", focusBottomDisplayName},
                                   SelfTest::Scale(5000ms)),
                  L"Pane contents not ready for folder-perf baseline.");
    state.Require(waitForItemCount(kExpectedItemCount, SelfTest::Scale(5000ms)),
                  std::format(L"Expected {} pane items for folder-perf baseline, got {}.",
                              kExpectedItemCount,
                              g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left)));
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView && IsWindow(folderView), L"Left FolderView hwnd invalid for folder-perf baseline.");
    if (! folderView || IsWindow(folderView) == FALSE)
    {
        return false;
    }

    RECT folderViewRect{};
    GetClientRect(folderView, &folderViewRect);
    const int folderViewWidth  = static_cast<int>(std::max<LONG>(0L, folderViewRect.right - folderViewRect.left));
    const int folderViewHeight = static_cast<int>(std::max<LONG>(0L, folderViewRect.bottom - folderViewRect.top));
    state.Require(folderViewWidth > 0 && folderViewHeight > 0,
                  std::format(L"FolderView client rect should be non-zero for folder-perf baseline, got {}x{}.", folderViewWidth, folderViewHeight));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto scenarioPerfBefore            = g_folderWindow.DebugGetWarmPanePerfSnapshot(FolderWindow::Pane::Left);
    uint64_t driveFolderViewPaintInvocations = 0;

    const auto driveFolderViewPaint = [&]() noexcept
    {
        ++driveFolderViewPaintInvocations;
        const auto perfSnapshotBefore           = g_folderWindow.DebugGetWarmPanePerfSnapshot(FolderWindow::Pane::Left);
        const uint64_t warmRenderingCallsBefore = g_folderWindow.DebugGetWarmPaneRenderingCallCount(FolderWindow::Pane::Left);
        state.Require(g_folderWindow.DebugWarmPaneRendering(FolderWindow::Pane::Left),
                      L"Failed to warm FolderView rendering resources for folder-perf baseline.");
        const auto perfSnapshotAfter           = g_folderWindow.DebugGetWarmPanePerfSnapshot(FolderWindow::Pane::Left);
        const uint64_t warmRenderingCallsAfter = g_folderWindow.DebugGetWarmPaneRenderingCallCount(FolderWindow::Pane::Left);
        state.Require(warmRenderingCallsAfter == warmRenderingCallsBefore + 1u,
                      std::format(L"Expected FolderView warm-render helper call count to increment from {} to {}, got {}.",
                                  warmRenderingCallsBefore,
                                  warmRenderingCallsBefore + 1u,
                                  warmRenderingCallsAfter));
        state.Require(perfSnapshotAfter.renderCalls > perfSnapshotBefore.renderCalls,
                      L"Expected FolderView render call count to increase during folder-perf warmup.");
        if (! state.failure.empty())
        {
            return;
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(SelfTest::Scale(60ms));
        PumpPendingMessages();
        InvalidateRect(folderView, nullptr, FALSE);
        UpdateWindow(folderView);
        PumpPendingMessages();
    };

    FocusFolderViewPane(FolderWindow::Pane::Left);
    driveFolderViewPaint();
    settleUiFor(SelfTest::Scale(300ms));

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, focusBottomDisplayName),
                  L"Failed to focus bottom target for folder-perf baseline.");
    state.Require(waitForFocusedItem(focusBottomDisplayName, SelfTest::Scale(3000ms)), L"Focus did not move to bottom target for folder-perf baseline.");
    driveFolderViewPaint();
    settleUiFor(SelfTest::Scale(300ms));

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, focusTopDisplayName),
                  L"Failed to focus top target for folder-perf baseline.");
    state.Require(waitForFocusedItem(focusTopDisplayName, SelfTest::Scale(3000ms)), L"Focus did not move to top target for folder-perf baseline.");
    driveFolderViewPaint();
    settleUiFor(SelfTest::Scale(200ms));

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, focusBottomDisplayName),
                  L"Failed to refocus bottom target for folder-perf baseline.");
    state.Require(waitForFocusedItem(focusBottomDisplayName, SelfTest::Scale(3000ms)), L"Focus did not return to bottom target for folder-perf baseline.");
    driveFolderViewPaint();
    settleUiFor(SelfTest::Scale(400ms));

    const uint64_t refreshBefore = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    g_folderWindow.CommandRefresh(FolderWindow::Pane::Left);
    state.Require(WaitForAtomicAtLeast(enumerationCount, 2u, SelfTest::Scale(5000ms)), L"Warm refresh enumeration did not complete for folder-perf baseline.");
    state.Require(g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left) >= refreshBefore + 1u,
                  L"Folder-perf baseline refresh did not increment force-refresh count.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {focusTopDisplayName, focusBottomDisplayName}, SelfTest::Scale(5000ms)),
                  L"Pane contents did not settle after warm refresh for folder-perf baseline.");
    driveFolderViewPaint();
    settleUiFor(SelfTest::Scale(500ms));

    const bool bitmapIconsResolved       = waitForBitmapIconCount(1u, SelfTest::Scale(2000ms));
    const auto scenarioPerfAfter         = g_folderWindow.DebugGetWarmPanePerfSnapshot(FolderWindow::Pane::Left);
    const uint64_t warmRenderingCalls    = scenarioPerfAfter.warmRenderingCalls - scenarioPerfBefore.warmRenderingCalls;
    const uint64_t deferredInitCalls     = scenarioPerfAfter.deferredInitCalls - scenarioPerfBefore.deferredInitCalls;
    const uint64_t renderCalls           = scenarioPerfAfter.renderCalls - scenarioPerfBefore.renderCalls;
    const uint64_t queueIconLoadingCalls = scenarioPerfAfter.queueIconLoadingCalls - scenarioPerfBefore.queueIconLoadingCalls;
    const uint64_t processIconQueueCalls = scenarioPerfAfter.processIconQueueCalls - scenarioPerfBefore.processIconQueueCalls;
    const uint64_t batchIconUpdateCalls  = scenarioPerfAfter.batchIconUpdateCalls - scenarioPerfBefore.batchIconUpdateCalls;
    const size_t bitmapIconCount         = g_folderWindow.DebugGetPaneBitmapIconCount(FolderWindow::Pane::Left);
    state.Require(
        bitmapIconsResolved,
        std::format(L"Folder-perf warmup should resolve stock icons into pane bitmaps; itemCount={} bitmapIconCount={} queued={} processed={} batchUpdates={}.",
                    kExpectedItemCount,
                    bitmapIconCount,
                    queueIconLoadingCalls,
                    processIconQueueCalls,
                    batchIconUpdateCalls));

    const std::wstring folderPerfArtifactText          = std::format(L"{{\n"
                                                                     L"  \"case\": \"folderView_perf_large_folder_baseline\",\n"
                                                                     L"  \"root\": \"{}\",\n"
                                                                     L"  \"itemCount\": {},\n"
                                                                     L"  \"bitmapIconCount\": {},\n"
                                                                     L"  \"clientWidthPx\": {},\n"
                                                                     L"  \"clientHeightPx\": {},\n"
                                                                     L"  \"enumerationCount\": {},\n"
                                                                     L"  \"warmRenderInvocations\": {},\n"
                                                                     L"  \"metrics\": {{\n"
                                                                     L"    \"warmRenderingCalls\": {},\n"
                                                                     L"    \"deferredInitCalls\": {},\n"
                                                                     L"    \"renderCalls\": {},\n"
                                                                     L"    \"queueIconLoadingCalls\": {},\n"
                                                                     L"    \"processIconQueueCalls\": {},\n"
                                                                     L"    \"batchIconUpdateCalls\": {}\n"
                                                                     L"  }}\n"
                                                                     L"}}\n",
                                                                     root.native(),
                                                                     kExpectedItemCount,
                                                                     g_folderWindow.DebugGetPaneBitmapIconCount(FolderWindow::Pane::Left),
                                                                     folderViewWidth,
                                                                     folderViewHeight,
                                                                     enumerationCount.load(std::memory_order_acquire),
                                                                     driveFolderViewPaintInvocations,
                                                                     warmRenderingCalls,
                                                                     deferredInitCalls,
                                                                     renderCalls,
                                                                     queueIconLoadingCalls,
                                                                     processIconQueueCalls,
                                                                     batchIconUpdateCalls);
    const std::filesystem::path folderPerfArtifactPath = SelfTest::GetPerfArtifactPath(L"folderView_perf_large_folder_baseline_metrics.json");
    const bool folderPerfArtifactWriteOk = ! folderPerfArtifactPath.empty() && SelfTest::WriteTextFile(folderPerfArtifactPath, folderPerfArtifactText);
    const bool folderPerfArtifactExists  = ! folderPerfArtifactPath.empty() && SelfTest::PathExists(folderPerfArtifactPath);
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                               std::format(L"folderView_perf_large_folder_baseline artifact path='{}' writeOk={} existsAfterWrite={}",
                                           folderPerfArtifactPath.native(),
                                           folderPerfArtifactWriteOk,
                                           folderPerfArtifactExists));
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"folderView_perf_large_folder_baseline custom trace: leaving case");
    state.Require(folderPerfArtifactWriteOk && folderPerfArtifactExists, L"Failed to write folder-perf baseline metrics artifact.");

    return state.failure.empty();
}

[[nodiscard]] bool TestToggleHiddenAndSystemFiles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"hidden_system_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create hidden/system test root.");

    const std::filesystem::path normal = root / L"normal.txt";
    const std::filesystem::path hidden = root / L"hidden.txt";
    const std::filesystem::path system = root / L"system.txt";
    state.Require(SelfTest::WriteTextFile(normal, "n"), L"Failed to create normal.txt.");
    state.Require(SelfTest::WriteTextFile(hidden, "h"), L"Failed to create hidden.txt.");
    state.Require(SelfTest::WriteTextFile(system, "s"), L"Failed to create system.txt.");

    const auto applyAttrs = [&](const std::filesystem::path& path, DWORD attrs, std::wstring_view label) noexcept
    {
        if (::SetFileAttributesW(path.c_str(), attrs) == FALSE)
        {
            const DWORD err = GetLastError();
            state.Require(false, std::format(L"Failed to SetFileAttributesW for {} (err={}).", label, err));
        }
    };

    applyAttrs(hidden, FILE_ATTRIBUTE_HIDDEN, L"hidden.txt");
    applyAttrs(system, FILE_ATTRIBUTE_SYSTEM, L"system.txt");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    const bool showHiddenBefore = g_folderWindow.GetShowHiddenFiles();
    const bool showSystemBefore = g_folderWindow.GetShowSystemFiles();
    const auto restoreFlags     = wil::scope_exit([&]
    {
        g_folderWindow.SetShowHiddenFiles(showHiddenBefore);
        g_folderWindow.SetShowSystemFiles(showSystemBefore);
    });

    g_folderWindow.SetShowHiddenFiles(true);
    g_folderWindow.SetShowSystemFiles(true);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for hidden/system test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"normal.txt", L"hidden.txt", L"system.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for hidden/system test.");

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"normal.txt"), L"Expected normal.txt visible.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"hidden.txt"), L"Expected hidden.txt visible when enabled.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"), L"Expected system.txt visible when enabled.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_HIDDEN_FILES, 0), 0);
    state.Require(WaitForAtomicAtLeast(enumCount, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Enumeration did not refresh after hidden toggle.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"hidden.txt"), L"Expected hidden.txt to be hidden when disabled.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"),
                  L"system.txt should remain visible when only hidden is disabled.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_SYSTEM_FILES, 0), 0);
    state.Require(WaitForAtomicAtLeast(enumCount, 3u, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Enumeration did not refresh after system toggle.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"), L"Expected system.txt to be hidden when disabled.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_HIDDEN_FILES, 0), 0);
    state.Require(WaitForAtomicAtLeast(enumCount, 4u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh after re-enabling hidden.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"hidden.txt"), L"Expected hidden.txt visible after re-enabling hidden.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"),
                  L"system.txt should remain hidden while system is disabled.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_PANE_SYSTEM_FILES, 0), 0);
    state.Require(WaitForAtomicAtLeast(enumCount, 5u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh after re-enabling system.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"system.txt"), L"Expected system.txt visible after re-enabling system.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectSameExtensionCommands(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"same_extension_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create same-extension test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "b"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"c.log", "c"), L"Failed to create c.log.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for same-extension test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.txt", L"c.log"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for same-extension test.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Failed to focus a.txt.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"c.log"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log selected before select-same-extension.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_SAME_EXTENSION, 0), 0);

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected after select-same-extension.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt selected after select-same-extension.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log to remain selected after select-same-extension.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u, L"Expected 3 selected items after select-same-extension.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_UNSELECT_SAME_EXTENSION, 0), 0);

    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected after unselect-same-extension.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt unselected after unselect-same-extension.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log to remain selected after unselect-same-extension.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected 1 selected item after unselect-same-extension.");

    return state.failure.empty();
}

[[nodiscard]] bool TestInvertSelectionCommand(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"invert_selection_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create invert-selection test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "b"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"c.log", "c"), L"Failed to create c.log.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for invert-selection test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.txt", L"c.log"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for invert-selection test.");

    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Failed to focus a.txt.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"c.log"; }, true);

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log selected before invert.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected before invert.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt unselected before invert.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected 1 selected item before invert.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_INVERT, 0), 0);

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected after invert.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt selected after invert.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log unselected after invert.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items after invert.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_INVERT, 0), 0);

    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected after second invert.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.txt"), L"Expected b.txt unselected after second invert.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.log"), L"Expected c.log selected after second invert.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected 1 selected item after second invert.");

    return state.failure.empty();
}

[[nodiscard]] bool TestHideNamesCommands(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"hide_names_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create hide-names test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    const bool showHiddenBefore = g_folderWindow.GetShowHiddenFiles();
    const bool showSystemBefore = g_folderWindow.GetShowSystemFiles();
    const auto restoreViewFlags = wil::scope_exit([&]
    {
        g_folderWindow.SetShowHiddenFiles(showHiddenBefore);
        g_folderWindow.SetShowSystemFiles(showSystemBefore);
    });

    g_folderWindow.SetShowHiddenFiles(true);
    g_folderWindow.SetShowSystemFiles(true);

    std::atomic<uint32_t> enumCount{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumCount.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for hide-names test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for hide-names test.");

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt visible before hiding names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log visible before hiding names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt visible before hiding names.");

    // Hide Selected Names
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_HIDE_SELECTED_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after Hide Selected Names.");
    }

    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be hidden after Hide Selected Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should remain visible after Hide Selected Names.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be hidden after Hide Selected Names.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter indicator active after hiding names.");

    // Show Hidden Names (clears hidden set)
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after Show Hidden Names.");
    }

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after Show Hidden Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should be visible after Show Hidden Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after Show Hidden Names.");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter indicator should be inactive after Show Hidden Names.");

    // Hide Unselected Names (requires at least one selected item)
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"b.log"; }, true);
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_HIDE_UNSELECTED_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after Hide Unselected Names.");
    }

    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be hidden after Hide Unselected Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should remain visible after Hide Unselected Names.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be hidden after Hide Unselected Names.");

    // Clear hidden names again
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after clearing hidden names.");
    }

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after clearing hidden names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should be visible after clearing hidden names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after clearing hidden names.");

    // Verify that Show Hidden Names does not clear the pane filter.
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);

        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after OK.");

        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after applying pane filter.");
    }

    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter expected active after applying a pane filter mask.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log expected to be filtered out by pane filter.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt"; }, true);
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_HIDE_SELECTED_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after hiding a.txt while pane filter is active.");
    }

    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be hidden while pane filter is active.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should remain visible while pane filter is active.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"),
                  L"b.log should remain filtered out while pane filter is active.");

    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, 0), 0);
        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after showing hidden names while pane filter is active.");
    }

    const FolderView::NameFilterState filterState = g_folderWindow.DebugGetNameFilterState(FolderWindow::Pane::Left);
    state.Require(filterState.enabled, L"Pane filter expected to remain enabled after Show Hidden Names.");
    state.Require(filterState.text == L"*.txt", L"Pane filter text expected to remain '*.txt' after Show Hidden Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after Show Hidden Names.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after Show Hidden Names.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should remain filtered out after Show Hidden Names.");

    // Clean up: disable pane filter.
    {
        const uint32_t before = enumCount.load(std::memory_order_acquire);

        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, false, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not reopen for disable.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after disabling.");

        state.Require(WaitForAtomicAtLeast(enumCount, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Enumeration did not refresh after disabling pane filter.");
    }

    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter indicator inactive after disabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"a.txt should be visible after disabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"b.log should be visible after disabling pane filter.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"c.txt should be visible after disabling pane filter.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionSaveRestoreCommands(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_save_restore_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-save/restore test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePath                                 = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Right);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for left pane in selection-save/restore test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for right pane in selection-save/restore test.");

    std::atomic<uint32_t> enumLeft{0};
    std::atomic<uint32_t> enumRight{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumLeft.fetch_add(1u, std::memory_order_release);
        }
    });
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumRight.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallbacks = wil::scope_exit([&]
    {
        g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {});
        g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Right, {});
    });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for selection-save/restore test.");
    state.Require(WaitForAtomicAtLeast(enumLeft, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Left pane enumeration did not complete for selection-save/restore test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Left pane contents not ready for selection-save/restore test.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set right pane path for selection-save/restore test.");
    state.Require(WaitForAtomicAtLeast(enumRight, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Right pane enumeration did not complete for selection-save/restore test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Right, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Right pane contents not ready for selection-save/restore test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected before save-selection.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt selected before save-selection.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items before save-selection.");

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SAVE_SELECTION, 0), 0);
    state.Require(g_folderWindow.HasSavedSelection(), L"Expected a saved selection after Save Selection command.");

    {
        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);

        state.Require(! clipText.empty(), L"Expected Save Selection to place Unicode text in the clipboard.");
        state.Require(clipText.find(L"a.txt") != std::wstring::npos, L"Expected clipboard text to contain a.txt.");
        state.Require(clipText.find(L"c.txt") != std::wstring::npos, L"Expected clipboard text to contain c.txt.");
    }

    // Rename c.txt to change casing before restore-selection, so restore must fall back to case-insensitive match.
    {
        const std::filesystem::path tempName = root / L"c_case_tmp.txt";
        std::filesystem::rename(root / L"c.txt", tempName, ec);
        state.Require(! ec, L"Failed to rename c.txt for case-insensitive restore-selection scenario.");

        std::filesystem::rename(tempName, root / L"C.TXT", ec);
        state.Require(! ec, L"Failed to rename temp file to C.TXT for case-insensitive restore-selection scenario.");

        const uint32_t before = enumRight.load(std::memory_order_acquire);
        g_folderWindow.CommandRefresh(FolderWindow::Pane::Right);
        state.Require(WaitForAtomicAtLeast(enumRight, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Right pane did not refresh after renaming c.txt for case-insensitive restore-selection scenario.");
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Right, [](std::wstring_view) noexcept { return false; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == 0u, L"Expected no selection in right pane before restore-selection.");

    FocusFolderViewPane(FolderWindow::Pane::Right);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_RESTORE, 0), 0);

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"a.txt"), L"Expected a.txt selected after restore-selection.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"C.TXT"), L"Expected C.TXT selected after restore-selection.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"b.log"), L"Expected b.log not selected after restore-selection.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == 2u, L"Expected 2 selected items after restore-selection.");

    std::filesystem::remove(root / L"a.txt", ec);
    state.Require(! std::filesystem::exists(root / L"a.txt"), L"Failed to remove a.txt for restore-selection missing-item scenario.");

    {
        const uint32_t before = enumRight.load(std::memory_order_acquire);
        g_folderWindow.CommandRefresh(FolderWindow::Pane::Right);
        state.Require(WaitForAtomicAtLeast(enumRight, before + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Right pane did not refresh after deleting a.txt.");
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Right, [](std::wstring_view) noexcept { return false; }, true);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_RESTORE, 0), 0);

    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"a.txt"),
                  L"Expected a.txt not selected after restore-selection when missing.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Right, L"C.TXT"),
                  L"Expected C.TXT selected after restore-selection when a.txt is missing.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Right) == 1u,
                  L"Expected 1 selected item after restore-selection when a.txt is missing.");

    return state.failure.empty();
}

[[nodiscard]] bool TestCopyTextCommands(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"copy_text_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create copy-text test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for copy-text test.");

    std::atomic<uint32_t> enumLeft{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, root))
        {
            enumLeft.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for copy-text test.");
    state.Require(WaitForAtomicAtLeast(enumLeft, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Left pane enumeration did not complete for copy-text test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Left pane contents not ready for copy-text test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"a.txt" || name == L"c.txt"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected in copy-text test.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt selected in copy-text test.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u, L"Expected 2 selected items in copy-text test.");

    const std::wstring rootText = root.native();
    const std::wstring aFull    = (root / L"a.txt").native();
    const std::wstring cFull    = (root / L"c.txt").native();
    const std::wstring aUnc     = BuildLocalAdministrativeUncPath(aFull);
    const std::wstring cUnc     = BuildLocalAdministrativeUncPath(cFull);

    {
        ClearClipboardContents(mainWindow);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT, 0), 0);

        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);
        state.Require(! clipText.empty(), L"Expected Copy Path + Name as Text to place Unicode text in the clipboard.");
        state.Require(clipText.find(aFull) != std::wstring::npos, L"Expected clipboard text to contain full path for a.txt.");
        state.Require(clipText.find(cFull) != std::wstring::npos, L"Expected clipboard text to contain full path for c.txt.");
    }

    {
        ClearClipboardContents(mainWindow);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_NAME_AS_TEXT, 0), 0);

        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);
        state.Require(! clipText.empty(), L"Expected Copy Name as Text to place Unicode text in the clipboard.");
        state.Require(clipText.find(L"a.txt") != std::wstring::npos, L"Expected clipboard text to contain a.txt.");
        state.Require(clipText.find(L"c.txt") != std::wstring::npos, L"Expected clipboard text to contain c.txt.");
        state.Require(clipText.find(aFull) == std::wstring::npos, L"Expected name-only clipboard text to exclude the full path for a.txt.");
        state.Require(clipText.find(cFull) == std::wstring::npos, L"Expected name-only clipboard text to exclude the full path for c.txt.");
    }

    {
        ClearClipboardContents(mainWindow);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_PATH_AS_TEXT, 0), 0);

        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);
        state.Require(! clipText.empty(), L"Expected Copy Path as Text to place Unicode text in the clipboard.");
        state.Require(clipText.find(rootText) != std::wstring::npos, L"Expected clipboard text to contain the containing folder path.");
        state.Require(clipText.find(aFull) == std::wstring::npos, L"Expected path-only clipboard text to exclude the full path for a.txt.");
        state.Require(clipText.find(cFull) == std::wstring::npos, L"Expected path-only clipboard text to exclude the full path for c.txt.");
    }

    {
        ClearClipboardContents(mainWindow);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_COPY_PATH_AND_FILE_NAME, 0), 0);

        const std::wstring clipText = ReadClipboardUnicodeText(mainWindow);
        state.Require(! clipText.empty(), L"Expected Copy UNC Path + Name as Text to place Unicode text in the clipboard.");
        state.Require(! aUnc.empty(), L"Expected a local-machine UNC path for a.txt in the UNC clipboard test.");
        state.Require(! cUnc.empty(), L"Expected a local-machine UNC path for c.txt in the UNC clipboard test.");
        state.Require(clipText.find(aUnc) != std::wstring::npos, L"Expected clipboard text to contain the UNC full path for a.txt.");
        state.Require(clipText.find(cUnc) != std::wstring::npos, L"Expected clipboard text to contain the UNC full path for c.txt.");
    }

    return state.failure.empty();
}

} // namespace (tests)

void RunViewCommandsCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"menu_load_selection_links_restore", [=](CaseState& state) noexcept {
        return TestLoadSelectionMenuLinksToRestoreSelection(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"menu_copy_text_group_contract", [=](CaseState& state) noexcept { return TestCopyTextCommandsMenuContract(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"dispatch_smoke_all_commands", [=](CaseState& state) noexcept { return TestDispatchAllCommandsSmoke(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"modeless_window_ownership", [=](CaseState& state) noexcept { return TestModelessWindowOwnership(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_fullScreen", [=](CaseState& state) noexcept { return TestFullScreenToggle(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_fullScreen_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestFullScreenKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_openDriveMenus", [=](CaseState& state) noexcept { return TestDriveMenuCommands(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_openDriveMenus_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestAppOpenDriveMenusKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_about_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestAppAboutKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_showShortcuts_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestAppShowShortcutsKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_compare_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestAppCompareKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_preferences_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestAppPreferencesKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_managePlugins_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestAppManagePluginsKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_focusAddressBar_tab_traversal", [=](CaseState& state) noexcept {
        return TestPaneFocusAddressBarTabTraversal(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_path_doubleClick_enters_edit_mode", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewPathDoubleClickEntersEditMode(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_path_region_keyboard_activation_enters_edit_mode", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewPathRegionKeyboardActivationEntersEditMode(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_path_ancestor_click_navigates_to_ancestor", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewPathAncestorClickNavigatesToAncestor(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewClickInUnfocusedPaneFocusesTargetPane(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_full_path_popup_edit_route", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewFullPathPopupEditRoute(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_full_path_popup_ancestor_click_navigates_to_ancestor", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewFullPathPopupAncestorClickNavigatesToAncestor(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_region_tab_traversal", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewRegionTraversal(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_menu_region_keyboard_activation_opens_menu", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewMenuRegionKeyboardActivationOpensMenu(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_disk_info_region_keyboard_activation_opens_menu", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewDiskInfoRegionKeyboardActivationOpensMenu(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_history_dropdown_keyboard_navigation", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewHistoryDropdownKeyboardNavigation(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_history_dropdown_escape_returns_focus_to_folder_view", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewHistoryDropdownEscapeReturnsFocusToFolderView(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_navigationView_edit_suggest_keyboard_routing", [=](CaseState& state) noexcept {
        return TestPaneNavigationViewEditSuggestKeyboardRouting(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_viewWidth", [=](CaseState& state) noexcept { return TestViewWidthAdjust(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_viewWidth_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestViewWidthKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_refresh", [=](CaseState& state) noexcept { return TestPaneRefresh(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"shortcut_functionbar_dispatch_refresh", [=](CaseState& state) noexcept {
        return TestShortcutFunctionBarDispatchRefresh(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_viewSpace", [=](CaseState& state) noexcept { return TestViewSpace(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_toggleUiChrome", [=](CaseState& state) noexcept { return TestToggleUiChrome(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_persistent_surface_renders_nonempty_labels", [=](CaseState& state) noexcept {
        return TestPersistentMainMenuBarRendersNonEmptyLabels(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_sc_keymenu_shows_temporary_dx_surface", [=](CaseState& state) noexcept {
        return TestMainMenuSystemKeyShowsTemporaryDxMenuBar(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_mnemonic_opens_dx_context_menu", [=](CaseState& state) noexcept {
        return TestMainMenuMnemonicOpensDxContextMenu(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_top_level_highlight_follows_keyboard_opened_root", [=](CaseState& state) noexcept {
        return TestMainMenuTopLevelHighlightFollowsKeyboardOpenedRoot(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_hidden_model_not_owner_draw", [=](CaseState& state) noexcept {
        return TestMainMenuModelRemainsPlainHMenu(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_popup_host_matches_flyout_contract", [=](CaseState& state) noexcept {
        return TestMainMenuPopupHostMatchesFlyoutContract(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_arrow_switches_top_level_popup", [=](CaseState& state) noexcept {
        return TestMainMenuArrowSwitchesTopLevelPopup(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_hover_switches_top_level_popup", [=](CaseState& state) noexcept {
        return TestMainMenuHoverSwitchesTopLevelPopup(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_persistent_direct_hover_switches_top_level_popup", [=](CaseState& state) noexcept {
        return TestMainMenuPersistentDirectHoverSwitchesTopLevelPopup(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup", [=](CaseState& state) noexcept {
        return TestMainMenuPersistentViewToPluginsHoverSwitchesPopup(mainWindow, state, false);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_persistent_view_to_files_hover_highlight_follows_pointer", [=](CaseState& state) noexcept {
        return TestMainMenuPersistentViewToFilesHoverHighlightFollowsPointer(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_persistent_mouseleave_hover_switches_popup", [=](CaseState& state) noexcept {
        return TestMainMenuPersistentViewToPluginsHoverSwitchesPopup(mainWindow, state, true);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_mouse_open_keeps_popup_selection_clear", [=](CaseState& state) noexcept {
        return TestMainMenuMouseOpenKeepsPopupSelectionClear(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_mouse_opened_popup_processes_keyboard_before_mouse_move", [=](CaseState& state) noexcept {
        return TestMainMenuMouseOpenedPopupProcessesKeyboardBeforeMouseMove(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_top_level_mapping_matches_raw_menu", [=](CaseState& state) noexcept {
        return TestMainMenuTopLevelMappingMatchesRawMenu(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_right_justified_items_anchor_to_trailing_edge", [=](CaseState& state) noexcept {
        return TestMainMenuRightJustifiedItemsAnchorToTrailingEdge(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_menuBar_submenu_placement_matches_spec", [=](CaseState& state) noexcept {
        return TestMainMenuSubmenuPlacementMatchesSpec(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_toggleUiChrome_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestToggleUiChromeKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_fileOperationsIssuesPane_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestAppFileOperationsIssuesPaneKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_swapPanes", [=](CaseState& state) noexcept { return TestSwapPanesCommand(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_swapPanes_keeps_navigation_shell_stable", [=](CaseState& state) noexcept {
        return TestSwapPanesKeepsNavigationShellStable(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_displayModeAndSort", [=](CaseState& state) noexcept { return TestDisplayModeAndSortCommands(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu", [=](CaseState& state) noexcept {
        return TestPaneStatusBarUsesOwnedWindowAndSortClickOpensMenu(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_empty_folder_state", [=](CaseState& state) noexcept { return TestEmptyFolderState(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"folderView_column_widths_audit", [=](CaseState& state) noexcept {
        return TestFolderViewColumnWidthsAudit(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_visible_column_widths", [=](CaseState& state) noexcept {
        return TestFolderViewVisibleColumnWidths(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_valid_images_shell_fail", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailValidImagesShellFail(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_aspect_ratio", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailAspectRatio(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_bad_files_fallback", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailBadFilesFallback(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_scroll_stress", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailScrollStress(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_scroll_requeues_visible", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailScrollRequeuesVisible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_resize_requeues_visible", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailResizeRequeuesVisible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_size_change_while_pending", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailSizeChangeWhilePending(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_return_to_normal_icon_size", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailReturnToNormalIconSize(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_thumbnail_sort_popup_slider", [=](CaseState& state) noexcept {
        return TestFolderViewThumbnailSortPopupSlider(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_perf_large_folder_baseline", [=](CaseState& state) noexcept {
        return TestFolderViewPerfLargeFolderBaseline(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_toggle_hidden_system", [=](CaseState& state) noexcept { return TestToggleHiddenAndSystemFiles(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_selection_same_extension", [=](CaseState& state) noexcept { return TestSelectSameExtensionCommands(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_invert", [=](CaseState& state) noexcept { return TestInvertSelectionCommand(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_hide_names", [=](CaseState& state) noexcept { return TestHideNamesCommands(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_selection_save_restore", [=](CaseState& state) noexcept { return TestSelectionSaveRestoreCommands(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_copy_text", [=](CaseState& state) noexcept { return TestCopyTextCommands(mainWindow, state); });
}

namespace
{
