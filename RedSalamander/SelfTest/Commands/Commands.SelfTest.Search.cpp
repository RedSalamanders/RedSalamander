// Commands.SelfTest.Search.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Search test family: local search, Find Files, and Quick Search command cases.

struct LocalSearchProbeCallback final : IFileSystemSearchCallback
{
    LocalSearchProbeCallback()                                           = default;
    LocalSearchProbeCallback(const LocalSearchProbeCallback&)            = delete;
    LocalSearchProbeCallback& operator=(const LocalSearchProbeCallback&) = delete;
    LocalSearchProbeCallback(LocalSearchProbeCallback&&)                 = delete;
    LocalSearchProbeCallback& operator=(LocalSearchProbeCallback&&)      = delete;

    std::atomic<uint32_t> matchCallbacks{0};
    std::atomic<uint32_t> progressCallbacks{0};
    std::atomic<uint32_t> completedProgressCallbacks{0};
    std::atomic<uint32_t> invalidProgressPayloads{0};
    std::atomic<uint32_t> warningFlags{FILESYSTEM_SEARCH_WARNING_NONE};
    std::atomic<uint32_t> shouldCancelCallbacks{0};
    std::atomic<uint32_t> concurrentCallbacks{0};
    std::atomic<uint32_t> cancelAfterMatchCallbacks{0};
    std::atomic<HRESULT> lastCompletedStatus{S_OK};
    std::atomic<bool> inCallback{false};

    HRESULT STDMETHODCALLTYPE FileSystemSearchMatch(const ::FileSystemSearchMatch* match, void* /*cookie*/) noexcept override
    {
        const bool alreadyInCallback = inCallback.exchange(true, std::memory_order_acq_rel);
        const auto callbackExit      = wil::scope_exit([&] { inCallback.store(false, std::memory_order_release); });
        if (alreadyInCallback)
        {
            concurrentCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        }

        if (! match || match->sizeBytes != sizeof(::FileSystemSearchMatch))
        {
            return E_INVALIDARG;
        }

        matchCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchProgress(const ::FileSystemSearchProgress* progress, void* /*cookie*/) noexcept override
    {
        const bool alreadyInCallback = inCallback.exchange(true, std::memory_order_acq_rel);
        const auto callbackExit      = wil::scope_exit([&] { inCallback.store(false, std::memory_order_release); });
        if (alreadyInCallback)
        {
            concurrentCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        }

        if (! progress || progress->sizeBytes != sizeof(::FileSystemSearchProgress))
        {
            invalidProgressPayloads.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        }

        progressCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        warningFlags.fetch_or(progress->warningFlags, std::memory_order_acq_rel);
        if (progress->phase == FILESYSTEM_SEARCH_PHASE_COMPLETED)
        {
            completedProgressCallbacks.fetch_add(1u, std::memory_order_acq_rel);
            lastCompletedStatus.store(progress->statusHint, std::memory_order_release);
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchShouldCancel(BOOL* pCancel, void* /*cookie*/) noexcept override
    {
        if (! pCancel)
        {
            return E_POINTER;
        }

        shouldCancelCallbacks.fetch_add(1u, std::memory_order_acq_rel);
        const uint32_t cancelAfter = cancelAfterMatchCallbacks.load(std::memory_order_acquire);
        *pCancel                   = (cancelAfter != 0u && matchCallbacks.load(std::memory_order_acquire) >= cancelAfter) ? TRUE : FALSE;
        return S_OK;
    }
};

struct BlockingDirectoryWatchCallback final : IFileSystemDirectoryWatchCallback
{
    BlockingDirectoryWatchCallback()                                                 = default;
    BlockingDirectoryWatchCallback(const BlockingDirectoryWatchCallback&)            = delete;
    BlockingDirectoryWatchCallback& operator=(const BlockingDirectoryWatchCallback&) = delete;
    BlockingDirectoryWatchCallback(BlockingDirectoryWatchCallback&&)                 = delete;
    BlockingDirectoryWatchCallback& operator=(BlockingDirectoryWatchCallback&&)      = delete;

    std::atomic<uint32_t> callbacks{0};
    std::atomic<uint32_t> invalidPayloads{0};
    std::atomic<bool> blockFirstCallback{false};
    std::atomic<bool> firstCallbackEntered{false};
    std::atomic<bool> firstCallbackExited{false};
    std::atomic<bool> allowFirstCallbackReturn{false};

    HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const ::FileSystemDirectoryChangeNotification* notification, void* /*cookie*/) noexcept override
    {
        if (! notification || notification->sizeBytes != sizeof(::FileSystemDirectoryChangeNotification))
        {
            invalidPayloads.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        }

        const uint32_t callbackIndex = callbacks.fetch_add(1u, std::memory_order_acq_rel) + 1u;
        if (callbackIndex == 1u && blockFirstCallback.load(std::memory_order_acquire))
        {
            firstCallbackEntered.store(true, std::memory_order_release);
            while (! allowFirstCallbackReturn.load(std::memory_order_acquire))
            {
                ::Sleep(1);
            }
            firstCallbackExited.store(true, std::memory_order_release);
        }

        return S_OK;
    }
};

[[nodiscard]] bool WaitForFlag(const std::atomic<bool>& flag, DWORD timeoutMs) noexcept
{
    const ULONGLONG deadline = ::GetTickCount64() + SelfTest::ScaleTimeout(timeoutMs);
    while (::GetTickCount64() < deadline)
    {
        if (flag.load(std::memory_order_acquire))
        {
            return true;
        }

        ::Sleep(10);
    }

    return flag.load(std::memory_order_acquire);
}

[[nodiscard]] std::vector<HWND> FindVisibleOwnedDxUiContextMenuWindowsForSearchTest(HWND ownerHwnd) noexcept
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

[[nodiscard]] HWND FindVisibleOwnedDxUiContextMenuWindowForSearchTest(HWND ownerHwnd) noexcept
{
    std::vector<HWND> windows = FindVisibleOwnedDxUiContextMenuWindowsForSearchTest(ownerHwnd);
    return windows.empty() ? nullptr : windows.front();
}

void DismissVisibleOwnedDxUiContextMenusForSearchTest(HWND ownerHwnd) noexcept
{
    using namespace std::chrono_literals;

    for (size_t attempt = 0u; attempt < 20u; ++attempt)
    {
        const std::vector<HWND> windows = FindVisibleOwnedDxUiContextMenuWindowsForSearchTest(ownerHwnd);
        if (windows.empty())
        {
            return;
        }

        for (const HWND popup : windows)
        {
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
        }
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }
}

[[nodiscard]] bool WaitForNoVisibleOwnedDxUiContextMenusForSearchTest(HWND ownerHwnd, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        if (FindVisibleOwnedDxUiContextMenuWindowsForSearchTest(ownerHwnd).empty())
        {
            return true;
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }

    return FindVisibleOwnedDxUiContextMenuWindowsForSearchTest(ownerHwnd).empty();
}

void JoinSearchPopupDriverWithUiPumping(std::jthread& popupDriver,
                                        const std::atomic<bool>& popupDriverDone,
                                        HWND ownerHwnd,
                                        std::chrono::milliseconds deadlineTimeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + deadlineTimeout;
    while (! popupDriverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }

    if (! popupDriverDone.load(std::memory_order_acquire))
    {
        popupDriver.request_stop();
        DismissVisibleOwnedDxUiContextMenusForSearchTest(ownerHwnd);
    }

    while (! popupDriverDone.load(std::memory_order_acquire))
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }

    if (popupDriver.joinable())
    {
        popupDriver.join();
    }
}

void CloseAllFindFilesWindowsForSearchTest() noexcept
{
    for (size_t attempt = 0; attempt < 16u && DebugGetFindFilesWindowCount() != 0u; ++attempt)
    {
        const HWND findWindow = GetFindFilesWindowHandle();
        if (! findWindow || IsWindow(findWindow) == FALSE)
        {
            PumpPendingMessages();
            continue;
        }

        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(std::chrono::milliseconds(3000))));
        PumpPendingMessages();
    }
}

[[nodiscard]] std::vector<std::filesystem::path> ReadFindClipboardDropPaths(HWND ownerWindow) noexcept
{
    using namespace std::chrono_literals;

    std::vector<std::filesystem::path> result;
    for (uint32_t attempt = 0; attempt < 20u; ++attempt)
    {
        if (OpenClipboard(ownerWindow) == 0)
        {
            std::this_thread::sleep_for(10ms);
            continue;
        }

        const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });
        HANDLE handle             = GetClipboardData(CF_HDROP);
        if (! handle)
        {
            return result;
        }

        const UINT fileCount = DragQueryFileW(static_cast<HDROP>(handle), 0xFFFFFFFFu, nullptr, 0u);
        result.reserve(fileCount);
        for (UINT index = 0; index < fileCount; ++index)
        {
            const UINT length = DragQueryFileW(static_cast<HDROP>(handle), index, nullptr, 0u);
            if (length == 0u)
            {
                continue;
            }

            std::wstring pathText(static_cast<size_t>(length) + 1u, L'\0');
            const UINT copied = DragQueryFileW(static_cast<HDROP>(handle), index, pathText.data(), length + 1u);
            if (copied == length)
            {
                pathText.resize(length);
                result.emplace_back(pathText);
            }
        }
        return result;
    }

    return result;
}

[[nodiscard]] std::optional<DWORD> ReadFindClipboardPreferredDropEffect(HWND ownerWindow) noexcept
{
    using namespace std::chrono_literals;

    for (uint32_t attempt = 0; attempt < 20u; ++attempt)
    {
        if (OpenClipboard(ownerWindow) == 0)
        {
            std::this_thread::sleep_for(10ms);
            continue;
        }

        const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });
        const UINT format         = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
        if (format == 0u)
        {
            return std::nullopt;
        }

        HANDLE handle = GetClipboardData(format);
        if (! handle)
        {
            return std::nullopt;
        }

        auto* effect = static_cast<DWORD*>(GlobalLock(handle));
        if (! effect)
        {
            return std::nullopt;
        }
        const DWORD resultEffect = *effect;
        GlobalUnlock(handle);
        return resultEffect;
    }

    return std::nullopt;
}

[[nodiscard]] bool ContainsFindClipboardPath(const std::vector<std::filesystem::path>& paths, const std::filesystem::path& expected) noexcept
{
    return std::any_of(paths.begin(), paths.end(), [&](const std::filesystem::path& path) noexcept { return OrdinalString::EqualsNoCasePath(path, expected); });
}

[[nodiscard]] std::wstring DescribeFindClipboardPaths(const std::vector<std::filesystem::path>& paths) noexcept
{
    std::wstring result;
    for (const std::filesystem::path& path : paths)
    {
        if (! result.empty())
        {
            result.append(L"; ");
        }
        result.append(path.native());
    }
    return result.empty() ? std::wstring(L"<empty>") : result;
}

[[nodiscard]] std::wstring DescribeFindClipboardEffect(const std::optional<DWORD>& effect) noexcept
{
    return effect.has_value() ? std::format(L"{}", effect.value()) : std::wstring(L"<none>");
}

void SendFindSplitMenuClick(HWND findWindow, CaseState& state) noexcept
{
    RECT findButtonRect{};
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::FindButton, findButtonRect), L"Failed to capture Find split button rect.");
    if (! state.failure.empty())
    {
        return;
    }

    const int clickX = std::max(findButtonRect.left, findButtonRect.right - 8);
    const int clickY = findButtonRect.top + ((findButtonRect.bottom - findButtonRect.top) / 2);
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
}

void SendFindTargetButtonClick(HWND findWindow, FindFilesDebugFocusTarget target, CaseState& state, std::wstring_view label) noexcept
{
    RECT rect{};
    state.Require(DebugGetFindFilesWindowTargetClientRect(target, rect), std::format(L"Failed to capture {} rect.", label));
    if (! state.failure.empty())
    {
        return;
    }

    const int width  = rect.right - rect.left;
    const int height = rect.bottom - rect.top;
    if (width <= 0 || height <= 0)
    {
        state.Require(false, std::format(L"{} rect was empty.", label));
        return;
    }

    const int clickX = (target == FindFilesDebugFocusTarget::FindButton) ? rect.left + std::max(2, width / 3) : rect.left + (width / 2);
    const int clickY = rect.top + (height / 2);
    SendMouseClickToResolvedPointWindow(findWindow, MAKELPARAM(clickX, clickY));
}

void SendFindResultCommand(HWND findWindow, unsigned int commandId) noexcept
{
    static_cast<void>(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid));
    PumpPendingMessages();
    SendMessageW(findWindow, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(commandId), 0), 0);
    PumpPendingMessages();
}

[[nodiscard]] bool IsFindSelfTestClipboardAvailable(std::wstring* reason = nullptr) noexcept
{
    using namespace std::chrono_literals;

    DWORD lastError = ERROR_SUCCESS;
    for (uint32_t attempt = 0; attempt < 10u; ++attempt)
    {
        if (OpenClipboard(nullptr) != 0)
        {
            CloseClipboard();
            return true;
        }

        lastError = GetLastError();
        std::this_thread::sleep_for(20ms);
    }

    if (reason)
    {
        reason->assign(std::format(L"OS clipboard unavailable for Find selftest (OpenClipboard error={}, openWindow=0x{:X}, owner=0x{:X}).",
                                   lastError,
                                   reinterpret_cast<UINT_PTR>(GetOpenClipboardWindow()),
                                   reinterpret_cast<UINT_PTR>(GetClipboardOwner())));
    }
    return false;
}

[[nodiscard]] bool SkipIfFindSelfTestClipboardUnavailable(CaseState& state) noexcept
{
    std::wstring reason;
    if (IsFindSelfTestClipboardAvailable(&reason))
    {
        return false;
    }

    state.Skip(reason);
    return true;
}

[[nodiscard]] bool InvokeFindSplitMenuItem(HWND findWindow,
                                           HWND ownerWindow,
                                           size_t itemIndex,
                                           CaseState& state,
                                           const std::function<void(const RedSalamander::DxUi::ContextMenuPopupDebugState&)>& validatePopup = {}) noexcept
{
    using namespace std::chrono_literals;

    std::atomic<bool> sawPopup{false};
    std::atomic<bool> stateReadable{false};
    std::atomic<bool> itemInvoked{false};
    std::atomic<bool> driverDone{false};

    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(ownerWindow);
            if (popup && IsWindow(popup) != FALSE)
            {
                sawPopup.store(true, std::memory_order_release);
                RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
                if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState))
                {
                    stateReadable.store(true, std::memory_order_release);
                    if (validatePopup)
                    {
                        validatePopup(popupState);
                    }

                    D2D1_RECT_F itemRectDip{};
                    if (RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(popup, itemIndex, itemRectDip))
                    {
                        const int clickX = static_cast<int>(std::lround((itemRectDip.left + itemRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) /
                                                                        static_cast<float>(USER_DEFAULT_SCREEN_DPI)));
                        const int clickY = static_cast<int>(std::lround((itemRectDip.top + itemRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) /
                                                                        static_cast<float>(USER_DEFAULT_SCREEN_DPI)));
                        const BOOL postedDown  = PostMessageW(popup, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
                        const BOOL postedUp    = PostMessageW(popup, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
                        const bool postedClick = postedDown != FALSE && postedUp != FALSE;
                        itemInvoked.store(postedClick, std::memory_order_release);
                        if (postedClick)
                        {
                            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                            while (IsWindow(popup) != FALSE && std::chrono::steady_clock::now() < closeDeadline)
                            {
                                std::this_thread::sleep_for(10ms);
                            }
                        }
                        if (IsWindow(popup) != FALSE)
                        {
                            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                            while (IsWindow(popup) != FALSE && std::chrono::steady_clock::now() < closeDeadline)
                            {
                                std::this_thread::sleep_for(10ms);
                            }
                        }
                        return;
                    }
                }

                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }

            std::this_thread::sleep_for(10ms);
        }
    });

    SendFindSplitMenuClick(findWindow, state);
    if (! state.failure.empty())
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, ownerWindow, SelfTest::Scale(5000ms));
    state.Require(sawPopup.load(std::memory_order_acquire), L"Find split button did not open a DxUI action menu.");
    state.Require(stateReadable.load(std::memory_order_acquire), L"Find split action menu did not expose readable debug state.");
    state.Require(itemInvoked.load(std::memory_order_acquire), L"Find split action menu item was not invoked.");
    return state.failure.empty();
}

[[nodiscard]] std::wstring FindResultMenuVkToTextForTest(uint32_t vk) noexcept
{
    vk &= 0xFFu;
    if (vk >= VK_F1 && vk <= VK_F24)
    {
        return std::format(L"F{}", static_cast<unsigned>(vk - VK_F1 + 1u));
    }
    if ((vk >= static_cast<uint32_t>(L'0') && vk <= static_cast<uint32_t>(L'9')) || (vk >= static_cast<uint32_t>(L'A') && vk <= static_cast<uint32_t>(L'Z')))
    {
        return std::wstring(1u, static_cast<wchar_t>(vk));
    }

    UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (scanCode == 0u)
    {
        return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
    }

    bool extended = false;
    switch (vk)
    {
        case VK_LEFT:
        case VK_UP:
        case VK_RIGHT:
        case VK_DOWN:
        case VK_PRIOR:
        case VK_NEXT:
        case VK_END:
        case VK_HOME:
        case VK_INSERT:
        case VK_DELETE: extended = true; break;
        default: break;
    }

    LPARAM lParam = static_cast<LPARAM>(scanCode) << 16;
    if (extended)
    {
        lParam |= (1 << 24);
    }

    wchar_t keyName[64]{};
    const int length = GetKeyNameTextW(static_cast<LONG>(lParam), keyName, static_cast<int>(std::size(keyName)));
    return length > 0 ? std::wstring(keyName, static_cast<size_t>(length)) : std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
}

[[nodiscard]] std::wstring FormatFindResultMenuShortcutForTest(const ShortcutManager::ShortcutChord& chord) noexcept
{
    std::wstring result;
    const auto appendPart = [&](std::wstring_view part)
    {
        if (part.empty())
        {
            return;
        }
        if (! result.empty())
        {
            result.append(L"+");
        }
        result.append(part);
    };

    if ((chord.modifiers & ShortcutManager::kModCtrl) != 0u)
    {
        appendPart(LoadEmbeddedStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((chord.modifiers & ShortcutManager::kModAlt) != 0u)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((chord.modifiers & ShortcutManager::kModShift) != 0u)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }
    appendPart(FindResultMenuVkToTextForTest(chord.vk));
    return result;
}

[[nodiscard]] std::optional<size_t> FindMenuItemInSection(const RedSalamander::DxUi::ContextMenuPopupDebugState& popupState,
                                                          std::wstring_view sectionText,
                                                          std::wstring_view itemText,
                                                          std::wstring_view acceleratorText) noexcept
{
    bool inSection = sectionText.empty();
    for (size_t index = 0; index < popupState.itemTexts.size(); ++index)
    {
        const RedSalamander::DxUi::MenuItemKind kind =
            index < popupState.itemKinds.size() ? popupState.itemKinds[index] : RedSalamander::DxUi::MenuItemKind::Standard;
        if (kind == RedSalamander::DxUi::MenuItemKind::Header)
        {
            if (inSection && popupState.itemTexts[index] != sectionText)
            {
                return std::nullopt;
            }
            inSection = popupState.itemTexts[index] == sectionText;
            continue;
        }
        if (! inSection || kind == RedSalamander::DxUi::MenuItemKind::Separator)
        {
            continue;
        }

        const std::wstring_view actualAccelerator =
            index < popupState.itemAcceleratorTexts.size() ? std::wstring_view(popupState.itemAcceleratorTexts[index]) : std::wstring_view{};
        if (popupState.itemTexts[index] == itemText && actualAccelerator == acceleratorText)
        {
            return index;
        }
    }

    return std::nullopt;
}

void RaiseSelfTestWindowForInput(HWND hwnd) noexcept;

enum class FindResultContextMenuOpenMode
{
    Pointer,
    Keyboard,
};

[[nodiscard]] bool InvokeFindResultContextMenuItem(
    HWND findWindow,
    HWND ownerWindow,
    CaseState& state,
    const std::function<std::optional<size_t>(const RedSalamander::DxUi::ContextMenuPopupDebugState&)>& resolveItemIndex,
    FindResultContextMenuOpenMode openMode,
    std::wstring_view failureContext) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    if (openMode == FindResultContextMenuOpenMode::Keyboard)
    {
        FindFilesDebugSnapshot latestSnapshot{};
        bool hasLatestSnapshot = false;
        const auto deadline    = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            if (findWindow && IsWindow(findWindow) != FALSE)
            {
                RaiseSelfTestWindowForInput(findWindow);
                static_cast<void>(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid));
            }

            PumpPendingMessages();
            FindFilesDebugSnapshot candidate{};
            if (GetFindFilesWindowHandle() == findWindow && DebugGetFindFilesWindowSnapshot(candidate))
            {
                latestSnapshot    = candidate;
                hasLatestSnapshot = true;
                const D2D1_RECT_F candidateRowRect = candidate.selectedResultRowRect;
                if (candidate.focusTarget == FindFilesDebugFocusTarget::ResultsGrid && candidate.selectedResultCount > 0u &&
                    candidate.hasWin32Focus && candidateRowRect.right > candidateRowRect.left &&
                    candidateRowRect.bottom > candidateRowRect.top)
                {
                    snapshot = candidate;
                    break;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        if (snapshot.focusTarget != FindFilesDebugFocusTarget::ResultsGrid)
        {
            snapshot = latestSnapshot;
        }
        state.Require(snapshot.focusTarget == FindFilesDebugFocusTarget::ResultsGrid && snapshot.selectedResultCount > 0u &&
                          snapshot.hasWin32Focus && snapshot.selectedResultRowRect.right > snapshot.selectedResultRowRect.left &&
                          snapshot.selectedResultRowRect.bottom > snapshot.selectedResultRowRect.top,
                      std::format(L"{} keyboard context-menu route did not settle focus on the Find results grid. {}",
                                  failureContext,
                                  hasLatestSnapshot ? DescribeFindSnapshotBrief(snapshot)
                                                    : std::format(L"[snapshot unavailable hwnd=0x{:X}]",
                                                                  static_cast<unsigned long long>(
                                                                      reinterpret_cast<uintptr_t>(GetFindFilesWindowHandle())))));
    }
    else
    {
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find result row before opening context menu.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const D2D1_RECT_F rowRect = snapshot.selectedResultRowRect;
    state.Require(rowRect.right > rowRect.left && rowRect.bottom > rowRect.top, L"Selected Find result should expose a context-menu click rectangle.");
    if (! state.failure.empty())
    {
        return false;
    }

    DismissVisibleOwnedDxUiContextMenusForSearchTest(ownerWindow);

    std::atomic<bool> sawPopup{false};
    std::atomic<bool> stateReadable{false};
    std::atomic<bool> itemFound{false};
    std::atomic<bool> itemInvoked{false};
    std::atomic<bool> driverDone{false};

    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(ownerWindow);
            if (! popup || IsWindow(popup) == FALSE)
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }

            sawPopup.store(true, std::memory_order_release);
            RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
            if (! RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState))
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }

            stateReadable.store(true, std::memory_order_release);
            const std::optional<size_t> itemIndex = resolveItemIndex(popupState);
            if (! itemIndex.has_value())
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }

            itemFound.store(true, std::memory_order_release);
            D2D1_RECT_F itemRectDip{};
            if (RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(popup, itemIndex.value(), itemRectDip))
            {
                const int clickX = static_cast<int>(std::lround((itemRectDip.left + itemRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) /
                                                                static_cast<float>(USER_DEFAULT_SCREEN_DPI)));
                const int clickY = static_cast<int>(std::lround((itemRectDip.top + itemRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) /
                                                                static_cast<float>(USER_DEFAULT_SCREEN_DPI)));
                SendMessageW(popup, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
                if (IsWindow(popup) != FALSE)
                {
                    SendMessageW(popup, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
                }
                itemInvoked.store(true, std::memory_order_release);
            }
            else
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
            while (! stopToken.stop_requested() && IsWindow(popup) != FALSE && std::chrono::steady_clock::now() < closeDeadline)
            {
                std::this_thread::sleep_for(10ms);
            }
            if (IsWindow(popup) != FALSE)
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
            }
            return;
        }
    });

    const UINT dpi = GetDpiForWindow(findWindow);
    const int clickX =
        static_cast<int>(std::lround((rowRect.left + rowRect.right) * 0.5f * static_cast<float>(dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI)));
    const int clickY =
        static_cast<int>(std::lround((rowRect.top + rowRect.bottom) * 0.5f * static_cast<float>(dpi) / static_cast<float>(USER_DEFAULT_SCREEN_DPI)));
    const auto sendOpenInput = [&]() noexcept
    {
        RaiseSelfTestWindowForInput(findWindow);
        PumpPendingMessages();
        if (openMode == FindResultContextMenuOpenMode::Pointer)
        {
            // The delivered WM_RBUTTONDOWN/UP carry the routing point via lParam; production
            // derives the context-menu anchor from the delivered message, not GetCursorPos.
            SendMessageW(findWindow, WM_MOUSEMOVE, 0, MAKELPARAM(clickX, clickY));
            SendMessageW(findWindow, WM_RBUTTONDOWN, MK_RBUTTON, MAKELPARAM(clickX, clickY));
            SendMessageW(findWindow, WM_RBUTTONUP, 0, MAKELPARAM(clickX, clickY));
        }
        else
        {
            SendMessageW(findWindow, WM_KEYDOWN, VK_APPS, 0);
            SendMessageW(findWindow, WM_KEYUP, VK_APPS, 0);
        }
    };

    sendOpenInput();

    const auto openDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    auto nextOpenAttempt    = std::chrono::steady_clock::now() + SelfTest::Scale(250ms);
    while (state.failure.empty() && ! driverDone.load(std::memory_order_acquire) && std::chrono::steady_clock::now() < openDeadline)
    {
        PumpPendingMessages();
        if (! sawPopup.load(std::memory_order_acquire) && std::chrono::steady_clock::now() >= nextOpenAttempt)
        {
            sendOpenInput();
            nextOpenAttempt = std::chrono::steady_clock::now() + SelfTest::Scale(250ms);
        }
        std::this_thread::sleep_for(10ms);
    }
    if (! state.failure.empty() || ! driverDone.load(std::memory_order_acquire))
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, ownerWindow, SelfTest::Scale(5000ms));
    state.Require(sawPopup.load(std::memory_order_acquire), std::format(L"{} did not open a DxUI context menu.", failureContext));
    state.Require(stateReadable.load(std::memory_order_acquire), std::format(L"{} did not expose readable menu state.", failureContext));
    state.Require(itemFound.load(std::memory_order_acquire), std::format(L"{} did not expose the expected menu item.", failureContext));
    state.Require(itemInvoked.load(std::memory_order_acquire), std::format(L"{} did not invoke the expected menu item.", failureContext));
    return state.failure.empty();
}

[[nodiscard]] bool ProbeFindSplitMenu(HWND findWindow,
                                      HWND ownerWindow,
                                      CaseState& state,
                                      const std::function<void(const RedSalamander::DxUi::ContextMenuPopupDebugState&)>& validatePopup) noexcept
{
    using namespace std::chrono_literals;

    std::atomic<bool> sawPopup{false};
    std::atomic<bool> stateReadable{false};
    std::atomic<bool> ownerButtonPressed{false};
    std::atomic<bool> driverDone{false};

    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(ownerWindow);
            if (popup && IsWindow(popup) != FALSE)
            {
                sawPopup.store(true, std::memory_order_release);
                RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
                if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState))
                {
                    stateReadable.store(true, std::memory_order_release);
                    FindFilesDebugSnapshot snapshot{};
                    if (DebugGetFindFilesWindowSnapshot(snapshot))
                    {
                        ownerButtonPressed.store(snapshot.findButtonPressed, std::memory_order_release);
                    }
                    if (validatePopup)
                    {
                        validatePopup(popupState);
                    }
                }

                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }

            std::this_thread::sleep_for(10ms);
        }
    });

    SendFindSplitMenuClick(findWindow, state);
    if (! state.failure.empty())
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, ownerWindow, SelfTest::Scale(5000ms));
    state.Require(sawPopup.load(std::memory_order_acquire), L"Find split button did not open a DxUI action menu.");
    state.Require(stateReadable.load(std::memory_order_acquire), L"Find split action menu did not expose readable debug state.");
    state.Require(ownerButtonPressed.load(std::memory_order_acquire), L"Find split button should stay highlighted while its action menu is open.");
    return state.failure.empty();
}

class SearchDirectedSelfTestInputWarning
{
public:
    SearchDirectedSelfTestInputWarning() noexcept
    {
        const int screenW = GetSystemMetrics(SM_CXSCREEN);
        const int screenH = GetSystemMetrics(SM_CYSCREEN);
        if (screenW <= 0 || screenH <= 0)
        {
            return;
        }

        const int width  = std::min(screenW, 900);
        const int height = std::min(screenH, 220);
        const int left   = (std::max)(0, (screenW - width) / 2);
        const int top    = (std::max)(0, (screenH - height) / 3);

        _font.reset(CreateFontW(-56,
                                0,
                                0,
                                0,
                                FW_BOLD,
                                FALSE,
                                FALSE,
                                FALSE,
                                DEFAULT_CHARSET,
                                OUT_DEFAULT_PRECIS,
                                CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY,
                                DEFAULT_PITCH | FF_SWISS,
                                L"Segoe UI"));

        HWND hwnd = CreateWindowExW(WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_LAYERED | WS_EX_TRANSPARENT,
                                    L"STATIC",
                                    L"don't touch the mouse",
                                    WS_POPUP | WS_VISIBLE | WS_BORDER | SS_CENTER | SS_CENTERIMAGE,
                                    left,
                                    top,
                                    width,
                                    height,
                                    nullptr,
                                    nullptr,
                                    GetModuleHandleW(nullptr),
                                    nullptr);
        if (! hwnd)
        {
            return;
        }

        _hwnd.reset(hwnd);
        if (_font)
        {
            SendMessageW(_hwnd.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_font.get()), TRUE);
        }
        static_cast<void>(SetLayeredWindowAttributes(_hwnd.get(), 0, 230, LWA_ALPHA));
        SetWindowPos(_hwnd.get(), HWND_TOPMOST, left, top, width, height, SWP_NOACTIVATE | SWP_SHOWWINDOW);
        UpdateWindow(_hwnd.get());
        PumpPendingMessages();
        _shownAt = GetTickCount64();
    }

    ~SearchDirectedSelfTestInputWarning() noexcept
    {
        if (_hwnd)
        {
            const ULONGLONG elapsedMs = GetTickCount64() - _shownAt;
            if (elapsedMs < 250u)
            {
                std::this_thread::sleep_for(std::chrono::milliseconds(250u - elapsedMs));
                PumpPendingMessages();
            }
            _hwnd.reset();
        }
    }

    SearchDirectedSelfTestInputWarning(const SearchDirectedSelfTestInputWarning&)            = delete;
    SearchDirectedSelfTestInputWarning& operator=(const SearchDirectedSelfTestInputWarning&) = delete;

private:
    wil::unique_hwnd _hwnd;
    wil::unique_hfont _font;
    ULONGLONG _shownAt = 0;
};

[[nodiscard]] bool ProbeFindSplitMenuStationaryHover(HWND findWindow, HWND ownerWindow, size_t itemIndex, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    std::optional<POINT> itemScreenCenter;
    std::atomic<bool> firstDriverDone{false};
    std::jthread firstOpenDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { firstDriverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(ownerWindow);
            if (popup && IsWindow(popup) != FALSE)
            {
                RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
                D2D1_RECT_F itemRectDip{};
                if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState) &&
                    RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(popup, itemIndex, itemRectDip))
                {
                    POINT point{static_cast<LONG>(std::lround((itemRectDip.left + itemRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) /
                                                              static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
                                static_cast<LONG>(std::lround((itemRectDip.top + itemRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) /
                                                              static_cast<float>(USER_DEFAULT_SCREEN_DPI)))};
                    if (ClientToScreen(popup, &point) != FALSE)
                    {
                        itemScreenCenter = point;
                    }
                }

                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }

            std::this_thread::sleep_for(10ms);
        }
    });

    SendFindSplitMenuClick(findWindow, state);
    if (! state.failure.empty())
    {
        firstOpenDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(firstOpenDriver, firstDriverDone, ownerWindow, SelfTest::Scale(5000ms));
    state.Require(itemScreenCenter.has_value(), L"Find split action menu did not expose an item center for stationary-hover validation.");
    state.Require(WaitForNoVisibleOwnedDxUiContextMenusForSearchTest(ownerWindow, SelfTest::Scale(1500ms)),
                  L"Find split action menu geometry popup did not close before stationary-hover validation.");
    if (! state.failure.empty() || ! itemScreenCenter.has_value())
    {
        return false;
    }

    const auto waitForCursorAtScreenPoint = [](POINT expected, std::chrono::milliseconds timeout, POINT& observed) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            if (GetCursorPos(&observed) != FALSE)
            {
                const auto dx = std::llabs(static_cast<long long>(observed.x) - static_cast<long long>(expected.x));
                const auto dy = std::llabs(static_cast<long long>(observed.y) - static_cast<long long>(expected.y));
                if (dx <= 1 && dy <= 1)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(10ms);
        }

        return GetCursorPos(&observed) != FALSE && std::llabs(static_cast<long long>(observed.x) - static_cast<long long>(expected.x)) <= 1 &&
               std::llabs(static_cast<long long>(observed.y) - static_cast<long long>(expected.y)) <= 1;
    };

    POINT cursorBefore{};
    const bool haveCursorBefore = GetCursorPos(&cursorBefore) != FALSE;
    std::optional<SearchDirectedSelfTestInputWarning> inputWarning;
    inputWarning.emplace();
    static_cast<void>(SetCursorPos(itemScreenCenter->x, itemScreenCenter->y));
    POINT cursorAfterSet{};
    state.Require(waitForCursorAtScreenPoint(itemScreenCenter.value(), SelfTest::Scale(1000ms), cursorAfterSet),
                  std::format(L"Find split action menu stationary-hover probe could not place the cursor at the menu item center. expected=({}, {}) actual=({}, {})",
                              itemScreenCenter->x,
                              itemScreenCenter->y,
                              cursorAfterSet.x,
                              cursorAfterSet.y));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto restoreCursor = wil::scope_exit([&]() noexcept
    {
        if (haveCursorBefore)
        {
            static_cast<void>(SetCursorPos(cursorBefore.x, cursorBefore.y));
        }
    });

    std::atomic<bool> hoverObserved{false};
    std::atomic<bool> cursorMovedToItem{false};
    std::atomic<bool> secondDriverDone{false};
    std::wstring hoverFailureDetails;
    std::jthread secondOpenDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { secondDriverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(ownerWindow);
            if (popup && IsWindow(popup) != FALSE)
            {
                RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
                D2D1_RECT_F itemRectDip{};
                if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState) &&
                    RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(popup, itemIndex, itemRectDip))
                {
                    POINT itemClient{
                        static_cast<LONG>(std::lround((itemRectDip.left + itemRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) /
                                                      static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
                        static_cast<LONG>(std::lround((itemRectDip.top + itemRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) /
                                                      static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
                    };
                    POINT itemCenter = itemClient;
                    RECT popupRect{};
                    if (ClientToScreen(popup, &itemCenter) != FALSE && GetWindowRect(popup, &popupRect) != FALSE)
                    {
                        const POINT deliveredPoint{itemCenter.x - popupRect.left, itemCenter.y - popupRect.top};
                        const LPARAM deliveredMove =
                            MAKELPARAM(static_cast<WORD>(static_cast<SHORT>(deliveredPoint.x)), static_cast<WORD>(static_cast<SHORT>(deliveredPoint.y)));

                        static_cast<void>(SetCursorPos(itemCenter.x, itemCenter.y));
                        const auto hoverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < hoverDeadline)
                        {
                            if (PostMessageW(popup, WM_MOUSEMOVE, 0, deliveredMove) != FALSE)
                            {
                                cursorMovedToItem.store(true, std::memory_order_release);
                            }
                            std::this_thread::sleep_for(25ms);

                            RedSalamander::DxUi::ContextMenuPopupDebugState hoverState{};
                            RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState paintState{};
                            if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, hoverState) &&
                                hoverState.hoveredIndex == std::optional<size_t>{itemIndex} &&
                                RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popup, itemIndex, paintState) && paintState.usesHighlightFill)
                            {
                                hoverObserved.store(true, std::memory_order_release);
                                break;
                            }
                        }

                        if (! hoverObserved.load(std::memory_order_acquire))
                        {
                            RedSalamander::DxUi::ContextMenuPopupDebugState finalState{};
                            static_cast<void>(RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, finalState));
                            RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState finalPaint{};
                            static_cast<void>(RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popup, itemIndex, finalPaint));
                            hoverFailureDetails =
                                std::format(L"Find split action menu did not repaint hover highlight for the delivered stationary pointer. "
                                            L"(hovered={}, keyboard={}, firstCenter=({},{}), currentCenter=({},{}), delivered=({},{}), "
                                            L"popup=({},{} {}x{}), row=({:.1f},{:.1f},{:.1f},{:.1f}), highlight={}, renders={})",
                                            finalState.hoveredIndex.has_value() ? static_cast<long long>(finalState.hoveredIndex.value()) : -1ll,
                                            finalState.keyboardIndex.has_value() ? static_cast<long long>(finalState.keyboardIndex.value()) : -1ll,
                                            itemScreenCenter->x,
                                            itemScreenCenter->y,
                                            itemCenter.x,
                                            itemCenter.y,
                                            deliveredPoint.x,
                                            deliveredPoint.y,
                                            popupRect.left,
                                            popupRect.top,
                                            popupRect.right - popupRect.left,
                                            popupRect.bottom - popupRect.top,
                                            itemRectDip.left,
                                            itemRectDip.top,
                                            itemRectDip.right,
                                            itemRectDip.bottom,
                                            finalPaint.usesHighlightFill ? L"true" : L"false",
                                            finalState.renderCount);
                        }
                    }
                }

                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }

            std::this_thread::sleep_for(10ms);
        }
    });

    SendFindSplitMenuClick(findWindow, state);
    if (! state.failure.empty())
    {
        secondOpenDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(secondOpenDriver, secondDriverDone, ownerWindow, SelfTest::Scale(5000ms));
    state.Require(cursorMovedToItem.load(std::memory_order_acquire), L"Failed to route the stationary Find split action menu pointer over the target item.");
    state.Require(hoverObserved.load(std::memory_order_acquire),
                  hoverFailureDetails.empty() ? L"Find split action menu did not highlight the item under a stationary pointer." : hoverFailureDetails);
    return state.failure.empty();
}

// Light-dismisses an owned DxUi context-menu popup by delivering a synthetic outside
// left-click to the popup's window proc. The menu modal loop reconstructs the screen
// point from the message lParam (ClientToScreen on the popup) and dismisses when it
// lands outside every popup -- so no live cursor is moved and the interactive user's
// pointer is never disturbed.
[[nodiscard]] bool PostSelfTestOutsideClickToPopup(HWND popup, POINT outsideScreenPoint) noexcept
{
    RECT popupRect{};
    if (! popup || GetWindowRect(popup, &popupRect) == FALSE)
    {
        return false;
    }

    const LPARAM lParam = MAKELPARAM(static_cast<WORD>(static_cast<SHORT>(outsideScreenPoint.x - popupRect.left)),
                                     static_cast<WORD>(static_cast<SHORT>(outsideScreenPoint.y - popupRect.top)));
    return PostMessageW(popup, WM_MOUSEMOVE, 0, lParam) != FALSE && PostMessageW(popup, WM_LBUTTONDOWN, MK_LBUTTON, lParam) != FALSE &&
           PostMessageW(popup, WM_LBUTTONUP, 0, lParam) != FALSE;
}

void RaiseSelfTestWindowForInput(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    const HWND foregroundWindow    = GetForegroundWindow();
    const DWORD currentThreadId    = GetCurrentThreadId();
    const DWORD foregroundThreadId = foregroundWindow ? GetWindowThreadProcessId(foregroundWindow, nullptr) : 0u;
    const bool attachedForegroundThread =
        foregroundThreadId != 0u && foregroundThreadId != currentThreadId && AttachThreadInput(foregroundThreadId, currentThreadId, TRUE) != FALSE;
    const auto detachForegroundThread = wil::scope_exit([&]() noexcept
    {
        if (attachedForegroundThread)
        {
            static_cast<void>(AttachThreadInput(foregroundThreadId, currentThreadId, FALSE));
        }
    });

    ShowWindow(hwnd, SW_SHOWNORMAL);
    static_cast<void>(BringWindowToTop(hwnd));
    static_cast<void>(SetActiveWindow(hwnd));
    static_cast<void>(SetForegroundWindow(hwnd));
    static_cast<void>(SetFocus(hwnd));
    SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    UpdateWindow(hwnd);
    PumpPendingMessages();
}

[[nodiscard]] bool FocusFindRootNavigationForKeyboard(HWND findWindow, std::chrono::milliseconds timeout, FindFilesDebugSnapshot* snapshot = nullptr) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot latest{};
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        if (findWindow && IsWindow(findWindow) != FALSE)
        {
            RaiseSelfTestWindowForInput(findWindow);
            static_cast<void>(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::RootCombo));
        }

        const bool focused = WaitForFindSnapshot(
            [](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.focusTarget == FindFilesDebugFocusTarget::RootCombo && value.rootNavigationVisible &&
                   value.rootNavigationEmbedded && value.rootNavigationHasWin32Focus && ! value.rootNavigationEditMode && value.hasWin32Focus &&
                   value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
        },
            SelfTest::Scale(150ms),
            &latest);
        if (snapshot)
        {
            *snapshot = latest;
        }
        if (focused)
        {
            return true;
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(25ms);
    }

    if (snapshot && DebugGetFindFilesWindowSnapshot(latest))
    {
        *snapshot = latest;
    }
    return false;
}

[[nodiscard]] HWND ResolveFindDestinationNavigationHwnd(HWND findWindow, RECT historyRect) noexcept;
[[nodiscard]] std::optional<POINT> ChooseOwnerPointOutsidePopup(HWND ownerWindow, HWND popup) noexcept;
[[nodiscard]] bool EnsureFindDestinationNavigationReadOnlyMode(HWND navHwnd, CaseState& state, std::wstring_view probeName) noexcept;
[[nodiscard]] HWND ResolveFindDestinationNavigationDxEditHost(HWND navHwnd) noexcept
{
    return FindWindowExW(navHwnd, nullptr, L"RedSalamander.NavigationView.DxHost", nullptr);
}

[[nodiscard]] std::optional<RECT> MapClientRectBetweenWindows(HWND fromWindow, HWND toWindow, RECT clientRect) noexcept
{
    POINT topLeft{clientRect.left, clientRect.top};
    POINT bottomRight{clientRect.right, clientRect.bottom};
    if (ClientToScreen(fromWindow, &topLeft) == FALSE || ClientToScreen(fromWindow, &bottomRight) == FALSE || ScreenToClient(toWindow, &topLeft) == FALSE ||
        ScreenToClient(toWindow, &bottomRight) == FALSE)
    {
        return std::nullopt;
    }

    return RECT{topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};
}

[[nodiscard]] bool ProbeFindDestinationHistoryMenu(HWND findWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find destination navigation snapshot before history-arrow click.");
    const RECT arrowRect = snapshot.destinationNavigationHistoryRect;
    state.Require(arrowRect.right > arrowRect.left && arrowRect.bottom > arrowRect.top,
                  L"Find destination navigation history arrow should expose a non-empty client rectangle.");
    if (! state.failure.empty())
    {
        return false;
    }

    POINT clickPointInFind{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    POINT clickPoint = clickPointInFind;
    state.Require(ClientToScreen(findWindow, &clickPoint) != FALSE, L"Failed to convert Find destination history arrow point to screen coordinates.");
    POINT arrowTopLeft{arrowRect.left, arrowRect.top};
    state.Require(ClientToScreen(findWindow, &arrowTopLeft) != FALSE, L"Failed to convert Find destination history arrow top edge to screen coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    RaiseSelfTestWindowForInput(findWindow);
    PumpPendingMessages();

    const HWND navHwnd = ResolveFindDestinationNavigationHwnd(findWindow, arrowRect);
    state.Require(navHwnd && IsWindow(navHwnd) != FALSE, L"Find destination navigation child should exist for history-arrow probe.");
    POINT clickPointInNav = clickPoint;
    state.Require(ScreenToClient(navHwnd, &clickPointInNav) != FALSE, L"Failed to map Find destination history arrow click point to the navigation child.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> sawPopup{false};
    std::atomic<bool> stateReadable{false};
    std::atomic<bool> popupHadHistoryItems{false};
    std::atomic<bool> popupRenderedInitialFrame{false};
    std::atomic<bool> popupAnchoredAboveFooter{false};
    std::atomic<bool> cursorMovedToItem{false};
    std::atomic<bool> hoverObserved{false};
    std::atomic<bool> outsideClickSent{false};
    std::atomic<bool> popupDismissed{false};
    std::atomic<bool> driverDone{false};
    std::wstring hoverFailureDetails;
    POINT cursorBefore{};
    const bool haveCursorBefore = GetCursorPos(&cursorBefore) != FALSE;
    const auto restoreCursor    = wil::scope_exit([&]() noexcept
    {
        if (haveCursorBefore)
        {
            static_cast<void>(SetCursorPos(cursorBefore.x, cursorBefore.y));
        }
    });

    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(findWindow);
            if (! popup || IsWindow(popup) == FALSE)
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }

            sawPopup.store(true, std::memory_order_release);
            RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
            if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState))
            {
                stateReadable.store(true, std::memory_order_release);
                popupHadHistoryItems.store(! popupState.itemTexts.empty(), std::memory_order_release);
                popupRenderedInitialFrame.store(popupState.renderCount > 0u, std::memory_order_release);
                popupAnchoredAboveFooter.store(popupState.surfaceRectPx.bottom <= arrowTopLeft.y + 1, std::memory_order_release);
            }

            D2D1_RECT_F itemRectDip{};
            if (RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(popup, 0u, itemRectDip))
            {
                RedSalamander::DxUi::ContextMenuPopupDebugState geometryState{};
                if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, geometryState))
                {
                    POINT itemClient{
                        static_cast<LONG>(std::lround((itemRectDip.left + itemRectDip.right) * 0.5f * static_cast<float>(geometryState.dpi) /
                                                      static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
                        static_cast<LONG>(std::lround((itemRectDip.top + itemRectDip.bottom) * 0.5f * static_cast<float>(geometryState.dpi) /
                                                      static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
                    };
                    POINT itemCenter = itemClient;
                    if (ClientToScreen(popup, &itemCenter) != FALSE)
                    {
                        static_cast<void>(SetCursorPos(itemCenter.x, itemCenter.y));
                        const LPARAM deliveredMove =
                            MAKELPARAM(static_cast<WORD>(static_cast<SHORT>(itemClient.x)), static_cast<WORD>(static_cast<SHORT>(itemClient.y)));
                        const auto hoverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < hoverDeadline)
                        {
                            if (PostMessageW(popup, WM_MOUSEMOVE, 0, deliveredMove) != FALSE)
                            {
                                cursorMovedToItem.store(true, std::memory_order_release);
                            }
                            std::this_thread::sleep_for(25ms);
                            RedSalamander::DxUi::ContextMenuPopupDebugState hoverState{};
                            RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState paintState{};
                            if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, hoverState) &&
                                hoverState.hoveredIndex == std::optional<size_t>{0u} &&
                                RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popup, 0u, paintState) && paintState.usesHighlightFill)
                            {
                                hoverObserved.store(true, std::memory_order_release);
                                break;
                            }
                        }
                        if (! hoverObserved.load(std::memory_order_acquire))
                        {
                            RedSalamander::DxUi::ContextMenuPopupDebugState finalState{};
                            static_cast<void>(RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, finalState));
                            RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState finalPaint{};
                            static_cast<void>(RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popup, 0u, finalPaint));
                            hoverFailureDetails = std::format(
                                L"Find destination history menu did not repaint hover highlight for the item under the delivered pointer. "
                                L"(hovered={}, keyboard={}, center=({},{}), delivered=({},{}), row=({:.1f},{:.1f},{:.1f},{:.1f}), highlight={}, renders={})",
                                finalState.hoveredIndex.has_value() ? static_cast<long long>(finalState.hoveredIndex.value()) : -1ll,
                                finalState.keyboardIndex.has_value() ? static_cast<long long>(finalState.keyboardIndex.value()) : -1ll,
                                itemCenter.x,
                                itemCenter.y,
                                itemClient.x,
                                itemClient.y,
                                itemRectDip.left,
                                itemRectDip.top,
                                itemRectDip.right,
                                itemRectDip.bottom,
                                finalPaint.usesHighlightFill ? L"true" : L"false",
                                finalState.renderCount);
                        }
                    }
                }
            }

            if (const std::optional<POINT> outsidePoint = ChooseOwnerPointOutsidePopup(findWindow, popup); outsidePoint.has_value())
            {
                outsideClickSent.store(PostSelfTestOutsideClickToPopup(popup, outsidePoint.value()), std::memory_order_release);
            }
            if (! outsideClickSent.load(std::memory_order_acquire))
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
            }

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                if (IsWindow(popup) == FALSE || IsWindowVisible(popup) == FALSE)
                {
                    popupDismissed.store(true, std::memory_order_release);
                    return;
                }

                std::this_thread::sleep_for(10ms);
            }
            if (IsWindow(popup) != FALSE)
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                const auto escapeDismissDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < escapeDismissDeadline)
                {
                    if (IsWindow(popup) == FALSE || IsWindowVisible(popup) == FALSE)
                    {
                        popupDismissed.store(true, std::memory_order_release);
                        return;
                    }

                    std::this_thread::sleep_for(10ms);
                }
            }
            return;
        }
    });

    RaiseSelfTestWindowForInput(findWindow);
    PumpPendingMessages();
    SendMessageW(navHwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickPointInNav.x, clickPointInNav.y));
    SendMessageW(navHwnd, WM_LBUTTONUP, 0, MAKELPARAM(clickPointInNav.x, clickPointInNav.y));
    if (! state.failure.empty())
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, findWindow, SelfTest::Scale(5000ms));

    state.Require(sawPopup.load(std::memory_order_acquire), L"Find destination history arrow did not open a DxUI history menu immediately.");
    state.Require(stateReadable.load(std::memory_order_acquire), L"Find destination history menu did not expose readable debug state.");
    state.Require(popupHadHistoryItems.load(std::memory_order_acquire), L"Find destination history menu opened without history items.");
    state.Require(popupRenderedInitialFrame.load(std::memory_order_acquire),
                  L"Find destination history menu did not have a rendered initial frame when it opened.");
    state.Require(popupAnchoredAboveFooter.load(std::memory_order_acquire),
                  L"Find destination history menu surface should open above the embedded destination footer.");
    state.Require(cursorMovedToItem.load(std::memory_order_acquire), L"Failed to route pointer movement over the Find destination history menu item.");
    state.Require(hoverObserved.load(std::memory_order_acquire),
                  hoverFailureDetails.empty()
                      ? L"Find destination history menu did not repaint hover highlight for the item under the delivered pointer."
                      : hoverFailureDetails);
    state.Require(outsideClickSent.load(std::memory_order_acquire), L"Failed to send real outside click for Find destination history menu light-dismiss.");
    state.Require(popupDismissed.load(std::memory_order_acquire), L"Find destination history menu did not close after outside-click and Escape cleanup.");
    return state.failure.empty();
}

[[nodiscard]] bool ProbeFindDestinationHistoryMenuFromActiveEditMode(HWND findWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find destination navigation snapshot before active-edit history probe.");
    state.Require(snapshot.destinationNavigationVisible, L"Find destination navigation bar should be visible before active-edit history probe.");
    const RECT arrowRect = snapshot.destinationNavigationHistoryRect;
    state.Require(arrowRect.right > arrowRect.left && arrowRect.bottom > arrowRect.top,
                  L"Find destination navigation history arrow should expose a non-empty client rectangle before active-edit history probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND navHwnd = ResolveFindDestinationNavigationHwnd(findWindow, arrowRect);
    state.Require(navHwnd && IsWindow(navHwnd) != FALSE, L"Find destination navigation child should exist for active-edit history probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    RaiseSelfTestWindowForInput(findWindow);
    SendMessageW(navHwnd, WM_MOUSELEAVE, 0, 0);
    PumpPendingMessages();

    FindFilesDebugSnapshot clearedHoverSnapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(clearedHoverSnapshot), L"Failed to capture destination navigation snapshot after clearing hover.");
    state.Require(! clearedHoverSnapshot.destinationNavigationHistoryHovered,
                  L"Destination navigation history hover should be clear before the active-edit hover probe starts.");
    if (! state.failure.empty())
    {
        return false;
    }

    POINT editPointInFind{std::max<LONG>(arrowRect.left - 24, 4), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    PumpPendingMessages();
    if (! state.failure.empty())
    {
        return false;
    }
    static_cast<void>(MapWindowPoints(findWindow, navHwnd, &editPointInFind, 1u));
    SendMessageW(navHwnd, WM_LBUTTONDBLCLK, 0, MAKELPARAM(editPointInFind.x, editPointInFind.y));
    PumpPendingMessages();

    FindFilesDebugSnapshot editModeSnapshot{};
    state.Require(
        WaitForFindSnapshot(
            [](const FindFilesDebugSnapshot& value) noexcept { return value.destinationNavigationEditMode; }, SelfTest::Scale(1500ms), &editModeSnapshot),
        std::format(L"Destination navigation should enter edit mode before active-edit history-arrow probe. {}", DescribeFindSnapshotBrief(editModeSnapshot)));
    const HWND editHost = ResolveFindDestinationNavigationDxEditHost(navHwnd);
    state.Require(editHost && IsWindow(editHost) != FALSE && IsWindowVisible(editHost) != FALSE,
                  L"Destination navigation edit host should be active before active-edit history-arrow probe.");
    state.Require(! editModeSnapshot.destinationNavigationHistoryHovered,
                  L"Destination navigation history hover should still be clear after entering edit mode from the path area.");
    const RECT activeArrowRect = editModeSnapshot.destinationNavigationHistoryRect;
    state.Require(activeArrowRect.right > activeArrowRect.left && activeArrowRect.bottom > activeArrowRect.top,
                  L"Find destination navigation history arrow should expose a non-empty client rectangle while edit mode is active.");
    if (! state.failure.empty())
    {
        return false;
    }

    POINT clickPoint{activeArrowRect.left + ((activeArrowRect.right - activeArrowRect.left) / 2),
                     activeArrowRect.top + ((activeArrowRect.bottom - activeArrowRect.top) / 2)};
    state.Require(ClientToScreen(findWindow, &clickPoint) != FALSE, L"Failed to convert active-edit history arrow click point to screen coordinates.");
    RECT editHostRect{};
    state.Require(GetWindowRect(editHost, &editHostRect) != FALSE, L"Failed to capture active destination edit-host screen rectangle.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND arrowHitWindow = WindowFromPoint(clickPoint);
    state.Require(PtInRect(&editHostRect, clickPoint) == 0,
                  std::format(L"Destination edit host must not overlap the history arrow; editHost=({}, {}, {}, {}), arrowCenter=({}, {}).",
                              editHostRect.left,
                              editHostRect.top,
                              editHostRect.right,
                              editHostRect.bottom,
                              clickPoint.x,
                              clickPoint.y));
    state.Require(arrowHitWindow != editHost && IsChild(editHost, arrowHitWindow) == FALSE,
                  std::format(L"Destination history arrow should not hit-test to the edit host while edit mode is active; hitWindow={:#x}, editHost={:#x}.",
                              reinterpret_cast<uintptr_t>(arrowHitWindow),
                              reinterpret_cast<uintptr_t>(editHost)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<RECT> historyRectInNav = MapClientRectBetweenWindows(findWindow, navHwnd, activeArrowRect);
    state.Require(historyRectInNav.has_value(), L"Failed to map active-edit history arrow rectangle to destination navigation coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }
    const RECT activeHistoryRectInNav = historyRectInNav.value();
    POINT hoverPointInNav{activeHistoryRectInNav.left + ((activeHistoryRectInNav.right - activeHistoryRectInNav.left) / 2),
                          activeHistoryRectInNav.top + ((activeHistoryRectInNav.bottom - activeHistoryRectInNav.top) / 2)};
    PumpPendingMessages();
    std::this_thread::sleep_for(40ms);
    PumpPendingMessages();

    SendMessageW(navHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(hoverPointInNav.x, hoverPointInNav.y));

    FindFilesDebugSnapshot hoverSnapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(hoverSnapshot), L"Failed to capture active-edit destination hover snapshot.");
    state.Require(hoverSnapshot.destinationNavigationHistoryHovered,
                  std::format(L"Destination navigation history arrow should still hover while path edit mode is active; "
                              L"hoverPt=({}, {}), menu={}, history={}, disk={}, segment={}, separator={}.",
                              hoverPointInNav.x,
                              hoverPointInNav.y,
                              hoverSnapshot.destinationNavigationMenuHovered ? 1 : 0,
                              hoverSnapshot.destinationNavigationHistoryHovered ? 1 : 0,
                              hoverSnapshot.destinationNavigationDiskHovered ? 1 : 0,
                              hoverSnapshot.destinationNavigationHoveredSegmentIndex,
                              hoverSnapshot.destinationNavigationHoveredSeparatorIndex));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto exitEditMode = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(editHost) != FALSE)
        {
            SendMessageW(editHost, WM_KEYDOWN, VK_ESCAPE, 0);
            SendMessageW(editHost, WM_KEYUP, VK_ESCAPE, 0);
        }
        PumpPendingMessages();
    });

    const uint64_t baselineDropdownOpenCount = editModeSnapshot.destinationNavigationHistoryDropdownOpenCount;
    std::atomic<bool> sawPopup{false};
    std::atomic<bool> popupDismissed{false};
    std::atomic<bool> driverDone{false};
    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(findWindow);
            if (! popup || IsWindow(popup) == FALSE)
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }

            sawPopup.store(true, std::memory_order_release);
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                if (IsWindow(popup) == FALSE || IsWindowVisible(popup) == FALSE)
                {
                    popupDismissed.store(true, std::memory_order_release);
                    return;
                }

                std::this_thread::sleep_for(10ms);
            }
            return;
        }
    });

    SendMessageW(navHwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(hoverPointInNav.x, hoverPointInNav.y));
    SendMessageW(navHwnd, WM_LBUTTONUP, 0, MAKELPARAM(hoverPointInNav.x, hoverPointInNav.y));
    if (! state.failure.empty())
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, findWindow, SelfTest::Scale(5000ms));

    FindFilesDebugSnapshot afterPopup{};
    state.Require(DebugGetFindFilesWindowSnapshot(afterPopup), L"Failed to capture destination navigation snapshot after active-edit history probe.");
    const bool dropdownOpened =
        sawPopup.load(std::memory_order_acquire) || afterPopup.destinationNavigationHistoryDropdownOpenCount > baselineDropdownOpenCount;
    const bool dropdownDismissed =
        popupDismissed.load(std::memory_order_acquire) ||
        (afterPopup.destinationNavigationHistoryDropdownOpenCount > baselineDropdownOpenCount && ! afterPopup.destinationNavigationHistoryDropdownVisible);
    state.Require(dropdownOpened,
                  std::format(L"Find destination history arrow should open its menu even while the destination path edit host is active; before={}, "
                              L"after={}, visible={}.",
                              baselineDropdownOpenCount,
                              afterPopup.destinationNavigationHistoryDropdownOpenCount,
                              afterPopup.destinationNavigationHistoryDropdownVisible ? 1 : 0));
    state.Require(dropdownDismissed, L"Find destination active-edit history menu opened but did not close from Escape.");
    return state.failure.empty();
}

[[nodiscard]] bool ProbeFindDestinationHistoryMenuThroughStaleEditHost(HWND findWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find destination navigation snapshot before stale edit-host probe.");
    state.Require(snapshot.destinationNavigationVisible, L"Find destination navigation bar should be visible before stale edit-host probe.");
    const RECT arrowRect = snapshot.destinationNavigationHistoryRect;
    state.Require(arrowRect.right > arrowRect.left && arrowRect.bottom > arrowRect.top,
                  L"Find destination navigation history arrow should expose a non-empty client rectangle before stale edit-host probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND navHwnd = ResolveFindDestinationNavigationHwnd(findWindow, arrowRect);
    state.Require(navHwnd && IsWindow(navHwnd) != FALSE, L"Find destination navigation child should exist for stale edit-host probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    RaiseSelfTestWindowForInput(findWindow);
    const LONG destinationLeft = static_cast<LONG>(std::lround(snapshot.destinationNavigationRect.left));
    POINT editPointInFind{std::max<LONG>(destinationLeft + 8, arrowRect.left - 20), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    PumpPendingMessages();
    if (! state.failure.empty())
    {
        return false;
    }
    static_cast<void>(MapWindowPoints(findWindow, navHwnd, &editPointInFind, 1u));
    SendMessageW(navHwnd, WM_LBUTTONDBLCLK, 0, MAKELPARAM(editPointInFind.x, editPointInFind.y));
    PumpPendingMessages();

    FindFilesDebugSnapshot editModeSnapshot{};
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.destinationNavigationEditMode; },
                                      SelfTest::Scale(1500ms),
                                      &editModeSnapshot),
                  std::format(L"Destination navigation should enter edit mode before stale edit-host probe. {}", DescribeFindSnapshotBrief(editModeSnapshot)));
    const HWND editHost = ResolveFindDestinationNavigationDxEditHost(navHwnd);
    state.Require(editHost && IsWindow(editHost) != FALSE, L"Destination navigation edit host should exist after entering edit mode.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(editHost, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(editHost, WM_KEYUP, VK_ESCAPE, 0);
    PumpPendingMessages();
    state.Require(EnsureFindDestinationNavigationReadOnlyMode(navHwnd, state, L"stale edit-host overlay probe"),
                  L"Destination navigation should leave edit mode before the stale edit-host overlay is forced visible.");
    if (! state.failure.empty())
    {
        return false;
    }
    ShowWindow(editHost, SW_SHOWNOACTIVATE);
    SetWindowPos(editHost, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
    PumpPendingMessages();

    POINT arrowCenterScreen{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    state.Require(ClientToScreen(findWindow, &arrowCenterScreen) != FALSE, L"Failed to convert stale edit-host history arrow center to screen coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LRESULT staleHostHit = SendMessageW(editHost, WM_NCHITTEST, 0, MAKELPARAM(arrowCenterScreen.x, arrowCenterScreen.y));
    PumpPendingMessages();
    state.Require(staleHostHit == HTTRANSPARENT,
                  std::format(L"Inactive destination edit host should retire as transparent on pointer hit-test; hit={}.", staleHostHit));
    state.Require(IsWindowVisible(editHost) == FALSE, L"Inactive destination edit host should hide when retired by pointer hit-test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::optional<RECT> historyRectInNav = MapClientRectBetweenWindows(findWindow, navHwnd, arrowRect);
    state.Require(historyRectInNav.has_value(), L"Failed to map stale edit-host history arrow rectangle to destination navigation coordinates.");
    if (! state.failure.empty())
    {
        return false;
    }
    const RECT historyRectNav = historyRectInNav.value();
    POINT clickPointInNav{historyRectNav.left + ((historyRectNav.right - historyRectNav.left) / 2),
                          historyRectNav.top + ((historyRectNav.bottom - historyRectNav.top) / 2)};

    const uint64_t baselineDropdownOpenCount = snapshot.destinationNavigationHistoryDropdownOpenCount;
    std::atomic<bool> sawPopup{false};
    std::atomic<bool> popupDismissed{false};
    std::atomic<bool> driverDone{false};
    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(findWindow);
            if (! popup || IsWindow(popup) == FALSE)
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }

            sawPopup.store(true, std::memory_order_release);
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                PumpPendingMessages();
                if (IsWindow(popup) == FALSE || IsWindowVisible(popup) == FALSE)
                {
                    popupDismissed.store(true, std::memory_order_release);
                    return;
                }

                std::this_thread::sleep_for(10ms);
            }
            return;
        }
    });

    RaiseSelfTestWindowForInput(findWindow);
    PumpPendingMessages();
    std::this_thread::sleep_for(30ms);
    PumpPendingMessages();
    SendMessageW(navHwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickPointInNav.x, clickPointInNav.y));
    SendMessageW(navHwnd, WM_LBUTTONUP, 0, MAKELPARAM(clickPointInNav.x, clickPointInNav.y));

    if (! state.failure.empty())
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, findWindow, SelfTest::Scale(5000ms));

    FindFilesDebugSnapshot afterPopup{};
    state.Require(DebugGetFindFilesWindowSnapshot(afterPopup), L"Failed to capture destination navigation snapshot after stale edit-host history probe.");
    const bool dropdownOpened =
        sawPopup.load(std::memory_order_acquire) || afterPopup.destinationNavigationHistoryDropdownOpenCount > baselineDropdownOpenCount;
    const bool dropdownDismissed =
        popupDismissed.load(std::memory_order_acquire) ||
        (afterPopup.destinationNavigationHistoryDropdownOpenCount > baselineDropdownOpenCount && ! afterPopup.destinationNavigationHistoryDropdownVisible);
    state.Require(dropdownOpened,
                  std::format(L"Find destination history arrow should open its menu even if a stale edit host is visible above the embedded navigation bar; "
                              L"before={}, after={}, visible={}.",
                              baselineDropdownOpenCount,
                              afterPopup.destinationNavigationHistoryDropdownOpenCount,
                              afterPopup.destinationNavigationHistoryDropdownVisible ? 1 : 0));
    state.Require(dropdownDismissed, L"Find destination history menu opened through stale edit host but did not close from Escape.");
    return state.failure.empty();
}

[[nodiscard]] HWND ResolveFindDestinationNavigationHwnd(HWND findWindow, RECT historyRect) noexcept
{
    POINT historyCenter{historyRect.left + ((historyRect.right - historyRect.left) / 2), historyRect.top + ((historyRect.bottom - historyRect.top) / 2)};
    for (HWND current = nullptr; (current = FindWindowExW(findWindow, current, L"RedSalamander.NavigationView", nullptr)) != nullptr;)
    {
        RECT childRect{};
        if (GetWindowRect(current, &childRect) == FALSE)
        {
            continue;
        }

        static_cast<void>(MapWindowPoints(nullptr, findWindow, reinterpret_cast<POINT*>(&childRect), 2));
        RECT hitRect = childRect;
        InflateRect(&hitRect, 2, 2);
        if (PtInRect(&hitRect, historyCenter) != FALSE)
        {
            return current;
        }
    }

    if (ClientToScreen(findWindow, &historyCenter) != FALSE)
    {
        for (HWND current = WindowFromPoint(historyCenter); current && current != findWindow; current = GetParent(current))
        {
            std::array<wchar_t, 128> className{};
            const int classNameLength = GetClassNameW(current, className.data(), static_cast<int>(className.size()));
            if (classNameLength > 0 &&
                std::wstring_view(className.data(), static_cast<size_t>(classNameLength)) == std::wstring_view(L"RedSalamander.NavigationView"))
            {
                return current;
            }
        }
    }

    return FindWindowExW(findWindow, nullptr, L"RedSalamander.NavigationView", nullptr);
}

[[nodiscard]] bool EnsureFindDestinationNavigationReadOnlyMode(HWND navHwnd, CaseState& state, std::wstring_view probeName) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture destination navigation mode before {}.", probeName));
    if (! state.failure.empty())
    {
        return false;
    }

    if (! snapshot.destinationNavigationEditMode)
    {
        return true;
    }

    const HWND editHost = ResolveFindDestinationNavigationDxEditHost(navHwnd);
    if (editHost && IsWindow(editHost) != FALSE)
    {
        SendMessageW(editHost, WM_KEYDOWN, VK_ESCAPE, 0);
        SendMessageW(editHost, WM_KEYUP, VK_ESCAPE, 0);
    }
    SendMessageW(navHwnd, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(navHwnd, WM_KEYUP, VK_ESCAPE, 0);
    PumpPendingMessages();

    FindFilesDebugSnapshot after{};
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.destinationNavigationVisible && ! value.destinationNavigationEditMode; },
                                      SelfTest::Scale(1000ms),
                                      &after),
                  std::format(L"Destination navigation should leave edit mode before {}; editMode={}, text='{}'.",
                              probeName,
                              after.destinationNavigationEditMode ? 1 : 0,
                              after.destinationNavigationText));
    return state.failure.empty();
}

[[nodiscard]] bool ProbeFindDestinationNavigationAcceptsNearOwnerQueuedHistoryClick(HWND findWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot),
                  L"Failed to capture Find destination navigation snapshot before near-owner queued history-click probe.");
    state.Require(snapshot.destinationNavigationVisible, L"Find destination navigation bar should be visible before near-owner queued history-click probe.");
    const RECT arrowRect = snapshot.destinationNavigationHistoryRect;
    state.Require(arrowRect.right > arrowRect.left && arrowRect.bottom > arrowRect.top,
                  L"Find destination navigation history arrow should expose a non-empty client rectangle before near-owner queued history-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND navHwnd = ResolveFindDestinationNavigationHwnd(findWindow, arrowRect);
    state.Require(navHwnd && IsWindow(navHwnd) != FALSE, L"Find destination navigation child should exist for near-owner queued history-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(EnsureFindDestinationNavigationReadOnlyMode(navHwnd, state, L"near-owner queued history-click probe"),
                  L"Find destination navigation should leave edit mode before near-owner queued history-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    POINT clickPointInNav{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    static_cast<void>(MapWindowPoints(findWindow, navHwnd, &clickPointInNav, 1u));
    RaiseSelfTestWindowForInput(findWindow);
    SendMessageW(navHwnd, WM_MOUSELEAVE, 0, 0);
    PumpPendingMessages();
    std::this_thread::sleep_for(40ms);
    PumpPendingMessages();

    const uint64_t baselineDropdownOpenCount = snapshot.destinationNavigationHistoryDropdownOpenCount;
    std::atomic<bool> sawPopup{false};
    std::atomic<bool> popupDismissed{false};
    std::atomic<bool> driverDone{false};
    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(findWindow);
            if (! popup || IsWindow(popup) == FALSE)
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }

            sawPopup.store(true, std::memory_order_release);
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                if (IsWindow(popup) == FALSE || IsWindowVisible(popup) == FALSE)
                {
                    popupDismissed.store(true, std::memory_order_release);
                    return;
                }

                std::this_thread::sleep_for(10ms);
            }
            return;
        }
    });

    SendMessageW(navHwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickPointInNav.x, clickPointInNav.y));
    SendMessageW(navHwnd, WM_LBUTTONUP, 0, MAKELPARAM(clickPointInNav.x, clickPointInNav.y));
    if (! state.failure.empty())
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, findWindow, SelfTest::Scale(4500ms));

    FindFilesDebugSnapshot afterPopup{};
    state.Require(DebugGetFindFilesWindowSnapshot(afterPopup),
                  L"Failed to capture destination navigation snapshot after near-owner queued history-click probe.");
    const bool dropdownOpened =
        sawPopup.load(std::memory_order_acquire) || afterPopup.destinationNavigationHistoryDropdownOpenCount > baselineDropdownOpenCount;
    const bool dropdownDismissed =
        popupDismissed.load(std::memory_order_acquire) ||
        (afterPopup.destinationNavigationHistoryDropdownOpenCount > baselineDropdownOpenCount && ! afterPopup.destinationNavigationHistoryDropdownVisible);
    state.Require(dropdownOpened,
                  std::format(L"Find destination navigation should accept an inside-history queued click when the live cursor only drifted to the same-owner "
                              L"fringe; before={}, after={}, visible={}.",
                              baselineDropdownOpenCount,
                              afterPopup.destinationNavigationHistoryDropdownOpenCount,
                              afterPopup.destinationNavigationHistoryDropdownVisible ? 1 : 0));
    state.Require(dropdownDismissed, L"Find destination near-owner queued history-click probe opened a popup that could not be dismissed by Escape.");
    return state.failure.empty();
}

[[nodiscard]] bool ProbeFindDestinationNavigationUsesDeliveredHoverPointAfterLiveCursorLeaves(HWND findWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find destination navigation snapshot before delivered-hover probe.");
    state.Require(snapshot.destinationNavigationVisible, L"Find destination navigation bar should be visible before delivered-hover probe.");
    const RECT arrowRect = snapshot.destinationNavigationHistoryRect;
    state.Require(arrowRect.right > arrowRect.left && arrowRect.bottom > arrowRect.top,
                  L"Find destination navigation history arrow should expose a non-empty client rectangle before delivered-hover probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND navHwnd = ResolveFindDestinationNavigationHwnd(findWindow, arrowRect);
    state.Require(navHwnd && IsWindow(navHwnd) != FALSE, L"Find destination navigation child should exist for delivered-hover probe.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(EnsureFindDestinationNavigationReadOnlyMode(navHwnd, state, L"delivered-hover probe"),
                  L"Find destination navigation should leave edit mode before delivered-hover probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    RaiseSelfTestWindowForInput(findWindow);
    SendMessageW(navHwnd, WM_MOUSELEAVE, 0, 0);
    PumpPendingMessages();

    // No live cursor warp: production hover routing reads the delivered WM_MOUSEMOVE point
    // below (never GetCursorPos), so it is authoritative regardless of the real cursor.
    PumpPendingMessages();
    std::this_thread::sleep_for(40ms);
    PumpPendingMessages();

    FindFilesDebugSnapshot clearedHoverSnapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(clearedHoverSnapshot), L"Failed to capture destination navigation snapshot after delivered-hover setup.");
    state.Require(! clearedHoverSnapshot.destinationNavigationHistoryHovered,
                  L"Destination navigation history hover should be clear before the delivered-hover message is injected.");
    if (! state.failure.empty())
    {
        return false;
    }

    POINT deliveredPointInNav{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    static_cast<void>(MapWindowPoints(findWindow, navHwnd, &deliveredPointInNav, 1u));
    SendMessageW(navHwnd, WM_MOUSEMOVE, 0, MAKELPARAM(deliveredPointInNav.x, deliveredPointInNav.y));

    FindFilesDebugSnapshot after{};
    state.Require(DebugGetFindFilesWindowSnapshot(after), L"Failed to capture Find destination navigation snapshot after delivered-hover probe.");
    state.Require(after.destinationNavigationHistoryHovered && ! after.destinationNavigationMenuHovered && ! after.destinationNavigationDiskHovered &&
                      after.destinationNavigationHoveredSegmentIndex == -1 && after.destinationNavigationHoveredSeparatorIndex == -1,
                  std::format(L"Find destination navigation should honor the delivered hover point even after the live cursor left; "
                              L"deliveredPt=({}, {}), arrowRect=({}, {}, {}, {}), menu={}, history={}, disk={}, segment={}, separator={}.",
                              deliveredPointInNav.x,
                              deliveredPointInNav.y,
                              arrowRect.left,
                              arrowRect.top,
                              arrowRect.right,
                              arrowRect.bottom,
                              after.destinationNavigationMenuHovered ? 1 : 0,
                              after.destinationNavigationHistoryHovered ? 1 : 0,
                              after.destinationNavigationDiskHovered ? 1 : 0,
                              after.destinationNavigationHoveredSegmentIndex,
                              after.destinationNavigationHoveredSeparatorIndex));
    if (! state.failure.empty())
    {
        return false;
    }

    InvalidateRect(navHwnd, nullptr, FALSE);
    UpdateWindow(navHwnd);

    FindFilesDebugSnapshot afterPaint{};
    state.Require(DebugGetFindFilesWindowSnapshot(afterPaint), L"Failed to capture Find destination navigation snapshot after delivered-hover paint probe.");
    state.Require(
        ! afterPaint.destinationNavigationMenuHovered && afterPaint.destinationNavigationHistoryHovered && ! afterPaint.destinationNavigationDiskHovered &&
            afterPaint.destinationNavigationHoveredSegmentIndex == -1 && afterPaint.destinationNavigationHoveredSeparatorIndex == -1,
        std::format(L"Find destination navigation paint must preserve the delivered-hover state; menu={}, history={}, disk={}, segment={}, separator={}.",
                    afterPaint.destinationNavigationMenuHovered ? 1 : 0,
                    afterPaint.destinationNavigationHistoryHovered ? 1 : 0,
                    afterPaint.destinationNavigationDiskHovered ? 1 : 0,
                    afterPaint.destinationNavigationHoveredSegmentIndex,
                    afterPaint.destinationNavigationHoveredSeparatorIndex));
    return state.failure.empty();
}

[[nodiscard]] bool ProbeFindDestinationNavigationUsesDeliveredHistoryClickAfterLiveCursorLeaves(HWND findWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find destination navigation snapshot before delivered-click probe.");
    state.Require(snapshot.destinationNavigationVisible, L"Find destination navigation bar should be visible before delivered-click probe.");
    const RECT arrowRect = snapshot.destinationNavigationHistoryRect;
    state.Require(arrowRect.right > arrowRect.left && arrowRect.bottom > arrowRect.top,
                  L"Find destination navigation history arrow should expose a non-empty client rectangle before delivered-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND navHwnd = ResolveFindDestinationNavigationHwnd(findWindow, arrowRect);
    state.Require(navHwnd && IsWindow(navHwnd) != FALSE, L"Find destination navigation child should exist for delivered-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(EnsureFindDestinationNavigationReadOnlyMode(navHwnd, state, L"delivered-click probe"),
                  L"Find destination navigation should leave edit mode before delivered-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    RaiseSelfTestWindowForInput(findWindow);
    SendMessageW(navHwnd, WM_MOUSELEAVE, 0, 0);
    PumpPendingMessages();

    // No live cursor warp: production click routing reads the delivered point below
    // (never GetCursorPos), so it is authoritative regardless of the real cursor.
    PumpPendingMessages();
    std::this_thread::sleep_for(40ms);
    PumpPendingMessages();

    POINT deliveredPointInNav{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    static_cast<void>(MapWindowPoints(findWindow, navHwnd, &deliveredPointInNav, 1u));

    const uint64_t baselineDropdownOpenCount = snapshot.destinationNavigationHistoryDropdownOpenCount;
    std::atomic<bool> sawPopup{false};
    std::atomic<bool> popupDismissed{false};
    std::atomic<bool> driverDone{false};
    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(4000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(findWindow);
            if (! popup || IsWindow(popup) == FALSE)
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }

            sawPopup.store(true, std::memory_order_release);
            PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);

            const auto closeDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < closeDeadline)
            {
                if (IsWindow(popup) == FALSE || IsWindowVisible(popup) == FALSE)
                {
                    popupDismissed.store(true, std::memory_order_release);
                    return;
                }

                std::this_thread::sleep_for(10ms);
            }
            return;
        }
    });

    SendMessageW(navHwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(deliveredPointInNav.x, deliveredPointInNav.y));
    SendMessageW(navHwnd, WM_LBUTTONUP, 0, MAKELPARAM(deliveredPointInNav.x, deliveredPointInNav.y));
    if (! state.failure.empty())
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, findWindow, SelfTest::Scale(4500ms));

    FindFilesDebugSnapshot afterPopup{};
    state.Require(DebugGetFindFilesWindowSnapshot(afterPopup), L"Failed to capture destination navigation snapshot after delivered-click probe.");
    const bool dropdownOpened =
        sawPopup.load(std::memory_order_acquire) || afterPopup.destinationNavigationHistoryDropdownOpenCount > baselineDropdownOpenCount;
    const bool dropdownDismissed =
        popupDismissed.load(std::memory_order_acquire) ||
        (afterPopup.destinationNavigationHistoryDropdownOpenCount > baselineDropdownOpenCount && ! afterPopup.destinationNavigationHistoryDropdownVisible);
    state.Require(dropdownOpened,
                  std::format(L"Find destination navigation should honor the delivered history click after the live cursor left; before={}, after={}, "
                              L"visible={}.",
                              baselineDropdownOpenCount,
                              afterPopup.destinationNavigationHistoryDropdownOpenCount,
                              afterPopup.destinationNavigationHistoryDropdownVisible ? 1 : 0));
    state.Require(dropdownDismissed, L"Find destination delivered-click probe opened a popup that could not be dismissed by Escape.");
    return state.failure.empty();
}

[[nodiscard]] bool ProbeFindDestinationNavigationUsesDeliveredDoubleClickAfterLiveCursorLeaves(HWND findWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.destinationNavigationVisible && ! value.destinationNavigationEditMode && ! value.destinationNavigationHistoryDropdownVisible &&
               value.destinationNavigationHistoryRect.right > value.destinationNavigationHistoryRect.left &&
               value.destinationNavigationHistoryRect.bottom > value.destinationNavigationHistoryRect.top;
    },
                      SelfTest::Scale(2000ms),
                      &snapshot),
                  std::format(L"Find destination navigation should be stable before delivered-double-click probe. {}", DescribeFindSnapshotBrief(snapshot)));
    const RECT arrowRect = snapshot.destinationNavigationHistoryRect;
    state.Require(arrowRect.right > arrowRect.left && arrowRect.bottom > arrowRect.top,
                  L"Find destination navigation history arrow should expose a non-empty client rectangle before delivered-double-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND navHwnd = ResolveFindDestinationNavigationHwnd(findWindow, arrowRect);
    state.Require(navHwnd && IsWindow(navHwnd) != FALSE, L"Find destination navigation child should exist for delivered-double-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(EnsureFindDestinationNavigationReadOnlyMode(navHwnd, state, L"delivered-double-click probe"),
                  L"Find destination navigation should leave edit mode before delivered-double-click probe.");
    if (! state.failure.empty())
    {
        return false;
    }
    RaiseSelfTestWindowForInput(findWindow);
    SendMessageW(navHwnd, WM_MOUSELEAVE, 0, 0);
    PumpPendingMessages();

    // No live cursor warp: production double-click routing reads the delivered point below
    // (never GetCursorPos), so it is authoritative regardless of the real cursor.
    PumpPendingMessages();
    std::this_thread::sleep_for(40ms);
    PumpPendingMessages();

    const LONG destinationLeft = static_cast<LONG>(std::lround(snapshot.destinationNavigationRect.left));
    POINT deliveredEditPointInNav{std::max<LONG>(destinationLeft + 8, arrowRect.left - 20), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    static_cast<void>(MapWindowPoints(findWindow, navHwnd, &deliveredEditPointInNav, 1u));

    SendMessageW(navHwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(deliveredEditPointInNav.x, deliveredEditPointInNav.y));
    SendMessageW(navHwnd, WM_LBUTTONUP, 0, MAKELPARAM(deliveredEditPointInNav.x, deliveredEditPointInNav.y));
    SendMessageW(navHwnd, WM_LBUTTONDBLCLK, MK_LBUTTON, MAKELPARAM(deliveredEditPointInNav.x, deliveredEditPointInNav.y));
    SendMessageW(navHwnd, WM_LBUTTONUP, 0, MAKELPARAM(deliveredEditPointInNav.x, deliveredEditPointInNav.y));
    PumpPendingMessages();

    FindFilesDebugSnapshot after{};
    state.Require(
        WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return value.destinationNavigationEditMode; }, SelfTest::Scale(1500ms), &after),
        std::format(L"Find destination navigation should honor the delivered double-click after the live cursor left. {}", DescribeFindSnapshotBrief(after)));

    const HWND editHost = ResolveFindDestinationNavigationDxEditHost(navHwnd);
    if (editHost && IsWindow(editHost) != FALSE)
    {
        state.Require(IsWindowVisible(editHost) != FALSE, L"Find destination navigation delivered double-click should leave a visible edit host.");
        SendMessageW(editHost, WM_KEYDOWN, VK_ESCAPE, 0);
        SendMessageW(editHost, WM_KEYUP, VK_ESCAPE, 0);
        PumpPendingMessages();

        FindFilesDebugSnapshot afterEscape{};
        state.Require(DebugGetFindFilesWindowSnapshot(afterEscape), L"Failed to capture destination navigation snapshot after delivered-double-click Escape.");
        state.Require(! afterEscape.destinationNavigationEditMode, L"Find destination navigation delivered double-click Escape should leave edit mode.");
    }
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogResultDrainsRespectQueuedChildInput(CaseState& state) noexcept
{
    PumpPendingMessages();

    wil::unique_hwnd parentWindow{CreateWindowExW(0, L"STATIC", L"", WS_POPUP, 0, 0, 16, 16, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr)};
    state.Require(parentWindow.get() != nullptr, L"Failed to create parent probe window for Find drain queue-order test.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::unique_hwnd childWindow{CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 8, 8, parentWindow.get(), nullptr, GetModuleHandleW(nullptr), nullptr)};
    state.Require(childWindow.get() != nullptr, L"Failed to create child probe window for Find drain queue-order test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        MSG message{};
        while (PeekMessageW(&message, childWindow.get(), 0, 0, PM_REMOVE) != 0)
        {
        }
        while (PeekMessageW(&message, parentWindow.get(), 0, 0, PM_REMOVE) != 0)
        {
        }
        childWindow.reset();
        parentWindow.reset();
    });

    state.Require(PostMessageW(childWindow.get(), WM_MOUSEMOVE, 0, MAKELPARAM(1, 1)) != FALSE,
                  L"Failed to queue child mouse input before Find result message.");
    state.Require(PostMessageW(parentWindow.get(), WndMsg::kFindSearchResults, 0, 0) != FALSE,
                  L"Failed to queue parent Find result message after child input.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! DebugFindFilesIsNextQueuedMessage(parentWindow.get(), WndMsg::kFindSearchResults),
                  L"Find result drains must not skip queued child-window input that is ahead of a result message.");

    MSG message{};
    state.Require(PeekMessageW(&message, childWindow.get(), WM_MOUSEMOVE, WM_MOUSEMOVE, PM_REMOVE) != 0,
                  L"Expected child mouse input to remain at the queue head for the drain-order probe.");
    state.Require(DebugFindFilesIsNextQueuedMessage(parentWindow.get(), WndMsg::kFindSearchResults),
                  L"Find result drain should become eligible once the older child input has been removed.");
    state.Require(PeekMessageW(&message, parentWindow.get(), WndMsg::kFindSearchResults, WndMsg::kFindSearchResults, PM_REMOVE) != 0,
                  L"Expected queued parent Find result message to remain after the child input.");

    return state.failure.empty();
}

[[nodiscard]] std::optional<POINT> ChooseOwnerPointOutsidePopup(HWND ownerWindow, HWND popup) noexcept
{
    RECT popupRect{};
    RECT clientRect{};
    if (GetClientRect(ownerWindow, &clientRect) == FALSE || GetWindowRect(popup, &popupRect) == FALSE)
    {
        return std::nullopt;
    }

    std::array<POINT, 4> candidates{{
        POINT{clientRect.left + 16, clientRect.top + 16},
        POINT{clientRect.right - 16, clientRect.top + 16},
        POINT{clientRect.left + 16, clientRect.bottom - 16},
        POINT{clientRect.right - 16, clientRect.bottom - 16},
    }};
    for (POINT candidate : candidates)
    {
        if (PtInRect(&clientRect, candidate) == FALSE || ClientToScreen(ownerWindow, &candidate) == FALSE)
        {
            continue;
        }
        if (PtInRect(&popupRect, candidate) == FALSE)
        {
            return candidate;
        }
    }

    POINT fallback{popupRect.right + 24, popupRect.top + 24};
    HMONITOR monitor = MonitorFromWindow(ownerWindow, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{sizeof(mi)};
    if (GetMonitorInfoW(monitor, &mi) != FALSE)
    {
        const LONG minX = mi.rcWork.left + 1;
        const LONG minY = mi.rcWork.top + 1;
        const LONG maxX = (std::max)(minX, mi.rcWork.right - 2);
        const LONG maxY = (std::max)(minY, mi.rcWork.bottom - 2);
        fallback.x      = (std::clamp)(fallback.x, minX, maxX);
        fallback.y      = (std::clamp)(fallback.y, minY, maxY);
    }
    return fallback;
}

[[nodiscard]] bool ProbeFindSplitMenuLivePointerRouting(HWND findWindow, HWND ownerWindow, size_t itemIndex, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    std::atomic<bool> sawPopup{false};
    std::atomic<bool> cursorMovedToItem{false};
    std::atomic<bool> hoverObserved{false};
    std::atomic<bool> outsideClickSent{false};
    std::atomic<bool> popupDismissed{false};
    std::atomic<bool> driverDone{false};
    std::wstring hoverFailureDetails;

    std::jthread menuDriver([&](std::stop_token stopToken) noexcept
    {
        const auto markDone = wil::scope_exit([&] { driverDone.store(true, std::memory_order_release); });
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < deadline)
        {
            const HWND popup = FindVisibleOwnedDxUiContextMenuWindowForSearchTest(ownerWindow);
            if (! popup || IsWindow(popup) == FALSE)
            {
                std::this_thread::sleep_for(10ms);
                continue;
            }

            sawPopup.store(true, std::memory_order_release);
            RedSalamander::DxUi::ContextMenuPopupDebugState popupState{};
            D2D1_RECT_F itemRectDip{};
            if (! RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, popupState) ||
                ! RedSalamander::DxUi::DebugGetContextMenuPopupItemRect(popup, itemIndex, itemRectDip))
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }

            POINT itemClient{
                static_cast<LONG>(std::lround((itemRectDip.left + itemRectDip.right) * 0.5f * static_cast<float>(popupState.dpi) /
                                              static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
                static_cast<LONG>(std::lround((itemRectDip.top + itemRectDip.bottom) * 0.5f * static_cast<float>(popupState.dpi) /
                                              static_cast<float>(USER_DEFAULT_SCREEN_DPI))),
            };
            POINT itemCenter = itemClient;
            if (ClientToScreen(popup, &itemCenter) == FALSE)
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }

            // The routing contract under test is the delivered WM_MOUSEMOVE point (production
            // never samples GetCursorPos), so the synthetic pointer message below is authoritative
            // and no live cursor move is needed.
            RECT popupRect{};
            if (GetWindowRect(popup, &popupRect) == FALSE)
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                return;
            }
            const POINT deliveredPoint{itemCenter.x - popupRect.left, itemCenter.y - popupRect.top};
            const LPARAM deliveredMove =
                MAKELPARAM(static_cast<WORD>(static_cast<SHORT>(deliveredPoint.x)), static_cast<WORD>(static_cast<SHORT>(deliveredPoint.y)));

            const auto hoverDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < hoverDeadline)
            {
                if (PostMessageW(popup, WM_MOUSEMOVE, 0, deliveredMove) != FALSE)
                {
                    cursorMovedToItem.store(true, std::memory_order_release);
                }
                std::this_thread::sleep_for(25ms);
                RedSalamander::DxUi::ContextMenuPopupDebugState hoverState{};
                RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState paintState{};
                if (RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, hoverState) && hoverState.hoveredIndex == std::optional<size_t>{itemIndex} &&
                    RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popup, itemIndex, paintState) && paintState.usesHighlightFill)
                {
                    hoverObserved.store(true, std::memory_order_release);
                    break;
                }
            }
            if (! hoverObserved.load(std::memory_order_acquire))
            {
                RedSalamander::DxUi::ContextMenuPopupDebugState finalState{};
                static_cast<void>(RedSalamander::DxUi::DebugGetContextMenuPopupState(popup, finalState));
                RedSalamander::DxUi::ContextMenuPopupItemPaintDebugState finalPaint{};
                static_cast<void>(RedSalamander::DxUi::DebugGetContextMenuPopupItemPaint(popup, itemIndex, finalPaint));
                hoverFailureDetails = std::format(L"Pointer movement over the Find split action menu did not produce hover highlight. "
                                                  L"(hovered={}, keyboard={}, center=({},{}), delivered=({},{}), popup=({},{} {}x{}), "
                                                  L"row=({:.1f},{:.1f},{:.1f},{:.1f}), highlight={}, renders={})",
                                                  finalState.hoveredIndex.has_value() ? static_cast<long long>(finalState.hoveredIndex.value()) : -1ll,
                                                  finalState.keyboardIndex.has_value() ? static_cast<long long>(finalState.keyboardIndex.value()) : -1ll,
                                                  itemCenter.x,
                                                  itemCenter.y,
                                                  deliveredPoint.x,
                                                  deliveredPoint.y,
                                                  popupRect.left,
                                                  popupRect.top,
                                                  popupRect.right - popupRect.left,
                                                  popupRect.bottom - popupRect.top,
                                                  itemRectDip.left,
                                                  itemRectDip.top,
                                                  itemRectDip.right,
                                                  itemRectDip.bottom,
                                                  finalPaint.usesHighlightFill ? L"true" : L"false",
                                                  finalState.renderCount);
            }

            const std::optional<POINT> outsidePoint = ChooseOwnerPointOutsidePopup(ownerWindow, popup);
            if (outsidePoint.has_value())
            {
                outsideClickSent.store(PostSelfTestOutsideClickToPopup(popup, outsidePoint.value()), std::memory_order_release);
            }

            const auto dismissDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
            while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < dismissDeadline)
            {
                if (IsWindow(popup) == FALSE)
                {
                    popupDismissed.store(true, std::memory_order_release);
                    return;
                }

                std::this_thread::sleep_for(10ms);
            }

            if (IsWindow(popup) != FALSE)
            {
                PostMessageW(popup, WM_KEYDOWN, VK_ESCAPE, 0);
                PostMessageW(popup, WM_KEYUP, VK_ESCAPE, 0);
                const auto escapeDismissDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                while (! stopToken.stop_requested() && std::chrono::steady_clock::now() < escapeDismissDeadline)
                {
                    if (IsWindow(popup) == FALSE)
                    {
                        popupDismissed.store(true, std::memory_order_release);
                        return;
                    }

                    std::this_thread::sleep_for(10ms);
                }
            }
            return;
        }
    });

    SendFindSplitMenuClick(findWindow, state);
    if (! state.failure.empty())
    {
        menuDriver.request_stop();
    }
    JoinSearchPopupDriverWithUiPumping(menuDriver, driverDone, ownerWindow, SelfTest::Scale(6000ms));

    state.Require(sawPopup.load(std::memory_order_acquire), L"Find split action menu did not open for live pointer routing validation.");
    state.Require(cursorMovedToItem.load(std::memory_order_acquire), L"Failed to route pointer movement over the Find split action menu item.");
    state.Require(hoverObserved.load(std::memory_order_acquire),
                  hoverFailureDetails.empty() ? L"Pointer movement over the Find split action menu did not produce hover highlight." : hoverFailureDetails);
    state.Require(outsideClickSent.load(std::memory_order_acquire), L"Failed to send outside click for Find split action menu light-dismiss validation.");
    state.Require(popupDismissed.load(std::memory_order_acquire), L"Find split action menu did not dismiss after outside-click and Escape cleanup.");
    return state.failure.empty();
}

[[nodiscard]] bool GetConfiguredLocalFileSystemForSelfTest(CaseState& state, std::string_view configurationJson, CreatedFileSystemInstance& created) noexcept
{
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    if (! configurationJson.empty())
    {
        wil::com_ptr<IInformations> informations;
        state.Require(CreateInformations(created.fileSystem, informations), L"Isolated local file system instance missing IInformations.");
        if (! informations)
        {
            return false;
        }

        const HRESULT configHr = informations->SetConfiguration(configurationJson.data());
        state.Require(SUCCEEDED(configHr),
                      std::format(L"Failed to configure isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(configHr)));
        if (FAILED(configHr))
        {
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool SearchServiceBackendSelectableForRoot(const std::filesystem::path& root, CaseState& state, std::wstring_view caseLabel) noexcept
{
    LocalSearchIndexCore::Repository repository;
    LocalSearchIndexCore::SupportInfo support{};
    const HRESULT probeHr = repository.ProbePath(root.native(), support);
    if (FAILED(probeHr))
    {
        state.Skip(std::format(L"{} requires an index-support probeable local root. ProbePath failed for '{}' with hr=0x{:08X}.",
                               caseLabel,
                               root.native(),
                               static_cast<unsigned long>(probeHr)));
        return false;
    }

    if (! support.indexable)
    {
        state.Skip(std::format(L"{} requires an indexable local root so the file-system plugin can select the search-service backend. Root '{}' is not "
                               L"indexable on this machine.",
                               caseLabel,
                               root.native()));
        return false;
    }

    return true;
}

[[nodiscard]] wil::com_ptr<IFileSystemSearch> QueryLocalFileSystemSearchForSelfTest(CaseState& state, IFileSystem* fs) noexcept
{
    if (! fs)
    {
        state.Require(false, L"Local file-system plugin is not loaded for search selftest.");
        return {};
    }

    wil::com_ptr<IFileSystemSearch> search;
    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemSearch), search.put_void());
    state.Require(SUCCEEDED(hr) && search != nullptr,
                  std::format(L"Local file-system plugin does not expose IFileSystemSearch: 0x{:08X}.", static_cast<unsigned long>(hr)));
    return search;
}

[[nodiscard]] wil::com_ptr<IFileSystemSearch> GetLocalFileSystemSearchForSelfTest(CaseState& state) noexcept
{
    constexpr std::wstring_view kLocalPluginId = L"builtin/file-system";

    wil::com_ptr<IFileSystem> fs = SelfTest::GetFileSystem(kLocalPluginId);
    if (! fs)
    {
        const HRESULT enableHr = FileSystemPluginManager::GetInstance().EnablePlugin(kLocalPluginId, g_settings);
        state.Require(SUCCEEDED(enableHr),
                      std::format(L"Failed to enable local file-system plugin for search selftest: 0x{:08X}.", static_cast<unsigned long>(enableHr)));
        fs = SelfTest::GetFileSystem(kLocalPluginId);
    }

    state.Require(fs != nullptr, L"Local file-system plugin is not loaded for search selftest.");
    if (! fs)
    {
        return {};
    }

    return QueryLocalFileSystemSearchForSelfTest(state, fs.get());
}

[[nodiscard]] bool TestLocalPluginInvalidRegexReportsSingleCompletion(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path caseRoot = suiteRoot / L"work" / (L"search_invalid_regex_" + NewGuidText());
    state.Require(SelfTest::RemoveAll(caseRoot), L"Failed to clear the indexed-search long-path fixture root.");
    const auto cleanup = wil::scope_exit([&]
    {
        if (! SelfTest::RemoveAll(caseRoot))
        {
            Debug::Warning(L"Commands selftest: failed to remove indexed-search long-path fixture '{}'.", caseRoot.wstring());
        }
    });

    state.Require(SelfTest::EnsureDirectory(caseRoot), L"Failed to create invalid-regex search root.");
    state.Require(SelfTest::WriteTextFile(caseRoot / L"candidate.txt", "needle"), L"Failed to create invalid-regex candidate file.");
    if (! state.failure.empty())
    {
        return false;
    }

    wil::com_ptr<IFileSystemSearch> search = GetLocalFileSystemSearchForSelfTest(state);
    if (! search || ! state.failure.empty())
    {
        return false;
    }

    const auto requireInvalidRegexReportedOnce = [&](std::wstring_view label,
                                                     FileSystemSearchNameMode nameMode,
                                                     const wchar_t* namePattern,
                                                     FileSystemSearchContentMode contentMode,
                                                     const wchar_t* contentPattern) noexcept
    {
        LocalSearchProbeCallback callback{};
        FileSystemSearchQuery query{};
        query.sizeBytes              = sizeof(FileSystemSearchQuery);
        query.rootPath               = caseRoot.c_str();
        query.namePattern            = namePattern;
        query.contentPattern         = contentPattern;
        query.flags                  = FILESYSTEM_SEARCH_INCLUDE_FILES;
        query.nameMode               = nameMode;
        query.contentMode            = contentMode;
        query.maxContentBytesPerFile = 0;
        query.maxSnippetCharacters   = 0;

        const HRESULT hr = search->Search(&query, &callback, nullptr);
        state.Require(hr == E_INVALIDARG, std::format(L"{} invalid regex should return E_INVALIDARG, got 0x{:08X}.", label, static_cast<unsigned long>(hr)));
        state.Require(callback.matchCallbacks.load(std::memory_order_acquire) == 0u,
                      std::format(L"{} invalid regex should not emit matches, got {}.", label, callback.matchCallbacks.load(std::memory_order_acquire)));
        state.Require(callback.invalidProgressPayloads.load(std::memory_order_acquire) == 0u,
                      std::format(L"{} invalid regex emitted malformed progress payloads (count={}).",
                                  label,
                                  callback.invalidProgressPayloads.load(std::memory_order_acquire)));
        state.Require(callback.progressCallbacks.load(std::memory_order_acquire) == 1u,
                      std::format(L"{} invalid regex should emit exactly one progress callback, got {}.",
                                  label,
                                  callback.progressCallbacks.load(std::memory_order_acquire)));
        state.Require(callback.completedProgressCallbacks.load(std::memory_order_acquire) == 1u,
                      std::format(L"{} invalid regex should emit exactly one completed progress callback, got {}.",
                                  label,
                                  callback.completedProgressCallbacks.load(std::memory_order_acquire)));
        state.Require((callback.warningFlags.load(std::memory_order_acquire) & FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED) != 0u,
                      std::format(L"{} invalid regex did not include REGEX_REJECTED warning flags (0x{:08X}).",
                                  label,
                                  callback.warningFlags.load(std::memory_order_acquire)));
        state.Require(callback.lastCompletedStatus.load(std::memory_order_acquire) == E_INVALIDARG,
                      std::format(L"{} invalid regex completed status should be E_INVALIDARG, got 0x{:08X}.",
                                  label,
                                  static_cast<unsigned long>(callback.lastCompletedStatus.load(std::memory_order_acquire))));
    };

    requireInvalidRegexReportedOnce(L"name", FILESYSTEM_SEARCH_NAME_REGEX, L"[", FILESYSTEM_SEARCH_CONTENT_DISABLED, nullptr);
    requireInvalidRegexReportedOnce(L"content", FILESYSTEM_SEARCH_NAME_WILDCARD, L"*", FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX, L"[");
    requireInvalidRegexReportedOnce(L"unsafe name", FILESYSTEM_SEARCH_NAME_REGEX, L"(a+)+", FILESYSTEM_SEARCH_CONTENT_DISABLED, nullptr);
    requireInvalidRegexReportedOnce(L"duplicate-alternative unsafe name", FILESYSTEM_SEARCH_NAME_REGEX, L"(a|a)*", FILESYSTEM_SEARCH_CONTENT_DISABLED, nullptr);
    requireInvalidRegexReportedOnce(L"prefix-overlap unsafe name", FILESYSTEM_SEARCH_NAME_REGEX, L"(a|ab)*x", FILESYSTEM_SEARCH_CONTENT_DISABLED, nullptr);

    return state.failure.empty();
}

[[nodiscard]] bool TestLocalPluginParallelSearchCancellationAndFanIn(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path caseRoot = suiteRoot / L"work" / (L"search_parallel_cancel_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(caseRoot, ec);
    const auto cleanup = wil::scope_exit([&]
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(caseRoot, cleanupEc);
    });

    state.Require(SelfTest::EnsureDirectory(caseRoot), L"Failed to create parallel-search root.");
    if (! state.failure.empty())
    {
        return false;
    }

    uint32_t expectedMatches = 0u;
    for (uint32_t index = 0u; index < 160u; ++index)
    {
        const std::filesystem::path path = caseRoot / std::format(L"match_flat_{:03}.txt", index);
        state.Require(SelfTest::WriteTextFile(path, "flat"), std::format(L"Failed to create {}.", path.wstring()));
        ++expectedMatches;
    }

    uint32_t expectedContentMatches = 0u;
    for (uint32_t index = 0u; index < 640u; ++index)
    {
        const std::filesystem::path path = caseRoot / std::format(L"content_match_{:03}.txt", index);
        state.Require(SelfTest::WriteTextFile(path, "needle"), std::format(L"Failed to create {}.", path.wstring()));
        ++expectedContentMatches;
    }

    const std::filesystem::path broadRoot = caseRoot / L"broad";
    state.Require(SelfTest::EnsureDirectory(broadRoot), L"Failed to create broad search root.");
    for (uint32_t dirIndex = 0u; dirIndex < 12u; ++dirIndex)
    {
        const std::filesystem::path dir = broadRoot / std::format(L"dir_{:02}", dirIndex);
        state.Require(SelfTest::EnsureDirectory(dir), std::format(L"Failed to create {}.", dir.wstring()));
        for (uint32_t fileIndex = 0u; fileIndex < 4u; ++fileIndex)
        {
            const std::filesystem::path path = dir / std::format(L"match_broad_{:02}_{:02}.txt", dirIndex, fileIndex);
            state.Require(SelfTest::WriteTextFile(path, "broad"), std::format(L"Failed to create {}.", path.wstring()));
            ++expectedMatches;
        }
    }

    std::filesystem::path deepDir = caseRoot / L"deep";
    state.Require(SelfTest::EnsureDirectory(deepDir), L"Failed to create deep search root.");
    for (uint32_t level = 0u; level < 8u; ++level)
    {
        deepDir /= std::format(L"d{:02}", level);
        state.Require(SelfTest::EnsureDirectory(deepDir), std::format(L"Failed to create {}.", deepDir.wstring()));
        const std::filesystem::path path = deepDir / std::format(L"match_deep_{:02}.txt", level);
        state.Require(SelfTest::WriteTextFile(path, "deep"), std::format(L"Failed to create {}.", path.wstring()));
        ++expectedMatches;
    }

    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    constexpr std::string_view kScanFourWalkerConfig = R"json({"searchBackendPreference":"scan","searchMaxDirectoryWalkers":4})json";
    if (! GetConfiguredLocalFileSystemForSelfTest(state, kScanFourWalkerConfig, created))
    {
        return false;
    }

    wil::com_ptr<IFileSystemSearch> search = QueryLocalFileSystemSearchForSelfTest(state, created.fileSystem.get());
    if (! search || ! state.failure.empty())
    {
        return false;
    }

    FileSystemSearchQuery query{};
    query.sizeBytes              = sizeof(FileSystemSearchQuery);
    query.rootPath               = caseRoot.c_str();
    query.namePattern            = L"match_*.txt";
    query.flags                  = static_cast<FileSystemSearchFlags>(FILESYSTEM_SEARCH_INCLUDE_FILES | FILESYSTEM_SEARCH_RECURSIVE);
    query.nameMode               = FILESYSTEM_SEARCH_NAME_WILDCARD;
    query.contentMode            = FILESYSTEM_SEARCH_CONTENT_DISABLED;
    query.maxContentBytesPerFile = 0;
    query.maxSnippetCharacters   = 0;

    LocalSearchProbeCallback fullCallback{};
    const HRESULT fullHr = search->Search(&query, &fullCallback, nullptr);
    state.Require(SUCCEEDED(fullHr), std::format(L"Full parallel scan should succeed, got 0x{:08X}.", static_cast<unsigned long>(fullHr)));
    state.Require(fullCallback.matchCallbacks.load(std::memory_order_acquire) == expectedMatches,
                  std::format(L"Full parallel scan should find {} matches across flat, broad, and deep trees; got {}.",
                              expectedMatches,
                              fullCallback.matchCallbacks.load(std::memory_order_acquire)));
    state.Require(
        fullCallback.concurrentCallbacks.load(std::memory_order_acquire) == 0u,
        std::format(L"Search callbacks must be serialized; saw {} concurrent entries.", fullCallback.concurrentCallbacks.load(std::memory_order_acquire)));
    state.Require(fullCallback.completedProgressCallbacks.load(std::memory_order_acquire) == 1u,
                  std::format(L"Full parallel scan should emit one completed progress callback, got {}.",
                              fullCallback.completedProgressCallbacks.load(std::memory_order_acquire)));
    state.Require(fullCallback.lastCompletedStatus.load(std::memory_order_acquire) == S_OK,
                  std::format(L"Full parallel scan completion status should be S_OK, got 0x{:08X}.",
                              static_cast<unsigned long>(fullCallback.lastCompletedStatus.load(std::memory_order_acquire))));

    LocalSearchProbeCallback cancelCallback{};
    cancelCallback.cancelAfterMatchCallbacks.store(1u, std::memory_order_release);
    const HRESULT cancelHr = search->Search(&query, &cancelCallback, nullptr);
    state.Require(cancelHr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Parallel scan should honor cancellation after first match, got 0x{:08X}.", static_cast<unsigned long>(cancelHr)));
    state.Require(cancelCallback.matchCallbacks.load(std::memory_order_acquire) <= 4u,
                  std::format(L"Parallel scan emitted too many matches after cancellation became visible: {}.",
                              cancelCallback.matchCallbacks.load(std::memory_order_acquire)));
    state.Require(cancelCallback.shouldCancelCallbacks.load(std::memory_order_acquire) > 0u,
                  L"Parallel scan did not poll FileSystemSearchShouldCancel during the match fan-in path.");
    state.Require(cancelCallback.concurrentCallbacks.load(std::memory_order_acquire) == 0u,
                  std::format(L"Cancelled parallel scan callbacks must be serialized; saw {} concurrent entries.",
                              cancelCallback.concurrentCallbacks.load(std::memory_order_acquire)));
    state.Require(cancelCallback.completedProgressCallbacks.load(std::memory_order_acquire) == 1u,
                  std::format(L"Cancelled parallel scan should emit one completed progress callback, got {}.",
                              cancelCallback.completedProgressCallbacks.load(std::memory_order_acquire)));
    state.Require(cancelCallback.lastCompletedStatus.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Cancelled parallel scan completion status should be ERROR_CANCELLED, got 0x{:08X}.",
                              static_cast<unsigned long>(cancelCallback.lastCompletedStatus.load(std::memory_order_acquire))));

    FileSystemSearchQuery contentQuery = query;
    contentQuery.namePattern           = L"content_match_*.txt";
    contentQuery.contentPattern        = L"needle";
    contentQuery.contentMode           = FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL;

    LocalSearchProbeCallback fullContentCallback{};
    const HRESULT fullContentHr = search->Search(&contentQuery, &fullContentCallback, nullptr);
    state.Require(SUCCEEDED(fullContentHr),
                  std::format(L"Full parallel content scan should succeed, got 0x{:08X}.", static_cast<unsigned long>(fullContentHr)));
    state.Require(fullContentCallback.matchCallbacks.load(std::memory_order_acquire) == expectedContentMatches,
                  std::format(L"Full parallel content scan should find {} matches, got {}.",
                              expectedContentMatches,
                              fullContentCallback.matchCallbacks.load(std::memory_order_acquire)));
    state.Require(fullContentCallback.completedProgressCallbacks.load(std::memory_order_acquire) == 1u,
                  std::format(L"Full parallel content scan should emit one completed progress callback, got {}.",
                              fullContentCallback.completedProgressCallbacks.load(std::memory_order_acquire)));

    LocalSearchProbeCallback cancelContentCallback{};
    cancelContentCallback.cancelAfterMatchCallbacks.store(1u, std::memory_order_release);
    const HRESULT cancelContentHr = search->Search(&contentQuery, &cancelContentCallback, nullptr);
    state.Require(
        cancelContentHr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
        std::format(L"Parallel content scan should honor cancellation during result drain, got 0x{:08X}.", static_cast<unsigned long>(cancelContentHr)));
    state.Require(cancelContentCallback.matchCallbacks.load(std::memory_order_acquire) <= 4u,
                  std::format(L"Parallel content scan emitted too many matches after cancellation became visible: {}.",
                              cancelContentCallback.matchCallbacks.load(std::memory_order_acquire)));
    state.Require(cancelContentCallback.shouldCancelCallbacks.load(std::memory_order_acquire) > 0u,
                  L"Parallel content scan did not poll FileSystemSearchShouldCancel during result drain.");
    state.Require(cancelContentCallback.completedProgressCallbacks.load(std::memory_order_acquire) == 1u,
                  std::format(L"Cancelled parallel content scan should emit one completed progress callback, got {}.",
                              cancelContentCallback.completedProgressCallbacks.load(std::memory_order_acquire)));
    state.Require(cancelContentCallback.lastCompletedStatus.load(std::memory_order_acquire) == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Cancelled parallel content scan completion status should be ERROR_CANCELLED, got 0x{:08X}.",
                              static_cast<unsigned long>(cancelContentCallback.lastCompletedStatus.load(std::memory_order_acquire))));

    return state.failure.empty();
}

[[nodiscard]] bool TestLocalPluginWatchUnwatchDrainsInflightCallback(CaseState& state) noexcept
{
    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path caseRoot = suiteRoot / L"work" / (L"watch_unwatch_drain_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(caseRoot, ec);
    const auto cleanup = wil::scope_exit([&]
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(caseRoot, cleanupEc);
    });

    state.Require(SelfTest::EnsureDirectory(caseRoot), L"Failed to create watch-drain root.");
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    if (! GetConfiguredLocalFileSystemForSelfTest(state, {}, created))
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryWatch> watch;
    const HRESULT watchQiHr = created.fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryWatch), watch.put_void());
    state.Require(SUCCEEDED(watchQiHr) && watch != nullptr,
                  std::format(L"Local file-system plugin does not expose IFileSystemDirectoryWatch: 0x{:08X}.", static_cast<unsigned long>(watchQiHr)));
    if (! watch || ! state.failure.empty())
    {
        return false;
    }

    BlockingDirectoryWatchCallback callback{};
    callback.blockFirstCallback.store(true, std::memory_order_release);

    const HRESULT watchHr = watch->WatchDirectory(caseRoot.c_str(), &callback, nullptr);
    state.Require(SUCCEEDED(watchHr), std::format(L"WatchDirectory failed: 0x{:08X}.", static_cast<unsigned long>(watchHr)));
    if (FAILED(watchHr))
    {
        return false;
    }

    state.Require(SelfTest::WriteTextFile(caseRoot / L"first.txt", "first"), L"Failed to create first watched file.");
    state.Require(WaitForFlag(callback.firstCallbackEntered, 5000u), L"Directory watch callback did not start after creating a file.");
    if (! state.failure.empty())
    {
        callback.allowFirstCallbackReturn.store(true, std::memory_order_release);
        static_cast<void>(watch->UnwatchDirectory(caseRoot.c_str()));
        return false;
    }

    std::jthread releaseThread([&callback]
    {
        ::Sleep(150);
        callback.allowFirstCallbackReturn.store(true, std::memory_order_release);
    });

    const ULONGLONG unwatchStartTick = ::GetTickCount64();
    const HRESULT unwatchHr          = watch->UnwatchDirectory(caseRoot.c_str());
    const ULONGLONG unwatchElapsed   = ::GetTickCount64() - unwatchStartTick;
    releaseThread.join();

    state.Require(SUCCEEDED(unwatchHr), std::format(L"UnwatchDirectory failed: 0x{:08X}.", static_cast<unsigned long>(unwatchHr)));
    state.Require(callback.firstCallbackExited.load(std::memory_order_acquire), L"UnwatchDirectory returned before the in-flight callback could exit.");
    state.Require(unwatchElapsed >= 100u, std::format(L"UnwatchDirectory did not appear to drain the blocked callback; elapsed={}ms.", unwatchElapsed));
    state.Require(callback.invalidPayloads.load(std::memory_order_acquire) == 0u,
                  std::format(L"Directory watch emitted malformed payloads (count={}).", callback.invalidPayloads.load(std::memory_order_acquire)));

    const uint32_t callbacksAfterUnwatch = callback.callbacks.load(std::memory_order_acquire);
    state.Require(SelfTest::WriteTextFile(caseRoot / L"after_unwatch.txt", "after"), L"Failed to create post-unwatch file.");
    ::Sleep(250);
    state.Require(callback.callbacks.load(std::memory_order_acquire) == callbacksAfterUnwatch,
                  std::format(L"Directory watch emitted callbacks after UnwatchDirectory returned: before={}, after={}.",
                              callbacksAfterUnwatch,
                              callback.callbacks.load(std::memory_order_acquire)));

    return state.failure.empty();
}

[[nodiscard]] bool TestLocalSearchIndexEnumerateStopsAfterFirstCandidate(CaseState& state) noexcept
{
    const SelfTest::TestSandbox sandbox =
        SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::Commands, L"search_index_stream");
    state.Require(sandbox.IsValid(), L"Indexed-search stream TestSandbox root unavailable.");
    if (! sandbox.IsValid())
    {
        return false;
    }

    const std::filesystem::path caseRoot   = sandbox.root;
    const std::filesystem::path dataRoot   = caseRoot / L"data";
    const std::filesystem::path nestedRoot = dataRoot / L"nested";
    constexpr size_t kSnapshotPaddingLength = 96u;
    constexpr size_t kLegacyMaxPath          = static_cast<size_t>(MAX_PATH);
    const std::wstring snapshotPadding(kSnapshotPaddingLength, L's');
    const std::filesystem::path snapshotRoot = caseRoot / L"snapshots" / snapshotPadding / snapshotPadding;

    std::error_code ec;
    std::filesystem::remove_all(caseRoot, ec);
    const auto cleanup = wil::scope_exit([&]
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(caseRoot, cleanupEc);
    });

    state.Require(SelfTest::EnsureDirectory(nestedRoot), L"Failed to create indexed-search nested folder.");
    state.Require(snapshotRoot.native().size() > kLegacyMaxPath,
                  std::format(L"Indexed-search snapshot fixture must exceed MAX_PATH; got {} characters.", snapshotRoot.native().size()));
    for (int index = 0; index < 12; ++index)
    {
        const std::filesystem::path folder = index < 6 ? dataRoot : nestedRoot;
        const std::filesystem::path file   = folder / std::format(L"match_{:02}.txt", index);
        state.Require(SelfTest::WriteTextFile(file, "streamed candidate"), std::format(L"Failed to create {}.", file.filename().native()));
    }
    state.Require(SelfTest::WriteTextFile(dataRoot / L"ignore.bin", "ignored"), L"Failed to create non-matching indexed-search file.");
    if (! state.failure.empty())
    {
        return false;
    }

    LocalSearchIndexCore::Repository repository({.snapshotRootDirectory = snapshotRoot.native()});
    LocalSearchIndexCore::SupportInfo support{};
    HRESULT hr = repository.ProbePath(dataRoot.native(), support);
    state.Require(SUCCEEDED(hr), std::format(L"ProbePath failed for indexed-search stream test: 0x{:08X}.", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(support.indexable, L"Indexed-search stream test requires an indexable local volume.");
    if (! support.indexable)
    {
        return false;
    }

    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = dataRoot.native();
    plan.namePattern        = L"match_*.txt";
    plan.nameMode           = FILESYSTEM_SEARCH_NAME_WILDCARD;
    plan.recursive          = true;
    plan.includeFiles       = true;
    plan.includeDirectories = false;

    EnumerateStopAfterFirstState callbackState{};
    LocalSearchIndexCore::QueryStats stats{};
    hr = repository.Enumerate(plan, nullptr, nullptr, &StopAfterFirstIndexedCandidate, &callbackState, &stats);
    state.Require(hr == S_OK, std::format(L"Enumerate should stop cleanly after the first candidate, got 0x{:08X}.", static_cast<unsigned long>(hr)));
    state.Require(callbackState.seenCandidates == 1u,
                  std::format(L"Expected Enumerate callback to see exactly one candidate, got {}.", callbackState.seenCandidates));
    state.Require(stats.candidateCount == 1u, std::format(L"Expected indexed candidate count to stop at 1, got {}.", stats.candidateCount));
    state.Require(stats.fileCount >= 12u, std::format(L"Expected indexed file count to include the prepared files, got {}.", stats.fileCount));
    state.Require(stats.snapshotSaved, L"Long-path indexed search should atomically save its snapshot.");
    state.Require(stats.snapshotFileBytes > 0u, L"Long-path indexed search should report the saved snapshot byte size.");
    state.Require(stats.snapshotPath.size() > kLegacyMaxPath,
                  std::format(L"Indexed-search snapshot path should exceed MAX_PATH; got {} characters.", stats.snapshotPath.size()));
    if (hr != S_OK)
    {
        return false;
    }

    LocalSearchIndexCore::Repository reloadRepository({.snapshotRootDirectory = snapshotRoot.native()});
    EnumerateStopAfterFirstState reloadCallbackState{};
    LocalSearchIndexCore::QueryStats reloadStats{};
    hr = reloadRepository.Enumerate(plan, nullptr, nullptr, &StopAfterFirstIndexedCandidate, &reloadCallbackState, &reloadStats);
    state.Require(hr == S_OK,
                  std::format(L"Long-path snapshot reload should stop cleanly after the first candidate, got 0x{:08X}.",
                              static_cast<unsigned long>(hr)));
    state.Require(reloadCallbackState.seenCandidates == 1u,
                  std::format(L"Expected long-path snapshot reload callback to see exactly one candidate, got {}.", reloadCallbackState.seenCandidates));
    state.Require(reloadStats.candidateCount == 1u,
                  std::format(L"Expected long-path snapshot reload candidate count to stop at 1, got {}.", reloadStats.candidateCount));
    state.Require(reloadStats.fileCount >= 12u,
                  std::format(L"Expected long-path snapshot reload file count to include the prepared files, got {}.", reloadStats.fileCount));
    state.Require(reloadStats.snapshotLoaded, L"Second indexed search should load the persisted long-path snapshot.");
    state.Require(reloadStats.snapshotFileBytes > 0u, L"Long-path snapshot reload should report the persisted snapshot byte size.");
    if (hr != S_OK)
    {
        return false;
    }

    hr = reloadRepository.CorruptSnapshotForTests(dataRoot.native(), LocalSearchIndexCore::SnapshotCorruptionMode::InvalidMagic);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Long-path snapshot corruption fixture failed with 0x{:08X}.", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    hr = reloadRepository.DropCachedVolumeForTests(dataRoot.native());
    state.Require(SUCCEEDED(hr),
                  std::format(L"Long-path snapshot cache drop failed with 0x{:08X}.", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    EnumerateStopAfterFirstState recoveredCallbackState{};
    LocalSearchIndexCore::QueryStats recoveredStats{};
    hr = reloadRepository.Enumerate(plan, nullptr, nullptr, &StopAfterFirstIndexedCandidate, &recoveredCallbackState, &recoveredStats);
    state.Require(hr == S_OK,
                  std::format(L"Long-path corrupt snapshot recovery should stop cleanly after the first candidate, got 0x{:08X}.",
                              static_cast<unsigned long>(hr)));
    state.Require(recoveredCallbackState.seenCandidates == 1u,
                  std::format(L"Expected long-path recovery callback to see exactly one candidate, got {}.", recoveredCallbackState.seenCandidates));
    state.Require(recoveredStats.rebuiltSnapshotCorruption, L"Long-path corrupt snapshot should be detected and rebuilt.");
    state.Require(recoveredStats.snapshotSaved, L"Long-path corrupt snapshot recovery should atomically replace the snapshot.");
    state.Require(recoveredStats.snapshotFileBytes > 0u, L"Long-path corrupt snapshot recovery should report the replacement snapshot byte size.");
    if (hr != S_OK)
    {
        return false;
    }

    hr = reloadRepository.InvalidateRoot(dataRoot.native(), true);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Long-path snapshot invalidation and deletion failed with 0x{:08X}.", static_cast<unsigned long>(hr)));
    return state.failure.empty();
}

[[nodiscard]] bool TestLocalProviderReselectAfterIndexStreamStaysResponsive(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const auto restorePane                                = wil::scope_exit([&]
    {
        if (! leftPluginBefore.empty())
        {
            static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        }
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"search_local_provider_reselect_after_index_stream: begin");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"search_local_provider_reselect_after_index_stream: watch-drain setup begin");
    if (! TestLocalPluginWatchUnwatchDrainsInflightCallback(state))
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"search_local_provider_reselect_after_index_stream: index-stream setup begin");
    if (! TestLocalSearchIndexEnumerateStopsAfterFirstCandidate(state))
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"search_local_provider_reselect_after_index_stream: predecessor setup done");
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"search_local_provider_reselect_after_index_stream: setting left pane local provider");
    const HRESULT hr = g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system");
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                               std::format(L"search_local_provider_reselect_after_index_stream: provider returned 0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(SUCCEEDED(hr), std::format(L"Local provider reselect failed after search-index stream selftest: 0x{:08X}.", static_cast<unsigned long>(hr)));
    const ULONGLONG pumpUntil = GetTickCount64() + static_cast<ULONGLONG>(SelfTest::Scale(50ms).count());
    while (GetTickCount64() < pumpUntil)
    {
        PumpPendingMessages();
        Sleep(1);
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"search_local_provider_reselect_after_index_stream: done");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogSearchOps(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || ! IsWindow(mainWindow))
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    ShortcutManager shortcuts;
    shortcuts.Load(ShortcutDefaults::CreateDefaultShortcuts());
    const auto findChordOpt = shortcuts.TryGetShortcutForCommand(L"cmd/pane/find");
    state.Require(findChordOpt.has_value(), L"Find default shortcut missing.");
    if (findChordOpt.has_value())
    {
        state.Require(findChordOpt->vk == VK_F7, L"Find default shortcut expected F7.");
        state.Require(findChordOpt->modifiers == ShortcutManager::kModAlt, L"Find default shortcut expected Alt+F7.");
    }
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: start");

    const std::optional<std::filesystem::path> leftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const auto restorePaths                                = wil::scope_exit([&]
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
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

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"find_dialog_" + NewGuidText());
    const std::filesystem::path sub  = root / L"sub";
    const std::filesystem::path bulk = root / L"bulk";

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, std::format(L"find_dialog_search_ops: creating fixtures root='{}'", root.native()));
    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create find dialog subdirectory.");
    state.Require(SelfTest::EnsureDirectory(bulk), L"Failed to create find dialog bulk directory.");
    state.Require(SelfTest::WriteTextFile(root / L"a.jsonl", "{\"message\":\"needle alpha\"}"), L"Failed to create a.jsonl.");
    state.Require(SelfTest::WriteTextFile(root / L"b.txt", "plain text"), L"Failed to create b.txt.");
    state.Require(SelfTest::WriteTextFile(sub / L"c.jsonl", "{\"message\":\"other\"}"), L"Failed to create c.jsonl.");
    state.Require(SelfTest::WriteTextFile(sub / L"d.txt", "content needle beta"), L"Failed to create d.txt.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: fixture files ready");

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: setting left pane local provider");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for find dialog test.");
    if (! state.failure.empty())
    {
        return false;
    }

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

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: navigating left pane");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for find dialog test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(3000ms)), L"Enumeration did not complete for find dialog test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.jsonl", L"b.txt", L"sub", L"bulk"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for find dialog test.");
    if (! state.failure.empty())
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: left pane ready");

    if (const HWND existing = GetFindFilesWindowHandle(); existing && IsWindow(existing))
    {
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: closing existing Find window");
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Find window did not close before find dialog test.");
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: dispatching find command");
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open from cmd/pane/find.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: Find window opened");

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, true), L"Failed to configure Find window options.");

    const auto waitMs = [](std::chrono::milliseconds value) noexcept -> uint32_t { return static_cast<uint32_t>(value.count()); };

    auto requireIdle = [&](std::wstring_view label) noexcept
    { state.Require(DebugWaitForFindFilesWindowIdle(waitMs(SelfTest::Scale(10000ms))), std::format(L"Find window did not become idle after {}.", label)); };

    struct RenderSnapshot
    {
        uint64_t dxRenderCount  = 0u;
        uint64_t gridPaintCount = 0u;
    };

    struct RenderSample
    {
        std::wstring label;
        std::wstring action;
        uint64_t waitUs         = 0u;
        uint64_t dxRenderDelta  = 0u;
        uint64_t gridPaintDelta = 0u;
    };

    std::vector<RenderSample> renderSamples;

    auto captureRenderCounts = [&]() noexcept -> RenderSnapshot
    {
        FindFilesDebugSnapshot snapshot{};
        if (! DebugGetFindFilesWindowSnapshot(snapshot))
        {
            return {};
        }
        return {.dxRenderCount = snapshot.dxRenderCount, .gridPaintCount = snapshot.resultGridPaintCount};
    };

    auto requireRenderedAfter = [&](RenderSnapshot before, std::wstring_view label) noexcept
    {
        std::wstring renderAction = L"idle";
        FindFilesDebugSnapshot current{};
        if (DebugGetFindFilesWindowSnapshot(current))
        {
            static_cast<void>(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid));
            if (current.resultListHasVerticalScrollbar && current.visibleResultRowCount < current.resultCount)
            {
                if (DebugScrollFindFilesWindowResultsByWheelDetents(-1))
                {
                    renderAction = L"scroll";
                }
            }
            else if (! current.fullPaths.empty() && DebugSelectFindFilesWindowResult(current.fullPaths.front()))
            {
                renderAction = L"select";
            }
        }

        const auto waitStartedAt = std::chrono::steady_clock::now();
        FindFilesDebugSnapshot rendered{};
        static_cast<void>(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& snapshot) noexcept {
            return snapshot.dxRenderCount > before.dxRenderCount || snapshot.resultGridPaintCount > before.gridPaintCount;
        }, SelfTest::Scale(3000ms), &rendered));
        const uint64_t waitUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - waitStartedAt).count());
        renderSamples.push_back(RenderSample{
            .label          = std::wstring(label),
            .action         = std::move(renderAction),
            .waitUs         = waitUs,
            .dxRenderDelta  = rendered.dxRenderCount >= before.dxRenderCount ? rendered.dxRenderCount - before.dxRenderCount : 0u,
            .gridPaintDelta = rendered.resultGridPaintCount >= before.gridPaintCount ? rendered.resultGridPaintCount - before.gridPaintCount : 0u,
        });
    };

    auto requireSnapshotCount = [&](size_t expectedCount, std::wstring_view label, std::initializer_list<std::filesystem::path> expectedPaths) noexcept
    {
        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot after {}.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(snapshot.resultCount == expectedCount,
                      std::format(L"Unexpected Find result count after {}. Expected {}, got {}.", label, expectedCount, snapshot.resultCount));
        for (const auto& path : expectedPaths)
        {
            const std::wstring expected = path.native();
            const bool found            = std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), expected) != snapshot.fullPaths.end();
            state.Require(found, std::format(L"Expected Find results after {} to contain '{}'.", label, expected));
        }
    };

    auto requireSortedResultNames = [&](size_t expectedCount, std::wstring_view label) noexcept
    {
        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture sorted Find snapshot after {}.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(snapshot.resultCount == expectedCount,
                      std::format(L"Unexpected sorted Find result count after {}. Expected {}, got {}.", label, expectedCount, snapshot.resultCount));
        if (! state.failure.empty())
        {
            return;
        }

        for (size_t index = 1; index < snapshot.fullPaths.size(); ++index)
        {
            const std::wstring previousName = std::filesystem::path(snapshot.fullPaths[index - 1]).filename().native();
            const std::wstring currentName  = std::filesystem::path(snapshot.fullPaths[index]).filename().native();
            state.Require(! OrdinalString::LessNoCase(currentName, previousName),
                          std::format(L"Find results should remain sorted by name after {}. '{}' appeared after '{}'.", label, currentName, previousName));
            if (! state.failure.empty())
            {
                return;
            }
        }
    };

    const RenderSnapshot findRenderBefore = captureRenderCounts();
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting Find");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for wildcard search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find operation.");
    requireIdle(L"Find");
    requireRenderedAfter(findRenderBefore, L"Find");
    requireSnapshotCount(2u, L"Find", {root / L"a.jsonl", sub / L"c.jsonl"});

    const RenderSnapshot appendRenderBefore = captureRenderCounts();
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting Append");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for append search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Append), L"Failed to start Append operation.");
    requireIdle(L"Append");
    requireRenderedAfter(appendRenderBefore, L"Append");
    requireSnapshotCount(4u, L"Append", {root / L"a.jsonl", root / L"b.txt", sub / L"c.jsonl", sub / L"d.txt"});

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting invalid-regex Intersect");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"[", L"", Common::Settings::SearchNameMode::Regex, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for invalid-regex intersect search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Intersect), L"Failed to start invalid-regex Intersect operation.");
    requireIdle(L"invalid-regex Intersect");

    FindFilesDebugSnapshot invalidIntersect{};
    state.Require(DebugGetFindFilesWindowSnapshot(invalidIntersect), L"Failed to capture invalid-regex Intersect snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(
        invalidIntersect.lastStatusHint == E_INVALIDARG,
        std::format(L"Expected invalid-regex Intersect to surface E_INVALIDARG, got 0x{:08X}.", static_cast<unsigned long>(invalidIntersect.lastStatusHint)));
    state.Require((invalidIntersect.warningFlags & FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED) != 0u,
                  std::format(L"Invalid-regex Intersect should report REGEX_REJECTED warning flags, got 0x{:08X}.",
                              static_cast<unsigned long>(invalidIntersect.warningFlags)));
    requireSnapshotCount(4u, L"invalid-regex Intersect", {root / L"a.jsonl", root / L"b.txt", sub / L"c.jsonl", sub / L"d.txt"});

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting Intersect");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"a*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for intersect search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Intersect), L"Failed to start Intersect operation.");
    requireIdle(L"Intersect");
    requireSnapshotCount(1u, L"Intersect", {root / L"a.jsonl"});

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting Subtract");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Subtract), L"Failed to start Subtract operation.");
    requireIdle(L"Subtract");
    requireSnapshotCount(0u, L"Subtract", {});

    std::wstring wildcardAllWithTranslatedControl = L"*";
    wildcardAllWithTranslatedControl.push_back(static_cast<wchar_t>(0x7F));
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting wildcard-all Find");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                std::move(wildcardAllWithTranslatedControl),
                                                L"",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for wildcard-all search with a translated control character.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start wildcard-all Find operation.");
    requireIdle(L"wildcard-all Find");
    requireSnapshotCount(4u, L"wildcard-all Find", {root / L"a.jsonl", root / L"b.txt", sub / L"c.jsonl", sub / L"d.txt"});

    FindFilesDebugSnapshot wildcardAllSnapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(wildcardAllSnapshot), L"Failed to capture wildcard-all Find snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(wildcardAllSnapshot.namePatternText == L"*",
                  std::format(L"Find dialog should sanitize translated control characters from wildcard-all patterns; got length {}.",
                              wildcardAllSnapshot.namePatternText.size()));

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting zero-match Intersect");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"missing-*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for zero-match intersect search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Intersect), L"Failed to start zero-match Intersect operation.");
    requireIdle(L"zero-match Intersect");
    requireSnapshotCount(0u, L"zero-match Intersect", {});

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting content Find");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"", L"needle", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for content search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, true), L"Failed to restore Find options for content search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start content Find operation.");
    requireIdle(L"content Find");
    requireSnapshotCount(2u, L"content Find", {root / L"a.jsonl", sub / L"d.txt"});

    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int i = 0; i < 48; ++i)
    {
        const std::filesystem::path bulkFile = bulk / std::format(L"bulk_{:03}.txt", i);
        state.Require(SelfTest::WriteTextFile(bulkFile, largeBody), std::format(L"Failed to create {}.", bulkFile.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: bulk fixtures ready");
    const RenderSnapshot sortedRenderBefore = captureRenderCounts();
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting sorted Find");
    state.Require(DebugSetFindFilesWindowResultSort(0u, false), L"Failed to enable ascending Name sort for Find results.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for sorted search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for sorted search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start sorted Find operation.");
    requireIdle(L"sorted Find");
    requireRenderedAfter(sortedRenderBefore, L"sorted Find");
    requireSortedResultNames(50u, L"sorted Find");

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting cancelled Intersect");
    state.Require(
        DebugConfigureFindFilesWindow(
            root.native(), L"*.txt", L"ZZZ_NOT_PRESENT_123456789", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextRegex),
        L"Failed to configure Find window for cancelled intersect search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for cancelled intersect search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Intersect), L"Failed to start cancelled Intersect search.");
    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Intersect search.");
    requireIdle(L"cancelled Intersect");

    FindFilesDebugSnapshot cancelledIntersect{};
    state.Require(DebugGetFindFilesWindowSnapshot(cancelledIntersect), L"Failed to capture cancelled Intersect snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(cancelledIntersect.searchActive == false, L"Find window remained active after cancelled Intersect.");
    requireSnapshotCount(50u, L"cancelled Intersect", {});

    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"find_dialog_search_ops: starting cancellation Find");
    state.Require(
        DebugConfigureFindFilesWindow(
            root.native(), L"*.txt", L"ZZZ_NOT_PRESENT_123456789", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextRegex),
        L"Failed to configure Find window for cancellation search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for cancellation search.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start cancellation search.");
    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search.");
    requireIdle(L"cancelled Find");

    FindFilesDebugSnapshot cancelled{};
    state.Require(DebugGetFindFilesWindowSnapshot(cancelled), L"Failed to capture cancelled Find snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(cancelled.searchActive == false, L"Find window remained active after cancellation.");
    state.Require(cancelled.lastStatusHint == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Expected cancelled Find HRESULT 0x{:08X}, got 0x{:08X}.",
                              static_cast<unsigned long>(HRESULT_FROM_WIN32(ERROR_CANCELLED)),
                              static_cast<unsigned long>(cancelled.lastStatusHint)));

    for (const auto& sample : renderSamples)
    {
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                                   std::format(L"find_dialog_search_ops render sample label='{}' action='{}' waitUs={} dxRenderDelta={} gridPaintDelta={}",
                                               sample.label,
                                               sample.action,
                                               sample.waitUs,
                                               sample.dxRenderDelta,
                                               sample.gridPaintDelta));
    }
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, std::format(L"find_dialog_search_ops render sampleCount={}", renderSamples.size()));

    PostMessageW(findWindow, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)), L"Find window did not close after operations test.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogTypedLocalRootOverridesStaleContext(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            static_cast<void>(SendMessageW(findWindow, WM_CLOSE, 0, 0));
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(5000ms)));
        }
    };
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_local_override";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create override test subdirectory.");
    state.Require(SelfTest::WriteTextFile(root / L"first.jsonl", "{\"message\":\"alpha\"}"), L"Failed to create first.jsonl.");
    state.Require(SelfTest::WriteTextFile(sub / L"second.jsonl", "{\"message\":\"beta\"}"), L"Failed to create second.jsonl.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext staleContext{};
    staleContext.pluginId        = L"builtin/file-system-gdrive";
    staleContext.pluginShortId   = L"gdrive";
    staleContext.instanceContext = L"@stale";
    staleContext.rootPluginPath  = std::filesystem::path(L"gdrive://@stale");

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-local-override");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(staleContext)), L"Failed to open Find window with stale context.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for stale-context override test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for stale-context override test.");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for stale-context override test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for stale-context override test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for stale-context override test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture stale-context override snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.lastStatusHint == S_OK,
                  std::format(L"Expected stale-context override search to succeed, got 0x{:08X}.", static_cast<unsigned long>(snapshot.lastStatusHint)));
    state.Require(snapshot.resultCount == 2u, std::format(L"Expected 2 results for stale-context override search, got {}.", snapshot.resultCount));

    for (const std::filesystem::path& expectedPath : {root / L"first.jsonl", sub / L"second.jsonl"})
    {
        const std::wstring expected = expectedPath.native();
        const bool found            = std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), expected) != snapshot.fullPaths.end();
        state.Require(found, std::format(L"Expected stale-context override results to contain '{}'.", expected));
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogOpensFromFocusedPaneAndAllowsMultipleInstances(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore                      = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore                     = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        g_settings.search = previousSearch;
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    CloseAllFindFilesWindowsForSearchTest();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path leftRoot  = suiteRoot / L"work" / L"find_focused_left";
    const std::filesystem::path rightRoot = suiteRoot / L"work" / L"find_focused_right";
    std::error_code ec;
    std::filesystem::remove_all(leftRoot, ec);
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(leftRoot), L"Failed to create left focused Find root.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create right focused Find root.");
    state.Require(SelfTest::WriteTextFile(leftRoot / L"left.txt", "left"), L"Failed to create left focused Find fixture.");
    state.Require(SelfTest::WriteTextFile(rightRoot / L"right.txt", "right"), L"Failed to create right focused Find fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = (suiteRoot / L"work" / L"stale_find_root").native();
    search.nameMode                               = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode                            = Common::Settings::SearchContentMode::Disabled;
    g_settings.search                             = search;

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set left pane to local file system for focused Find test.");
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set right pane to local file system for focused Find test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftRoot);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, leftRoot, SelfTest::Scale(3000ms)), L"Left pane did not reach focused Find test root.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)), L"Right pane did not reach focused Find test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"First cmd/pane/find dispatch failed.");
    const HWND firstFindWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(firstFindWindow != nullptr && IsWindow(firstFindWindow) != FALSE, L"First Find window did not open.");

    FindFilesDebugSnapshot firstSnapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return value.rootText == leftRoot.native(); },
                                      SelfTest::Scale(3000ms),
                                      &firstSnapshot),
                  std::format(L"First Find window should use focused left pane root, not stale settings. {}", DescribeFindSnapshotBrief(firstSnapshot)));
    state.Require(DebugGetFindFilesWindowCount() == 1u, std::format(L"Expected one Find window after first open, got {}.", DebugGetFindFilesWindowCount()));
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
    FocusFolderViewPane(FolderWindow::Pane::Right);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Second cmd/pane/find dispatch failed.");
    const HWND secondFindWindow = WaitForWindow(
        [&]() noexcept
    {
        const HWND hwnd = GetFindFilesWindowHandle();
        return hwnd && hwnd != firstFindWindow ? hwnd : nullptr;
    },
        SelfTest::Scale(5000ms));
    state.Require(secondFindWindow != nullptr && IsWindow(secondFindWindow) != FALSE, L"Second Find window did not open as a separate instance.");

    FindFilesDebugSnapshot secondSnapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return value.rootText == rightRoot.native(); },
                                      SelfTest::Scale(3000ms),
                                      &secondSnapshot),
                  std::format(L"Second Find window should use focused right pane root. {}", DescribeFindSnapshotBrief(secondSnapshot)));
    state.Require(DebugGetFindFilesWindowCount() >= 2u,
                  std::format(L"Find command should allow multiple modeless instances; got {}.", DebugGetFindFilesWindowCount()));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRecursiveLocalSearchAndIndexAvailability(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore                      = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore                     = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        g_settings.search = previousSearch;
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    CloseAllFindFilesWindowsForSearchTest();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_recursive_local_mode";
    const std::filesystem::path sub  = root / L"sub";
    const std::filesystem::path deep = sub / L"deep";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(deep), L"Failed to create recursive Find test folders.");
    state.Require(SelfTest::WriteTextFile(root / L"root-hit.txt", "root"), L"Failed to create root recursive Find fixture.");
    state.Require(SelfTest::WriteTextFile(sub / L"sub-hit.txt", "sub"), L"Failed to create sub recursive Find fixture.");
    state.Require(SelfTest::WriteTextFile(deep / L"deep-hit.txt", "deep"), L"Failed to create deep recursive Find fixture.");
    state.Require(SelfTest::WriteTextFile(deep / L"ignored.log", "ignored"), L"Failed to create recursive Find ignored fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND findWindow = nullptr;
    std::optional<std::filesystem::path> ignoredLeftBefore;
    state.Require(OpenFindWindowFromLocalPaneRoot(mainWindow, root, {L"root-hit.txt", L"sub"}, findWindow, ignoredLeftBefore),
                  L"Find window did not open from local recursive test root.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(
            L"relative-root-is-not-indexable", L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to set non-indexable root text for Find index preference validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to request Prefer indexed backend for non-indexable root validation.");
    FindFilesDebugSnapshot disabledSnapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(disabledSnapshot), L"Failed to capture disabled index preference Find snapshot.");
    state.Require(! disabledSnapshot.preferIndexEnabled && ! disabledSnapshot.preferIndexChecked,
                  L"Prefer indexed backend must be disabled and unchecked when the current root cannot use the indexed backend.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure recursive local Find search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to force recursive local scan options.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start recursive local Find search.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Recursive local Find search did not become idle.");

    const std::vector<std::wstring> expectedPaths = {
        (root / L"root-hit.txt").native(),
        (sub / L"sub-hit.txt").native(),
        (deep / L"deep-hit.txt").native(),
    };
    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        if (value.resultCount != expectedPaths.size())
        {
            return false;
        }
        return std::all_of(expectedPaths.begin(), expectedPaths.end(), [&](const std::wstring& expected) noexcept {
            return std::find(value.fullPaths.begin(), value.fullPaths.end(), expected) != value.fullPaths.end();
        });
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Recursive local Find should include root, subfolder, and deep files. {}", DescribeFindSnapshotBrief(snapshot)));
    state.Require(std::find(snapshot.resultPathTexts.begin(), snapshot.resultPathTexts.end(), L"sub") != snapshot.resultPathTexts.end(),
                  L"Find Path column should display the immediate subfolder for nested results.");
    state.Require(std::find(snapshot.resultPathTexts.begin(), snapshot.resultPathTexts.end(), (std::filesystem::path(L"sub") / L"deep").native()) !=
                      snapshot.resultPathTexts.end(),
                  L"Find Path column should display multi-level subfolders for deep nested results.");
    state.Require(snapshot.backend == FILESYSTEM_SEARCH_BACKEND_SCAN,
                  std::format(L"Unchecked Find indexed-backend preference must force local scan; backend was {}.", snapshot.backend));
    state.Require(
        snapshot.resultIconIndices.size() == expectedPaths.size(),
        std::format(L"Find should expose one shell icon index per result; got {} for {} results.", snapshot.resultIconIndices.size(), expectedPaths.size()));
    state.Require(std::all_of(snapshot.resultIconIndices.begin(), snapshot.resultIconIndices.end(), [](int iconIndex) noexcept { return iconIndex >= 0; }),
                  L"Find results should resolve real shell icon indices instead of generic text glyphs.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogDestinationNavigationStaleEditHostHitTesting(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> leftBefore                      = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> rightBefore                     = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        g_settings.search = previousSearch;
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    CloseAllFindFilesWindowsForSearchTest();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / L"find_destination_nav_stale_edit";
    const std::filesystem::path rightRoot = suiteRoot / L"work" / L"find_destination_nav_stale_edit_target";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    std::filesystem::remove_all(rightRoot, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find destination navigation source root.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create Find destination navigation target root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create Find destination navigation fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND findWindow = nullptr;
    std::optional<std::filesystem::path> ignoredLeftBefore;
    state.Require(OpenFindWindowFromLocalPaneRoot(mainWindow, root, {L"alpha.txt"}, findWindow, ignoredLeftBefore),
                  L"Find window did not open from destination navigation stale edit-host test root.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set right pane to local file-system plugin for destination navigation stale edit-host test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)),
                  L"Right pane did not reach destination navigation stale edit-host test root.");
    state.Require(DebugSetFindFilesWindowDestinationPath(rightRoot.native()), L"Failed to set explicit Find destination for stale edit-host navigation test.");

    FindFilesDebugSnapshot readySnapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.destinationNavigationVisible && value.destinationNavigationEmbedded &&
               value.destinationNavigationText.find(rightRoot.native()) != std::wstring::npos && value.destinationNavigationHistoryCount > 0u &&
               value.destinationNavigationHistoryRect.right > value.destinationNavigationHistoryRect.left &&
               value.destinationNavigationHistoryRect.bottom > value.destinationNavigationHistoryRect.top;
    },
                      SelfTest::Scale(3000ms),
                      &readySnapshot),
                  std::format(L"Find destination navigation did not become ready for stale edit-host hit-test. {}", DescribeFindSnapshotBrief(readySnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ProbeFindDestinationNavigationUsesDeliveredHoverPointAfterLiveCursorLeaves(findWindow, state),
                  L"Find destination navigation should use delivered hover points after the live cursor left the embedded control.");
    state.Require(ProbeFindDestinationNavigationUsesDeliveredHistoryClickAfterLiveCursorLeaves(findWindow, state),
                  L"Find destination navigation should use delivered history clicks after the live cursor left the embedded control.");
    state.Require(ProbeFindDestinationNavigationUsesDeliveredDoubleClickAfterLiveCursorLeaves(findWindow, state),
                  L"Find destination navigation should use delivered double-clicks after the live cursor left the embedded control.");
    state.Require(ProbeFindDestinationNavigationAcceptsNearOwnerQueuedHistoryClick(findWindow, state),
                  L"Find destination navigation should accept queued history clicks when the live cursor is still in the same-owner fringe.");
    state.Require(ProbeFindDestinationHistoryMenuFromActiveEditMode(findWindow, state),
                  L"Find destination navigation should keep history hover/click routing live while the embedded path editor is active.");
    state.Require(ProbeFindDestinationHistoryMenuThroughStaleEditHost(findWindow, state),
                  L"Find destination navigation should retire a stale edit host before routing history input to the embedded NavigationView.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogPartialCompletionRemovesOnlyKnownSources(CaseState& state) noexcept
{
    const std::array<FindFilesDebugSourceOutcome, 3> partialOutcomes = {
        FindFilesDebugSourceOutcome{.sourceIndex = 0u, .status = S_OK},
        FindFilesDebugSourceOutcome{.sourceIndex = 1u, .status = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)},
        FindFilesDebugSourceOutcome{.sourceIndex = 2u, .status = S_FALSE},
    };
    const std::vector<size_t> partialCompleted = DebugSelectKnownCompletedFindFilesSourceIndicesForTests(
        4u, partialOutcomes, HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY));
    state.Require(partialCompleted == std::vector<size_t>{0u},
                  L"Find partial completion should remove only the source with a known S_OK outcome.");

    const std::vector<size_t> fullCompleted = DebugSelectKnownCompletedFindFilesSourceIndicesForTests(3u, {}, S_OK);
    state.Require(fullCompleted == std::vector<size_t>({0u, 1u, 2u}),
                  L"Find full completion without per-source callbacks should preserve the legacy all-complete fallback.");

    const std::array<FindFilesDebugSourceOutcome, 1> uncertainOutcome = {
        FindFilesDebugSourceOutcome{.sourceIndex = 7u, .status = S_OK},
    };
    state.Require(DebugSelectKnownCompletedFindFilesSourceIndicesForTests(2u, uncertainOutcome, E_FAIL).empty(),
                  L"Find completion should preserve rows for out-of-range or otherwise uncertain source outcomes.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogResultShortcutsUseShellClipboardAndFileActions(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    if (SkipIfFindSelfTestClipboardUnavailable(state))
    {
        return true;
    }

    const std::optional<std::filesystem::path> leftBefore                      = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto viewersBefore                                                   = g_settings.fileActions.viewers;
    const auto editorsBefore                                                   = g_settings.fileActions.editors;
    const std::optional<Common::Settings::ShortcutsSettings> shortcutsBefore   = g_settings.shortcuts;
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        static_cast<void>(CloseFileOperationsPopupForSelfTest(g_folderWindow.DebugGetFileOperationState()));
        g_settings.fileActions.viewers = viewersBefore;
        g_settings.fileActions.editors = editorsBefore;
        g_settings.shortcuts           = shortcutsBefore;
        DebugReloadShortcutsFromSettings();
        g_settings.search = previousSearch;
        HostClearTestPromptResultOverride();
        HostResetTestPromptRequestCount();
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    CloseAllFindFilesWindowsForSearchTest();
    state.Require(CloseFileOperationsPopupForSelfTest(g_folderWindow.DebugGetFileOperationState()),
                  L"Failed to quiesce pre-existing file operations before Find shortcut file-action validation.");
    state.Require(CloseNonMainTopLevelWindowsForSelfTest(GetCurrentProcessId(), mainWindow, SelfTest::Scale(3000ms)),
                  L"Failed to close inherited top-level windows before Find shortcut file-action validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root             = suiteRoot / L"work" / L"find_result_shortcuts";
    const std::filesystem::path rightRoot        = suiteRoot / L"work" / L"find_result_shortcuts_destination";
    const std::filesystem::path explicitRoot     = suiteRoot / L"work" / L"find_result_shortcuts_explicit_destination";
    const std::filesystem::path file             = root / L"alpha.findshortcut";
    const std::filesystem::path explicitCopyFile = root / L"explicit-copy.findshortcut";
    const std::filesystem::path copyFile         = root / L"copy.findshortcut";
    const std::filesystem::path deleteFile       = root / L"delete.findshortcut";
    const std::filesystem::path moveFile         = root / L"move.findshortcut";
    const std::filesystem::path permanentFile    = root / L"permanent.findshortcut";
    const std::filesystem::path nestedCutDirA    = root / L"sub-a";
    const std::filesystem::path nestedCutDirB    = root / L"sub-b";
    const std::filesystem::path nestedCutFileA   = nestedCutDirA / L"nested-cut-a.findshortcut";
    const std::filesystem::path nestedCutFileB   = nestedCutDirB / L"nested-cut-b.findshortcut";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    std::filesystem::remove_all(rightRoot, ec);
    std::filesystem::remove_all(explicitRoot, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find shortcut test root.");
    state.Require(SelfTest::EnsureDirectory(rightRoot), L"Failed to create Find shortcut destination root.");
    state.Require(SelfTest::EnsureDirectory(explicitRoot), L"Failed to create Find shortcut explicit destination root.");
    state.Require(SelfTest::EnsureDirectory(nestedCutDirA), L"Failed to create Find shortcut nested cut directory A.");
    state.Require(SelfTest::EnsureDirectory(nestedCutDirB), L"Failed to create Find shortcut nested cut directory B.");
    state.Require(SelfTest::WriteTextFile(file, "shortcut"), L"Failed to create Find shortcut test file.");
    state.Require(SelfTest::WriteTextFile(explicitCopyFile, "explicit-copy"), L"Failed to create Find shortcut explicit copy test file.");
    state.Require(SelfTest::WriteTextFile(copyFile, "copy"), L"Failed to create Find shortcut copy test file.");
    state.Require(SelfTest::WriteTextFile(deleteFile, "delete"), L"Failed to create Find shortcut delete test file.");
    state.Require(SelfTest::WriteTextFile(moveFile, "move"), L"Failed to create Find shortcut move test file.");
    state.Require(SelfTest::WriteTextFile(permanentFile, "permanent"), L"Failed to create Find shortcut permanent-delete test file.");
    state.Require(SelfTest::WriteTextFile(nestedCutFileA, "nested-a"), L"Failed to create Find shortcut nested cut file A.");
    state.Require(SelfTest::WriteTextFile(nestedCutFileB, "nested-b"), L"Failed to create Find shortcut nested cut file B.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts();
    auto bindShortcut    = [](std::vector<Common::Settings::ShortcutBinding>& bindings, uint32_t vk, uint32_t modifiers, std::wstring_view commandId) noexcept
    {
        RemoveTestShortcutBinding(bindings, vk, modifiers);
        Common::Settings::ShortcutBinding binding{};
        binding.vk        = vk;
        binding.modifiers = modifiers;
        binding.commandId = commandId;
        bindings.push_back(std::move(binding));
    };
    bindShortcut(g_settings.shortcuts.value().folderView, static_cast<uint32_t>(L'Q'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardCopy");
    bindShortcut(g_settings.shortcuts.value().folderView, static_cast<uint32_t>(L'W'), ShortcutManager::kModCtrl, L"cmd/pane/clipboardCut");
    bindShortcut(g_settings.shortcuts.value().folderView, static_cast<uint32_t>(L'D'), ShortcutManager::kModCtrl, L"cmd/pane/moveToRecycleBin");
    bindShortcut(g_settings.shortcuts.value().folderView,
                 static_cast<uint32_t>(L'D'),
                 ShortcutManager::kModCtrl | ShortcutManager::kModShift,
                 L"cmd/pane/permanentDelete");
    bindShortcut(g_settings.shortcuts.value().functionBar, VK_F9, 0u, L"cmd/pane/view");
    bindShortcut(g_settings.shortcuts.value().functionBar, VK_F9, ShortcutManager::kModCtrl, L"cmd/pane/alternateView");
    bindShortcut(g_settings.shortcuts.value().functionBar, VK_F11, 0u, L"cmd/pane/edit");
    bindShortcut(g_settings.shortcuts.value().functionBar, VK_F11, ShortcutManager::kModCtrl, L"cmd/pane/alternateEdit");

    g_settings.fileActions.viewers = Common::Settings::ViewerFileActionsSettings{};
    Common::Settings::FileActionDefinition viewer{};
    viewer.id               = L"find-shortcut-viewer";
    viewer.displayName      = L"Find Shortcut Viewer";
    viewer.enabled          = true;
    viewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    viewer.executablePath   = ResolveCommandProcessorPath();
    viewer.arguments        = L"/C if exist {FullPath} echo viewer>find-view-marker.txt";
    viewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(viewer, {L".findshortcut"});
    g_settings.fileActions.viewers.actions.push_back(std::move(viewer));

    Common::Settings::FileActionDefinition alternateViewer{};
    alternateViewer.id               = L"find-shortcut-alternate-viewer";
    alternateViewer.displayName      = L"Find Shortcut Alternate Viewer";
    alternateViewer.enabled          = true;
    alternateViewer.kind             = Common::Settings::FileActionKind::ExternalProgram;
    alternateViewer.executablePath   = ResolveCommandProcessorPath();
    alternateViewer.arguments        = L"/C if exist {FullPath} echo alternate-viewer>find-alt-view-marker.txt";
    alternateViewer.workingDirectory = L"{Path}";
    TestSetActionExtensions(alternateViewer, {L".findshortcut"});
    g_settings.fileActions.viewers.actions.push_back(std::move(alternateViewer));
    g_settings.fileActions.viewers.associations.push_back(TestViewerAssociation(L".findshortcut", L"find-shortcut-viewer", L"find-shortcut-alternate-viewer"));

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    Common::Settings::FileActionDefinition editor{};
    editor.id               = L"find-shortcut-editor";
    editor.displayName      = L"Find Shortcut Editor";
    editor.enabled          = true;
    editor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    editor.executablePath   = ResolveCommandProcessorPath();
    editor.arguments        = L"/C if exist {FullPath} echo editor>find-edit-marker.txt";
    editor.workingDirectory = L"{Path}";
    TestSetActionExtensions(editor, {L".findshortcut"});
    g_settings.fileActions.editors.actions.push_back(std::move(editor));

    Common::Settings::FileActionDefinition alternateEditor{};
    alternateEditor.id               = L"find-shortcut-alternate-editor";
    alternateEditor.displayName      = L"Find Shortcut Alternate Editor";
    alternateEditor.enabled          = true;
    alternateEditor.kind             = Common::Settings::FileActionKind::ExternalProgram;
    alternateEditor.executablePath   = ResolveCommandProcessorPath();
    alternateEditor.arguments        = L"/C if exist {FullPath} echo alternate-editor>find-alt-edit-marker.txt";
    alternateEditor.workingDirectory = L"{Path}";
    TestSetActionExtensions(alternateEditor, {L".findshortcut"});
    g_settings.fileActions.editors.actions.push_back(std::move(alternateEditor));
    g_settings.fileActions.editors.associations.push_back(TestEditorAssociation(L".findshortcut", L"find-shortcut-editor", L"find-shortcut-alternate-editor"));

    HWND findWindow = nullptr;
    std::optional<std::filesystem::path> ignoredLeftBefore;
    state.Require(
        OpenFindWindowFromLocalPaneRoot(mainWindow,
                                        root,
                                        {L"alpha.findshortcut", L"copy.findshortcut", L"delete.findshortcut", L"move.findshortcut", L"permanent.findshortcut"},
                                        findWindow,
                                        ignoredLeftBefore),
        L"Find window did not open from shortcut test root.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Right, L"builtin/file-system")),
                  L"Failed to set right pane to local file-system plugin for Find shortcut move test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Right, rightRoot, SelfTest::Scale(3000ms)),
                  L"Right pane did not reach Find shortcut move destination root.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.findshortcut", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find shortcut search.");
    state.Require(DebugSetFindFilesWindowOptions(false, true, false, false, false), L"Failed to configure Find shortcut options.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find shortcut search.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())), L"Find shortcut search did not become idle.");
    const auto findSnapshotContainsPath = [](const FindFilesDebugSnapshot& value, const std::filesystem::path& path) noexcept
    {
        return std::find_if(value.fullPaths.begin(), value.fullPaths.end(), [&](const std::wstring& fullPath) noexcept {
                   return OrdinalString::EqualsNoCasePath(std::wstring_view(fullPath), path);
               }) != value.fullPaths.end();
    };
    const auto selectFindShortcutResult = [&](const std::filesystem::path& path, std::wstring_view context) noexcept -> bool
    {
        FindFilesDebugSnapshot availableSnapshot{};
        const bool present = WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return findSnapshotContainsPath(value, path); },
                                                 SelfTest::Scale(3000ms),
                                                 &availableSnapshot);
        if (! present)
        {
            state.Require(false,
                          std::format(L"Find shortcut result '{}' was not available before {}. {}",
                                      path.native(),
                                      context,
                                      DescribeFindSnapshotBrief(availableSnapshot)));
            return false;
        }

        const bool selected = DebugSelectFindFilesWindowResult(path.native());
        FindFilesDebugSnapshot selectedSnapshot{};
        const bool selectionSettled = selected && WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.selectedResultCount == 1u && OrdinalString::EqualsNoCasePath(std::wstring_view(value.selectedResultFullPath), path);
        },
                                                                      SelfTest::Scale(1500ms),
                                                                      &selectedSnapshot);
        if (! selectionSettled)
        {
            static_cast<void>(DebugGetFindFilesWindowSnapshot(selectedSnapshot));
            state.Require(false,
                          std::format(L"Find shortcut result '{}' did not become the stable single selection before {}; selectedCall={}. {}",
                                      path.native(),
                                      context,
                                      selected ? 1 : 0,
                                      DescribeFindSnapshotBrief(selectedSnapshot)));
            return false;
        }

        return true;
    };
    if (! selectFindShortcutResult(file, L"initial shortcut validation"))
    {
        return false;
    }
    RaiseSelfTestWindowForInput(findWindow);
    PumpPendingMessages();
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to focus Find results grid for shortcut test.");
    FindFilesDebugSnapshot focusSnapshot{};
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid && value.selectedResultCount == 1u && value.hasWin32Focus; },
                                      SelfTest::Scale(3000ms),
                                      &focusSnapshot),
                  std::format(L"Find results grid did not focus before shortcut test. {}", DescribeFindSnapshotBrief(focusSnapshot)));
    state.Require(focusSnapshot.rootNavigationVisible, L"Find Look in field should be rendered by an embedded navigation bar.");
    state.Require(focusSnapshot.rootNavigationEmbedded, L"Find Look in navigation bar should use the embedded Find-dialog visual mode.");
    state.Require(! focusSnapshot.rootNavigationEditMode, L"Find Look in navigation bar should start in read-only breadcrumb mode.");
    state.Require(focusSnapshot.rootNavigationText.find(root.native()) != std::wstring::npos,
                  std::format(L"Find Look in navigation text should mirror the configured root '{}'.", root.native()));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto sendKey = [&](WPARAM key) noexcept
    {
        SendMessageW(findWindow, WM_KEYDOWN, key, 0);
        SendMessageW(findWindow, WM_KEYUP, key, 0);
        PumpPendingMessages();
    };

    ClearClipboardContents(findWindow);
    SendFindResultCommand(findWindow, IDM_PANE_CLIPBOARD_COPY);
    std::vector<std::filesystem::path> copyPaths;
    std::optional<DWORD> copyEffect;
    for (size_t retry = 0; retry < 20u &&
                           (! ContainsFindClipboardPath(copyPaths, file) || copyEffect.value_or(DROPEFFECT_NONE) != DROPEFFECT_COPY);
         ++retry)
    {
        copyPaths  = ReadFindClipboardDropPaths(findWindow);
        copyEffect = ReadFindClipboardPreferredDropEffect(findWindow);
        if (! ContainsFindClipboardPath(copyPaths, file) || copyEffect.value_or(DROPEFFECT_NONE) != DROPEFFECT_COPY)
        {
            std::this_thread::sleep_for(20ms);
            PumpPendingMessages();
        }
    }
    FindFilesDebugSnapshot copyClipboardSnapshot{};
    static_cast<void>(DebugGetFindFilesWindowSnapshot(copyClipboardSnapshot));
    state.Require(ContainsFindClipboardPath(copyPaths, file),
                  std::format(L"Find configured clipboard-copy command should put the selected result into CF_HDROP. expected='{}', actual=[{}], effect={}, {}",
                              file.native(),
                              DescribeFindClipboardPaths(copyPaths),
                              DescribeFindClipboardEffect(copyEffect),
                              DescribeFindSnapshotBrief(copyClipboardSnapshot)));
    state.Require(copyEffect.value_or(DROPEFFECT_NONE) == DROPEFFECT_COPY,
                  std::format(L"Find configured clipboard-copy command should publish Preferred DropEffect = COPY. actual={}, paths=[{}], {}",
                              DescribeFindClipboardEffect(copyEffect),
                              DescribeFindClipboardPaths(copyPaths),
                              DescribeFindSnapshotBrief(copyClipboardSnapshot)));

    ClearClipboardContents(findWindow);
    SendFindResultCommand(findWindow, IDM_PANE_CLIPBOARD_CUT);
    std::vector<std::filesystem::path> cutPaths;
    std::optional<DWORD> cutEffect;
    for (size_t retry = 0; retry < 20u &&
                           (! ContainsFindClipboardPath(cutPaths, file) || cutEffect.value_or(DROPEFFECT_NONE) != DROPEFFECT_MOVE);
         ++retry)
    {
        cutPaths  = ReadFindClipboardDropPaths(findWindow);
        cutEffect = ReadFindClipboardPreferredDropEffect(findWindow);
        if (! ContainsFindClipboardPath(cutPaths, file) || cutEffect.value_or(DROPEFFECT_NONE) != DROPEFFECT_MOVE)
        {
            std::this_thread::sleep_for(20ms);
            PumpPendingMessages();
        }
    }
    FindFilesDebugSnapshot cutClipboardSnapshot{};
    static_cast<void>(DebugGetFindFilesWindowSnapshot(cutClipboardSnapshot));
    state.Require(ContainsFindClipboardPath(cutPaths, file),
                  std::format(L"Find configured clipboard-cut command should put the selected result into CF_HDROP. expected='{}', actual=[{}], effect={}, {}",
                              file.native(),
                              DescribeFindClipboardPaths(cutPaths),
                              DescribeFindClipboardEffect(cutEffect),
                              DescribeFindSnapshotBrief(cutClipboardSnapshot)));
    state.Require(cutEffect.value_or(DROPEFFECT_NONE) == DROPEFFECT_MOVE,
                  std::format(L"Find configured clipboard-cut command should publish Preferred DropEffect = MOVE. actual={}, paths=[{}], {}",
                              DescribeFindClipboardEffect(cutEffect),
                              DescribeFindClipboardPaths(cutPaths),
                              DescribeFindSnapshotBrief(cutClipboardSnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForPath = [](const std::filesystem::path& path, std::chrono::milliseconds timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (SelfTest::PathExists(path))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }
        return SelfTest::PathExists(path);
    };
    const auto waitForPathMissing = [](const std::filesystem::path& path, std::chrono::milliseconds timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (! SelfTest::PathExists(path))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }
        return ! SelfTest::PathExists(path);
    };

    FindFilesDebugSnapshot beforePrimaryView{};
    state.Require(DebugGetFindFilesWindowSnapshot(beforePrimaryView), L"Failed to capture Find snapshot before primary viewer launch.");
    SendFindResultCommand(findWindow, IDM_PANE_VIEW);
    state.Require(waitForPath(root / L"find-view-marker.txt", SelfTest::Scale(5000ms)),
                  L"Find configured view command should launch the configured primary viewer for the selected result.");
    FindFilesDebugSnapshot afterPrimaryView{};
    state.Require(DebugGetFindFilesWindowSnapshot(afterPrimaryView), L"Failed to capture Find snapshot after primary viewer launch.");
    state.Require(afterPrimaryView.debugResultActionFocusRestoreRequestCount == beforePrimaryView.debugResultActionFocusRestoreRequestCount,
                  std::format(L"Find viewer command should not schedule a Find foreground/focus restore that can bury the viewer; before={}, after={}. {}",
                              beforePrimaryView.debugResultActionFocusRestoreRequestCount,
                              afterPrimaryView.debugResultActionFocusRestoreRequestCount,
                              DescribeFindSnapshotBrief(afterPrimaryView)));
    std::error_code markerCleanupEc;
    std::filesystem::remove(root / L"find-view-marker.txt", markerCleanupEc);

    FindFilesDebugSnapshot helpSnapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(helpSnapshot), L"Failed to capture Find shortcut help button snapshot.");
    state.Require(helpSnapshot.helpButtonEnabled, L"Find result actions help button should be enabled.");
    state.Require(helpSnapshot.statusStripBlendsWithWindowBackground,
                  L"Find action-row status should blend with the window background instead of drawing a card surface.");
    RECT helpRect{};
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::HelpButton, helpRect),
                  L"Find result actions help button should expose a debug target rectangle.");
    RECT parentRect{};
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::ParentButton, parentRect),
                  L"Find Go to folder button should expose a debug target rectangle for compact help comparison.");
    state.Require((helpRect.right - helpRect.left) > 0 && (helpRect.right - helpRect.left) < (parentRect.right - parentRect.left),
                  L"Find result actions help button should stay narrower than the regular action buttons.");
    RECT findButtonRect{};
    RECT openButtonRect{};
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::FindButton, findButtonRect),
                  L"Find button should expose a debug target rectangle for action-row spacing validation.");
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::OpenButton, openButtonRect),
                  L"Find Open button should expose a debug target rectangle for action-row spacing validation.");
    const LONG openGapPx    = openButtonRect.left - findButtonRect.right;
    const LONG maxOpenGapPx = MulDiv(40, static_cast<int>(GetDpiForWindow(findWindow)), 96);
    state.Require(openGapPx >= 0 && openGapPx <= maxOpenGapPx,
                  std::format(L"Find Open button should sit next to Find Now when Cancel is unavailable; gap was {} px.", openGapPx));
    RECT findClient{};
    state.Require(GetClientRect(findWindow, &findClient) != FALSE && findClient.right > findClient.left,
                  L"Find window should expose a non-empty client rect for help button alignment validation.");
    const LONG helpRightInsetPx = findClient.right - helpRect.right;
    const LONG maxHelpInsetPx   = MulDiv(64, static_cast<int>(GetDpiForWindow(findWindow)), 96);
    state.Require(helpRightInsetPx >= 0 && helpRightInsetPx <= maxHelpInsetPx,
                  std::format(L"Find result actions help button should be right aligned; right inset was {} px.", helpRightInsetPx));
    state.Require(helpSnapshot.statusStripSectionCount == 1u, L"Find status strip should be a single action-row status section.");
    state.Require(helpSnapshot.destinationNavigationVisible, L"Find destination navigation bar should be visible at the bottom of the window.");
    state.Require(helpSnapshot.destinationNavigationText.find(rightRoot.native()) != std::wstring::npos,
                  std::format(L"Find destination navigation should display the other pane path '{}'; saw '{}'.",
                              rightRoot.native(),
                              helpSnapshot.destinationNavigationText));
    state.Require(helpSnapshot.destinationNavigationRect.bottom > helpSnapshot.destinationNavigationRect.top,
                  L"Find destination navigation bar should expose a non-empty debug rectangle.");
    state.Require(helpSnapshot.destinationNavigationEmbedded,
                  L"Find destination navigation bar should use embedded destination styling instead of full pane chrome.");
    state.Require(helpSnapshot.destinationNavigationHistoryCount > 0u,
                  L"Find destination navigation bar should seed pane history so the history arrow has an immediate menu.");
    state.Require(helpSnapshot.destinationNavigationHistoryRect.right > helpSnapshot.destinationNavigationHistoryRect.left &&
                      helpSnapshot.destinationNavigationHistoryRect.bottom > helpSnapshot.destinationNavigationHistoryRect.top,
                  L"Find destination navigation bar should expose a non-empty history-arrow rectangle.");
    state.Require(ProbeFindDestinationNavigationUsesDeliveredHoverPointAfterLiveCursorLeaves(findWindow, state),
                  L"Find destination navigation should use delivered hover points after the live cursor left the embedded control.");
    state.Require(ProbeFindDestinationNavigationUsesDeliveredHistoryClickAfterLiveCursorLeaves(findWindow, state),
                  L"Find destination navigation should use delivered history clicks after the live cursor left the embedded control.");
    state.Require(ProbeFindDestinationHistoryMenuFromActiveEditMode(findWindow, state),
                  L"Find destination navigation should route history-arrow clicks even when the destination edit host is active.");
    state.Require(ProbeFindDestinationHistoryMenu(findWindow, state),
                  L"Find destination navigation history arrow should open and close its menu without waiting for unrelated pointer movement.");

    constexpr std::wstring_view kAlertOverlayWindowClassName = L"RedSalamander.AlertOverlayWindow";
    const auto getFindHelpOverlay             = [&]() noexcept -> HWND { return FindVisibleDescendantWindowByClass(findWindow, kAlertOverlayWindowClassName); };
    const auto describeWindowForFindHelpClick = [](HWND hwnd) -> std::wstring
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            return L"<none>";
        }

        std::array<wchar_t, 128> className{};
        const int classNameLength = GetClassNameW(hwnd, className.data(), static_cast<int>(className.size()));
        RECT rect{};
        GetWindowRect(hwnd, &rect);
        return std::format(L"hwnd=0x{:X}, class='{}', visible={}, enabled={}, rect=({}, {}, {}, {})",
                           reinterpret_cast<UINT_PTR>(hwnd),
                           classNameLength > 0 ? std::wstring_view(className.data(), static_cast<size_t>(classNameLength)) : std::wstring_view(L"<unknown>"),
                           IsWindowVisible(hwnd) != FALSE ? 1 : 0,
                           IsWindowEnabled(hwnd) != FALSE ? 1 : 0,
                           rect.left,
                           rect.top,
                           rect.right,
                           rect.bottom);
    };
    const auto waitForOverlayHidden = [](HWND overlay, std::chrono::milliseconds timeout) noexcept -> bool
    {
        const ULONGLONG deadline = GetTickCount64() + static_cast<ULONGLONG>(timeout.count());
        while (GetTickCount64() < deadline && IsWindow(overlay) != FALSE && IsWindowVisible(overlay) != FALSE)
        {
            PumpPendingMessages();
            Sleep(10);
        }
        return IsWindow(overlay) == FALSE || IsWindowVisible(overlay) == FALSE;
    };
    if (const HWND existingOverlay = getFindHelpOverlay(); existingOverlay && IsWindow(existingOverlay) != FALSE)
    {
        SendMessageW(existingOverlay, WM_KEYDOWN, VK_ESCAPE, 0);
        SendMessageW(existingOverlay, WM_KEYUP, VK_ESCAPE, 0);
        static_cast<void>(waitForOverlayHidden(existingOverlay, SelfTest::Scale(3000ms)));
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::HelpButton), L"Failed to focus Find result actions help button.");
    SendMessageW(findWindow, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_RETURN, 0);
    const HWND immediateHelpOverlay = getFindHelpOverlay();
    state.Require(immediateHelpOverlay != nullptr && IsWindow(immediateHelpOverlay) != FALSE,
                  L"Find result actions help should open synchronously before any pointer movement or paint pump.");
    const HWND helpOverlay = immediateHelpOverlay ? immediateHelpOverlay : WaitForWindow(getFindHelpOverlay, SelfTest::Scale(3000ms));
    state.Require(helpOverlay != nullptr && IsWindow(helpOverlay) != FALSE, L"Find result actions help should open an alert overlay.");
    state.Require(GetFocus() == helpOverlay, L"Find result actions help should take keyboard focus so Escape closes the message.");
    if (helpOverlay && IsWindow(helpOverlay) != FALSE)
    {
        RedSalamander::Ui::AlertOverlayWindowDebugSnapshot immediateOverlay{};
        state.Require(RedSalamander::Ui::DebugGetAlertOverlayWindowSnapshot(helpOverlay, immediateOverlay),
                      L"Find result actions help should expose alert overlay debug state immediately after opening.");
        const RECT immediateCloseRect = immediateOverlay.closeRectPx;
        state.Require(immediateOverlay.visible && immediateOverlay.hasLayout && immediateCloseRect.right > immediateCloseRect.left &&
                          immediateCloseRect.bottom > immediateCloseRect.top,
                      std::format(L"Find result actions help should prepare close hit geometry before WM_PAINT or mouse movement; visible={}, hasLayout={}, "
                                  L"closeRect=({}, {}, {}, {}), paints={}.",
                                  immediateOverlay.visible ? 1 : 0,
                                  immediateOverlay.hasLayout ? 1 : 0,
                                  immediateCloseRect.left,
                                  immediateCloseRect.top,
                                  immediateCloseRect.right,
                                  immediateCloseRect.bottom,
                                  immediateOverlay.paintCount));
        state.Require(immediateOverlay.hasBackdropBitmap,
                      L"Find result actions help should capture the owner backdrop before first paint so the modal scrim does not render over black.");
        PumpPendingMessages();
        RedSalamander::Ui::AlertOverlayWindowDebugSnapshot overlayPaint{};
        state.Require(RedSalamander::Ui::DebugGetAlertOverlayWindowSnapshot(helpOverlay, overlayPaint),
                      L"Find result actions help should expose alert overlay debug paint state.");
        state.Require(overlayPaint.visible && overlayPaint.hasLayout && overlayPaint.paintCount > 0u && overlayPaint.minimumDrawOpacity >= 0.95f,
                      std::format(L"Find result actions help should render a visible first frame without mouse movement; visible={}, hasLayout={}, paints={}, "
                                  L"opacity={:.3f}, minimumOpacity={:.3f}.",
                                  overlayPaint.visible ? 1 : 0,
                                  overlayPaint.hasLayout ? 1 : 0,
                                  overlayPaint.paintCount,
                                  overlayPaint.lastDrawOpacity,
                                  overlayPaint.minimumDrawOpacity));
        state.Require(overlayPaint.lastDrawScrimOpacity >= 0.45f,
                      std::format(L"Find result actions help should apply a real modal scrim opacity; saw {:.3f}.", overlayPaint.lastDrawScrimOpacity));
        state.Require(overlayPaint.usesSharedCloseChrome,
                      L"Find result actions help close affordance should render through shared AlertOverlay/DxUi chrome while preserving the overlay glyph.");
        RECT originalFindWindowRect{};
        state.Require(GetWindowRect(findWindow, &originalFindWindowRect) != FALSE,
                      L"Failed to capture the Find window rectangle before modal backdrop resize validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        const int originalFindWidth  = std::max(1L, originalFindWindowRect.right - originalFindWindowRect.left);
        const int originalFindHeight = std::max(1L, originalFindWindowRect.bottom - originalFindWindowRect.top);
        state.Require(SetWindowPos(findWindow,
                                   nullptr,
                                   originalFindWindowRect.left,
                                   originalFindWindowRect.top,
                                   originalFindWidth + 47,
                                   originalFindHeight + 31,
                                   SWP_NOZORDER | SWP_NOACTIVATE) != FALSE,
                      L"Failed to resize the Find window while the modal help overlay was visible.");
        PumpPendingMessages();
        RedSalamander::Ui::AlertOverlayWindowDebugSnapshot resizedOverlay{};
        state.Require(RedSalamander::Ui::DebugGetAlertOverlayWindowSnapshot(helpOverlay, resizedOverlay),
                      L"Find result actions help should expose alert overlay debug state after owner resize.");
        state.Require(
            resizedOverlay.backdropCaptureCount > overlayPaint.backdropCaptureCount && resizedOverlay.clientSizePx.cx > overlayPaint.clientSizePx.cx &&
                resizedOverlay.clientSizePx.cy > overlayPaint.clientSizePx.cy && resizedOverlay.backdropSizePx.cx == resizedOverlay.clientSizePx.cx &&
                resizedOverlay.backdropSizePx.cy == resizedOverlay.clientSizePx.cy,
            std::format(
                L"Find result actions help should refresh its captured backdrop after owner resize; captures {}->{}, client {}x{}->{}x{}, backdrop {}x{}.",
                overlayPaint.backdropCaptureCount,
                resizedOverlay.backdropCaptureCount,
                overlayPaint.clientSizePx.cx,
                overlayPaint.clientSizePx.cy,
                resizedOverlay.clientSizePx.cx,
                resizedOverlay.clientSizePx.cy,
                resizedOverlay.backdropSizePx.cx,
                resizedOverlay.backdropSizePx.cy));
        static_cast<void>(SetWindowPos(findWindow,
                                       nullptr,
                                       originalFindWindowRect.left,
                                       originalFindWindowRect.top,
                                       originalFindWidth,
                                       originalFindHeight,
                                       SWP_NOZORDER | SWP_NOACTIVATE));
        PumpPendingMessages();
        RedSalamander::Ui::AlertOverlayWindowDebugSnapshot restoredOverlay{};
        state.Require(RedSalamander::Ui::DebugGetAlertOverlayWindowSnapshot(helpOverlay, restoredOverlay),
                      L"Find result actions help should expose alert overlay debug state after restoring owner size.");
        const auto helpTextState = CollectVisibleDescendantNamedElementState(helpOverlay, UIA_TextControlTypeId);
        state.Require(helpTextState.has_value() && helpTextState->name.find(rightRoot.native()) != std::wstring::npos,
                      std::format(L"Find result actions help should display destination folder '{}'; saw '{}'.",
                                  rightRoot.native(),
                                  helpTextState.has_value() ? helpTextState->name : L"<missing>"));
        const RECT closeRect = restoredOverlay.closeRectPx;
        POINT closeCenter{closeRect.left + ((closeRect.right - closeRect.left) / 2), closeRect.top + ((closeRect.bottom - closeRect.top) / 2)};
        SendMessageW(helpOverlay, WM_LBUTTONUP, 0, MAKELPARAM(closeCenter.x, closeCenter.y));
        const bool overlayHidden = waitForOverlayHidden(helpOverlay, SelfTest::Scale(3000ms));
        RedSalamander::Ui::AlertOverlayWindowDebugSnapshot afterReleaseOverlay{};
        static_cast<void>(RedSalamander::Ui::DebugGetAlertOverlayWindowSnapshot(helpOverlay, afterReleaseOverlay));
        state.Require(overlayHidden,
                      std::format(L"Find result actions help close glyph should close on mouse-up even if activation consumed the mouse-down; "
                                  L"client=({}, {}), visible={}, down={}, up={}, dismiss={}, upPart={}.",
                                  closeCenter.x,
                                  closeCenter.y,
                                  afterReleaseOverlay.visible ? 1 : 0,
                                  afterReleaseOverlay.mouseDownCount,
                                  afterReleaseOverlay.mouseUpCount,
                                  afterReleaseOverlay.dismissCount,
                                  afterReleaseOverlay.lastMouseUpHitPart));
    }

    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::HelpButton), L"Failed to refocus Find result actions help button.");
    sendKey(VK_RETURN);
    const HWND clickHelpOverlay = WaitForWindow(getFindHelpOverlay, SelfTest::Scale(3000ms));
    state.Require(clickHelpOverlay != nullptr && IsWindow(clickHelpOverlay) != FALSE, L"Find result actions help should reopen for real-click validation.");
    if (clickHelpOverlay && IsWindow(clickHelpOverlay) != FALSE)
    {
        RedSalamander::Ui::AlertOverlayWindowDebugSnapshot clickOverlay{};
        state.Require(RedSalamander::Ui::DebugGetAlertOverlayWindowSnapshot(clickHelpOverlay, clickOverlay),
                      L"Find result actions help should expose alert overlay debug state before real-click validation.");
        const RECT closeRect = clickOverlay.closeRectPx;
        POINT closeCenter{closeRect.left + ((closeRect.right - closeRect.left) / 2), closeRect.top + ((closeRect.bottom - closeRect.top) / 2)};
        RaiseSelfTestWindowForInput(clickHelpOverlay);
        SendMessageW(clickHelpOverlay, WM_LBUTTONDOWN, 0, MAKELPARAM(closeCenter.x, closeCenter.y));
        SendMessageW(clickHelpOverlay, WM_LBUTTONUP, 0, MAKELPARAM(closeCenter.x, closeCenter.y));
        const bool overlayHidden = waitForOverlayHidden(clickHelpOverlay, SelfTest::Scale(3000ms));
        RedSalamander::Ui::AlertOverlayWindowDebugSnapshot afterClickOverlay{};
        static_cast<void>(RedSalamander::Ui::DebugGetAlertOverlayWindowSnapshot(clickHelpOverlay, afterClickOverlay));
        state.Require(overlayHidden,
                      std::format(L"Clicking the Find result actions help close glyph should close the message without title-bar mouse movement; "
                                  L"click=({}, {}), focus={}, foreground={}, visible={}, down={}, up={}, dismiss={}, downPart={}, upPart={}, "
                                  L"downPt=({}, {}), upPt=({}, {}).",
                                  closeCenter.x,
                                  closeCenter.y,
                                  describeWindowForFindHelpClick(GetFocus()),
                                  describeWindowForFindHelpClick(GetForegroundWindow()),
                                  afterClickOverlay.visible ? 1 : 0,
                                  afterClickOverlay.mouseDownCount,
                                  afterClickOverlay.mouseUpCount,
                                  afterClickOverlay.dismissCount,
                                  afterClickOverlay.lastMouseDownHitPart,
                                  afterClickOverlay.lastMouseUpHitPart,
                                  afterClickOverlay.lastMouseDownPointPx.x,
                                  afterClickOverlay.lastMouseDownPointPx.y,
                                  afterClickOverlay.lastMouseUpPointPx.x,
                                  afterClickOverlay.lastMouseUpPointPx.y));
    }

    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::HelpButton), L"Failed to refocus Find result actions help button.");
    sendKey(VK_RETURN);
    const HWND escapeHelpOverlay = WaitForWindow(getFindHelpOverlay, SelfTest::Scale(3000ms));
    state.Require(escapeHelpOverlay != nullptr && IsWindow(escapeHelpOverlay) != FALSE, L"Find result actions help should reopen for Escape validation.");
    if (escapeHelpOverlay && IsWindow(escapeHelpOverlay) != FALSE)
    {
        const HWND escapeTarget = GetFocus() ? GetFocus() : escapeHelpOverlay;
        SendMessageW(escapeTarget, WM_KEYDOWN, VK_ESCAPE, 0);
        SendMessageW(escapeTarget, WM_KEYUP, VK_ESCAPE, 0);
        state.Require(waitForOverlayHidden(escapeHelpOverlay, SelfTest::Scale(3000ms)), L"Escape should close the Find result actions help message.");
    }

    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.findshortcut", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure recursive Find shortcut search.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to enable recursive Find shortcut options.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start recursive Find shortcut search.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Recursive Find shortcut search did not become idle.");
    state.Require(DebugSelectFindFilesWindowResults({nestedCutFileA.native(), nestedCutFileB.native()}),
                  L"Failed to select multiple Find shortcut results from different subfolders.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to focus Find results grid before multi-subfolder cut.");
    FindFilesDebugSnapshot multiCutSelectionSnapshot{};
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid && value.selectedResultCount == 2u && value.hasWin32Focus; },
                                      SelfTest::Scale(3000ms),
                                      &multiCutSelectionSnapshot),
                  std::format(L"Find results grid did not expose two selected nested rows before multi-subfolder cut. {}",
                              DescribeFindSnapshotBrief(multiCutSelectionSnapshot)));
    const std::filesystem::path clickedMenuExpectedPath(multiCutSelectionSnapshot.selectedResultFullPath);
    state.Require(! clickedMenuExpectedPath.empty() && (OrdinalString::EqualsNoCasePath(clickedMenuExpectedPath, nestedCutFileA) ||
                                                        OrdinalString::EqualsNoCasePath(clickedMenuExpectedPath, nestedCutFileB)),
                  std::format(L"Find selected-row context-menu probe should identify the clicked row path; selected='{}'.",
                              multiCutSelectionSnapshot.selectedResultFullPath));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring clickedSectionText    = LoadStringResource(nullptr, IDS_FIND_RESULT_MENU_THIS_ITEM);
    const std::wstring selectionSectionText  = FormatStringResource(nullptr, IDS_FIND_RESULT_MENU_SELECTION_FMT, 2u);
    const std::wstring clipboardCopyText     = LoadStringResource(nullptr, IDS_CMD_CLIPBOARD_COPY);
    const std::wstring copyToDestinationText = LoadStringResource(nullptr, IDS_FIND_RESULT_MENU_COPY_TO_DESTINATION);
    const std::wstring deleteText            = LoadStringResource(nullptr, IDS_CMD_MOVE_TO_RECYCLE_BIN);
    ShortcutManager resultMenuShortcutManager;
    state.Require(g_settings.shortcuts.has_value(), L"Find result context menu shortcut test should have active shortcut settings.");
    if (g_settings.shortcuts.has_value())
    {
        resultMenuShortcutManager.Load(g_settings.shortcuts.value());
    }
    const auto resolveResultMenuShortcutText = [&](std::wstring_view commandId) -> std::wstring
    {
        const std::optional<ShortcutManager::ShortcutChord> chord = resultMenuShortcutManager.TryGetShortcutForCommand(commandId);
        return chord.has_value() ? FormatFindResultMenuShortcutForTest(chord.value()) : std::wstring{};
    };
    const std::optional<ShortcutManager::ShortcutChord> clipboardCopyChord = resultMenuShortcutManager.TryGetShortcutForCommand(L"cmd/pane/clipboardCopy");
    state.Require(clipboardCopyChord.has_value(), L"Find result context menu should resolve Clipboard Copy through ShortcutManager.");
    const std::wstring clipboardCopyAccel = clipboardCopyChord.has_value() ? FormatFindResultMenuShortcutForTest(clipboardCopyChord.value()) : std::wstring{};
    const std::wstring copyToDestinationAccel = resolveResultMenuShortcutText(L"cmd/pane/copyToOtherPane");
    const std::wstring deleteAccel            = resolveResultMenuShortcutText(L"cmd/pane/moveToRecycleBin");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<bool> resultMenuHasExpectedShape{false};
    std::atomic<bool> resultMenuHasSelectionCopy{false};

    ClearClipboardContents(findWindow);
    state.Require(InvokeFindResultContextMenuItem(findWindow,
                                                  findWindow,
                                                  state,
                                                  [&](const RedSalamander::DxUi::ContextMenuPopupDebugState& popupState) noexcept -> std::optional<size_t>
    {
        const std::optional<size_t> clickedCopy   = FindMenuItemInSection(popupState, clickedSectionText, clipboardCopyText, clipboardCopyAccel);
        const std::optional<size_t> selectionCopy = FindMenuItemInSection(popupState, selectionSectionText, clipboardCopyText, clipboardCopyAccel);
        const bool hasTwoSections                 = std::ranges::find(popupState.itemTexts, clickedSectionText) != popupState.itemTexts.end() &&
                                                    std::ranges::find(popupState.itemTexts, selectionSectionText) != popupState.itemTexts.end();
        resultMenuHasExpectedShape.store(hasTwoSections && clickedCopy.has_value(), std::memory_order_release);
        resultMenuHasSelectionCopy.store(selectionCopy.has_value(), std::memory_order_release);
        return clickedCopy;
    },
                                                  FindResultContextMenuOpenMode::Pointer,
                                                  L"Find result context menu clicked-item Clipboard Copy"),
                  L"Find result context menu clicked-item Clipboard Copy failed.");
    std::vector<std::filesystem::path> clickedMenuCopyPaths;
    for (size_t retry = 0u; retry < 20u && clickedMenuCopyPaths.empty(); ++retry)
    {
        clickedMenuCopyPaths = ReadFindClipboardDropPaths(findWindow);
        if (clickedMenuCopyPaths.empty())
        {
            std::this_thread::sleep_for(20ms);
            PumpPendingMessages();
        }
    }
    state.Require(resultMenuHasExpectedShape.load(std::memory_order_acquire),
                  L"Find result context menu should expose separate clicked-item and selection sections when multiple results are selected.");
    state.Require(resultMenuHasSelectionCopy.load(std::memory_order_acquire),
                  std::format(L"Find result context menu selection section should expose Clipboard Copy with {}.", clipboardCopyAccel));
    state.Require(clickedMenuCopyPaths.size() == 1u && ContainsFindClipboardPath(clickedMenuCopyPaths, clickedMenuExpectedPath),
                  std::format(L"Find result clicked-item context action should operate only on the row that opened the menu; expected='{}', actual=[{}].",
                              clickedMenuExpectedPath.native(),
                              DescribeFindClipboardPaths(clickedMenuCopyPaths)));

    state.Require(DebugSelectFindFilesWindowResults({nestedCutFileA.native(), nestedCutFileB.native()}),
                  L"Failed to restore multiple Find shortcut results after clicked-item menu validation.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to refocus Find results grid before selection-menu validation.");
    ClearClipboardContents(findWindow);
    state.Require(InvokeFindResultContextMenuItem(findWindow,
                                                  findWindow,
                                                  state,
                                                  [&](const RedSalamander::DxUi::ContextMenuPopupDebugState& popupState) noexcept -> std::optional<size_t>
    { return FindMenuItemInSection(popupState, selectionSectionText, clipboardCopyText, clipboardCopyAccel); },
                                                  FindResultContextMenuOpenMode::Keyboard,
                                                  L"Find result context menu selection Clipboard Copy"),
                  L"Find result context menu selection Clipboard Copy failed.");
    std::vector<std::filesystem::path> selectionMenuCopyPaths;
    for (size_t retry = 0u; retry < 20u && selectionMenuCopyPaths.size() != 2u; ++retry)
    {
        selectionMenuCopyPaths = ReadFindClipboardDropPaths(findWindow);
        if (selectionMenuCopyPaths.size() != 2u)
        {
            std::this_thread::sleep_for(20ms);
            PumpPendingMessages();
        }
    }
    state.Require(selectionMenuCopyPaths.size() == 2u && ContainsFindClipboardPath(selectionMenuCopyPaths, nestedCutFileA) &&
                      ContainsFindClipboardPath(selectionMenuCopyPaths, nestedCutFileB),
                  L"Find result selection context action should operate on the whole selected result set.");
    state.Require(DebugSelectFindFilesWindowResults({nestedCutFileA.native(), nestedCutFileB.native()}),
                  L"Failed to restore multiple Find shortcut results after selection-menu validation.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to refocus Find results grid before multi-subfolder cut.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(findWindow);
    SendFindResultCommand(findWindow, IDM_PANE_CLIPBOARD_CUT);
    std::vector<std::filesystem::path> multiCutPaths;
    std::optional<DWORD> multiCutEffect;
    std::wstring multiCutText;
    for (size_t retry = 0u; retry < 20u && (multiCutPaths.size() != 2u || ! multiCutEffect.has_value() || multiCutText.empty()); ++retry)
    {
        multiCutPaths  = ReadFindClipboardDropPaths(findWindow);
        multiCutEffect = ReadFindClipboardPreferredDropEffect(findWindow);
        multiCutText   = ReadClipboardUnicodeText(findWindow);
        if (multiCutPaths.size() != 2u || ! multiCutEffect.has_value() || multiCutText.empty())
        {
            std::this_thread::sleep_for(20ms);
            PumpPendingMessages();
        }
    }
    state.Require(multiCutPaths.size() == 2u, std::format(L"Find cut command should publish two nested CF_HDROP paths; got {}.", multiCutPaths.size()));
    state.Require(ContainsFindClipboardPath(multiCutPaths, nestedCutFileA), L"Find cut command should include nested cut file A in CF_HDROP.");
    state.Require(ContainsFindClipboardPath(multiCutPaths, nestedCutFileB), L"Find cut command should include nested cut file B in CF_HDROP.");
    state.Require(multiCutEffect.value_or(DROPEFFECT_NONE) == DROPEFFECT_MOVE, L"Find cut command should publish Preferred DropEffect = MOVE.");
    state.Require(multiCutText.find(nestedCutFileA.native()) != std::wstring::npos && multiCutText.find(nestedCutFileB.native()) != std::wstring::npos,
                  L"Find cut command should publish full selected paths as Unicode text fallback for nested results.");
    if (! selectFindShortcutResult(file, L"multi-subfolder cut validation"))
    {
        return false;
    }
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to refocus Find results grid after multi-subfolder cut validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendFindResultCommand(findWindow, IDM_PANE_VIEW);
    state.Require(waitForPath(root / L"find-view-marker.txt", SelfTest::Scale(5000ms)),
                  L"Find configured view command should launch the configured primary viewer for the selected result.");

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to refocus Find results grid before alternate view.");
    SendFindResultCommand(findWindow, IDM_PANE_ALTERNATE_VIEW);
    state.Require(waitForPath(root / L"find-alt-view-marker.txt", SelfTest::Scale(5000ms)),
                  L"Find configured alternate-view command should launch the configured alternate viewer for the selected result.");

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to refocus Find results grid before edit.");
    SendFindResultCommand(findWindow, IDM_PANE_EDIT);
    state.Require(waitForPath(root / L"find-edit-marker.txt", SelfTest::Scale(5000ms)),
                  L"Find configured edit command should launch the configured primary editor for the selected result.");

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to refocus Find results grid before alternate edit.");
    SendFindResultCommand(findWindow, IDM_PANE_ALTERNATE_EDIT);
    state.Require(waitForPath(root / L"find-alt-edit-marker.txt", SelfTest::Scale(5000ms)),
                  L"Find configured alternate-edit command should launch the configured alternate editor for the selected result.");

    state.Require(DebugSetFindFilesWindowDestinationPath(explicitRoot.native()), L"Failed to set Find shortcut explicit destination navigation path.");
    FindFilesDebugSnapshot explicitDestinationSnapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.destinationNavigationText.find(explicitRoot.native()) != std::wstring::npos; },
                                      SelfTest::Scale(3000ms),
                                      &explicitDestinationSnapshot),
                  std::format(L"Find destination navigation should accept explicit path '{}'. {}",
                              explicitRoot.native(),
                              DescribeFindSnapshotBrief(explicitDestinationSnapshot)));
    if (! selectFindShortcutResult(explicitCopyFile, L"explicit copy validation"))
    {
        return false;
    }
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to refocus Find results grid before explicit copy.");
    state.Require(InvokeFindResultContextMenuItem(findWindow,
                                                  findWindow,
                                                  state,
                                                  [&](const RedSalamander::DxUi::ContextMenuPopupDebugState& popupState) noexcept -> std::optional<size_t>
    { return FindMenuItemInSection(popupState, std::wstring_view{}, copyToDestinationText, copyToDestinationAccel); },
                                                  FindResultContextMenuOpenMode::Keyboard,
                                                  L"Find result context menu Copy to Destination"),
                  L"Find result context menu Copy to Destination failed.");
    const std::filesystem::path explicitCopiedFile = explicitRoot / explicitCopyFile.filename();
    state.Require(waitForPath(explicitCopiedFile, SelfTest::Scale(5000ms)),
                  L"Find context-menu copy-to-destination action should honor the explicit destination navigation path.");
    state.Require(! SelfTest::PathExists(rightRoot / explicitCopyFile.filename()),
                  L"Find explicit destination copy should not fall back to the opposite pane destination.");
    state.Require(SelfTest::PathExists(explicitCopyFile), L"Find explicit destination copy should keep the source file.");
    FindFilesDebugSnapshot afterExplicitCopy{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.statusText.find(explicitRoot.native()) != std::wstring::npos; },
                                      SelfTest::Scale(3000ms),
                                      &afterExplicitCopy),
                  std::format(L"Find explicit destination copy status should show the destination path '{}'. {}",
                              explicitRoot.native(),
                              DescribeFindSnapshotBrief(afterExplicitCopy)));
    state.Require(DebugSetFindFilesWindowDestinationPath(rightRoot.native()),
                  L"Failed to restore Find shortcut destination navigation path to the opposite pane root.");

    if (! selectFindShortcutResult(copyFile, L"copy-to-other-pane validation"))
    {
        return false;
    }
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to refocus Find results grid before copy.");
    SendFindResultCommand(findWindow, IDM_PANE_COPY_TO_OTHER);
    const std::filesystem::path copiedFile = rightRoot / copyFile.filename();
    state.Require(waitForPath(copiedFile, SelfTest::Scale(5000ms)),
                  L"Find configured copy-to-other-pane command should create the selected result in the other pane.");
    state.Require(SelfTest::PathExists(copyFile), L"Find configured copy-to-other-pane command should keep the source file.");
    FindFilesDebugSnapshot afterCopyToOtherPane{};
    state.Require(
        WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return value.statusText.find(rightRoot.native()) != std::wstring::npos; },
                            SelfTest::Scale(3000ms),
                            &afterCopyToOtherPane),
        std::format(
            L"Find copy-to-other-pane status should show the destination path '{}'. {}", rightRoot.native(), DescribeFindSnapshotBrief(afterCopyToOtherPane)));

    if (! selectFindShortcutResult(moveFile, L"move-to-other-pane validation"))
    {
        return false;
    }
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to refocus Find results grid before move.");
    SendFindResultCommand(findWindow, IDM_PANE_MOVE_TO_OTHER);
    state.Require(waitForPathMissing(moveFile, SelfTest::Scale(5000ms)),
                  L"Find configured move-to-other-pane command should move the selected result through file operations.");
    const std::filesystem::path movedFile = rightRoot / moveFile.filename();
    state.Require(waitForPath(movedFile, SelfTest::Scale(5000ms)),
                  L"Find configured move-to-other-pane command should create the selected result in the other pane.");
    FindFilesDebugSnapshot afterMoveToOtherPane{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return std::find(value.fullPaths.begin(), value.fullPaths.end(), moveFile.native()) == value.fullPaths.end() &&
               value.statusText.find(rightRoot.native()) != std::wstring::npos;
    },
                      SelfTest::Scale(3000ms),
                      &afterMoveToOtherPane),
                  std::format(L"Find result list should remove a moved result and show destination path '{}' after move-to-other-pane is accepted. {}",
                              rightRoot.native(),
                              DescribeFindSnapshotBrief(afterMoveToOtherPane)));

    if (! selectFindShortcutResult(deleteFile, L"move-to-recycle-bin validation"))
    {
        return false;
    }
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid), L"Failed to refocus Find results grid before delete.");
    state.Require(InvokeFindResultContextMenuItem(findWindow,
                                                  findWindow,
                                                  state,
                                                  [&](const RedSalamander::DxUi::ContextMenuPopupDebugState& popupState) noexcept -> std::optional<size_t>
    { return FindMenuItemInSection(popupState, std::wstring_view{}, deleteText, deleteAccel); },
                                                  FindResultContextMenuOpenMode::Pointer,
                                                  L"Find result context menu Move to Recycle Bin"),
                  L"Find result context menu Move to Recycle Bin failed.");
    state.Require(waitForPathMissing(deleteFile, SelfTest::Scale(5000ms)),
                  L"Find context-menu move-to-recycle action should delete the selected result through file operations.");
    FindFilesDebugSnapshot afterRecycleDelete{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return std::find(value.fullPaths.begin(), value.fullPaths.end(), deleteFile.native()) == value.fullPaths.end(); },
                                      SelfTest::Scale(3000ms),
                                      &afterRecycleDelete),
                  std::format(L"Find result list should remove a recycled result after delete is accepted. {}", DescribeFindSnapshotBrief(afterRecycleDelete)));

    if (! selectFindShortcutResult(permanentFile, L"permanent-delete validation"))
    {
        return false;
    }
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to refocus Find results grid before permanent-delete cancel.");
    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_CANCEL);
    SendFindResultCommand(findWindow, IDM_PANE_PERMANENT_DELETE);
    state.Require(HostGetTestPromptRequestCount() == 1u,
                  std::format(L"Find permanent-delete command should request one confirmation prompt on cancel; saw {}.", HostGetTestPromptRequestCount()));
    state.Require(SelfTest::PathExists(permanentFile), L"Canceled Find permanent-delete command should leave the selected result on disk.");
    FindFilesDebugSnapshot afterPermanentCancel{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return std::find(value.fullPaths.begin(), value.fullPaths.end(), permanentFile.native()) != value.fullPaths.end(); },
                                      SelfTest::Scale(3000ms),
                                      &afterPermanentCancel),
                  std::format(L"Canceled Find permanent-delete command should keep the result row. {}", DescribeFindSnapshotBrief(afterPermanentCancel)));

    HostResetTestPromptRequestCount();
    HostSetTestPromptResultOverride(HOST_PROMPT_RESULT_OK);
    SendFindResultCommand(findWindow, IDM_PANE_PERMANENT_DELETE);
    state.Require(HostGetTestPromptRequestCount() == 1u,
                  std::format(L"Find permanent-delete command should request one confirmation prompt on OK; saw {}.", HostGetTestPromptRequestCount()));
    state.Require(waitForPathMissing(permanentFile, SelfTest::Scale(5000ms)),
                  L"Confirmed Find permanent-delete command should remove the selected result through file operations.");
    FindFilesDebugSnapshot afterPermanentDelete{};
    state.Require(
        WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return std::find(value.fullPaths.begin(), value.fullPaths.end(), permanentFile.native()) == value.fullPaths.end(); },
                            SelfTest::Scale(3000ms),
                            &afterPermanentDelete),
        std::format(L"Find result list should remove a permanently deleted result after confirmation. {}", DescribeFindSnapshotBrief(afterPermanentDelete)));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogLargeLocalSearchUsesIncrementalUpdates(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            static_cast<void>(SendMessageW(findWindow, WM_CLOSE, 0, 0));
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(5000ms)));
        }
    };
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_large_incremental";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create incremental-update test subdirectory.");
    constexpr size_t kRootMatchCount = 640u;
    constexpr size_t kSubMatchCount  = 640u;
    for (size_t index = 0; index < kRootMatchCount; ++index)
    {
        const std::filesystem::path filePath = root / std::format(L"root_{:04}.txt", index);
        state.Require(SelfTest::WriteTextFile(filePath, "{\"message\":\"root\"}\n"), std::format(L"Failed to create '{}'.", filePath.filename().native()));
    }
    for (size_t index = 0; index < kSubMatchCount; ++index)
    {
        const std::filesystem::path filePath = sub / std::format(L"sub_{:04}.txt", index);
        state.Require(SelfTest::WriteTextFile(filePath, "{\"message\":\"sub\"}\n"), std::format(L"Failed to create '{}'.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    state.Require(GetConfiguredLocalFileSystemForSelfTest(state, R"({"searchBackendPreference":"scan","searchMaxDirectoryWalkers":4})", created),
                  L"Failed to create scan-only local filesystem for incremental-update test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.fileSystem     = created.fileSystem;
    context.pluginId       = L"builtin/file-system";
    context.pluginShortId  = L"file";
    context.rootPluginPath = root;

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-large-incremental");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for incremental-update test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for incremental-update test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in incremental-update test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure Find options for incremental-update test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for incremental-update test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for incremental-update test.");

    constexpr size_t kExpectedMatches = kRootMatchCount + kSubMatchCount;
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())),
                  L"Find window did not become idle for incremental-update test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture incremental-update snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.lastStatusHint == S_OK,
                  std::format(L"Expected incremental-update search to succeed, got 0x{:08X}.", static_cast<unsigned long>(snapshot.lastStatusHint)));
    state.Require(snapshot.resultCount == kExpectedMatches,
                  std::format(L"Expected {} incremental-update results, got {}.", kExpectedMatches, snapshot.resultCount));
    state.Require(snapshot.incrementalResultRefreshCount != 0u, L"Wildcard Find search did not apply any result-batch refresh before the Find UI settled.");
    state.Require(snapshot.incrementalVisibleResultRefreshCount != 0u,
                  L"Wildcard Find search did not expose a visible result batch before the Find UI settled.");
    state.Require(snapshot.resultListFullRebuildCount == 0u,
                  std::format(L"Append-only Find search should not rebuild the whole results list, got {} rebuild(s).", snapshot.resultListFullRebuildCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRunningStatusShowsPhaseAndPath(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_running_status";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create running-status test subdirectory.");
    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int index = 0; index < 32; ++index)
    {
        const std::filesystem::path filePath =
            (index % 2 == 0) ? (root / std::format(L"root_{:03}.json", index)) : (sub / std::format(L"sub_{:03}.json", index));
        state.Require(SelfTest::WriteTextFile(filePath, largeBody), std::format(L"Failed to create '{}'.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext staleContext{};
    staleContext.pluginId        = L"builtin/file-system-onedrive-business";
    staleContext.pluginShortId   = L"onedrive";
    staleContext.instanceContext = L"@stale";
    staleContext.rootPluginPath  = std::filesystem::path(L"onedrive://@stale");

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-running-status");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(staleContext)), L"Failed to open Find window for running-status test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for running-status test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in running-status test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure Find options for running-status test.");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                L"*.json",
                                                L"ZZZ_NOT_PRESENT_123456789",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for running-status test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for running-status test.");

    const auto hasRunningStatus = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.searchActive && snapshot.statusText.contains(L"Searching with ") && snapshot.statusText.contains(L"last progress ") &&
               snapshot.statusText.find(root.native()) != std::wstring::npos &&
               (snapshot.statusText.contains(L"initializing") || snapshot.statusText.contains(L"enumerating") ||
                snapshot.statusText.contains(L"content scan") || snapshot.statusText.contains(L"index lookup"));
    };
    FindFilesDebugSnapshot active{};
    state.Require(DebugGetFindFilesWindowSnapshot(active), L"Failed to capture immediate running-status snapshot.");
    if (state.failure.empty() && ! hasRunningStatus(active))
    {
        state.Require(WaitForFindSnapshot(hasRunningStatus, SelfTest::Scale(10000ms), &active),
                      std::format(L"Running Find status did not expose phase/path details. Last status='{}'.", active.statusText));
    }
    if (! state.failure.empty())
    {
        static_cast<void>(DebugCancelFindFilesWindowSearch());
        static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())));
        return false;
    }

    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search for running-status test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle after running-status cancellation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogCloseDuringActiveSearchDoesNotJoinUi(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    CloseAllFindFilesWindowsForSearchTest();

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        DebugReleaseFindFilesWindowSearchRunBlocker();
        DebugConfigureFindFilesWindowSearchRunBlocker(false);
        CloseAllFindFilesWindowsForSearchTest();
        g_settings.search = previousSearch;
    });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_active_close_async";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create active-close Find root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha\n"), L"Failed to create active-close Find fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-active-close");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for active-close validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for active-close validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for active-close validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure Find options for active-close validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    DebugConfigureFindFilesWindowSearchRunBlocker(true);
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for active-close validation.");
    state.Require(DebugWaitForFindFilesWindowSearchRunBlocked(static_cast<uint32_t>(SelfTest::Scale(5000ms).count())),
                  L"Find search worker did not reach the active-close blocker.");

    FindFilesDebugSnapshot active{};
    state.Require(DebugGetFindFilesWindowSnapshot(active) && active.searchActive, L"Find search should be active before exercising active-close validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::jthread delayedRelease([&](std::stop_token stopToken) noexcept
    {
        const ULONGLONG deadline = GetTickCount64() + SelfTest::ScaleTimeout(1'200);
        while (! stopToken.stop_requested() && GetTickCount64() < deadline)
        {
            Sleep(10);
        }
        if (! stopToken.stop_requested())
        {
            DebugReleaseFindFilesWindowSearchRunBlocker();
        }
    });

    const ULONGLONG closeStartTick = GetTickCount64();
    state.Require(PostMessageW(findWindow, WM_CLOSE, 0, 0) != FALSE, L"Failed to post active-close WM_CLOSE.");
    PumpPendingMessages();
    const ULONGLONG closeElapsedMs = GetTickCount64() - closeStartTick;

    delayedRelease.request_stop();
    delayedRelease.join();

    state.Require(closeElapsedMs < SelfTest::ScaleTimeout(500), std::format(L"Active Find close blocked the UI thread for {} ms.", closeElapsedMs));
    state.Require(IsWindow(findWindow) != FALSE, L"Active Find close should defer destruction while the worker is still blocked.");
    state.Require(IsWindowVisible(findWindow) == FALSE, L"Active Find close should hide the modeless window while cancellation drains.");
    if (! state.failure.empty())
    {
        DebugReleaseFindFilesWindowSearchRunBlocker();
        DebugConfigureFindFilesWindowSearchRunBlocker(false);
        return false;
    }

    DebugReleaseFindFilesWindowSearchRunBlocker();
    DebugConfigureFindFilesWindowSearchRunBlocker(false);
    state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(5000ms)),
                  L"Deferred active Find close did not destroy the window after the worker completed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogDropsStaleSearchEpochPayloads(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::optional<std::filesystem::path> rightBefore                     = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);
    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    std::optional<std::filesystem::path> leftBefore;
    const auto cleanup = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        g_settings.search = previousSearch;
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
        if (rightBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Right, rightBefore.value());
        }
    });

    CloseAllFindFilesWindowsForSearchTest();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root      = suiteRoot / L"work" / (L"find_stale_epoch_" + NewGuidText());
    const std::filesystem::path alphaPath = root / L"alpha.txt";
    const std::filesystem::path stalePath = root / L"stale-from-old-epoch.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create stale-epoch Find root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha\n"), L"Failed to create stale-epoch alpha fixture.");
    state.Require(SelfTest::WriteTextFile(stalePath, "stale\n"), L"Failed to create stale-epoch rejected fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND findWindow = nullptr;
    state.Require(OpenFindWindowFromLocalPaneRoot(mainWindow, root, {L"alpha.txt", L"stale-from-old-epoch.txt"}, findWindow, leftBefore),
                  L"Find window did not open from local stale-epoch test root.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"alpha.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure stale-epoch Find search.");
    state.Require(DebugSetFindFilesWindowOptions(false, true, false, false, false), L"Failed to configure stale-epoch Find options.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start stale-epoch baseline Find search.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Stale-epoch baseline Find search did not become idle.");

    FindFilesDebugSnapshot before{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return ! value.searchActive && value.resultCount == 1u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), alphaPath.native()) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &before),
                  std::format(L"Stale-epoch baseline result did not settle. {}", DescribeFindSnapshotBrief(before)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugPostFindFilesWindowStaleSearchPayloads(stalePath.native()), L"Failed to post stale Find search payloads.");
    for (int attempt = 0; attempt < 20; ++attempt)
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(1ms);
    }

    FindFilesDebugSnapshot after{};
    state.Require(DebugGetFindFilesWindowSnapshot(after), L"Failed to capture stale-epoch Find snapshot after injected payloads.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(after.resultCount == before.resultCount,
                  std::format(L"Stale epoch result payload mutated visible results. Before {}, after {}.", before.resultCount, after.resultCount));
    state.Require(std::find(after.fullPaths.begin(), after.fullPaths.end(), stalePath.native()) == after.fullPaths.end(),
                  L"Stale epoch result payload inserted a rejected path into the visible result set.");
    state.Require(after.lastStatusHint == before.lastStatusHint,
                  std::format(L"Stale epoch completion changed status from 0x{:08X} to 0x{:08X}.",
                              static_cast<unsigned long>(before.lastStatusHint),
                              static_cast<unsigned long>(after.lastStatusHint)));
    state.Require(after.warningFlags == before.warningFlags,
                  std::format(L"Stale epoch progress/completion changed warning flags from 0x{:08X} to 0x{:08X}.", before.warningFlags, after.warningFlags));
    state.Require(after.phase == before.phase, std::format(L"Stale epoch progress/completion changed phase from {} to {}.", before.phase, after.phase));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogServiceStatusShowsBackendDiagnostics(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    CloseAllFindFilesWindowsForSearchTest();
    DismissVisibleOwnedDxUiContextMenusForSearchTest(mainWindow);
    const auto cleanup = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        DismissVisibleOwnedDxUiContextMenusForSearchTest(mainWindow);
    });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_service_status";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create service-status test subdirectory.");
    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int index = 0; index < 48; ++index)
    {
        const std::filesystem::path filePath =
            (index % 2 == 0) ? (root / std::format(L"root_{:03}.json", index)) : (sub / std::format(L"sub_{:03}.json", index));
        state.Require(SelfTest::WriteTextFile(filePath, largeBody), std::format(L"Failed to create '{}'.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    if (! SearchServiceBackendSelectableForRoot(root, state, L"Find service backend-status test"))
    {
        return true;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring pipeName             = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, pipeName.c_str()) != 0,
                  L"Failed to override the search service pipe for Find backend-status test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    const std::filesystem::path sqlitePath = root / L"find-service-status.sqlite3";
    ForegroundSearchServiceProcess service;
    std::wstring serviceError;
    const std::wstring serviceArgs = std::format(L"--store-backend=sqlite --sqlite-path=\"{}\"", sqlitePath.wstring());
    state.Require(service.Start(pipeName, 32u, serviceArgs, true, serviceError), serviceError);
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForBrokerReady = [&](SearchServiceBroker::ServiceStatus& outStatus) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(10000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outStatus = {};
            if (SUCCEEDED(SearchServiceBroker::GetStatus(outStatus)) && outStatus.pipeName == pipeName)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outStatus = {};
        return SUCCEEDED(SearchServiceBroker::GetStatus(outStatus)) && outStatus.pipeName == pipeName;
    };

    SearchServiceBroker::ServiceStatus initialStatus{};
    state.Require(waitForBrokerReady(initialStatus), std::format(L"Search service broker did not become ready for the isolated pipe '{}'.", pipeName));
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    if (! info)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"service\"}");
    state.Require(SUCCEEDED(setHr),
                  std::format(L"Failed to configure service backend for Find backend-status test. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.fileSystem     = created.fileSystem;
    context.pluginId       = std::wstring(kBuiltinLocalFileSystemId);
    context.pluginShortId  = L"file";
    context.rootPluginPath = root;

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-service-status");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for service backend-status test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for service backend-status test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in service backend-status test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for service backend-status test.");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                L"*.json",
                                                L"ZZZ_NOT_PRESENT_123456789",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for service backend-status test.");

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for service backend-status test.");

    const auto leftInitialPlaceholder = [](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.backend != FILESYSTEM_SEARCH_BACKEND_UNKNOWN || snapshot.phase != FILESYSTEM_SEARCH_PHASE_INITIALIZING ||
               snapshot.warningFlags != FILESYSTEM_SEARCH_WARNING_NONE || snapshot.statusText != L"Ready" || snapshot.hasServiceStatus ||
               ! snapshot.backendStatusText.empty();
    };

    FindFilesDebugSnapshot firstProgress{};
    state.Require(WaitForFindSnapshot(leftInitialPlaceholder, SelfTest::Scale(5000ms), &firstProgress),
                  std::format(L"Service-backed Find search did not transition out of the initial placeholder status. status='{}' backend='{}' active={} "
                              L"phase={} warnings=0x{:08X}.",
                              firstProgress.statusText,
                              firstProgress.backendStatusText,
                              firstProgress.searchActive ? 1 : 0,
                              static_cast<unsigned>(firstProgress.phase),
                              static_cast<unsigned long>(firstProgress.warningFlags)));
    if (! state.failure.empty())
    {
        static_cast<void>(DebugCancelFindFilesWindowSearch());
        static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())));
        return false;
    }
    const auto containsAny = [](std::wstring_view text, std::initializer_list<std::wstring_view> tokens) noexcept
    {
        for (const std::wstring_view token : tokens)
        {
            if (! token.empty() && text.find(token) != std::wstring_view::npos)
            {
                return true;
            }
        }
        return false;
    };

    const auto waitForServiceDiagnostics = [&](FindFilesDebugSnapshot& outActive) noexcept
    {
        return WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& snapshot) noexcept
        {
            return snapshot.backend == FILESYSTEM_SEARCH_BACKEND_SERVICE && snapshot.hasServiceStatus &&
                   (snapshot.warningFlags & FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX) != 0u && snapshot.backendStatusText.contains(L"db ") &&
                   snapshot.backendStatusText.contains(L" | search ") && snapshot.backendStatusText.contains(L"live-scan") &&
                   containsAny(snapshot.backendStatusText, {L"store-missing", L"store-invalid", L"warmup-running", L"store-stale"});
        },
            SelfTest::Scale(15000ms),
            &outActive);
    };

    FindFilesDebugSnapshot active{};
    bool reachedDiagnostics = waitForServiceDiagnostics(active);
    if (! reachedDiagnostics)
    {
        state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel the initial Find search after missing service backend diagnostics.");
        state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())),
                      L"Find window did not become idle after cancelling the initial service backend-status attempt.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to restart Find search for the service backend-status retry.");
        if (! state.failure.empty())
        {
            return false;
        }

        firstProgress = {};
        state.Require(WaitForFindSnapshot(leftInitialPlaceholder, SelfTest::Scale(5000ms), &firstProgress),
                      std::format(L"Service-backed Find retry did not transition out of the initial placeholder status. status='{}' backend='{}' active={} "
                                  L"phase={} warnings=0x{:08X}.",
                                  firstProgress.statusText,
                                  firstProgress.backendStatusText,
                                  firstProgress.searchActive ? 1 : 0,
                                  static_cast<unsigned>(firstProgress.phase),
                                  static_cast<unsigned long>(firstProgress.warningFlags)));
        if (! state.failure.empty())
        {
            static_cast<void>(DebugCancelFindFilesWindowSearch());
            static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())));
            return false;
        }

        reachedDiagnostics = waitForServiceDiagnostics(active);
    }

    state.Require(reachedDiagnostics,
                  std::format(L"Service-backed Find status did not expose database diagnostics. Last status='{}' backend='{}'.",
                              active.statusText,
                              active.backendStatusText));
    if (! state.failure.empty())
    {
        static_cast<void>(DebugCancelFindFilesWindowSearch());
        static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())));
        return false;
    }

    state.Require((active.warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) == 0u,
                  std::format(L"Healthy service degraded path should not report service unavailable. warnings=0x{:08X}.",
                              static_cast<unsigned long>(active.warningFlags)));
    state.Require(active.statusText.find(active.backendStatusText) != std::wstring::npos,
                  std::format(L"Running Find status should include the backend diagnostics text. status='{}' backend='{}'.",
                              active.statusText,
                              active.backendStatusText));

    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search for service backend-status test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())),
                  L"Find window did not become idle after service backend-status cancellation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogServiceUnavailableWarningIsDistinct(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    CloseAllFindFilesWindowsForSearchTest();
    DismissVisibleOwnedDxUiContextMenusForSearchTest(mainWindow);
    const auto cleanup = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        DismissVisibleOwnedDxUiContextMenusForSearchTest(mainWindow);
    });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_service_unavailable";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create service-unavailable test root.");
    state.Require(SelfTest::WriteTextFile(root / L"match.jsonl", "{\"message\":\"seed\"}"), L"Failed to create match.jsonl for service-unavailable test.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! SearchServiceBackendSelectableForRoot(root, state, L"Find unavailable-service warning test"))
    {
        return true;
    }

    const std::wstring previousPipeOverride = GetEnvVarTrimmed(SearchServiceBroker::kPipeNameEnvVar);
    const std::wstring unavailablePipe      = MakeUniquePipeName();
    state.Require(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, unavailablePipe.c_str()) != 0,
                  L"Failed to override the search service pipe for the unavailable-service Find test.");
    const auto restorePipeOverride = wil::scope_exit([&] noexcept
    {
        const wchar_t* restoreValue = previousPipeOverride.empty() ? nullptr : previousPipeOverride.c_str();
        static_cast<void>(::SetEnvironmentVariableW(SearchServiceBroker::kPipeNameEnvVar, restoreValue));
    });

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    if (! info)
    {
        return false;
    }

    const HRESULT setHr = info->SetConfiguration("{\"searchBackendPreference\":\"service\"}");
    state.Require(SUCCEEDED(setHr),
                  std::format(L"Failed to configure service backend for unavailable-service Find test. hr=0x{:08X}", static_cast<unsigned long>(setHr)));
    if (FAILED(setHr))
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.fileSystem     = created.fileSystem;
    context.pluginId       = std::wstring(kBuiltinLocalFileSystemId);
    context.pluginShortId  = L"file";
    context.rootPluginPath = root;

    const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-service-unavailable");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for unavailable-service warning test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for unavailable-service warning test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in unavailable-service warning test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for unavailable-service warning test.");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for unavailable-service warning test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for unavailable-service warning test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(15000ms).count())),
                  L"Find window did not become idle for unavailable-service warning test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture unavailable-service warning snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring serviceUnavailableText = LoadStringResource(nullptr, IDS_FIND_WARNING_SERVICE_UNAVAILABLE);
    state.Require(
        snapshot.lastStatusHint == S_OK,
        std::format(L"Unavailable-service Find search should still succeed via fallback. hr=0x{:08X}.", static_cast<unsigned long>(snapshot.lastStatusHint)));
    state.Require(snapshot.resultCount == 1u, std::format(L"Unavailable-service Find search should still return one match, got {}.", snapshot.resultCount));
    state.Require((snapshot.warningFlags & FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE) != 0u,
                  std::format(L"Unavailable-service Find search should report the service-unavailable warning. warnings=0x{:08X}.",
                              static_cast<unsigned long>(snapshot.warningFlags)));
    state.Require(snapshot.statusText.contains(serviceUnavailableText),
                  std::format(L"Unavailable-service Find status should surface '{}' in the UI, got '{}'.", serviceUnavailableText, snapshot.statusText));
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogFailureShowsReadableStatus(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    CloseAllFindFilesWindowsForSearchTest();
    DismissVisibleOwnedDxUiContextMenusForSearchTest(mainWindow);
    const auto cleanup = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        DismissVisibleOwnedDxUiContextMenusForSearchTest(mainWindow);
    });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path paneRoot    = suiteRoot / L"work" / L"find_dialog_failure_status";
    const std::filesystem::path missingRoot = paneRoot / L"missing";

    std::error_code ec;
    std::filesystem::remove_all(paneRoot, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(paneRoot), L"Failed to create failure-status pane root.");
    state.Require(SelfTest::WriteTextFile(paneRoot / L"seed.jsonl", "{\"message\":\"seed\"}"), L"Failed to create seed.jsonl.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for failure-status test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, paneRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, paneRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for failure-status test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"seed.jsonl"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for failure-status test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in failure-status test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for failure-status test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(findWindow, mainWindow), L"Find window should be an independent top-level window in failure-status test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure Find options for failure-status test.");
    state.Require(DebugConfigureFindFilesWindow(
                      missingRoot.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for failure-status test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for failure-status test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for failure-status test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture failure-status snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FAILED(snapshot.lastStatusHint), L"Expected failure-status test search to fail.");
    const std::wstring expectedHr = std::format(L"0x{:08X}", static_cast<unsigned long>(snapshot.lastStatusHint));
    state.Require(snapshot.statusText.starts_with(L"Search failed with "),
                  std::format(L"Expected readable failure status with backend context, got '{}'.", snapshot.statusText));
    state.Require(snapshot.statusText.find(L": ") != std::wstring::npos,
                  std::format(L"Readable failure status should include a backend/value separator, got '{}'.", snapshot.statusText));
    state.Require(snapshot.statusText.contains(expectedHr),
                  std::format(L"Readable failure status should preserve the HRESULT code '{}', got '{}'.", expectedHr, snapshot.statusText));
    state.Require(snapshot.statusText != std::format(L"Search failed ({})", expectedHr), L"Failure status regressed to the old raw-HRESULT format.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogUsesDxUiHostWithNoVisibleChildControls(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const auto waitForFindWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms)); };

    const auto validateFindHostSurface = [&](std::wstring_view context) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/find during {}.", context));

        const HWND findWindow = waitForFindWindow();
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, std::format(L"Find window did not open during {}.", context));
        if (! findWindow || IsWindow(findWindow) == FALSE)
        {
            return false;
        }

        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.usesDxUiHost, std::format(L"Find window is not attached to the shared DxUi host during {}.", context));
        state.Require(snapshot.hasStatusStrip, std::format(L"Find window should expose a retained DxUi StatusStrip during {}.", context));
        state.Require(snapshot.statusStripVisible, std::format(L"Find window StatusStrip should remain visible during {}.", context));
        state.Require(
            snapshot.statusStripHeightDip >= 20.0f,
            std::format(L"Find window StatusStrip should keep a usable height during {}; observed {:.2f} DIP.", context, snapshot.statusStripHeightDip));
        state.Require(WindowExposesUiaProvider(findWindow), std::format(L"Find window should answer WM_GETOBJECT during {}.", context));
        state.Require(snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Find window should not expose visible child-control fallback during {}; got {} visible child window(s).",
                                  context,
                                  snapshot.visibleChildWindowCount));
        state.Require(snapshot.dxResizeFailureCount == 0u,
                      std::format(L"Find window should not report DX resize failures during {}; saw {}.", context, snapshot.dxResizeFailureCount));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(findWindow);
        state.Require(uiaPatternStats.has_value(), std::format(L"Failed to collect live UI Automation pattern statistics for Find window during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Find window should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->editControlCount + uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Find window should expose visible UI Automation edit or combo descendants during {}.", context));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation ValuePattern support during {}.", context));
            state.Require(uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Find window should expose a visible UI Automation command button during {}.", context));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation InvokePattern support during {}.", context));
        }

        auto valueState = CollectVisibleDescendantValuePatternState(findWindow, UIA_EditControlTypeId);
        if (! valueState.has_value())
        {
            valueState = CollectVisibleDescendantValuePatternState(findWindow, UIA_ComboBoxControlTypeId);
        }
        state.Require(valueState.has_value(), std::format(L"Find window should expose a visible editable DX field during {}.", context));
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly, std::format(L"Find window editable DX field should not be read-only during {}.", context));
            state.Require(! valueState->name.empty(), std::format(L"Find window editable DX field should expose a stable accessible name during {}.", context));
        }

        const auto buttonState = CollectVisibleDescendantNamedElementState(findWindow, UIA_ButtonControlTypeId);
        state.Require(buttonState.has_value(), std::format(L"Find window should expose a visible DX command button during {}.", context));
        if (buttonState.has_value())
        {
            state.Require(! buttonState->name.empty(),
                          std::format(L"Find window visible DX command button should expose a stable accessible name during {}.", context));
        }

        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)), std::format(L"Find window did not close cleanly during {}.", context));
        return state.failure.empty();
    };

    if (! validateFindHostSurface(L"the initial Find DX host baseline probe"))
    {
        return false;
    }

    if (! validateFindHostSurface(L"the reopened Find DX host baseline probe"))
    {
        return false;
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogExposesLiveUiaSelectionAndInputs(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_live_uia";
    const std::filesystem::path subA = root / L"sub_a";
    const std::filesystem::path subB = root / L"sub_b";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(subA), L"Failed to create first Find UIA test subdirectory.");
    state.Require(SelfTest::EnsureDirectory(subB), L"Failed to create second Find UIA test subdirectory.");
    state.Require(SelfTest::WriteTextFile(subA / L"inside.txt", "payload-a"), L"Failed to create first Find UIA test file.");
    state.Require(SelfTest::WriteTextFile(subB / L"inside.txt", "payload-b"), L"Failed to create second Find UIA test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find UIA test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find UIA test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub_a", L"sub_b"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Find UIA test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring subAPath = subA.native();
    const std::wstring subBPath = subB.native();

    const auto runLiveUiaCycle = [&](std::wstring_view phaseLabel) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/find during {}.", phaseLabel));

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, std::format(L"Find window did not open during {}.", phaseLabel));
        if (! findWindow || IsWindow(findWindow) == FALSE)
        {
            return false;
        }

        state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), std::format(L"Failed to configure Find options during {}.", phaseLabel));
        state.Require(DebugConfigureFindFilesWindow(
                          root.native(), L"sub*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                      std::format(L"Failed to configure Find window during {}.", phaseLabel));
        state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), std::format(L"Failed to start Find search during {}.", phaseLabel));
        state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                      std::format(L"Find window did not become idle during {}.", phaseLabel));

        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot during {}.", phaseLabel));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.usesDxUiHost, std::format(L"Find window should stay attached to the shared DxUi host during {}.", phaseLabel));
        state.Require(
            snapshot.visibleChildWindowCount <= 1u,
            std::format(L"Find window should not expose visible child-control fallback during {}; saw {}.", phaseLabel, snapshot.visibleChildWindowCount));
        state.Require(snapshot.resultCount == 2u,
                      std::format(L"Find UIA test should expose exactly two visible results during {}; saw {}.", phaseLabel, snapshot.resultCount));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(findWindow);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for Find window during {}.", phaseLabel));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Find window should expose visible UI Automation descendants during {}.", phaseLabel));
            state.Require(uiaPatternStats->editControlCount + uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Find window should expose visible UI Automation edit or combo descendants during {}.", phaseLabel));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation ValuePattern support during {}.", phaseLabel));
            state.Require(uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation TogglePattern support during {}.", phaseLabel));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        const auto requireFindSelection = [&](const std::wstring& fullPath, const std::wstring_view expectedLeaf, const std::wstring_view label) noexcept
        {
            state.Require(DebugSelectFindFilesWindowResult(fullPath),
                          std::format(L"Failed to select '{}' in Find results during {} ({})", fullPath, phaseLabel, label));
            if (! state.failure.empty())
            {
                return false;
            }

            FindFilesDebugSnapshot selectedSnapshot{};
            state.Require(DebugGetFindFilesWindowSnapshot(selectedSnapshot),
                          std::format(L"Failed to capture Find snapshot after {} during {}.", label, phaseLabel));
            if (! state.failure.empty())
            {
                return false;
            }

            state.Require(
                selectedSnapshot.selectedResultCount == 1u,
                std::format(
                    L"Find should expose exactly one selected result after {} during {}; saw {}.", label, phaseLabel, selectedSnapshot.selectedResultCount));

            const auto selectionState = CollectVisibleDescendantSelectionPatternState(findWindow, UIA_DataGridControlTypeId);
            state.Require(
                selectionState.has_value(),
                std::format(L"Failed to collect live UI Automation selection state for the Find results grid after {} during {}.", label, phaseLabel));
            if (! selectionState.has_value())
            {
                return false;
            }

            state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId, L"Find results host should expose a UI Automation DataGrid control.");
            state.Require(selectionState->hasSelectionPattern, L"Find results grid should expose SelectionPattern.");
            state.Require(selectionState->selectionCount == 1u,
                          std::format(L"Find results grid should expose exactly one selected UIA row after {} during {}; saw {}.",
                                      label,
                                      phaseLabel,
                                      selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId, L"Find selected UIA row should expose the DataItem control type.");
            state.Require(selectionState->selectedHasSelectionItemPattern, L"Find selected UIA row should expose SelectionItemPattern.");
            state.Require(! selectionState->selectedName.empty(), L"Find selected UIA row should expose a non-empty accessible name.");
            state.Require(selectionState->selectedName.find(expectedLeaf) != std::wstring::npos,
                          std::format(L"Find selected UIA row name '{}' should include the selected directory name '{}' after {} during {}.",
                                      selectionState->selectedName,
                                      expectedLeaf,
                                      label,
                                      phaseLabel));
            return state.failure.empty();
        };

        state.Require(requireFindSelection(subAPath, subA.filename().native(), L"selecting the first Find result"),
                      std::format(L"Find should keep the UIA results-grid selection synchronized for the first result during {}.", phaseLabel));
        state.Require(requireFindSelection(subBPath, subB.filename().native(), L"switching to the second Find result"),
                      std::format(L"Find should keep the UIA results-grid selection synchronized for the second result during {}.", phaseLabel));
        state.Require(requireFindSelection(subAPath, subA.filename().native(), L"switching back to the first Find result"),
                      std::format(L"Find should keep the UIA results-grid selection synchronized when switching back during {}.", phaseLabel));
        return state.failure.empty();
    };

    state.Require(runLiveUiaCycle(L"initial live UIA validation"), L"Find window initial live UIA validation cycle failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(GetFindFilesWindowHandle() == nullptr || IsWindow(GetFindFilesWindowHandle()) == FALSE,
                  L"Find window should fully close before the reopened live UIA validation cycle.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runLiveUiaCycle(L"reopened live UIA validation"), L"Find window reopened live UIA validation cycle failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogLongRunScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Find scrolling validation.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_long_run_scrolling";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find long-run scrolling test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kResultCount = 48u;
    for (size_t index = 0; index < kResultCount; ++index)
    {
        const std::filesystem::path subdir = root / std::format(L"dir_{:02}", index);
        state.Require(SelfTest::EnsureDirectory(subdir), std::format(L"Failed to create Find long-run scrolling subdirectory '{}'.", subdir.native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const auto cleanupFindWindow = wil::scope_exit([&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow) != FALSE)
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find long-run scrolling test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find long-run scrolling test.");
    FocusFolderViewPane(FolderWindow::Pane::Left);

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in Find long-run scrolling test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for long-run scrolling validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(false, false, true, false, false), L"Failed to configure Find options for long-run scrolling validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"dir_*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for long-run scrolling validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for long-run scrolling validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return ! value.searchActive && value.resultCount >= kResultCount; },
                                      SelfTest::Scale(15000ms),
                                      &snapshot),
                  L"Find window did not finish loading enough results for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the Find results grid for long-run scrolling validation.");
    state.Require(snapshot.usesDxUiHost, L"Find window should stay attached to the shared DxUi host during long-run scrolling validation.");
    state.Require(snapshot.visibleChildWindowCount <= 1u,
                  std::format(L"Find window should not expose visible child-control fallback during long-run scrolling validation; saw {}.",
                              snapshot.visibleChildWindowCount));
    state.Require(snapshot.visibleResultRowCount > 0u, L"Find results grid should expose visible rows before long-run scrolling validation.");
    state.Require(snapshot.visibleResultColumnCount > 0u, L"Find results grid should expose visible columns before long-run scrolling validation.");
    state.Require(snapshot.visibleResultRowCount < snapshot.resultCount,
                  std::format(L"Find results grid should stay virtualized during long-run scrolling validation; visible rows={} total results={}.",
                              snapshot.visibleResultRowCount,
                              snapshot.resultCount));
    state.Require(snapshot.resultListHasVerticalScrollbar, L"Find results grid should expose a vertical scrollbar during long-run scrolling validation.");
    state.Require(snapshot.dxResizeFailureCount == 0u,
                  std::format(L"Find results grid should start with zero DX resize failures; saw {}.", snapshot.dxResizeFailureCount));

    const size_t initialVisibleRows        = snapshot.visibleResultRowCount;
    const size_t initialVisibleColumns     = snapshot.visibleResultColumnCount;
    const uint32_t initialFullRebuildCount = snapshot.resultListFullRebuildCount;
    const uint64_t initialResizeCount      = snapshot.dxResizeCount;
    uint64_t previousRenderCount           = snapshot.dxRenderCount;

    size_t acceptedScrollChunks = 0u;
    for (size_t chunk = 0; chunk < 12u; ++chunk)
    {
        const bool scrolled = DebugScrollFindFilesWindowResultsByWheelDetents(-1);
        if (! scrolled)
        {
            break;
        }
        ++acceptedScrollChunks;

        state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return value.dxRenderCount > previousRenderCount; },
                                          SelfTest::Scale(3000ms),
                                          &snapshot),
                      std::format(L"Find results grid did not repaint after long-run scroll chunk {}.", chunk));
        if (! state.failure.empty())
        {
            return false;
        }

        previousRenderCount = snapshot.dxRenderCount;
        state.Require(snapshot.resultCount >= kResultCount,
                      std::format(L"Find results grid lost results during long-run scroll chunk {}; saw {}.", chunk, snapshot.resultCount));
        state.Require(snapshot.visibleResultRowCount > 0u && snapshot.visibleResultRowCount <= initialVisibleRows + 1u,
                      std::format(L"Find results grid visible row work became unbounded during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleResultRowCount,
                                  initialVisibleRows));
        state.Require(snapshot.visibleResultColumnCount == initialVisibleColumns,
                      std::format(L"Find results grid visible column work changed unexpectedly during chunk {}; saw {} vs baseline {}.",
                                  chunk,
                                  snapshot.visibleResultColumnCount,
                                  initialVisibleColumns));
        state.Require(snapshot.visibleResultCellCount <= snapshot.visibleResultRowCount * snapshot.visibleResultColumnCount,
                      std::format(L"Find results grid visible cell work became inconsistent during chunk {}; saw {} cells for {} rows and {} columns.",
                                  chunk,
                                  snapshot.visibleResultCellCount,
                                  snapshot.visibleResultRowCount,
                                  snapshot.visibleResultColumnCount));
        state.Require(snapshot.resultListHasVerticalScrollbar,
                      std::format(L"Find results grid lost its vertical scrollbar during long-run scroll chunk {}.", chunk));
        state.Require(snapshot.resultListFullRebuildCount == initialFullRebuildCount,
                      std::format(L"Find results grid unexpectedly rebuilt its full list during chunk {}; rebuild count moved from {} to {}.",
                                  chunk,
                                  initialFullRebuildCount,
                                  snapshot.resultListFullRebuildCount));
        state.Require(snapshot.dxResizeCount == initialResizeCount,
                      std::format(L"Find results grid churned DX host resizes during chunk {}; resize count moved from {} to {}.",
                                  chunk,
                                  initialResizeCount,
                                  snapshot.dxResizeCount));
        state.Require(snapshot.dxResizeFailureCount == 0u,
                      std::format(L"Find results grid hit DX resize failures during chunk {}; saw {}.", chunk, snapshot.dxResizeFailureCount));
    }

    state.Require(acceptedScrollChunks > 0u, L"Find results grid should accept at least one bounded wheel scroll before reaching the lower edge.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow) != FALSE)
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_long_run_open_close";
    const std::filesystem::path subA = root / L"sub_a";
    const std::filesystem::path subB = root / L"sub_b";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(subA), L"Failed to create first Find churn subdirectory.");
    state.Require(SelfTest::EnsureDirectory(subB), L"Failed to create second Find churn subdirectory.");
    state.Require(SelfTest::WriteTextFile(subA / L"inside.txt", "payload-a"), L"Failed to create first Find churn file.");
    state.Require(SelfTest::WriteTextFile(subB / L"inside.txt", "payload-b"), L"Failed to create second Find churn file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find churn test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find churn test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub_a", L"sub_b"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Find churn test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring subAPath = subA.native();
    const std::wstring subBPath = subB.native();

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        closeFindWindow();
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/find during cycle {}.", cycle));

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, std::format(L"Find window did not open during cycle {}.", cycle));
        if (! findWindow || IsWindow(findWindow) == FALSE)
        {
            return false;
        }

        state.Require(! IsOwnedBy(findWindow, mainWindow), std::format(L"Find window should be an independent top-level window during cycle {}.", cycle));
        state.Require(WindowExposesUiaProvider(findWindow), std::format(L"Find window should answer WM_GETOBJECT during cycle {}.", cycle));
        state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), std::format(L"Failed to configure Find options during cycle {}.", cycle));
        state.Require(DebugConfigureFindFilesWindow(
                          root.native(), L"sub*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                      std::format(L"Failed to configure Find window during cycle {}.", cycle));
        state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), std::format(L"Failed to start Find search during cycle {}.", cycle));
        state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                      std::format(L"Find window did not become idle during cycle {}.", cycle));

        const std::wstring& selectedPath     = ((cycle % 2u) == 0u) ? subAPath : subBPath;
        const std::wstring_view expectedLeaf = ((cycle % 2u) == 0u) ? L"sub_a" : L"sub_b";
        state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' during cycle {}.", selectedPath, cycle));

        FindFilesDebugSnapshot snapshot{};
        state.Require(DebugGetFindFilesWindowSnapshot(snapshot), std::format(L"Failed to capture Find snapshot during cycle {}.", cycle));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.usesDxUiHost, std::format(L"Find window should stay attached to the shared DxUi host during cycle {}.", cycle));
        state.Require(
            snapshot.visibleChildWindowCount <= 1u,
            std::format(L"Find window should not expose visible child-control fallback during cycle {}; saw {}.", cycle, snapshot.visibleChildWindowCount));
        state.Require(snapshot.resultCount == 2u,
                      std::format(L"Find window should keep exactly two visible results during cycle {}; saw {}.", cycle, snapshot.resultCount));
        state.Require(snapshot.selectedResultCount == 1u,
                      std::format(L"Find window should keep exactly one selected result during cycle {}; saw {}.", cycle, snapshot.selectedResultCount));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(findWindow);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for Find window during cycle {}.", cycle));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Find window should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(uiaPatternStats->editControlCount + uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Find window should expose visible editable form descendants during cycle {}.", cycle));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation ValuePattern support during cycle {}.", cycle));
            state.Require(uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Find window should expose live UI Automation TogglePattern support during cycle {}.", cycle));
        }

        const auto selectionState = CollectVisibleDescendantSelectionPatternState(findWindow, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(),
                      std::format(L"Failed to collect live UI Automation selection state for the Find results grid during cycle {}.", cycle));
        if (selectionState.has_value())
        {
            state.Require(selectionState->hasSelectionPattern, std::format(L"Find results grid should expose SelectionPattern during cycle {}.", cycle));
            state.Require(
                selectionState->selectionCount == 1u,
                std::format(L"Find results grid should keep exactly one selected row during cycle {}; saw {}.", cycle, selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                          std::format(L"Find selected UIA row should remain a DataItem during cycle {}.", cycle));
            state.Require(selectionState->selectedHasSelectionItemPattern,
                          std::format(L"Find selected UIA row should keep SelectionItemPattern during cycle {}.", cycle));
            state.Require(
                selectionState->selectedName.find(expectedLeaf) != std::wstring::npos,
                std::format(L"Find selected UIA row name should track '{}' during cycle {}; saw '{}'.", expectedLeaf, cycle, selectionState->selectedName));
        }

        auto valueState = CollectVisibleDescendantValuePatternState(findWindow, UIA_EditControlTypeId);
        if (! valueState.has_value())
        {
            valueState = CollectVisibleDescendantValuePatternState(findWindow, UIA_ComboBoxControlTypeId);
        }
        state.Require(valueState.has_value(), std::format(L"Failed to collect UI Automation ValuePattern state for the Find form during cycle {}.", cycle));
        if (valueState.has_value())
        {
            state.Require(! valueState->isReadOnly, std::format(L"Find visible DX edit surface should remain editable during cycle {}.", cycle));
            state.Require(! valueState->name.empty(),
                          std::format(L"Find visible DX edit surface should expose a stable accessible name during cycle {}.", cycle));
        }

        const auto buttonState = CollectVisibleDescendantNamedElementState(findWindow, UIA_ButtonControlTypeId);
        state.Require(buttonState.has_value(), std::format(L"Failed to collect a visible DX command button state for the Find window during cycle {}.", cycle));
        if (buttonState.has_value())
        {
            state.Require(! buttonState->name.empty(),
                          std::format(L"Find visible DX command button should expose a stable accessible name during cycle {}.", cycle));
        }

        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)), std::format(L"Find window did not close cleanly during cycle {}.", cycle));
        state.Require(GetFindFilesWindowHandle() == nullptr || IsWindow(GetFindFilesWindowHandle()) == FALSE,
                      std::format(L"Find window should not linger after close during cycle {}.", cycle));
    }

    state.Require(GetFindFilesWindowHandle() == nullptr || IsWindow(GetFindFilesWindowHandle()) == FALSE,
                  L"Find window should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogDirectoryActivationNavigatesIntoSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_directory_activation";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create activation test subdirectory.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create activation test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find activation test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find activation test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Find activation test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in directory activation test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for directory activation test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure Find options for directory activation test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for directory activation test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for directory activation test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for directory activation test.");

    const std::wstring subPath = sub.native();
    state.Require(DebugSelectFindFilesWindowResult(subPath), std::format(L"Failed to select '{}' in Find results.", subPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot for directory activation test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.selectedResultCount == 1u, std::format(L"Expected one selected Find result, got {}.", snapshot.selectedResultCount));
    state.Require(snapshot.openButtonEnabled, L"Open button should be enabled after selecting a Find result.");
    state.Require(snapshot.parentButtonEnabled, L"Parent button should be enabled after selecting a Find result.");

    state.Require(DebugActivateSelectedFindFilesWindowResult(), L"Failed to activate selected Find result.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sub, SelfTest::Scale(5000ms)),
                  L"Activating directory Find result did not navigate into the selected directory.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogOpenParentKeepsDirectoryFocusedInParent(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_open_parent";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create open-parent test subdirectory.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create open-parent test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find open-parent test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Find open-parent test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Find open-parent test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in open-parent test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for open-parent test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure Find options for open-parent test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for open-parent test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for open-parent test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for open-parent test.");

    const std::wstring subPath = sub.native();
    state.Require(DebugSelectFindFilesWindowResult(subPath), std::format(L"Failed to select '{}' in Find results for open-parent test.", subPath));
    state.Require(DebugOpenSelectedFindFilesWindowResultParent(), L"Failed to invoke Go to folder on the selected Find result.");

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                  L"Go to folder did not navigate to the selected result parent folder.");

    const auto focusDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < focusDeadline)
    {
        MSG msg{};
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        if (g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"sub")
        {
            return true;
        }

        std::this_thread::sleep_for(10ms);
    }

    state.Require(false,
                  std::format(L"Go to folder should focus 'sub' in '{}', got focus on '{}'.",
                              root.native(),
                              std::wstring(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left))));
    return false;
}

[[nodiscard]] bool TestFindDialogEnterFromCheckboxInvokesDefaultSearch(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_default_enter";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create default-enter test directory.");
    state.Require(SelfTest::WriteTextFile(root / L"match.txt", "payload"), L"Failed to create default-enter test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in default-enter test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for default-enter test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"match.txt", L"", Common::Settings::SearchNameMode::Literal, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for default-enter test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for default-enter test.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::RecursiveCheck), L"Failed to focus recursive checkbox before default-enter test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture pre-search snapshot for default-enter test.");
    state.Require(snapshot.focusTarget == FindFilesDebugFocusTarget::RecursiveCheck,
                  L"Recursive checkbox did not keep focus before Enter default-button routing.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_RETURN, 0);

    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for default-enter test.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.resultCount == 1u; }, SelfTest::Scale(3000ms), &snapshot),
                  L"Pressing Enter from a non-editor Find control did not trigger the default Find action.");
    state.Require(! snapshot.statusText.empty(), L"Default Enter should publish a completed Find status after the search runs.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogPointerClickTogglesRecursiveCheckbox(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
            static_cast<void>(WaitForFindWindowUnavailable(SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                  L"Shortcut dispatch failed for cmd/pane/find in recursive-checkbox pointer test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for recursive-checkbox pointer test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(false, true, false, true, false), L"Failed to configure Find options for recursive-checkbox pointer test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && ! value.searchActive && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u &&
               ! value.recursiveChecked;
    },
                      SelfTest::Scale(10000ms),
                      &snapshot),
                  L"Find window did not reach the expected baseline state before recursive-checkbox pointer test.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT recursiveRect{};
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::RecursiveCheck, recursiveRect),
                  L"Failed to capture recursive-checkbox client rect for pointer test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int clickX                     = recursiveRect.left + ((recursiveRect.right - recursiveRect.left) / 2);
    const int clickY                     = recursiveRect.top + ((recursiveRect.bottom - recursiveRect.top) / 2);
    const LPARAM clickPoint              = MAKELPARAM(clickX, clickY);
    const uint64_t baselineRenderCount   = snapshot.dxRenderCount;
    const uint64_t baselineResizeCount   = snapshot.dxResizeCount;
    const size_t baselineVisibleChildren = snapshot.visibleChildWindowCount;

    const auto clickRecursive = [&]() noexcept
    {
        SendMessageW(findWindow, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(findWindow, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    const auto requireRecursiveState = [&](const bool expectedChecked, std::wstring_view label) noexcept
    {
        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && ! value.searchActive && value.visibleChildWindowCount == baselineVisibleChildren && value.dxResizeFailureCount == 0u &&
                   value.dxResizeCount == baselineResizeCount && value.focusTarget == FindFilesDebugFocusTarget::RecursiveCheck &&
                   value.recursiveChecked == expectedChecked && value.dxRenderCount >= baselineRenderCount;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      std::format(L"Recursive checkbox did not reach the expected state after {}.", label));
    };

    clickRecursive();
    requireRecursiveState(true, L"first pointer click");
    if (! state.failure.empty())
    {
        return false;
    }

    clickRecursive();
    requireRecursiveState(false, L"second pointer click");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogEscapeClosesPopupBeforeCancel(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in popup-escape test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for popup-escape test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameModeCombo), L"Failed to focus name-mode combo before popup-escape test.");

    SendMessageW(findWindow, WM_SYSKEYDOWN, VK_DOWN, 0);
    SendMessageW(findWindow, WM_SYSKEYUP, VK_DOWN, 0);

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.nameModePopupOpen && value.focusTarget == FindFilesDebugFocusTarget::NameModeCombo; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Alt+Down did not open the focused Find mode combo popup.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_ESCAPE, 0);

    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return ! value.nameModePopupOpen && value.focusTarget == FindFilesDebugFocusTarget::NameModeCombo && ! value.searchActive; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Esc should close the focused Find combo popup before any cancel routing.");
    state.Require(IsWindow(findWindow) != FALSE, L"Esc should not close the Find window while dismissing a combo popup.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogEscapeFromDxControlClosesCancel(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
            static_cast<void>(WaitForFindWindowUnavailable(SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in cancel-escape test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for cancel-escape test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::RecursiveCheck), L"Failed to focus recursive checkbox before cancel-escape test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::RecursiveCheck && ! value.searchActive && value.usesDxUiHost &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(10000ms),
                      &snapshot),
                  L"Find window did not reach the expected focused DX idle state before cancel-escape test.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_ESCAPE, 0);

    const bool closedAfterEscape = WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms));
    if (! closedAfterEscape && IsWindow(findWindow) != FALSE)
    {
        PostMessageW(findWindow, WM_CLOSE, 0, 0);
        static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
    }

    state.Require(closedAfterEscape, L"Esc from a focused DX Find control should close the dialog through cancel routing.");
    state.Require(GetFindFilesWindowHandle() == nullptr || IsWindow(GetFindFilesWindowHandle()) == FALSE,
                  L"Find window should not remain open after Esc cancel routing.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogAccessKeysFocusExpectedFields(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in access-key test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for access-key test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    const auto expectFocusAfterMnemonic = [&](wchar_t mnemonic, FindFilesDebugFocusTarget expected, std::wstring_view label) noexcept
    {
        FindFilesDebugSnapshot snapshot{};
        SendMessageW(findWindow, WM_SYSCHAR, mnemonic, 0);
        state.Require(WaitForFindSnapshot([expected](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == expected; },
                                          SelfTest::Scale(2000ms),
                                          &snapshot),
                      std::format(L"Mnemonic '{}' did not focus the expected Find field '{}'.", mnemonic, label));
    };

    expectFocusAfterMnemonic(L'l', FindFilesDebugFocusTarget::RootCombo, L"Look in");
    expectFocusAfterMnemonic(L'n', FindFilesDebugFocusTarget::NameCombo, L"Named");
    expectFocusAfterMnemonic(L'c', FindFilesDebugFocusTarget::ContentCombo, L"Containing");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogGridFocusedEnterActivatesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_grid_enter";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create grid-enter test subdirectory.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create grid-enter test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for grid-enter test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for grid-enter test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for grid-enter test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in grid-enter test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for grid-enter test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure Find options for grid-enter test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for grid-enter test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for grid-enter test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for grid-enter test.");

    const std::wstring subPath = sub.native();
    state.Require(DebugSelectFindFilesWindowResult(subPath), std::format(L"Failed to select '{}' in Find results for grid-enter test.", subPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot for grid-enter test.");
    state.Require(snapshot.focusTarget == FindFilesDebugFocusTarget::ResultsGrid,
                  L"Selecting a Find result for grid-enter test should focus the results grid.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_RETURN, 0);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sub, SelfTest::Scale(5000ms)),
                  L"Pressing Enter on a focused Find result did not activate the selected row.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogGridDoubleClickActivatesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_grid_double_click";
    const std::filesystem::path sub  = root / L"sub";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create find_grid_double_click test folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for grid double-click test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for grid double-click test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"sub"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for grid double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in grid double-click test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for grid double-click test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure Find options for grid double-click test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for grid double-click test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for grid double-click test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle for grid double-click test.");

    const std::wstring subPath = sub.native();
    state.Require(DebugSelectFindFilesWindowResult(subPath), std::format(L"Failed to select '{}' in Find results for grid double-click test.", subPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot for grid double-click test.");
    state.Require(snapshot.focusTarget == FindFilesDebugFocusTarget::ResultsGrid,
                  L"Selecting a Find result for grid double-click test should focus the results grid.");
    state.Require(snapshot.selectedResultRowRect.right > snapshot.selectedResultRowRect.left &&
                      snapshot.selectedResultRowRect.bottom > snapshot.selectedResultRowRect.top,
                  L"Find results grid did not expose selected-row geometry for grid double-click test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LPARAM doubleClickPoint = DipPointToClientLParam(findWindow,
                                                           (snapshot.selectedResultRowRect.left + snapshot.selectedResultRowRect.right) * 0.5f,
                                                           (snapshot.selectedResultRowRect.top + snapshot.selectedResultRowRect.bottom) * 0.5f);
    SendMouseDoubleClickToResolvedPointWindow(findWindow, doubleClickPoint);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sub, SelfTest::Scale(5000ms)),
                  L"Double-clicking a focused Find result did not activate the selected row.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogTabTraversalMatchesExpectedOrder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path findRoot = suiteRoot / L"work" / L"find_tab_traversal";
    const std::filesystem::path subDir   = findRoot / L"sub";
    const std::filesystem::path hitFile  = subDir / L"needle.txt";

    std::error_code ec;
    std::filesystem::remove_all(findRoot, ec);
    state.Require(SelfTest::EnsureDirectory(subDir), L"Failed to create find_tab_traversal folder.");
    state.Require(SelfTest::WriteTextFile(hitFile, "needle\n"), L"Failed to seed find_tab_traversal result file.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in tab-traversal test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for tab-traversal test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(
                      findRoot.native(), L"*.txt", L"needle", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for tab-traversal test.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, true), L"Failed to configure Find options for tab-traversal test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for tab-traversal test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for tab-traversal test.");
    state.Require(DebugSelectFindFilesWindowResult(hitFile.native()),
                  std::format(L"Failed to select '{}' before Find tab-traversal validation.", hitFile.native()));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return ! value.searchActive && value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u &&
               value.resultCount == 1u && value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled &&
               value.appendButtonEnabled && value.intersectButtonEnabled && value.subtractButtonEnabled && value.visibleResultRowCount > 0u &&
               value.visibleResultColumnCount > 0u && value.visibleResultCellCount > 0u;
    },
                      SelfTest::Scale(5000ms),
                      &snapshot),
                  L"Find window did not settle to the expected DX results state before tab-traversal validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(FocusFindRootNavigationForKeyboard(findWindow, SelfTest::Scale(3000ms), &snapshot),
                  std::format(L"Root combo did not take active Win32 focus before Find tab-traversal validation. {}", DescribeFindSnapshotBrief(snapshot)));
    state.Require(! snapshot.searchActive && snapshot.resultCount == 1u && snapshot.selectedResultCount == 1u,
                  std::format(L"Find results state changed while focusing the root combo before tab traversal. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;

    const auto sendTab = [&](bool reverse, FindFilesDebugFocusTarget expected, std::wstring_view label) noexcept
    {
        const FindFilesDebugFocusTarget beforeFocus = snapshot.focusTarget;
        std::array<BYTE, 256> keyboardState{};
        const bool restoreKeyboardState = reverse && GetKeyboardState(keyboardState.data()) != FALSE;
        const auto keyboardStateRestore = wil::scope_exit([&]() noexcept
        {
            if (restoreKeyboardState)
            {
                SetKeyboardState(keyboardState.data());
            }
        });
        if (reverse && restoreKeyboardState)
        {
            std::array<BYTE, 256> shiftedState = keyboardState;
            shiftedState[VK_SHIFT] |= 0x80u;
            shiftedState[VK_LSHIFT] |= 0x80u;
            SetKeyboardState(shiftedState.data());
        }

        const auto sendKeyMessages = [&]() noexcept
        {
            if (reverse)
            {
                SendMessageW(findWindow, WM_KEYDOWN, VK_SHIFT, 0);
            }
            SendMessageW(findWindow, WM_KEYDOWN, VK_TAB, 0);
            SendMessageW(findWindow, WM_KEYUP, VK_TAB, 0);
            if (reverse)
            {
                SendMessageW(findWindow, WM_KEYUP, VK_SHIFT, 0);
            }
        };

        const auto reachedExpected = [&]() noexcept
        {
            return WaitForFindSnapshot(
                [&](const FindFilesDebugSnapshot& value) noexcept
            {
                return value.focusTarget == expected && ! value.searchActive && value.usesDxUiHost && value.visibleChildWindowCount <= 1u &&
                       value.dxResizeFailureCount == 0u && value.resultCount == 1u && value.selectedResultCount == 1u &&
                       value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
                       value.visibleResultCellCount == baselineVisibleCellCount;
            },
                SelfTest::Scale(2000ms),
                &snapshot);
        };

        sendKeyMessages();
        if (! reachedExpected() && snapshot.focusTarget == beforeFocus)
        {
            sendKeyMessages();
            static_cast<void>(reachedExpected());
        }

        state.Require(snapshot.focusTarget == expected,
                      std::format(L"{} focus target not reached during Find tab traversal. {}", label, DescribeFindSnapshotBrief(snapshot)));
    };

    sendTab(false, FindFilesDebugFocusTarget::NameCombo, L"Name combo");
    sendTab(false, FindFilesDebugFocusTarget::NameModeCombo, L"Name mode combo");
    sendTab(false, FindFilesDebugFocusTarget::ContentCombo, L"Content combo");
    sendTab(false, FindFilesDebugFocusTarget::ContentModeCombo, L"Content mode combo");
    sendTab(false, FindFilesDebugFocusTarget::RecursiveCheck, L"Recursive checkbox");
    sendTab(false, FindFilesDebugFocusTarget::IncludeFilesCheck, L"Include files checkbox");
    sendTab(false, FindFilesDebugFocusTarget::IncludeDirectoriesCheck, L"Include directories checkbox");
    sendTab(false, FindFilesDebugFocusTarget::FollowSymlinksCheck, L"Follow symlinks checkbox");
    sendTab(false, FindFilesDebugFocusTarget::MatchCaseNameCheck, L"Match-case name checkbox");
    sendTab(false, FindFilesDebugFocusTarget::MatchCaseContentCheck, L"Match-case content checkbox");
    sendTab(false, FindFilesDebugFocusTarget::PreferIndexCheck, L"Prefer index checkbox");
    sendTab(false, FindFilesDebugFocusTarget::WantSnippetsCheck, L"Include snippets checkbox");
    sendTab(false, FindFilesDebugFocusTarget::FindButton, L"Find button");
    sendTab(false, FindFilesDebugFocusTarget::OpenButton, L"Open button");
    sendTab(false, FindFilesDebugFocusTarget::ParentButton, L"Go to folder button");
    sendTab(false, FindFilesDebugFocusTarget::HelpButton, L"result actions help button");
    sendTab(false, FindFilesDebugFocusTarget::ResultsGrid, L"results grid");
    sendTab(false, FindFilesDebugFocusTarget::RootCombo, L"wrapped root combo");

    sendTab(true, FindFilesDebugFocusTarget::ResultsGrid, L"reverse wrapped results grid");
    sendTab(true, FindFilesDebugFocusTarget::HelpButton, L"reverse result actions help button");
    sendTab(true, FindFilesDebugFocusTarget::ParentButton, L"reverse Go to folder button");
    sendTab(true, FindFilesDebugFocusTarget::OpenButton, L"reverse Open button");
    sendTab(true, FindFilesDebugFocusTarget::FindButton, L"reverse Find button");
    sendTab(true, FindFilesDebugFocusTarget::WantSnippetsCheck, L"reverse Include snippets checkbox");
    sendTab(true, FindFilesDebugFocusTarget::PreferIndexCheck, L"reverse Prefer index checkbox");
    sendTab(true, FindFilesDebugFocusTarget::MatchCaseContentCheck, L"reverse Match-case content checkbox");
    sendTab(true, FindFilesDebugFocusTarget::MatchCaseNameCheck, L"reverse Match-case name checkbox");
    sendTab(true, FindFilesDebugFocusTarget::FollowSymlinksCheck, L"reverse Follow symlinks checkbox");
    sendTab(true, FindFilesDebugFocusTarget::IncludeDirectoriesCheck, L"reverse Include directories checkbox");
    sendTab(true, FindFilesDebugFocusTarget::IncludeFilesCheck, L"reverse Include files checkbox");
    sendTab(true, FindFilesDebugFocusTarget::RecursiveCheck, L"reverse Recursive checkbox");
    sendTab(true, FindFilesDebugFocusTarget::ContentModeCombo, L"reverse Content mode combo");
    sendTab(true, FindFilesDebugFocusTarget::ContentCombo, L"reverse Content combo");
    sendTab(true, FindFilesDebugFocusTarget::NameModeCombo, L"reverse Name mode combo");
    sendTab(true, FindFilesDebugFocusTarget::NameCombo, L"reverse Name combo");
    sendTab(true, FindFilesDebugFocusTarget::RootCombo, L"reverse Root combo");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogEditableComboKeyboardEditingKeys(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window is invalid for Find editable-combo keyboard editing test.");
        return false;
    }

    CloseAllFindFilesWindowsForSearchTest();
    const auto cleanupFindWindows = wil::scope_exit([]() noexcept { CloseAllFindFilesWindowsForSearchTest(); });

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for Find editable-combo keyboard editing test.");
    const std::filesystem::path editRoot = suiteRoot / L"work" / L"find_editable_combo_keyboard" / L"alpha" / L"beta" / L"gamma";
    std::error_code cleanupError;
    std::filesystem::remove_all(editRoot.parent_path().parent_path().parent_path(), cleanupError);
    state.Require(SelfTest::EnsureDirectory(editRoot), std::format(L"Failed to create Find editable-combo keyboard editing root '{}'.", editRoot.native()));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring editRootText            = editRoot.native();
    const std::wstring editRootParentText      = editRoot.parent_path().native() + L"\\";
    const std::wstring editRootGrandparentText = editRoot.parent_path().parent_path().native() + L"\\";
    const std::filesystem::path paneRoot       = editRoot.parent_path().parent_path().parent_path();

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Find editable-combo keyboard editing test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, paneRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, paneRoot, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for Find editable-combo keyboard editing test.");
    if (! state.failure.empty())
    {
        return false;
    }

    RaiseSelfTestWindowForInput(mainWindow);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                  L"Shortcut dispatch failed for cmd/pane/find in editable-combo keyboard editing test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for editable-combo keyboard editing test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesDebugSnapshot readySnapshot{};
    const bool rootNavigationReady = WaitForFindSnapshot(
        [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.rootNavigationVisible && value.rootNavigationEmbedded && value.visibleChildWindowCount <= 1u &&
               value.dxResizeFailureCount == 0u;
    },
        SelfTest::Scale(3000ms),
        &readySnapshot);
    state.Require(
        rootNavigationReady,
        std::format(L"Find root navigation did not become ready before editable-combo keyboard editing test. {}", DescribeFindSnapshotBrief(readySnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const auto waitForRootText = [&](std::wstring_view expected, std::wstring_view label) noexcept
    {
        FindFilesDebugSnapshot snapshot{};
        const bool textReady = WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.rootNavigationEditMode && value.rootNavigationText == expected && value.visibleChildWindowCount <= 1u &&
                   value.dxResizeFailureCount == 0u;
        },
            SelfTest::Scale(2000ms),
            &snapshot);
        state.Require(textReady,
                      std::format(L"Find root navigation edit text should be '{}' after {} but was '{}' (focusTarget={}, editMode={}). {}",
                                  expected,
                                  label,
                                  snapshot.rootNavigationText,
                                  static_cast<int>(snapshot.focusTarget),
                                  snapshot.rootNavigationEditMode ? 1 : 0,
                                  DescribeFindSnapshotBrief(snapshot)));
    };

    const auto prepareRootText = [&](std::wstring text) noexcept -> HWND
    {
        FindFilesDebugSnapshot beforePrepareSnapshot{};
        if (DebugGetFindFilesWindowSnapshot(beforePrepareSnapshot) && beforePrepareSnapshot.rootNavigationEditMode)
        {
            CloseAllFindFilesWindowsForSearchTest();
            RaiseSelfTestWindowForInput(mainWindow);
            FocusFolderViewPane(FolderWindow::Pane::Left);
            state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"),
                          L"Shortcut dispatch failed while reopening Find window for editable-combo keyboard editing test.");
            const HWND reopenedFindWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
            state.Require(reopenedFindWindow != nullptr && IsWindow(reopenedFindWindow) != FALSE,
                          L"Find window did not reopen while preparing editable-combo keyboard editing test.");

            FindFilesDebugSnapshot reopenedSnapshot{};
            state.Require(WaitForFindSnapshot(
                              [](const FindFilesDebugSnapshot& value) noexcept
            {
                return value.usesDxUiHost && value.rootNavigationVisible && value.rootNavigationEmbedded && ! value.rootNavigationEditMode &&
                       value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
            },
                              SelfTest::Scale(3000ms),
                              &reopenedSnapshot),
                          std::format(L"Reopened Find root navigation did not become ready before preparing next keyboard-editing value. {}",
                                      DescribeFindSnapshotBrief(reopenedSnapshot)));
            if (! state.failure.empty())
            {
                return nullptr;
            }
        }

        const size_t caretIndex         = text.size();
        const std::wstring expectedText = text;
        state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::RootCombo, std::move(text)),
                      L"Failed to set root combo text for editable-combo keyboard editing test.");

        const HWND activeFindWindow = GetFindFilesWindowHandle();
        state.Require(activeFindWindow != nullptr && IsWindow(activeFindWindow) != FALSE,
                      L"Find window handle is invalid while preparing editable-combo keyboard editing test.");

        FindFilesDebugSnapshot populatedSnapshot{};
        const bool rootTextPrepared = WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.rootNavigationVisible && value.rootNavigationEmbedded && ! value.rootNavigationEditMode &&
                   value.rootNavigationText == expectedText && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
        },
            SelfTest::Scale(2000ms),
            &populatedSnapshot);
        state.Require(rootTextPrepared,
                      std::format(L"Find root navigation did not show prepared text '{}' before edit mode; saw '{}'. {}",
                                  expectedText,
                                  populatedSnapshot.rootNavigationText,
                                  DescribeFindSnapshotBrief(populatedSnapshot)));

        FindFilesDebugSnapshot focusedSnapshot{};
        const bool rootFocused = FocusFindRootNavigationForKeyboard(activeFindWindow, SelfTest::Scale(3000ms), &focusedSnapshot);
        state.Require(
            rootFocused,
            std::format(L"Find root navigation did not retain active Win32 focus before entering edit mode. {}", DescribeFindSnapshotBrief(focusedSnapshot)));
        PumpPendingMessages();

        HWND target = focusedSnapshot.rootNavigationHwnd;
        if (! target || IsWindow(target) == FALSE)
        {
            FindFilesDebugSnapshot refreshedFocusSnapshot{};
            if (DebugGetFindFilesWindowSnapshot(refreshedFocusSnapshot))
            {
                target = refreshedFocusSnapshot.rootNavigationHwnd;
            }
        }
        if (! target || IsWindow(target) == FALSE)
        {
            target = GetFocus();
        }
        state.Require(target != nullptr && IsWindow(target) != FALSE, L"Focused root navigation target is invalid before entering edit mode.");
        if (target && IsWindow(target) != FALSE)
        {
            SetFocus(target);
            SendMessageW(target, WM_KEYDOWN, VK_RETURN, 0);
            SendMessageW(target, WM_KEYUP, VK_RETURN, 0);
        }
        PumpPendingMessages();

        FindFilesDebugSnapshot editSnapshot{};
        const bool editHostOpened = WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            const bool editHostAvailable = (value.rootNavigationEditInputHwnd && IsWindow(value.rootNavigationEditInputHwnd) != FALSE) ||
                                           (value.rootNavigationEditHostHwnd && IsWindow(value.rootNavigationEditHostHwnd) != FALSE);
            return value.usesDxUiHost && value.rootNavigationVisible && value.rootNavigationEmbedded && value.rootNavigationEditMode &&
                   value.rootNavigationText == expectedText && editHostAvailable && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
        },
            SelfTest::Scale(2000ms),
            &editSnapshot);
        state.Require(editHostOpened,
                      std::format(L"Find root navigation edit host did not open with text '{}' but was '{}'. {}",
                                  expectedText,
                                  editSnapshot.rootNavigationText,
                                  DescribeFindSnapshotBrief(editSnapshot)));

        if (! state.failure.empty())
        {
            return nullptr;
        }

        target = editSnapshot.rootNavigationEditInputHwnd;
        if (! target || IsWindow(target) == FALSE)
        {
            target = editSnapshot.rootNavigationEditHostHwnd;
        }
        if (! target || IsWindow(target) == FALSE)
        {
            target = GetFocus();
        }
        if (target && IsWindow(target) != FALSE)
        {
            SetFocus(target);
        }
        state.Require(target != nullptr && IsWindow(target) != FALSE, L"Focused root navigation edit target is invalid for keyboard editing test.");
        if (target && IsWindow(target) != FALSE)
        {
            SendMessageW(target, EM_SETSEL, static_cast<WPARAM>(caretIndex), static_cast<LPARAM>(caretIndex));
            DWORD selectionStart = 0u;
            DWORD selectionEnd   = 0u;
            SendMessageW(target, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));
            state.Require(selectionStart == caretIndex && selectionEnd == caretIndex,
                          std::format(L"Root navigation edit selection should collapse to {} but was {}..{}.", caretIndex, selectionStart, selectionEnd));
        }
        return target;
    };

    const auto sendKey = [](HWND keyboardTarget, WPARAM key) noexcept
    {
        if (keyboardTarget && IsWindow(keyboardTarget) != FALSE)
        {
            SendMessageW(keyboardTarget, WM_KEYDOWN, key, 0);
            SendMessageW(keyboardTarget, WM_KEYUP, key, 0);
        }
        PumpPendingMessages();
    };

    const auto sendChar = [](HWND keyboardTarget, wchar_t ch) noexcept
    {
        if (keyboardTarget && IsWindow(keyboardTarget) != FALSE)
        {
            SendMessageW(keyboardTarget, WM_CHAR, static_cast<WPARAM>(ch), 0);
        }
        PumpPendingMessages();
    };

    const auto sendCtrlKey = [](HWND keyboardTarget, WPARAM key, WPARAM translatedChar) noexcept
    {
        if (keyboardTarget && IsWindow(keyboardTarget) != FALSE)
        {
            SendMessageW(keyboardTarget, WM_KEYDOWN, VK_CONTROL, 0);
            SendMessageW(keyboardTarget, WM_KEYDOWN, key, 0);
            if (translatedChar != 0u)
            {
                SendMessageW(keyboardTarget, WM_CHAR, translatedChar, 0);
            }
            SendMessageW(keyboardTarget, WM_KEYUP, key, 0);
            SendMessageW(keyboardTarget, WM_KEYUP, VK_CONTROL, 0);
        }
        PumpPendingMessages();
    };

    HWND target = prepareRootText(editRootText);
    sendCtrlKey(target, static_cast<WPARAM>(L'A'), 0x01u);
    sendChar(target, L'Z');
    waitForRootText(L"Z", L"Ctrl+A followed by character input");
    if (! state.failure.empty())
    {
        return false;
    }

    target = prepareRootText(editRootText);
    waitForRootText(editRootText, L"preparing Ctrl+Backspace");
    sendCtrlKey(target, VK_BACK, 0x7Fu);
    waitForRootText(editRootParentText, L"Ctrl+Backspace");
    sendCtrlKey(target, VK_BACK, 0x7Fu);
    waitForRootText(editRootGrandparentText, L"repeated Ctrl+Backspace");
    if (! state.failure.empty())
    {
        return false;
    }

    target = prepareRootText(L"abcd");
    SendMessageW(target, EM_SETSEL, 1, 1);
    sendKey(target, VK_DELETE);
    waitForRootText(L"acd", L"Delete");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogModeTypeaheadUpdatesSelectionAndDependencies(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    CloseAllFindFilesWindowsForSearchTest();
    DismissVisibleOwnedDxUiContextMenusForSearchTest(mainWindow);
    const auto cleanup = wil::scope_exit([&] noexcept
    {
        CloseAllFindFilesWindowsForSearchTest();
        DismissVisibleOwnedDxUiContextMenusForSearchTest(mainWindow);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in mode-typeahead test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for mode-typeahead test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(L"C:\\", L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Find window for mode-typeahead test.");

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameModeCombo), L"Failed to focus name-mode combo before typeahead test.");
    SendMessageW(findWindow, WM_CHAR, L'r', 0);

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::NameModeCombo && value.nameModeSelectedIndex.has_value() &&
               value.nameModeSelectedIndex.value() == 2u;
    },
                      SelfTest::Scale(2000ms),
                      &snapshot),
                  L"Typing on the focused name-mode combo did not select Regex via DX combo typeahead.");

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ContentModeCombo),
                  L"Failed to focus content-mode combo before typeahead dependency test.");
    SendMessageW(findWindow, WM_CHAR, L'r', 0);
    const bool contentModeReady = WaitForFindSnapshot(
        [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::ContentModeCombo && value.contentModeSelectedIndex.has_value() &&
               value.contentModeSelectedIndex.value() == 2u && value.matchCaseContentEnabled && value.wantSnippetsEnabled;
    },
        SelfTest::Scale(2000ms),
        &snapshot);
    state.Require(contentModeReady,
                  std::format(L"Typing on the focused content-mode combo did not select Regex and enable content-specific options. "
                              L"(focus={}, selected={}, matchCaseEnabled={}, wantSnippetsEnabled={}, popupOpen={}, contentText='{}')",
                              static_cast<uint32_t>(snapshot.focusTarget),
                              snapshot.contentModeSelectedIndex ? std::format(L"{}", snapshot.contentModeSelectedIndex.value()) : L"<none>",
                              snapshot.matchCaseContentEnabled ? L"true" : L"false",
                              snapshot.wantSnippetsEnabled ? L"true" : L"false",
                              snapshot.contentModePopupOpen ? L"true" : L"false",
                              snapshot.contentPatternText));

    std::this_thread::sleep_for(SelfTest::Scale(1200ms));
    PumpPendingMessages();
    SendMessageW(findWindow, WM_CHAR, L'd', 0);
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::ContentModeCombo && value.contentModeSelectedIndex.has_value() &&
               value.contentModeSelectedIndex.value() == 0u && ! value.matchCaseContentEnabled && ! value.wantSnippetsEnabled;
    },
                      SelfTest::Scale(2000ms),
                      &snapshot),
                  L"Selecting Disabled via DX combo typeahead should disable the content-specific option row.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogCommandEnablementMatchesIdleRunningAndSelectionStates(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_command_enablement";
    const std::filesystem::path sub  = root / L"sub";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create command-enablement test subdirectory.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create command-enablement result file.");

    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int index = 0; index < 24; ++index)
    {
        const std::filesystem::path filePath =
            (index % 2 == 0) ? (root / std::format(L"root_{:03}.json", index)) : (sub / std::format(L"sub_{:03}.json", index));
        state.Require(SelfTest::WriteTextFile(filePath, largeBody),
                      std::format(L"Failed to create '{}' for command-enablement test.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for command-enablement test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for command-enablement test.");
    if (! state.failure.empty())
    {
        return false;
    }

    RaiseSelfTestWindowForInput(mainWindow);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in command-enablement test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for command-enablement test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, true, true, false, true), L"Failed to configure Find options for command-enablement test.");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                L"*.json",
                                                L"ZZZ_NOT_PRESENT_123456789",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure Find window for command-enablement test.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture idle Find snapshot for command-enablement test.");
    state.Require(snapshot.statusText.empty(), L"Idle Find action-row status should stay blank instead of showing Ready.");
    state.Require(snapshot.statusStripBlendsWithWindowBackground,
                  L"Find action-row status should blend with the window background instead of drawing a card surface.");
    state.Require(snapshot.findButtonEnabled && ! snapshot.appendButtonEnabled && ! snapshot.intersectButtonEnabled && ! snapshot.subtractButtonEnabled,
                  L"Find should be enabled while idle, but result-set action menu options should stay disabled until the result list has items.");
    std::atomic<bool> initialMenuShapeValid{false};
    std::atomic<bool> initialSetOpsDisabled{false};
    state.Require(ProbeFindSplitMenu(findWindow,
                                     findWindow,
                                     state,
                                     [&](const RedSalamander::DxUi::ContextMenuPopupDebugState& popupState) noexcept
    {
        const bool shapeValid = popupState.itemTexts.size() == 4u && popupState.itemEnabled.size() == 4u &&
                                popupState.itemTexts[0] == LoadStringResource(nullptr, IDS_FIND_ACTION_FIND) &&
                                popupState.itemTexts[1] == LoadStringResource(nullptr, IDS_FIND_ACTION_REFINE_INTERSECT) &&
                                popupState.itemTexts[2] == LoadStringResource(nullptr, IDS_FIND_ACTION_REFINE_SUBTRACT) &&
                                popupState.itemTexts[3] == LoadStringResource(nullptr, IDS_FIND_ACTION_APPEND_TO_FOUND);
        initialMenuShapeValid.store(shapeValid, std::memory_order_release);
        initialSetOpsDisabled.store(shapeValid && popupState.itemEnabled[0] && ! popupState.itemEnabled[1] && ! popupState.itemEnabled[2] &&
                                        ! popupState.itemEnabled[3],
                                    std::memory_order_release);
    }),
                  L"Failed to probe initial Find split action menu.");
    state.Require(initialMenuShapeValid.load(std::memory_order_acquire),
                  L"Find split action menu should expose Find, Intersect, Subtract, and Append entries.");
    state.Require(initialSetOpsDisabled.load(std::memory_order_acquire),
                  L"Find split action menu should disable Append/Intersect/Subtract until the result list contains items.");
    state.Require(ProbeFindSplitMenuStationaryHover(findWindow, findWindow, 0u, state),
                  L"Find split action menu should highlight the item under a stationary pointer.");
    state.Require(ProbeFindSplitMenuLivePointerRouting(findWindow, findWindow, 0u, state),
                  L"Find split action menu should process live pointer hover and light-dismiss messages.");
    state.Require(! snapshot.cancelButtonEnabled && ! snapshot.openButtonEnabled && ! snapshot.parentButtonEnabled,
                  L"Cancel/Open/Parent buttons should be disabled while idle with no selection.");
    state.Require(snapshot.rootComboEnabled && snapshot.nameComboEnabled && snapshot.nameModeComboEnabled && snapshot.contentComboEnabled &&
                      snapshot.contentModeComboEnabled,
                  L"Form combos should remain enabled while the Find window is idle.");
    state.Require(snapshot.matchCaseContentEnabled && snapshot.wantSnippetsEnabled,
                  L"Content-mode-dependent controls should be enabled when content mode is active.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for command-enablement test.");
    const auto hasActiveCommandState = [](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.searchActive && ! value.findButtonEnabled && ! value.appendButtonEnabled && ! value.intersectButtonEnabled &&
               ! value.subtractButtonEnabled && value.cancelButtonEnabled && ! value.openButtonEnabled && ! value.parentButtonEnabled &&
               ! value.rootComboEnabled && ! value.nameComboEnabled && ! value.nameModeComboEnabled && ! value.contentComboEnabled &&
               ! value.contentModeComboEnabled;
    };
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture immediate active Find command-state snapshot.");
    if (state.failure.empty() && ! hasActiveCommandState(snapshot))
    {
        state.Require(WaitForFindSnapshot(hasActiveCommandState, SelfTest::Scale(10000ms), &snapshot),
                      std::format(L"Active Find search did not expose the expected disabled/enabled command state. "
                                  L"find={} append={} intersect={} subtract={} cancel={} open={} parent={} rootCombo={} nameCombo={} nameMode={} "
                                  L"contentCombo={} contentMode={} {}",
                                  snapshot.findButtonEnabled ? 1 : 0,
                                  snapshot.appendButtonEnabled ? 1 : 0,
                                  snapshot.intersectButtonEnabled ? 1 : 0,
                                  snapshot.subtractButtonEnabled ? 1 : 0,
                                  snapshot.cancelButtonEnabled ? 1 : 0,
                                  snapshot.openButtonEnabled ? 1 : 0,
                                  snapshot.parentButtonEnabled ? 1 : 0,
                                  snapshot.rootComboEnabled ? 1 : 0,
                                  snapshot.nameComboEnabled ? 1 : 0,
                                  snapshot.nameModeComboEnabled ? 1 : 0,
                                  snapshot.contentComboEnabled ? 1 : 0,
                                  snapshot.contentModeComboEnabled ? 1 : 0,
                                  DescribeFindSnapshotBrief(snapshot)));
    }
    if (! state.failure.empty())
    {
        static_cast<void>(DebugCancelFindFilesWindowSearch());
        static_cast<void>(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())));
        return false;
    }
    RECT activeCancelRect{};
    RECT activeHelpRect{};
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::CancelButton, activeCancelRect),
                  L"Active Find Cancel button should expose a visible debug target rectangle.");
    state.Require(DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget::HelpButton, activeHelpRect),
                  L"Active Find help button should expose a visible debug target rectangle.");
    const LONG activeCancelHelpGapPx    = activeHelpRect.left - activeCancelRect.right;
    const LONG maxActiveCancelHelpGapPx = MulDiv(40, static_cast<int>(GetDpiForWindow(findWindow)), 96);
    state.Require(activeCancelRect.right <= activeHelpRect.left && activeCancelHelpGapPx >= 0 && activeCancelHelpGapPx <= maxActiveCancelHelpGapPx,
                  std::format(L"Active Find Cancel button should appear immediately left of the help button; gap was {} px.", activeCancelHelpGapPx));

    state.Require(DebugCancelFindFilesWindowSearch(), L"Failed to cancel Find search for command-enablement test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find window did not become idle after cancellation in command-enablement test.");

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to reconfigure Find window for selection-state command-enablement test.");
    state.Require(DebugSetFindFilesWindowOptions(true, false, true, false, false),
                  L"Failed to configure selection-state Find options for command-enablement test.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start selection-state Find search for command-enablement test.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Selection-state Find search did not become idle in command-enablement test.");
    state.Require(DebugSelectFindFilesWindowResult(sub.native()), std::format(L"Failed to select '{}' for command-enablement test.", sub.native()));
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept
    {
        return ! value.searchActive && value.findButtonEnabled && value.appendButtonEnabled && value.intersectButtonEnabled && value.subtractButtonEnabled &&
               ! value.cancelButtonEnabled && value.openButtonEnabled && value.parentButtonEnabled && value.selectedResultCount == 1u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Selecting a Find result did not restore idle commands and enable Open/Parent.");
    std::atomic<bool> populatedSetOpsEnabled{false};
    state.Require(ProbeFindSplitMenu(findWindow,
                                     findWindow,
                                     state,
                                     [&](const RedSalamander::DxUi::ContextMenuPopupDebugState& popupState) noexcept
    {
        populatedSetOpsEnabled.store(popupState.itemEnabled.size() == 4u && popupState.itemEnabled[0] && popupState.itemEnabled[1] &&
                                         popupState.itemEnabled[2] && popupState.itemEnabled[3],
                                     std::memory_order_release);
    }),
                  L"Failed to probe populated Find split action menu.");
    state.Require(populatedSetOpsEnabled.load(std::memory_order_acquire),
                  L"Find split action menu should enable Append/Intersect/Subtract after the result list contains items.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogActionButtonsActivateExpectedCommands(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Find action-button command activation validation"))
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_action_buttons";
    const std::filesystem::path sub  = root / L"sub";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sub), L"Failed to create action-buttons test directory.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.jsonl", "payload"), L"Failed to create alpha.jsonl.");
    state.Require(SelfTest::WriteTextFile(root / L"apple.txt", "payload"), L"Failed to create apple.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "payload"), L"Failed to create beta.txt.");
    state.Require(SelfTest::WriteTextFile(sub / L"inside.txt", "payload"), L"Failed to create nested test file.");

    std::string largeBody(2u * 1024u * 1024u, 'x');
    largeBody.back() = '\n';
    for (int index = 0; index < 24; ++index)
    {
        const std::filesystem::path filePath =
            (index % 2 == 0) ? (root / std::format(L"cancel_{:03}.json", index)) : (sub / std::format(L"cancel_{:03}.json", index));
        state.Require(SelfTest::WriteTextFile(filePath, largeBody),
                      std::format(L"Failed to create '{}' for action-buttons cancellation coverage.", filePath.filename().native()));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for action-buttons test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for action-buttons test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.jsonl", L"apple.txt", L"beta.txt", L"sub"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for action-buttons test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/find"), L"Shortcut dispatch failed for cmd/pane/find in action-buttons test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for action-buttons test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::filesystem::path& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path.native()) != snapshot.fullPaths.end(); };

    const auto clickButton = [&](FindFilesDebugFocusTarget target, std::wstring_view label) noexcept
    { SendFindTargetButtonClick(findWindow, target, state, label); };

    FindFilesDebugSnapshot snapshot{};

    state.Require(DebugSetFindFilesWindowOptions(false, true, false, true, false), L"Failed to configure root-level Find options for action-buttons test.");
    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure initial Find search for action-buttons test.");
    Trace(L"action-buttons: invoking initial Find button");
    clickButton(FindFilesDebugFocusTarget::FindButton, L"Find button before initial search");
    Trace(L"action-buttons: waiting for initial Find idle");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Initial Find search did not become idle in action-buttons test.");
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.resultCount == 1u && containsPath(value, root / L"alpha.jsonl"); },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Find button delivered click did not add the expected initial result.");
    const std::filesystem::path alphaJsonl = root / L"alpha.jsonl";
    state.Require(DebugSelectFindFilesWindowResult(alphaJsonl.native()), L"Failed to select alpha.jsonl before file-open disposition validation.");
    FindFilesDebugOpenDisposition fileOpenDisposition = FindFilesDebugOpenDisposition::None;
    state.Require(DebugGetSelectedFindFilesWindowOpenDisposition(false, fileOpenDisposition),
                  L"Find file-open disposition should be available for the selected file result.");
    state.Require(fileOpenDisposition == FindFilesDebugOpenDisposition::DefaultOpenFile,
                  std::format(L"Find Open on a file should default-open the file without pane navigation; disposition was {}.",
                              static_cast<uint32_t>(fileOpenDisposition)));

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure second Append search for action-buttons test.");
    Trace(L"action-buttons: invoking Append menu item");
    state.Require(InvokeFindSplitMenuItem(findWindow, findWindow, 3u, state), L"Find split menu Append item did not invoke.");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"action-buttons: waiting for Append idle");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Second Append search did not become idle in action-buttons test.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 3u && containsPath(value, root / L"alpha.jsonl") && containsPath(value, root / L"apple.txt") &&
               containsPath(value, root / L"beta.txt");
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Append button live UIA invoke did not merge the expected TXT results. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"a*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Intersect search for action-buttons test.");
    Trace(L"action-buttons: invoking Intersect menu item");
    state.Require(InvokeFindSplitMenuItem(findWindow, findWindow, 1u, state), L"Find split menu Intersect item did not invoke.");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"action-buttons: waiting for Intersect idle");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Intersect search did not become idle in action-buttons test.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && containsPath(value, root / L"alpha.jsonl") && containsPath(value, root / L"apple.txt") &&
               ! containsPath(value, root / L"beta.txt");
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Intersect button live UIA invoke did not retain only the expected results. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugConfigureFindFilesWindow(
                      root.native(), L"*.jsonl", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
                  L"Failed to configure Subtract search for action-buttons test.");
    Trace(L"action-buttons: invoking Subtract menu item");
    state.Require(InvokeFindSplitMenuItem(findWindow, findWindow, 2u, state), L"Find split menu Subtract item did not invoke.");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"action-buttons: waiting for Subtract idle");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Subtract search did not become idle in action-buttons test.");
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.resultCount == 1u && containsPath(value, root / L"apple.txt") && ! containsPath(value, root / L"alpha.jsonl"); },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  std::format(L"Subtract button live UIA invoke did not remove the expected JSONL result. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, true, false, false, false), L"Failed to configure cancellation options for action-buttons test.");
    state.Require(DebugConfigureFindFilesWindow(root.native(),
                                                L"*.json",
                                                L"ZZZ_NOT_PRESENT_123456789",
                                                Common::Settings::SearchNameMode::Wildcard,
                                                Common::Settings::SearchContentMode::TextLiteral),
                  L"Failed to configure cancellation search for action-buttons test.");
    Trace(L"action-buttons: invoking cancellation Find button");
    clickButton(FindFilesDebugFocusTarget::FindButton, L"Find button before cancellation search");
    Trace(L"action-buttons: waiting for cancellation search active-or-complete");
    const bool cancellationSearchObserved = WaitForFindSnapshot(
        [](const FindFilesDebugSnapshot& value) noexcept
    {
        const bool submittedCancellationQuery =
            value.submittedNamePatternText == L"*.json" && value.submittedContentPatternText == L"ZZZ_NOT_PRESENT_123456789";
        if (value.searchActive)
        {
            return submittedCancellationQuery && value.cancelButtonEnabled && ! value.findButtonEnabled;
        }

        // On fast machines the synthetic content scan can complete before the
        // polling loop observes the active state. That still proves the UIA
        // Find invoke reached BeginSearch; Cancel coverage remains conditional
        // on actually seeing the active state.
        return submittedCancellationQuery && ! value.cancelButtonEnabled && value.findButtonEnabled;
    },
        SelfTest::Scale(10000ms),
        &snapshot);
    state.Require(cancellationSearchObserved,
                  std::format(L"Find button delivered click did not start or complete the cancellable search. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.searchActive)
    {
        Trace(L"action-buttons: invoking Cancel button");
        clickButton(FindFilesDebugFocusTarget::CancelButton, L"Cancel button during active search");
        Trace(L"action-buttons: waiting for Cancel idle");
        state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                      L"Cancel button delivered click did not return Find to idle.");
    }
    Trace(L"action-buttons: verifying cancellation idle state");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return ! value.searchActive && ! value.cancelButtonEnabled && value.findButtonEnabled; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  std::format(L"Cancellation search did not restore the idle Find command state. {}", DescribeFindSnapshotBrief(snapshot)));

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false), L"Failed to configure directory-result options for action-buttons test.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"sub", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure directory-result search for action-buttons test.");
    Trace(L"action-buttons: invoking directory Find button");
    clickButton(FindFilesDebugFocusTarget::FindButton, L"Find button before directory-result search");
    Trace(L"action-buttons: waiting for directory Find idle");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Directory-result search did not become idle in action-buttons test.");
    FindFilesDebugSnapshot directorySnapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return ! value.searchActive && value.resultCount == 1u && containsPath(value, sub); },
                                      SelfTest::Scale(3000ms),
                                      &directorySnapshot),
                  std::format(L"Directory-result search did not expose the expected Find result before selection. {}",
                              DescribeFindSnapshotBrief(directorySnapshot)));
    state.Require(DebugSelectFindFilesWindowResult(sub.native()), std::format(L"Failed to select '{}' for action-buttons test.", sub.native()));
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && containsPath(value, sub); },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Selecting the directory result did not enable Open/Parent for action-buttons test.");

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    Trace(L"action-buttons: invoking Open button");
    clickButton(FindFilesDebugFocusTarget::OpenButton, L"Open button before directory activation");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sub, SelfTest::Scale(5000ms)),
                  std::format(L"Open button delivered click did not navigate the focused pane into the selected directory; focusedPane={} activePane={} left='{}' right='{}'.",
                              static_cast<int>(g_folderWindow.GetFocusedPane()),
                              static_cast<int>(g_folderWindow.GetActivePane()),
                              g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left).value_or(std::filesystem::path{}).native(),
                              g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right).value_or(std::filesystem::path{}).native()));

    state.Require(DebugSelectFindFilesWindowResult(sub.native()),
                  std::format(L"Failed to reselect '{}' before Parent action in action-buttons test.", sub.native()));
    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    Trace(L"action-buttons: invoking Go to folder button");
    clickButton(FindFilesDebugFocusTarget::ParentButton, L"Parent button before parent navigation");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                  std::format(L"Parent button delivered click did not return the focused pane to the selected directory parent; focusedPane={} activePane={} "
                              L"left='{}' right='{}'.",
                              static_cast<int>(g_folderWindow.GetFocusedPane()),
                              static_cast<int>(g_folderWindow.GetActivePane()),
                              g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left).value_or(std::filesystem::path{}).native(),
                              g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right).value_or(std::filesystem::path{}).native()));
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.jsonl", L"apple.txt", L"beta.txt", L"sub"}, SelfTest::Scale(3000ms)),
                  L"Parent button delivered click did not restore the parent directory contents.");

    Trace(L"action-buttons: completed");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresPersistedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_persisted_grid_layout";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create persisted-grid-layout test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode          = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode       = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout = {
        Common::Settings::GridColumnLayoutEntry{.columnId = L"path", .displayIndex = 0u, .widthDip = 420.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"modified", .displayIndex = 1u, .widthDip = 180.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"name", .displayIndex = 2u, .widthDip = 260.0f},
    };
    g_settings.search = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-persisted-grid-layout");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for persisted-grid-layout test.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for persisted-grid-layout test.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [](const FindFilesDebugSnapshot& value) noexcept { return value.resultColumnIds.size() >= 5u; }, SelfTest::Scale(2000ms), &snapshot),
                  L"Find snapshot did not expose grid columns for persisted-grid-layout test.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.resultColumnIds.size() == snapshot.resultColumnWidthsDip.size(), L"Grid layout snapshot shape mismatch.");
    state.Require(snapshot.resultColumnIds.size() >= 5u, L"Expected base Find grid columns in persisted-grid-layout snapshot.");
    if (snapshot.resultColumnIds.size() >= 5u && snapshot.resultColumnWidthsDip.size() >= 5u)
    {
        state.Require(snapshot.resultColumnIds[0] == L"path", L"Persisted Find grid layout did not move Path to display index 0.");
        state.Require(snapshot.resultColumnIds[1] == L"modified", L"Persisted Find grid layout did not move Modified to display index 1.");
        state.Require(snapshot.resultColumnIds[2] == L"name", L"Persisted Find grid layout did not move Name to display index 2.");
        state.Require(std::fabs(snapshot.resultColumnWidthsDip[0] - 420.0f) <= 0.1f, L"Persisted Find grid layout did not restore the Path column width.");
        state.Require(std::fabs(snapshot.resultColumnWidthsDip[1] - 180.0f) <= 0.1f, L"Persisted Find grid layout did not restore the Modified column width.");
        state.Require(std::fabs(snapshot.resultColumnWidthsDip[2] - 260.0f) <= 0.1f, L"Persisted Find grid layout did not restore the Name column width.");
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Find grid layout close path should keep search settings present.");
    if (g_settings.search.has_value())
    {
        const auto& savedLayout = g_settings.search->resultsGridLayout;
        state.Require(savedLayout.size() >= 5u, L"Closing the Find window should persist the active grid layout.");
        if (savedLayout.size() >= 3u)
        {
            state.Require(savedLayout[0].columnId == L"path", L"Persisted Find layout should keep Path first after close.");
            state.Require(savedLayout[1].columnId == L"modified", L"Persisted Find layout should keep Modified second after close.");
            state.Require(savedLayout[2].columnId == L"name", L"Persisted Find layout should keep Name third after close.");
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogHeaderDragReordersColumnsWithoutSort(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto restoreSearch                                                   = wil::scope_exit([&] noexcept { g_settings.search = previousSearch; });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_header_drag_reorder";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find header-reorder test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find header-reorder subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find header reorder.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find header reorder.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-header-reorder");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for header reorder validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for header reorder validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for header reorder validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for header reorder validation.");
    if (! DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find))
    {
        FindFilesDebugSnapshot startFailureSnapshot{};
        static_cast<void>(DebugGetFindFilesWindowSnapshot(startFailureSnapshot));
        state.Require(false, std::format(L"Failed to start Find search for header reorder validation. {}", DescribeFindSnapshotBrief(startFailureSnapshot)));
    }
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for header reorder validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for header reorder validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot before header reorder validation.");
    state.Require(snapshot.usesDxUiHost && snapshot.resultCount == 2u && snapshot.selectedResultCount == 1u && snapshot.visibleChildWindowCount <= 1u &&
                      snapshot.resultColumnIds.size() >= 2u && snapshot.firstResultHeaderRect.right > snapshot.firstResultHeaderRect.left &&
                      snapshot.secondResultHeaderRect.right > snapshot.secondResultHeaderRect.left && snapshot.resultColumnIds[0] == L"name" &&
                      snapshot.resultColumnIds[1] == L"path" && snapshot.dxResizeFailureCount == 0u && containsPath(snapshot, selectedPath),
                  std::format(L"Find window did not expose the expected baseline results-grid state before header reorder validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    const bool sawNestedPathText =
        std::find(snapshot.resultPathTexts.begin(), snapshot.resultPathTexts.end(), std::wstring(L"sub")) != snapshot.resultPathTexts.end();
    state.Require(sawNestedPathText, L"Recursive Find results should show the containing subfolder in the Path column.");
    state.Require(snapshot.visibleResultIconCellCount >= snapshot.visibleResultRowCount && snapshot.visibleResultRowCount > 0u,
                  L"Find results should render a file/folder icon cell for each visible row.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    state.Require(DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
                  L"Failed to reorder the visible Path column ahead of Name through the Find results grid debug seam.");

    const bool reordered = WaitForFindSnapshot(
        [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
        SelfTest::Scale(10000ms),
        &snapshot);
    state.Require(reordered,
                  std::format(L"Find header drag should reorder visible columns without losing selection or bounded visible work. {}",
                              DescribeFindSnapshotBrief(snapshot)));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogCopyFollowsReorderedColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    if (SkipIfFindSelfTestClipboardUnavailable(state))
    {
        return true;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto restoreSearch                                                   = wil::scope_exit([&] noexcept { g_settings.search = previousSearch; });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_copy_reordered_columns";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reordered-copy test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find reordered-copy subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reordered-copy.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find reordered-copy.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-copy-reordered-columns");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reordered-copy validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reordered-copy validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reordered-copy validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reordered-copy validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reordered-copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reordered-copy validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath         = (root / L"sub" / L"beta.txt").native();
    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for reordered-copy validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the expected baseline results-grid state before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const D2D1_RECT_F headerRect = snapshot.secondResultHeaderRect;
    const LPARAM dragStart       = DipPointToClientLParam(findWindow, (headerRect.left + headerRect.right) * 0.5f, (headerRect.top + headerRect.bottom) * 0.5f);
    const LPARAM dragTarget      = DipPointToClientLParam(findWindow, snapshot.firstResultHeaderRect.left + 12.0f, (headerRect.top + headerRect.bottom) * 0.5f);

    SendMouseDragToResolvedPointWindow(findWindow, dragStart, dragTarget);

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not settle on the reordered visible column order before reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the Find results grid before reordered-copy validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  L"Find results grid did not take focus before reordered-copy validation.");

    ClearClipboardContents(findWindow);

    std::wstring copiedSelection;
    for (size_t commandAttempt = 0u; commandAttempt < 3u && copiedSelection.empty(); ++commandAttempt)
    {
        SendFindResultCommand(findWindow, IDM_PANE_CLIPBOARD_COPY);
        for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
        {
            PumpPendingMessages();
            copiedSelection = ReadClipboardUnicodeText(findWindow);
            if (copiedSelection.empty())
            {
                std::this_thread::sleep_for(20ms);
            }
        }
    }

    state.Require(! copiedSelection.empty(), L"Find copy command should copy the reordered visible row content to the clipboard.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Find clipboard copy should start with the visible Path column after header reorder.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Find clipboard copy should still include the selected result name after header reorder.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogReorderedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto restoreSearch                                                   = wil::scope_exit([&] noexcept { g_settings.search = previousSearch; });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_reordered_columns_survive_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reorder/sort test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find reorder/sort subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reorder/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find reorder/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-reordered-sort-cycles");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reorder/sort validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reorder/sort validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reorder/sort validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reorder/sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reorder/sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder/sort validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    const std::wstring otherPath    = (root / L"alpha.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for reorder/sort validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath) && containsPath(value, otherPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before reorder/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;

    state.Require(DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
                  L"Failed to reorder the visible Path column ahead of Name through the deterministic debug layout path before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not settle on the reordered visible column order before sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.left + snapshot.firstResultHeaderRect.right) * 0.5f));
    const LONG clickY = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.top + snapshot.firstResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && value.fullPaths.size() >= 2u && value.fullPaths[0] == selectedPath && value.fullPaths[1] == otherPath &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find reordered-column sort cycles should keep the visible display order while sorting the visible Path column.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogHeaderResizeChangesVisibleWidth(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto restoreSearch                                                   = wil::scope_exit([&] noexcept { g_settings.search = previousSearch; });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_header_resize";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find header-resize test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find header-resize subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find header resize.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find header resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-header-resize");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for header resize validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for header resize validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for header resize validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for header resize validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for header resize validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for header resize validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for header resize validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the expected baseline results-grid state before header resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    const float baselineFirstHeaderWidth    = std::max(0.0f, snapshot.firstResultHeaderRect.right - snapshot.firstResultHeaderRect.left);
    const float baselineFirstColumnWidthDip = snapshot.resultColumnWidthsDip[0];
    const float baselineSecondHeaderLeft    = snapshot.secondResultHeaderRect.left;

    const auto waitForHeaderResize = [&]() noexcept
    {
        return WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            const float newFirstHeaderWidth = std::max(0.0f, value.firstResultHeaderRect.right - value.firstResultHeaderRect.left);
            return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
                   value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
                   newFirstHeaderWidth >= baselineFirstHeaderWidth + 20.0f && value.resultColumnWidthsDip[0] >= baselineFirstColumnWidthDip + 20.0f &&
                   value.secondResultHeaderRect.left > baselineSecondHeaderLeft + 10.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
                   value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
                   value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
                   value.visibleResultCellCount <= baselineVisibleCellCount &&
                   value.visibleResultCellCount <= value.visibleResultRowCount * value.visibleResultColumnCount && value.dxResizeFailureCount == 0u &&
                   containsPath(value, selectedPath);
        },
            SelfTest::Scale(3000ms),
            &snapshot);
    };

    state.Require(SendFindFirstVisibleHeaderResizeDrag(findWindow, snapshot, 48.0f),
                  L"Failed to locate the live Find header-resize grip before header resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! waitForHeaderResize())
    {
        state.Require(SendFindFirstVisibleHeaderResizeDragToHost(findWindow, snapshot, 48.0f),
                      L"Failed to retry the live Find header-resize drag against the host window.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const float currentFirstHeaderWidth = std::max(0.0f, snapshot.firstResultHeaderRect.right - snapshot.firstResultHeaderRect.left);
    const float currentSecondHeaderLeft = snapshot.secondResultHeaderRect.left;
    state.Require(waitForHeaderResize(),
                  std::format(L"Find header resize should widen the visible Name column without reordering columns or losing the selected row. "
                              L"baselineFirstHeaderW={:.1f} currentFirstHeaderW={:.1f} baselineFirstColumnW={:.1f} currentFirstColumnW={:.1f} "
                              L"baselineSecondHeaderLeft={:.1f} currentSecondHeaderLeft={:.1f} baselineVisibleRows={} currentVisibleRows={} "
                              L"baselineVisibleCols={} currentVisibleCols={} baselineVisibleCells={} currentVisibleCells={}. {}",
                              baselineFirstHeaderWidth,
                              currentFirstHeaderWidth,
                              baselineFirstColumnWidthDip,
                              snapshot.resultColumnWidthsDip.size() >= 1u ? snapshot.resultColumnWidthsDip[0] : 0.0f,
                              baselineSecondHeaderLeft,
                              currentSecondHeaderLeft,
                              baselineVisibleRowCount,
                              snapshot.visibleResultRowCount,
                              baselineVisibleColumnCount,
                              snapshot.visibleResultColumnCount,
                              baselineVisibleCellCount,
                              snapshot.visibleResultCellCount,
                              DescribeFindSnapshotBrief(snapshot)));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogReorderedColumnsSurviveSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_reordered_columns_survive_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reorder/rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find reorder/rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reorder/rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find reorder/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-reordered-search-rerun");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reorder/rerun validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reorder/rerun validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reorder/rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reorder/rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reorder/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder/rerun validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before reorder/rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;

    state.Require(DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
                  L"Failed to reorder the visible Path column ahead of Name through the deterministic debug layout path before search rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find header drag did not settle on the reordered visible column order before search rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to focus the Find name combo before reordered-layout rerun validation.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"beta*"),
                  L"Failed to narrow the Find name combo text for reordered-layout rerun validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::NameCombo && value.namePatternText == L"beta*" && value.rootText == root.native() &&
               value.contentPatternText.empty() && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the narrowed beta* query before reordered-layout rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to rerun Find search for reordered-layout rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle for reordered-layout rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.visibleResultColumnCount == baselineVisibleColumnCount && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find reordered visible column order should survive a narrowed search rerun. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to refocus the Find name combo before restoring the reordered-layout baseline search.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"*.txt"),
                  L"Failed to restore the Find baseline name combo text after reordered-layout rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after reordered-layout rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after reordered-layout rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find reordered visible column order should restore after returning to the baseline search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogResizedColumnsSurviveSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_resized_columns_survive_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find resize/rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find resize/rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find resize/rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find resize/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-resized-search-rerun");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for resize/rerun validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for resize/rerun validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for resize/rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for resize/rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for resize/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for resize/rerun validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before resize/rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    const float baselineFirstColumnWidthDip = snapshot.resultColumnWidthsDip[0];
    const auto waitForHeaderResize          = [&]() noexcept
    {
        return WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
                   value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
                   value.resultColumnWidthsDip[0] >= baselineFirstColumnWidthDip + 20.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
                   value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
                   value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
                   value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
        },
            SelfTest::Scale(3000ms),
            &snapshot);
    };

    state.Require(SendFindFirstVisibleHeaderResizeDrag(findWindow, snapshot, 48.0f),
                  L"Failed to locate the live Find header-resize grip before resize/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! waitForHeaderResize())
    {
        state.Require(SendFindFirstVisibleHeaderResizeDragToHost(findWindow, snapshot, 48.0f),
                      L"Failed to retry the live Find header-resize drag against the host window before resize/rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        waitForHeaderResize(),
        std::format(L"Find deterministic resize/rerun validation did not settle on the widened Name column. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to focus the Find name combo before resized-layout rerun validation.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"beta*"),
                  L"Failed to narrow the Find name combo text for resized-layout rerun validation.");
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot) && snapshot.namePatternText == L"beta*",
                  std::format(L"Find window did not accept the narrowed beta* query immediately before resized-layout rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.focusTarget == FindFilesDebugFocusTarget::NameCombo && value.namePatternText == L"beta*" && value.rootText == root.native() &&
               value.contentPatternText.empty() && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnWidthsDip[0] >= widenedWidthDip - 1.0f &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the narrowed beta* query before resized-layout rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to rerun Find search for resized-layout rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle for resized-layout rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
               value.resultColumnWidthsDip[0] >= widenedWidthDip - 1.0f && value.visibleResultColumnCount > 0u &&
               value.visibleResultColumnCount <= baselineVisibleColumnCount && value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find widened Name column should survive a narrowed search rerun. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to refocus the Find name combo before restoring the resized-layout baseline search.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"*.txt"),
                  L"Failed to restore the Find baseline name combo text after resized-layout rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after resized-layout rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after resized-layout rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
               value.resultColumnWidthsDip[0] >= widenedWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
               value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
               value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find widened Name column should restore after returning to the baseline search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogReorderedResizedColumnsSurviveSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_reordered_resized_columns_survive_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reorder+resize/rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create Find reorder+resize/rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reorder+resize/rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for Find reorder+resize/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-reordered-resized-search-rerun");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reorder+resize/rerun validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reorder+resize/rerun validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reorder+resize/rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reorder+resize/rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reorder+resize/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder+resize/rerun validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before reorder+resize/rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    state.Require(DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
                  L"Failed to reorder the visible Path column ahead of Name before combined-view-state persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not settle on the reordered visible column order before reorder+resize/rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        std::format(L"Find reordered first visible column did not widen before reorder+resize/rerun validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to rerun Find search for reorder+resize/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder+resize/rerun validation.");

    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= widenedFirstVisibleWidthDip - 1.0f && value.visibleResultColumnCount > 0u &&
               value.visibleResultColumnCount <= baselineVisibleColumnCount && value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
            SelfTest::Scale(10000ms),
            &snapshot),
        std::format(L"Find reordered visible layout and widened first column should survive a search rerun together. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::NameCombo),
                  L"Failed to refocus the Find name combo before restoring the reorder+resize baseline search.");
    state.Require(DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget::NameCombo, L"*.txt"),
                  L"Failed to restore the Find baseline name combo text after reorder+resize/rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after reorder+resize/rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after reorder+resize/rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= widenedFirstVisibleWidthDip - 1.0f && value.visibleResultColumnCount > 0u &&
               value.visibleResultColumnCount <= baselineVisibleColumnCount && value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount &&
               value.visibleResultCellCount > 0u && value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find reordered visible layout and widened first column should restore after returning to the baseline search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogReorderedResizedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_reordered_resized_columns_survive_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find reorder+resize/sort test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find reorder+resize/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find reorder+resize/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for Find reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-reordered-resized-sort-cycles");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reorder+resize/sort validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reorder+resize/sort validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reorder+resize/sort validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reorder+resize/sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reorder+resize/sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reorder+resize/sort validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for reorder+resize/sort validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;

    state.Require(
        DebugReorderFindFilesWindowVisibleResultColumn(1u, 0u),
        L"Failed to reorder the visible Path column ahead of Name through the deterministic debug layout path before reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not settle on the reordered visible column order before reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic reorder+resize/sort validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];
    const LONG sortClickX = static_cast<LONG>(std::lround((snapshot.secondResultHeaderRect.left + snapshot.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((snapshot.secondResultHeaderRect.top + snapshot.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    const std::wstring firstExpected  = (root / L"gamma.txt").native();
    const std::wstring secondExpected = (root / L"beta.txt").native();
    const std::wstring thirdExpected  = (root / L"alpha.txt").native();
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= widenedFirstVisibleWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
               value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
               value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u && value.fullPaths.size() >= 3u &&
               value.fullPaths[0] == firstExpected && value.fullPaths[1] == secondExpected && value.fullPaths[2] == thirdExpected &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find sort cycles should keep the reordered visible layout and widened first visible column intact together. {}",
                              DescribeFindSnapshotBrief(snapshot)));

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogResizedColumnsSurviveSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_resized_columns_survive_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find resize/sort test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find resize/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find resize/sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for Find resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    FindFilesPaneContext context{};
    context.rootPluginPath = root;
    const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-resized-sort-cycles");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for resize/sort validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for resize/sort validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for resize/sort validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for resize/sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for resize/sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for resize/sort validation.");

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::wstring selectedPath = (root / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for resize/sort validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before resize/sort validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = snapshot.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = snapshot.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = snapshot.visibleResultCellCount;
    const float baselineFirstColumnWidthDip = snapshot.resultColumnWidthsDip[0];
    const auto waitForResizedNameColumn     = [&]() noexcept
    {
        return WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" &&
                   value.resultColumnIds[1] == L"path" && value.resultColumnWidthsDip[0] >= baselineFirstColumnWidthDip + 20.0f &&
                   value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount > 0u &&
                   value.visibleResultColumnCount <= baselineVisibleColumnCount && value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount &&
                   value.visibleResultCellCount > 0u && value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u &&
                   containsPath(value, selectedPath);
        },
            SelfTest::Scale(3000ms),
            &snapshot);
    };

    state.Require(SendFindFirstVisibleHeaderResizeDrag(findWindow, snapshot, 48.0f),
                  L"Failed to locate the live Find header-resize grip before reorder+resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! waitForResizedNameColumn())
    {
        state.Require(SendFindFirstVisibleHeaderResizeDragToHost(findWindow, snapshot, 48.0f),
                      L"Failed to retry the live Find header-resize drag against the host window before reorder+resize/sort validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        waitForResizedNameColumn(),
        std::format(L"Find header resize did not settle on the widened Name column before resize/sort validation. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const float widenedWidthDip = snapshot.resultColumnWidthsDip[0];
    state.Require(DebugSetFindFilesWindowResultSort(0u, true), L"Failed to enable descending Name sort before resize/sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring firstExpected  = (root / L"gamma.txt").native();
    const std::wstring secondExpected = (root / L"beta.txt").native();
    const std::wstring thirdExpected  = (root / L"alpha.txt").native();
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 3u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
               value.resultColumnWidthsDip[0] >= widenedWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount > 0u && value.visibleResultColumnCount <= baselineVisibleColumnCount &&
               value.visibleResultColumnCount + 1u >= baselineVisibleColumnCount && value.visibleResultCellCount > 0u &&
               value.visibleResultCellCount <= baselineVisibleCellCount && value.dxResizeFailureCount == 0u && value.fullPaths.size() >= 3u &&
               value.fullPaths[0] == firstExpected && value.fullPaths[1] == secondExpected && value.fullPaths[2] == thirdExpected &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find sort cycles should keep the widened Name column intact while sorting the visible Name header.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogHeaderClickSortsResults(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_header_click_sort";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Find header-click-sort test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for Find header-click sort.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for Find header-click sort.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for Find header-click sort.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = previousSearch.value_or(Common::Settings::SearchDialogSettings{});
    search.resultsGridLayout.clear();
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.nameMode       = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode    = Common::Settings::SearchContentMode::Disabled;
    g_settings.search     = search;

    std::optional<std::filesystem::path> leftBefore;
    const auto restorePath = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    HWND findWindow = nullptr;
    state.Require(OpenFindWindowFromLocalPaneRoot(mainWindow, root, {L"alpha.txt", L"beta.txt", L"gamma.txt"}, findWindow, leftBefore),
                  L"Find window did not open from a deterministic local pane root for header-click sort validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for header-click sort validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for header-click sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for header-click sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for header-click sort validation.");

    const std::wstring selectedPath = (root / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for header-click sort validation.", selectedPath));

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot before header-click sort validation.");
    state.Require(snapshot.usesDxUiHost && snapshot.resultCount == 3u && snapshot.selectedResultCount == 1u && snapshot.visibleChildWindowCount <= 1u &&
                      snapshot.resultColumnIds.size() >= 2u && snapshot.firstResultHeaderRect.right > snapshot.firstResultHeaderRect.left &&
                      snapshot.resultColumnIds[0] == L"name" && snapshot.resultColumnIds[1] == L"path" && snapshot.dxResizeFailureCount == 0u,
                  std::format(L"Find window did not expose the baseline results-grid state before header-click sort validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const LONG clickX = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.left + snapshot.firstResultHeaderRect.right) * 0.5f));
    const LONG clickY = static_cast<LONG>(std::lround((snapshot.firstResultHeaderRect.top + snapshot.firstResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(clickX, clickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(clickX, clickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" &&
               value.dxResizeFailureCount == 0u && value.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), value.fullPaths.begin());
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header clicks did not sort the visible results into descending Name order while preserving selection and visible layout.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresResizedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_resized_grid_layout";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create resized-layout test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create resized-layout subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for resized-layout validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for resized-layout validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    std::optional<std::filesystem::path> leftBefore;
    const auto restorePath = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for resized-layout validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for resized-layout validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-resized-layout-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for resized-layout validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for resized-layout validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for resized-layout validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for resized-layout validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for resized-layout validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before resized-layout persistence validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstHeaderWidth    = std::max(0.0f, snapshot.firstResultHeaderRect.right - snapshot.firstResultHeaderRect.left);
    const float baselineFirstColumnWidthDip = snapshot.resultColumnWidthsDip[0];
    state.Require(SendFindFirstVisibleHeaderResizeDrag(findWindow, snapshot, 48.0f),
                  L"Failed to locate the live Find header-resize grip before resized-layout persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        const float newFirstHeaderWidth = std::max(0.0f, value.firstResultHeaderRect.right - value.firstResultHeaderRect.left);
        return value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && newFirstHeaderWidth >= baselineFirstHeaderWidth + 20.0f &&
               value.resultColumnWidthsDip[0] >= baselineFirstColumnWidthDip + 20.0f && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header resize did not widen the visible Name column before persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedWidthDip = snapshot.resultColumnWidthsDip[0];

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after a live header resize.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    const auto& savedLayout = g_settings.search->resultsGridLayout;
    state.Require(savedLayout.size() >= 2u, L"Closing the Find window should persist the resized grid layout.");
    if (savedLayout.size() >= 2u)
    {
        state.Require(savedLayout[0].columnId == L"name", L"Persisted resized Find layout should keep Name first after close.");
        state.Require(savedLayout[1].columnId == L"path", L"Persisted resized Find layout should keep Path second after close.");
        state.Require(std::fabs(savedLayout[0].widthDip - resizedWidthDip) <= 1.0f,
                      L"Persisted resized Find layout should keep the widened Name column width after close.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-resized-layout-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && std::fabs(value.resultColumnWidthsDip[0] - resizedWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the persisted resized Name column width.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresReorderedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_reordered_grid_layout";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create reordered-layout test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create reordered-layout subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for reordered-layout validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for reordered-layout validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for reordered-layout validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reordered-layout validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-reordered-layout-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for reordered-layout validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for reordered-layout validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reordered-layout validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reordered-layout validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for reordered-layout validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before reordered-layout persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SendFindSecondVisibleHeaderAheadOfFirst(findWindow, snapshot),
                  L"Failed to drag the live Find Path header ahead of Name before persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not reorder the live results-grid columns before persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after a live header reorder.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    const auto& savedLayout = g_settings.search->resultsGridLayout;
    state.Require(savedLayout.size() >= 2u, L"Closing the Find window should persist the reordered grid layout.");
    if (savedLayout.size() >= 2u)
    {
        state.Require(savedLayout[0].columnId == L"path", L"Persisted reordered Find layout should keep Path first after close.");
        state.Require(savedLayout[1].columnId == L"name", L"Persisted reordered Find layout should keep Name second after close.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-reordered-layout-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the persisted reordered column layout.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredReorderedLayoutCopyFollowsVisibleColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    if (SkipIfFindSelfTestClipboardUnavailable(state))
    {
        return true;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_reordered_copy";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored-reordered-copy test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored-reordered-copy subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for restored-reordered-copy validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for restored-reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern.clear();
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored-reordered-copy validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for restored-reordered-copy validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-reordered-copy-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored-reordered-copy validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for restored-reordered-copy validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for restored-reordered-copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored-reordered-copy validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for restored-reordered-copy validation.", selectedPath));

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.secondResultHeaderRect.right > value.secondResultHeaderRect.left && value.resultColumnIds[0] == L"name" &&
               value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored-reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SendFindSecondVisibleHeaderAheadOfFirst(findWindow, snapshot),
                  L"Failed to drag the live Find Path header ahead of Name before restored-reordered-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find header drag did not reorder the live results-grid columns before restored-reordered-copy persistence validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after reordered-copy setup.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->resultsGridLayout.size() >= 2u,
                  L"Closing the Find window should persist the reordered layout before restored-copy validation.");
    if (g_settings.search->resultsGridLayout.size() >= 2u)
    {
        state.Require(g_settings.search->resultsGridLayout[0].columnId == L"path",
                      L"Persisted reordered Find layout should keep Path first before restored-copy validation.");
        state.Require(g_settings.search->resultsGridLayout[1].columnId == L"name",
                      L"Persisted reordered Find layout should keep Name second before restored-copy validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-reordered-copy-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the persisted reordered column layout before restored-copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to reconfigure Find window after restored reordered-layout reopen.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to restore Find options after restored reordered-layout reopen.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored reordered-layout copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Restarted Find search did not become idle for restored reordered-layout copy validation.");
    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored reordered-layout reopen.", selectedPath));

    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 2u && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected reordered results-grid state before restored-copy validation.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the reopened Find results grid before restored reordered-copy validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find results grid did not take focus before restored reordered-copy validation.");

    ClearClipboardContents(findWindow);
    SendFindResultCommand(findWindow, IDM_PANE_CLIPBOARD_COPY);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Find copy command should copy the restored reordered visible row content to the clipboard after reopen.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Reopened Find clipboard copy should start with the restored visible Path column after header reorder persistence.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Reopened Find clipboard copy should still include the selected result name after restored header reorder.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresPersistedSortOrder(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_persisted_sort_order";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create persisted-sort test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for persisted-sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for persisted-sort validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for persisted-sort validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode          = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode       = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId      = L"name";
    search.sortDescending    = true;
    search.resultsGridLayout = {
        Common::Settings::GridColumnLayoutEntry{.columnId = L"path", .displayIndex = 0u, .widthDip = 320.0f},
        Common::Settings::GridColumnLayoutEntry{.columnId = L"name", .displayIndex = 1u, .widthDip = 240.0f},
    };
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)), L"Failed to open Find window for persisted-sort validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for persisted-sort validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-sort-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowResultSort(0u, true), L"Failed to enable descending Name sort for persisted-sort validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for persisted-sort validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for persisted-sort validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value); },
                                      SelfTest::Scale(3000ms),
                                      &snapshot),
                  std::format(L"Find window did not expose the expected descending Name sort order before persistence validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after a live sort change.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->sortColumnId == L"name", L"Closing the Find window should persist the logical Name sort column id.");
    state.Require(g_settings.search->sortDescending, L"Closing the Find window should persist descending sort direction.");
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-sort-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to restart Find search for persisted-sort restore validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for persisted-sort restore validation.");

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value); },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find window did not restore the persisted descending Name sort order.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresReorderedSortedGridLayout(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_reordered_sorted_grid_layout";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create reordered-sorted-layout test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for reordered-sorted-layout validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for reordered-sorted-layout validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for reordered-sorted-layout validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search{};
    search.lastRoot           = root.native();
    search.lastNamePattern    = L"*.txt";
    search.lastContentPattern.clear();
    search.recursive          = true;
    search.includeFiles       = true;
    search.includeDirectories = false;
    search.followSymlinks     = false;
    search.matchCaseName      = false;
    search.matchCaseContent   = false;
    search.preferIndex        = false;
    search.wantSnippets       = false;
    search.nameMode           = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode        = Common::Settings::SearchContentMode::Disabled;
    search.maxResults         = 0u;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for reordered-sorted-layout validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for reordered-sorted-layout validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-reordered-sorted-layout-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for reordered-sorted-layout validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for reordered-sorted-layout validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before reordered-sorted-layout validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  std::format(L"Find deterministic header reorder did not move the visible Path column ahead of Name before combined restore validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowResultSort(0u, true), L"Failed to enable descending logical Name sort for reordered-sorted-layout validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(5000ms),
                      &snapshot),
                  std::format(L"Find logical Name sort did not stay correct after visible header reorder. {}", DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->resultsGridLayout.size() >= 2u, L"Closing the Find window should persist the reordered visible layout.");
    if (g_settings.search->resultsGridLayout.size() >= 2u)
    {
        state.Require(g_settings.search->resultsGridLayout[0].columnId == L"path", L"Persisted combined Find layout should keep Path first after close.");
        state.Require(g_settings.search->resultsGridLayout[1].columnId == L"name", L"Persisted combined Find layout should keep Name second after close.");
    }
    state.Require(g_settings.search->sortColumnId == L"name", L"Persisted combined Find sort should remain bound to the logical Name column id.");
    state.Require(g_settings.search->sortDescending, L"Persisted combined Find sort should remain descending after close.");
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-reordered-sorted-layout-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for reordered-sorted-layout restore validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for reordered-sorted-layout restore validation.");

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered layout plus logical Name sort.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoresCombinedViewState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restores_combined_view_state";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create combined-view-state test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for combined-view-state validation.");
    state.Require(SelfTest::WriteTextFile(root / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for combined-view-state validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"), L"Failed to create gamma.txt for combined-view-state validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    std::optional<std::filesystem::path> leftBefore;
    const auto restorePath = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* /*themeTag*/) noexcept -> HWND
    {
        HWND findWindow = nullptr;
        state.Require(OpenFindWindowFromLocalPaneRoot(mainWindow, root, {L"alpha.txt", L"beta.txt", L"gamma.txt"}, findWindow, leftBefore),
                      L"Find window did not open from a deterministic local pane root for combined-view-state validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-combined-view-state-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot initializationSnapshot{};
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.findButtonEnabled && value.rootComboEnabled && value.nameComboEnabled && value.nameModeComboEnabled; },
                                      SelfTest::Scale(3000ms),
                                      &initializationSnapshot),
                  std::format(L"Find window did not finish initializing its DX shell before combined-view-state validation. {}",
                              DescribeFindSnapshotBrief(initializationSnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for combined-view-state validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for combined-view-state validation.");
    FindFilesDebugSnapshot configuredSnapshot{};
    state.Require(
        WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept
    { return value.usesDxUiHost && value.findButtonEnabled && value.rootComboEnabled && value.nameComboEnabled && value.nameModeComboEnabled; },
                            SelfTest::Scale(1000ms),
                            &configuredSnapshot),
        std::format(L"Find window lost its live DX shell after combined-view-state configuration. {}", DescribeFindSnapshotBrief(configuredSnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for combined-view-state validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for combined-view-state validation.");
    FindFilesDebugSnapshot postSearchSnapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(postSearchSnapshot), L"Failed to capture Find snapshot immediately after combined-view-state search.");
    state.Require(postSearchSnapshot.usesDxUiHost,
                  std::format(L"Find window lost its DX host immediately after combined-view-state search. {}", DescribeFindSnapshotBrief(postSearchSnapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPath = (root / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to select '{}' before combined-view-state baseline validation.", selectedPath));
    FindFilesDebugSnapshot snapshot{};
    state.Require(DebugGetFindFilesWindowSnapshot(snapshot), L"Failed to capture Find snapshot before combined-view-state validation.");
    state.Require(snapshot.usesDxUiHost && snapshot.visibleChildWindowCount <= 1u && snapshot.resultColumnIds.size() >= 2u &&
                      snapshot.resultColumnWidthsDip.size() >= 2u && snapshot.firstResultHeaderRect.right > snapshot.firstResultHeaderRect.left &&
                      snapshot.secondResultHeaderRect.right > snapshot.secondResultHeaderRect.left && snapshot.resultColumnIds[0] == L"name" &&
                      snapshot.resultColumnIds[1] == L"path" && snapshot.selectedResultCount == 1u &&
                      std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), selectedPath) != snapshot.fullPaths.end() &&
                      snapshot.dxResizeFailureCount == 0u,
                  std::format(L"Find window did not expose the baseline results-grid state before combined-view-state validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Find combined-view-state validation did not reorder the visible Path column ahead of Name through the deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find combined-view-state validation did not widen the visible reordered Path column through the deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true), L"Failed to enable descending logical Name sort for combined-view-state validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.visibleChildWindowCount <= 1u && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not stay correct before combined-view-state persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->resultsGridLayout.size() >= 2u, L"Closing the Find window should persist the combined visible grid layout.");
    if (g_settings.search->resultsGridLayout.size() >= 2u)
    {
        state.Require(g_settings.search->resultsGridLayout[0].columnId == L"path", L"Persisted combined Find layout should keep Path first after close.");
        state.Require(g_settings.search->resultsGridLayout[1].columnId == L"name", L"Persisted combined Find layout should keep Name second after close.");
        state.Require(std::fabs(g_settings.search->resultsGridLayout[0].widthDip - resizedFirstVisibleWidthDip) <= 1.0f,
                      L"Persisted combined Find layout should keep the widened first visible column width after close.");
    }
    state.Require(g_settings.search->sortColumnId == L"name", L"Persisted combined Find sort should remain bound to the logical Name column id.");
    state.Require(g_settings.search->sortDescending, L"Persisted combined Find sort should remain descending after close.");
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-combined-view-state-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for combined-view-state validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to restore Find options for combined-view-state validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to restart Find search for combined-view-state restore validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for combined-view-state restore validation.");

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered layout, resized width, and logical Name sort.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumns(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    if (SkipIfFindSelfTestClipboardUnavailable(state))
    {
        return true;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_copy";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state copy test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state copy subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state copy validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state copy validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state copy validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for restored combined-view-state copy validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state copy validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state copy validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state copy validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Find restored combined-view-state copy validation did not reorder the visible Path column ahead of Name through the deterministic debug "
                  L"layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find restored combined-view-state copy validation did not widen the visible reordered Path column through the deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state copy validation.");
    const auto formatExpectedDescendingOrder = [&]() -> std::wstring
    {
        std::wstring result;
        for (const std::wstring& path : expectedDescendingPaths)
        {
            if (! result.empty())
            {
                result.append(L",");
            }
            result.append(std::filesystem::path(path).filename().native());
        }
        return result;
    };
    const auto formatPersistedFindLayout = []() -> std::wstring
    {
        if (! g_settings.search.has_value())
        {
            return L"<none>";
        }

        std::wstring result;
        for (const Common::Settings::GridColumnLayoutEntry& entry : g_settings.search->resultsGridLayout)
        {
            if (! result.empty())
            {
                result.append(L",");
            }
            result.append(std::format(L"{}:{:.1f}", entry.columnId, entry.widthDip));
        }
        return result;
    };
    const bool prePersistCombinedStateSettled = WaitForFindSnapshot(
        [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
        SelfTest::Scale(3000ms),
        &snapshot);
    state.Require(prePersistCombinedStateSettled,
                  std::format(L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state copy persistence. "
                              L"expectedFirstWidth={:.1f} expectedOrder=[{}] live={} persistedSort='{}' persistedDescending={} persistedLayout=[{}]",
                              resizedFirstVisibleWidthDip,
                              formatExpectedDescendingOrder(),
                              DescribeFindSnapshotBrief(snapshot),
                              g_settings.search.has_value() ? g_settings.search->sortColumnId : std::wstring(L"<none>"),
                              g_settings.search.has_value() && g_settings.search->sortDescending ? 1 : 0,
                              formatPersistedFindLayout()));
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    state.Require(g_settings.search->resultsGridLayout.size() >= 2u,
                  L"Closing the Find window should persist the combined visible grid layout before restored combined copy validation.");
    if (g_settings.search->resultsGridLayout.size() >= 2u)
    {
        state.Require(g_settings.search->resultsGridLayout[0].columnId == L"path",
                      L"Persisted combined Find layout should keep Path first before restored combined copy validation.");
        state.Require(g_settings.search->resultsGridLayout[1].columnId == L"name",
                      L"Persisted combined Find layout should keep Name second before restored combined copy validation.");
        state.Require(std::fabs(g_settings.search->resultsGridLayout[0].widthDip - resizedFirstVisibleWidthDip) <= 1.0f,
                      L"Persisted combined Find layout should keep the widened first visible column width before restored combined copy validation.");
    }
    state.Require(g_settings.search->sortColumnId == L"name",
                  L"Persisted combined Find sort should remain bound to the logical Name column id before restored combined copy validation.");
    state.Require(g_settings.search->sortDescending, L"Persisted combined Find sort should remain descending before restored combined copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state copy validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state copy validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state copy validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state copy validation.");

    const std::wstring selectedPath         = (root / L"sub" / L"beta.txt").native();
    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before restored combined copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before restored combined copy validation.");
    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the reopened Find results grid before restored combined copy validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find results grid did not take focus before restored combined copy validation.");

    ClearClipboardContents(findWindow);
    SendFindResultCommand(findWindow, IDM_PANE_CLIPBOARD_COPY);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Find copy command should copy the restored combined visible row content to the clipboard after reopen.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Reopened Find clipboard copy should start with the restored visible Path column after combined view-state persistence.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Reopened Find clipboard copy should still include the selected result name after restored combined view-state persistence.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateSurvivesSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for restored combined-view-state rerun validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state rerun validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Find restored combined-view-state rerun validation did not reorder the visible Path column ahead of Name through the deterministic debug "
                  L"layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find restored combined-view-state rerun validation did not widen the visible reordered Path column through the deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state rerun persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state rerun validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before rerun validation.");

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state reopen.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state reopen.");

    const auto restoredCombinedRerunStateOk = [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.submittedNamePatternText == L"*.txt" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               containsPath(value, selectedPath);
    };
    state.Require(DebugGetFindFilesWindowSnapshot(reopened), L"Failed to capture Find snapshot after restored combined-view-state baseline rerun.");
    state.Require(
        restoredCombinedRerunStateOk(reopened),
        std::format(L"Reopened combined Find view-state should survive a baseline search rerun. expectedWidth={:.1f} observedWidth={:.1f} "
                    L"widthDelta={:.1f} submittedNameOk={} submittedNameLen={} containsSelected={} expectedOrder={} selectedPath='{}' Last snapshot: {}",
                    resizedFirstVisibleWidthDip,
                    reopened.resultColumnWidthsDip.empty() ? 0.0f : reopened.resultColumnWidthsDip[0],
                    reopened.resultColumnWidthsDip.empty() ? resizedFirstVisibleWidthDip
                                                           : std::fabs(reopened.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip),
                    reopened.submittedNamePatternText == L"*.txt" ? 1 : 0,
                    reopened.submittedNamePatternText.size(),
                    containsPath(reopened, selectedPath) ? 1 : 0,
                    matchesExpectedOrder(reopened) ? 1 : 0,
                    selectedPath,
                    DescribeFindSnapshotBrief(reopened)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateSurvivesSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state sort-cycles test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state sort-cycles subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state sort-cycles validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state sort-cycles validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedDescendingOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state sort-cycles validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state sort-cycles validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-sort-cycles-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state sort-cycles validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state sort-cycles validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state sort-cycles validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Failed to apply the deterministic Find header reorder before restored combined-view-state sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state sort-cycles validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state sort-cycles validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedDescendingOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state sort-cycles persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-sort-cycles-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state sort-cycles validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state sort-cycles validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state sort-cycles validation.");

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedDescendingOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    state.Require(DebugSetFindFilesWindowResultSort(0u, false),
                  L"Failed to cycle reopened combined Find sort ascending before restored combined copy-after-sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to restore reopened combined Find sort descending before restored combined copy-after-sort-cycles validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && matchesExpectedDescendingOrder(value) &&
               containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should survive repeated sort cycles.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumnsAfterSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    if (SkipIfFindSelfTestClipboardUnavailable(state))
    {
        return true;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_copy_after_sort_cycles";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state copy-after-sort-cycles test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state copy-after-sort-cycles subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state copy-after-sort-cycles validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state copy-after-sort-cycles validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-after-sort-cycles-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state copy-after-sort-cycles validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Find restored combined-view-state copy-after-sort-cycles validation did not reorder the visible Path column ahead of Name through the "
                  L"deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find restored combined-view-state copy-after-sort-cycles validation did not widen the visible reordered Path column through the "
                  L"deterministic debug layout path.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state copy-after-sort-cycles persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-after-sort-cycles-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state copy-after-sort-cycles validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state copy-after-sort-cycles validation.");

    const std::wstring selectedPath         = (root / L"sub" / L"beta.txt").native();
    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";

    FindFilesDebugSnapshot reopened{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not restore the combined reordered, resized, and sorted state before restored combined copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not expose the expected selected combined results-grid state before restored combined copy-after-sort-cycles validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        std::format(L"Reopened combined Find view-state should survive repeated sort cycles before restored combined copy-after-sort-cycles validation. "
                    L"expectedWidth>={:.1f}, rows={}/{}, cols={}/{}, cells={}/{}, selected={}, orderOk={}, containsSelected={}, {}",
                    resizedFirstVisibleWidthDip - 1.0f,
                    reopened.visibleResultRowCount,
                    baselineVisibleRowCount,
                    reopened.visibleResultColumnCount,
                    baselineVisibleColumnCount,
                    reopened.visibleResultCellCount,
                    baselineVisibleCellCount,
                    reopened.selectedResultCount,
                    matchesExpectedOrder(reopened) ? 1 : 0,
                    containsPath(reopened, selectedPath) ? 1 : 0,
                    DescribeFindSnapshotBrief(reopened)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the reopened Find results grid before restored combined copy-after-sort-cycles validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find results grid did not take focus before restored combined copy-after-sort-cycles validation.");

    ClearClipboardContents(findWindow);
    SendFindResultCommand(findWindow, IDM_PANE_CLIPBOARD_COPY);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(),
                  L"Find copy command should copy the restored combined visible row content to the clipboard after reopen and sort cycles.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Reopened Find clipboard copy should still start with the restored visible Path column after combined view-state sort churn.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Reopened Find clipboard copy should still include the selected result name after combined view-state sort churn.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumnsAfterSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }
    if (SkipIfFindSelfTestClipboardUnavailable(state))
    {
        return true;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_copy_after_search_rerun";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create restored combined-view-state copy-after-search-rerun test root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create restored combined-view-state copy-after-search-rerun subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"),
                  L"Failed to create alpha.txt for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"),
                  L"Failed to create beta.txt for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma needle\n"),
                  L"Failed to create gamma.txt for restored combined-view-state copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*.txt";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const auto containsPath = [](const FindFilesDebugSnapshot& snapshot, const std::wstring& path) noexcept
    { return std::find(snapshot.fullPaths.begin(), snapshot.fullPaths.end(), path) != snapshot.fullPaths.end(); };

    const std::vector<std::wstring> expectedDescendingPaths = {
        (root / L"gamma.txt").native(),
        (root / L"sub" / L"beta.txt").native(),
        (root / L"alpha.txt").native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state copy-after-search-rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state copy-after-search-rerun validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-after-search-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to configure Find options for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state copy-after-search-rerun validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path" && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose the baseline results-grid state before restored combined-view-state copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                  L"Failed to apply the deterministic Find header reorder before restored combined-view-state copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state "
                  L"copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state copy-after-search-rerun validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state copy-after-search-rerun persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-copy-after-search-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false),
                  L"Failed to restore Find options for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state copy-after-search-rerun validation.");

    const std::wstring selectedPath         = (root / L"sub" / L"beta.txt").native();
    const std::wstring expectedRelativePath = L"sub";
    const std::wstring expectedName         = L"beta.txt";

    FindFilesDebugSnapshot reopened{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not restore the combined reordered, resized, and sorted state before restored combined copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not expose the expected selected combined results-grid state before restored combined copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state reopen.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state reopen.");

    const auto restoredCombinedCopyRerunStateOk = [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && containsPath(value, selectedPath);
    };
    state.Require(DebugGetFindFilesWindowSnapshot(reopened),
                  L"Failed to capture Find snapshot after restored combined copy-after-search-rerun baseline rerun.");
    state.Require(restoredCombinedCopyRerunStateOk(reopened),
                  std::format(L"Reopened combined Find view-state should survive a baseline search rerun before restored combined copy-after-search-rerun "
                              L"validation. expectedWidth={:.1f} observedWidth={:.1f} widthDelta={:.1f} submittedNameOk={} submittedNameLen={} "
                              L"containsSelected={} expectedOrder={} "
                              L"selectedPath='{}' Last snapshot: {}",
                              resizedFirstVisibleWidthDip,
                              reopened.resultColumnWidthsDip.empty() ? 0.0f : reopened.resultColumnWidthsDip[0],
                              reopened.resultColumnWidthsDip.empty() ? resizedFirstVisibleWidthDip
                                                                     : std::fabs(reopened.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip),
                              reopened.submittedNamePatternText == L"*.txt" ? 1 : 0,
                              reopened.submittedNamePatternText.size(),
                              containsPath(reopened, selectedPath) ? 1 : 0,
                              matchesExpectedOrder(reopened) ? 1 : 0,
                              selectedPath,
                              DescribeFindSnapshotBrief(reopened)));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state copy-after-search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state copy-after-search-rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) && containsPath(value, selectedPath);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun before restored combined "
                  L"copy-after-search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget::ResultsGrid),
                  L"Failed to focus the reopened Find results grid before restored combined copy-after-search-rerun validation.");
    state.Require(WaitForFindSnapshot([](const FindFilesDebugSnapshot& value) noexcept { return value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid; },
                                      SelfTest::Scale(3000ms),
                                      &reopened),
                  L"Reopened Find results grid did not take focus before restored combined copy-after-search-rerun validation.");

    ClearClipboardContents(findWindow);
    SendFindResultCommand(findWindow, IDM_PANE_CLIPBOARD_COPY);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(findWindow);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(),
                  L"Find copy command should copy the restored combined visible row content to the clipboard after reopen and search rerun.");
    state.Require(copiedSelection.rfind((expectedRelativePath + L"\t"), 0u) == 0u,
                  L"Reopened Find clipboard copy should still start with the restored visible Path column after combined view-state search rerun.");
    state.Require(copiedSelection.find(expectedName) != std::wstring::npos,
                  L"Reopened Find clipboard copy should still include the selected result name after combined view-state search rerun.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateGridEnterActivatesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_enter_activation";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir), L"Failed to create alpha directory for restored combined-view-state Enter activation validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir), L"Failed to create beta directory for restored combined-view-state Enter activation validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir), L"Failed to create gamma directory for restored combined-view-state Enter activation validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state Enter activation validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state Enter activation validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state Enter activation validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state Enter activation validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-enter-activation-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state Enter activation validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state Enter activation validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state Enter activation validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state Enter activation validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state Enter activation validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state Enter activation validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state Enter "
                      L"activation validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state Enter activation validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state Enter activation validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state Enter activation persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-enter-activation-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for Enter activation validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for Enter activation validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state Enter activation validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state Enter activation validation.");

    const std::wstring selectedPath = betaDir.native();

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before Enter activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(findWindow, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(findWindow, WM_KEYUP, VK_RETURN, 0);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Pressing Enter on the restored combined Find result did not activate the selected directory.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateGridDoubleClickActivatesSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_double_click_activation";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state double-click activation validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir), L"Failed to create beta directory for restored combined-view-state double-click activation validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state double-click activation validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state double-click activation validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state double-click activation validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state double-click activation validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state double-click activation validation.");
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-double-click-activation-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state double-click activation validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state double-click activation validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state double-click activation validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state double-click activation validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state double-click activation validation. {}",
                    DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state double-click activation validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state double-click "
                      L"activation validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state double-click activation validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state double-click activation validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state double-click activation persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-double-click-activation-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for double-click activation validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for double-click activation validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state double-click activation validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state double-click activation validation.");

    const std::wstring selectedPath = betaDir.native();

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               value.selectedResultRowRect.right > value.selectedResultRowRect.left && value.selectedResultRowRect.bottom > value.selectedResultRowRect.top &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before double-click activation validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const LPARAM doubleClickPoint = DipPointToClientLParam(findWindow,
                                                           (reopened.selectedResultRowRect.left + reopened.selectedResultRowRect.right) * 0.5f,
                                                           (reopened.selectedResultRowRect.top + reopened.selectedResultRowRect.bottom) * 0.5f);
    SendMouseDoubleClickToResolvedPointWindow(findWindow, doubleClickPoint);

    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Double-clicking the restored combined Find result did not activate the selected directory.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_action_buttons";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir), L"Failed to create alpha directory for restored combined-view-state action-button validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir), L"Failed to create beta directory for restored combined-view-state action-button validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir), L"Failed to create gamma directory for restored combined-view-state action-button validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state action-button validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state action-button validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state action-button validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state action-button validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
                      std::format(L"Find window did not expose its live DX shell before restored combined-view-state action-button validation. {}",
                                  DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state action-button validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state action-button validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state action-button validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state action-button validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state action-button validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state action-button validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state action-button "
                      L"validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state action-button validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state action-button validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state action-button persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for action-button validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for action-button validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state action-button validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state action-button validation.");

    const std::wstring selectedPath     = betaDir.native();
    const std::wstring openButtonText   = LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN);
    const std::wstring parentButtonText = LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT);
    state.Require(! openButtonText.empty(), L"Failed to resolve Open button caption for restored combined-view-state action-button validation.");
    state.Require(! parentButtonText.empty(), L"Failed to resolve Go to folder button caption for restored combined-view-state action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before action-button validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByNameWithMessagePump(
                      findWindow, UIA_ButtonControlTypeId, openButtonText, L"restored combined Find Open action button"),
                  L"Open button did not expose live UIA InvokePattern interaction after restored combined state was reapplied.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Open button did not activate the selected directory after restored combined view state was reapplied.");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' before restored combined-view-state Parent action validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-enable Go to folder on the selected combined-state result.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByNameWithMessagePump(
                      findWindow, UIA_ButtonControlTypeId, parentButtonText, L"restored combined Find Go to folder action button"),
                  L"Go to folder button did not expose live UIA InvokePattern interaction after restored combined state was reapplied.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                  L"Go to folder button did not navigate back to the selected directory parent after restored combined view state was reapplied.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Go to folder button did not restore the parent directory contents after restored combined view state was reapplied.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelectionAfterSortCycles(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
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

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_action_buttons_after_sort_cycles";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state action-button sort-cycle validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir), L"Failed to create beta directory for restored combined-view-state action-button sort-cycle validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state action-button sort-cycle validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state action-button sort-cycle validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state action-button sort-cycle validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state action-button sort-cycle validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state action-button sort-cycle validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
                      std::format(L"Find window did not expose its live DX shell before restored combined-view-state action-button sort-cycle validation. {}",
                                  DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-after-sort-cycles-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state action-button sort-cycle validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state action-button sort-cycle validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state action-button sort-cycle validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state action-button sort-cycle validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state action-button sort-cycle validation. {}",
                    DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state action-button sort-cycle validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state action-button "
                      L"sort-cycle validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
                  L"Find deterministic restored combined-view-state action-button sort-cycle validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state action-button sort-cycle validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state action-button sort-cycle persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-after-sort-cycles-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for action-button sort-cycle validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for action-button sort-cycle validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state action-button sort-cycle validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state action-button sort-cycle validation.");

    const std::wstring selectedPath     = betaDir.native();
    const std::wstring openButtonText   = LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN);
    const std::wstring parentButtonText = LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT);
    state.Require(! openButtonText.empty(), L"Failed to resolve Open button caption for restored combined-view-state action-button sort-cycle validation.");
    state.Require(! parentButtonText.empty(),
                  L"Failed to resolve Go to folder button caption for restored combined-view-state action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before action-button sort-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled &&
               value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should preserve Open and Go to folder enablement across repeated sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByNameWithMessagePump(
                      findWindow, UIA_ButtonControlTypeId, openButtonText, L"restored combined Find Open action button after sort cycles"),
                  L"Open button did not expose live UIA InvokePattern interaction after restored combined state and reopened sort cycles.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Open button did not activate the selected directory after restored combined view state and reopened sort cycles.");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' before restored combined-view-state Parent action sort-cycle validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-enable Go to folder on the selected combined-state result after sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByNameWithMessagePump(
                      findWindow, UIA_ButtonControlTypeId, parentButtonText, L"restored combined Find Go to folder action button after sort cycles"),
                  L"Go to folder button did not expose live UIA InvokePattern interaction after restored combined state and reopened sort cycles.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
                  L"Go to folder button did not navigate back to the selected directory parent after restored combined view state and reopened sort cycles.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Go to folder button did not restore the parent directory contents after restored combined view state and reopened sort cycles.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelectionAfterSortCyclesAndSearchRerun(HWND mainWindow,
                                                                                                                      CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.failure = L"Main window unavailable for restored combined-view-state action-button sort-cycle search-rerun validation.";
        return false;
    }

    const auto closeFindWindow = []() noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_action_buttons_after_sort_cycles_and_search_rerun";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir),
                  L"Failed to create beta directory for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state action-button sort-cycle search-rerun validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state action-button sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state action-button sort-cycle search-rerun validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(
            WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
            std::format(
                L"Find window did not expose its live DX shell before restored combined-view-state action-button sort-cycle search-rerun validation. {}",
                DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-after-sort-cycles-and-search-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state action-button sort-cycle search-rerun validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state action-button sort-cycle "
                              L"search-rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(
            ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
            L"Failed to apply the deterministic Find header reorder before restored combined-view-state action-button sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state action-button "
                      L"sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find deterministic restored combined-view-state action-button sort-cycle search-rerun validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state action-button sort-cycle "
                  L"search-rerun persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-action-buttons-after-sort-cycles-and-search-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for action-button sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for action-button sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state action-button sort-cycle search-rerun validation.");

    const std::wstring selectedPath     = betaDir.native();
    const std::wstring openButtonText   = LoadStringResource(nullptr, IDS_FIND_ACTION_OPEN);
    const std::wstring parentButtonText = LoadStringResource(nullptr, IDS_FIND_ACTION_PARENT);
    state.Require(! openButtonText.empty(),
                  L"Failed to resolve Open button caption for restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(! parentButtonText.empty(),
                  L"Failed to resolve Go to folder button caption for restored combined-view-state action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesDebugSnapshot reopened{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not restore the combined reordered, resized, and sorted state before action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not expose the expected selected combined results-grid state before action-button sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));

    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled &&
               value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f && value.visibleResultRowCount == baselineVisibleRowCount &&
               value.visibleResultColumnCount == baselineVisibleColumnCount && value.visibleResultCellCount == baselineVisibleCellCount &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should preserve Open and Go to folder enablement across repeated sort cycles before the search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"beta*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to narrow Find search after restored combined-view-state sort cycles.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun narrowed Find search after restored combined-view-state sort cycles.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle after restored combined-view-state sort cycles.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               value.fullPaths.size() == 1u && value.fullPaths[0] == selectedPath;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should survive a narrowed search rerun after sort cycles while preserving the visible layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state action-button sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state action-button sort-cycle search-rerun validation.");

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value) && std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun after sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state sort-cycle search rerun.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-expose the expected selected combined results-grid state after sort cycles and search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByNameWithMessagePump(
                      findWindow, UIA_ButtonControlTypeId, openButtonText, L"restored combined Find Open action button after sort cycles and search rerun"),
                  L"Open button did not expose live UIA InvokePattern interaction after restored combined state, sort cycles, and search rerun.");
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Open button did not activate the selected directory after restored combined view state, sort cycles, and search rerun.");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' before restored combined-view-state Parent action sort-cycle search-rerun validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.openButtonEnabled && value.parentButtonEnabled && value.resultColumnIds.size() >= 2u &&
               value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-enable Go to folder on the selected combined-state result after sort cycles and search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByNameWithMessagePump(findWindow,
                                                              UIA_ButtonControlTypeId,
                                                              parentButtonText,
                                                              L"restored combined Find Go to folder action button after sort cycles and search rerun"),
                  L"Go to folder button did not expose live UIA InvokePattern interaction after restored combined state, sort cycles, and search rerun.");
    state.Require(
        WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(5000ms)),
        L"Go to folder button did not navigate back to the selected directory parent after restored combined view state, sort cycles, and search rerun.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Go to folder button did not restore the parent directory contents after restored combined view state, sort cycles, and search rerun.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateGridEnterActivatesSelectionAfterSortCyclesAndSearchRerun(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.failure = L"Main window unavailable for restored combined-view-state Enter sort-cycle search-rerun validation.";
        return false;
    }

    const auto closeFindWindow = []() noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_enter_after_sort_cycles_and_search_rerun";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir),
                  L"Failed to create beta directory for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state Enter sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state Enter sort-cycle search-rerun validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state Enter sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state Enter sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state Enter sort-cycle search-rerun validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(
            WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
            std::format(L"Find window did not expose its live DX shell before restored combined-view-state Enter sort-cycle search-rerun validation. {}",
                        DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-enter-after-sort-cycles-and-search-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state Enter sort-cycle search-rerun validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        std::format(
            L"Find window did not expose the baseline results-grid state before restored combined-view-state Enter sort-cycle search-rerun validation. {}",
            DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
                      L"Failed to apply the deterministic Find header reorder before restored combined-view-state Enter sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state Enter "
                      L"sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find deterministic restored combined-view-state Enter sort-cycle search-rerun validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &snapshot),
        L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state Enter sort-cycle search-rerun persistence.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-enter-after-sort-cycles-and-search-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for Enter sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for Enter sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state Enter sort-cycle search-rerun validation.");

    const std::wstring selectedPath = betaDir.native();
    FindFilesDebugSnapshot reopened{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not restore the combined reordered, resized, and sorted state before Enter sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not expose the expected selected combined results-grid state before Enter sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should preserve selected Enter activation state across repeated sort cycles before the search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"beta*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to narrow Find search after restored combined-view-state Enter sort cycles.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun narrowed Find search after restored combined-view-state Enter sort cycles.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle after restored combined-view-state Enter sort cycles.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               value.fullPaths.size() == 1u && value.fullPaths[0] == selectedPath;
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should survive a narrowed search rerun after Enter sort cycles while preserving the visible layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state Enter sort-cycle search-rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value) && std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun after Enter sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state Enter sort-cycle search rerun.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-expose the expected selected combined results-grid state after Enter sort cycles and search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND enterTarget = GetFocus();
    if (! enterTarget || (enterTarget != findWindow && IsChild(findWindow, enterTarget) == FALSE))
    {
        enterTarget = findWindow;
    }

    SendMessageW(enterTarget, WM_KEYDOWN, VK_RETURN, 0);
    SendMessageW(enterTarget, WM_KEYUP, VK_RETURN, 0);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Pressing Enter on the restored combined Find result did not activate the selected directory after sort cycles and search rerun.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateGridDoubleClickActivatesSelectionAfterSortCyclesAndSearchRerun(HWND mainWindow,
                                                                                                                         CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.failure = L"Main window unavailable for restored combined-view-state double-click sort-cycle search-rerun validation.";
        return false;
    }

    const auto closeFindWindow = []() noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const std::optional<Common::Settings::SearchDialogSettings> previousSearch = g_settings.search;
    const auto cleanup                                                         = wil::scope_exit([&] noexcept
    {
        closeFindWindow();
        g_settings.search = previousSearch;
    });
    const auto traceStep                                                       = [](std::wstring_view step) noexcept
    {
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                                   std::format(L"find_dialog_restored_combined_view_state_grid_doubleClick_after_sort_cycles_and_search_rerun {}", step));
    };

    closeFindWindow();
    traceStep(L"begin");

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root     = suiteRoot / L"work" / L"find_dialog_restored_combined_view_state_double_click_after_sort_cycles_and_search_rerun";
    const std::filesystem::path alphaDir = root / L"alpha";
    const std::filesystem::path betaDir  = root / L"beta";
    const std::filesystem::path gammaDir = root / L"gamma";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();
    state.Require(SelfTest::EnsureDirectory(alphaDir),
                  L"Failed to create alpha directory for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(betaDir),
                  L"Failed to create beta directory for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(SelfTest::EnsureDirectory(gammaDir),
                  L"Failed to create gamma directory for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(SelfTest::WriteTextFile(betaDir / L"inside.txt", "payload"),
                  L"Failed to create beta/inside.txt for restored combined-view-state double-click sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for restored combined-view-state double-click sort-cycle search-rerun validation.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha", L"beta", L"gamma"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for restored combined-view-state double-click sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    Common::Settings::SearchDialogSettings search = g_settings.search.value_or(Common::Settings::SearchDialogSettings{});
    search.lastRoot                               = root.native();
    search.lastNamePattern                        = L"*";
    search.lastContentPattern.clear();
    search.nameMode    = Common::Settings::SearchNameMode::Wildcard;
    search.contentMode = Common::Settings::SearchContentMode::Disabled;
    search.sortColumnId.clear();
    search.sortDescending = false;
    search.resultsGridLayout.clear();
    g_settings.search = search;

    const std::vector<std::wstring> expectedDescendingPaths = {
        gammaDir.native(),
        betaDir.native(),
        alphaDir.native(),
    };

    const auto matchesExpectedOrder = [&](const FindFilesDebugSnapshot& snapshot) noexcept
    {
        return snapshot.resultCount == expectedDescendingPaths.size() && snapshot.fullPaths.size() >= expectedDescendingPaths.size() &&
               std::equal(expectedDescendingPaths.begin(), expectedDescendingPaths.end(), snapshot.fullPaths.begin());
    };

    const auto openFindWindow = [&](const wchar_t* themeTag) noexcept -> HWND
    {
        FindFilesPaneContext context{};
        context.rootPluginPath = root;
        const AppTheme theme   = ResolveAppTheme(ThemeMode::Dark, themeTag);
        state.Require(ShowFindFilesWindow(mainWindow, g_settings, theme, std::move(context)),
                      L"Failed to open Find window for restored combined-view-state double-click sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return nullptr;
        }

        const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE,
                      L"Find window did not open for restored combined-view-state double-click sort-cycle search-rerun validation.");
        FindFilesDebugSnapshot readySnapshot{};
        state.Require(
            WaitForFindWindowReady(SelfTest::Scale(3000ms), &readySnapshot),
            std::format(L"Find window did not expose its live DX shell before restored combined-view-state double-click sort-cycle search-rerun validation. {}",
                        DescribeFindSnapshotBrief(readySnapshot)));
        return findWindow;
    };

    HWND findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-double-click-after-sort-cycles-and-search-rerun-save");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to configure Find options for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to start Find search for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Find search did not become idle for restored combined-view-state double-click sort-cycle search-rerun validation.");
    traceStep(L"initial_search_idle");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.secondResultHeaderRect.right > value.secondResultHeaderRect.left &&
               ((value.resultColumnIds[0] == L"name" && value.resultColumnIds[1] == L"path") ||
                (value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name")) &&
               value.dxResizeFailureCount == 0u;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  std::format(L"Find window did not expose the baseline results-grid state before restored combined-view-state double-click sort-cycle "
                              L"search-rerun validation. {}",
                              DescribeFindSnapshotBrief(snapshot)));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.resultColumnIds[0] == L"name")
    {
        state.Require(
            ApplyFindVisibleHeaderReorderViaDebug(snapshot, SelfTest::Scale(3000ms)),
            L"Failed to apply the deterministic Find header reorder before restored combined-view-state double-click sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
                   value.resultColumnWidthsDip.size() >= 2u && value.dxResizeFailureCount == 0u;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      L"Find deterministic header reorder did not move the visible Path column ahead of Name before restored combined-view-state double-click "
                      L"sort-cycle search-rerun validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(
        ApplyFindFirstVisibleColumnResizeViaDebug(snapshot, 48.0f, SelfTest::Scale(3000ms)),
        L"Find deterministic restored combined-view-state double-click sort-cycle search-rerun validation did not widen the reordered first visible column.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidthDip = snapshot.resultColumnWidthsDip[0];

    state.Require(DebugSetFindFilesWindowResultSort(0u, true),
                  L"Failed to enable descending logical Name sort for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               value.resultColumnWidthsDip.size() >= 2u && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find logical Name sort or widened reordered layout did not settle before restored combined-view-state double-click sort-cycle search-rerun "
                  L"persistence.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"saved_layout_sorted");

    closeFindWindow();
    state.Require(g_settings.search.has_value(), L"Closing the Find window should persist search settings after combined reorder, resize, and sort changes.");
    if (! g_settings.search.has_value())
    {
        return false;
    }

    findWindow = openFindWindow(L"find-selftest-restored-combined-view-state-double-click-after-sort-cycles-and-search-rerun-restore");
    if (! findWindow || IsWindow(findWindow) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetFindFilesWindowOptions(true, false, true, true, false),
                  L"Failed to restore Find options after combined view-state reopen for double-click sort-cycle search-rerun validation.");
    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore Find query after combined view-state reopen for double-click sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to restart Find search for restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Reopened Find search did not become idle for restored combined-view-state double-click sort-cycle search-rerun validation.");
    traceStep(L"reopened_search_idle");

    const std::wstring selectedPath = betaDir.native();
    FindFilesDebugSnapshot reopened{};
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && matchesExpectedOrder(value);
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not restore the combined reordered, resized, and sorted state before double-click sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"reopened_selected");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state reopen.", selectedPath));
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               value.selectedResultRowRect.right > value.selectedResultRowRect.left && value.selectedResultRowRect.bottom > value.selectedResultRowRect.top &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened Find window did not expose the expected selected combined results-grid state before double-click sort-cycle search-rerun validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"sort_cycles_settled");

    const size_t baselineVisibleRowCount    = reopened.visibleResultRowCount;
    const size_t baselineVisibleColumnCount = reopened.visibleResultColumnCount;
    const size_t baselineVisibleCellCount   = reopened.visibleResultCellCount;
    const LONG sortClickX = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.left + reopened.secondResultHeaderRect.right) * 0.5f));
    const LONG sortClickY = static_cast<LONG>(std::lround((reopened.secondResultHeaderRect.top + reopened.secondResultHeaderRect.bottom) * 0.5f));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(sortClickX, sortClickY));
    SendMessageW(findWindow, WM_LBUTTONUP, 0, MAKELPARAM(sortClickX, sortClickY));

    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.selectedResultCount == 1u && value.visibleChildWindowCount <= 1u &&
               value.resultColumnIds.size() >= 2u && value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && value.resultColumnWidthsDip[0] >= resizedFirstVisibleWidthDip - 1.0f &&
               value.visibleResultRowCount == baselineVisibleRowCount && value.visibleResultColumnCount == baselineVisibleColumnCount &&
               value.visibleResultCellCount == baselineVisibleCellCount && value.dxResizeFailureCount == 0u && matchesExpectedOrder(value) &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened combined Find view-state should preserve selected double-click activation state across repeated sort cycles before the search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"beta*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to narrow Find search after restored combined-view-state double-click sort cycles.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun narrowed Find search after restored combined-view-state double-click sort cycles.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Narrowed Find search did not become idle after restored combined-view-state double-click sort cycles.");
    state.Require(
        WaitForFindSnapshot(
            [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == 1u && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               value.fullPaths.size() == 1u && value.fullPaths[0] == selectedPath;
    },
            SelfTest::Scale(3000ms),
            &reopened),
        L"Reopened combined Find view-state should survive a narrowed search rerun after double-click sort cycles while preserving the visible layout.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"narrowed_rerun_settled");

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to restore baseline Find search after restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find),
                  L"Failed to rerun baseline Find search after restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(DebugWaitForFindFilesWindowIdle(static_cast<uint32_t>(SelfTest::Scale(10000ms).count())),
                  L"Baseline Find search did not become idle after restored combined-view-state double-click sort-cycle search-rerun validation.");
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.resultCount == expectedDescendingPaths.size() && value.visibleChildWindowCount <= 1u && value.resultColumnIds.size() >= 2u &&
               value.resultColumnWidthsDip.size() >= 2u && value.resultColumnIds[0] == L"path" && value.resultColumnIds[1] == L"name" &&
               std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f && value.dxResizeFailureCount == 0u &&
               matchesExpectedOrder(value) && std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened combined Find view-state should restore after returning to the baseline search rerun after double-click sort cycles.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"baseline_rerun_settled");

    state.Require(DebugSelectFindFilesWindowResult(selectedPath),
                  std::format(L"Failed to reselect '{}' after restored combined-view-state double-click sort-cycle search rerun.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.resultColumnIds.size() >= 2u && value.resultColumnIds[0] == L"path" &&
               value.resultColumnIds[1] == L"name" && std::fabs(value.resultColumnWidthsDip[0] - resizedFirstVisibleWidthDip) <= 1.0f &&
               value.dxResizeFailureCount == 0u && value.focusTarget == FindFilesDebugFocusTarget::ResultsGrid &&
               value.selectedResultRowRect.right > value.selectedResultRowRect.left && value.selectedResultRowRect.bottom > value.selectedResultRowRect.top &&
               std::find(value.fullPaths.begin(), value.fullPaths.end(), selectedPath) != value.fullPaths.end();
    },
                      SelfTest::Scale(3000ms),
                      &reopened),
                  L"Reopened Find window did not re-expose the expected selected combined results-grid state after double-click sort cycles and search rerun.");
    if (! state.failure.empty())
    {
        return false;
    }
    traceStep(L"final_selection_ready");

    const LPARAM doubleClickPoint = DipPointToClientLParam(findWindow,
                                                           (reopened.selectedResultRowRect.left + reopened.selectedResultRowRect.right) * 0.5f,
                                                           (reopened.selectedResultRowRect.top + reopened.selectedResultRowRect.bottom) * 0.5f);
    traceStep(L"dispatch_double_click");
    SendMouseDoubleClickToResolvedPointWindow(findWindow, doubleClickPoint);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, betaDir, SelfTest::Scale(5000ms)),
                  L"Double-clicking the restored combined Find result did not activate the selected directory after sort cycles and search rerun.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogThemeCycleKeepsGridLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_theme_cycle";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create theme-cycle Find root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create theme-cycle Find subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha needle\n"), L"Failed to create alpha.txt for theme-cycle Find test.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta needle\n"), L"Failed to create beta.txt for theme-cycle Find test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.rootPluginPath      = root;
    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-theme-cycle-initial");
    state.Require(ShowFindFilesWindow(mainWindow, g_settings, initialTheme, std::move(context)), L"Failed to open Find window for theme-cycle validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for theme-cycle validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for theme-cycle validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for theme-cycle validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for theme-cycle validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot([&](const FindFilesDebugSnapshot& value) noexcept { return value.usesDxUiHost && value.resultCount == 2u; },
                                      SelfTest::Scale(10000ms),
                                      &snapshot),
                  L"Find window did not produce the expected results for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for theme-cycle validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.selectedResultRowFillArgb != 0u && value.selectedResultRowTextArgb != 0u &&
               value.resultsGridFolderViewMode;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose a selected row for theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring expectedLeaf                = std::filesystem::path(selectedPath).filename().native();
    const std::wstring expectedContainingFolder    = std::filesystem::path(selectedPath).parent_path().native();
    const uint32_t expectedFolderViewRainbowHash32 = FolderItemStableHash32(expectedContainingFolder, expectedLeaf);
    std::wstring baselineSelectedName;
    const auto requireSelectedUiaRowState = [&](std::wstring_view label) noexcept
    {
        const auto selectionState = CollectVisibleDescendantSelectionPatternState(findWindow, UIA_DataGridControlTypeId);
        state.Require(selectionState.has_value(), std::format(L"Failed to collect Find results-grid UIA selection state after {}.", label));
        if (! selectionState.has_value())
        {
            return false;
        }

        state.Require(selectionState->rootControlType == UIA_DataGridControlTypeId,
                      std::format(L"Find results-grid root should stay a DataGrid after {}.", label));
        state.Require(selectionState->hasSelectionPattern, std::format(L"Find results-grid should keep SelectionPattern after {}.", label));
        state.Require(selectionState->selectionCount == 1u,
                      std::format(L"Find results-grid should keep exactly one selected UIA row after {}; saw {}.", label, selectionState->selectionCount));
        state.Require(selectionState->selectedControlType == UIA_DataItemControlTypeId,
                      std::format(L"Find selected UIA row should stay a DataItem after {}.", label));
        state.Require(selectionState->selectedHasSelectionItemPattern, std::format(L"Find selected UIA row should keep SelectionItemPattern after {}.", label));
        state.Require(! selectionState->selectedName.empty(), std::format(L"Find selected UIA row should keep a non-empty accessible name after {}.", label));
        state.Require(selectionState->selectedName.find(expectedLeaf) != std::wstring::npos,
                      std::format(L"Find selected UIA row name '{}' should still include '{}' after {}.", selectionState->selectedName, expectedLeaf, label));
        if (! state.failure.empty())
        {
            return false;
        }

        if (baselineSelectedName.empty())
        {
            baselineSelectedName = selectionState->selectedName;
        }

        state.Require(selectionState->selectedName == baselineSelectedName,
                      std::format(L"Find selected UIA row accessible name should stay stable after {}; expected '{}', saw '{}'.",
                                  label,
                                  baselineSelectedName,
                                  selectionState->selectedName));
        return state.failure.empty();
    };

    state.Require(requireSelectedUiaRowState(L"the baseline theme-cycle selection capture"),
                  L"Find window baseline UIA selection state was not stable before theme churn.");
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
    const auto channelsNear = [](const D2D1_COLOR_F& actual, const D2D1_COLOR_F& expected) noexcept
    {
        constexpr float kTolerance = 1.5f / 255.0f;
        return std::fabs(actual.r - expected.r) <= kTolerance && std::fabs(actual.g - expected.g) <= kTolerance &&
               std::fabs(actual.b - expected.b) <= kTolerance;
    };

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        const uint64_t previousRenderCount = snapshot.dxRenderCount;
        UpdateFindFilesWindowsTheme(theme);
        state.Require(WaitForFindSnapshot(
                          [&](const FindFilesDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.visibleChildWindowCount <= 1u && value.resultCount == 2u && value.selectedResultCount == 1u &&
                   value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast && value.themeRainbow == theme.menu.rainbowMode &&
                   value.selectedResultRowFillArgb != 0u && value.selectedResultRowTextArgb != 0u && value.dxRenderCount >= previousRenderCount;
        },
                          SelfTest::Scale(3000ms),
                          &snapshot),
                      std::format(L"Find window did not settle after switching to {} theme.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(IsWindow(findWindow) != FALSE, std::format(L"Find window did not survive the {} theme update.", label));
        state.Require(snapshot.resultsGridFolderViewMode, std::format(L"Find results grid did not keep folder-view visual mode after {} theme update.", label));
        state.Require(snapshot.selectedResultRowUsesRainbow == expectRainbow,
                      std::format(L"Find selected-row rainbow state mismatch after {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast, std::format(L"Find high-contrast state mismatch after {} theme update.", label));
        state.Require(snapshot.selectedResultRowFillArgb != snapshot.selectedResultRowTextArgb,
                      std::format(L"Find selected-row colors collapsed to the same value after {} theme update.", label));

        const float minimumContrast = expectHighContrast ? 4.5f : 3.0f;
        state.Require(contrastRatio(snapshot.selectedResultRowFillArgb, snapshot.selectedResultRowTextArgb) >= minimumContrast,
                      std::format(L"Find selected-row text contrast dropped below {:.1f}:1 after {} theme update.", minimumContrast, label));
        if (expectRainbow)
        {
            const D2D1_COLOR_F actualFill   = unpackColor(snapshot.selectedResultRowFillArgb);
            const D2D1_COLOR_F expectedFill = ColorFromHSV(static_cast<float>(expectedFolderViewRainbowHash32 % 360u), 0.85f, theme.dark ? 0.75f : 0.90f, 1.0f);
            state.Require(channelsNear(actualFill, expectedFill),
                          std::format(L"Find selected-row Rainbow RGB does not match FolderView stable-hash color after {} theme update.", label));
        }
        state.Require(requireSelectedUiaRowState(std::format(L"the {} theme update", label)),
                      std::format(L"Find selected UIA row state did not remain stable after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"find-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"find-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"find-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"find-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestFindDialogCompactModeShrinksResultsGridMetrics(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeFindWindow = [&] noexcept
    {
        if (const HWND findWindow = GetFindFilesWindowHandle(); findWindow && IsWindow(findWindow))
        {
            PostMessageW(findWindow, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(findWindow, SelfTest::Scale(3000ms)));
        }
    };

    const auto cleanup = wil::scope_exit([&] noexcept { closeFindWindow(); });

    closeFindWindow();

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"find_dialog_compact_mode";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create compact-mode Find root.");
    state.Require(SelfTest::EnsureDirectory(root / L"sub"), L"Failed to create compact-mode Find subfolder.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha compact test\n"), L"Failed to create alpha.txt for compact-mode Find test.");
    state.Require(SelfTest::WriteTextFile(root / L"sub" / L"beta.txt", "beta compact test\n"), L"Failed to create beta.txt for compact-mode Find test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FindFilesPaneContext context{};
    context.rootPluginPath    = root;
    AppTheme standardTheme    = ResolveAppTheme(ThemeMode::Dark, L"find-selftest-compact-standard");
    standardTheme.compactMode = false;

    state.Require(ShowFindFilesWindow(mainWindow, g_settings, standardTheme, std::move(context)), L"Failed to open Find window for compact-mode validation.");

    const HWND findWindow = WaitForWindow([] noexcept { return GetFindFilesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(findWindow != nullptr && IsWindow(findWindow) != FALSE, L"Find window did not open for compact-mode validation.");
    if (! findWindow || IsWindow(findWindow) == FALSE)
    {
        return false;
    }

    state.Require(
        DebugConfigureFindFilesWindow(root.native(), L"*.txt", L"", Common::Settings::SearchNameMode::Wildcard, Common::Settings::SearchContentMode::Disabled),
        L"Failed to configure Find window for compact-mode validation.");
    state.Require(DebugSetFindFilesWindowOptions(true, true, false, true, false), L"Failed to configure Find options for compact-mode validation.");
    state.Require(DebugStartFindFilesWindowSearch(FindFilesDebugOperation::Find), L"Failed to start Find search for compact-mode validation.");

    FindFilesDebugSnapshot snapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.firstResultHeaderRect.right > value.firstResultHeaderRect.left &&
               value.firstResultHeaderRect.bottom > value.firstResultHeaderRect.top && value.statusStripVisible && value.statusStripHeightDip > 0.0f &&
               ! value.themeCompactMode;
    },
                      SelfTest::Scale(10000ms),
                      &snapshot),
                  L"Find window did not reach the standard-density baseline state for compact-mode validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPath = (root / L"sub" / L"beta.txt").native();
    state.Require(DebugSelectFindFilesWindowResult(selectedPath), std::format(L"Failed to select '{}' for compact-mode validation.", selectedPath));
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.selectedResultCount == 1u && value.selectedResultRowRect.right > value.selectedResultRowRect.left &&
               value.selectedResultRowRect.bottom > value.selectedResultRowRect.top && value.statusStripHeightDip > 0.0f && ! value.themeCompactMode;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not expose a selected row for compact-mode validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineHeaderHeight = std::max(0.0f, snapshot.firstResultHeaderRect.bottom - snapshot.firstResultHeaderRect.top);
    const float baselineRowHeight    = std::max(0.0f, snapshot.selectedResultRowRect.bottom - snapshot.selectedResultRowRect.top);
    const float baselineStatusHeight = snapshot.statusStripHeightDip;
    const bool sawNestedPathText =
        std::find(snapshot.resultPathTexts.begin(), snapshot.resultPathTexts.end(), std::wstring(L"sub")) != snapshot.resultPathTexts.end();
    state.Require(sawNestedPathText, L"Recursive Find results should show the containing subfolder in the Path column.");
    state.Require(
        baselineRowHeight <= 32.0f,
        std::format(L"Standard-density Find result rows without snippets should use shared compact grid metrics, but measured {:.2f} dip.", baselineRowHeight));
    if (! state.failure.empty())
    {
        return false;
    }

    AppTheme compactTheme              = standardTheme;
    compactTheme.compactMode           = true;
    const uint64_t baselineRenderCount = snapshot.dxRenderCount;
    UpdateFindFilesWindowsTheme(compactTheme);

    FindFilesDebugSnapshot compactSnapshot{};
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && value.themeCompactMode &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.firstResultHeaderRect.bottom > value.firstResultHeaderRect.top &&
               value.selectedResultRowRect.right > value.selectedResultRowRect.left && value.selectedResultRowRect.bottom > value.selectedResultRowRect.top &&
               value.statusStripHeightDip > 0.0f && value.dxRenderCount >= baselineRenderCount;
    },
                      SelfTest::Scale(3000ms),
                      &compactSnapshot),
                  L"Find window did not settle into compact density.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float compactHeaderHeight = std::max(0.0f, compactSnapshot.firstResultHeaderRect.bottom - compactSnapshot.firstResultHeaderRect.top);
    const float compactRowHeight    = std::max(0.0f, compactSnapshot.selectedResultRowRect.bottom - compactSnapshot.selectedResultRowRect.top);

    state.Require(compactHeaderHeight + 0.5f < baselineHeaderHeight,
                  std::format(L"Find compact header height should shrink below the standard height; baseline {:.2f} dip vs compact {:.2f} dip.",
                              baselineHeaderHeight,
                              compactHeaderHeight));
    state.Require(compactRowHeight + 0.5f < baselineRowHeight,
                  std::format(L"Find compact row height should shrink below the standard height; baseline {:.2f} dip vs compact {:.2f} dip.",
                              baselineRowHeight,
                              compactRowHeight));
    state.Require(
        compactRowHeight <= 26.0f,
        std::format(L"Compact-density Find result rows without snippets should stay near the shared grid compact row height, but measured {:.2f} dip.",
                    compactRowHeight));
    state.Require(compactSnapshot.statusStripHeightDip + 0.5f < baselineStatusHeight,
                  std::format(L"Find compact status strip height should shrink below the standard height; baseline {:.2f} dip vs compact {:.2f} dip.",
                              baselineStatusHeight,
                              compactSnapshot.statusStripHeightDip));
    if (! state.failure.empty())
    {
        return false;
    }

    UpdateFindFilesWindowsTheme(standardTheme);
    state.Require(WaitForFindSnapshot(
                      [&](const FindFilesDebugSnapshot& value) noexcept
    {
        return value.usesDxUiHost && value.resultCount == 2u && value.selectedResultCount == 1u && ! value.themeCompactMode &&
               value.firstResultHeaderRect.right > value.firstResultHeaderRect.left && value.selectedResultRowRect.right > value.selectedResultRowRect.left &&
               value.statusStripHeightDip > 0.0f;
    },
                      SelfTest::Scale(3000ms),
                      &snapshot),
                  L"Find window did not return to standard density after compact-mode validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float restoredHeaderHeight = std::max(0.0f, snapshot.firstResultHeaderRect.bottom - snapshot.firstResultHeaderRect.top);
    const float restoredRowHeight    = std::max(0.0f, snapshot.selectedResultRowRect.bottom - snapshot.selectedResultRowRect.top);
    state.Require(restoredHeaderHeight + 0.5f >= baselineHeaderHeight,
                  std::format(L"Find header height did not restore after leaving compact mode; baseline {:.2f} dip vs restored {:.2f} dip.",
                              baselineHeaderHeight,
                              restoredHeaderHeight));
    state.Require(restoredRowHeight + 0.5f >= baselineRowHeight,
                  std::format(L"Find row height did not restore after leaving compact mode; baseline {:.2f} dip vs restored {:.2f} dip.",
                              baselineRowHeight,
                              restoredRowHeight));
    state.Require(snapshot.statusStripHeightDip + 0.5f >= baselineStatusHeight,
                  std::format(L"Find status strip height did not restore after leaving compact mode; baseline {:.2f} dip vs restored {:.2f} dip.",
                              baselineStatusHeight,
                              snapshot.statusStripHeightDip));
    return state.failure.empty();
}

[[nodiscard]] bool QuickSearchSnapshotHasMatch(
    const FolderView::IncrementalSearchDebugSnapshot& snapshot, std::wstring_view displayName, UINT32 startPosition, UINT32 length, bool startsWith) noexcept
{
    return std::any_of(snapshot.matches.begin(), snapshot.matches.end(), [&](const FolderView::IncrementalSearchDebugMatch& match) noexcept {
        return match.displayName == displayName && match.range.startPosition == startPosition && match.range.length == length && match.startsWith == startsWith;
    });
}

[[nodiscard]] bool TestPaneQuickSearchIntegratedNavigation(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const DWORD processId = GetCurrentProcessId();
    CloseAllFindFilesWindowsForSearchTest();
    state.Require(PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"Quick Search focus validation"),
                  L"Failed to isolate the main window before Quick Search focus validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / (L"quick_search_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create quick-search test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"alpine.log", "alpine"), L"Failed to create alpine.log for quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"beta-alpha.txt", "beta"), L"Failed to create beta-alpha.txt for quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma"), L"Failed to create gamma.txt for quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"space name.txt", "space"), L"Failed to create space name.txt for quick-search test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::SortBy leftSortByBefore                 = g_folderWindow.GetSortBy(FolderWindow::Pane::Left);
    const FolderView::SortDirection leftSortDirectionBefore   = g_folderWindow.GetSortDirection(FolderWindow::Pane::Left);
    const bool leftExtensionsVisibleBefore                    = g_folderWindow.GetFileExtensionsVisible(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        g_folderWindow.SetSort(FolderWindow::Pane::Left, leftSortByBefore, leftSortDirectionBefore);
        g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Left, leftExtensionsVisibleBefore);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for quick-search test.");
    g_folderWindow.SetSort(FolderWindow::Pane::Left, FolderView::SortBy::Name, FolderView::SortDirection::Ascending);
    g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Left, true);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for quick-search test.");
    state.Require(
        WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"alpine.log", L"beta-alpha.txt", L"gamma.txt", L"space name.txt"}, SelfTest::Scale(3000ms)),
        L"Pane contents not ready for quick-search test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"gamma.txt"), L"Failed to focus gamma.txt before quick-search test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto baselineTopLevelWindows = SnapshotTopLevelWindowsForProcess(processId);
    const auto raiseMainWindowForInput = [&]() noexcept
    {
        const HWND rootWindow = GetAncestor(mainWindow, GA_ROOT);
        RaiseSelfTestWindowForInput(rootWindow && IsWindow(rootWindow) != FALSE ? rootWindow : mainWindow);
    };
    const auto clearSyntheticTextInputModifiers = []() noexcept
    {
        std::array<BYTE, 256> keyboardState{};
        if (GetKeyboardState(keyboardState.data()) == FALSE)
        {
            return;
        }

        for (const BYTE key : {static_cast<BYTE>(VK_SHIFT),
                               static_cast<BYTE>(VK_LSHIFT),
                               static_cast<BYTE>(VK_RSHIFT),
                               static_cast<BYTE>(VK_CONTROL),
                               static_cast<BYTE>(VK_LCONTROL),
                               static_cast<BYTE>(VK_RCONTROL),
                               static_cast<BYTE>(VK_MENU),
                               static_cast<BYTE>(VK_LMENU),
                               static_cast<BYTE>(VK_RMENU)})
        {
            keyboardState[key] = 0;
        }
        static_cast<void>(SetKeyboardState(keyboardState.data()));
    };
    raiseMainWindowForInput();
    FocusFolderViewPane(FolderWindow::Pane::Left);
    HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Left folder view handle unavailable for quick-search test.");
    if (! folderView)
    {
        return false;
    }

    const auto paneLabel        = [](FolderWindow::Pane pane) noexcept -> std::wstring_view { return pane == FolderWindow::Pane::Left ? L"Left" : L"Right"; };
    const auto focusDiagnostics = [&]()
    {
        return std::format(L"focus=0x{:X}, focusedFolderView=0x{:X}, focusedPane={}, expectedLeftFolderView=0x{:X}",
                           reinterpret_cast<uintptr_t>(GetFocus()),
                           reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd()),
                           paneLabel(g_folderWindow.GetFocusedPane()),
                           reinterpret_cast<uintptr_t>(folderView));
    };
    const auto ensureLeftFolderViewFocus = [&](std::wstring_view context) noexcept -> bool
    {
        raiseMainWindowForInput();
        if (const HWND rootWindow = GetAncestor(mainWindow, GA_ROOT); rootWindow && GetActiveWindow() != rootWindow)
        {
            SetActiveWindow(rootWindow);
        }

        g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
        FocusFolderViewPane(FolderWindow::Pane::Left);
        folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
        if (! folderView || IsWindow(folderView) == FALSE)
        {
            state.Require(false, std::format(L"Left folder view handle unavailable while restoring Quick Search focus during {}.", context));
            return false;
        }

        SetFocus(folderView);
        PumpPendingMessages();
        const bool focused = WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, folderView, SelfTest::Scale(1500ms));
        state.Require(focused, std::format(L"Left folder view did not have stable focus during {}; {}.", context, focusDiagnostics()));
        return focused;
    };
    const auto waitForLeftFolderViewFocusPassive = [&](std::chrono::milliseconds timeout) noexcept -> bool
    {
        using namespace std::chrono_literals;

        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
            if (folderView && IsWindow(folderView) != FALSE && GetFocus() == folderView && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
                g_folderWindow.GetFocusedPane() == FolderWindow::Pane::Left)
            {
                return true;
            }

            std::this_thread::sleep_for(10ms);
        }

        PumpPendingMessages();
        folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
        return folderView && IsWindow(folderView) != FALSE && GetFocus() == folderView && g_folderWindow.GetFocusedFolderViewHwnd() == folderView &&
               g_folderWindow.GetFocusedPane() == FolderWindow::Pane::Left;
    };
    const auto sendQuickSearchChar = [&](wchar_t character) noexcept
    {
        folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
        if (! folderView || IsWindow(folderView) == FALSE)
        {
            state.Require(false, L"Left folder view handle unavailable while typing Quick Search text.");
            return;
        }

        clearSyntheticTextInputModifiers();
        SendMessageW(folderView, WM_CHAR, static_cast<WPARAM>(character), 0);
        PumpPendingMessages();
    };
    const auto waitForQuickSearchSnapshot =
        [&](const auto& predicate, std::chrono::milliseconds timeout, FolderView::IncrementalSearchDebugSnapshot* outSnapshot) noexcept -> bool
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        FolderView::IncrementalSearchDebugSnapshot current{};
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, current) && predicate(current))
            {
                if (outSnapshot)
                {
                    *outSnapshot = std::move(current);
                }
                return true;
            }
            std::this_thread::sleep_for(10ms);
        }

        if (g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, current))
        {
            const bool matched = predicate(current);
            if (outSnapshot)
            {
                *outSnapshot = std::move(current);
            }
            return matched;
        }
        return false;
    };
    const auto appendQuickSearchChar = [&](wchar_t character, std::wstring_view expectedQuery, std::wstring_view context) noexcept -> bool
    {
        sendQuickSearchChar(character);
        if (! state.failure.empty())
        {
            return false;
        }

        FolderView::IncrementalSearchDebugSnapshot typed{};
        const bool advanced = waitForQuickSearchSnapshot([expectedQuery](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept {
            return value.active && value.query == expectedQuery;
        }, SelfTest::Scale(1500ms), &typed);
        state.Require(advanced,
                      std::format(L"Quick Search {} should advance to query '{}'; active={}, query='{}', focused='{}'; {}.",
                                  context,
                                  expectedQuery,
                                  typed.active ? 1 : 0,
                                  typed.query,
                                  typed.focusedDisplayName,
                                  focusDiagnostics()));
        return advanced;
    };
    const auto activateQuickSearchThroughCommand = [&](std::wstring_view context, FolderView::IncrementalSearchDebugSnapshot& outSnapshot) noexcept -> bool
    {
        for (int attempt = 1; attempt <= 3; ++attempt)
        {
            if (! ensureLeftFolderViewFocus(context))
            {
                return false;
            }

            clearSyntheticTextInputModifiers();
            static_cast<void>(SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_QUICK_SEARCH, 0), 0));
            PumpPendingMessages();

            const bool activated = waitForQuickSearchSnapshot([](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept {
                return value.active && value.query.empty();
            }, SelfTest::Scale(500ms), &outSnapshot);
            const bool focusRestored = activated && waitForLeftFolderViewFocusPassive(SelfTest::Scale(500ms));
            const bool stillEmpty = focusRestored && waitForQuickSearchSnapshot([](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept {
                return value.active && value.query.empty();
            }, SelfTest::Scale(250ms), &outSnapshot);
            if (stillEmpty)
            {
                return true;
            }

            SelfTest::AppendSelfTestTrace(std::format(L"{} activation attempt {} did not settle; activated={}, focusRestored={}, stillEmpty={}, "
                                                      L"active={}, query='{}', focused='{}'; {}.",
                                                      context,
                                                      attempt,
                                                      activated ? 1 : 0,
                                                      focusRestored ? 1 : 0,
                                                      stillEmpty ? 1 : 0,
                                                      outSnapshot.active ? 1 : 0,
                                                      outSnapshot.query,
                                                      outSnapshot.focusedDisplayName,
                                                      focusDiagnostics()));
        }

        state.Require(false,
                      std::format(L"{} should enter search mode; active={}, query='{}', focused='{}'; {}.",
                                  context,
                                  outSnapshot.active ? 1 : 0,
                                  outSnapshot.query,
                                  outSnapshot.focusedDisplayName,
                                  focusDiagnostics()));
        return false;
    };
    const auto quickSearchMatchDisplayOrder = [](const FolderView::IncrementalSearchDebugSnapshot& value)
    {
        std::vector<std::wstring> displayNames;
        displayNames.reserve(value.matches.size());
        for (const FolderView::IncrementalSearchDebugMatch& match : value.matches)
        {
            displayNames.push_back(match.displayName);
        }
        return displayNames;
    };
    const auto findNextQuickSearchMatchName = [](const std::vector<std::wstring>& matchDisplayOrder,
                                                 std::wstring_view currentDisplayName) -> std::optional<std::wstring>
    {
        if (matchDisplayOrder.empty())
        {
            return std::nullopt;
        }

        const auto current = std::find_if(matchDisplayOrder.begin(), matchDisplayOrder.end(), [&](const std::wstring& displayName) noexcept {
            return displayName == currentDisplayName;
        });
        if (current == matchDisplayOrder.end())
        {
            return std::nullopt;
        }

        auto next = current;
        ++next;
        if (next == matchDisplayOrder.end())
        {
            next = matchDisplayOrder.begin();
        }
        return *next;
    };
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, folderView, SelfTest::Scale(1000ms)),
                  std::format(L"Left folder view did not have stable focus before Quick Search activation; {}.", focusDiagnostics()));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_QUICK_SEARCH, 0), 0);
    PumpPendingMessages();

    FolderView::IncrementalSearchDebugSnapshot snapshot{};
    state.Require(waitForQuickSearchSnapshot([](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept
    { return value.active && value.query.empty(); },
                                             SelfTest::Scale(1500ms),
                                             &snapshot),
                  std::format(L"Quick Search command should activate integrated pane search with an empty query; active={}, query='{}', focused='{}'; {}.",
                              snapshot.active ? 1 : 0,
                              snapshot.query,
                              snapshot.focusedDisplayName,
                              focusDiagnostics()));
    state.Require(waitForLeftFolderViewFocusPassive(SelfTest::Scale(1000ms)),
                  std::format(L"Left folder view did not retain stable focus after Quick Search activation; {}.", focusDiagnostics()));
    if (! state.failure.empty())
    {
        return false;
    }

    if (! appendQuickSearchChar(L'a', L"a", L"initial typing") || ! appendQuickSearchChar(L'l', L"al", L"initial typing"))
    {
        return false;
    }

    state.Require(waitForQuickSearchSnapshot([](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept
    { return value.active && value.query == L"al"; },
                                             SelfTest::Scale(1500ms),
                                             &snapshot),
                  std::format(L"Quick Search query should be 'al'; active={}, query='{}', focused='{}'; {}.",
                              snapshot.active ? 1 : 0,
                              snapshot.query,
                              snapshot.focusedDisplayName,
                              focusDiagnostics()));
    const std::wstring initialQuickSearchFocus = snapshot.focusedDisplayName;
    const std::vector<std::wstring> alMatchDisplayOrder = quickSearchMatchDisplayOrder(snapshot);
    state.Require(initialQuickSearchFocus == L"alpha.txt" || initialQuickSearchFocus == L"alpine.log",
                  std::format(L"Quick Search should select one of the starts-with matches; got '{}'.", initialQuickSearchFocus));
    state.Require(snapshot.matches.size() == 3u, std::format(L"Quick Search should highlight three containing matches; got {}.", snapshot.matches.size()));
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"alpha.txt", 0u, 2u, true), L"Quick Search should highlight prefix match alpha.txt.");
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"alpine.log", 0u, 2u, true), L"Quick Search should highlight prefix match alpine.log.");
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"beta-alpha.txt", 5u, 2u, false), L"Quick Search should highlight contained match beta-alpha.txt.");
    const std::optional<std::wstring> expectedAfterFirstDown = findNextQuickSearchMatchName(alMatchDisplayOrder, initialQuickSearchFocus);
    state.Require(expectedAfterFirstDown.has_value(),
                  std::format(L"Quick Search initial focus '{}' was absent from the visible match order.", initialQuickSearchFocus));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(folderView, WM_KEYDOWN, VK_DOWN, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after first match navigation.");
    state.Require(snapshot.focusedDisplayName == expectedAfterFirstDown.value(),
                  std::format(L"Quick Search Down should move to the next visible match; expected '{}', got '{}'.",
                              expectedAfterFirstDown.value(),
                              snapshot.focusedDisplayName));
    const std::wstring focusAfterFirstDown = snapshot.focusedDisplayName;
    const std::optional<std::wstring> expectedAfterSecondDown = findNextQuickSearchMatchName(alMatchDisplayOrder, focusAfterFirstDown);
    state.Require(expectedAfterSecondDown.has_value(),
                  std::format(L"Quick Search first Down focus '{}' was absent from the visible match order.", focusAfterFirstDown));
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(folderView, WM_KEYDOWN, VK_DOWN, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after second match navigation.");
    state.Require(snapshot.focusedDisplayName == expectedAfterSecondDown.value(),
                  std::format(L"Quick Search Down should continue through visible containing matches; expected '{}', got '{}'.",
                              expectedAfterSecondDown.value(),
                              snapshot.focusedDisplayName));
    const std::wstring acceptedQuickSearchMatch = expectedAfterSecondDown.value();

    SendMessageW(folderView, WM_KEYDOWN, VK_RETURN, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after Enter.");
    state.Require(! snapshot.active, L"Quick Search Enter should accept the current item and exit search mode.");
    state.Require(snapshot.query.empty(), L"Quick Search Enter should clear the query.");
    state.Require(snapshot.focusedDisplayName == acceptedQuickSearchMatch,
                  std::format(L"Quick Search Enter should keep focus on the accepted item; expected '{}', got '{}'.",
                              acceptedQuickSearchMatch,
                              snapshot.focusedDisplayName));
    state.Require(g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left).value_or(std::filesystem::path{}) == root,
                  L"Quick Search Enter should not navigate away from the folder.");

    state.Require(WaitForNoNonBaselineWindows(processId, baselineTopLevelWindows, mainWindow, SelfTest::Scale(3000ms)),
                  L"Quick Search Enter activation left a transient top-level window open before no-match reactivation.");
    raiseMainWindowForInput();
    FocusFolderViewPane(FolderWindow::Pane::Left);
    folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Left folder view handle unavailable before Quick Search no-match reactivation.");
    if (! state.failure.empty())
    {
        return false;
    }
    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, folderView, SelfTest::Scale(1000ms)),
                  std::format(L"Left folder view did not have stable focus before Quick Search no-match reactivation; {}.", focusDiagnostics()));
    if (! state.failure.empty())
    {
        return false;
    }

    if (! activateQuickSearchThroughCommand(L"Quick Search no-match reactivation", snapshot))
    {
        return false;
    }
    FolderView::IncrementalSearchDebugSnapshot rightSnapshot{};
    const bool haveRightSnapshot = g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Right, rightSnapshot);
    state.Require(snapshot.active,
                  std::format(L"Quick Search no-match reactivation should enter search mode; active={}, query='{}', focused='{}'; "
                              L"rightSnapshot={}, rightActive={}, rightQuery='{}', rightFocused='{}'; {}.",
                              snapshot.active ? 1 : 0,
                              snapshot.query,
                              snapshot.focusedDisplayName,
                              haveRightSnapshot ? 1 : 0,
                              haveRightSnapshot && rightSnapshot.active ? 1 : 0,
                              haveRightSnapshot ? rightSnapshot.query : std::wstring{},
                              haveRightSnapshot ? rightSnapshot.focusedDisplayName : std::wstring{},
                              focusDiagnostics()));
    state.Require(snapshot.query.empty(), std::format(L"Quick Search no-match reactivation should clear the query; got '{}'.", snapshot.query));
    state.Require(waitForLeftFolderViewFocusPassive(SelfTest::Scale(1000ms)),
                  std::format(L"Left folder view did not retain stable focus after Quick Search no-match reactivation; {}.", focusDiagnostics()));
    if (! state.failure.empty())
    {
        return false;
    }

    if (! appendQuickSearchChar(L'z', L"z", L"no-match typing"))
    {
        return false;
    }
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot), L"Quick Search no-match snapshot should be available.");
    state.Require(snapshot.active,
                  std::format(L"Quick Search no-match state should remain active; active={}, query='{}', matches={}, focused='{}'; {}.",
                              snapshot.active ? 1 : 0,
                              snapshot.query,
                              snapshot.matches.size(),
                              snapshot.focusedDisplayName,
                              focusDiagnostics()));
    state.Require(snapshot.query == L"z", std::format(L"Quick Search no-match query should remain visible; got '{}'.", snapshot.query));
    state.Require(snapshot.matches.empty(), std::format(L"Quick Search no-match state should expose zero matches; got {}.", snapshot.matches.size()));

    SendMessageW(folderView, WM_KEYDOWN, VK_ESCAPE, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after Escape.");
    state.Require(! snapshot.active, L"Quick Search Escape should exit search mode.");
    state.Require(snapshot.query.empty(), L"Quick Search Escape should clear the query.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! activateQuickSearchThroughCommand(L"Quick Search shortcut-routed Space reactivation", snapshot))
    {
        return false;
    }
    state.Require(waitForLeftFolderViewFocusPassive(SelfTest::Scale(1000ms)),
                  std::format(L"Left folder view did not retain stable focus after Quick Search shortcut-routed Space reactivation; {}.", focusDiagnostics()));
    if (! state.failure.empty())
    {
        return false;
    }
    if (! appendQuickSearchChar(L's', L"s", L"shortcut-routed Space setup") || ! appendQuickSearchChar(L'p', L"sp", L"shortcut-routed Space setup") ||
        ! appendQuickSearchChar(L'a', L"spa", L"shortcut-routed Space setup") || ! appendQuickSearchChar(L'c', L"spac", L"shortcut-routed Space setup") ||
        ! appendQuickSearchChar(L'e', L"space", L"shortcut-routed Space setup"))
    {
        return false;
    }

    state.Require(waitForQuickSearchSnapshot([](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept
    { return value.active && value.query == L"space"; },
                                             SelfTest::Scale(1500ms),
                                             &snapshot),
                  std::format(L"Quick Search should be active with query 'space' before shortcut-routed Space; active={}, query='{}', focused='{}'; {}.",
                              snapshot.active ? 1 : 0,
                              snapshot.query,
                              snapshot.focusedDisplayName,
                              focusDiagnostics()));
    state.Require(waitForLeftFolderViewFocusPassive(SelfTest::Scale(1000ms)),
                  std::format(L"Left folder view did not retain stable focus before shortcut-routed Space; {}.", focusDiagnostics()));
    state.Require(
        waitForQuickSearchSnapshot([](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept { return value.active && value.query == L"space"; },
                                   SelfTest::Scale(500ms),
                                   &snapshot),
        std::format(L"Quick Search should remain active after stabilizing focus before shortcut-routed Space; active={}, query='{}', focused='{}'; {}.",
                    snapshot.active ? 1 : 0,
                    snapshot.query,
                    snapshot.focusedDisplayName,
                    focusDiagnostics()));
    if (! state.failure.empty())
    {
        return false;
    }

    const bool shortcutSpaceDispatched = DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/selectCalculateDirectorySizeNext");
    state.Require(shortcutSpaceDispatched, L"Shortcut-routed Space should dispatch while Quick Search is active.");
    PumpPendingMessages();
    state.Require(waitForQuickSearchSnapshot([](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept
    { return value.active && value.query == L"space "; },
                                             SelfTest::Scale(1500ms),
                                             &snapshot),
                  std::format(L"Quick Search should include shortcut-routed Space in the query; active={}, query='{}', focused='{}'; {}.",
                              snapshot.active ? 1 : 0,
                              snapshot.query,
                              snapshot.focusedDisplayName,
                              focusDiagnostics()));
    state.Require(snapshot.focusedDisplayName == L"space name.txt",
                  std::format(L"Quick Search should keep focus on the spaced filename; got '{}'.", snapshot.focusedDisplayName));
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"space name.txt", 0u, 6u, true),
                  L"Quick Search should expose a prefix match that includes shortcut-routed Space.");
    state.Require(waitForLeftFolderViewFocusPassive(SelfTest::Scale(1000ms)),
                  std::format(L"Left folder view did not retain stable focus after shortcut-routed Space; {}.", focusDiagnostics()));
    state.Require(
        waitForQuickSearchSnapshot([](const FolderView::IncrementalSearchDebugSnapshot& value) noexcept { return value.active && value.query == L"space "; },
                                   SelfTest::Scale(500ms),
                                   &snapshot),
        std::format(L"Quick Search should remain active after stabilizing focus after shortcut-routed Space; active={}, query='{}', focused='{}'; {}.",
                    snapshot.active ? 1 : 0,
                    snapshot.query,
                    snapshot.focusedDisplayName,
                    focusDiagnostics()));
    if (! state.failure.empty())
    {
        return false;
    }

    if (! appendQuickSearchChar(L'n', L"space n", L"shortcut-routed Space follow-up"))
    {
        return false;
    }
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after shortcut-routed Space follow-up.");
    state.Require(snapshot.focusedDisplayName == L"space name.txt",
                  std::format(L"Quick Search should keep focus on the spaced filename after the next character; got '{}'.", snapshot.focusedDisplayName));

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneQuickSearchUsesVisualNamesWhenExtensionsHidden(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"quick_search_hidden_extensions_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create hidden-extension quick-search test root.");
    state.Require(SelfTest::WriteTextFile(root / L"alpha.txt", "alpha"), L"Failed to create alpha.txt for hidden-extension quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"alpine.log", "alpine"), L"Failed to create alpine.log for hidden-extension quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"beta-alpha.txt", "beta"), L"Failed to create beta-alpha.txt for hidden-extension quick-search test.");
    state.Require(SelfTest::WriteTextFile(root / L"gamma.txt", "gamma"), L"Failed to create gamma.txt for hidden-extension quick-search test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const FolderView::SortBy leftSortByBefore                 = g_folderWindow.GetSortBy(FolderWindow::Pane::Left);
    const FolderView::SortDirection leftSortDirectionBefore   = g_folderWindow.GetSortDirection(FolderWindow::Pane::Left);
    const bool leftExtensionsVisibleBefore                    = g_folderWindow.GetFileExtensionsVisible(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        g_folderWindow.SetSort(FolderWindow::Pane::Left, leftSortByBefore, leftSortDirectionBefore);
        g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Left, leftExtensionsVisibleBefore);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for hidden-extension quick-search test.");
    g_folderWindow.SetSort(FolderWindow::Pane::Left, FolderView::SortBy::Name, FolderView::SortDirection::Ascending);
    g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Left, false);
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for hidden-extension quick-search test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"alpine.log", L"beta-alpha.txt", L"gamma.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for hidden-extension quick-search test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"gamma.txt"),
                  L"Failed to focus gamma.txt before hidden-extension quick-search test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    HWND folderView = g_folderWindow.GetFolderViewHwnd(FolderWindow::Pane::Left);
    state.Require(folderView != nullptr && IsWindow(folderView) != FALSE, L"Left folder view handle unavailable for hidden-extension quick-search test.");
    if (! folderView)
    {
        return false;
    }

    state.Require(WaitForFolderViewPaneFocus(FolderWindow::Pane::Left, folderView, SelfTest::Scale(1000ms)),
                  L"Left folder view did not have stable focus before hidden-extension Quick Search activation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_QUICK_SEARCH, 0), 0);
    PumpPendingMessages();
    SendMessageW(folderView, WM_CHAR, static_cast<WPARAM>(L'a'), 0);
    SendMessageW(folderView, WM_CHAR, static_cast<WPARAM>(L'l'), 0);
    PumpPendingMessages();

    FolderView::IncrementalSearchDebugSnapshot snapshot{};
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after typing with extensions hidden.");
    state.Require(snapshot.active, L"Quick Search should remain active after typing with extensions hidden.");
    state.Require(snapshot.query == L"al", std::format(L"Quick Search query should be 'al'; got '{}'.", snapshot.query));
    state.Require(snapshot.focusedDisplayName == L"alpha",
                  std::format(L"Quick Search snapshot should expose the visible focused label; got '{}'.", snapshot.focusedDisplayName));
    state.Require(snapshot.matches.size() == 3u,
                  std::format(L"Quick Search should expose three visible-name matches with extensions hidden; got {}.", snapshot.matches.size()));
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"alpha", 0u, 2u, true), L"Quick Search should expose visible prefix match alpha.");
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"alpine", 0u, 2u, true), L"Quick Search should expose visible prefix match alpine.");
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"beta-alpha", 5u, 2u, false), L"Quick Search should expose visible contained match beta-alpha.");

    SendMessageW(folderView, WM_KEYDOWN, VK_ESCAPE, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"gamma.txt"),
                  L"Failed to refocus gamma.txt before hidden-extension Quick Search navigation check.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_QUICK_SEARCH, 0), 0);
    PumpPendingMessages();
    SendMessageW(folderView, WM_CHAR, static_cast<WPARAM>(L't'), 0);
    PumpPendingMessages();

    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available for hidden-extension extension-text query.");
    state.Require(snapshot.active, L"Quick Search should remain active for hidden-extension extension-text query.");
    state.Require(snapshot.query == L"t", std::format(L"Quick Search query should be 't'; got '{}'.", snapshot.query));
    state.Require(snapshot.focusedDisplayName == L"beta-alpha",
                  std::format(L"Quick Search should ignore hidden extension text while selecting the visible match; got '{}'.", snapshot.focusedDisplayName));
    state.Require(snapshot.matches.size() == 1u,
                  std::format(L"Quick Search should expose only one visible-name match for 't' with extensions hidden; got {}.", snapshot.matches.size()));
    state.Require(QuickSearchSnapshotHasMatch(snapshot, L"beta-alpha", 2u, 1u, false),
                  L"Quick Search should expose the visible contained match beta-alpha for 't'.");

    SendMessageW(folderView, WM_KEYDOWN, VK_DOWN, 0);
    PumpPendingMessages();
    state.Require(g_folderWindow.DebugGetIncrementalSearchSnapshot(FolderWindow::Pane::Left, snapshot),
                  L"Quick Search snapshot should be available after hidden-extension match navigation.");
    state.Require(snapshot.focusedDisplayName == L"beta-alpha",
                  std::format(L"Quick Search Down should stay on the only visible-name match instead of matching hidden extensions; got '{}'.",
                              snapshot.focusedDisplayName));

    SendMessageW(folderView, WM_KEYDOWN, VK_ESCAPE, 0);
    PumpPendingMessages();
    return state.failure.empty();
}

} // namespace (tests)

void RunSearchCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"search_local_plugin_invalid_regex_reports_single_completion", [](CaseState& state) noexcept {
        return TestLocalPluginInvalidRegexReportsSingleCompletion(state);
    });
    SelfTest::RunCase(options, suite, L"search_local_plugin_parallel_cancel_fanin", [](CaseState& state) noexcept {
        return TestLocalPluginParallelSearchCancellationAndFanIn(state);
    });
    SelfTest::RunCase(options, suite, L"filesystem_local_watch_unwatch_drains_inflight_callback", [](CaseState& state) noexcept {
        return TestLocalPluginWatchUnwatchDrainsInflightCallback(state);
    });
    SelfTest::RunCase(options, suite, L"search_local_index_stream_stop_after_first", [](CaseState& state) noexcept {
        return TestLocalSearchIndexEnumerateStopsAfterFirstCandidate(state);
    });
    SelfTest::RunCase(options, suite, L"search_local_provider_reselect_after_index_stream_stays_responsive", [=](CaseState& state) noexcept {
        return TestLocalProviderReselectAfterIndexStreamStaysResponsive(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_find_dialog_search_ops", [=](CaseState& state) noexcept { return TestFindDialogSearchOps(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_active_close_does_not_join_ui", [=](CaseState& state) noexcept {
        return TestFindDialogCloseDuringActiveSearchDoesNotJoinUi(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_drops_stale_search_epoch_payloads", [=](CaseState& state) noexcept {
        return TestFindDialogDropsStaleSearchEpochPayloads(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_result_drains_respect_child_input_queue_order", [](CaseState& state) noexcept {
        return TestFindDialogResultDrainsRespectQueuedChildInput(state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_partial_completion_removes_only_known_sources", [](CaseState& state) noexcept {
        return TestFindDialogPartialCompletionRemovesOnlyKnownSources(state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_quickSearch_integrated_navigation", [=](CaseState& state) noexcept {
        return TestPaneQuickSearchIntegratedNavigation(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_quickSearch_uses_visual_names_when_extensions_hidden", [=](CaseState& state) noexcept {
        return TestPaneQuickSearchUsesVisualNamesWhenExtensionsHidden(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_local_root_overrides_stale_context", [=](CaseState& state) noexcept {
        return TestFindDialogTypedLocalRootOverridesStaleContext(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_opens_from_focused_pane_and_allows_multiple_instances", [=](CaseState& state) noexcept {
        return TestFindDialogOpensFromFocusedPaneAndAllowsMultipleInstances(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_recursive_local_search_and_index_availability", [=](CaseState& state) noexcept {
        return TestFindDialogRecursiveLocalSearchAndIndexAvailability(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_destination_navigation_stale_edit_host_hit_testing", [=](CaseState& state) noexcept {
        return TestFindDialogDestinationNavigationStaleEditHostHitTesting(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions", [=](CaseState& state) noexcept {
        return TestFindDialogResultShortcutsUseShellClipboardAndFileActions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_large_local_search_uses_incremental_updates", [=](CaseState& state) noexcept {
        return TestFindDialogLargeLocalSearchUsesIncrementalUpdates(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_running_status_shows_phase_and_path", [=](CaseState& state) noexcept {
        return TestFindDialogRunningStatusShowsPhaseAndPath(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_service_status_shows_backend_diagnostics", [=](CaseState& state) noexcept {
        return TestFindDialogServiceStatusShowsBackendDiagnostics(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_service_unavailable_warning_is_distinct", [=](CaseState& state) noexcept {
        return TestFindDialogServiceUnavailableWarningIsDistinct(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_failure_status_is_readable", [=](CaseState& state) noexcept {
        return TestFindDialogFailureShowsReadableStatus(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_uses_dxui_host_without_visible_child_controls", [=](CaseState& state) noexcept {
        return TestFindDialogUsesDxUiHostWithNoVisibleChildControls(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_exposes_live_uia_selection_and_inputs", [=](CaseState& state) noexcept {
        return TestFindDialogExposesLiveUiaSelectionAndInputs(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_long_run_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestFindDialogLongRunScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestFindDialogLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_directory_activation_navigates_into_selection", [=](CaseState& state) noexcept {
        return TestFindDialogDirectoryActivationNavigatesIntoSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_open_parent_keeps_directory_focused_in_parent", [=](CaseState& state) noexcept {
        return TestFindDialogOpenParentKeepsDirectoryFocusedInParent(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_enter_from_checkbox_invokes_default_search", [=](CaseState& state) noexcept {
        return TestFindDialogEnterFromCheckboxInvokesDefaultSearch(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_pointer_click_toggles_recursive_checkbox", [=](CaseState& state) noexcept {
        return TestFindDialogPointerClickTogglesRecursiveCheckbox(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_escape_closes_popup_before_cancel", [=](CaseState& state) noexcept {
        return TestFindDialogEscapeClosesPopupBeforeCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_escape_from_dx_control_closes_cancel", [=](CaseState& state) noexcept {
        return TestFindDialogEscapeFromDxControlClosesCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_access_keys_focus_expected_fields", [=](CaseState& state) noexcept {
        return TestFindDialogAccessKeysFocusExpectedFields(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_grid_enter_activates_selection", [=](CaseState& state) noexcept {
        return TestFindDialogGridFocusedEnterActivatesSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_grid_doubleClick_activates_selection", [=](CaseState& state) noexcept {
        return TestFindDialogGridDoubleClickActivatesSelection(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_tab_traversal_matches_expected_order", [=](CaseState& state) noexcept {
        return TestFindDialogTabTraversalMatchesExpectedOrder(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_editable_combo_keyboard_editing_keys", [=](CaseState& state) noexcept {
        return TestFindDialogEditableComboKeyboardEditingKeys(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_mode_typeahead_updates_selection_and_dependencies", [=](CaseState& state) noexcept {
        return TestFindDialogModeTypeaheadUpdatesSelectionAndDependencies(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_command_enablement_matches_idle_running_and_selection_states", [=](CaseState& state) noexcept {
        return TestFindDialogCommandEnablementMatchesIdleRunningAndSelectionStates(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_action_buttons_activate_expected_commands", [=](CaseState& state) noexcept {
        return TestFindDialogActionButtonsActivateExpectedCommands(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_persisted_grid_layout", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresPersistedGridLayout(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_header_drag_reorders_columns_without_sort", [=](CaseState& state) noexcept {
        return TestFindDialogHeaderDragReordersColumnsWithoutSort(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_copy_follows_reordered_columns", [=](CaseState& state) noexcept {
        return TestFindDialogCopyFollowsReorderedColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_reordered_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogReorderedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_header_resize_changes_visible_width", [=](CaseState& state) noexcept {
        return TestFindDialogHeaderResizeChangesVisibleWidth(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_reordered_columns_survive_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogReorderedColumnsSurviveSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_resized_columns_survive_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogResizedColumnsSurviveSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_reordered_resized_columns_survive_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogReorderedResizedColumnsSurviveSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_reordered_resized_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogReorderedResizedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_resized_columns_survive_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogResizedColumnsSurviveSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_header_click_sorts_results", [=](CaseState& state) noexcept {
        return TestFindDialogHeaderClickSortsResults(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_resized_grid_layout", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresResizedGridLayout(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_reordered_grid_layout", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresReorderedGridLayout(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_reordered_layout_copy_follows_visible_columns", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredReorderedLayoutCopyFollowsVisibleColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_persisted_sort_order", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresPersistedSortOrder(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_reordered_sorted_grid_layout", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresReorderedSortedGridLayout(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restores_combined_view_state", [=](CaseState& state) noexcept {
        return TestFindDialogRestoresCombinedViewState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumns(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_survives_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateSurvivesSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_survives_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateSurvivesSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns_after_sort_cycles", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumnsAfterSortCycles(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns_after_search_rerun", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateCopyFollowsVisibleColumnsAfterSearchRerun(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_grid_enter_activates_selection", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateGridEnterActivatesSelection(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection",
                      [=](CaseState& state) noexcept { return TestFindDialogRestoredCombinedViewStateGridDoubleClickActivatesSelection(mainWindow, state); });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_grid_enter_activates_selection_after_sort_cycles_and_search_rerun",
                      [=](CaseState& state) noexcept
    { return TestFindDialogRestoredCombinedViewStateGridEnterActivatesSelectionAfterSortCyclesAndSearchRerun(mainWindow, state); });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection_after_sort_cycles_and_search_rerun",
                      [=](CaseState& state) noexcept
    { return TestFindDialogRestoredCombinedViewStateGridDoubleClickActivatesSelectionAfterSortCyclesAndSearchRerun(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_restored_combined_view_state_action_buttons_activate_selection", [=](CaseState& state) noexcept {
        return TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelection(mainWindow, state);
    });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_action_buttons_activate_selection_after_sort_cycles",
                      [=](CaseState& state) noexcept
    { return TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelectionAfterSortCycles(mainWindow, state); });
    SelfTest::RunCase(options,
                      suite,
                      L"cmd_pane_find_dialog_restored_combined_view_state_action_buttons_activate_selection_after_sort_cycles_and_search_rerun",
                      [=](CaseState& state) noexcept
    { return TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelectionAfterSortCyclesAndSearchRerun(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_theme_cycle_keeps_grid_legible", [=](CaseState& state) noexcept {
        return TestFindDialogThemeCycleKeepsGridLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_find_dialog_compact_mode_shrinks_results_grid_metrics", [=](CaseState& state) noexcept {
        return TestFindDialogCompactModeShrinksResultsGridMetrics(mainWindow, state);
    });
}

namespace
{
