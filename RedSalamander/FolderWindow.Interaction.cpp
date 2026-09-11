#include "FolderWindowInternal.h"

#include "CommandDispatch.h"
#include "CommandRegistry.h"
#include "CommandRuntimeState.h"
#include "FluentIcons.h"

#include <windowsx.h>

LRESULT FolderWindow::OnSetCursor(HWND cursorWindow, UINT hitTest, UINT mouseMsg)
{
    if (! _hWnd)
    {
        return 0;
    }

    const LPARAM messagePos = static_cast<LPARAM>(GetMessagePos());
    POINT pt{GET_X_LPARAM(messagePos), GET_Y_LPARAM(messagePos)};
    if (ScreenToClient(_hWnd.get(), &pt) != FALSE)
    {
        if (OnSetCursor(pt))
        {
            return TRUE;
        }
    }
    return DefWindowProcW(
        _hWnd.get(), WM_SETCURSOR, reinterpret_cast<WPARAM>(cursorWindow), MAKELPARAM(static_cast<WORD>(hitTest), static_cast<WORD>(mouseMsg)));
}

void FolderWindow::OnSetFocus()
{
    FocusPanePreferredTarget(_activePane);
}

void FolderWindow::UpdatePaneFocusStates() noexcept
{
    const HWND focusedHwnd = GetFocus();
    const auto containsFocus = [focusedHwnd](HWND paneWindow) noexcept
    {
        return paneWindow && focusedHwnd && IsWindow(paneWindow) != FALSE &&
            (focusedHwnd == paneWindow || IsChild(paneWindow, focusedHwnd) != FALSE);
    };
    const auto inPane = [&](const PaneState& state) noexcept
    {
        return containsFocus(state.hFolderView.get()) || containsFocus(state.hNavigationView.get()) ||
            containsFocus(state.hFilterBar.get()) || containsFocus(state.hStatusBar.get()) ||
            containsFocus(state.hPreviewTabs.get()) || containsFocus(state.hPreviewContent.get()) ||
            containsFocus(state.terminalHwnd);
    };
    const bool inLeftPane = inPane(_leftPane);
    const bool inRightPane = inPane(_rightPane);
    if (focusedHwnd && (inLeftPane || inRightPane))
    {
        _lastFocusedPaneChild = focusedHwnd;
    }

    const Pane focusedPane = GetFocusedPane();
    SetActivePane(focusedPane);

    _leftPane.folderView.SetPaneFocused(focusedPane == Pane::Left);
    _rightPane.folderView.SetPaneFocused(focusedPane == Pane::Right);

    _leftPane.navigationView.SetPaneFocused(focusedPane == Pane::Left);
    _rightPane.navigationView.SetPaneFocused(focusedPane == Pane::Right);
}

HWND FolderWindow::GetPanePreferredFocusTarget(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.terminalTabSelected && state.terminalHwnd && IsWindow(state.terminalHwnd) != FALSE)
    {
        return state.terminalHwnd;
    }
    const bool inPane =
        (state.hFolderView && (_lastFocusedPaneChild == state.hFolderView.get() || IsChild(state.hFolderView.get(), _lastFocusedPaneChild))) ||
        (state.hNavigationView && (_lastFocusedPaneChild == state.hNavigationView.get() || IsChild(state.hNavigationView.get(), _lastFocusedPaneChild)));
    if (_lastFocusedPaneChild && IsWindow(_lastFocusedPaneChild) && inPane)
    {
        return _lastFocusedPaneChild;
    }

    return state.hFolderView.get();
}

void FolderWindow::FocusPaneFolderView(Pane pane) noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.hFolderView)
    {
        SetFocus(state.hFolderView.get());
    }
}

void FolderWindow::FocusPanePreferredTarget(Pane pane) noexcept
{
    const HWND target = GetPanePreferredFocusTarget(pane);
    if (target)
    {
        SetFocus(target);
    }
}

void FolderWindow::SetActivePane(Pane pane) noexcept
{
    if (_activePane == pane)
    {
        return;
    }

    _activePane = pane;

    if (_theme.menu.rainbowMode)
    {
        constexpr uint32_t kHueStepDegrees = 47u;
        _statusBarRainbowHueDegrees        = (_statusBarRainbowHueDegrees + kHueStepDegrees) % 360u;

        PaneState& state            = pane == Pane::Left ? _leftPane : _rightPane;
        state.statusFocusHueDegrees = _statusBarRainbowHueDegrees;
    }

    if (_leftPane.hStatusBar)
    {
        InvalidateRect(_leftPane.hStatusBar.get(), nullptr, FALSE);
    }
    if (_rightPane.hStatusBar)
    {
        InvalidateRect(_rightPane.hStatusBar.get(), nullptr, FALSE);
    }
}

FolderWindow::Pane FolderWindow::GetFocusedPane() const noexcept
{
    return GetPaneFromChild(GetFocus());
}

HWND FolderWindow::GetFocusedFolderViewHwnd() const noexcept
{
    const HWND focused = GetFocus();
    if (! focused)
    {
        return nullptr;
    }

    if (_leftPane.hFolderView && (focused == _leftPane.hFolderView.get() || IsChild(_leftPane.hFolderView.get(), focused)))
    {
        return _leftPane.hFolderView.get();
    }

    if (_rightPane.hFolderView && (focused == _rightPane.hFolderView.get() || IsChild(_rightPane.hFolderView.get(), focused)))
    {
        return _rightPane.hFolderView.get();
    }

    return nullptr;
}

bool FolderWindow::IsFocusInNavigationView() const noexcept
{
    const HWND focused = GetFocus();
    if (! focused)
    {
        return false;
    }

    return (_leftPane.hNavigationView && (focused == _leftPane.hNavigationView.get() || IsChild(_leftPane.hNavigationView.get(), focused))) ||
           (_rightPane.hNavigationView && (focused == _rightPane.hNavigationView.get() || IsChild(_rightPane.hNavigationView.get(), focused)));
}

bool FolderWindow::IsTerminalInputTarget(HWND targetWindow) const noexcept
{
    if (! targetWindow)
    {
        return false;
    }

    const auto containsTarget = [targetWindow](HWND terminalWindow) noexcept
    {
        return terminalWindow && IsWindow(terminalWindow) != FALSE &&
            (targetWindow == terminalWindow || IsChild(terminalWindow, targetWindow) != FALSE);
    };
    return containsTarget(_leftPane.terminalHwnd) || containsTarget(_rightPane.terminalHwnd);
}

bool FolderWindow::QueryTerminalHostCommandState(HWND invocationOrigin,
                                                 std::wstring_view commandId,
                                                 CommandRuntimeState& state) noexcept
{
    state = {};
    state.enabled = false;
    const auto contains = [invocationOrigin](HWND window) noexcept
    {
        return invocationOrigin != nullptr && window != nullptr && IsWindow(window) != FALSE &&
            (invocationOrigin == window || IsChild(window, invocationOrigin) != FALSE);
    };
    const std::optional<Pane> terminalPane = contains(_leftPane.terminalHwnd) ? std::optional{Pane::Left}
        : contains(_rightPane.terminalHwnd) ? std::optional{Pane::Right}
                                            : std::nullopt;
    if (! terminalPane.has_value())
    {
        // The application-level opener must work from a folder before any
        // terminal exists. Use the same active launch source as dispatch.
        if (commandId == L"cmd/terminal/openFloatingWindow")
        {
            state.enabled = GetActiveTerminalLaunchPath().has_value();
        }
        return true;
    }

    const Pane contextPane = terminalPane.value();
    PaneState& context = contextPane == Pane::Left ? _leftPane : _rightPane;
    TerminalViewState view{};
    view.sizeBytes = sizeof(view);
    if (context.terminal && SUCCEEDED(context.terminal->GetViewState(&view)))
    {
        wil::unique_cotaskmem_string title(view.title.data);
        wil::unique_cotaskmem_string status(view.status.data);
        state.terminalIdentityPresent = true;
        state.terminalInstanceId = view.instanceId;
        state.terminalSessionGeneration = view.sessionGeneration;

        const bool pluginAction = IsTerminalPluginActionId(commandId);
        if (pluginAction)
        {
            wil::com_ptr<ITerminalActions> actions;
            if (FAILED(context.terminal->QueryInterface(__uuidof(ITerminalActions), actions.put_void())) || ! actions)
            {
                return true;
            }
            TerminalActionRequest request{};
            request.sizeBytes = sizeof(request);
            request.commandId = {commandId.data(), static_cast<uint32_t>(commandId.size())};
            request.instanceId = view.instanceId;
            request.sessionGeneration = view.sessionGeneration;
            TerminalActionState actionState{};
            actionState.sizeBytes = sizeof(actionState);
            if (SUCCEEDED(actions->GetActionState(&request, &actionState)))
            {
                state.enabled = actionState.enabled != 0u;
                state.checked = actionState.checked != 0u;
            }
            return true;
        }
    }

    if (commandId == L"cmd/terminal/close" || commandId == L"cmd/terminal/contextMenu" ||
        commandId == L"cmd/terminal/sessionMenu")
    {
        state.enabled = context.terminal != nullptr && context.terminalOpen;
        return true;
    }
    if (commandId == L"cmd/terminal/tab/new")
    {
        state.enabled = ResolveTerminalCommandWorkingDirectory(contextPane).has_value();
        return true;
    }
    if (commandId == L"cmd/terminal/openFloatingWindow")
    {
        state.enabled = ResolveTerminalCommandWorkingDirectory(contextPane).has_value();
        return true;
    }
    if (commandId == L"cmd/terminal/tab/next" || commandId == L"cmd/terminal/tab/previous" ||
        commandId == L"cmd/terminal/tab/last")
    {
        state.enabled = true;
        return true;
    }
    constexpr std::wstring_view selectPrefix = L"cmd/terminal/tab/select/";
    if (commandId.starts_with(selectPrefix) && commandId.size() == selectPrefix.size() + 1u)
    {
        const bool previewAvailable = _previewSourcePane.has_value() && OppositePane(_previewSourcePane.value()) == contextPane;
        switch (commandId.back())
        {
            case L'1': state.enabled = true; break;
            case L'2': state.enabled = previewAvailable; break;
            case L'3': state.enabled = context.terminalOpen; break;
            default: state.enabled = false; break;
        }
        return true;
    }
    return true;
}

bool FolderWindow::ExecuteTerminalHostCommand(std::wstring_view commandId) noexcept
{
    const HWND focus = GetFocus();
    const auto contains = [focus](HWND window) noexcept
    {
        return focus && window && IsWindow(window) != FALSE && (focus == window || IsChild(window, focus) != FALSE);
    };
    const std::optional<Pane> terminalPane = contains(_leftPane.terminalHwnd) ? std::optional{Pane::Left}
        : contains(_rightPane.terminalHwnd) ? std::optional{Pane::Right}
                                            : std::nullopt;
    const Pane contextPane = terminalPane.value_or(_activePane);
    PaneState& context = contextPane == Pane::Left ? _leftPane : _rightPane;

    const auto executePluginAction = [&]() noexcept
    {
        if (! terminalPane.has_value() || ! context.terminal)
        {
            return false;
        }
        wil::com_ptr<ITerminalActions> actions;
        if (FAILED(context.terminal->QueryInterface(__uuidof(ITerminalActions), actions.put_void())) || ! actions)
        {
            return false;
        }
        TerminalViewState view{};
        view.sizeBytes = sizeof(view);
        if (FAILED(context.terminal->GetViewState(&view)))
        {
            return false;
        }
        CoTaskMemFree(view.title.data);
        CoTaskMemFree(view.status.data);
        TerminalActionRequest request{};
        request.sizeBytes = sizeof(request);
        request.commandId = {commandId.data(), static_cast<uint32_t>(commandId.size())};
        request.instanceId = view.instanceId;
        request.sessionGeneration = view.sessionGeneration;
        return actions->ExecuteAction(&request) == S_OK;
    };

    const auto showCommandMenu = [&](std::span<const std::wstring_view> actionIds) noexcept
    {
        const HWND owner = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        if (owner == nullptr || IsWindow(owner) == FALSE)
        {
            return false;
        }
        wil::unique_hmenu menu(CreatePopupMenu());
        if (! menu)
        {
            return false;
        }

        constexpr UINT kFirstTerminalMenuCommand = 0x7500u;
        for (size_t index = 0u; index < actionIds.size(); ++index)
        {
            const std::wstring_view actionId = actionIds[index];
            if (actionId.empty())
            {
                if (AppendMenuW(menu.get(), MF_SEPARATOR, 0u, nullptr) == FALSE)
                {
                    return false;
                }
                continue;
            }
            const CommandInfo* info = FindCommandInfo(actionId);
            std::wstring label = info != nullptr && info->displayNameStringId != 0u
                ? LoadStringResource(nullptr, info->displayNameStringId)
                : std::wstring(actionId);
            CommandRuntimeState actionState{};
            const bool enabled = QueryTerminalHostCommandState(focus, actionId, actionState) && actionState.enabled;
            const UINT flags = static_cast<UINT>(MF_STRING | (enabled ? MF_ENABLED : MF_GRAYED));
            if (AppendMenuW(menu.get(), flags, kFirstTerminalMenuCommand + static_cast<UINT>(index), label.c_str()) == FALSE)
            {
                return false;
            }
        }

        POINT anchor{};
        const HWND terminalWindow = terminalPane.has_value() ? context.terminalHwnd : nullptr;
        if (terminalWindow != nullptr && IsWindow(terminalWindow) != FALSE)
        {
            RECT bounds{};
            if (GetWindowRect(terminalWindow, &bounds) != FALSE)
            {
                anchor = {bounds.left + MulDiv(12, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI),
                          bounds.top + MulDiv(12, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI)};
            }
            POINT caret{};
            if (focus != nullptr && (focus == terminalWindow || IsChild(terminalWindow, focus) != FALSE) && GetCaretPos(&caret) != FALSE &&
                ClientToScreen(focus, &caret) != FALSE)
            {
                anchor = caret;
            }
        }
        if (anchor.x == 0 && anchor.y == 0)
        {
            RECT ownerBounds{};
            if (GetWindowRect(owner, &ownerBounds) != FALSE)
            {
                anchor = {ownerBounds.left + MulDiv(12, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI),
                          ownerBounds.top + MulDiv(12, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI)};
            }
        }

        SetForegroundWindow(owner);
        const UINT selected = static_cast<UINT>(TrackPopupMenuEx(
            menu.get(), TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, anchor.x, anchor.y, owner, nullptr));
        if (selected < kFirstTerminalMenuCommand || selected >= kFirstTerminalMenuCommand + actionIds.size())
        {
            return true;
        }
        const std::wstring_view selectedAction = actionIds[selected - kFirstTerminalMenuCommand];
        return ! selectedAction.empty() && ExecuteTerminalHostCommand(selectedAction);
    };

    if (commandId == L"cmd/terminal/contextMenu")
    {
        constexpr std::array<std::wstring_view, 7u> actions{{
            L"cmd/terminal/copy",
            L"cmd/terminal/paste",
            L"cmd/terminal/selectAll",
            std::wstring_view{},
            L"cmd/terminal/find",
            L"cmd/terminal/suggestions",
            L"cmd/terminal/close",
        }};
        return showCommandMenu(actions);
    }
    if (commandId == L"cmd/terminal/sessionMenu")
    {
        constexpr std::array<std::wstring_view, 4u> actions{{
            L"cmd/terminal/tab/new",
            L"cmd/terminal/openFloatingWindow",
            std::wstring_view{},
            L"cmd/terminal/close",
        }};
        return showCommandMenu(actions);
    }

    if (IsTerminalPluginActionId(commandId))
    {
        return executePluginAction();
    }
    if (commandId == L"cmd/terminal/close")
    {
        if (! terminalPane.has_value())
        {
            return false;
        }
        CloseTerminalPane(terminalPane.value());
        SetPaneContentTab(terminalPane.value(), 0u);
        FocusPaneFolderView(terminalPane.value());
        return true;
    }
    if (commandId == L"cmd/terminal/tab/new")
    {
        const std::optional<std::filesystem::path> workingDirectory = ResolveTerminalCommandWorkingDirectory(contextPane);
        const Pane targetPane = OppositePane(contextPane);
        return workingDirectory.has_value() && SUCCEEDED(OpenTerminalPane(targetPane, workingDirectory.value()));
    }
    if (commandId == L"cmd/terminal/openFloatingWindow")
    {
        const HWND owner = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        if (owner == nullptr || IsWindow(owner) == FALSE)
        {
            return false;
        }
        return DispatchApplicationCommand(owner,
                                          commandId,
                                          RedSalamander::Ui::CommandInvocationSource::Programmatic);
    }

    const auto selectContent = [&](size_t index) noexcept
    {
        const bool previewAvailable = _previewSourcePane.has_value() && OppositePane(_previewSourcePane.value()) == contextPane;
        if ((index == 1u && ! previewAvailable) || (index == 2u && ! context.terminalOpen) || index > 2u)
        {
            return false;
        }
        SetPaneContentTab(contextPane, index);
        SetActivePane(contextPane);
        FocusPanePreferredTarget(contextPane);
        return true;
    };
    if (commandId == L"cmd/terminal/tab/next" || commandId == L"cmd/terminal/tab/previous")
    {
        std::vector<size_t> available{0u};
        if (_previewSourcePane.has_value() && OppositePane(_previewSourcePane.value()) == contextPane)
        {
            available.push_back(1u);
        }
        if (context.terminalOpen)
        {
            available.push_back(2u);
        }
        const size_t current = context.terminalTabSelected ? 2u : context.previewTabSelected ? 1u : 0u;
        const auto currentIt = std::ranges::find(available, current);
        const size_t currentOffset = currentIt == available.end() ? 0u : static_cast<size_t>(currentIt - available.begin());
        const bool forward = commandId == L"cmd/terminal/tab/next";
        const size_t targetOffset = forward ? (currentOffset + 1u) % available.size()
                                            : (currentOffset + available.size() - 1u) % available.size();
        return selectContent(available[targetOffset]);
    }
    constexpr std::wstring_view selectPrefix = L"cmd/terminal/tab/select/";
    if (commandId.starts_with(selectPrefix) && commandId.size() == selectPrefix.size() + 1u)
    {
        const wchar_t digit = commandId.back();
        if (digit >= L'1' && digit <= L'3')
        {
            // Stable embedded identities are Folder=1, Preview=2, Terminal=3.
            return selectContent(static_cast<size_t>(digit - L'1'));
        }
        return false;
    }
    if (commandId == L"cmd/terminal/tab/last")
    {
        const bool previewAvailable = _previewSourcePane.has_value() && OppositePane(_previewSourcePane.value()) == contextPane;
        return selectContent(context.terminalOpen ? 2u : previewAvailable ? 1u : 0u);
    }
    if (commandId == L"cmd/pane/focus/left" || commandId == L"cmd/pane/focus/right" || commandId == L"cmd/pane/switchPaneFocus")
    {
        const Pane target = commandId == L"cmd/pane/focus/left" ? Pane::Left
            : commandId == L"cmd/pane/focus/right" ? Pane::Right
                                                    : OppositePane(contextPane);
        SetActivePane(target);
        FocusPanePreferredTarget(target);
        return true;
    }
    if (commandId == L"cmd/pane/resizeSplitter/left" || commandId == L"cmd/pane/resizeSplitter/right")
    {
        const float widthPx = static_cast<float>(std::max<LONG>(1, _clientSize.cx));
        const float stepPx = static_cast<float>(MulDiv(16, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI));
        const float direction = commandId == L"cmd/pane/resizeSplitter/left" ? -1.0f : 1.0f;
        SetSplitRatio(_splitRatio + (direction * stepPx / widthPx));
        return true;
    }
    return false;
}

void STDMETHODCALLTYPE FolderWindow::OnTerminalEvent(const TerminalEvent* event, void* cookie) noexcept
{
    if (cookie != this || event == nullptr || event->sizeBytes < sizeof(TerminalEvent) ||
        event->kind != TerminalEventKind::RootSessionExited || event->finalSnapshotComplete == 0u || ! _hWnd)
    {
        return;
    }

    const HWND appWindow = GetAncestor(_hWnd.get(), GA_ROOT);
    if (appWindow == nullptr || IsWindow(appWindow) == FALSE)
    {
        return;
    }
    auto payload = std::make_unique<TerminalEvent>(*event);
    static_cast<void>(PostMessagePayload(appWindow, WndMsg::kTerminalSessionExited, 0u, std::move(payload)));
}

void FolderWindow::HandleTerminalSessionExited(const TerminalEvent& event) noexcept
{
    if (event.sizeBytes < sizeof(TerminalEvent) || event.kind != TerminalEventKind::RootSessionExited ||
        event.exitCodePresent == 0u || event.finalSnapshotComplete == 0u)
    {
        return;
    }

    const auto closeMatchingPane = [&](Pane pane, PaneState& state) noexcept
    {
        if (! state.terminal)
        {
            return false;
        }
        TerminalViewState view{};
        view.sizeBytes = sizeof(view);
        const HRESULT stateHr = state.terminal->GetViewState(&view);
        const bool matches = SUCCEEDED(stateHr) &&
            memcmp(view.instanceId.bytes, event.instanceId.bytes, sizeof(event.instanceId.bytes)) == 0 &&
            view.sessionGeneration == event.sessionGeneration && view.activity.lifecycleState == TerminalLifecycleState::Exited &&
            view.exitCodePresent != 0u && view.exitCode == event.exitCode && view.finalSnapshotComplete != 0u;
        CoTaskMemFree(view.title.data);
        CoTaskMemFree(view.status.data);
        if (! matches)
        {
            return false;
        }

        CloseTerminalPane(pane);
        SetPaneContentTab(pane, 0u);
        SetActivePane(pane);
        FocusPaneFolderView(pane);
        return true;
    };

    if (! closeMatchingPane(Pane::Left, _leftPane))
    {
        static_cast<void>(closeMatchingPane(Pane::Right, _rightPane));
    }
}

HRESULT FolderWindow::RouteTerminalShortcut(HWND targetWindow,
                                            std::wstring_view commandId,
                                            const MSG& message,
                                            uint32_t normalizedModifiers,
                                            TerminalShortcutRoute& route) noexcept
{
    route = TerminalShortcutRoute::PassThrough;
    if (! targetWindow || commandId.empty() || commandId.size() > (std::numeric_limits<uint32_t>::max)())
    {
        return E_INVALIDARG;
    }

    const auto containsTarget = [targetWindow](HWND terminalWindow) noexcept
    {
        return terminalWindow && IsWindow(terminalWindow) != FALSE &&
            (targetWindow == terminalWindow || IsChild(terminalWindow, targetWindow) != FALSE);
    };
    PaneState* pane = containsTarget(_leftPane.terminalHwnd) ? &_leftPane
                                                            : containsTarget(_rightPane.terminalHwnd) ? &_rightPane : nullptr;
    if (pane == nullptr || ! pane->terminal)
    {
        return E_HANDLE;
    }

    wil::com_ptr<ITerminalActions> actions;
    HRESULT hr = pane->terminal->QueryInterface(__uuidof(ITerminalActions), actions.put_void());
    if (FAILED(hr) || ! actions)
    {
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    TerminalViewState view{};
    view.sizeBytes = sizeof(view);
    hr = pane->terminal->GetViewState(&view);
    const auto freeViewText = wil::scope_exit([&]() noexcept
    {
        CoTaskMemFree(view.title.data);
        CoTaskMemFree(view.status.data);
    });
    if (FAILED(hr))
    {
        return hr;
    }

    uint32_t modifierFlags = normalizedModifiers & 0x7u;
    const auto addModifier = [&](int virtualKey, uint32_t flag) noexcept
    {
        if ((GetKeyState(virtualKey) & 0x8000) != 0)
        {
            modifierFlags |= flag;
        }
    };
    addModifier(VK_LCONTROL, TerminalShortcutModifierLeftCtrl);
    addModifier(VK_RCONTROL, TerminalShortcutModifierRightCtrl);
    addModifier(VK_LMENU, TerminalShortcutModifierLeftAlt);
    addModifier(VK_RMENU, TerminalShortcutModifierRightAlt);
    addModifier(VK_LSHIFT, TerminalShortcutModifierLeftShift);
    addModifier(VK_RSHIFT, TerminalShortcutModifierRightShift);

    TerminalShortcutRequest request{};
    request.sizeBytes = sizeof(request);
    request.commandId = TerminalUtf16Span{.data = commandId.data(), .length = static_cast<uint32_t>(commandId.size())};
    request.message = message.message;
    request.virtualKey = static_cast<uint32_t>(message.wParam);
    request.scanCode = Common::Keyboard::ScanCodeFromKeyMessageLParam(message.lParam);
    request.extended = Common::Keyboard::IsExtendedKeyMessageLParam(message.lParam) ? 1u : 0u;
    request.systemKey = message.message == WM_SYSKEYDOWN ? 1u : 0u;
    request.repeatCount = static_cast<uint32_t>(static_cast<ULONG_PTR>(message.lParam) & 0xFFFFu);
    request.previousDown = (static_cast<ULONG_PTR>(message.lParam) & (1ull << 30u)) != 0u ? 1u : 0u;
    request.modifierFlags = modifierFlags;
    request.instanceId = view.instanceId;
    request.sessionGeneration = view.sessionGeneration;
    return actions->RouteShortcut(&request, &route);
}

bool FolderWindow::HandlePanePointerFocus(HWND targetWindow) noexcept
{
    if (! _hWnd || ! _settings || ! targetWindow || IsWindow(targetWindow) == FALSE || GetCapture() != nullptr)
    {
        return false;
    }

    const Common::Settings::MouseSettings mouse = _settings->mouse.value_or(Common::Settings::MouseSettings{});
    const bool terminalDisplayed = (_leftPane.terminalOpen && _leftPane.terminalTabSelected) ||
        (_rightPane.terminalOpen && _rightPane.terminalTabSelected);
    if (! Common::Settings::ShouldPaneFocusFollowPointer(mouse, terminalDisplayed))
    {
        return false;
    }

    const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! root || GetForegroundWindow() != root)
    {
        return false;
    }

    const auto containsTarget = [targetWindow](HWND paneWindow) noexcept
    {
        return paneWindow && IsWindow(paneWindow) != FALSE &&
            (targetWindow == paneWindow || IsChild(paneWindow, targetWindow) != FALSE);
    };
    const auto targetBelongsToPane = [&](const PaneState& state) noexcept
    {
        return containsTarget(state.hNavigationView.get()) || containsTarget(state.hFolderView.get()) ||
            containsTarget(state.hFilterBar.get()) || containsTarget(state.hStatusBar.get()) ||
            containsTarget(state.hPreviewTabs.get()) || containsTarget(state.hPreviewContent.get()) ||
            containsTarget(state.terminalHwnd);
    };

    std::optional<Pane> targetPane;
    if (targetBelongsToPane(_leftPane))
    {
        targetPane = Pane::Left;
    }
    else if (targetBelongsToPane(_rightPane))
    {
        targetPane = Pane::Right;
    }
    if (! targetPane.has_value())
    {
        return false;
    }

    PaneState& targetState = targetPane.value() == Pane::Left ? _leftPane : _rightPane;
    if (targetState.previewTabSelected)
    {
        return false;
    }

    const HWND currentFocus = GetFocus();
    const auto containsFocus = [currentFocus](HWND paneWindow) noexcept
    {
        return paneWindow && currentFocus &&
            (currentFocus == paneWindow || IsChild(paneWindow, currentFocus) != FALSE);
    };
    const bool focusAlreadyInTargetPane = containsFocus(targetState.hNavigationView.get()) || containsFocus(targetState.hFolderView.get()) ||
        containsFocus(targetState.hFilterBar.get()) || containsFocus(targetState.hStatusBar.get()) || containsFocus(targetState.hPreviewTabs.get()) ||
        containsFocus(targetState.hPreviewContent.get()) || containsFocus(targetState.terminalHwnd);
    if (focusAlreadyInTargetPane)
    {
        return false;
    }

    HWND focusTarget = targetState.terminalTabSelected ? targetState.terminalHwnd : GetPanePreferredFocusTarget(targetPane.value());
    if (! focusTarget || IsWindow(focusTarget) == FALSE || IsWindowVisible(focusTarget) == FALSE)
    {
        return false;
    }

    Debug::Perf::Scope perf(L"pane.focus_follows_pointer.transition_us");
    perf.SetValue0(targetPane.value() == Pane::Left ? 0u : 1u);
    perf.SetValue1(terminalDisplayed ? 1u : 0u);

    SetFocus(focusTarget);
    const bool focused = GetFocus() == focusTarget || IsChild(focusTarget, GetFocus()) != FALSE;
    if (focused)
    {
        SetActivePane(targetPane.value());
        UpdatePaneFocusStates();
    }
    perf.SetHr(focused ? S_OK : E_FAIL);
    return focused;
}

HWND FolderWindow::GetFolderViewHwnd(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.hFolderView.get();
}

bool FolderWindow::TryRestoreActivePaneFolderViewFocus() noexcept
{
    if (! _hWnd)
    {
        return false;
    }
    const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! root)
    {
        return false;
    }
    if (const HWND activeWindow = GetActiveWindow(); activeWindow && activeWindow != root)
    {
        return false;
    }
    if (const HWND foregroundWindow = GetForegroundWindow(); foregroundWindow && foregroundWindow != root)
    {
        return false;
    }
    if (GetFocusedFolderViewHwnd() != nullptr)
    {
        return false;
    }

    FocusPaneFolderView(_activePane);
    return GetFocusedFolderViewHwnd() != nullptr;
}

void FolderWindow::RequestRestoreFolderViewFocus(HWND folderView) noexcept
{
    Pane pane;
    if (_leftPane.hFolderView && folderView == _leftPane.hFolderView.get())
    {
        pane = Pane::Left;
    }
    else if (_rightPane.hFolderView && folderView == _rightPane.hFolderView.get())
    {
        pane = Pane::Right;
    }
    else
    {
        return;
    }

    const HWND root = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! root)
    {
        return;
    }
    if (const HWND activeWindow = GetActiveWindow(); activeWindow && activeWindow != root)
    {
        return;
    }
    if (const HWND foregroundWindow = GetForegroundWindow(); foregroundWindow && foregroundWindow != root)
    {
        return;
    }

    SetActivePane(pane);
    if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
    {
        if (IsIconic(_hWnd.get()) != FALSE)
        {
            ShowWindow(_hWnd.get(), SW_RESTORE);
        }
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
        static_cast<void>(SetActiveWindow(_hWnd.get()));
    }

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (state.hNavigationView && IsWindow(state.hNavigationView.get()) != FALSE &&
        PostMessageW(state.hNavigationView.get(), WndMsg::kNavigationViewRestoreFolderFocus, 0, 0))
    {
        return;
    }

    FocusPaneFolderView(pane);
}

FolderWindow::Pane FolderWindow::GetPaneFromChild(HWND child) const noexcept
{
    if (! child)
    {
        return _activePane;
    }

    const auto belongsTo = [child](const PaneState& state) noexcept
    {
        const auto contains = [child](HWND paneWindow) noexcept
        {
            return paneWindow && IsWindow(paneWindow) != FALSE &&
                (child == paneWindow || IsChild(paneWindow, child) != FALSE);
        };
        return contains(state.hFolderView.get()) || contains(state.hNavigationView.get()) ||
            contains(state.hFilterBar.get()) || contains(state.hStatusBar.get()) ||
            contains(state.hPreviewTabs.get()) || contains(state.hPreviewContent.get()) || contains(state.terminalHwnd);
    };

    if (belongsTo(_leftPane))
    {
        return Pane::Left;
    }
    if (belongsTo(_rightPane))
    {
        return Pane::Right;
    }

    return _activePane;
}

void FolderWindow::OnLButtonDown(POINT pt)
{
    const SplitterArrowZone arrowZone = HitTestSplitterArrow(pt);
    if (arrowZone != SplitterArrowZone::None)
    {
        SetHoveredSplitterArrowZone(arrowZone);
        ToggleZoomPanel(GetSplitterArrowTargetPane(arrowZone));
        return;
    }

    if (PtInRect(&_splitterRect, pt))
    {
        _draggingSplitter     = true;
        _splitterDragOffsetPx = pt.x - _splitterRect.left;
        SetCapture(_hWnd.get());
        return;
    }

    if (pt.x < _splitterRect.left)
    {
        SetActivePane(Pane::Left);
        if (GetFocusedPane() != Pane::Left)
        {
            FocusPaneFolderView(Pane::Left);
        }
    }
    else if (pt.x > _splitterRect.right)
    {
        SetActivePane(Pane::Right);
        if (GetFocusedPane() != Pane::Right)
        {
            FocusPaneFolderView(Pane::Right);
        }
    }
}

void FolderWindow::OnLButtonDblClk(POINT pt)
{
    if (HitTestSplitterArrow(pt) != SplitterArrowZone::None || ! PtInRect(&_splitterRect, pt))
    {
        return;
    }

    _draggingSplitter = false;
    ReleaseCapture();
    SetSplitRatio(0.5f);
}

void FolderWindow::OnLButtonUp()
{
    if (_draggingSplitter)
    {
        _draggingSplitter = false;
        ReleaseCapture();
    }
}

void FolderWindow::OnMouseMove(POINT pt)
{
    if (! _draggingSplitter)
    {
        const SplitterArrowZone arrowZone = HitTestSplitterArrow(pt);
        SetHoveredSplitterArrowZone(arrowZone);
        if (arrowZone != SplitterArrowZone::None)
        {
            TrackSplitterMouseLeave();
        }
        return;
    }

    SetHoveredSplitterArrowZone(SplitterArrowZone::None);

    const int splitterWidth  = _splitterRect.right - _splitterRect.left;
    const int availableWidth = std::max(0L, _clientSize.cx - splitterWidth);
    if (availableWidth <= 0)
    {
        return;
    }

    int desiredLeftWidth = pt.x - _splitterDragOffsetPx;
    desiredLeftWidth     = std::clamp(desiredLeftWidth, 0, availableWidth);

    const float ratio = static_cast<float>(desiredLeftWidth) / static_cast<float>(availableWidth);
    SetSplitRatio(ratio);

    if (_hWnd)
    {
        UpdateWindow(_hWnd.get());
    }
}

void FolderWindow::OnCaptureChanged()
{
    _draggingSplitter = false;
}

void FolderWindow::OnMouseLeave()
{
    _trackingSplitterMouseLeave = false;
    SetHoveredSplitterArrowZone(SplitterArrowZone::None);
}

bool FolderWindow::OnSetCursor(POINT pt)
{
    if (HitTestSplitterArrow(pt) != SplitterArrowZone::None)
    {
        SetCursor(LoadCursor(nullptr, IDC_HAND));
        return true;
    }

    if (PtInRect(&_splitterRect, pt))
    {
        SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
        return true;
    }
    return false;
}

void FolderWindow::OnParentNotify(UINT eventMsg, UINT childId)
{
    if (eventMsg != WM_LBUTTONDOWN && eventMsg != WM_RBUTTONDOWN && eventMsg != WM_MBUTTONDOWN)
    {
        return;
    }

    if (childId == kLeftNavigationId || childId == kLeftFolderViewId)
    {
        SetActivePane(Pane::Left);
        if (GetFocusedPane() != Pane::Left)
        {
            FocusPaneFolderView(Pane::Left);
        }
    }
    else if (childId == kRightNavigationId || childId == kRightFolderViewId)
    {
        SetActivePane(Pane::Right);
        if (GetFocusedPane() != Pane::Right)
        {
            FocusPaneFolderView(Pane::Right);
        }
    }
}

RECT FolderWindow::GetSplitterArrowRect(SplitterArrowZone zone) const noexcept
{
    if (zone == SplitterArrowZone::None || _splitterRect.right <= _splitterRect.left || _splitterRect.bottom <= _splitterRect.top)
    {
        return RECT{};
    }

    const int dpi            = std::max(1, static_cast<int>(_dpi));
    const int navHeight      = std::max(1, MulDiv(NavigationView::kHeight, dpi, USER_DEFAULT_SCREEN_DPI));
    const int splitterHeight = std::max(0L, _splitterRect.bottom - _splitterRect.top);
    if (splitterHeight <= 0)
    {
        return RECT{};
    }

    const int arrowHeight = std::min(navHeight, splitterHeight);
    if ((arrowHeight * 2) > splitterHeight)
    {
        const LONG midpoint = _splitterRect.top + (splitterHeight / 2);
        if (zone == SplitterArrowZone::Left)
        {
            return RECT{_splitterRect.left, _splitterRect.top, _splitterRect.right, midpoint};
        }

        return RECT{_splitterRect.left, midpoint, _splitterRect.right, _splitterRect.bottom};
    }

    if (zone == SplitterArrowZone::Left)
    {
        return RECT{_splitterRect.left, _splitterRect.top, _splitterRect.right, _splitterRect.top + arrowHeight};
    }

    return RECT{_splitterRect.left, _splitterRect.bottom - arrowHeight, _splitterRect.right, _splitterRect.bottom};
}

FolderWindow::Pane FolderWindow::GetSplitterArrowTargetPane(SplitterArrowZone zone) const noexcept
{
    if (_zoomedPane == Pane::Left)
    {
        return Pane::Right;
    }

    if (_zoomedPane == Pane::Right)
    {
        return Pane::Left;
    }

    return zone == SplitterArrowZone::Right ? Pane::Right : Pane::Left;
}

wchar_t FolderWindow::GetSplitterArrowGlyph(SplitterArrowZone zone) const noexcept
{
    return GetSplitterArrowTargetPane(zone) == Pane::Left ? FluentIcons::kChevronRightSmall : FluentIcons::kChevronLeftSmall;
}

FolderWindow::SplitterArrowZone FolderWindow::HitTestSplitterArrow(POINT pt) const noexcept
{
    const RECT leftArrow = GetSplitterArrowRect(SplitterArrowZone::Left);
    if (PtInRect(&leftArrow, pt) != FALSE)
    {
        return SplitterArrowZone::Left;
    }

    const RECT rightArrow = GetSplitterArrowRect(SplitterArrowZone::Right);
    if (PtInRect(&rightArrow, pt) != FALSE)
    {
        return SplitterArrowZone::Right;
    }

    return SplitterArrowZone::None;
}

void FolderWindow::SetHoveredSplitterArrowZone(SplitterArrowZone zone) noexcept
{
    if (_hoveredSplitterArrowZone == zone)
    {
        return;
    }

    const RECT previousRect   = GetSplitterArrowRect(_hoveredSplitterArrowZone);
    _hoveredSplitterArrowZone = zone;
    const RECT nextRect       = GetSplitterArrowRect(_hoveredSplitterArrowZone);

    if (_hWnd)
    {
        if (previousRect.right > previousRect.left && previousRect.bottom > previousRect.top)
        {
            InvalidateRect(_hWnd.get(), &previousRect, FALSE);
        }
        if (nextRect.right > nextRect.left && nextRect.bottom > nextRect.top)
        {
            InvalidateRect(_hWnd.get(), &nextRect, FALSE);
        }
    }
}

void FolderWindow::TrackSplitterMouseLeave() noexcept
{
    if (_trackingSplitterMouseLeave || ! _hWnd)
    {
        return;
    }

    TRACKMOUSEEVENT tme{};
    tme.cbSize    = sizeof(tme);
    tme.dwFlags   = TME_LEAVE;
    tme.hwndTrack = _hWnd.get();
    if (TrackMouseEvent(&tme) != FALSE)
    {
        _trackingSplitterMouseLeave = true;
    }
}

#ifdef ENABLE_TESTS
bool FolderWindow::DebugGetSplitterSnapshot(FolderWindowSplitterDebugSnapshot& out) const noexcept
{
    out                      = {};
    out.splitterRect         = _splitterRect;
    out.leftArrowRect        = GetSplitterArrowRect(SplitterArrowZone::Left);
    out.rightArrowRect       = GetSplitterArrowRect(SplitterArrowZone::Right);
    out.leftArrowTargetPane  = GetSplitterArrowTargetPane(SplitterArrowZone::Left);
    out.rightArrowTargetPane = GetSplitterArrowTargetPane(SplitterArrowZone::Right);
    out.leftArrowGlyph       = GetSplitterArrowGlyph(SplitterArrowZone::Left);
    out.rightArrowGlyph      = GetSplitterArrowGlyph(SplitterArrowZone::Right);
    out.arrowColor           = GetSplitterArrowColor();
    out.gripColor            = GetSplitterGripColor();
    out.arrowChevronSizePx   = GetSplitterArrowChevronSizePx();
    out.gripDotSizePx        = GetSplitterGripDotSizePx();
    out.leftArrowCursorHand =
        HitTestSplitterArrow({out.leftArrowRect.left + ((out.leftArrowRect.right - out.leftArrowRect.left) / 2),
                              out.leftArrowRect.top + ((out.leftArrowRect.bottom - out.leftArrowRect.top) / 2)}) == SplitterArrowZone::Left;
    out.rightArrowCursorHand =
        HitTestSplitterArrow({out.rightArrowRect.left + ((out.rightArrowRect.right - out.rightArrowRect.left) / 2),
                              out.rightArrowRect.top + ((out.rightArrowRect.bottom - out.rightArrowRect.top) / 2)}) == SplitterArrowZone::Right;

    if (_hoveredSplitterArrowZone == SplitterArrowZone::Left)
    {
        out.hoveredArrowPane = Pane::Left;
    }
    else if (_hoveredSplitterArrowZone == SplitterArrowZone::Right)
    {
        out.hoveredArrowPane = Pane::Right;
    }

    return true;
}

bool FolderWindow::DebugHoverSplitterArrow(Pane pane) noexcept
{
    const SplitterArrowZone zone = pane == Pane::Left ? SplitterArrowZone::Left : SplitterArrowZone::Right;
    const RECT arrowRect         = GetSplitterArrowRect(zone);
    if (arrowRect.right <= arrowRect.left || arrowRect.bottom <= arrowRect.top)
    {
        return false;
    }

    const POINT pt{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    OnMouseMove(pt);
    return _hoveredSplitterArrowZone == zone;
}

bool FolderWindow::DebugClickSplitterArrow(Pane pane) noexcept
{
    const SplitterArrowZone zone = pane == Pane::Left ? SplitterArrowZone::Left : SplitterArrowZone::Right;
    const RECT arrowRect         = GetSplitterArrowRect(zone);
    if (arrowRect.right <= arrowRect.left || arrowRect.bottom <= arrowRect.top)
    {
        return false;
    }

    const POINT pt{arrowRect.left + ((arrowRect.right - arrowRect.left) / 2), arrowRect.top + ((arrowRect.bottom - arrowRect.top) / 2)};
    OnLButtonDown(pt);
    return true;
}
#endif
