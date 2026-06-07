// Commands.SelfTest.Dialogs.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Dialogs test family: 48 test functions.

template <typename WorkerFunc> void RunAboutDialogModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND about = WaitForWindow([] noexcept { return GetAboutDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        workerFunc(about);
    });

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_ABOUT, 0), 0);
    worker.join();
}

template <typename Task> [[nodiscard]] auto RunDialogUiaTaskWithMessagePump(std::wstring_view label, Task&& task) noexcept -> decltype(task())
{
    using Result = decltype(task());
    static_assert(std::is_default_constructible_v<Result>,
                  "RunDialogUiaTaskWithMessagePump requires a default-constructible result so timeout fallback can return a safe sentinel.");

    using namespace std::chrono_literals;

    using TaskType = std::decay_t<Task>;
    struct SharedState final
    {
        SharedState()                              = default;
        SharedState(const SharedState&)            = delete;
        SharedState& operator=(const SharedState&) = delete;
        SharedState(SharedState&&)                 = delete;
        SharedState& operator=(SharedState&&)      = delete;

        std::optional<Result> result;
        std::atomic<bool> done = false;
    };

    auto sharedState = std::make_shared<SharedState>();
    auto sharedTask  = std::make_shared<TaskType>(std::forward<Task>(task));

    std::jthread worker([sharedState, sharedTask](std::stop_token) noexcept
    {
        sharedState->result = (*sharedTask)();
        sharedState->done.store(true, std::memory_order_release);
    });

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (! sharedState->done.load(std::memory_order_acquire))
    {
        if (std::chrono::steady_clock::now() >= deadline)
        {
            Trace(std::format(L"Dialog UIA task timed out during '{}'; returning default result and detaching worker.", label));
            worker.request_stop();
            worker.detach();
            return Result{};
        }

        PumpPendingMessages();
        std::this_thread::sleep_for(10ms);
    }

    worker.join();
    return std::move(sharedState->result).value_or(Result{});
}

template <typename WorkerFunc> void RunRenamePromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt  = WaitForWindow([] noexcept { return GetFolderViewRenamePromptHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_RENAME, 0), 0);
    worker.join();
}

template <typename WorkerFunc> void RunCreateDirectoryPromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow([] noexcept { return GetFolderViewCreateDirectoryPromptHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CREATE_DIR, 0), 0);
    worker.join();
}

template <typename WorkerFunc> void RunEditNewPromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt  = WaitForWindow([] noexcept { return GetFolderViewEditNewPromptHandle(); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_EDIT_NEW, 0), 0);
    worker.join();
}

template <typename WorkerFunc> void RunChangeAttributesOptionsPromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow(
            [mainWindow]() noexcept -> HWND
        {
            const HWND dlg = GetChangeAttributesOptionsPromptHandle();
            if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
            {
                return nullptr;
            }
            return dlg;
        },
            SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_ATTRIBUTES, 0), 0);
    worker.join();
}

template <typename WorkerFunc> void RunMakeFileListOptionsPromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow(
            [mainWindow]() noexcept -> HWND
        {
            const HWND dlg = GetMakeFileListOptionsPromptHandle();
            if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
            {
                return nullptr;
            }
            return dlg;
        },
            SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_MAKE_FILE_LIST, 0), 0);
    worker.join();
}

template <typename WorkerFunc> void RunArchivePackPromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow(
            [mainWindow]() noexcept -> HWND
        {
            const HWND dlg = GetArchivePackPromptHandle();
            if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
            {
                return nullptr;
            }
            return dlg;
        },
            SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_PACK, 0), 0);
    worker.join();
}

template <typename WorkerFunc> void RunArchiveUnpackPromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow(
            [mainWindow]() noexcept -> HWND
        {
            const HWND dlg = GetArchiveUnpackPromptHandle();
            if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
            {
                return nullptr;
            }
            return dlg;
        },
            SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_UNPACK, 0), 0);
    worker.join();
}

[[nodiscard]] std::optional<UiaDescendantPatternStats> WaitForVisiblePromptButtonStats(HWND hwnd, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    std::optional<UiaDescendantPatternStats> lastStats;
    while (std::chrono::steady_clock::now() < deadline)
    {
        lastStats = CollectVisibleUiaDescendantPatternStats(hwnd);
        if (lastStats.has_value() && lastStats->buttonControlCount > 0u && lastStats->invokePatternCount > 0u)
        {
            return lastStats;
        }

        std::this_thread::sleep_for(20ms);
    }

    return lastStats.has_value() ? lastStats : CollectVisibleUiaDescendantPatternStats(hwnd);
}

template <typename WorkerFunc> void RunChangeCasePromptModalCycle(HWND mainWindow, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow(
            [mainWindow]() noexcept -> HWND
        {
            const HWND dlg = GetFolderViewChangeCasePromptHandle();
            if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
            {
                return nullptr;
            }
            return dlg;
        },
            SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_CASE, 0), 0);
    worker.join();
}

template <typename WorkerFunc> void RunSelectionMaskPromptModalCycle(HWND mainWindow, const UINT commandId, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow(
            [mainWindow]() noexcept -> HWND
        {
            const HWND dlg = GetFolderViewSelectionMaskPromptHandle();
            if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
            {
                return nullptr;
            }
            return dlg;
        },
            SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
    worker.join();
}

template <typename WorkerFunc> void RunPaneFilterPromptModalCycle(HWND mainWindow, const UINT commandId, WorkerFunc&& workerFunc) noexcept
{
    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND prompt = WaitForWindow(
            [mainWindow]() noexcept -> HWND
        {
            const HWND dlg = GetFolderViewPaneFilterPromptHandle();
            if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
            {
                return nullptr;
            }
            return dlg;
        },
            SelfTest::Scale(std::chrono::milliseconds{5000}));
        const auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (prompt && IsWindow(prompt) != FALSE)
            {
                PostMessageW(prompt, WM_CLOSE, 0, 0);
                static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(std::chrono::milliseconds{3000})));
            }
        });

        workerFunc(prompt);
    });

    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
    worker.join();
}

[[nodiscard]] bool TestAboutDialogUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetAboutDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{3000})),
                      L"Existing About dialog did not close before DXUI validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const auto validateAboutSurface = [&](std::wstring_view label) noexcept -> bool
    {
        struct ProbeResult final
        {
            HWND about                     = nullptr;
            bool sawDialog                 = false;
            bool independentTopLevel       = false;
            size_t visibleChildWindowCount = 0u;
            bool exposesUiaProvider        = false;
            std::optional<UiaDescendantPatternStats> uiaPatternStats;
            std::optional<UiaNamedElementState> aboutTextState;
            std::optional<UiaNamedElementState> buttonState;
            bool closed = false;
        } probe{};

        RunAboutDialogModalCycle(mainWindow,
                                 [&](HWND about) noexcept
        {
            probe.about     = about;
            probe.sawDialog = about != nullptr && IsWindow(about) != FALSE;
            if (! probe.sawDialog)
            {
                return;
            }

            probe.independentTopLevel     = ! IsOwnedBy(about, mainWindow);
            probe.visibleChildWindowCount = CountVisibleChildWindows(about);
            probe.exposesUiaProvider      = WindowExposesUiaProvider(about);
            probe.uiaPatternStats         = CollectVisibleUiaDescendantPatternStats(about);
            probe.aboutTextState          = CollectVisibleDescendantNamedElementState(about, UIA_TextControlTypeId);
            probe.buttonState             = CollectVisibleDescendantNamedElementState(about, UIA_ButtonControlTypeId);

            PostMessageW(about, WM_CLOSE, 0, 0);
            probe.closed = WaitForWindowClosed(about, SelfTest::Scale(std::chrono::milliseconds{3000}));
        });

        state.Require(probe.sawDialog, std::format(L"About dialog did not open during {}.", label));
        if (! probe.sawDialog)
        {
            return false;
        }

        state.Require(probe.independentTopLevel, L"About dialog should be an independent top-level window.");
        state.Require(probe.visibleChildWindowCount == 0u, L"About dialog should not expose visible child-control fallback.");
        state.Require(probe.exposesUiaProvider, L"About dialog should answer WM_GETOBJECT with a UI Automation provider.");

        state.Require(probe.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the About dialog during {}.", label));
        if (probe.uiaPatternStats.has_value())
        {
            state.Require(probe.uiaPatternStats->visibleElementCount > 0u, L"About dialog should expose visible UI Automation descendants.");
            state.Require(probe.uiaPatternStats->buttonControlCount > 0u,
                          L"About dialog should expose a visible UI Automation button descendant for the DX footer action.");
            state.Require(probe.uiaPatternStats->invokePatternCount > 0u,
                          L"About dialog should expose live UI Automation InvokePattern support for the DX footer action.");
        }

        state.Require(probe.aboutTextState.has_value(), std::format(L"About dialog should expose a visible UI Automation text descendant during {}.", label));
        if (probe.aboutTextState.has_value())
        {
            state.Require(probe.aboutTextState->name.find(L"RedSalamander") != std::wstring::npos,
                          std::format(L"About dialog text surface should expose product text during {}; saw '{}'.", label, probe.aboutTextState->name));
        }

        state.Require(probe.buttonState.has_value(), std::format(L"About dialog should expose a visible DX footer button during {}.", label));
        if (probe.buttonState.has_value())
        {
            state.Require(! probe.buttonState->name.empty(),
                          std::format(L"About dialog visible DX footer button should expose a stable accessible name during {}.", label));
        }

        state.Require(probe.closed, std::format(L"About dialog did not close cleanly during {}.", label));

        return state.failure.empty();
    };

    state.Require(validateAboutSurface(L"the initial About surface pass"), L"About dialog should expose the expected DX surface on the initial pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(validateAboutSurface(L"the reopened About surface pass"), L"About dialog should expose the same DX surface after reopen.");
    return state.failure.empty();
}

[[nodiscard]] bool TestAboutDialogLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetAboutDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing About dialog did not close before live interaction validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const auto runCycle = [&](std::wstring_view context) noexcept
    {
        struct ProbeResult final
        {
            HWND about                     = nullptr;
            bool sawDialog                 = false;
            bool independentTopLevel       = false;
            size_t visibleChildWindowCount = 0u;
            bool exposesUiaProvider        = false;
            std::optional<UiaDescendantPatternStats> uiaPatternStats;
            std::optional<UiaNamedElementState> aboutTextState;
            bool invokedButton = false;
            bool closed        = false;
        } probe{};

        RunAboutDialogModalCycle(mainWindow,
                                 [&](HWND about) noexcept
        {
            probe.about     = about;
            probe.sawDialog = about != nullptr && IsWindow(about) != FALSE;
            if (! probe.sawDialog)
            {
                return;
            }

            probe.independentTopLevel     = ! IsOwnedBy(about, mainWindow);
            probe.visibleChildWindowCount = CountVisibleChildWindows(about);
            probe.exposesUiaProvider      = WindowExposesUiaProvider(about);
            probe.uiaPatternStats         = CollectVisibleUiaDescendantPatternStats(about);
            probe.aboutTextState          = CollectVisibleDescendantNamedElementState(about, UIA_TextControlTypeId);
            probe.invokedButton           = InvokeVisibleDescendantByName(about, UIA_ButtonControlTypeId, L"");
            probe.closed                  = WaitForWindowClosed(about, SelfTest::Scale(3000ms));
        });

        state.Require(probe.sawDialog, std::format(L"About dialog did not open for {}.", context));
        state.Require(probe.independentTopLevel, std::format(L"About dialog should remain an independent top-level window during {}.", context));
        state.Require(probe.visibleChildWindowCount == 0u, std::format(L"About dialog should not expose visible child-control fallback during {}.", context));
        state.Require(probe.exposesUiaProvider, std::format(L"About dialog should answer WM_GETOBJECT during {}.", context));
        state.Require(probe.uiaPatternStats.has_value(), std::format(L"Failed to collect live UI Automation stats for About dialog during {}.", context));
        if (probe.uiaPatternStats.has_value())
        {
            state.Require(probe.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"About dialog should expose a visible DX footer button during {}.", context));
            state.Require(probe.uiaPatternStats->invokePatternCount > 0u, std::format(L"About dialog should expose InvokePattern during {}.", context));
        }

        state.Require(probe.aboutTextState.has_value(), std::format(L"About dialog should expose a visible text descendant during {}.", context));
        if (probe.aboutTextState.has_value())
        {
            state.Require(probe.aboutTextState->name.find(L"RedSalamander") != std::wstring::npos,
                          std::format(L"About dialog text surface should expose product text during {}; saw '{}'.", context, probe.aboutTextState->name));
        }
        state.Require(probe.invokedButton, std::format(L"Failed to invoke the visible DX footer button on About dialog during {}.", context));
        state.Require(probe.closed, std::format(L"About dialog did not close after live UIA InvokePattern interaction during {}.", context));
        return state.failure.empty();
    };

    state.Require(runCycle(L"initial live interaction"), L"Initial About dialog live DX validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runCycle(L"reopened live interaction"), L"Reopened About dialog live DX close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestAboutDialogAccessKeyRoutesOk(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetAboutDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing About dialog did not close before access-key validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeResult final
    {
        HWND about                     = nullptr;
        bool sawDialog                 = false;
        size_t visibleChildWindowCount = 0u;
        bool exposesUiaProvider        = false;
        bool closed                    = false;
    } probe{};

    RunAboutDialogModalCycle(mainWindow,
                             [&](HWND about) noexcept
    {
        probe.about     = about;
        probe.sawDialog = about != nullptr && IsWindow(about) != FALSE;
        if (! probe.sawDialog)
        {
            return;
        }

        probe.visibleChildWindowCount = CountVisibleChildWindows(about);
        probe.exposesUiaProvider      = WindowExposesUiaProvider(about);
        SendMessageW(about, WM_SYSCHAR, static_cast<WPARAM>(L'o'), 0);
        probe.closed = WaitForWindowClosed(about, SelfTest::Scale(3000ms));
    });

    state.Require(probe.sawDialog, L"About dialog did not open for access-key validation.");
    if (! probe.sawDialog)
    {
        return false;
    }

    state.Require(probe.visibleChildWindowCount == 0u, L"About dialog should not expose visible child-control fallback during access-key validation.");
    state.Require(probe.exposesUiaProvider, L"About dialog should answer WM_GETOBJECT during access-key validation.");
    state.Require(probe.closed, L"About dialog did not close after the DX OK access key.");
    state.Require(GetAboutDialogHandle() == nullptr || IsWindow(GetAboutDialogHandle()) == FALSE,
                  L"About dialog should not remain open after the DX OK access key.");
    return state.failure.empty();
}

[[nodiscard]] bool TestAboutDialogEnterAndEscapeRouteOk(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExisting = [&]() noexcept
    {
        if (const HWND existing = GetAboutDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    closeExisting();

    const auto runCycle = [&](std::wstring_view context, const WPARAM key) noexcept
    {
        struct ProbeResult final
        {
            HWND about                     = nullptr;
            bool sawDialog                 = false;
            size_t visibleChildWindowCount = 0u;
            bool exposesUiaProvider        = false;
            bool closed                    = false;
        } probe{};

        RunAboutDialogModalCycle(mainWindow,
                                 [&](HWND about) noexcept
        {
            probe.about     = about;
            probe.sawDialog = about != nullptr && IsWindow(about) != FALSE;
            if (! probe.sawDialog)
            {
                return;
            }

            probe.visibleChildWindowCount = CountVisibleChildWindows(about);
            probe.exposesUiaProvider      = WindowExposesUiaProvider(about);
            SendMessageW(about, WM_KEYDOWN, key, 0);
            SendMessageW(about, WM_KEYUP, key, 0);
            probe.closed = WaitForWindowClosed(about, SelfTest::Scale(3000ms));
        });

        state.Require(probe.sawDialog, std::format(L"About dialog did not open for {}.", context));
        if (! probe.sawDialog)
        {
            return false;
        }

        state.Require(probe.visibleChildWindowCount == 0u, std::format(L"About dialog should not expose visible child-control fallback during {}.", context));
        state.Require(probe.exposesUiaProvider, std::format(L"About dialog should answer WM_GETOBJECT during {}.", context));
        state.Require(probe.closed, std::format(L"About dialog did not close after {}.", context));
        return state.failure.empty();
    };

    state.Require(runCycle(L"Enter default-button routing", VK_RETURN), L"About dialog Enter default-button validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runCycle(L"Escape cancel routing", VK_ESCAPE), L"About dialog Escape cancel validation failed.");
    state.Require(GetAboutDialogHandle() == nullptr || IsWindow(GetAboutDialogHandle()) == FALSE,
                  L"About dialog should not remain open after Enter/Escape keyboard-flow validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestAboutDialogLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingAbout = [&]() noexcept
    {
        if (const HWND existing = GetAboutDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    closeExistingAbout();

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        struct CycleResult final
        {
            HWND about                     = nullptr;
            bool sawDialog                 = false;
            bool independentTopLevel       = false;
            size_t visibleChildWindowCount = 0u;
            bool exposesUiaProvider        = false;
            std::optional<UiaDescendantPatternStats> uiaPatternStats;
            std::optional<UiaNamedElementState> aboutTextState;
            std::optional<UiaNamedElementState> buttonState;
            bool closed = false;
        } cycleResult{};

        RunAboutDialogModalCycle(mainWindow,
                                 [&](HWND about) noexcept
        {
            cycleResult.about     = about;
            cycleResult.sawDialog = about != nullptr && IsWindow(about) != FALSE;
            if (! cycleResult.sawDialog)
            {
                return;
            }

            cycleResult.independentTopLevel     = ! IsOwnedBy(about, mainWindow);
            cycleResult.visibleChildWindowCount = CountVisibleChildWindows(about);
            cycleResult.exposesUiaProvider      = WindowExposesUiaProvider(about);
            cycleResult.uiaPatternStats         = CollectVisibleUiaDescendantPatternStats(about);
            cycleResult.aboutTextState          = CollectVisibleDescendantNamedElementState(about, UIA_TextControlTypeId);
            cycleResult.buttonState             = CollectVisibleDescendantNamedElementState(about, UIA_ButtonControlTypeId);
            PostMessageW(about, WM_CLOSE, 0, 0);
            cycleResult.closed = WaitForWindowClosed(about, SelfTest::Scale(3000ms));
        });

        state.Require(cycleResult.sawDialog, std::format(L"About dialog did not open during cycle {}.", cycle));
        if (! cycleResult.sawDialog)
        {
            return false;
        }

        state.Require(cycleResult.independentTopLevel, std::format(L"About dialog should be an independent top-level window during cycle {}.", cycle));
        state.Require(cycleResult.visibleChildWindowCount == 0u,
                      std::format(L"About dialog should not expose visible child-control fallback during cycle {}.", cycle));
        state.Require(cycleResult.exposesUiaProvider, std::format(L"About dialog should answer WM_GETOBJECT during cycle {}.", cycle));

        state.Require(cycleResult.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the About dialog during cycle {}.", cycle));
        if (cycleResult.uiaPatternStats.has_value())
        {
            state.Require(cycleResult.uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"About dialog should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"About dialog should expose a visible UI Automation button descendant during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"About dialog should expose live UI Automation InvokePattern support during cycle {}.", cycle));
        }

        state.Require(cycleResult.aboutTextState.has_value(),
                      std::format(L"About dialog should expose a visible UI Automation text descendant during cycle {}.", cycle));
        if (cycleResult.aboutTextState.has_value())
        {
            state.Require(
                cycleResult.aboutTextState->name.find(L"RedSalamander") != std::wstring::npos,
                std::format(L"About dialog text surface should expose product text during cycle {}; saw '{}'.", cycle, cycleResult.aboutTextState->name));
        }

        state.Require(cycleResult.buttonState.has_value(), std::format(L"About dialog should expose a visible DX footer button during cycle {}.", cycle));
        if (cycleResult.buttonState.has_value())
        {
            state.Require(! cycleResult.buttonState->name.empty(),
                          std::format(L"About dialog visible DX footer button should expose a stable accessible name during cycle {}.", cycle));
        }
        state.Require(cycleResult.closed, std::format(L"About dialog did not close cleanly during cycle {}.", cycle));
    }

    state.Require(GetAboutDialogHandle() == nullptr || IsWindow(GetAboutDialogHandle()) == FALSE, L"About dialog should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestAppPromptUsesAlertOverlayWindow(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    constexpr std::wstring_view kAlertOverlayWindowClassName = L"RedSalamander.AlertOverlayWindow";

    const auto getOverlayWindow = [&]() noexcept -> HWND { return FindVisibleDescendantWindowByClass(mainWindow, kAlertOverlayWindowClassName); };

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(false);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });

    const auto closeExistingOverlay = [&]() noexcept
    {
        if (const HWND existing = getOverlayWindow(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(existing, WM_KEYUP, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    closeExistingOverlay();

    struct PromptAutomationResult final
    {
        PromptAutomationResult()                                         = default;
        PromptAutomationResult(const PromptAutomationResult&)            = delete;
        PromptAutomationResult& operator=(const PromptAutomationResult&) = delete;
        PromptAutomationResult(PromptAutomationResult&&)                 = delete;
        PromptAutomationResult& operator=(PromptAutomationResult&&)      = delete;

        std::atomic<bool> sawOverlay{false};
        std::atomic<bool> closed{false};
        HWND overlay            = nullptr;
        HRESULT hr              = E_FAIL;
        HostPromptResult result = HOST_PROMPT_RESULT_NONE;
    };

    const auto runPrompt = [&](std::wstring_view title,
                               std::wstring_view message,
                               WPARAM closeKey,
                               HostPromptResult expectedResult,
                               std::wstring_view operationName,
                               std::wstring_view invokeButtonName = {}) noexcept
    {
        PromptAutomationResult promptResult{};

        std::jthread automation([&](std::stop_token) noexcept
        {
            const HWND overlay   = WaitForWindow(getOverlayWindow, SelfTest::Scale(5000ms));
            promptResult.overlay = overlay;
            state.Require(overlay != nullptr && IsWindow(overlay) != FALSE, std::format(L"Alert overlay prompt did not open for {}.", operationName));
            if (! overlay || IsWindow(overlay) == FALSE)
            {
                return;
            }

            promptResult.sawOverlay.store(true, std::memory_order_release);

            state.Require(IsOwnedBy(overlay, mainWindow),
                          std::format(L"Alert overlay prompt should be owned by the main window root for {}.", operationName));
            state.Require(CountVisibleChildWindows(overlay) == 0u,
                          std::format(L"Alert overlay prompt should not expose visible child fallback for {}.", operationName));
            state.Require(WindowExposesUiaProvider(overlay), std::format(L"Alert overlay prompt should answer WM_GETOBJECT for {}.", operationName));

            RECT overlayRect{};
            state.Require(GetClientRect(overlay, &overlayRect) != FALSE && overlayRect.right > overlayRect.left && overlayRect.bottom > overlayRect.top,
                          std::format(L"Alert overlay prompt should expose a non-empty client area for {}.", operationName));
            const auto uiaPatternStats = WaitForVisiblePromptButtonStats(overlay, SelfTest::Scale(3000ms));
            state.Require(uiaPatternStats.has_value(),
                          std::format(L"Failed to collect live UI Automation pattern statistics for the alert overlay prompt during {}.", operationName));
            if (uiaPatternStats.has_value() && ! invokeButtonName.empty())
            {
                state.Require(uiaPatternStats->buttonControlCount > 0u,
                              std::format(L"Alert overlay prompt should expose a visible UI Automation button descendant during {}.", operationName));
                state.Require(uiaPatternStats->invokePatternCount > 0u,
                              std::format(L"Alert overlay prompt should expose live UI Automation InvokePattern support during {}.", operationName));
            }

            if (! state.failure.empty())
            {
                SendMessageW(overlay, WM_KEYDOWN, closeKey, 0);
                SendMessageW(overlay, WM_KEYUP, closeKey, 0);
            }
            else if (! invokeButtonName.empty())
            {
                const bool invoked = InvokeVisibleDescendantByName(overlay, UIA_ButtonControlTypeId, invokeButtonName);
                state.Require(invoked,
                              std::format(L"Failed to invoke the visible DX alert overlay prompt button '{}' through live UIA interaction during {}.",
                                          invokeButtonName,
                                          operationName));
                if (! invoked)
                {
                    SendMessageW(overlay, WM_KEYDOWN, closeKey, 0);
                    SendMessageW(overlay, WM_KEYUP, closeKey, 0);
                }
            }
            else
            {
                SendMessageW(overlay, WM_KEYDOWN, closeKey, 0);
                SendMessageW(overlay, WM_KEYUP, closeKey, 0);
            }

            if (IsWindow(overlay) != FALSE)
            {
                promptResult.closed.store(WaitForWindowClosed(overlay, SelfTest::Scale(3000ms)), std::memory_order_release);
            }
            else
            {
                promptResult.closed.store(true, std::memory_order_release);
            }
        });

        HostPromptRequest request{};
        request.version       = 1u;
        request.sizeBytes     = sizeof(request);
        request.scope         = HOST_ALERT_SCOPE_APPLICATION;
        request.severity      = HOST_ALERT_INFO;
        request.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
        request.targetWindow  = nullptr;
        request.title         = title.data();
        request.message       = message.data();
        request.defaultResult = HOST_PROMPT_RESULT_OK;

        promptResult.hr = HostShowPrompt(request, nullptr, &promptResult.result);
        automation.join();

        state.Require(SUCCEEDED(promptResult.hr),
                      std::format(L"Host prompt should succeed for {}; hr=0x{:08X}.", operationName, static_cast<unsigned int>(promptResult.hr)));
        state.Require(promptResult.result == expectedResult,
                      std::format(L"Alert overlay prompt returned unexpected result for {}; saw {} expected {}.",
                                  operationName,
                                  static_cast<unsigned int>(promptResult.result),
                                  static_cast<unsigned int>(expectedResult)));
        state.Require(promptResult.closed.load(std::memory_order_acquire), std::format(L"Alert overlay prompt did not close cleanly after {}.", operationName));
    };

    runPrompt(L"Alert overlay prompt self-test",
              L"Cancel the real host-services prompt through the DX alert overlay surface.",
              VK_ESCAPE,
              HOST_PROMPT_RESULT_CANCEL,
              L"cancel",
              L"Cancel");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(getOverlayWindow() == nullptr, L"Alert overlay prompt should fully dismiss after the live DX cancel pass before reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    runPrompt(L"Alert overlay prompt self-test",
              L"Accept the real host-services prompt through the reopened DX alert overlay surface.",
              VK_RETURN,
              HOST_PROMPT_RESULT_OK,
              L"reopened accept",
              L"OK");
    return state.failure.empty();
}

[[nodiscard]] bool TestAppPromptAccessKeysRouteExpectedActions(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    constexpr std::wstring_view kAlertOverlayWindowClassName = L"RedSalamander.AlertOverlayWindow";

    const auto getOverlayWindow = [&]() noexcept -> HWND { return FindVisibleDescendantWindowByClass(mainWindow, kAlertOverlayWindowClassName); };

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(false);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });

    const auto closeExistingOverlay = [&]() noexcept
    {
        if (const HWND existing = getOverlayWindow(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(existing, WM_KEYUP, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    closeExistingOverlay();

    const auto runPrompt = [&](wchar_t mnemonic, HostPromptResult expectedResult, std::wstring_view operationName) noexcept
    {
        struct PromptAutomationResult final
        {
            HRESULT hr              = E_FAIL;
            HostPromptResult result = HOST_PROMPT_RESULT_NONE;
            HWND overlay            = nullptr;
            bool closed             = false;
        };

        PromptAutomationResult promptResult{};

        std::jthread automation([&](std::stop_token) noexcept
        {
            const HWND overlay   = WaitForWindow(getOverlayWindow, SelfTest::Scale(5000ms));
            promptResult.overlay = overlay;
            state.Require(overlay != nullptr && IsWindow(overlay) != FALSE, std::format(L"Alert overlay prompt did not open for {}.", operationName));
            if (! overlay || IsWindow(overlay) == FALSE)
            {
                return;
            }

            state.Require(IsOwnedBy(overlay, mainWindow),
                          std::format(L"Alert overlay prompt should be owned by the main window root for {}.", operationName));
            state.Require(CountVisibleChildWindows(overlay) == 0u,
                          std::format(L"Alert overlay prompt should not expose visible child fallback for {}.", operationName));
            state.Require(WindowExposesUiaProvider(overlay), std::format(L"Alert overlay prompt should answer WM_GETOBJECT for {}.", operationName));

            const auto uiaPatternStats = WaitForVisiblePromptButtonStats(overlay, SelfTest::Scale(3000ms));
            state.Require(uiaPatternStats.has_value(),
                          std::format(L"Failed to collect live UI Automation pattern statistics for the alert overlay prompt during {}.", operationName));
            if (uiaPatternStats.has_value())
            {
                state.Require(uiaPatternStats->buttonControlCount > 0u,
                              std::format(L"Alert overlay prompt should expose a visible DX button during {}.", operationName));
                state.Require(uiaPatternStats->invokePatternCount > 0u,
                              std::format(L"Alert overlay prompt should expose live InvokePattern support during {}.", operationName));
            }

            const auto buttonState = CollectVisibleDescendantNamedElementState(overlay, UIA_ButtonControlTypeId);
            state.Require(buttonState.has_value(), std::format(L"Failed to collect a visible DX alert overlay button state during {}.", operationName));
            if (buttonState.has_value())
            {
                state.Require(! buttonState->name.empty(),
                              std::format(L"Alert overlay prompt visible DX button should expose a stable accessible name during {}.", operationName));
            }

            SendMessageW(overlay, WM_SYSCHAR, static_cast<WPARAM>(state.failure.empty() ? mnemonic : L'c'), 0);
            promptResult.closed = WaitForWindowClosed(overlay, SelfTest::Scale(3000ms));
        });

        HostPromptRequest request{};
        request.version       = 1u;
        request.sizeBytes     = sizeof(request);
        request.scope         = HOST_ALERT_SCOPE_APPLICATION;
        request.severity      = HOST_ALERT_INFO;
        request.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
        request.targetWindow  = nullptr;
        request.title         = L"Alert overlay prompt access-key self-test";
        request.message       = L"Exercise the real host-services prompt through the DX alert overlay access-key route.";
        request.defaultResult = HOST_PROMPT_RESULT_OK;

        promptResult.hr = HostShowPrompt(request, nullptr, &promptResult.result);
        automation.join();

        state.Require(SUCCEEDED(promptResult.hr),
                      std::format(L"Host prompt should succeed for {}; hr=0x{:08X}.", operationName, static_cast<unsigned int>(promptResult.hr)));
        state.Require(promptResult.result == expectedResult,
                      std::format(L"Alert overlay prompt returned unexpected result for {}; saw {} expected {}.",
                                  operationName,
                                  static_cast<unsigned int>(promptResult.result),
                                  static_cast<unsigned int>(expectedResult)));
        state.Require(promptResult.closed, std::format(L"Alert overlay prompt did not close cleanly after {}.", operationName));
    };

    runPrompt(L'c', HOST_PROMPT_RESULT_CANCEL, L"DX Cancel access key");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(getOverlayWindow() == nullptr, L"Alert overlay prompt should fully dismiss after the DX Cancel access key before reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    runPrompt(L'o', HOST_PROMPT_RESULT_OK, L"DX OK access key");
    return state.failure.empty();
}

[[nodiscard]] bool TestAppPromptLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    constexpr std::wstring_view kAlertOverlayWindowClassName = L"RedSalamander.AlertOverlayWindow";

    const auto getOverlayWindow = [&]() noexcept -> HWND { return FindVisibleDescendantWindowByClass(mainWindow, kAlertOverlayWindowClassName); };

    const bool autoAcceptPromptsBefore = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(false);
    const auto restoreAutoAcceptPrompts = wil::scope_exit([&]() noexcept { HostSetAutoAcceptPrompts(autoAcceptPromptsBefore); });

    const auto closeExistingOverlay = [&]() noexcept
    {
        if (const HWND existing = getOverlayWindow(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_KEYDOWN, VK_ESCAPE, 0);
            PostMessageW(existing, WM_KEYUP, VK_ESCAPE, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    closeExistingOverlay();

    const auto runPromptCycle = [&](size_t iteration, WPARAM closeKey, HostPromptResult expectedResult, std::wstring_view action) noexcept
    {
        struct PromptCycleResult final
        {
            HRESULT hr              = E_FAIL;
            HostPromptResult result = HOST_PROMPT_RESULT_NONE;
            HWND overlay            = nullptr;
            bool closedCleanly      = false;
        };

        PromptCycleResult cycleResult{};

        std::wstring title   = std::format(L"Alert overlay prompt churn {}", iteration + 1u);
        std::wstring message = std::format(L"Exercise the real host-services alert overlay prompt churn path ({}) iteration {}.", action, iteration + 1u);

        std::jthread automation([&](std::stop_token) noexcept
        {
            const HWND overlay  = WaitForWindow(getOverlayWindow, SelfTest::Scale(5000ms));
            cycleResult.overlay = overlay;
            state.Require(overlay != nullptr && IsWindow(overlay) != FALSE,
                          std::format(L"Alert overlay prompt did not open on churn iteration {}.", iteration + 1u));
            if (! overlay || IsWindow(overlay) == FALSE)
            {
                return;
            }

            state.Require(IsOwnedBy(overlay, mainWindow),
                          std::format(L"Alert overlay prompt should stay owned by the main window root on churn iteration {}.", iteration + 1u));
            state.Require(CountVisibleChildWindows(overlay) == 0u,
                          std::format(L"Alert overlay prompt should keep visible child fallback at zero on churn iteration {}.", iteration + 1u));
            state.Require(WindowExposesUiaProvider(overlay),
                          std::format(L"Alert overlay prompt should keep answering WM_GETOBJECT on churn iteration {}.", iteration + 1u));

            RECT overlayRect{};
            state.Require(GetClientRect(overlay, &overlayRect) != FALSE && overlayRect.right > overlayRect.left && overlayRect.bottom > overlayRect.top,
                          std::format(L"Alert overlay prompt should expose a non-empty client area on churn iteration {}.", iteration + 1u));

            const auto uiaPatternStats = WaitForVisiblePromptButtonStats(overlay, SelfTest::Scale(3000ms));
            state.Require(
                uiaPatternStats.has_value(),
                std::format(L"Failed to collect live UI Automation pattern statistics for the alert overlay prompt on churn iteration {}.", iteration + 1u));
            if (uiaPatternStats.has_value())
            {
                state.Require(
                    uiaPatternStats->buttonControlCount > 0u,
                    std::format(L"Alert overlay prompt should expose a visible UI Automation button descendant on churn iteration {}.", iteration + 1u));
                state.Require(
                    uiaPatternStats->invokePatternCount > 0u,
                    std::format(L"Alert overlay prompt should expose live UI Automation InvokePattern support on churn iteration {}.", iteration + 1u));
            }

            const auto buttonState = CollectVisibleDescendantNamedElementState(overlay, UIA_ButtonControlTypeId);
            state.Require(buttonState.has_value(),
                          std::format(L"Failed to collect a visible DX alert overlay button state on churn iteration {}.", iteration + 1u));
            if (buttonState.has_value())
            {
                state.Require(
                    ! buttonState->name.empty(),
                    std::format(L"Alert overlay prompt visible DX button should expose a stable accessible name on churn iteration {}.", iteration + 1u));
            }

            const auto textState = CollectVisibleDescendantNamedElementState(overlay, UIA_TextControlTypeId);
            state.Require(textState.has_value(),
                          std::format(L"Failed to collect a visible DX alert overlay text state on churn iteration {}.", iteration + 1u));
            if (textState.has_value())
            {
                state.Require(
                    ! textState->name.empty(),
                    std::format(L"Alert overlay prompt visible DX text surface should expose a stable accessible name on churn iteration {}.", iteration + 1u));
            }

            SendMessageW(overlay, WM_KEYDOWN, state.failure.empty() ? closeKey : VK_ESCAPE, 0);
            SendMessageW(overlay, WM_KEYUP, state.failure.empty() ? closeKey : VK_ESCAPE, 0);
            cycleResult.closedCleanly = WaitForWindowClosed(overlay, SelfTest::Scale(3000ms));
        });

        HostPromptRequest request{};
        request.version       = 1u;
        request.sizeBytes     = sizeof(request);
        request.scope         = HOST_ALERT_SCOPE_APPLICATION;
        request.severity      = HOST_ALERT_INFO;
        request.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
        request.targetWindow  = nullptr;
        request.title         = title.c_str();
        request.message       = message.c_str();
        request.defaultResult = HOST_PROMPT_RESULT_OK;

        cycleResult.hr = HostShowPrompt(request, nullptr, &cycleResult.result);
        automation.join();

        state.Require(
            SUCCEEDED(cycleResult.hr),
            std::format(L"Host prompt should succeed on churn iteration {}; hr=0x{:08X}.", iteration + 1u, static_cast<unsigned int>(cycleResult.hr)));
        state.Require(cycleResult.result == expectedResult,
                      std::format(L"Alert overlay prompt returned unexpected result on churn iteration {}; saw {} expected {}.",
                                  iteration + 1u,
                                  static_cast<unsigned int>(cycleResult.result),
                                  static_cast<unsigned int>(expectedResult)));
        state.Require(cycleResult.closedCleanly, std::format(L"Alert overlay prompt did not close cleanly on churn iteration {}.", iteration + 1u));
        state.Require(getOverlayWindow() == nullptr,
                      std::format(L"Alert overlay prompt should not leave a visible overlay behind after churn iteration {}.", iteration + 1u));
        return state.failure.empty();
    };

    constexpr size_t kIterationCount = 12u;
    for (size_t iteration = 0; iteration < kIterationCount; ++iteration)
    {
        const bool accept = (iteration % 2u) == 0u;
        if (! runPromptCycle(
                iteration, accept ? VK_RETURN : VK_ESCAPE, accept ? HOST_PROMPT_RESULT_OK : HOST_PROMPT_RESULT_CANCEL, accept ? L"accept" : L"cancel"))
        {
            return false;
        }
    }

    return state.failure.empty();
}

struct FatalErrorReadableSurfaceProbe final
{
    std::optional<UiaDescendantPatternStats> uiaPatternStats;
    std::optional<UiaReadableTextState> readableTextState;
    std::optional<UiaNamedElementState> buttonState;
};

[[nodiscard]] FatalErrorReadableSurfaceProbe CollectFatalErrorReadableSurfaceProbe(HWND dialog) noexcept
{
    FatalErrorReadableSurfaceProbe probe{};
    probe.uiaPatternStats    = CollectVisibleUiaDescendantPatternStats(dialog);
    probe.readableTextState  = CollectVisibleDescendantReadableTextState(dialog, UIA_EditControlTypeId);
    probe.buttonState        = CollectVisibleDescendantNamedElementState(dialog, UIA_ButtonControlTypeId);
    return probe;
}

[[nodiscard]] bool FatalErrorReadableSurfaceProbeSettled(const FatalErrorReadableSurfaceProbe& probe) noexcept
{
    return probe.uiaPatternStats.has_value() && probe.uiaPatternStats->visibleElementCount > 0u && probe.uiaPatternStats->buttonControlCount > 0u &&
           probe.uiaPatternStats->invokePatternCount > 0u &&
           (probe.uiaPatternStats->valuePatternCount > 0u || probe.uiaPatternStats->textPatternCount > 0u) && probe.readableTextState.has_value() &&
           probe.buttonState.has_value();
}

[[nodiscard]] FatalErrorReadableSurfaceProbe WaitForFatalErrorReadableSurfaceProbe(HWND dialog) noexcept
{
    using namespace std::chrono_literals;

    FatalErrorReadableSurfaceProbe probe{};
    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        probe = CollectFatalErrorReadableSurfaceProbe(dialog);
        if (FatalErrorReadableSurfaceProbeSettled(probe))
        {
            return probe;
        }

        std::this_thread::sleep_for(20ms);
    }

    return CollectFatalErrorReadableSurfaceProbe(dialog);
}

[[nodiscard]] bool UiaReadableTextIsReadOnlyOrUnknown(const UiaReadableTextState& state) noexcept
{
    return ! state.readOnlyKnown || state.isReadOnly;
}

[[nodiscard]] bool UiaReadableTextContains(const UiaReadableTextState& state, std::wstring_view expectedText)
{
    return NormalizeComparisonNewlines(state.value).find(NormalizeComparisonNewlines(expectedText)) != std::wstring::npos;
}

[[nodiscard]] bool TestFatalErrorDialogUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetFatalErrorDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing fatal-error dialog did not close before DXUI validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeState final
    {
        std::atomic<bool> sawDialog{false};
        std::atomic<bool> ownedByMainWindow{false};
        std::atomic<bool> exposesProvider{false};
        std::atomic<bool> closed{false};
        std::atomic<size_t> visibleChildCount{std::numeric_limits<size_t>::max()};
        FatalErrorReadableSurfaceProbe surfaceProbe;

        ProbeState()                             = default;
        ProbeState(const ProbeState&)            = delete;
        ProbeState& operator=(const ProbeState&) = delete;
        ProbeState(ProbeState&&)                 = delete;
        ProbeState& operator=(ProbeState&&)      = delete;
    };

    const auto runFatalSurfacePass = [&](std::wstring_view label) noexcept -> bool
    {
        ProbeState probe{};
        std::jthread closeThread([&](std::stop_token) noexcept
        {
            const HWND dialog = WaitForWindow([] noexcept { return GetFatalErrorDialogHandle(); }, SelfTest::Scale(5000ms));
            if (! dialog || IsWindow(dialog) == FALSE)
            {
                return;
            }

            probe.sawDialog.store(true, std::memory_order_release);
            probe.ownedByMainWindow.store(IsOwnedBy(dialog, mainWindow), std::memory_order_release);
            probe.exposesProvider.store(WindowExposesUiaProvider(dialog), std::memory_order_release);
            probe.visibleChildCount.store(CountVisibleChildWindows(dialog), std::memory_order_release);
            probe.surfaceProbe = WaitForFatalErrorReadableSurfaceProbe(dialog);

            PostMessageW(dialog, WM_CLOSE, 0, 0);
            probe.closed.store(WaitForWindowClosed(dialog, SelfTest::Scale(5000ms)), std::memory_order_release);
        });

        DebugShowFatalErrorDialog(
            mainWindow, L"Fatal Error DXUI Self-Test", L"Sample fatal error text for DXUI validation.\r\nSecond line for multiline coverage.");
        closeThread.join();

        state.Require(probe.sawDialog.load(std::memory_order_acquire), std::format(L"Fatal-error dialog did not open during {}.", label));
        state.Require(probe.ownedByMainWindow.load(std::memory_order_acquire),
                      std::format(L"Fatal-error dialog should be owned by the main window during {}.", label));
        state.Require(probe.visibleChildCount.load(std::memory_order_acquire) == 0u,
                      std::format(L"Fatal-error dialog should not expose visible child-control fallback during {}.", label));
        state.Require(probe.exposesProvider.load(std::memory_order_acquire),
                      std::format(L"Fatal-error dialog should answer WM_GETOBJECT with a UI Automation provider during {}.", label));
        state.Require(probe.surfaceProbe.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the fatal-error dialog during {}.", label));
        if (probe.surfaceProbe.uiaPatternStats.has_value())
        {
            state.Require(probe.surfaceProbe.uiaPatternStats->visibleElementCount > 0u, L"Fatal-error dialog should expose visible UI Automation descendants.");
            state.Require(probe.surfaceProbe.uiaPatternStats->buttonControlCount > 0u,
                          L"Fatal-error dialog should expose a visible UI Automation button descendant.");
            state.Require(probe.surfaceProbe.uiaPatternStats->invokePatternCount > 0u,
                          L"Fatal-error dialog should expose live UI Automation InvokePattern support for the DX footer action.");
            state.Require(probe.surfaceProbe.uiaPatternStats->valuePatternCount > 0u || probe.surfaceProbe.uiaPatternStats->textPatternCount > 0u,
                          L"Fatal-error dialog should expose live UI Automation readable text-pattern support for the DX message surface.");
        }

        state.Require(probe.surfaceProbe.readableTextState.has_value(),
                      std::format(L"Failed to collect UI Automation readable text state for the fatal-error message surface during {}.", label));
        if (probe.surfaceProbe.readableTextState.has_value())
        {
            state.Require(UiaReadableTextIsReadOnlyOrUnknown(probe.surfaceProbe.readableTextState.value()), L"Fatal-error message surface should remain read-only.");
            state.Require(UiaReadableTextContains(probe.surfaceProbe.readableTextState.value(), L"Sample fatal error text for DXUI validation."),
                          std::format(L"Fatal-error dialog readable text should expose the message text during {}; saw '{}'.",
                                      label,
                                      probe.surfaceProbe.readableTextState->value));
        }
        state.Require(probe.surfaceProbe.buttonState.has_value(), std::format(L"Fatal-error dialog should expose a visible DX footer button during {}.", label));
        if (probe.surfaceProbe.buttonState.has_value())
        {
            state.Require(! probe.surfaceProbe.buttonState->name.empty(),
                          std::format(L"Fatal-error dialog visible DX footer button should expose a stable accessible name during {}.", label));
        }
        state.Require(probe.closed.load(std::memory_order_acquire), std::format(L"Fatal-error dialog did not close cleanly during {}.", label));
        return state.failure.empty();
    };

    state.Require(runFatalSurfacePass(L"the initial fatal-error surface pass"),
                  L"Fatal-error dialog should expose the expected DX surface on the initial pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runFatalSurfacePass(L"the reopened fatal-error surface pass"), L"Fatal-error dialog should expose the same DX surface after reopen.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFatalErrorDialogLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetFatalErrorDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing fatal-error dialog did not close before live interaction validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring message = L"Fatal error live interaction text.\r\nSecond line for close-button coverage.";

    struct ProbeState final
    {
        bool sawDialog           = false;
        bool ownedByMainWindow   = false;
        bool exposesProvider     = false;
        bool invokedButton       = false;
        bool closed              = false;
        size_t visibleChildCount = std::numeric_limits<size_t>::max();
        FatalErrorReadableSurfaceProbe surfaceProbe;
    };

    const auto runLivePass = [&](std::wstring_view context) noexcept
    {
        ProbeState probe{};
        std::jthread worker([&](std::stop_token) noexcept
        {
            const HWND dialog = WaitForWindow([] noexcept { return GetFatalErrorDialogHandle(); }, SelfTest::Scale(5000ms));
            if (! dialog || IsWindow(dialog) == FALSE)
            {
                return;
            }

            probe.sawDialog         = true;
            probe.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);
            probe.visibleChildCount = CountVisibleChildWindows(dialog);
            probe.exposesProvider   = WindowExposesUiaProvider(dialog);
            probe.surfaceProbe      = WaitForFatalErrorReadableSurfaceProbe(dialog);

            probe.invokedButton = InvokeVisibleDescendantByName(dialog, UIA_ButtonControlTypeId, L"");
            if (! probe.invokedButton && IsWindow(dialog) != FALSE)
            {
                PostMessageW(dialog, WM_CLOSE, 0, 0);
            }
            probe.closed = WaitForWindowClosed(dialog, SelfTest::Scale(5000ms));
        });

        DebugShowFatalErrorDialog(mainWindow, L"Fatal Error DXUI Live Interaction", message.c_str());
        worker.join();

        state.Require(probe.sawDialog, std::format(L"Fatal-error dialog did not open for {}.", context));
        state.Require(probe.ownedByMainWindow, std::format(L"Fatal-error dialog should remain owned by the main window during {}.", context));
        state.Require(probe.visibleChildCount == 0u, std::format(L"Fatal-error dialog should not expose visible child-control fallback during {}.", context));
        state.Require(probe.exposesProvider, std::format(L"Fatal-error dialog should answer WM_GETOBJECT during {}.", context));
        state.Require(probe.surfaceProbe.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation stats for fatal-error dialog during {}.", context));
        if (probe.surfaceProbe.uiaPatternStats.has_value())
        {
            state.Require(probe.surfaceProbe.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Fatal-error dialog should expose a visible DX footer button during {}.", context));
            state.Require(probe.surfaceProbe.uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Fatal-error dialog should expose InvokePattern during {}.", context));
            state.Require(probe.surfaceProbe.uiaPatternStats->valuePatternCount > 0u || probe.surfaceProbe.uiaPatternStats->textPatternCount > 0u,
                          std::format(L"Fatal-error dialog should expose readable text-pattern support during {}.", context));
        }

        state.Require(probe.surfaceProbe.readableTextState.has_value(),
                      std::format(L"Failed to collect fatal-error readable text state during {}.", context));
        if (probe.surfaceProbe.readableTextState.has_value())
        {
            state.Require(UiaReadableTextIsReadOnlyOrUnknown(probe.surfaceProbe.readableTextState.value()),
                          std::format(L"Fatal-error message surface should remain read-only during {}.", context));
            state.Require(
                UiaReadableTextContains(probe.surfaceProbe.readableTextState.value(), message),
                std::format(L"Fatal-error dialog readable text should expose the live message text during {}; saw '{}'.",
                            context,
                            probe.surfaceProbe.readableTextState->value));
        }
        state.Require(probe.invokedButton, std::format(L"Failed to invoke the visible DX footer button on fatal-error dialog during {}.", context));
        state.Require(probe.closed, std::format(L"Fatal-error dialog did not close after live UIA InvokePattern interaction during {}.", context));
        return state.failure.empty();
    };

    state.Require(runLivePass(L"initial live interaction"), L"Initial fatal-error dialog live DX validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runLivePass(L"reopened live interaction"), L"Reopened fatal-error dialog live DX validation failed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFatalErrorDialogAccessKeyRoutesOk(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetFatalErrorDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)), L"Existing fatal-error dialog did not close before access-key validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeState final
    {
        bool sawDialog           = false;
        bool exposesProvider     = false;
        bool accessKeyClosed     = false;
        bool fallbackClosed      = false;
        size_t visibleChildCount = std::numeric_limits<size_t>::max();
    } probe;

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND dialog = WaitForWindow([] noexcept { return GetFatalErrorDialogHandle(); }, SelfTest::Scale(5000ms));
        if (! dialog || IsWindow(dialog) == FALSE)
        {
            return;
        }

        probe.sawDialog         = true;
        probe.visibleChildCount = CountVisibleChildWindows(dialog);
        probe.exposesProvider   = WindowExposesUiaProvider(dialog);

        SendMessageW(dialog, WM_SYSCHAR, static_cast<WPARAM>(L'o'), 0);
        probe.accessKeyClosed = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
        if (! probe.accessKeyClosed && IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            probe.fallbackClosed = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
        }
        else
        {
            probe.fallbackClosed = true;
        }
    });

    DebugShowFatalErrorDialog(mainWindow, L"Fatal Error DXUI access-key", L"Fatal dialog access-key validation message.");
    worker.join();

    state.Require(probe.sawDialog, L"Fatal-error dialog did not open for access-key validation.");
    state.Require(probe.visibleChildCount == 0u, L"Fatal-error dialog should not expose visible child-control fallback during access-key validation.");
    state.Require(probe.exposesProvider, L"Fatal-error dialog should answer WM_GETOBJECT during access-key validation.");
    state.Require(probe.accessKeyClosed, L"Fatal-error dialog did not close after the DX OK access key.");
    state.Require(probe.fallbackClosed, L"Fatal-error dialog did not close after access-key validation cleanup.");
    state.Require(GetFatalErrorDialogHandle() == nullptr || IsWindow(GetFatalErrorDialogHandle()) == FALSE,
                  L"Fatal-error dialog should not remain open after the DX OK access key.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFatalErrorDialogEnterAndEscapeRouteOk(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExisting = [&]() noexcept
    {
        if (const HWND existing = GetFatalErrorDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };
    closeExisting();

    const auto runCycle = [&](std::wstring_view context, const WPARAM key) noexcept
    {
        struct ProbeState final
        {
            bool sawDialog           = false;
            bool exposesProvider     = false;
            bool keyClosed           = false;
            bool fallbackClosed      = false;
            size_t visibleChildCount = std::numeric_limits<size_t>::max();
        } probe;

        std::jthread worker([&](std::stop_token) noexcept
        {
            const HWND dialog = WaitForWindow([] noexcept { return GetFatalErrorDialogHandle(); }, SelfTest::Scale(5000ms));
            if (! dialog || IsWindow(dialog) == FALSE)
            {
                return;
            }

            probe.sawDialog         = true;
            probe.visibleChildCount = CountVisibleChildWindows(dialog);
            probe.exposesProvider   = WindowExposesUiaProvider(dialog);

            SendMessageW(dialog, WM_KEYDOWN, key, 0);
            SendMessageW(dialog, WM_KEYUP, key, 0);
            probe.keyClosed = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
            if (! probe.keyClosed && IsWindow(dialog) != FALSE)
            {
                PostMessageW(dialog, WM_CLOSE, 0, 0);
                probe.fallbackClosed = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
            }
            else
            {
                probe.fallbackClosed = true;
            }
        });

        DebugShowFatalErrorDialog(mainWindow, L"Fatal Error DXUI keyboard-flow", L"Fatal dialog keyboard-flow validation message.");
        worker.join();

        state.Require(probe.sawDialog, std::format(L"Fatal-error dialog did not open for {}.", context));
        state.Require(probe.visibleChildCount == 0u, std::format(L"Fatal-error dialog should not expose visible child-control fallback during {}.", context));
        state.Require(probe.exposesProvider, std::format(L"Fatal-error dialog should answer WM_GETOBJECT during {}.", context));
        state.Require(probe.keyClosed, std::format(L"Fatal-error dialog did not close after {}.", context));
        state.Require(probe.fallbackClosed, std::format(L"Fatal-error dialog did not close after {} cleanup.", context));
        return state.failure.empty();
    };

    state.Require(runCycle(L"Enter default-button routing", VK_RETURN), L"Fatal-error dialog Enter default-button validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runCycle(L"Escape cancel routing", VK_ESCAPE), L"Fatal-error dialog Escape cancel validation failed.");
    state.Require(GetFatalErrorDialogHandle() == nullptr || IsWindow(GetFatalErrorDialogHandle()) == FALSE,
                  L"Fatal-error dialog should not remain open after Enter/Escape keyboard-flow validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFatalErrorDialogLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingFatalError = [&]() noexcept
    {
        if (const HWND existing = GetFatalErrorDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    closeExistingFatalError();

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const std::wstring caption = std::format(L"Fatal Error DXUI churn {}", cycle);
        const std::wstring message = std::format(L"Fatal error churn message {}.\r\nSecond line {}.", cycle, cycle);

        struct ProbeState final
        {
            bool sawDialog           = false;
            bool ownedByMainWindow   = false;
            bool exposesProvider     = false;
            bool closed              = false;
            size_t visibleChildCount = std::numeric_limits<size_t>::max();
            FatalErrorReadableSurfaceProbe surfaceProbe;
        } probe;

        std::jthread worker([&](std::stop_token) noexcept
        {
            const HWND dialog = WaitForWindow([] noexcept { return GetFatalErrorDialogHandle(); }, SelfTest::Scale(5000ms));
            if (! dialog || IsWindow(dialog) == FALSE)
            {
                return;
            }

            probe.sawDialog         = true;
            probe.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);
            probe.visibleChildCount = CountVisibleChildWindows(dialog);
            probe.exposesProvider   = WindowExposesUiaProvider(dialog);
            probe.surfaceProbe      = WaitForFatalErrorReadableSurfaceProbe(dialog);

            PostMessageW(dialog, WM_CLOSE, 0, 0);
            probe.closed = WaitForWindowClosed(dialog, SelfTest::Scale(5000ms));
        });

        DebugShowFatalErrorDialog(mainWindow, caption.c_str(), message.c_str());
        worker.join();

        state.Require(probe.sawDialog, std::format(L"Fatal-error dialog did not open during cycle {}.", cycle));
        state.Require(probe.ownedByMainWindow, std::format(L"Fatal-error dialog should be owned by the main window during cycle {}.", cycle));
        state.Require(probe.visibleChildCount == 0u,
                      std::format(L"Fatal-error dialog should not expose visible child-control fallback during cycle {}.", cycle));
        state.Require(probe.exposesProvider, std::format(L"Fatal-error dialog should answer WM_GETOBJECT during cycle {}.", cycle));

        state.Require(probe.surfaceProbe.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the fatal-error dialog during cycle {}.", cycle));
        if (probe.surfaceProbe.uiaPatternStats.has_value())
        {
            state.Require(probe.surfaceProbe.uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Fatal-error dialog should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(probe.surfaceProbe.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Fatal-error dialog should expose a visible UI Automation button descendant during cycle {}.", cycle));
            state.Require(probe.surfaceProbe.uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Fatal-error dialog should expose live UI Automation InvokePattern support during cycle {}.", cycle));
            state.Require(probe.surfaceProbe.uiaPatternStats->valuePatternCount > 0u || probe.surfaceProbe.uiaPatternStats->textPatternCount > 0u,
                          std::format(L"Fatal-error dialog should expose live UI Automation readable text-pattern support during cycle {}.", cycle));
        }

        state.Require(probe.surfaceProbe.readableTextState.has_value(),
                      std::format(L"Failed to collect UI Automation readable text state for the fatal-error dialog during cycle {}.", cycle));
        if (probe.surfaceProbe.readableTextState.has_value())
        {
            state.Require(UiaReadableTextIsReadOnlyOrUnknown(probe.surfaceProbe.readableTextState.value()),
                          std::format(L"Fatal-error message surface should remain read-only during cycle {}.", cycle));
            state.Require(! probe.surfaceProbe.readableTextState->name.empty(),
                          std::format(L"Fatal-error message surface should expose a stable accessible name during cycle {}.", cycle));
            state.Require(UiaReadableTextContains(probe.surfaceProbe.readableTextState.value(), message),
                          std::format(L"Fatal-error dialog should expose the live message text during cycle {}; saw '{}'.",
                                      cycle,
                                      probe.surfaceProbe.readableTextState->value));
        }

        state.Require(probe.surfaceProbe.buttonState.has_value(),
                      std::format(L"Failed to collect a visible DX footer-button state for the fatal-error dialog during cycle {}.", cycle));
        if (probe.surfaceProbe.buttonState.has_value())
        {
            state.Require(! probe.surfaceProbe.buttonState->name.empty(),
                          std::format(L"Fatal-error dialog visible DX footer button should expose a stable accessible name during cycle {}.", cycle));
        }

        state.Require(probe.closed, std::format(L"Fatal-error dialog did not close cleanly during cycle {}.", cycle));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(GetFatalErrorDialogHandle() == nullptr || IsWindow(GetFatalErrorDialogHandle()) == FALSE,
                  L"Fatal-error dialog should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFatalErrorDialogLongRunScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetFatalErrorDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)),
                      L"Existing fatal-error dialog did not close before long-run scrolling validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring message = L"Fatal dialog multiline scrolling validation header.";
    for (size_t index = 0u; index < 96u; ++index)
    {
        message.append(std::format(L"\r\nLine {:03}: fatal dialog read-only multiline viewport pressure with wrapped detail segment alpha beta gamma delta "
                                   L"epsilon zeta eta theta iota kappa lambda.",
                                   index));
    }

    struct WorkerResult final
    {
        bool sawDialog         = false;
        bool ownedByMainWindow = false;
        bool scrolledDown      = false;
        bool restoredToTop     = false;
        bool closed            = false;
        std::wstring failureMessage;
        FatalErrorDialogDebugSnapshot baselineSnapshot{};
        FatalErrorDialogDebugSnapshot finalSnapshot{};
        std::optional<UiaReadableTextState> readableTextState;
        std::optional<UiaNamedElementState> buttonState;
    } workerResult;

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND dialog = WaitForWindow([] noexcept { return GetFatalErrorDialogHandle(); }, SelfTest::Scale(5000ms));
        if (! dialog || IsWindow(dialog) == FALSE)
        {
            workerResult.failureMessage = L"Fatal-error dialog did not open for long-run scrolling validation.";
            return;
        }

        workerResult.sawDialog         = true;
        workerResult.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

        const auto waitForSnapshot = [&](const auto& predicate, FatalErrorDialogDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                outSnapshot = {};
                if (DebugGetFatalErrorDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetFatalErrorDialogSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        auto setFailure = [&](std::wstring messageText) noexcept
        {
            if (workerResult.failureMessage.empty())
            {
                workerResult.failureMessage = std::move(messageText);
            }
        };

        FatalErrorDialogDebugSnapshot snapshot{};
        if (! waitForSnapshot(
                [&](const FatalErrorDialogDebugSnapshot& value) noexcept
        {
            return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.bodyVisibleLineCount > 0u &&
                   value.bodyTotalLineCount > value.bodyVisibleLineCount && value.bodyCanScrollVertically && value.resizeFailureCount == 0u &&
                   value.messageText.find(L"Line 095:") != std::wstring::npos;
        },
                snapshot))
        {
            setFailure(L"Fatal-error dialog did not expose a scrollable DX read-only multiline surface.");
        }

        if (workerResult.failureMessage.empty())
        {
            RECT rect{};
            if (! GetWindowRect(dialog, &rect))
            {
                setFailure(L"Failed to capture fatal-error dialog rect for long-run scrolling validation.");
            }
            else
            {
                constexpr int kTargetWidth  = 420;
                constexpr int kTargetHeight = 240;
                if (! SetWindowPos(dialog, nullptr, 0, 0, kTargetWidth, kTargetHeight, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE))
                {
                    setFailure(L"Failed to resize fatal-error dialog for long-run scrolling validation.");
                }
                else if (! waitForSnapshot(
                             [&](const FatalErrorDialogDebugSnapshot& value) noexcept
                {
                    return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.bodyVisibleLineCount > 0u &&
                           value.bodyTotalLineCount > value.bodyVisibleLineCount && value.bodyCanScrollVertically && value.resizeFailureCount == 0u;
                },
                             snapshot))
                {
                    setFailure(L"Fatal-error dialog did not settle after the narrow-width scrolling-validation resize.");
                }
            }
        }

        workerResult.baselineSnapshot             = snapshot;
        size_t previousFirstVisibleLine           = snapshot.bodyFirstVisibleLine;
        const uint64_t baselineResizeFailureCount = snapshot.resizeFailureCount;
        for (size_t chunk = 0u; chunk < 8u && workerResult.failureMessage.empty(); ++chunk)
        {
            if (! DebugScrollFatalErrorDialogByWheelDetents(-4))
            {
                setFailure(std::format(L"Fatal-error dialog did not accept long-run scroll chunk {}.", chunk));
                break;
            }

            if (! waitForSnapshot(
                    [&](const FatalErrorDialogDebugSnapshot& value) noexcept
            {
                return value.bodyFirstVisibleLine > previousFirstVisibleLine && value.visibleChildWindowCount == 0u && value.bodyCanScrollVertically &&
                       value.bodyVisibleLineCount > 0u && value.bodyTotalLineCount == workerResult.baselineSnapshot.bodyTotalLineCount &&
                       value.resizeFailureCount == baselineResizeFailureCount;
            },
                    snapshot))
            {
                setFailure(std::format(L"Fatal-error dialog did not advance its multiline viewport during long-run scroll chunk {}.", chunk));
                break;
            }

            previousFirstVisibleLine = snapshot.bodyFirstVisibleLine;
        }

        workerResult.finalSnapshot = snapshot;
        workerResult.scrolledDown  = previousFirstVisibleLine > workerResult.baselineSnapshot.bodyFirstVisibleLine;
        if (workerResult.failureMessage.empty() && ! workerResult.scrolledDown)
        {
            setFailure(L"Fatal-error dialog should scroll away from the top during long-run validation.");
        }

        if (workerResult.failureMessage.empty())
        {
            workerResult.readableTextState = CollectVisibleDescendantReadableTextState(dialog, UIA_EditControlTypeId);
            workerResult.buttonState       = CollectVisibleDescendantNamedElementState(dialog, UIA_ButtonControlTypeId);
        }

        while (previousFirstVisibleLine > workerResult.baselineSnapshot.bodyFirstVisibleLine && workerResult.failureMessage.empty())
        {
            if (! DebugScrollFatalErrorDialogByWheelDetents(4))
            {
                setFailure(L"Fatal-error dialog did not accept reverse wheel scrolling back toward the top.");
                break;
            }

            if (! waitForSnapshot(
                    [&](const FatalErrorDialogDebugSnapshot& value) noexcept
            {
                return value.bodyFirstVisibleLine < previousFirstVisibleLine && value.visibleChildWindowCount == 0u &&
                       value.resizeFailureCount == baselineResizeFailureCount;
            },
                    snapshot))
            {
                setFailure(L"Fatal-error dialog did not move back toward the top during reverse scroll normalization.");
                break;
            }

            previousFirstVisibleLine = snapshot.bodyFirstVisibleLine;
        }

        workerResult.finalSnapshot = snapshot;
        workerResult.restoredToTop = previousFirstVisibleLine == workerResult.baselineSnapshot.bodyFirstVisibleLine;
        if (IsWindow(dialog) != FALSE)
        {
            PostMessageW(dialog, WM_CLOSE, 0, 0);
            workerResult.closed = WaitForWindowClosed(dialog, SelfTest::Scale(5000ms));
        }
    });

    DebugShowFatalErrorDialog(mainWindow, L"Fatal Error DXUI Long-Run Scroll", message.c_str());
    worker.join();

    state.Require(workerResult.failureMessage.empty(), workerResult.failureMessage);
    state.Require(workerResult.sawDialog, L"Fatal-error dialog worker did not observe the dialog.");
    state.Require(workerResult.ownedByMainWindow, L"Fatal-error dialog should be owned by the main window during long-run scrolling validation.");
    state.Require(workerResult.baselineSnapshot.usesDxUiHost, L"Fatal-error dialog should expose its DX host during long-run scrolling validation.");
    state.Require(workerResult.baselineSnapshot.visibleChildWindowCount == 0u,
                  L"Fatal-error dialog should not expose visible child-control fallback during long-run scrolling validation.");
    state.Require(workerResult.baselineSnapshot.bodyVisibleLineCount > 0u,
                  L"Fatal-error dialog should expose visible multiline rows before long-run scrolling validation.");
    state.Require(workerResult.baselineSnapshot.bodyTotalLineCount > workerResult.baselineSnapshot.bodyVisibleLineCount,
                  L"Fatal-error dialog should expose vertical overflow before long-run scrolling validation.");
    state.Require(workerResult.scrolledDown, L"Fatal-error dialog should scroll away from the top during long-run validation.");
    state.Require(workerResult.restoredToTop, L"Fatal-error dialog should restore its multiline viewport to the top after reverse wheel scrolling.");
    state.Require(workerResult.finalSnapshot.resizeFailureCount == workerResult.baselineSnapshot.resizeFailureCount,
                  std::format(L"Fatal-error dialog resize-failure count changed during long-run scrolling; baseline={} final={}.",
                              workerResult.baselineSnapshot.resizeFailureCount,
                              workerResult.finalSnapshot.resizeFailureCount));
    state.Require(workerResult.readableTextState.has_value(), L"Failed to collect fatal-error readable text state after long-run scrolling.");
    if (workerResult.readableTextState.has_value())
    {
        state.Require(! workerResult.readableTextState->readOnlyKnown || workerResult.readableTextState->isReadOnly,
                      L"Fatal-error message surface should remain read-only after long-run scrolling.");
        state.Require(workerResult.readableTextState->value.find(L"Line 095:") != std::wstring::npos,
                      L"Fatal-error readable text should still expose the wrapped multiline payload after long-run scrolling.");
    }
    state.Require(workerResult.buttonState.has_value(), L"Fatal-error dialog should keep a visible DX footer button after long-run scrolling.");
    if (workerResult.buttonState.has_value())
    {
        state.Require(! workerResult.buttonState->name.empty(),
                      L"Fatal-error dialog DX footer button should keep a stable accessible name after long-run scrolling.");
    }
    state.Require(workerResult.closed, L"Fatal-error dialog did not close cleanly after long-run scrolling validation.");
    state.Require(GetFatalErrorDialogHandle() == nullptr || IsWindow(GetFatalErrorDialogHandle()) == FALSE,
                  L"Fatal-error dialog should not remain open after long-run scrolling validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestFatalErrorDialogThemeMatrixKeepsMessageLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeExistingFatalError = [&]() noexcept
    {
        if (const HWND existing = GetFatalErrorDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    const std::wstring originalThemeId = g_settings.theme.currentThemeId;
    const auto restoreTheme            = wil::scope_exit([&]() noexcept
    {
        g_settings.theme.currentThemeId = originalThemeId;
        closeExistingFatalError();
    });

    closeExistingFatalError();

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

    struct ThemeExpectation final
    {
        std::wstring label;
        std::wstring themeId;
        ThemeMode mode          = ThemeMode::Dark;
        bool expectRainbow      = false;
        bool expectHighContrast = false;
    };

    const std::array themeCases{
        ThemeExpectation{L"dark", L"builtin/dark", ThemeMode::Dark, false, false},
        ThemeExpectation{L"light", L"builtin/light", ThemeMode::Light, false, false},
        ThemeExpectation{L"rainbow", L"builtin/rainbow", ThemeMode::Rainbow, true, false},
        ThemeExpectation{L"high-contrast", L"builtin/highContrast", ThemeMode::HighContrast, false, true},
    };

    const auto runThemeCase = [&](const ThemeExpectation& themeCase) noexcept
    {
        closeExistingFatalError();
        g_settings.theme.currentThemeId = themeCase.themeId;

        const AppTheme expectedTheme = ResolveAppTheme(themeCase.mode, std::format(L"fatal-error-selftest-theme-{}", themeCase.label));
        const std::wstring caption   = std::format(L"Fatal Error {}", themeCase.label);
        const std::wstring message   = std::format(L"Theme validation for {}.\r\nSecond line keeps the read-only DX message surface active.", themeCase.label);

        struct WorkerResult final
        {
            bool sawDialog         = false;
            bool ownedByMainWindow = false;
            bool closed            = false;
            std::wstring failureMessage;
            FatalErrorDialogDebugSnapshot snapshot{};
            FatalErrorReadableSurfaceProbe surfaceProbe;
        } workerResult;

        std::jthread worker([&](std::stop_token) noexcept
        {
            const HWND dialog = WaitForWindow([]() noexcept { return GetFatalErrorDialogHandle(); }, SelfTest::Scale(5000ms));
            if (! dialog || IsWindow(dialog) == FALSE)
            {
                workerResult.failureMessage = std::format(L"Fatal-error dialog did not open for {} theme validation.", themeCase.label);
                return;
            }

            workerResult.sawDialog         = true;
            workerResult.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

            const auto waitForSnapshot = [&](const auto& predicate) noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{3000});
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    workerResult.snapshot = {};
                    if (DebugGetFatalErrorDialogSnapshot(workerResult.snapshot) && predicate(workerResult.snapshot))
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds{20});
                }

                workerResult.snapshot = {};
                return DebugGetFatalErrorDialogSnapshot(workerResult.snapshot) && predicate(workerResult.snapshot);
            };

            if (! waitForSnapshot([&](const FatalErrorDialogDebugSnapshot& value) noexcept
            {
                return value.usesDxUiHost && value.visibleChildWindowCount == 0u && value.themeDark == expectedTheme.dark &&
                       value.themeHighContrast == expectedTheme.highContrast && value.themeRainbow == expectedTheme.menu.rainbowMode &&
                       value.bodyFillArgb != 0u && value.bodyTextArgb != 0u && value.messageText.find(themeCase.label) != std::wstring::npos;
            }))
            {
                workerResult.failureMessage = std::format(L"Fatal-error dialog did not settle on the expected {} theme surface.", themeCase.label);
                return;
            }

            workerResult.surfaceProbe = WaitForFatalErrorReadableSurfaceProbe(dialog);

            PostMessageW(dialog, WM_CLOSE, 0, 0);
            workerResult.closed = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
            if (! workerResult.closed)
            {
                workerResult.failureMessage = std::format(L"Fatal-error dialog did not close cleanly after {} theme validation.", themeCase.label);
            }
        });

        DebugShowFatalErrorDialog(mainWindow, caption.c_str(), message.c_str());
        worker.join();

        state.Require(workerResult.failureMessage.empty(), workerResult.failureMessage);
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(workerResult.sawDialog, std::format(L"Fatal-error dialog should open for {} theme validation.", themeCase.label));
        state.Require(workerResult.ownedByMainWindow,
                      std::format(L"Fatal-error dialog should be owned by the main window during {} theme validation.", themeCase.label));
        state.Require(workerResult.snapshot.usesDxUiHost,
                      std::format(L"Fatal-error dialog should stay on the DX host path during {} theme validation.", themeCase.label));
        state.Require(workerResult.snapshot.visibleChildWindowCount == 0u,
                      std::format(L"Fatal-error dialog should not expose visible child fallback during {} theme validation.", themeCase.label));
        state.Require(workerResult.snapshot.themeDark == expectedTheme.dark,
                      std::format(L"Fatal-error dialog dark-theme flag mismatch during {} theme validation.", themeCase.label));
        state.Require(workerResult.snapshot.themeRainbow == themeCase.expectRainbow,
                      std::format(L"Fatal-error dialog rainbow-theme flag mismatch during {} theme validation.", themeCase.label));
        state.Require(workerResult.snapshot.themeHighContrast == themeCase.expectHighContrast,
                      std::format(L"Fatal-error dialog high-contrast flag mismatch during {} theme validation.", themeCase.label));
        state.Require(workerResult.snapshot.bodyFillArgb != workerResult.snapshot.bodyTextArgb,
                      std::format(L"Fatal-error dialog message colors collapsed after {} theme selection.", themeCase.label));

        const float minimumContrast = themeCase.expectHighContrast ? 4.5f : 3.0f;
        state.Require(contrastRatio(workerResult.snapshot.bodyFillArgb, workerResult.snapshot.bodyTextArgb) >= minimumContrast,
                      std::format(L"Fatal-error dialog message contrast dropped below {:.1f}:1 after {} theme selection.", minimumContrast, themeCase.label));

        state.Require(workerResult.surfaceProbe.readableTextState.has_value(),
                      std::format(L"Failed to collect fatal-error readable text state for {} theme validation.", themeCase.label));
        if (workerResult.surfaceProbe.readableTextState.has_value())
        {
            state.Require(UiaReadableTextIsReadOnlyOrUnknown(workerResult.surfaceProbe.readableTextState.value()),
                          std::format(L"Fatal-error message surface should stay read-only during {} theme validation.", themeCase.label));
            state.Require(UiaReadableTextContains(workerResult.surfaceProbe.readableTextState.value(), message),
                          std::format(L"Fatal-error dialog should expose the live {} theme message text.", themeCase.label));
        }

        state.Require(workerResult.surfaceProbe.buttonState.has_value(),
                      std::format(L"Failed to collect fatal-error button state for {} theme validation.", themeCase.label));
        if (workerResult.surfaceProbe.buttonState.has_value())
        {
            state.Require(! workerResult.surfaceProbe.buttonState->name.empty(),
                          std::format(L"Fatal-error dialog button should expose a stable accessible name during {} theme validation.", themeCase.label));
        }

        state.Require(workerResult.closed, std::format(L"Fatal-error dialog should close after {} theme validation.", themeCase.label));
    };

    for (const auto& themeCase : themeCases)
    {
        runThemeCase(themeCase);
        if (! state.failure.empty())
        {
            return false;
        }
    }

    state.Require(GetFatalErrorDialogHandle() == nullptr || IsWindow(GetFatalErrorDialogHandle()) == FALSE,
                  L"Fatal-error dialog should not remain open after theme-matrix validation.");
    return state.failure.empty();
}

[[nodiscard]] std::optional<UiaNamedElementState> CollectVisibleSplashTextState(HWND splash, std::wstring_view expectedText) noexcept
{
    wil::com_ptr<IUIAutomationElement> element;
    if (FindMatchingVisibleDescendantElement(splash, UIA_TextControlTypeId, expectedText, element.put()) && element)
    {
        UiaNamedElementState state{};
        state.controlType = UIA_TextControlTypeId;
        state.name.assign(expectedText);
        return state;
    }

    return CollectVisibleDescendantNamedElementState(splash, UIA_TextControlTypeId);
}

[[nodiscard]] bool TestSplashScreenUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto openSplash = [&](std::wstring_view text) noexcept -> HWND
    {
        SplashScreen::SetOwner(mainWindow);
        SplashScreen::IfExistSetText(text.data());
        SplashScreen::BeginDelayedOpen(std::chrono::milliseconds{0}, GetModuleHandleW(nullptr));
        return WaitForWindow([] noexcept { return SplashScreen::GetHwnd(); }, SelfTest::Scale(5000ms));
    };
    const auto closeSplashWindow = [&](HWND splash, std::wstring_view label) noexcept -> bool
    {
        Trace(std::format(L"splash_surface: requesting close during '{}'", label));
        SplashScreen::RequestCloseIfExist();

        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const bool targetClosed                         = ! splash || IsWindow(splash) == FALSE;
            const SplashScreen::DebugSnapshot debugSnapshot = SplashScreen::DebugGetSnapshot();
            if (targetClosed && ! debugSnapshot.threadStarted && ! debugSnapshot.hasHwnd)
            {
                Trace(std::format(L"splash_surface: close completed during '{}' stage={} lastError={} comHr=0x{:08X}",
                                  label,
                                  debugSnapshot.stage,
                                  debugSnapshot.lastError,
                                  static_cast<unsigned long>(debugSnapshot.comHr)));
                return true;
            }

            std::this_thread::sleep_for(10ms);
        }

        const SplashScreen::DebugSnapshot debugSnapshot = SplashScreen::DebugGetSnapshot();
        state.Require(false,
                      std::format(L"Splash window did not close cleanly during {}; stage={} lastError={} comHr=0x{:08X} threadStarted={} hasHwnd={}.",
                                  label,
                                  debugSnapshot.stage,
                                  debugSnapshot.lastError,
                                  static_cast<unsigned long>(debugSnapshot.comHr),
                                  debugSnapshot.threadStarted ? L"yes" : L"no",
                                  debugSnapshot.hasHwnd ? L"yes" : L"no"));
        return false;
    };
    HWND splash            = nullptr;
    const auto closeSplash = wil::scope_exit([&]() noexcept { static_cast<void>(closeSplashWindow(splash, L"cleanup")); });

    const HWND existing = SplashScreen::GetHwnd();
    if (! closeSplashWindow(existing, L"initial cleanup before DXUI validation") || ! state.failure.empty())
    {
        return false;
    }

    const auto validateSplashSurface = [&](std::wstring_view label) noexcept -> bool
    {
        splash                                          = openSplash(L"Splash DXUI self-test");
        const SplashScreen::DebugSnapshot debugSnapshot = SplashScreen::DebugGetSnapshot();
        state.Require(splash != nullptr && IsWindow(splash) != FALSE,
                      std::format(L"Splash window did not open during {}. stage={} lastError={} comHr=0x{:08X} threadStarted={} hasHwnd={}.",
                                  label,
                                  debugSnapshot.stage,
                                  debugSnapshot.lastError,
                                  static_cast<unsigned long>(debugSnapshot.comHr),
                                  debugSnapshot.threadStarted ? L"yes" : L"no",
                                  debugSnapshot.hasHwnd ? L"yes" : L"no"));
        if (! splash || IsWindow(splash) == FALSE)
        {
            return false;
        }

        state.Require(IsOwnedBy(splash, mainWindow), L"Splash window should be owned by the main window.");
        state.Require(CountVisibleChildWindows(splash) == 0u, L"Splash window should not expose visible child-control fallback.");
        state.Require(WindowExposesUiaProvider(splash), L"Splash window should answer WM_GETOBJECT with a UI Automation provider.");

        Trace(std::format(L"splash_surface: collecting UIA pattern stats during '{}'", label));
        const auto uiaPatternStats = RunDialogUiaTaskWithMessagePump(label, [splash]() noexcept { return CollectVisibleUiaDescendantPatternStats(splash); });
        Trace(std::format(L"splash_surface: collected UIA pattern stats during '{}' hasValue={}", label, uiaPatternStats.has_value() ? 1 : 0));
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the splash window during {}.", label));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u, L"Splash window should expose visible UI Automation descendants.");
            state.Require(uiaPatternStats->textControlCount > 0u,
                          std::format(L"Splash window should expose a visible UI Automation text descendant during {}.", label));
        }

        Trace(std::format(L"splash_surface: collecting UIA text state during '{}'", label));
        const auto splashTextState =
            RunDialogUiaTaskWithMessagePump(label, [splash]() noexcept { return CollectVisibleSplashTextState(splash, L"Splash DXUI self-test"); });
        Trace(std::format(L"splash_surface: collected UIA text state during '{}' hasValue={}", label, splashTextState.has_value() ? 1 : 0));
        state.Require(splashTextState.has_value(), std::format(L"Splash window should expose a visible UI Automation text descendant during {}.", label));
        if (splashTextState.has_value())
        {
            state.Require(! splashTextState->name.empty(), std::format(L"Splash text surface should expose a stable accessible name during {}.", label));
            state.Require(splashTextState->name.find(L"Splash DXUI self-test") != std::wstring::npos,
                          std::format(L"Splash text surface should expose the live status text during {}; saw '{}'.", label, splashTextState->name));
        }

        return state.failure.empty();
    };

    state.Require(validateSplashSurface(L"the initial splash surface pass"), L"Splash window should expose the expected DX surface on the initial pass.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closeSplashWindow(splash, L"between splash surface passes"), L"Splash window should close cleanly between DX surface validation passes.");
    if (! state.failure.empty())
    {
        return false;
    }
    splash = nullptr;

    state.Require(validateSplashSurface(L"the reopened splash surface pass"), L"Splash window should expose the same DX surface after reopen.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSplashScreenLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto closeSplash = wil::scope_exit([]() noexcept { SplashScreen::CloseIfExist(); });

    const HWND existing = SplashScreen::GetHwnd();
    SplashScreen::CloseIfExist();
    if (existing && IsWindow(existing) != FALSE)
    {
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(5000ms)), L"Existing splash window did not close before long-run DXUI validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const std::wstring text = std::format(L"Splash DXUI churn {}", cycle);
        SplashScreen::SetOwner(mainWindow);
        SplashScreen::IfExistSetText(text.c_str());
        SplashScreen::BeginDelayedOpen(std::chrono::milliseconds{0}, GetModuleHandleW(nullptr));

        const HWND splash                               = WaitForWindow([] noexcept { return SplashScreen::GetHwnd(); }, SelfTest::Scale(5000ms));
        const SplashScreen::DebugSnapshot debugSnapshot = SplashScreen::DebugGetSnapshot();
        state.Require(splash != nullptr && IsWindow(splash) != FALSE,
                      std::format(L"Splash window did not open during cycle {}. stage={} lastError={} comHr=0x{:08X} threadStarted={} hasHwnd={}.",
                                  cycle,
                                  debugSnapshot.stage,
                                  debugSnapshot.lastError,
                                  static_cast<unsigned long>(debugSnapshot.comHr),
                                  debugSnapshot.threadStarted ? L"yes" : L"no",
                                  debugSnapshot.hasHwnd ? L"yes" : L"no"));
        if (! splash || IsWindow(splash) == FALSE)
        {
            return false;
        }

        state.Require(IsOwnedBy(splash, mainWindow), std::format(L"Splash window should be owned by the main window during cycle {}.", cycle));
        state.Require(CountVisibleChildWindows(splash) == 0u,
                      std::format(L"Splash window should not expose visible child-control fallback during cycle {}.", cycle));
        state.Require(WindowExposesUiaProvider(splash), std::format(L"Splash window should answer WM_GETOBJECT during cycle {}.", cycle));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(splash);
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the splash window during cycle {}.", cycle));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Splash window should expose visible UI Automation descendants during cycle {}.", cycle));
        }

        const auto splashTextState = CollectVisibleSplashTextState(splash, text);
        state.Require(splashTextState.has_value(), std::format(L"Splash window should expose a visible UI Automation text descendant during cycle {}.", cycle));
        if (splashTextState.has_value())
        {
            state.Require(! splashTextState->name.empty(), std::format(L"Splash text surface should expose a stable accessible name during cycle {}.", cycle));
            state.Require(splashTextState->name.find(text) != std::wstring::npos,
                          std::format(L"Splash text surface should expose the live status text during cycle {}; saw '{}'.", cycle, splashTextState->name));
        }

        SplashScreen::CloseIfExist();
        state.Require(WaitForWindowClosed(splash, SelfTest::Scale(5000ms)), std::format(L"Splash window did not close cleanly during cycle {}.", cycle));
    }

    state.Require(SplashScreen::GetHwnd() == nullptr || IsWindow(SplashScreen::GetHwnd()) == FALSE,
                  L"Splash window should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestSplashScreenLiveDxTextUpdate(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const auto cleanupSplash = wil::scope_exit([]() noexcept { SplashScreen::CloseIfExist(); });

    const HWND existing = SplashScreen::GetHwnd();
    SplashScreen::CloseIfExist();
    if (existing && IsWindow(existing) != FALSE)
    {
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(5000ms)), L"Existing splash window did not close before live text-update validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kInitialText         = L"Splash DXUI live text initial";
    constexpr std::wstring_view kUpdatedText         = L"Splash DXUI live text updated";
    constexpr std::wstring_view kReopenedInitialText = L"Splash DXUI live text reopened initial";
    constexpr std::wstring_view kReopenedUpdatedText = L"Splash DXUI live text reopened updated";

    SplashScreen::SetOwner(mainWindow);

    const auto openSplash = [&](std::wstring_view initialText, std::wstring_view context) noexcept
    {
        SplashScreen::IfExistSetText(initialText);
        SplashScreen::BeginDelayedOpen(std::chrono::milliseconds{0}, GetModuleHandleW(nullptr));

        const HWND splash                               = WaitForWindow([] noexcept { return SplashScreen::GetHwnd(); }, SelfTest::Scale(5000ms));
        const SplashScreen::DebugSnapshot debugSnapshot = SplashScreen::DebugGetSnapshot();
        state.Require(splash != nullptr && IsWindow(splash) != FALSE,
                      std::format(L"Splash window did not open for {}. stage={} lastError={} comHr=0x{:08X} threadStarted={} hasHwnd={}.",
                                  context,
                                  debugSnapshot.stage,
                                  debugSnapshot.lastError,
                                  static_cast<unsigned long>(debugSnapshot.comHr),
                                  debugSnapshot.threadStarted ? L"yes" : L"no",
                                  debugSnapshot.hasHwnd ? L"yes" : L"no"));
        if (splash && IsWindow(splash) != FALSE)
        {
            state.Require(IsOwnedBy(splash, mainWindow), std::format(L"Splash window should be owned by the main window during {}.", context));
            state.Require(CountVisibleChildWindows(splash) == 0u, std::format(L"Splash window should not expose visible child fallback during {}.", context));
            state.Require(WindowExposesUiaProvider(splash), std::format(L"Splash window should answer WM_GETOBJECT during {}.", context));
        }
        return splash;
    };

    const auto waitForSplashText = [&](const HWND splash, std::wstring_view expectedText) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const auto textState = CollectVisibleSplashTextState(splash, expectedText);
            if (textState.has_value() && textState->name.find(expectedText) != std::wstring::npos)
            {
                return textState;
            }

            std::this_thread::sleep_for(20ms);
        }

        return CollectVisibleSplashTextState(splash, expectedText);
    };

    const auto validateSplashTextCycle =
        [&](const HWND splash, std::wstring_view initialText, std::wstring_view updatedText, std::wstring_view context) noexcept
    {
        const auto initialTextState = CollectVisibleSplashTextState(splash, initialText);
        state.Require(initialTextState.has_value(),
                      std::format(L"Splash window should expose a visible UI Automation text descendant for the initial state during {}.", context));
        if (initialTextState.has_value())
        {
            state.Require(initialTextState->name.find(initialText) != std::wstring::npos,
                          std::format(L"Splash text surface should expose the initial live status text during {}; saw '{}'.", context, initialTextState->name));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        SplashScreen::IfExistSetText(updatedText);
        const auto updatedTextState = waitForSplashText(splash, updatedText);
        state.Require(
            updatedTextState.has_value(),
            std::format(L"Splash window should keep exposing a visible UI Automation text descendant after the live text update during {}.", context));
        if (updatedTextState.has_value())
        {
            state.Require(updatedTextState->name.find(updatedText) != std::wstring::npos,
                          std::format(L"Splash text surface should expose the updated live status text during {}; saw '{}'.", context, updatedTextState->name));
            state.Require(updatedTextState->name.find(initialText) == std::wstring::npos,
                          std::format(L"Splash text surface should replace the initial text after the live update during {}; saw '{}'.",
                                      context,
                                      updatedTextState->name));
        }
        return state.failure.empty();
    };

    const auto closeSplash = [&](const HWND splash, std::wstring_view context) noexcept
    {
        SplashScreen::CloseIfExist();
        state.Require(WaitForWindowClosed(splash, SelfTest::Scale(5000ms)), std::format(L"Splash window did not close cleanly during {}.", context));
        return state.failure.empty();
    };

    const HWND splash = openSplash(kInitialText, L"initial live text-update validation");
    if (! splash || IsWindow(splash) == FALSE)
    {
        return false;
    }

    state.Require(validateSplashTextCycle(splash, kInitialText, kUpdatedText, L"initial live text-update validation"),
                  L"Initial splash live DX text-update validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closeSplash(splash, L"initial live text-update validation"), L"Initial splash live DX close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedSplash = openSplash(kReopenedInitialText, L"reopened live text-update validation");
    if (! reopenedSplash || IsWindow(reopenedSplash) == FALSE)
    {
        return false;
    }

    state.Require(validateSplashTextCycle(reopenedSplash, kReopenedInitialText, kReopenedUpdatedText, L"reopened live text-update validation"),
                  L"Reopened splash live DX text-update validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closeSplash(reopenedSplash, L"reopened live text-update validation"), L"Reopened splash live DX close validation failed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneChangeCasePromptUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"change_case_prompt_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change-case prompt test root.");
    state.Require(SelfTest::WriteTextFile(root / L"foo.txt", "a"), L"Failed to create foo.txt.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewChangeCasePromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for change-case prompt test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"foo.txt"; }, true);

    closePrompt();

    struct ChangeCaseSurfaceCycleResult final
    {
        HWND prompt                 = nullptr;
        bool opened                 = false;
        bool ownedByMainWindow      = false;
        bool capturedSnapshot       = false;
        bool setSelections          = false;
        bool capturedEditedSnapshot = false;
        bool actionIssued           = false;
        bool closed                 = false;
        FolderViewChangeCasePromptDebugSnapshot snapshot{};
        FolderViewChangeCasePromptDebugSnapshot editedSnapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
        std::optional<UiaTogglePatternState> toggleState;
        std::optional<UiaNamedElementState> buttonState;
    };

    const auto runCycle =
        [&](bool accept, const size_t styleIndex, const size_t targetIndex, const bool includeSubdirs, std::wstring_view failureContext) noexcept
    {
        const auto expectedExampleResult = [](size_t expectedStyleIndex, size_t expectedTargetIndex) noexcept -> std::wstring
        {
            constexpr std::wstring_view kSample = L"Sample File.TXT";
            ChangeCase::Options options{};

            switch (std::min(expectedStyleIndex, static_cast<size_t>(3u)))
            {
                case 0u: options.style = ChangeCase::CaseStyle::Lower; break;
                case 1u: options.style = ChangeCase::CaseStyle::Upper; break;
                case 2u: options.style = ChangeCase::CaseStyle::PartiallyMixed; break;
                case 3u: options.style = ChangeCase::CaseStyle::Mixed; break;
            }

            switch (std::min(expectedTargetIndex, static_cast<size_t>(2u)))
            {
                case 0u: options.target = ChangeCase::ChangeTarget::WholeFilename; break;
                case 1u: options.target = ChangeCase::ChangeTarget::OnlyName; break;
                case 2u: options.target = ChangeCase::ChangeTarget::OnlyExtension; break;
            }

            return ChangeCase::TransformLeafName(kSample, options);
        };

        ChangeCaseSurfaceCycleResult cycle{};
        RunChangeCasePromptModalCycle(mainWindow,
                                      [&](const HWND prompt) noexcept
        {
            cycle.prompt = prompt;
            cycle.opened = prompt != nullptr && IsWindow(prompt) != FALSE;
            if (! cycle.opened)
            {
                return;
            }

            cycle.ownedByMainWindow      = IsOwnedBy(prompt, mainWindow);
            cycle.capturedSnapshot       = DebugGetFolderViewChangeCasePromptSnapshot(cycle.snapshot);
            cycle.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
            cycle.setSelections          = DebugSetFolderViewChangeCasePromptSelections(styleIndex, targetIndex, includeSubdirs);
            cycle.capturedEditedSnapshot = DebugGetFolderViewChangeCasePromptSnapshot(cycle.editedSnapshot);
            cycle.toggleState            = CollectVisibleDescendantTogglePatternState(prompt);
            cycle.buttonState            = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
            cycle.actionIssued           = accept ? DebugConfirmFolderViewChangeCasePrompt() : DebugCancelFolderViewChangeCasePrompt();
            if (cycle.actionIssued)
            {
                cycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
            }
        });

        state.Require(cycle.opened, std::format(L"Change-case prompt did not open for {}.", failureContext));
        state.Require(cycle.ownedByMainWindow, std::format(L"Change-case prompt should be owned by the main window during {}.", failureContext));
        state.Require(cycle.capturedSnapshot, std::format(L"Failed to capture change-case prompt snapshot during {}.", failureContext));
        if (! cycle.capturedSnapshot || ! state.failure.empty())
        {
            return cycle;
        }

        state.Require(cycle.snapshot.usesDxUiHost, std::format(L"Change-case prompt should use a DxUi host during {}.", failureContext));
        state.Require(cycle.snapshot.visibleChildWindowCount == 0u,
                      std::format(L"Change-case prompt should not expose visible fallback child controls during {}.", failureContext));
        state.Require(cycle.snapshot.includeSubdirsEnabled,
                      std::format(L"Change-case prompt should expose an enabled include-subdirectories toggle during {}.", failureContext));

        state.Require(cycle.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation stats for change-case prompt during {}.", failureContext));
        if (cycle.uiaPatternStats.has_value())
        {
            state.Require(cycle.uiaPatternStats->comboBoxControlCount >= 2u,
                          std::format(L"Change-case prompt should expose visible UI Automation combo descendants during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Change-case prompt should expose TogglePattern for include-subdirectories during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Change-case prompt should expose a visible UI Automation command button during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Change-case prompt should expose InvokePattern for its visible DX command button during {}.", failureContext));
            state.Require(
                cycle.uiaPatternStats->visibleElementCount >= 4u,
                std::format(L"Change-case prompt should expose visible UI Automation descendants for inputs and commands during {}.", failureContext));
        }

        state.Require(cycle.setSelections, std::format(L"Failed to set change-case prompt selections during {}.", failureContext));
        state.Require(cycle.capturedEditedSnapshot, std::format(L"Failed to recapture change-case prompt snapshot during {}.", failureContext));
        if (cycle.capturedEditedSnapshot)
        {
            state.Require(cycle.editedSnapshot.styleIndex == styleIndex,
                          std::format(L"Change-case prompt style index did not stick during {}.", failureContext));
            state.Require(cycle.editedSnapshot.targetIndex == targetIndex,
                          std::format(L"Change-case prompt target index did not stick during {}.", failureContext));
            state.Require(cycle.editedSnapshot.includeSubdirsChecked == includeSubdirs,
                          std::format(L"Change-case prompt include-subdirectories state did not stick during {}.", failureContext));
            state.Require(cycle.snapshot.exampleText.find(L"Sample File.TXT") != std::wstring::npos &&
                              cycle.snapshot.exampleText.find(L" -> ") != std::wstring::npos,
                          std::format(L"Change-case prompt should show a before/after example during {}.", failureContext));
            state.Require(cycle.editedSnapshot.exampleText.find(expectedExampleResult(styleIndex, targetIndex)) != std::wstring::npos,
                          std::format(L"Change-case prompt example should update to the edited selections during {}.", failureContext));
            if (styleIndex != 0u || targetIndex != 0u)
            {
                state.Require(cycle.editedSnapshot.exampleText != cycle.snapshot.exampleText,
                              std::format(L"Change-case prompt example should change when selections change during {}.", failureContext));
            }
        }

        state.Require(cycle.toggleState.has_value(), std::format(L"Change-case prompt should expose a visible DX toggle during {}.", failureContext));
        if (cycle.toggleState.has_value())
        {
            state.Require(! cycle.toggleState->name.empty(),
                          std::format(L"Change-case prompt visible DX toggle should expose a stable accessible name during {}.", failureContext));
            state.Require(cycle.toggleState->toggleState == (includeSubdirs ? ToggleState_On : ToggleState_Off),
                          std::format(L"Change-case prompt visible DX toggle state should match the live prompt snapshot during {}.", failureContext));
        }

        state.Require(cycle.buttonState.has_value(), std::format(L"Change-case prompt should expose a visible DX command button during {}.", failureContext));
        if (cycle.buttonState.has_value())
        {
            state.Require(! cycle.buttonState->name.empty(),
                          std::format(L"Change-case prompt visible DX command button should expose a stable accessible name during {}.", failureContext));
        }

        state.Require(cycle.actionIssued, std::format(L"Failed to close the change-case prompt through the DX action path during {}.", failureContext));
        state.Require(cycle.closed, std::format(L"Change-case prompt did not close after the DX action path during {}.", failureContext));
        return cycle;
    };

    runCycle(true, 1u, 0u, true, L"initial confirm pass");
    if (! state.failure.empty())
    {
        return false;
    }

    ChangeCase::Options confirmedOptions{};
    confirmedOptions.target          = ChangeCase::ChangeTarget::WholeFilename;
    confirmedOptions.style           = ChangeCase::CaseStyle::Upper;
    const std::wstring confirmedName = ChangeCase::TransformLeafName(L"foo.txt", confirmedOptions);
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {confirmedName}, SelfTest::Scale(3000ms)),
                  L"Change-case prompt confirm should leave the renamed item visible before the reopened cancel pass.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [&](std::wstring_view name) noexcept { return name == confirmedName; }, true);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == confirmedName,
                  L"Change-case prompt reopened cancel pass should focus the renamed item.");
    if (! state.failure.empty())
    {
        return false;
    }

    runCycle(false, 0u, 0u, false, L"reopened cancel pass");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneChangeCasePromptLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"change_case_prompt_live_dx_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change-case prompt live interaction root.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring currentName = L"rename_me.txt";
    state.Require(SelfTest::WriteTextFile(root / currentName, "rename"), L"Failed to create initial change-case live interaction file.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewChangeCasePromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for change-case live interaction test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for change-case live interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto readSingleLeafName = [&](const std::filesystem::path& directory) noexcept -> std::optional<std::wstring>
    {
        std::error_code enumEc;
        std::filesystem::directory_iterator it(directory, enumEc);
        if (enumEc)
        {
            return std::nullopt;
        }
        if (it == std::filesystem::directory_iterator{})
        {
            return std::nullopt;
        }

        const std::wstring name = it->path().filename().native();
        it.increment(enumEc);
        if (enumEc || it != std::filesystem::directory_iterator{})
        {
            return std::nullopt;
        }
        return name;
    };
    const auto waitForLeafName = [&](std::wstring_view expectedName, const auto timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const auto actualName = readSingleLeafName(root);
            if (actualName.has_value() && actualName.value() == expectedName)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        const auto actualName = readSingleLeafName(root);
        return actualName.has_value() && actualName.value() == expectedName;
    };

    const auto mutateIncludeSubdirsToggle = [&](const HWND prompt, const ToggleState expectedState, std::wstring_view context) noexcept
    {
        FolderViewChangeCasePromptDebugSnapshot snapshot{};
        state.Require(DebugGetFolderViewChangeCasePromptSnapshot(snapshot), std::format(L"Failed to capture change-case prompt snapshot during {}.", context));
        state.Require(snapshot.usesDxUiHost, std::format(L"Change-case prompt should use a DxUi host during {}.", context));
        state.Require(snapshot.visibleChildWindowCount == 0u,
                      std::format(L"Change-case prompt should not expose visible fallback child controls during {}.", context));
        state.Require(snapshot.includeSubdirsEnabled,
                      std::format(L"Change-case prompt should expose an enabled include-subdirectories toggle during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto initialToggleState = CollectVisibleDescendantTogglePatternState(prompt);
        state.Require(initialToggleState.has_value(), std::format(L"Change-case prompt should expose a visible TogglePattern descendant during {}.", context));
        if (! initialToggleState.has_value() || ! state.failure.empty())
        {
            return false;
        }

        state.Require(! initialToggleState->name.empty(),
                      std::format(L"Change-case prompt TogglePattern descendant should expose a stable accessible name during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const std::wstring toggleName = initialToggleState->name;
        if (initialToggleState->toggleState != expectedState)
        {
            state.Require(ToggleVisibleDescendantByName(prompt, toggleName),
                          std::format(L"Failed to toggle change-case prompt include-subdirectories control '{}' during {}.", toggleName, context));
            if (! state.failure.empty())
            {
                return false;
            }
        }

        const auto waitForToggleState = [&](const ToggleState value) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(prompt, toggleName);
                if (toggleState.has_value() && toggleState->toggleState == value)
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(prompt, toggleName);
            return toggleState.has_value() && toggleState->toggleState == value;
        };

        state.Require(waitForToggleState(expectedState),
                      std::format(L"Change-case prompt TogglePattern state did not update after live UIA interaction during {}.", context));

        FolderViewChangeCasePromptDebugSnapshot editedSnapshot{};
        state.Require(DebugGetFolderViewChangeCasePromptSnapshot(editedSnapshot),
                      std::format(L"Failed to recapture change-case prompt snapshot during {}.", context));
        state.Require(editedSnapshot.includeSubdirsChecked == (expectedState == ToggleState_On),
                      std::format(L"Change-case prompt include-subdirectories state should match the live UIA toggle during {}.", context));
        return state.failure.empty();
    };

    const auto focusItemForPrompt = [&](std::wstring_view expectedName, std::wstring_view context) noexcept
    {
        state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {std::wstring{expectedName}}, SelfTest::Scale(3000ms)),
                      std::format(L"Pane contents not settled on '{}' before opening the change-case prompt for {}.", expectedName, context));
        g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
            FolderWindow::Pane::Left, [&](std::wstring_view name) noexcept { return name == expectedName; }, true);
        state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedName,
                      std::format(L"Change-case live interaction expected focus on '{}' before opening the prompt for {}.", expectedName, context));
        if (! state.failure.empty())
        {
            return false;
        }

        FocusFolderViewPane(FolderWindow::Pane::Left);
        return state.failure.empty();
    };
    const auto setToggleStateViaUia =
        [&](const HWND prompt, const ToggleState expectedState, std::wstring& toggleName, FolderViewChangeCasePromptDebugSnapshot& editedSnapshot) noexcept
    {
        const auto initialToggleState = CollectVisibleDescendantTogglePatternState(prompt);
        if (! initialToggleState.has_value() || initialToggleState->name.empty())
        {
            return false;
        }

        toggleName = initialToggleState->name;
        if (initialToggleState->toggleState != expectedState)
        {
            if (! ToggleVisibleDescendantByName(prompt, toggleName))
            {
                return false;
            }

            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(prompt, toggleName);
                if (toggleState.has_value() && toggleState->toggleState == expectedState)
                {
                    break;
                }
                std::this_thread::sleep_for(20ms);
            }

            const auto toggleState = CollectVisibleDescendantTogglePatternStateByName(prompt, toggleName);
            if (! toggleState.has_value() || toggleState->toggleState != expectedState)
            {
                return false;
            }
        }

        return DebugGetFolderViewChangeCasePromptSnapshot(editedSnapshot);
    };

    constexpr size_t kUpperStyleIndex          = 1u;
    constexpr size_t kWholeFilenameTarget      = 0u;
    constexpr ToggleState kToggleEnabledState  = ToggleState_On;
    constexpr ToggleState kToggleDisabledState = ToggleState_Off;

    if (! focusItemForPrompt(currentName, L"live change-case cancel"))
    {
        return false;
    }

    struct ChangeCaseLiveCycleResult final
    {
        HWND prompt             = nullptr;
        bool opened             = false;
        bool ownedByMainWindow  = false;
        bool exposesUiaProvider = false;
        bool capturedSnapshot   = false;
        FolderViewChangeCasePromptDebugSnapshot snapshot{};
        std::optional<UiaTogglePatternState> initialToggleState;
        bool setSelections = false;
        bool toggled       = false;
        std::wstring toggleName;
        FolderViewChangeCasePromptDebugSnapshot editedSnapshot{};
        bool actionInvoked = false;
        bool closed        = false;
    };

    ChangeCaseLiveCycleResult cancelCycle{};
    RunChangeCasePromptModalCycle(mainWindow,
                                  [&](const HWND prompt) noexcept
    {
        cancelCycle.prompt = prompt;
        cancelCycle.opened = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! cancelCycle.opened)
        {
            return;
        }

        cancelCycle.ownedByMainWindow  = IsOwnedBy(prompt, mainWindow);
        cancelCycle.exposesUiaProvider = WindowExposesUiaProvider(prompt);
        cancelCycle.capturedSnapshot   = DebugGetFolderViewChangeCasePromptSnapshot(cancelCycle.snapshot);
        cancelCycle.initialToggleState = CollectVisibleDescendantTogglePatternState(prompt);
        cancelCycle.setSelections      = DebugSetFolderViewChangeCasePromptSelections(kUpperStyleIndex, kWholeFilenameTarget, false);
        if (cancelCycle.setSelections)
        {
            cancelCycle.toggled = setToggleStateViaUia(prompt, kToggleEnabledState, cancelCycle.toggleName, cancelCycle.editedSnapshot);
        }
        cancelCycle.actionInvoked = DebugCancelFolderViewChangeCasePrompt();
        if (cancelCycle.actionInvoked)
        {
            cancelCycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        }
    });

    state.Require(cancelCycle.opened, L"Change-case prompt did not open for live change-case cancel.");
    state.Require(cancelCycle.ownedByMainWindow, L"Change-case prompt should be owned by the main window during live change-case cancel.");
    state.Require(cancelCycle.exposesUiaProvider, L"Change-case prompt should answer WM_GETOBJECT during live change-case cancel.");
    state.Require(cancelCycle.capturedSnapshot, L"Failed to capture initial change-case prompt snapshot before live cancel interaction.");
    state.Require(cancelCycle.snapshot.usesDxUiHost, L"Change-case prompt should use a DxUi host before live cancel interaction.");
    state.Require(cancelCycle.snapshot.visibleChildWindowCount == 0u,
                  L"Change-case prompt should not expose visible fallback child controls before live cancel interaction.");
    state.Require(cancelCycle.snapshot.includeSubdirsEnabled,
                  L"Change-case prompt should expose an enabled include-subdirectories toggle before live cancel interaction.");
    state.Require(cancelCycle.setSelections, L"Failed to seed change-case prompt selections before live cancel interaction.");
    state.Require(cancelCycle.toggled, L"Change-case prompt live DX toggle mutation failed before cancel.");
    state.Require(! cancelCycle.toggleName.empty(),
                  L"Change-case prompt TogglePattern descendant should expose a stable accessible name during live change-case cancel.");
    state.Require(cancelCycle.editedSnapshot.includeSubdirsChecked,
                  L"Change-case prompt include-subdirectories state should match the live UIA toggle during live change-case cancel.");
    state.Require(cancelCycle.actionInvoked, L"Change-case prompt did not close through the debug cancel path after live DX toggle interaction.");
    state.Require(cancelCycle.closed, L"Change-case prompt did not close after the debug cancel path during live DX interaction.");
    FolderViewChangeCasePromptDebugSnapshot initialSnapshot = cancelCycle.snapshot;
    if (! state.failure.empty())
    {
        return false;
    }

    ChangeCase::Options cancelledOptions{};
    cancelledOptions.target          = ChangeCase::ChangeTarget::WholeFilename;
    cancelledOptions.style           = ChangeCase::CaseStyle::Upper;
    const std::wstring cancelledName = ChangeCase::TransformLeafName(currentName, cancelledOptions);
    const auto actualCancelledName   = readSingleLeafName(root);
    state.Require(actualCancelledName.has_value(), L"Change-case live cancel interaction should leave exactly one file on disk.");
    if (actualCancelledName.has_value())
    {
        state.Require(actualCancelledName.value() == currentName,
                      std::format(L"Change-case live cancel interaction should preserve '{}'; saw '{}'.", currentName, actualCancelledName.value()));
        state.Require(actualCancelledName.value() != cancelledName, L"Change-case live cancel interaction should not apply the pending rename.");
    }
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                  L"Pane items should remain on the original filename after live change-case cancel interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! focusItemForPrompt(currentName, L"live change-case confirm"))
    {
        return false;
    }

    ChangeCaseLiveCycleResult confirmCycle{};
    RunChangeCasePromptModalCycle(mainWindow,
                                  [&](const HWND prompt) noexcept
    {
        confirmCycle.prompt = prompt;
        confirmCycle.opened = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! confirmCycle.opened)
        {
            return;
        }

        confirmCycle.ownedByMainWindow  = IsOwnedBy(prompt, mainWindow);
        confirmCycle.exposesUiaProvider = WindowExposesUiaProvider(prompt);
        confirmCycle.capturedSnapshot   = DebugGetFolderViewChangeCasePromptSnapshot(confirmCycle.snapshot);
        confirmCycle.initialToggleState = CollectVisibleDescendantTogglePatternState(prompt);
        confirmCycle.setSelections      = DebugSetFolderViewChangeCasePromptSelections(kUpperStyleIndex, kWholeFilenameTarget, false);
        if (confirmCycle.setSelections)
        {
            confirmCycle.toggled = setToggleStateViaUia(prompt, kToggleEnabledState, confirmCycle.toggleName, confirmCycle.editedSnapshot);
        }
        confirmCycle.actionInvoked = DebugConfirmFolderViewChangeCasePrompt();
        if (confirmCycle.actionInvoked)
        {
            confirmCycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        }
    });

    state.Require(confirmCycle.opened, L"Change-case prompt did not open for live change-case confirm.");
    state.Require(confirmCycle.ownedByMainWindow, L"Reopened change-case prompt should be owned by the main window before live confirm interaction.");
    state.Require(confirmCycle.exposesUiaProvider, L"Reopened change-case prompt should answer WM_GETOBJECT before live confirm interaction.");
    state.Require(confirmCycle.capturedSnapshot, L"Failed to capture reopened change-case prompt snapshot before live confirm interaction.");
    state.Require(confirmCycle.snapshot.usesDxUiHost, L"Reopened change-case prompt should use a DxUi host before live confirm interaction.");
    state.Require(confirmCycle.snapshot.visibleChildWindowCount == 0u,
                  L"Reopened change-case prompt should not expose visible fallback child controls before live confirm interaction.");
    state.Require(confirmCycle.snapshot.styleIndex == initialSnapshot.styleIndex,
                  L"Reopened change-case prompt should restore its baseline style selection before live confirm interaction.");
    state.Require(confirmCycle.snapshot.targetIndex == initialSnapshot.targetIndex,
                  L"Reopened change-case prompt should restore its baseline target selection before live confirm interaction.");
    state.Require(confirmCycle.snapshot.includeSubdirsChecked == initialSnapshot.includeSubdirsChecked,
                  L"Reopened change-case prompt should restore its baseline include-subdirectories state before live confirm interaction.");
    state.Require(confirmCycle.initialToggleState.has_value(),
                  L"Reopened change-case prompt should expose a visible TogglePattern descendant before live confirm interaction.");
    if (confirmCycle.initialToggleState.has_value())
    {
        state.Require(confirmCycle.initialToggleState->toggleState == kToggleDisabledState,
                      L"Reopened change-case prompt should restore its baseline include-subdirectories toggle state before live confirm interaction.");
    }
    state.Require(confirmCycle.setSelections, L"Failed to reseed change-case prompt selections before live confirm interaction.");
    state.Require(confirmCycle.toggled, L"Change-case prompt live DX toggle mutation failed before confirm.");
    state.Require(! confirmCycle.toggleName.empty(),
                  L"Change-case prompt TogglePattern descendant should expose a stable accessible name during live change-case confirm.");
    state.Require(confirmCycle.editedSnapshot.includeSubdirsChecked,
                  L"Change-case prompt include-subdirectories state should match the live UIA toggle during live change-case confirm.");
    state.Require(confirmCycle.actionInvoked, L"Change-case prompt did not close through the debug confirm path after live DX toggle interaction.");
    state.Require(confirmCycle.closed, L"Change-case prompt did not close after the debug confirm path during live DX interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    ChangeCase::Options confirmOptions{};
    confirmOptions.target            = ChangeCase::ChangeTarget::WholeFilename;
    confirmOptions.style             = ChangeCase::CaseStyle::Upper;
    const std::wstring confirmedName = ChangeCase::TransformLeafName(currentName, confirmOptions);
    state.Require(waitForLeafName(confirmedName, SelfTest::Scale(5000ms)), L"Change-case live interaction did not rename the file after confirm.");
    const auto actualConfirmedName = readSingleLeafName(root);
    state.Require(actualConfirmedName.has_value(), L"Change-case live confirm interaction should leave exactly one file on disk.");
    if (actualConfirmedName.has_value())
    {
        state.Require(
            actualConfirmedName.value() == confirmedName,
            std::format(L"Change-case live confirm interaction should rename the file to '{}'; saw '{}'.", confirmedName, actualConfirmedName.value()));
    }
    currentName = confirmedName;
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneChangeCasePromptLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"change_case_prompt_churn_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change-case prompt churn root.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring currentName = L"rename_me.txt";
    state.Require(SelfTest::WriteTextFile(root / currentName, "rename"), L"Failed to create initial change-case prompt churn file.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewChangeCasePromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for change-case prompt churn test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for change-case prompt churn test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto readSingleLeafName = [&](const std::filesystem::path& directory) noexcept -> std::optional<std::wstring>
    {
        std::error_code enumEc;
        std::filesystem::directory_iterator it(directory, enumEc);
        if (enumEc)
        {
            return std::nullopt;
        }
        if (it == std::filesystem::directory_iterator{})
        {
            return std::nullopt;
        }

        const std::wstring name = it->path().filename().native();
        it.increment(enumEc);
        if (enumEc || it != std::filesystem::directory_iterator{})
        {
            return std::nullopt;
        }
        return name;
    };
    const auto waitForLeafName = [&](std::wstring_view expectedName, const auto timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const auto actualName = readSingleLeafName(root);
            if (actualName.has_value() && actualName.value() == expectedName)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        const auto actualName = readSingleLeafName(root);
        return actualName.has_value() && actualName.value() == expectedName;
    };

    constexpr size_t kWholeFilenameTarget = 0u;
    constexpr size_t kCycles              = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const bool accept          = (cycle % 2u) == 0u;
        const bool includeSubdirs  = (cycle % 3u) == 0u;
        const size_t acceptOrdinal = cycle / 2u;
        const size_t styleIndex    = accept ? (((acceptOrdinal % 2u) == 0u) ? 1u : 0u) : (((acceptOrdinal % 2u) == 0u) ? 0u : 1u);

        const auto actualBeforeName = readSingleLeafName(root);
        state.Require(actualBeforeName.has_value(), std::format(L"Expected exactly one file before change-case cycle {}.", cycle));
        if (actualBeforeName.has_value())
        {
            state.Require(actualBeforeName.value() == currentName,
                          std::format(L"Expected '{}' before change-case cycle {}; saw '{}'.", currentName, cycle, actualBeforeName.value()));
        }
        state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                      std::format(L"Pane contents not settled before change-case cycle {}.", cycle));
        g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
            FolderWindow::Pane::Left, [&](std::wstring_view name) noexcept { return name == currentName; }, true);
        state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == currentName,
                      std::format(L"Change-case churn expected focus on '{}' before cycle {}.", currentName, cycle));
        if (! state.failure.empty())
        {
            closePrompt();
            return false;
        }

        struct ChangeCaseCycleResult final
        {
            HWND prompt            = nullptr;
            bool opened            = false;
            bool ownedByMainWindow = false;
            bool exposesUia        = false;
            FolderViewChangeCasePromptDebugSnapshot snapshot{};
            bool capturedSnapshot = false;
            std::optional<UiaDescendantPatternStats> uiaPatternStats;
            bool setSelections = false;
            FolderViewChangeCasePromptDebugSnapshot editedSnapshot{};
            bool capturedEditedSnapshot = false;
            std::optional<UiaTogglePatternState> toggleState;
            std::optional<UiaNamedElementState> buttonState;
            bool actionTriggered = false;
            bool closed          = false;
        } cycleResult{};

        closePrompt();
        RunChangeCasePromptModalCycle(mainWindow,
                                      [&](const HWND prompt) noexcept
        {
            cycleResult.prompt            = prompt;
            cycleResult.opened            = prompt != nullptr && IsWindow(prompt) != FALSE;
            cycleResult.ownedByMainWindow = prompt != nullptr && IsOwnedBy(prompt, mainWindow);
            cycleResult.exposesUia        = prompt != nullptr && WindowExposesUiaProvider(prompt);
            if (! cycleResult.opened)
            {
                return;
            }

            cycleResult.capturedSnapshot       = DebugGetFolderViewChangeCasePromptSnapshot(cycleResult.snapshot);
            cycleResult.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
            cycleResult.setSelections          = DebugSetFolderViewChangeCasePromptSelections(styleIndex, kWholeFilenameTarget, includeSubdirs);
            cycleResult.capturedEditedSnapshot = DebugGetFolderViewChangeCasePromptSnapshot(cycleResult.editedSnapshot);
            cycleResult.toggleState            = CollectVisibleDescendantTogglePatternState(prompt);
            cycleResult.buttonState            = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
            cycleResult.actionTriggered        = accept ? DebugConfirmFolderViewChangeCasePrompt() : DebugCancelFolderViewChangeCasePrompt();
            if (cycleResult.actionTriggered)
            {
                cycleResult.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
            }
        });

        state.Require(cycleResult.opened, std::format(L"Change-case prompt did not open during cycle {}.", cycle));
        if (! cycleResult.opened)
        {
            closePrompt();
            return false;
        }

        state.Require(cycleResult.ownedByMainWindow, std::format(L"Change-case prompt should be owned by the main window during cycle {}.", cycle));
        state.Require(cycleResult.exposesUia, std::format(L"Change-case prompt should answer WM_GETOBJECT during cycle {}.", cycle));
        state.Require(cycleResult.capturedSnapshot, std::format(L"Failed to capture change-case prompt snapshot during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.usesDxUiHost, std::format(L"Change-case prompt should use a DxUi host during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.visibleChildWindowCount == 0u,
                      std::format(L"Change-case prompt should not expose visible fallback child controls during cycle {}; saw {}.",
                                  cycle,
                                  cycleResult.snapshot.visibleChildWindowCount));
        state.Require(cycleResult.snapshot.includeSubdirsEnabled,
                      std::format(L"Change-case prompt should expose an enabled include-subdirectories toggle during cycle {}.", cycle));

        state.Require(cycleResult.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect UI Automation stats for change-case prompt during cycle {}.", cycle));
        if (cycleResult.uiaPatternStats.has_value())
        {
            state.Require(cycleResult.uiaPatternStats->visibleElementCount >= 4u,
                          std::format(L"Change-case prompt should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->comboBoxControlCount >= 2u,
                          std::format(L"Change-case prompt should expose visible combo descendants during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Change-case prompt should expose TogglePattern during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Change-case prompt should expose visible command buttons during cycle {}.", cycle));
        }

        state.Require(cycleResult.setSelections, std::format(L"Failed to set change-case prompt selections during cycle {}.", cycle));
        state.Require(cycleResult.capturedEditedSnapshot, std::format(L"Failed to recapture change-case prompt snapshot during cycle {}.", cycle));
        state.Require(cycleResult.editedSnapshot.styleIndex == styleIndex,
                      std::format(L"Change-case prompt style index did not stick during cycle {}.", cycle));
        state.Require(cycleResult.editedSnapshot.targetIndex == kWholeFilenameTarget,
                      std::format(L"Change-case prompt target index did not stick during cycle {}.", cycle));
        state.Require(cycleResult.editedSnapshot.includeSubdirsChecked == includeSubdirs,
                      std::format(L"Change-case prompt include-subdirectories state did not stick during cycle {}.", cycle));

        state.Require(cycleResult.toggleState.has_value(),
                      std::format(L"Failed to collect UI Automation toggle state for change-case prompt during cycle {}.", cycle));
        if (cycleResult.toggleState.has_value())
        {
            const ToggleState expectedToggleState = includeSubdirs ? ToggleState_On : ToggleState_Off;
            state.Require(cycleResult.toggleState->toggleState == expectedToggleState,
                          std::format(L"Change-case prompt toggle state mismatch during cycle {}.", cycle));
            state.Require(! cycleResult.toggleState->name.empty(),
                          std::format(L"Change-case prompt visible DX toggle should expose a stable accessible name during cycle {}.", cycle));
        }

        state.Require(cycleResult.buttonState.has_value(),
                      std::format(L"Change-case prompt should expose a visible DX command button during cycle {}.", cycle));
        if (cycleResult.buttonState.has_value())
        {
            state.Require(! cycleResult.buttonState->name.empty(),
                          std::format(L"Change-case prompt visible DX command button should expose a stable accessible name during cycle {}.", cycle));
        }

        state.Require(cycleResult.actionTriggered, std::format(L"Failed to close the change-case prompt through the DX action path during cycle {}.", cycle));
        state.Require(cycleResult.closed, std::format(L"Change-case prompt did not close after cycle {}.", cycle));

        if (accept)
        {
            ChangeCase::Options options{};
            options.target                  = ChangeCase::ChangeTarget::WholeFilename;
            options.style                   = (styleIndex == 0u) ? ChangeCase::CaseStyle::Lower : ChangeCase::CaseStyle::Upper;
            const std::wstring expectedName = ChangeCase::TransformLeafName(currentName, options);
            state.Require(waitForLeafName(expectedName, SelfTest::Scale(5000ms)),
                          std::format(L"Change-case prompt did not rename '{}' to '{}' during cycle {}.", currentName, expectedName, cycle));
            const auto actualAfterName = readSingleLeafName(root);
            state.Require(actualAfterName.has_value(), std::format(L"Expected exactly one file after change-case cycle {}.", cycle));
            if (actualAfterName.has_value())
            {
                state.Require(actualAfterName.value() == expectedName,
                              std::format(L"Change-case cycle {} should leave '{}'; saw '{}'.", cycle, expectedName, actualAfterName.value()));
            }
            state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {expectedName}, SelfTest::Scale(3000ms)),
                          std::format(L"Pane items did not settle on '{}' after change-case cycle {}.", expectedName, cycle));
            currentName = expectedName;
        }
        else
        {
            const auto actualAfterName = readSingleLeafName(root);
            state.Require(actualAfterName.has_value(), std::format(L"Expected exactly one file after change-case cancel cycle {}.", cycle));
            if (actualAfterName.has_value())
            {
                state.Require(actualAfterName.value() == currentName,
                              std::format(L"Cancel should preserve '{}' after change-case cycle {}; saw '{}'.", currentName, cycle, actualAfterName.value()));
            }
            state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                          std::format(L"Pane items should remain on '{}' after change-case cancel cycle {}.", currentName, cycle));
        }

        state.Require(GetFolderViewChangeCasePromptHandle() == nullptr || IsWindow(GetFolderViewChangeCasePromptHandle()) == FALSE,
                      std::format(L"Change-case prompt should not remain open after cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closePrompt();
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestChangeCaseDialogAndMultiSelection(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"change_case_dialog_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change_case_dialog root.");

    const std::filesystem::path foo = root / L"foo.txt";
    const std::filesystem::path bar = root / L"bar.baz";
    state.Require(SelfTest::WriteTextFile(foo, "a"), L"Failed to create foo.txt.");
    state.Require(SelfTest::WriteTextFile(bar, "b"), L"Failed to create bar.baz.");

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
                  L"Failed to set local file-system plugin for change-case dialog test.");

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
                  L"Failed to set left pane path for change-case dialog test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for change-case dialog test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"foo.txt", L"bar.baz"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Folder contents not ready for change-case dialog test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"foo.txt" || name == L"bar.baz"; }, true);

    ChangeCasePromptAutomationState first{};
    std::jthread okCloser([&](std::stop_token) noexcept { AutomateChangeCasePrompt(mainWindow, first, 1u, 0u, false, true); });
    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_CASE, 0), 0);
    okCloser.join();

    state.Require(first.sawDialog.load(std::memory_order_acquire), L"Change Case dialog did not open.");
    state.Require(first.closed.load(std::memory_order_acquire), L"Change Case dialog did not close after OK.");
    state.Require(first.includeEnabled.load(std::memory_order_acquire), L"Change Case include-subdirectories checkbox unexpectedly disabled.");

    const auto renameDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{5000});
    while (std::chrono::steady_clock::now() < renameDeadline)
    {
        PumpPendingMessages();
        if (std::filesystem::exists(root / L"FOO.TXT", ec) && std::filesystem::exists(root / L"BAR.BAZ", ec))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(std::filesystem::exists(root / L"FOO.TXT", ec), L"Change case did not rename foo.txt to FOO.TXT.");
    state.Require(std::filesystem::exists(root / L"BAR.BAZ", ec), L"Change case did not rename bar.baz to BAR.BAZ.");
    state.Require(ForceRefreshPaneForCommandSelfTest(mainWindow, FolderWindow::Pane::Left, SelfTest::Scale(3000ms)),
                  L"Failed to refresh the left pane after the first change-case operation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"FOO.TXT", L"BAR.BAZ"}, SelfTest::Scale(3000ms)),
                  L"Pane did not show renamed change-case items before reopening the dialog.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"FOO.TXT" || name == L"BAR.BAZ"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 2u,
                  L"Expected two renamed items selected before reopening the Change Case dialog.");
    if (! state.failure.empty())
    {
        return false;
    }

    ChangeCasePromptAutomationState second{};
    std::jthread cancelCloser([&](std::stop_token) noexcept { AutomateChangeCasePrompt(mainWindow, second, 0u, 0u, false, false); });
    FocusFolderViewPane(FolderWindow::Pane::Left);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_CHANGE_CASE, 0), 0);
    cancelCloser.join();

    state.Require(second.sawDialog.load(std::memory_order_acquire), L"Change Case dialog did not reopen after completing an operation.");
    state.Require(second.closed.load(std::memory_order_acquire), L"Change Case dialog did not close after Cancel.");

    return state.failure.empty();
}

[[nodiscard]] bool TestChangeCaseCore(CaseState& state) noexcept
{
    using ChangeCase::CaseStyle;
    using ChangeCase::ChangeTarget;

    ChangeCase::Options options{};
    options.style  = CaseStyle::PartiallyMixed;
    options.target = ChangeTarget::WholeFilename;

    state.Require(ChangeCase::TransformLeafName(L"hello_world.TXT", options) == L"Hello_World.txt", L"TransformLeafName partially-mixed failed.");

    options.style  = CaseStyle::Upper;
    options.target = ChangeTarget::OnlyExtension;
    state.Require(ChangeCase::TransformLeafName(L"file.txt", options) == L"file.TXT", L"TransformLeafName upper ext failed.");

    wil::com_ptr<IFileSystem> fs = SelfTest::GetFileSystem(L"builtin/file-system");
    state.Require(static_cast<bool>(fs), L"builtin/file-system plugin not available.");
    if (! fs)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable.");
    if (suiteRoot.empty())
    {
        return false;
    }

    const std::filesystem::path root = suiteRoot / L"work" / L"change_case";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create change-case work directory.");

    const std::filesystem::path a      = root / L"Foo.TXT";
    const std::filesystem::path b      = root / L"bar.BAZ";
    const std::filesystem::path subdir = root / L"subdir";
    const std::filesystem::path nested = subdir / L"Nested.TXT";
    state.Require(SelfTest::WriteTextFile(a, "a"), L"Failed to create Foo.TXT.");
    state.Require(SelfTest::WriteTextFile(b, "b"), L"Failed to create bar.BAZ.");
    state.Require(SelfTest::EnsureDirectory(subdir), L"Failed to create subdir.");
    state.Require(SelfTest::WriteTextFile(nested, "c"), L"Failed to create Nested.TXT.");

    struct ProgressCapture final
    {
        bool sawEnumerating = false;
        bool sawRenaming    = false;

        uint64_t maxScannedFolders = 0;
        uint64_t maxScannedEntries = 0;
        uint64_t plannedRenames    = 0;
        uint64_t completedRenames  = 0;
    };

    ChangeCase::Options apply{};
    apply.style          = CaseStyle::Lower;
    apply.target         = ChangeTarget::WholeFilename;
    apply.includeSubdirs = true;

    ProgressCapture capture{};
    const auto onProgress = [](const ChangeCase::ProgressUpdate& update, void* cookie) noexcept
    {
        auto* cap = static_cast<ProgressCapture*>(cookie);
        if (! cap)
        {
            return;
        }

        if (update.phase == ChangeCase::ProgressUpdate::Phase::Enumerating)
        {
            cap->sawEnumerating = true;
        }
        else if (update.phase == ChangeCase::ProgressUpdate::Phase::Renaming)
        {
            cap->sawRenaming = true;
        }

        if (update.scannedFolders > cap->maxScannedFolders)
        {
            cap->maxScannedFolders = update.scannedFolders;
        }
        if (update.scannedEntries > cap->maxScannedEntries)
        {
            cap->maxScannedEntries = update.scannedEntries;
        }
        if (update.plannedRenames > cap->plannedRenames)
        {
            cap->plannedRenames = update.plannedRenames;
        }
        cap->completedRenames = update.completedRenames;
    };

    const HRESULT hr = ChangeCase::ApplyToPaths(*fs, {a, b, subdir}, apply, {}, onProgress, &capture);
    state.Require(SUCCEEDED(hr), std::format(L"ApplyToPaths failed (hr=0x{:08X}).", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require(capture.sawEnumerating, L"ChangeCase progress callback did not report Enumerating.");
    state.Require(capture.sawRenaming, L"ChangeCase progress callback did not report Renaming.");
    state.Require(capture.maxScannedFolders >= 1u, L"ChangeCase expected to scan at least one folder when includeSubdirs is enabled.");
    state.Require(capture.maxScannedEntries >= 1u, L"ChangeCase expected to scan at least one entry when includeSubdirs is enabled.");
    state.Require(capture.plannedRenames == 3u, std::format(L"ChangeCase planned renames mismatch (expected 3, got {}).", capture.plannedRenames));
    state.Require(capture.completedRenames == 3u, std::format(L"ChangeCase completed renames mismatch (expected 3, got {}).", capture.completedRenames));

    std::unordered_set<std::wstring> names;
    for (const auto& entry : std::filesystem::directory_iterator(root, ec))
    {
        if (ec)
        {
            break;
        }
        names.insert(entry.path().filename().wstring());
    }

    state.Require(names.contains(L"foo.txt"), L"Expected foo.txt after change case.");
    state.Require(names.contains(L"bar.baz"), L"Expected bar.baz after change case.");
    state.Require(names.contains(L"subdir"), L"Expected subdir entry after change case.");

    ec.clear();
    std::unordered_set<std::wstring> subNames;
    for (const auto& entry : std::filesystem::directory_iterator(subdir, ec))
    {
        if (ec)
        {
            break;
        }
        subNames.insert(entry.path().filename().wstring());
    }
    state.Require(subNames.contains(L"nested.txt"), L"Expected nested.txt after change case includeSubdirs.");

    const std::filesystem::path rootUpper = suiteRoot / L"work" / L"change_case_upper";
    ec.clear();
    std::filesystem::remove_all(rootUpper, ec);
    state.Require(SelfTest::EnsureDirectory(rootUpper), L"Failed to create change_case_upper root.");

    const std::filesystem::path upperA = rootUpper / L"foo.txt";
    const std::filesystem::path upperB = rootUpper / L"bar.baz";
    state.Require(SelfTest::WriteTextFile(upperA, "x"), L"Failed to create foo.txt.");
    state.Require(SelfTest::WriteTextFile(upperB, "y"), L"Failed to create bar.baz.");

    ChangeCase::Options upper{};
    upper.style          = CaseStyle::Upper;
    upper.target         = ChangeTarget::WholeFilename;
    upper.includeSubdirs = false;

    const HRESULT upperHr = ChangeCase::ApplyToPaths(*fs, {upperA, upperB}, upper);
    state.Require(SUCCEEDED(upperHr), std::format(L"ApplyToPaths upper failed (hr=0x{:08X}).", static_cast<unsigned long>(upperHr)));
    if (FAILED(upperHr))
    {
        return false;
    }

    ec.clear();
    std::unordered_set<std::wstring> upperNames;
    for (const auto& entry : std::filesystem::directory_iterator(rootUpper, ec))
    {
        if (ec)
        {
            break;
        }
        upperNames.insert(entry.path().filename().wstring());
    }

    state.Require(upperNames.contains(L"FOO.TXT"), L"Expected FOO.TXT after change case upper.");
    state.Require(upperNames.contains(L"BAR.BAZ"), L"Expected BAR.BAZ after change case upper.");
    return state.failure.empty();
}

[[nodiscard]] bool WaitForAtomicAtLeast(const std::atomic<uint32_t>& value, uint32_t expected, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();

        if (value.load(std::memory_order_acquire) >= expected)
        {
            return true;
        }

        std::this_thread::sleep_for(20ms);
    }

    return value.load(std::memory_order_acquire) >= expected;
}

[[nodiscard]] bool TestMaskSyntaxWildcardMatching(CaseState& state) noexcept
{
    const auto empty = MaskSyntax::ParseWildcardMask(L"");
    state.Require(empty.includePatterns.empty() && empty.excludePatterns.empty(), L"Empty mask should have no patterns.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"anything", empty), L"Empty mask should match any text.");

    const auto aQuestion = MaskSyntax::ParseWildcardMask(L"*.a?");
    state.Require(MaskSyntax::MatchesWildcardMask(L"x.al", aQuestion), L"*.a? should match x.al.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"anything.ai", aQuestion), L"*.a? should match anything.ai.");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"x.a", aQuestion), L"*.a? should not match x.a (missing char).");

    const auto includeList = MaskSyntax::ParseWildcardMask(L" *.txt ;  *.doc ");
    state.Require(includeList.includePatterns.size() == 2u, L"Include list should parse two patterns.");
    state.Require(includeList.excludePatterns.empty(), L"Include list should have no excludes.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"a.txt", includeList), L"Include list should match a.txt.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"b.DOC", includeList), L"Include list should match b.DOC (case-insensitive).");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"c.pdf", includeList), L"Include list should not match c.pdf.");

    const auto escapedSemicolon = MaskSyntax::ParseWildcardMask(L"one;;mask");
    state.Require(escapedSemicolon.includePatterns.size() == 1u, L"Escaped semicolon mask should parse one pattern.");
    state.Require(escapedSemicolon.includePatterns[0] == L"one;mask", L"Escaped semicolon should keep a single ';' in the token.");

    const auto excludeOnly = MaskSyntax::ParseWildcardMask(L"|*.tmp;*.bak");
    state.Require(excludeOnly.includePatterns.empty(), L"Exclude-only mask should have empty include patterns.");
    state.Require(excludeOnly.excludePatterns.size() == 2u, L"Exclude-only mask should parse two exclude patterns.");
    state.Require(MaskSyntax::MatchesWildcardMask(L"keep.txt", excludeOnly), L"Exclude-only mask should include keep.txt.");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"drop.tmp", excludeOnly), L"Exclude-only mask should exclude drop.tmp.");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"drop.BAK", excludeOnly), L"Exclude-only mask should exclude drop.BAK (case-insensitive).");

    const auto includeExclude = MaskSyntax::ParseWildcardMask(L"*.txt|a*");
    state.Require(MaskSyntax::MatchesWildcardMask(L"foo.txt", includeExclude), L"Include/exclude mask should match foo.txt.");
    state.Require(! MaskSyntax::MatchesWildcardMask(L"a.txt", includeExclude), L"Include/exclude mask should exclude a.txt.");

    std::vector<std::wstring> history = {L"  *.txt  ", L"*.TXT", L"", L"*.doc", L"*.DOC", L"*.json"};
    MaskSyntax::NormalizeWildcardMaskHistory(history, 3u);
    state.Require(history.size() == 3u, L"NormalizeWildcardMaskHistory should clamp to max items and dedupe.");
    state.Require(history[0] == L"*.txt", L"NormalizeWildcardMaskHistory should trim whitespace and keep first entry.");
    state.Require(history[1] == L"*.doc", L"NormalizeWildcardMaskHistory should dedupe case-insensitively.");

    MaskSyntax::AddToWildcardMaskHistory(history, 3u, L"  *.JSON ");
    state.Require(history.size() == 3u, L"AddToWildcardMaskHistory should not exceed max items.");
    state.Require(history[0] == L"*.json", L"AddToWildcardMaskHistory should move existing (case-insensitive) entry to front.");

    return state.failure.empty();
}

void AutomatePaneFilterDialog(HWND mainWindow, PaneFilterDialogAutomationState& dlgState, bool enabled, std::wstring_view maskText, bool accept) noexcept
{
    const HWND dlg = WaitForWindow(
        [mainWindow]() noexcept -> HWND
    {
        const HWND dlg = GetFolderViewPaneFilterPromptHandle();
        if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
        {
            return nullptr;
        }

        return dlg;
    },
        SelfTest::Scale(std::chrono::milliseconds{10000}));
    if (! dlg)
    {
        return;
    }

    dlgState.sawDialog.store(true, std::memory_order_release);

    static_cast<void>(DebugSetFolderViewPaneFilterPromptEnabled(enabled));
    static_cast<void>(DebugSetFolderViewPaneFilterPromptText(maskText));
    static_cast<void>(accept ? DebugConfirmFolderViewPaneFilterPrompt() : DebugCancelFolderViewPaneFilterPrompt());

    bool closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    if (! closed)
    {
        PostMessageW(dlg, WM_CLOSE, 0, 0);
        PostMessageW(dlg, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(dlg, WM_KEYUP, VK_ESCAPE, 0);
        closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    }

    dlgState.closed.store(closed, std::memory_order_release);
}

[[nodiscard]] bool TestPaneFilterPromptUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_prompt_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-filter prompt test root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create pane-filter prompt test file.");

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
                  L"Failed to set local file-system plugin for pane-filter live interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewPaneFilterPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    closePrompt();
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for pane-filter prompt test.");

    struct PaneFilterSurfaceCycleResult final
    {
        HWND prompt                 = nullptr;
        bool opened                 = false;
        bool ownedByMainWindow      = false;
        bool capturedSnapshot       = false;
        bool setEnabled             = false;
        bool setText                = false;
        bool expandHelp             = false;
        bool capturedExpandedHelp   = false;
        bool capturedEditedSnapshot = false;
        bool actionIssued           = false;
        bool closed                 = false;
        FolderViewPaneFilterPromptDebugSnapshot snapshot{};
        FolderViewPaneFilterPromptDebugSnapshot expandedHelpSnapshot{};
        FolderViewPaneFilterPromptDebugSnapshot editedSnapshot{};
        std::optional<UiaValuePatternState> valueState;
        std::optional<UiaValuePatternState> editedValueState;
        std::optional<UiaTogglePatternState> toggleState;
        std::optional<UiaTogglePatternState> editedToggleState;
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
        std::optional<UiaNamedElementState> buttonState;
    };

    const auto runCycle = [&](bool enabled, std::wstring_view requestedText, bool accept, std::wstring_view failureContext) noexcept
    {
        PaneFilterSurfaceCycleResult cycle{};
        RunPaneFilterPromptModalCycle(mainWindow,
                                      IDM_LEFT_FILTER,
                                      [&](const HWND prompt) noexcept
        {
            cycle.prompt = prompt;
            cycle.opened = prompt != nullptr && IsWindow(prompt) != FALSE;
            if (! cycle.opened)
            {
                return;
            }

            cycle.ownedByMainWindow      = IsOwnedBy(prompt, mainWindow);
            cycle.capturedSnapshot       = DebugGetFolderViewPaneFilterPromptSnapshot(cycle.snapshot);
            cycle.valueState             = CollectVisibleDescendantValuePatternState(prompt, UIA_ComboBoxControlTypeId);
            cycle.toggleState            = CollectVisibleDescendantTogglePatternState(prompt);
            cycle.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
            cycle.buttonState            = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
            cycle.expandHelp             = DebugSetFolderViewPaneFilterPromptHelpExpanded(true);
            cycle.capturedExpandedHelp   = DebugGetFolderViewPaneFilterPromptSnapshot(cycle.expandedHelpSnapshot);
            cycle.setEnabled             = DebugSetFolderViewPaneFilterPromptEnabled(enabled);
            cycle.setText                = DebugSetFolderViewPaneFilterPromptTextAndNotify(requestedText);
            cycle.capturedEditedSnapshot = DebugGetFolderViewPaneFilterPromptSnapshot(cycle.editedSnapshot);
            cycle.editedValueState       = CollectVisibleDescendantValuePatternState(prompt, UIA_ComboBoxControlTypeId);
            cycle.editedToggleState      = CollectVisibleDescendantTogglePatternState(prompt);
            cycle.actionIssued           = accept ? DebugConfirmFolderViewPaneFilterPrompt() : DebugCancelFolderViewPaneFilterPrompt();
            if (cycle.actionIssued)
            {
                cycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
            }
        });

        state.Require(cycle.opened, std::format(L"Pane Filter prompt did not open for {}.", failureContext));
        state.Require(cycle.ownedByMainWindow, std::format(L"Pane Filter prompt should be owned by the main window during {}.", failureContext));
        state.Require(cycle.capturedSnapshot, std::format(L"Failed to capture pane-filter prompt snapshot during {}.", failureContext));
        if (! cycle.capturedSnapshot || ! state.failure.empty())
        {
            return cycle;
        }

        state.Require(cycle.snapshot.usesDxUiHost, std::format(L"Pane Filter prompt should use a DxUi host during {}.", failureContext));
        state.Require(cycle.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Pane Filter prompt should not expose more than one visible child window during {}; saw {}.",
                                  failureContext,
                                  cycle.snapshot.visibleChildWindowCount));
        state.Require(cycle.snapshot.historyComboVisible, std::format(L"Pane Filter prompt should use an editable history combo during {}.", failureContext));
        state.Require(! cycle.snapshot.historyButtonVisible,
                      std::format(L"Pane Filter prompt should not expose a separate History button during {}.", failureContext));
        state.Require(cycle.snapshot.historyItemCount > 0u,
                      std::format(L"Pane Filter prompt should load filter history into the combo during {}.", failureContext));

        state.Require(cycle.valueState.has_value(), std::format(L"Pane Filter prompt should expose a visible editable DX field during {}.", failureContext));
        if (cycle.valueState.has_value())
        {
            state.Require(! cycle.valueState->isReadOnly, std::format(L"Pane Filter prompt field should be editable during {}.", failureContext));
            state.Require(! cycle.valueState->name.empty(),
                          std::format(L"Pane Filter prompt editable DX field should expose a stable accessible name during {}.", failureContext));
            state.Require(cycle.valueState->value == cycle.snapshot.text,
                          std::format(L"Pane Filter prompt field value should match the live prompt snapshot during {}.", failureContext));
        }

        state.Require(cycle.toggleState.has_value(), std::format(L"Pane Filter prompt should expose a visible toggle control during {}.", failureContext));
        if (cycle.toggleState.has_value())
        {
            state.Require(! cycle.toggleState->name.empty(),
                          std::format(L"Pane Filter prompt toggle should expose a stable accessible name during {}.", failureContext));
            state.Require(cycle.toggleState->toggleState == (cycle.snapshot.enabled ? ToggleState_On : ToggleState_Off),
                          std::format(L"Pane Filter prompt toggle state should match the live prompt snapshot during {}.", failureContext));
        }

        state.Require(cycle.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation stats for pane-filter prompt during {}.", failureContext));
        if (cycle.uiaPatternStats.has_value())
        {
            state.Require(cycle.uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Pane Filter prompt should expose a visible UI Automation history combo during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Pane Filter prompt should expose ValuePattern for the filter field during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Pane Filter prompt should expose TogglePattern for the enable control during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Pane Filter prompt should expose a visible UI Automation command button during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Pane Filter prompt should expose InvokePattern for its visible DX command button during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->visibleElementCount >= 4u,
                          std::format(L"Pane Filter prompt should expose visible UI Automation descendants for the field, toggle, and commands during {}.",
                                      failureContext));
        }

        state.Require(cycle.buttonState.has_value(), std::format(L"Pane Filter prompt should expose a visible DX command button during {}.", failureContext));
        if (cycle.buttonState.has_value())
        {
            state.Require(! cycle.buttonState->name.empty(),
                          std::format(L"Pane Filter prompt visible DX command button should expose a stable accessible name during {}.", failureContext));
        }

        state.Require(cycle.expandHelp, std::format(L"Failed to expand pane-filter mask syntax help during {}.", failureContext));
        state.Require(cycle.capturedExpandedHelp, std::format(L"Failed to capture expanded pane-filter help snapshot during {}.", failureContext));
        if (cycle.capturedExpandedHelp)
        {
            state.Require(cycle.expandedHelpSnapshot.helpExpanded,
                          std::format(L"Pane Filter prompt mask syntax help should report expanded during {}.", failureContext));
            state.Require(
                cycle.expandedHelpSnapshot.commandButtonsFitInClient,
                std::format(L"Pane Filter prompt expanded mask syntax should keep command buttons inside the client area during {}.", failureContext));
        }

        state.Require(cycle.setEnabled, std::format(L"Failed to set the pane-filter toggle during {}.", failureContext));
        state.Require(cycle.setText, std::format(L"Failed to set pane-filter prompt text during {}.", failureContext));
        state.Require(cycle.capturedEditedSnapshot, std::format(L"Failed to capture edited pane-filter prompt snapshot during {}.", failureContext));
        const bool expectedEnabledAfterEdit = ! requestedText.empty();
        if (cycle.capturedEditedSnapshot)
        {
            state.Require(cycle.editedSnapshot.enabled == expectedEnabledAfterEdit,
                          std::format(L"Pane Filter prompt should map empty edited text to filter-off and non-empty edited text to filter-on during {}.",
                                      failureContext));
        }
        state.Require(cycle.editedValueState.has_value(),
                      std::format(L"Pane Filter prompt should keep the editable field visible after editing during {}.", failureContext));
        if (cycle.editedValueState.has_value())
        {
            state.Require(cycle.editedValueState->value == cycle.editedSnapshot.text,
                          std::format(L"Pane Filter prompt field value should track the edited prompt text during {}.", failureContext));
        }
        state.Require(cycle.editedToggleState.has_value(),
                      std::format(L"Pane Filter prompt should keep the toggle visible after editing during {}.", failureContext));
        if (cycle.editedToggleState.has_value())
        {
            state.Require(cycle.editedToggleState->toggleState == (expectedEnabledAfterEdit ? ToggleState_On : ToggleState_Off),
                          std::format(L"Pane Filter prompt toggle should match the edited enabled state during {}.", failureContext));
        }
        state.Require(cycle.actionIssued, std::format(L"Failed to {} the pane-filter prompt during {}.", accept ? L"confirm" : L"cancel", failureContext));
        state.Require(cycle.closed, std::format(L"Pane Filter prompt did not close after {} during {}.", accept ? L"confirm" : L"cancel", failureContext));
        return cycle;
    };

    static_cast<void>(runCycle(true, L"*.txt", true, L"initial confirm pass"));
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left),
                  L"Pane Filter prompt confirm should activate the filter during the initial confirm pass.");

    static_cast<void>(runCycle(false, L"*.log", false, L"reopened cancel pass"));
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left),
                  L"Canceling the reopened pane-filter prompt should preserve the previously confirmed active filter.");

    static_cast<void>(runCycle(true, L"", true, L"empty confirm pass"));
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left),
                  L"Confirming the pane-filter prompt with empty edited text should turn the filter off.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneFilterPromptLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    Trace(L"pane_filter_prompt_live_dx: begin");

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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_prompt_live_dx_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-filter live interaction test root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create pane-filter live interaction test file.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create pane-filter live interaction secondary test file.");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"pane_filter_prompt_live_dx: files-ready");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewPaneFilterPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    Trace(L"pane_filter_prompt_live_dx: set-path");
    closePrompt();
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    Trace(L"pane_filter_prompt_live_dx: path-issued");

    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for pane-filter live interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"pane_filter_prompt_live_dx: items-ready");

    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left),
                  L"Pane Filter prompt live interaction should start with the filter inactive.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto runInteraction = [&](std::wstring_view expectedInitialMask,
                                    const ToggleState expectedInitialToggleState,
                                    std::wstring_view expectedMask,
                                    bool accept,
                                    std::wstring_view failureContext) noexcept
    {
        struct PaneFilterLiveCycleResult final
        {
            HWND prompt             = nullptr;
            bool opened             = false;
            bool ownedByMainWindow  = false;
            bool exposesUiaProvider = false;
            bool capturedSnapshot   = false;
            FolderViewPaneFilterPromptDebugSnapshot snapshot{};
            bool setText       = false;
            bool enabledToggle = false;
            FolderViewPaneFilterPromptDebugSnapshot editedSnapshot{};
            bool recapturedSnapshot = false;
            bool actionIssued       = false;
            bool closed             = false;
        };

        PaneFilterLiveCycleResult cycle{};
        RunPaneFilterPromptModalCycle(mainWindow,
                                      IDM_LEFT_FILTER,
                                      [&](const HWND prompt) noexcept
        {
            cycle.prompt = prompt;
            cycle.opened = prompt != nullptr && IsWindow(prompt) != FALSE;
            if (! cycle.opened)
            {
                return;
            }

            cycle.ownedByMainWindow  = IsOwnedBy(prompt, mainWindow);
            cycle.exposesUiaProvider = WindowExposesUiaProvider(prompt);
            cycle.capturedSnapshot   = DebugGetFolderViewPaneFilterPromptSnapshot(cycle.snapshot);
            cycle.setText            = DebugSetFolderViewPaneFilterPromptText(expectedMask);
            cycle.enabledToggle      = DebugSetFolderViewPaneFilterPromptEnabled(true);
            cycle.recapturedSnapshot = DebugGetFolderViewPaneFilterPromptSnapshot(cycle.editedSnapshot);
            cycle.actionIssued       = accept ? DebugConfirmFolderViewPaneFilterPrompt() : DebugCancelFolderViewPaneFilterPrompt();
            if (cycle.actionIssued)
            {
                cycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
            }
        });

        state.Require(cycle.opened, std::format(L"Pane Filter prompt did not open for {}.", failureContext));
        state.Require(cycle.ownedByMainWindow, std::format(L"Pane Filter prompt should be owned by the main window during {}.", failureContext));
        state.Require(cycle.exposesUiaProvider, std::format(L"Pane Filter prompt should answer WM_GETOBJECT during {}.", failureContext));
        state.Require(cycle.capturedSnapshot, std::format(L"Failed to capture pane-filter prompt snapshot during {}.", failureContext));
        state.Require(cycle.snapshot.usesDxUiHost, std::format(L"Pane Filter prompt should use a DxUi host during {}.", failureContext));
        state.Require(cycle.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Pane Filter prompt should not expose more than one visible child window during {}; saw {}.",
                                  failureContext,
                                  cycle.snapshot.visibleChildWindowCount));
        state.Require(cycle.snapshot.text == expectedInitialMask,
                      std::format(L"Pane Filter prompt text should reopen with '{}' before {}.", expectedInitialMask, failureContext));
        state.Require(
            (cycle.snapshot.enabled ? ToggleState_On : ToggleState_Off) == expectedInitialToggleState,
            std::format(L"Pane Filter prompt toggle should reopen in state {} before {}.", static_cast<int>(expectedInitialToggleState), failureContext));
        state.Require(cycle.setText, std::format(L"Failed to set pane-filter prompt text via debug hook during {}.", failureContext));
        state.Require(cycle.enabledToggle, std::format(L"Failed to enable pane-filter prompt via debug hook during {}.", failureContext));
        state.Require(cycle.recapturedSnapshot, std::format(L"Failed to recapture pane-filter prompt snapshot during {}.", failureContext));
        state.Require(cycle.editedSnapshot.enabled,
                      std::format(L"Pane Filter prompt should be enabled after live DX toggle interaction during {}.", failureContext));
        state.Require(cycle.editedSnapshot.text == expectedMask,
                      std::format(L"Pane Filter prompt text should match the live DX edit interaction during {}.", failureContext));
        state.Require(cycle.actionIssued,
                      std::format(L"Pane Filter prompt {} action did not complete via debug hook during {}.", accept ? L"OK" : L"Cancel", failureContext));
        state.Require(cycle.closed, std::format(L"Pane Filter prompt did not close after {} during {}.", accept ? L"confirmation" : L"cancel", failureContext));
        return state.failure.empty();
    };

    state.Require(runInteraction(L"", ToggleState_Off, L"*.txt", false, L"live Cancel interaction"), L"Pane Filter prompt live Cancel interaction failed.");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left),
                  L"Pane Filter prompt live DX cancel interaction should leave the filter inactive.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runInteraction(L"", ToggleState_Off, L"*.txt", true, L"live OK interaction after Cancel reopen"),
                  L"Pane Filter prompt live OK interaction failed.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Pane Filter prompt live DX OK interaction should activate the filter.");
    Trace(L"pane_filter_prompt_live_dx: end");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneFilterPromptLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_prompt_churn_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-filter prompt churn test root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewPaneFilterPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for pane-filter prompt churn test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for pane-filter prompt churn test.");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Pane-filter churn test should start with the filter inactive.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const size_t mode       = cycle % 4u;
        const bool enabled      = mode == 0u || mode == 1u;
        const bool accept       = mode == 0u || mode == 2u;
        const bool filterBefore = g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left);

        struct WorkerResult final
        {
            HWND prompt            = nullptr;
            bool sawPrompt         = false;
            bool ownedByMainWindow = false;
            bool capturedSnapshot  = false;
            bool setEnabled        = false;
            bool setText           = false;
            bool actionIssued      = false;
            std::optional<UiaDescendantPatternStats> uiaPatternStats;
            std::optional<UiaValuePatternState> valueState;
            std::optional<UiaTogglePatternState> toggleState;
            std::wstring buttonName;
            FolderViewPaneFilterPromptDebugSnapshot snapshot{};
        } workerResult{};

        std::jthread worker([&](std::stop_token) noexcept
        {
            workerResult.prompt = WaitForWindow(
                [mainWindow]() noexcept -> HWND
            {
                const HWND dlg = GetFolderViewPaneFilterPromptHandle();
                if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
                {
                    return nullptr;
                }

                return dlg;
            },
                SelfTest::Scale(5000ms));
            workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
            if (! workerResult.sawPrompt)
            {
                return;
            }

            workerResult.ownedByMainWindow = IsOwnedBy(workerResult.prompt, mainWindow);
            workerResult.capturedSnapshot  = DebugGetFolderViewPaneFilterPromptSnapshot(workerResult.snapshot);
            workerResult.uiaPatternStats   = CollectVisibleUiaDescendantPatternStats(workerResult.prompt);
            workerResult.valueState        = CollectVisibleDescendantValuePatternState(workerResult.prompt, UIA_ComboBoxControlTypeId);
            workerResult.toggleState       = CollectVisibleDescendantTogglePatternState(workerResult.prompt);
            if (const auto buttonState = CollectVisibleDescendantNamedElementState(workerResult.prompt, UIA_ButtonControlTypeId); buttonState.has_value())
            {
                workerResult.buttonName = buttonState->name;
            }
            workerResult.setEnabled   = DebugSetFolderViewPaneFilterPromptEnabled(enabled);
            workerResult.setText      = DebugSetFolderViewPaneFilterPromptText(L"*.txt");
            workerResult.actionIssued = accept ? DebugConfirmFolderViewPaneFilterPrompt() : DebugCancelFolderViewPaneFilterPrompt();
        });

        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        worker.join();

        state.Require(workerResult.sawPrompt, std::format(L"Pane-filter prompt did not open during cycle {}.", cycle));
        state.Require(workerResult.ownedByMainWindow, std::format(L"Pane-filter prompt should be owned by the main window during cycle {}.", cycle));
        state.Require(workerResult.capturedSnapshot, std::format(L"Failed to capture pane-filter prompt snapshot during cycle {}.", cycle));
        state.Require(workerResult.snapshot.usesDxUiHost, std::format(L"Pane-filter prompt should use a DxUi host during cycle {}.", cycle));
        state.Require(workerResult.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Pane-filter prompt should not expose more than one visible child window during cycle {}; saw {}.",
                                  cycle,
                                  workerResult.snapshot.visibleChildWindowCount));
        state.Require(workerResult.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect UI Automation stats for pane-filter prompt during cycle {}.", cycle));
        if (workerResult.uiaPatternStats.has_value())
        {
            state.Require(workerResult.uiaPatternStats->visibleElementCount >= 4u,
                          std::format(L"Pane-filter prompt should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats->editControlCount + workerResult.uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Pane-filter prompt should expose a visible editable field during cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Pane-filter prompt should expose ValuePattern during cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats->togglePatternCount > 0u,
                          std::format(L"Pane-filter prompt should expose TogglePattern during cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Pane-filter prompt should expose visible command buttons during cycle {}.", cycle));
        }
        state.Require(workerResult.valueState.has_value(),
                      std::format(L"Failed to collect pane-filter prompt visible DX edit ValuePattern state during cycle {}.", cycle));
        if (workerResult.valueState.has_value())
        {
            state.Require(! workerResult.valueState->isReadOnly,
                          std::format(L"Pane-filter prompt visible DX edit surface should remain editable during cycle {}.", cycle));
            state.Require(! workerResult.valueState->name.empty(),
                          std::format(L"Pane-filter prompt visible DX edit surface should expose a stable accessible name during cycle {}.", cycle));
        }
        state.Require(workerResult.toggleState.has_value(),
                      std::format(L"Failed to collect pane-filter prompt visible DX toggle state during cycle {}.", cycle));
        if (workerResult.toggleState.has_value())
        {
            state.Require(! workerResult.toggleState->name.empty(),
                          std::format(L"Pane-filter prompt visible DX toggle should expose a stable accessible name during cycle {}.", cycle));
        }
        state.Require(! workerResult.buttonName.empty(),
                      std::format(L"Pane-filter prompt visible DX command button should expose a stable accessible name during cycle {}.", cycle));
        state.Require(workerResult.setEnabled, std::format(L"Failed to set pane-filter enabled state during cycle {}.", cycle));
        state.Require(workerResult.setText, std::format(L"Failed to set pane-filter text during cycle {}.", cycle));
        state.Require(workerResult.actionIssued, std::format(L"Failed to close the pane-filter prompt through the DX action path during cycle {}.", cycle));

        if (accept)
        {
            if (enabled)
            {
                state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left),
                              std::format(L"Pane filter should be active after enabled accept cycle {}.", cycle));
                state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt"}, SelfTest::Scale(3000ms)),
                              std::format(L"Pane items should reflect the active *.txt filter after cycle {}.", cycle));
            }
            else
            {
                state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left),
                              std::format(L"Pane filter should be inactive after disabled accept cycle {}.", cycle));
                state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log"}, SelfTest::Scale(3000ms)),
                              std::format(L"Pane items should reflect the disabled filter after cycle {}.", cycle));
            }
        }
        else
        {
            state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) == filterBefore,
                          std::format(L"Pane-filter cancel cycle {} should not change the active filter state.", cycle));
            if (filterBefore)
            {
                state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt"}, SelfTest::Scale(3000ms)),
                              std::format(L"Pane-filter cancel cycle {} should preserve the filtered item set.", cycle));
            }
            else
            {
                state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log"}, SelfTest::Scale(3000ms)),
                              std::format(L"Pane-filter cancel cycle {} should preserve the unfiltered item set.", cycle));
            }
        }

        const HWND lingeringPrompt = GetFolderViewPaneFilterPromptHandle();
        state.Require(lingeringPrompt == nullptr || IsWindow(lingeringPrompt) == FALSE,
                      std::format(L"Pane-filter prompt should not remain open after cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closePrompt();
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestFilterWatermarkBadgeForEmptyState(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"filter_watermark_empty_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create filter-watermark test root.");

    const std::filesystem::path emptyChild    = root / L"empty";
    const std::filesystem::path nonEmptyChild = root / L"non_empty";
    state.Require(SelfTest::EnsureDirectory(emptyChild), L"Failed to create empty child folder.");
    state.Require(SelfTest::EnsureDirectory(nonEmptyChild), L"Failed to create non-empty child folder.");
    state.Require(SelfTest::WriteTextFile(nonEmptyChild / L"a.txt", "a"), L"Failed to create a.txt.");

    const std::wstring leftPluginBefore                   = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto clearLeftPaneOverlayState = []() noexcept
    {
        g_folderWindow.DismissPaneAlertOverlay(FolderWindow::Pane::Left);
        g_folderWindow.SetPaneEmptyStateMessage(FolderWindow::Pane::Left, {});
        g_folderWindow.SetPaneBackgroundWatermark(FolderWindow::Pane::Left, {}, false);
    };

    clearLeftPaneOverlayState();
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for filter-watermark empty-state test.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::atomic<uint32_t> enumEmpty{0};
    std::atomic<uint32_t> enumNonEmpty{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, emptyChild))
        {
            enumEmpty.fetch_add(1u, std::memory_order_release);
        }
        else if (OrdinalString::EqualsNoCasePath(folder, nonEmptyChild))
        {
            enumNonEmpty.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    const auto describeLeftPaneState = []() noexcept
    {
        const std::optional<std::filesystem::path> path = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
        const FolderView::NameFilterState filterState   = g_folderWindow.DebugGetNameFilterState(FolderWindow::Pane::Left);
        FolderView::AlertOverlayDebugSnapshot alert{};
        const bool alertSnapshot = g_folderWindow.DebugGetPaneAlertSnapshot(FolderWindow::Pane::Left, alert);
        return std::format(L"path='{}', itemCount={}, filterActive={}, filterEnabled={}, filterText='{}', emptyActive={}, watermark={}, "
                           L"alertSnapshot={}, alertVisible={}, alertKind={}, alertSeverity={}, alertClosable={}, alertBlocksInput={}, "
                           L"alertHr=0x{:08X}, alertTitle='{}', alertMessage='{}'",
                           path.has_value() ? path->wstring() : std::wstring{L"<none>"},
                           g_folderWindow.DebugGetItemCount(FolderWindow::Pane::Left),
                           g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) ? 1 : 0,
                           filterState.enabled ? 1 : 0,
                           filterState.text,
                           g_folderWindow.DebugIsEmptyFolderStateActive(FolderWindow::Pane::Left) ? 1 : 0,
                           static_cast<int>(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left)),
                           alertSnapshot ? 1 : 0,
                           alert.visible ? 1 : 0,
                           static_cast<int>(alert.kind),
                           static_cast<int>(alert.severity),
                           alert.closable ? 1 : 0,
                           alert.blocksInput ? 1 : 0,
                           static_cast<unsigned long>(alert.hr),
                           alert.title,
                           alert.message);
    };
    const auto waitForLeftPaneState = [&](auto&& predicate, std::chrono::milliseconds timeout) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + timeout;
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (predicate())
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        return predicate();
    };

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, emptyChild);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, emptyChild, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path for filter-watermark test.");
    state.Require(WaitForAtomicAtLeast(enumEmpty, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for empty folder in filter-watermark test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const uint32_t enumBeforeReset = enumEmpty.load(std::memory_order_acquire);
    clearLeftPaneOverlayState();
    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(WaitForAtomicAtLeast(enumEmpty, enumBeforeReset + 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh for empty folder after clearing inherited pane visibility state.");
    state.Require(waitForLeftPaneState(
                      []() noexcept
    {
        return ! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left) && g_folderWindow.DebugIsEmptyFolderStateActive(FolderWindow::Pane::Left) &&
               g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::None;
    },
                      SelfTest::Scale(3000ms)),
                  std::format(L"Expected unfiltered empty-folder state before applying pane filter; {}.", describeLeftPaneState()));
    if (! state.failure.empty())
    {
        return false;
    }

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after OK.");
    }

    state.Require(WaitForAtomicAtLeast(enumEmpty, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh for empty folder after applying pane filter.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter active after applying pane filter on empty folder.");
    state.Require(! g_folderWindow.DebugIsEmptyFolderStateActive(FolderWindow::Pane::Left),
                  L"Filter-empty state should suppress the generic empty-folder placeholder.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::Background,
                  L"Expected full-pane funnel placeholder when filter is active and no rows are visible.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, nonEmptyChild);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, nonEmptyChild, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set left pane path to non-empty folder in filter-watermark test.");
    state.Require(WaitForAtomicAtLeast(enumNonEmpty, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for non-empty folder in filter-watermark test.");

    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter inactive by default on non-empty folder.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt visible in non-empty folder.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::None,
                  L"Expected no filter watermark for non-empty folder with filter disabled.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open for non-empty folder.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after OK for non-empty folder.");
    }

    state.Require(WaitForAtomicAtLeast(enumNonEmpty, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh for non-empty folder after applying pane filter.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter active after applying pane filter on non-empty folder.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt still visible under *.txt filter.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::Background,
                  L"Expected background watermark when filter is active and items are visible.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, false, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not reopen for disable.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after disabling.");
    }

    state.Require(WaitForAtomicAtLeast(enumNonEmpty, 3u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh for non-empty folder after disabling pane filter.");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Expected filter inactive after disabling.");
    state.Require(g_folderWindow.DebugGetFilterWatermarkVisualMode(FolderWindow::Pane::Left) == FolderView::FilterWatermarkVisualMode::None,
                  L"Expected no watermark after disabling.");

    return state.failure.empty();
}

struct SelectionMaskDialogAutomationState final
{
    std::atomic<bool> sawDialog{false};
    std::atomic<bool> closed{false};

    SelectionMaskDialogAutomationState()                                                     = default;
    SelectionMaskDialogAutomationState(const SelectionMaskDialogAutomationState&)            = delete;
    SelectionMaskDialogAutomationState& operator=(const SelectionMaskDialogAutomationState&) = delete;
    SelectionMaskDialogAutomationState(SelectionMaskDialogAutomationState&&)                 = delete;
    SelectionMaskDialogAutomationState& operator=(SelectionMaskDialogAutomationState&&)      = delete;
};

void AutomateSelectionMaskDialog(
    HWND mainWindow, SelectionMaskDialogAutomationState& dlgState, std::wstring_view caption, std::wstring_view maskText, bool accept) noexcept
{
    const HWND dlg = WaitForWindow(
        [mainWindow]() noexcept -> HWND
    {
        const HWND dlg = GetFolderViewSelectionMaskPromptHandle();
        if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
        {
            return nullptr;
        }
        return dlg;
    },
        SelfTest::Scale(std::chrono::milliseconds{10000}));
    if (! dlg)
    {
        return;
    }

    dlgState.sawDialog.store(true, std::memory_order_release);

    FolderViewSelectionMaskPromptDebugSnapshot snapshot{};
    if (! DebugGetFolderViewSelectionMaskPromptSnapshot(snapshot))
    {
        return;
    }
    dlgState.sawDialog.store(snapshot.title == std::wstring(caption), std::memory_order_release);

    static_cast<void>(DebugSetFolderViewSelectionMaskPromptText(maskText));
    static_cast<void>(accept ? DebugConfirmFolderViewSelectionMaskPrompt() : DebugCancelFolderViewSelectionMaskPrompt());

    bool closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    if (! closed)
    {
        PostMessageW(dlg, WM_CLOSE, 0, 0);
        PostMessageW(dlg, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(dlg, WM_KEYUP, VK_ESCAPE, 0);
        closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    }

    dlgState.closed.store(closed, std::memory_order_release);
}

void AutomateChangeCasePrompt(
    HWND mainWindow, ChangeCasePromptAutomationState& dlgState, size_t styleIndex, size_t targetIndex, bool includeSubdirs, bool accept) noexcept
{
    const HWND dlg = WaitForWindow(
        [mainWindow]() noexcept -> HWND
    {
        const HWND prompt = GetFolderViewChangeCasePromptHandle();
        if (! prompt || IsWindow(prompt) == FALSE || ! IsOwnedBy(prompt, mainWindow))
        {
            return nullptr;
        }
        return prompt;
    },
        SelfTest::Scale(std::chrono::milliseconds{10000}));
    if (! dlg)
    {
        return;
    }

    dlgState.sawDialog.store(true, std::memory_order_release);

    FolderViewChangeCasePromptDebugSnapshot snapshot{};
    if (! DebugGetFolderViewChangeCasePromptSnapshot(snapshot))
    {
        return;
    }
    dlgState.includeEnabled.store(snapshot.includeSubdirsEnabled, std::memory_order_release);

    static_cast<void>(DebugSetFolderViewChangeCasePromptSelections(styleIndex, targetIndex, includeSubdirs));
    static_cast<void>(accept ? DebugConfirmFolderViewChangeCasePrompt() : DebugCancelFolderViewChangeCasePrompt());

    bool closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    if (! closed)
    {
        PostMessageW(dlg, WM_CLOSE, 0, 0);
        PostMessageW(dlg, WM_KEYDOWN, VK_ESCAPE, 0);
        PostMessageW(dlg, WM_KEYUP, VK_ESCAPE, 0);
        closed = WaitForWindowClosed(dlg, SelfTest::Scale(std::chrono::milliseconds{5000}));
    }

    dlgState.closed.store(closed, std::memory_order_release);
}

[[nodiscard]] bool TestSelectionMaskPromptUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_mask_prompt_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-mask prompt test root.");

    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
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

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for selection-mask prompt test.");
    struct SelectionMaskSurfaceCycleResult final
    {
        HWND prompt                 = nullptr;
        bool opened                 = false;
        bool ownedByMainWindow      = false;
        bool capturedSnapshot       = false;
        bool setText                = false;
        bool capturedEditedSnapshot = false;
        bool actionIssued           = false;
        bool closed                 = false;
        FolderViewSelectionMaskPromptDebugSnapshot snapshot{};
        FolderViewSelectionMaskPromptDebugSnapshot editedSnapshot{};
        std::optional<UiaValuePatternState> valueState;
        std::optional<UiaValuePatternState> editedValueState;
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
        std::optional<UiaNamedElementState> buttonState;
    };

    const auto runCycle = [&](bool accept, std::wstring_view requestedText, std::wstring_view failureContext) noexcept
    {
        SelectionMaskSurfaceCycleResult cycle{};
        RunSelectionMaskPromptModalCycle(mainWindow,
                                         IDM_PANE_SELECTION_SELECT_DIALOG,
                                         [&](const HWND prompt) noexcept
        {
            cycle.prompt = prompt;
            cycle.opened = prompt != nullptr && IsWindow(prompt) != FALSE;
            if (! cycle.opened)
            {
                return;
            }

            cycle.ownedByMainWindow      = IsOwnedBy(prompt, mainWindow);
            cycle.capturedSnapshot       = DebugGetFolderViewSelectionMaskPromptSnapshot(cycle.snapshot);
            cycle.valueState             = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
            cycle.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
            cycle.buttonState            = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
            cycle.setText                = DebugSetFolderViewSelectionMaskPromptText(requestedText);
            cycle.capturedEditedSnapshot = DebugGetFolderViewSelectionMaskPromptSnapshot(cycle.editedSnapshot);
            cycle.editedValueState       = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
            cycle.actionIssued           = accept ? DebugConfirmFolderViewSelectionMaskPrompt() : DebugCancelFolderViewSelectionMaskPrompt();
            if (cycle.actionIssued)
            {
                cycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
            }
        });

        state.Require(cycle.opened, std::format(L"Selection-mask prompt did not open for {}.", failureContext));
        state.Require(cycle.ownedByMainWindow, std::format(L"Selection-mask prompt should be owned by the main window during {}.", failureContext));
        state.Require(cycle.capturedSnapshot, std::format(L"Failed to capture selection-mask prompt snapshot during {}.", failureContext));
        if (! cycle.capturedSnapshot || ! state.failure.empty())
        {
            return cycle;
        }

        state.Require(cycle.snapshot.usesDxUiHost, std::format(L"Selection-mask prompt should use a DxUi host during {}.", failureContext));
        state.Require(cycle.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Selection-mask prompt should not expose more than one visible child window during {}; saw {}.",
                                  failureContext,
                                  cycle.snapshot.visibleChildWindowCount));
        state.Require(cycle.snapshot.title == LoadStringResource(nullptr, IDS_CAPTION_SELECTION_MASK_SELECT),
                      std::format(L"Selection-mask prompt should expose the Select caption through the DX surface during {}.", failureContext));
        state.Require(cycle.valueState.has_value(), std::format(L"Selection-mask prompt should expose a visible editable DX field during {}.", failureContext));
        if (cycle.valueState.has_value())
        {
            state.Require(! cycle.valueState->isReadOnly, std::format(L"Selection-mask prompt field should be editable during {}.", failureContext));
            state.Require(! cycle.valueState->name.empty(),
                          std::format(L"Selection-mask prompt editable DX field should expose a stable accessible name during {}.", failureContext));
            state.Require(cycle.valueState->value == cycle.snapshot.text,
                          std::format(L"Selection-mask prompt field value should match the live prompt snapshot during {}.", failureContext));
        }

        state.Require(cycle.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation stats for selection-mask prompt during {}.", failureContext));
        if (cycle.uiaPatternStats.has_value())
        {
            state.Require(cycle.uiaPatternStats->editControlCount + cycle.uiaPatternStats->comboBoxControlCount > 0u,
                          std::format(L"Selection-mask prompt should expose a visible UI Automation editable field during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Selection-mask prompt should expose ValuePattern for the mask field during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Selection-mask prompt should expose a visible UI Automation command button during {}.", failureContext));
            state.Require(cycle.uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Selection-mask prompt should expose InvokePattern for its visible DX command button during {}.", failureContext));
            state.Require(
                cycle.uiaPatternStats->visibleElementCount >= 3u,
                std::format(L"Selection-mask prompt should expose visible UI Automation descendants for the field and commands during {}.", failureContext));
        }

        state.Require(cycle.buttonState.has_value(),
                      std::format(L"Selection-mask prompt should expose a visible DX command button during {}.", failureContext));
        if (cycle.buttonState.has_value())
        {
            state.Require(! cycle.buttonState->name.empty(),
                          std::format(L"Selection-mask prompt visible DX command button should expose a stable accessible name during {}.", failureContext));
        }

        state.Require(cycle.setText, std::format(L"Failed to set selection-mask prompt text during {}.", failureContext));
        state.Require(cycle.capturedEditedSnapshot, std::format(L"Failed to capture edited selection-mask prompt snapshot during {}.", failureContext));
        state.Require(cycle.editedValueState.has_value(),
                      std::format(L"Selection-mask prompt should keep the editable field visible after editing during {}.", failureContext));
        if (cycle.editedValueState.has_value() && cycle.capturedEditedSnapshot)
        {
            state.Require(cycle.editedValueState->value == cycle.editedSnapshot.text,
                          std::format(L"Selection-mask prompt field value should track the edited prompt text during {}.", failureContext));
        }
        state.Require(cycle.actionIssued, std::format(L"Failed to {} the selection-mask prompt during {}.", accept ? L"confirm" : L"cancel", failureContext));
        state.Require(cycle.closed, std::format(L"Selection-mask prompt did not close after {} during {}.", accept ? L"confirm" : L"cancel", failureContext));
        return cycle;
    };

    static_cast<void>(runCycle(true, L"*.txt", L"initial confirm pass"));
    if (! state.failure.empty())
    {
        return false;
    }
    Trace(L"pane_filter_prompt_live_dx: pre-interaction");

    static_cast<void>(runCycle(false, L"*.log", L"reopened cancel pass"));

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionMaskPromptLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_mask_prompt_live_dx_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-mask prompt live interaction test root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewSelectionMaskPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for selection-mask prompt live interaction test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for selection-mask prompt live interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto runInteraction = [&](const UINT commandId,
                                    const UINT expectedCaptionId,
                                    [[maybe_unused]] std::wstring_view expectedInitialMask,
                                    std::wstring_view mask,
                                    const auto& verifySelection,
                                    const auto& verifyCancelSelection,
                                    std::wstring_view failureContext) noexcept
    {
        const auto runPass = [&](bool accept,
                                 bool useLiveValueInteraction,
                                 bool requireExpectedInitialValue,
                                 std::wstring_view expectedInitialValue,
                                 std::wstring_view passContext,
                                 const auto& verifyAfterClose,
                                 std::wstring* observedInitialValue) noexcept
        {
            closePrompt();
            FocusFolderViewPane(FolderWindow::Pane::Left);
            struct WorkerResult final
            {
                HWND prompt                 = nullptr;
                bool sawPrompt              = false;
                bool ownedByMainWindow      = false;
                bool capturedSnapshot       = false;
                bool capturedEditedSnapshot = false;
                bool setValue               = false;
                bool editUpdated            = false;
                bool invokedAction          = false;
                bool closedAfterAction      = false;
                std::optional<UiaValuePatternState> initialValueState;
                FolderViewSelectionMaskPromptDebugSnapshot snapshot{};
                FolderViewSelectionMaskPromptDebugSnapshot editedSnapshot{};
                std::wstring editName;
            } workerResult{};

            const std::wstring expectedCaption = LoadStringResource(nullptr, expectedCaptionId);
            std::jthread worker([&](std::stop_token) noexcept
            {
                workerResult.prompt = WaitForWindow(
                    [mainWindow]() noexcept -> HWND
                {
                    const HWND dlg = GetFolderViewSelectionMaskPromptHandle();
                    if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
                    {
                        return nullptr;
                    }
                    return dlg;
                },
                    SelfTest::Scale(std::chrono::milliseconds{5000}));
                workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
                if (! workerResult.sawPrompt)
                {
                    return;
                }

                workerResult.ownedByMainWindow = IsOwnedBy(workerResult.prompt, mainWindow);
                workerResult.capturedSnapshot  = DebugGetFolderViewSelectionMaskPromptSnapshot(workerResult.snapshot);
                workerResult.initialValueState = CollectVisibleDescendantValuePatternState(workerResult.prompt, UIA_EditControlTypeId);
                if (! workerResult.initialValueState.has_value())
                {
                    return;
                }

                workerResult.editName = workerResult.initialValueState->name;
                workerResult.setValue = useLiveValueInteraction
                                            ? SetVisibleDescendantValueByName(workerResult.prompt, UIA_EditControlTypeId, workerResult.editName, mask)
                                            : DebugSetFolderViewSelectionMaskPromptText(mask);
                if (! workerResult.setValue)
                {
                    return;
                }

                if (useLiveValueInteraction)
                {
                    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{3000});
                    while (std::chrono::steady_clock::now() < deadline)
                    {
                        PumpPendingMessages();
                        const auto valueState =
                            CollectVisibleDescendantValuePatternStateByName(workerResult.prompt, UIA_EditControlTypeId, workerResult.editName);
                        if (valueState.has_value() && valueState->value == mask)
                        {
                            workerResult.editUpdated = true;
                            break;
                        }
                        std::this_thread::sleep_for(std::chrono::milliseconds{20});
                    }
                    if (! workerResult.editUpdated)
                    {
                        const auto valueState =
                            CollectVisibleDescendantValuePatternStateByName(workerResult.prompt, UIA_EditControlTypeId, workerResult.editName);
                        workerResult.editUpdated = valueState.has_value() && valueState->value == mask;
                    }
                }
                else
                {
                    workerResult.editUpdated = true;
                }
                if (! workerResult.editUpdated)
                {
                    return;
                }

                workerResult.capturedEditedSnapshot = DebugGetFolderViewSelectionMaskPromptSnapshot(workerResult.editedSnapshot);
                workerResult.invokedAction          = accept ? DebugConfirmFolderViewSelectionMaskPrompt() : DebugCancelFolderViewSelectionMaskPrompt();
                if (! workerResult.invokedAction)
                {
                    return;
                }
                workerResult.closedAfterAction = WaitForWindowClosed(workerResult.prompt, SelfTest::Scale(std::chrono::milliseconds{3000}));
            });

            SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
            worker.join();

            state.Require(workerResult.sawPrompt, std::format(L"Selection-mask prompt did not open for {}.", passContext));
            state.Require(workerResult.ownedByMainWindow, std::format(L"Selection-mask prompt should be owned by the main window during {}.", passContext));
            state.Require(workerResult.capturedSnapshot, std::format(L"Failed to capture selection-mask prompt snapshot during {}.", passContext));
            if (! workerResult.capturedSnapshot || ! state.failure.empty())
            {
                return false;
            }

            state.Require(workerResult.snapshot.usesDxUiHost, std::format(L"Selection-mask prompt should use a DxUi host during {}.", passContext));
            state.Require(workerResult.snapshot.visibleChildWindowCount <= 1u,
                          std::format(L"Selection-mask prompt should not expose more than one visible child window during {}; saw {}.",
                                      passContext,
                                      workerResult.snapshot.visibleChildWindowCount));
            state.Require(workerResult.snapshot.title == expectedCaption,
                          std::format(L"Selection-mask prompt should expose the expected caption during {}.", passContext));
            state.Require(workerResult.initialValueState.has_value(),
                          std::format(L"Selection-mask prompt should expose a visible editable DX field during {}.", passContext));
            if (! workerResult.initialValueState.has_value() || ! state.failure.empty())
            {
                return false;
            }
            if (observedInitialValue)
            {
                *observedInitialValue = workerResult.initialValueState->value;
            }

            state.Require(! workerResult.initialValueState->name.empty(),
                          std::format(L"Selection-mask prompt DX edit should expose a stable accessible name during {}.", passContext));
            state.Require(! workerResult.initialValueState->isReadOnly,
                          std::format(L"Selection-mask prompt DX edit should remain editable during {}.", passContext));
            if (requireExpectedInitialValue)
            {
                state.Require(workerResult.initialValueState->value == expectedInitialValue,
                              std::format(L"Selection-mask prompt DX edit should reopen with '{}' before {}.", expectedInitialValue, passContext));
            }
            state.Require(workerResult.setValue,
                          std::format(L"Failed to set selection-mask prompt DX edit '{}' during {}.", workerResult.editName, passContext));
            state.Require(workerResult.editUpdated,
                          std::format(L"Selection-mask prompt text did not update after {} interaction during {}.",
                                      useLiveValueInteraction ? L"live UIA edit" : L"debug seed",
                                      passContext));
            state.Require(workerResult.capturedEditedSnapshot, std::format(L"Failed to recapture selection-mask prompt snapshot during {}.", passContext));

            state.Require(
                workerResult.invokedAction,
                std::format(L"Selection-mask prompt did not close through the debug {} path during {}.", accept ? L"confirm" : L"cancel", passContext));
            state.Require(workerResult.closedAfterAction,
                          std::format(L"Selection-mask prompt did not close after the debug {} path during {}.", accept ? L"confirm" : L"cancel", passContext));
            if (! state.failure.empty())
            {
                return false;
            }

            verifyAfterClose();
            return state.failure.empty();
        };

        std::wstring baselineInitialMask;
        state.Require(runPass(false, true, false, L"", std::format(L"{} cancel path", failureContext), verifyCancelSelection, &baselineInitialMask),
                      std::format(L"Selection-mask prompt cancel path failed during {}.", failureContext));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(runPass(true, false, true, baselineInitialMask, std::format(L"{} confirm path", failureContext), verifySelection, nullptr),
                      std::format(L"Selection-mask prompt confirm path failed during {}.", failureContext));
        return state.failure.empty();
    };

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"b.log"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log selected before live Select interaction.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected item before live Select interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runInteraction(IDM_PANE_SELECTION_SELECT_DIALOG,
                                 IDS_CAPTION_SELECTION_MASK_SELECT,
                                 L"",
                                 L"*.txt",
                                 [&]() noexcept
    {
        state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected after live Select interaction.");
        state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt selected after live Select interaction.");
        state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                      L"Expected b.log to remain selected after live Select interaction.");
        state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u, L"Expected three selected items after live Select interaction.");
    },
                                 [&]() noexcept
    {
        state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                      L"Expected a.txt unselected after canceling the live Select interaction.");
        state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                      L"Expected c.txt unselected after canceling the live Select interaction.");
        state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                      L"Expected b.log to remain selected after canceling the live Select interaction.");
        state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                      L"Expected one selected item after canceling the live Select interaction.");
    },
                                 L"live Select interaction"),
                  L"Selection-mask Select live interaction failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runInteraction(IDM_PANE_SELECTION_UNSELECT_DIALOG,
                                 IDS_CAPTION_SELECTION_MASK_UNSELECT,
                                 L"",
                                 L"*.txt",
                                 [&]() noexcept
    {
        state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected after live Unselect interaction.");
        state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt unselected after live Unselect interaction.");
        state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                      L"Expected b.log to remain selected after live Unselect interaction.");
        state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected item after live Unselect interaction.");
    },
                                 [&]() noexcept
    {
        state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                      L"Expected a.txt to remain selected after canceling the live Unselect interaction.");
        state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                      L"Expected c.txt to remain selected after canceling the live Unselect interaction.");
        state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                      L"Expected b.log to remain selected after canceling the live Unselect interaction.");
        state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u,
                      L"Expected three selected items after canceling the live Unselect interaction.");
    },
                                 L"live Unselect interaction"),
                  L"Selection-mask Unselect live interaction failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionMaskPromptLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_mask_prompt_churn_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-mask prompt churn test root.");
    state.Require(SelfTest::WriteTextFile(root / L"a.txt", "a"), L"Failed to create a.txt.");
    state.Require(SelfTest::WriteTextFile(root / L"b.log", "b"), L"Failed to create b.log.");
    state.Require(SelfTest::WriteTextFile(root / L"c.txt", "c"), L"Failed to create c.txt.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewSelectionMaskPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set pane path for selection-mask churn test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for selection-mask churn test.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"b.log"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log selected before selection-mask churn.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected item before selection-mask churn.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectCaption   = LoadStringResource(nullptr, IDS_CAPTION_SELECTION_MASK_SELECT);
    const std::wstring unselectCaption = LoadStringResource(nullptr, IDS_CAPTION_SELECTION_MASK_UNSELECT);
    constexpr size_t kCycles           = 12u;

    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const size_t mode               = cycle % 4u;
        const bool useSelectCommand     = mode == 0u || mode == 3u;
        const bool accept               = mode == 0u || mode == 2u;
        const UINT commandId            = useSelectCommand ? IDM_PANE_SELECTION_SELECT_DIALOG : IDM_PANE_SELECTION_UNSELECT_DIALOG;
        const std::wstring_view caption = useSelectCommand ? std::wstring_view{selectCaption} : std::wstring_view{unselectCaption};
        const bool hadATxtSelected      = g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt");
        const bool hadBLogSelected      = g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log");
        const bool hadCTxtSelected      = g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt");
        const size_t selectedBefore     = g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left);

        struct WorkerResult final
        {
            HWND prompt            = nullptr;
            bool sawPrompt         = false;
            bool ownedByMainWindow = false;
            bool capturedSnapshot  = false;
            bool setText           = false;
            bool actionIssued      = false;
            std::optional<UiaDescendantPatternStats> uiaPatternStats;
            std::optional<UiaValuePatternState> valueState;
            std::wstring buttonName;
            FolderViewSelectionMaskPromptDebugSnapshot snapshot{};
        } workerResult{};

        std::jthread worker([&](std::stop_token) noexcept
        {
            workerResult.prompt = WaitForWindow(
                [mainWindow]() noexcept -> HWND
            {
                const HWND dlg = GetFolderViewSelectionMaskPromptHandle();
                if (! dlg || IsWindow(dlg) == FALSE || ! IsOwnedBy(dlg, mainWindow))
                {
                    return nullptr;
                }
                return dlg;
            },
                SelfTest::Scale(5000ms));
            workerResult.sawPrompt = workerResult.prompt != nullptr && IsWindow(workerResult.prompt) != FALSE;
            if (! workerResult.sawPrompt)
            {
                return;
            }

            workerResult.ownedByMainWindow = IsOwnedBy(workerResult.prompt, mainWindow);
            workerResult.capturedSnapshot  = DebugGetFolderViewSelectionMaskPromptSnapshot(workerResult.snapshot);
            workerResult.uiaPatternStats   = CollectVisibleUiaDescendantPatternStats(workerResult.prompt);
            workerResult.valueState        = CollectVisibleDescendantValuePatternState(workerResult.prompt, UIA_EditControlTypeId);
            if (const auto buttonState = CollectVisibleDescendantNamedElementState(workerResult.prompt, UIA_ButtonControlTypeId); buttonState.has_value())
            {
                workerResult.buttonName = buttonState->name;
            }
            workerResult.setText      = DebugSetFolderViewSelectionMaskPromptText(L"*.txt");
            workerResult.actionIssued = accept ? DebugConfirmFolderViewSelectionMaskPrompt() : DebugCancelFolderViewSelectionMaskPrompt();
        });

        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
        worker.join();

        state.Require(workerResult.sawPrompt, std::format(L"Selection-mask prompt did not open during cycle {}.", cycle));
        state.Require(workerResult.ownedByMainWindow, std::format(L"Selection-mask prompt should be owned by the main window during cycle {}.", cycle));
        state.Require(workerResult.capturedSnapshot, std::format(L"Failed to capture selection-mask prompt snapshot during cycle {}.", cycle));
        state.Require(workerResult.snapshot.usesDxUiHost, std::format(L"Selection-mask prompt should use a DxUi host during cycle {}.", cycle));
        state.Require(workerResult.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Selection-mask prompt should not expose more than one visible child window during cycle {}; saw {}.",
                                  cycle,
                                  workerResult.snapshot.visibleChildWindowCount));
        state.Require(workerResult.snapshot.title == caption, std::format(L"Selection-mask prompt caption mismatch during cycle {}.", cycle));
        state.Require(workerResult.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect UI Automation stats for selection-mask prompt during cycle {}.", cycle));
        if (workerResult.uiaPatternStats.has_value())
        {
            state.Require(workerResult.uiaPatternStats->visibleElementCount >= 3u,
                          std::format(L"Selection-mask prompt should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats->editControlCount > 0u,
                          std::format(L"Selection-mask prompt should expose a visible editable field during cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Selection-mask prompt should expose ValuePattern during cycle {}.", cycle));
            state.Require(workerResult.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Selection-mask prompt should expose visible command buttons during cycle {}.", cycle));
        }
        state.Require(workerResult.valueState.has_value(),
                      std::format(L"Failed to collect selection-mask prompt visible DX edit ValuePattern state during cycle {}.", cycle));
        if (workerResult.valueState.has_value())
        {
            state.Require(! workerResult.valueState->isReadOnly,
                          std::format(L"Selection-mask prompt visible DX edit surface should remain editable during cycle {}.", cycle));
            state.Require(! workerResult.valueState->name.empty(),
                          std::format(L"Selection-mask prompt visible DX edit surface should expose a stable accessible name during cycle {}.", cycle));
        }
        state.Require(! workerResult.buttonName.empty(),
                      std::format(L"Selection-mask prompt visible DX command button should expose a stable accessible name during cycle {}.", cycle));
        state.Require(workerResult.setText, std::format(L"Failed to set the selection mask text during cycle {}.", cycle));
        state.Require(workerResult.actionIssued, std::format(L"Failed to close the selection-mask prompt through the DX action path during cycle {}.", cycle));

        if (accept)
        {
            if (useSelectCommand)
            {
                state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                              std::format(L"a.txt should be selected after select-mask accept cycle {}.", cycle));
                state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                              std::format(L"b.log should remain selected after select-mask accept cycle {}.", cycle));
                state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                              std::format(L"c.txt should be selected after select-mask accept cycle {}.", cycle));
                state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u,
                              std::format(L"Expected three selected items after select-mask accept cycle {}.", cycle));
            }
            else
            {
                state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"),
                              std::format(L"a.txt should be unselected after unselect-mask accept cycle {}.", cycle));
                state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"),
                              std::format(L"b.log should remain selected after unselect-mask accept cycle {}.", cycle));
                state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"),
                              std::format(L"c.txt should be unselected after unselect-mask accept cycle {}.", cycle));
                state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u,
                              std::format(L"Expected one selected item after unselect-mask accept cycle {}.", cycle));
            }
        }
        else
        {
            state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt") == hadATxtSelected,
                          std::format(L"a.txt selection should remain unchanged after selection-mask cancel cycle {}.", cycle));
            state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log") == hadBLogSelected,
                          std::format(L"b.log selection should remain unchanged after selection-mask cancel cycle {}.", cycle));
            state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt") == hadCTxtSelected,
                          std::format(L"c.txt selection should remain unchanged after selection-mask cancel cycle {}.", cycle));
            state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == selectedBefore,
                          std::format(L"Selection count should remain unchanged after selection-mask cancel cycle {}.", cycle));
        }

        const HWND lingeringPrompt = GetFolderViewSelectionMaskPromptHandle();
        state.Require(lingeringPrompt == nullptr || IsWindow(lingeringPrompt) == FALSE,
                      std::format(L"Selection-mask prompt should not remain open after cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closePrompt();
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneFilterDialogFiltering(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    Trace(L"pane_filter_dialog_filtering: begin");

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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane filter test root.");

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

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Failed to set pane path for filter test.");
    state.Require(WaitForAtomicAtLeast(enumerationCount, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for filter test.");

    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt in folder view before filtering.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log in folder view before filtering.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt in folder view before filtering.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        Trace(L"pane_filter_dialog_filtering: enabling filter command");
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();
        Trace(L"pane_filter_dialog_filtering: enabling filter dialog closed");

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after OK.");
    }

    Trace(L"pane_filter_dialog_filtering: waiting enable enumeration");
    state.Require(WaitForAtomicAtLeast(enumerationCount, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh after applying filter.");

    Trace(L"pane_filter_dialog_filtering: checking enable state");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter state expected to be active after applying a mask.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt after filtering.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt after filtering.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log to be filtered out.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, false, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        Trace(L"pane_filter_dialog_filtering: disabling filter command");
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();
        Trace(L"pane_filter_dialog_filtering: disabling filter dialog closed");

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not reopen.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close after disabling filter.");
    }

    Trace(L"pane_filter_dialog_filtering: waiting disable enumeration");
    state.Require(WaitForAtomicAtLeast(enumerationCount, 3u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not refresh after disabling filter.");

    Trace(L"pane_filter_dialog_filtering: checking disable state");
    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter state expected to be inactive when disabled.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log to return after disabling filter.");

    Trace(L"pane_filter_dialog_filtering: end");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneFilterHistoryRestore(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_filter_history_" + NewGuidText());
    const std::filesystem::path a    = root / L"A";
    const std::filesystem::path b    = root / L"B";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(a), L"Failed to create A folder.");
    state.Require(SelfTest::EnsureDirectory(b), L"Failed to create B folder.");

    state.Require(SelfTest::WriteTextFile(a / L"keep.txt", "a"), L"Failed to create keep.txt.");
    state.Require(SelfTest::WriteTextFile(a / L"drop.log", "b"), L"Failed to create drop.log.");
    state.Require(SelfTest::WriteTextFile(b / L"other.txt", "c"), L"Failed to create other.txt.");

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);

    std::atomic<uint32_t> enumA{0};
    std::atomic<uint32_t> enumB{0};
    g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left,
                                                       [&](const std::filesystem::path& folder) noexcept
    {
        if (OrdinalString::EqualsNoCasePath(folder, a))
        {
            enumA.fetch_add(1u, std::memory_order_release);
        }
        else if (OrdinalString::EqualsNoCasePath(folder, b))
        {
            enumB.fetch_add(1u, std::memory_order_release);
        }
    });
    const auto clearEnumCallback = wil::scope_exit([&] { g_folderWindow.SetPaneEnumerationCompletedCallback(FolderWindow::Pane::Left, {}); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, a);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, a, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Failed to set pane path to A.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"keep.txt", L"drop.log"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for A.");

    {
        PaneFilterDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomatePaneFilterDialog(mainWindow, dlg, true, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_LEFT_FILTER, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Pane Filter dialog did not open for A.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Pane Filter dialog did not close for A.");
    }

    state.Require(WaitForAtomicAtLeast(enumA, 2u, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Enumeration did not refresh for A after filtering.");
    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter expected active on A.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"drop.log"), L"A/drop.log expected to be filtered out.");

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, b);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, b, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Failed to set pane path to B.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"other.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for B.");

    state.Require(! g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter expected inactive on B (no saved filter state).");

    g_folderWindow.CommandHistoryBack(FolderWindow::Pane::Left);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, a, SelfTest::Scale(std::chrono::milliseconds{3000})), L"HistoryBack did not navigate to A.");
    state.Require(WaitForAtomicAtLeast(enumA, 3u, SelfTest::Scale(std::chrono::milliseconds{3000})), L"Enumeration did not complete after HistoryBack to A.");

    state.Require(g_folderWindow.DebugIsNameFilterActive(FolderWindow::Pane::Left), L"Filter expected restored when navigating back to A from history.");
    state.Require(! g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"drop.log"), L"A/drop.log expected to remain filtered after restore.");

    return state.failure.empty();
}

[[nodiscard]] bool TestSelectionMaskDialogs(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"selection_mask_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create selection-mask test root.");

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
                  L"Failed to set local file-system plugin for selection-mask test.");

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
                  L"Failed to set pane path for selection-mask test.");
    state.Require(WaitForAtomicAtLeast(enumCount, 1u, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Enumeration did not complete for selection-mask test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"a.txt", L"b.log", L"c.txt"}, SelfTest::Scale(std::chrono::milliseconds{3000})),
                  L"Pane contents not ready for selection-mask test.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"b.log"; }, true);
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log selected before Select dialog.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected item before Select dialog.");

    const std::wstring selectCaption = LoadStringResource(nullptr, IDS_CAPTION_SELECTION_MASK_SELECT);
    {
        SelectionMaskDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomateSelectionMaskDialog(mainWindow, dlg, selectCaption, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_SELECT_DIALOG, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Select dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Select dialog did not close after OK.");
    }

    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt selected after Select dialog.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt selected after Select dialog.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log to remain selected after Select dialog.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u, L"Expected three selected items after Select dialog.");

    const std::wstring unselectCaption = LoadStringResource(nullptr, IDS_CAPTION_SELECTION_MASK_UNSELECT);
    {
        SelectionMaskDialogAutomationState dlg{};
        std::jthread okCloser([&](std::stop_token) noexcept { AutomateSelectionMaskDialog(mainWindow, dlg, unselectCaption, L"*.txt", true); });
        FocusFolderViewPane(FolderWindow::Pane::Left);
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_PANE_SELECTION_UNSELECT_DIALOG, 0), 0);
        okCloser.join();

        state.Require(dlg.sawDialog.load(std::memory_order_acquire), L"Unselect dialog did not open.");
        state.Require(dlg.closed.load(std::memory_order_acquire), L"Unselect dialog did not close after OK.");
    }

    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"a.txt"), L"Expected a.txt unselected after Unselect dialog.");
    state.Require(! g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"c.txt"), L"Expected c.txt unselected after Unselect dialog.");
    state.Require(g_folderWindow.DebugIsItemSelected(FolderWindow::Pane::Left, L"b.log"), L"Expected b.log to remain selected after Unselect dialog.");
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected one selected item after Unselect dialog.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRenamePromptUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_rename_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-rename test directory.");

    const std::filesystem::path original = root / L"rename_me.txt";
    const std::filesystem::path renamed  = root / L"renamed_file.txt";
    state.Require(SelfTest::WriteTextFile(original, "rename"), L"Failed to create pane-rename test file.");
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

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for rename test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for rename test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"rename_me.txt"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for rename test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"rename_me.txt"; }, true);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"rename_me.txt", L"Rename test expected focus on rename_me.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closePrompt = [&] noexcept
    {
        if (const HWND prompt = GetFolderViewRenamePromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&] noexcept { closePrompt(); });
    closePrompt();

    const auto resolveExpectedSelectionEnd = [](std::wstring_view text) noexcept -> size_t
    {
        const size_t dot = text.find_last_of(L'.');
        if (dot == std::wstring_view::npos || dot == 0u)
        {
            return text.size();
        }

        return dot;
    };

    struct RenameSurfaceCycleResult final
    {
        HWND prompt            = nullptr;
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool capturedSnapshot  = false;
        FolderViewRenamePromptDebugSnapshot snapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
        std::optional<UiaValuePatternState> valueState;
        std::optional<UiaNamedElementState> buttonState;
        bool setText                = false;
        bool capturedEditedSnapshot = false;
        FolderViewRenamePromptDebugSnapshot editedSnapshot{};
        std::optional<UiaValuePatternState> editedValueState;
        bool actionTriggered = false;
        bool closed          = false;
    };

    const auto requireDxSurface = [&](const RenameSurfaceCycleResult& result, std::wstring_view expectedText, std::wstring_view failureContext) noexcept
    {
        state.Require(result.sawPrompt, std::format(L"Rename prompt did not open for {}.", failureContext));
        if (! result.sawPrompt)
        {
            return;
        }

        state.Require(result.ownedByMainWindow, std::format(L"Rename prompt should be owned by the main window during {}.", failureContext));
        state.Require(result.capturedSnapshot, std::format(L"Failed to capture rename prompt snapshot during {}.", failureContext));
        state.Require(result.snapshot.usesDxUiHost, std::format(L"Rename prompt should use the shared DxUi host during {}.", failureContext));
        state.Require(result.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Rename prompt should not expose more than one visible child window during {}; saw {}.",
                                  failureContext,
                                  result.snapshot.visibleChildWindowCount));
        state.Require(result.snapshot.text == expectedText,
                      std::format(L"Rename prompt should expose '{}' through the live snapshot during {}.", expectedText, failureContext));
        state.Require(result.snapshot.selectionStart == 0u,
                      std::format(L"Rename prompt should select from the beginning of the name during {}.", failureContext));
        state.Require(result.snapshot.selectionEnd == resolveExpectedSelectionEnd(expectedText),
                      std::format(L"Rename prompt should select only the basename during {}; saw selection end {} for '{}'.",
                                  failureContext,
                                  result.snapshot.selectionEnd,
                                  expectedText));

        state.Require(result.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation stats for rename prompt during {}.", failureContext));
        if (result.uiaPatternStats.has_value())
        {
            state.Require(result.uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Rename prompt should expose visible UI Automation descendants during {}.", failureContext));
            state.Require(result.uiaPatternStats->editControlCount > 0u,
                          std::format(L"Rename prompt should expose a visible UI Automation edit descendant during {}.", failureContext));
            state.Require(result.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Rename prompt should expose ValuePattern for the rename field during {}.", failureContext));
            state.Require(result.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Rename prompt should expose a visible UI Automation command button during {}.", failureContext));
            state.Require(result.uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Rename prompt should expose InvokePattern for its visible DX command button during {}.", failureContext));
            state.Require(result.uiaPatternStats->visibleElementCount >= 3u,
                          std::format(L"Rename prompt should expose the edit field and visible UI Automation command controls during {}.", failureContext));
        }

        state.Require(result.valueState.has_value(),
                      std::format(L"Failed to collect UI Automation ValuePattern state for rename prompt during {}.", failureContext));
        if (result.valueState.has_value())
        {
            state.Require(! result.valueState->isReadOnly, std::format(L"Rename prompt field should remain editable during {}.", failureContext));
            state.Require(! result.valueState->name.empty(),
                          std::format(L"Rename prompt editable DX field should expose a stable accessible name during {}.", failureContext));
            state.Require(result.valueState->value == result.snapshot.text,
                          std::format(L"Rename prompt field should expose the live prompt text through ValuePattern during {}; saw '{}'.",
                                      failureContext,
                                      result.valueState->value));
        }

        state.Require(result.buttonState.has_value(), std::format(L"Rename prompt should expose a visible DX command button during {}.", failureContext));
        if (result.buttonState.has_value())
        {
            state.Require(! result.buttonState->name.empty(),
                          std::format(L"Rename prompt visible DX command button should expose a stable accessible name during {}.", failureContext));
        }
    };

    RenameSurfaceCycleResult confirmCycle{};
    RunRenamePromptModalCycle(mainWindow,
                              [&](const HWND prompt) noexcept
    {
        confirmCycle.prompt    = prompt;
        confirmCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! confirmCycle.sawPrompt)
        {
            return;
        }

        confirmCycle.ownedByMainWindow      = IsOwnedBy(prompt, mainWindow);
        confirmCycle.capturedSnapshot       = DebugGetFolderViewRenamePromptSnapshot(confirmCycle.snapshot);
        confirmCycle.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
        confirmCycle.valueState             = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        confirmCycle.buttonState            = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
        confirmCycle.setText                = DebugSetFolderViewRenamePromptText(L"  renamed_file.txt  ");
        confirmCycle.capturedEditedSnapshot = DebugGetFolderViewRenamePromptSnapshot(confirmCycle.editedSnapshot);
        confirmCycle.editedValueState       = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        confirmCycle.actionTriggered        = DebugConfirmFolderViewRenamePrompt();
        confirmCycle.closed                 = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    requireDxSurface(confirmCycle, L"rename_me.txt", L"initial confirm pass");
    state.Require(confirmCycle.setText, L"Failed to set rename prompt text during initial confirm pass.");
    state.Require(confirmCycle.capturedEditedSnapshot, L"Failed to capture edited rename prompt snapshot during initial confirm pass.");
    state.Require(confirmCycle.editedValueState.has_value(),
                  L"Failed to collect edited UI Automation ValuePattern state for rename prompt during initial confirm pass.");
    if (confirmCycle.editedValueState.has_value())
    {
        state.Require(confirmCycle.editedValueState->value == confirmCycle.editedSnapshot.text,
                      std::format(L"Rename prompt ValuePattern should track the edited text during the initial confirm pass; saw '{}'.",
                                  confirmCycle.editedValueState->value));
    }
    state.Require(confirmCycle.actionTriggered, L"Failed to confirm rename prompt during initial confirm pass.");
    state.Require(confirmCycle.closed, L"Rename prompt did not close after confirm during initial confirm pass.");

    const auto renameDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < renameDeadline)
    {
        PumpPendingMessages();
        ec.clear();
        if (std::filesystem::exists(renamed, ec))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(std::filesystem::exists(renamed, ec), L"Rename command did not rename the file.");
    state.Require(! std::filesystem::exists(original, ec), L"Original file still exists after rename.");
    const auto waitForFocusedItem = [&](std::wstring_view expectedName) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedName)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        return g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == expectedName;
    };
    state.Require(waitForFocusedItem(L"renamed_file.txt"), L"Rename test expected focus on renamed_file.txt after rename.");

    RenameSurfaceCycleResult cancelCycle{};
    RunRenamePromptModalCycle(mainWindow,
                              [&](const HWND prompt) noexcept
    {
        cancelCycle.prompt    = prompt;
        cancelCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! cancelCycle.sawPrompt)
        {
            return;
        }

        cancelCycle.ownedByMainWindow      = IsOwnedBy(prompt, mainWindow);
        cancelCycle.capturedSnapshot       = DebugGetFolderViewRenamePromptSnapshot(cancelCycle.snapshot);
        cancelCycle.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
        cancelCycle.valueState             = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        cancelCycle.buttonState            = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
        cancelCycle.setText                = DebugSetFolderViewRenamePromptText(L"should_not_apply.txt");
        cancelCycle.capturedEditedSnapshot = DebugGetFolderViewRenamePromptSnapshot(cancelCycle.editedSnapshot);
        cancelCycle.editedValueState       = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        cancelCycle.actionTriggered        = DebugCancelFolderViewRenamePrompt();
        cancelCycle.closed                 = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    requireDxSurface(cancelCycle, L"renamed_file.txt", L"reopened cancel pass");
    state.Require(cancelCycle.setText, L"Failed to set rename prompt cancel text during reopened cancel pass.");
    state.Require(cancelCycle.capturedEditedSnapshot, L"Failed to capture edited rename prompt snapshot during reopened cancel pass.");
    state.Require(cancelCycle.editedValueState.has_value(),
                  L"Failed to collect edited UI Automation ValuePattern state for rename prompt during reopened cancel pass.");
    if (cancelCycle.editedValueState.has_value())
    {
        state.Require(cancelCycle.editedValueState->value == cancelCycle.editedSnapshot.text,
                      std::format(L"Rename prompt ValuePattern should track the edited text during the reopened cancel pass; saw '{}'.",
                                  cancelCycle.editedValueState->value));
    }
    state.Require(cancelCycle.actionTriggered, L"Failed to cancel rename prompt during reopened cancel pass.");
    state.Require(cancelCycle.closed, L"Rename prompt did not close after cancel during reopened cancel pass.");

    state.Require(std::filesystem::exists(renamed, ec), L"Cancel should keep the renamed file.");
    state.Require(! std::filesystem::exists(root / L"should_not_apply.txt", ec), L"Cancel should not rename the file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRenamePromptLongInitialSelectionStaysClipped(HWND mainWindow, CaseState& state) noexcept
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

    const std::wstring longName      = L"stepfather-s-day-b93327ff-9c23-45e5-a777-eab3c75f474d-1080.txt";
    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_rename_long_selection_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-rename long-selection test directory.");
    state.Require(SelfTest::WriteTextFile(root / longName, "rename long selection"), L"Failed to create pane-rename long-selection file.");
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

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for rename long-selection test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for rename long-selection test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {longName}, SelfTest::Scale(3000ms)), L"Pane contents not ready for rename long-selection test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [&](std::wstring_view name) noexcept { return name == longName; }, true);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == longName,
                  L"Rename long-selection test expected focus on the long file name.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closePrompt = [&] noexcept
    {
        if (const HWND prompt = GetFolderViewRenamePromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&] noexcept { closePrompt(); });
    closePrompt();

    struct RenameLongSelectionCycle final
    {
        HWND prompt           = nullptr;
        bool sawPrompt        = false;
        bool capturedSnapshot = false;
        FolderViewRenamePromptDebugSnapshot snapshot{};
        bool actionTriggered = false;
        bool closed          = false;
    };

    RenameLongSelectionCycle cycle{};
    RunRenamePromptModalCycle(mainWindow,
                              [&](const HWND prompt) noexcept
    {
        cycle.prompt    = prompt;
        cycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! cycle.sawPrompt)
        {
            return;
        }

        cycle.capturedSnapshot = DebugGetFolderViewRenamePromptSnapshot(cycle.snapshot);
        cycle.actionTriggered  = DebugCancelFolderViewRenamePrompt();
        cycle.closed           = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(cycle.sawPrompt, L"Rename prompt did not open for long initial selection.");
    state.Require(cycle.capturedSnapshot, L"Failed to capture rename prompt long initial selection snapshot.");
    if (cycle.capturedSnapshot)
    {
        const size_t dot = longName.find_last_of(L'.');
        state.Require(cycle.snapshot.text == longName, L"Rename prompt long initial selection should expose the selected file name.");
        state.Require(cycle.snapshot.selectionStart == 0u, L"Rename prompt long initial selection should start at the beginning of the file name.");
        state.Require(cycle.snapshot.selectionEnd == dot, L"Rename prompt long initial selection should select only the basename.");
        state.Require(cycle.snapshot.hasSelectionPaintRect, L"Rename prompt long initial selection should expose a selection paint rect.");
        state.Require(cycle.snapshot.horizontalScrollDip > 0.0f,
                      L"Rename prompt long initial selection should horizontally scroll to the selected basename end.");

        constexpr float kClipEpsilonDip = 0.5f;
        state.Require(cycle.snapshot.selectionPaintRect.left >= cycle.snapshot.textRect.left - kClipEpsilonDip,
                      std::format(L"Rename prompt long initial selection should clip left; selection left={} text left={}.",
                                  cycle.snapshot.selectionPaintRect.left,
                                  cycle.snapshot.textRect.left));
        state.Require(cycle.snapshot.selectionPaintRect.right <= cycle.snapshot.textRect.right + kClipEpsilonDip,
                      std::format(L"Rename prompt long initial selection should clip right; selection right={} text right={}.",
                                  cycle.snapshot.selectionPaintRect.right,
                                  cycle.snapshot.textRect.right));
    }
    state.Require(cycle.actionTriggered, L"Failed to cancel rename prompt after long initial selection snapshot.");
    state.Require(cycle.closed, L"Rename prompt did not close after long initial selection snapshot.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRenamePromptLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_rename_live_dx_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-rename live interaction test directory.");

    const std::filesystem::path original = root / L"rename_me.txt";
    const std::filesystem::path renamed  = root / L"renamed_via_live_dx.txt";
    state.Require(SelfTest::WriteTextFile(original, "rename"), L"Failed to create pane-rename live interaction test file.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewRenamePromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for rename live interaction test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for rename live interaction test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"rename_me.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for rename live interaction test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"rename_me.txt"; }, true);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"rename_me.txt",
                  L"Rename live interaction test expected focus on rename_me.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct RenameLiveCycleResult final
    {
        HWND prompt            = nullptr;
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool capturedSnapshot  = false;
        FolderViewRenamePromptDebugSnapshot snapshot{};
        std::optional<UiaValuePatternState> initialValueState;
        bool setValue = false;
        std::optional<UiaValuePatternState> mutatedValueState;
        bool capturedEditedSnapshot = false;
        FolderViewRenamePromptDebugSnapshot editedSnapshot{};
        bool invokeAction = false;
        bool closed       = false;
    };

    const auto requireLiveCycle =
        [&](const RenameLiveCycleResult& result, std::wstring_view expectedStart, std::wstring_view expectedEdited, std::wstring_view context) noexcept
    {
        state.Require(result.sawPrompt, std::format(L"Rename prompt did not open for {}.", context));
        if (! result.sawPrompt)
        {
            return;
        }

        state.Require(result.ownedByMainWindow, std::format(L"Rename prompt should be owned by the main window during {}.", context));
        state.Require(result.capturedSnapshot, std::format(L"Failed to capture rename prompt snapshot during {}.", context));
        state.Require(result.snapshot.usesDxUiHost, std::format(L"Rename prompt should use the shared DxUi host during {}.", context));
        state.Require(result.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Rename prompt should not expose more than one visible child window during {}; saw {}.",
                                  context,
                                  result.snapshot.visibleChildWindowCount));
        state.Require(result.snapshot.text == expectedStart, std::format(L"Rename prompt text should start as '{}' during {}.", expectedStart, context));
        state.Require(result.initialValueState.has_value(), std::format(L"Rename prompt should expose a visible editable DX field during {}.", context));
        if (result.initialValueState.has_value())
        {
            state.Require(! result.initialValueState->name.empty(),
                          std::format(L"Rename prompt DX edit should expose a stable accessible name during {}.", context));
            state.Require(! result.initialValueState->isReadOnly, std::format(L"Rename prompt DX edit should remain editable during {}.", context));
            state.Require(result.initialValueState->value == expectedStart,
                          std::format(L"Rename prompt ValuePattern should start as '{}' during {}.", expectedStart, context));
        }
        state.Require(result.setValue, std::format(L"Failed to set rename prompt DX edit during {}.", context));
        state.Require(result.mutatedValueState.has_value(),
                      std::format(L"Rename prompt DX edit did not update after live UIA interaction during {}.", context));
        if (result.mutatedValueState.has_value())
        {
            state.Require(result.mutatedValueState->value == expectedEdited,
                          std::format(L"Rename prompt DX edit should update to '{}' during {}.", expectedEdited, context));
        }
        state.Require(result.capturedEditedSnapshot, std::format(L"Failed to recapture rename prompt snapshot during {}.", context));
        state.Require(result.invokeAction, std::format(L"Rename prompt did not close through the debug action path during {}.", context));
        state.Require(result.closed, std::format(L"Rename prompt did not close after {}.", context));
    };

    const auto renameDeadline = [&](const std::filesystem::path& expectedPath) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            ec.clear();
            if (std::filesystem::exists(expectedPath, ec))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        ec.clear();
        return std::filesystem::exists(expectedPath, ec);
    };

    RenameLiveCycleResult cancelCycle{};
    RunRenamePromptModalCycle(mainWindow,
                              [&](const HWND prompt) noexcept
    {
        cancelCycle.prompt    = prompt;
        cancelCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! cancelCycle.sawPrompt)
        {
            return;
        }

        cancelCycle.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        cancelCycle.capturedSnapshot  = DebugGetFolderViewRenamePromptSnapshot(cancelCycle.snapshot);
        cancelCycle.initialValueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        if (cancelCycle.initialValueState.has_value() && ! cancelCycle.initialValueState->name.empty())
        {
            cancelCycle.setValue = DebugSetFolderViewRenamePromptText(L"should_not_apply_live.txt");
            if (cancelCycle.setValue)
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    cancelCycle.mutatedValueState =
                        CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, cancelCycle.initialValueState->name);
                    cancelCycle.capturedEditedSnapshot = DebugGetFolderViewRenamePromptSnapshot(cancelCycle.editedSnapshot);
                    if (cancelCycle.mutatedValueState.has_value() && cancelCycle.mutatedValueState->value == L"should_not_apply_live.txt" &&
                        cancelCycle.capturedEditedSnapshot && cancelCycle.editedSnapshot.text == L"should_not_apply_live.txt")
                    {
                        break;
                    }
                    std::this_thread::sleep_for(20ms);
                }
            }
        }
        if (! cancelCycle.capturedEditedSnapshot)
        {
            cancelCycle.capturedEditedSnapshot = DebugGetFolderViewRenamePromptSnapshot(cancelCycle.editedSnapshot);
        }
        cancelCycle.invokeAction = DebugCancelFolderViewRenamePrompt();
        cancelCycle.closed       = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });
    requireLiveCycle(cancelCycle, L"rename_me.txt", L"should_not_apply_live.txt", L"live cancel interaction");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(std::filesystem::exists(original, ec), L"Cancel should keep the original file before confirm.");
    state.Require(! std::filesystem::exists(root / L"should_not_apply_live.txt", ec), L"Cancel should not rename the file after live interaction.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"rename_me.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane items did not restore the original file after live DX cancel.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"rename_me.txt"; }, true);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"rename_me.txt",
                  L"Rename live interaction test expected focus on rename_me.txt after cancel.");
    if (! state.failure.empty())
    {
        return false;
    }

    RenameLiveCycleResult confirmCycle{};
    RunRenamePromptModalCycle(mainWindow,
                              [&](const HWND prompt) noexcept
    {
        confirmCycle.prompt    = prompt;
        confirmCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! confirmCycle.sawPrompt)
        {
            return;
        }

        confirmCycle.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        confirmCycle.capturedSnapshot  = DebugGetFolderViewRenamePromptSnapshot(confirmCycle.snapshot);
        confirmCycle.initialValueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        if (confirmCycle.initialValueState.has_value() && ! confirmCycle.initialValueState->name.empty())
        {
            confirmCycle.setValue = DebugSetFolderViewRenamePromptText(L"renamed_via_live_dx.txt");
            if (confirmCycle.setValue)
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    confirmCycle.mutatedValueState =
                        CollectVisibleDescendantValuePatternStateByName(prompt, UIA_EditControlTypeId, confirmCycle.initialValueState->name);
                    confirmCycle.capturedEditedSnapshot = DebugGetFolderViewRenamePromptSnapshot(confirmCycle.editedSnapshot);
                    if (confirmCycle.mutatedValueState.has_value() && confirmCycle.mutatedValueState->value == L"renamed_via_live_dx.txt" &&
                        confirmCycle.capturedEditedSnapshot && confirmCycle.editedSnapshot.text == L"renamed_via_live_dx.txt")
                    {
                        break;
                    }
                    std::this_thread::sleep_for(20ms);
                }
            }
        }
        if (! confirmCycle.capturedEditedSnapshot)
        {
            confirmCycle.capturedEditedSnapshot = DebugGetFolderViewRenamePromptSnapshot(confirmCycle.editedSnapshot);
        }
        confirmCycle.invokeAction = DebugConfirmFolderViewRenamePrompt();
        confirmCycle.closed       = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });
    requireLiveCycle(confirmCycle, L"rename_me.txt", L"renamed_via_live_dx.txt", L"live confirm interaction");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(renameDeadline(renamed), L"Rename live interaction did not rename the file.");
    state.Require(! std::filesystem::exists(original, ec), L"Original file still exists after rename live interaction.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"renamed_via_live_dx.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane items did not settle on the renamed file after live interaction.");

    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
        FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"renamed_via_live_dx.txt"; }, true);
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"renamed_via_live_dx.txt",
                  L"Rename live interaction test expected focus on renamed_via_live_dx.txt after confirm.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneRenamePromptLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"pane_rename_prompt_churn_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create pane-rename prompt churn test directory.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::wstring currentName = L"rename_me.txt";
    state.Require(SelfTest::WriteTextFile(root / currentName, "rename"), L"Failed to create initial pane-rename churn file.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewRenamePromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.DebugResetPaneVisibilityState(FolderWindow::Pane::Left);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for pane-rename prompt churn test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for pane-rename prompt churn test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for pane-rename prompt churn test.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        struct RenameChurnCycleResult final
        {
            HWND prompt            = nullptr;
            bool sawPrompt         = false;
            bool ownedByMainWindow = false;
            bool capturedSnapshot  = false;
            FolderViewRenamePromptDebugSnapshot snapshot{};
            std::optional<UiaDescendantPatternStats> uiaPatternStats;
            std::optional<UiaValuePatternState> valueState;
            std::optional<UiaNamedElementState> buttonState;
            bool setText         = false;
            bool actionTriggered = false;
            bool closed          = false;
        } cycleResult{};

        const bool accept                         = (cycle % 2u) == 0u;
        const std::wstring requestedName          = accept ? std::format(L"renamed_{:02}.txt", cycle) : std::format(L"should_not_apply_{:02}.txt", cycle);
        const std::filesystem::path beforePath    = root / currentName;
        const std::filesystem::path requestedPath = root / requestedName;

        state.Require(std::filesystem::exists(beforePath, ec), std::format(L"Expected '{}' to exist before rename cycle {}.", currentName, cycle));
        state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                      std::format(L"Pane contents not settled before rename cycle {}.", cycle));
        g_folderWindow.SetPaneSelectionByDisplayNamePredicate(
            FolderWindow::Pane::Left, [&](std::wstring_view name) noexcept { return name == currentName; }, true);
        state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == currentName,
                      std::format(L"Rename churn expected focus on '{}' before cycle {}.", currentName, cycle));
        if (! state.failure.empty())
        {
            closePrompt();
            return false;
        }

        RunRenamePromptModalCycle(mainWindow,
                                  [&](const HWND prompt) noexcept
        {
            cycleResult.prompt    = prompt;
            cycleResult.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
            if (! cycleResult.sawPrompt)
            {
                return;
            }

            cycleResult.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
            cycleResult.capturedSnapshot  = DebugGetFolderViewRenamePromptSnapshot(cycleResult.snapshot);
            cycleResult.uiaPatternStats   = CollectVisibleUiaDescendantPatternStats(prompt);
            cycleResult.valueState        = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
            cycleResult.buttonState       = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
            cycleResult.setText           = DebugSetFolderViewRenamePromptText(requestedName);
            cycleResult.actionTriggered   = accept ? DebugConfirmFolderViewRenamePrompt() : DebugCancelFolderViewRenamePrompt();
            cycleResult.closed            = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        });

        state.Require(cycleResult.sawPrompt, std::format(L"Rename prompt did not open during cycle {}.", cycle));
        if (! cycleResult.sawPrompt)
        {
            closePrompt();
            return false;
        }

        state.Require(cycleResult.ownedByMainWindow, std::format(L"Rename prompt should be owned by the main window during cycle {}.", cycle));
        state.Require(cycleResult.capturedSnapshot, std::format(L"Failed to capture rename prompt snapshot during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.usesDxUiHost, std::format(L"Rename prompt should use a DxUi host during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Rename prompt should not expose more than one visible child window during cycle {}; saw {}.",
                                  cycle,
                                  cycleResult.snapshot.visibleChildWindowCount));
        state.Require(cycleResult.snapshot.text == currentName, std::format(L"Rename prompt text should start with '{}' during cycle {}.", currentName, cycle));

        state.Require(cycleResult.uiaPatternStats.has_value(), std::format(L"Failed to collect UI Automation stats for rename prompt during cycle {}.", cycle));
        if (cycleResult.uiaPatternStats.has_value())
        {
            state.Require(cycleResult.uiaPatternStats->visibleElementCount >= 3u,
                          std::format(L"Rename prompt should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->editControlCount > 0u,
                          std::format(L"Rename prompt should expose a visible editable field during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Rename prompt should expose ValuePattern during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Rename prompt should expose visible command buttons during cycle {}.", cycle));
        }

        state.Require(cycleResult.valueState.has_value(),
                      std::format(L"Rename prompt should expose a visible editable ValuePattern field during cycle {}.", cycle));
        if (cycleResult.valueState.has_value())
        {
            state.Require(! cycleResult.valueState->isReadOnly, std::format(L"Rename prompt field should remain editable during cycle {}.", cycle));
            state.Require(! cycleResult.valueState->name.empty(),
                          std::format(L"Rename prompt visible DX edit surface should expose a stable accessible name during cycle {}.", cycle));
            state.Require(cycleResult.valueState->value == currentName,
                          std::format(L"Rename prompt ValuePattern should start with '{}' during cycle {}.", currentName, cycle));
        }

        state.Require(cycleResult.buttonState.has_value(), std::format(L"Rename prompt should expose a visible DX command button during cycle {}.", cycle));
        if (cycleResult.buttonState.has_value())
        {
            state.Require(! cycleResult.buttonState->name.empty(),
                          std::format(L"Rename prompt visible DX command button should expose a stable accessible name during cycle {}.", cycle));
        }

        state.Require(cycleResult.setText, std::format(L"Failed to set rename prompt text during cycle {}.", cycle));
        state.Require(cycleResult.actionTriggered, std::format(L"Failed to close the rename prompt through the DX action path during cycle {}.", cycle));
        state.Require(cycleResult.closed, std::format(L"Rename prompt did not close after cycle {}.", cycle));

        if (accept)
        {
            const auto renameDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
            while (std::chrono::steady_clock::now() < renameDeadline)
            {
                PumpPendingMessages();
                ec.clear();
                if (std::filesystem::exists(requestedPath, ec))
                {
                    break;
                }
                std::this_thread::sleep_for(20ms);
            }

            state.Require(std::filesystem::exists(requestedPath, ec), std::format(L"Rename cycle {} did not create '{}'.", cycle, requestedName));
            state.Require(! std::filesystem::exists(beforePath, ec), std::format(L"Rename cycle {} should remove '{}'.", cycle, currentName));
            currentName = requestedName;
            state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                          std::format(L"Pane items did not settle on '{}' after rename cycle {}.", currentName, cycle));
        }
        else
        {
            state.Require(std::filesystem::exists(beforePath, ec), std::format(L"Rename cancel cycle {} should keep '{}'.", cycle, currentName));
            state.Require(! std::filesystem::exists(requestedPath, ec), std::format(L"Rename cancel cycle {} should not create '{}'.", cycle, requestedName));
            state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {currentName}, SelfTest::Scale(3000ms)),
                          std::format(L"Pane items should remain on '{}' after rename cancel cycle {}.", currentName, cycle));
        }

        const HWND lingeringPrompt = GetFolderViewRenamePromptHandle();
        state.Require(lingeringPrompt == nullptr || IsWindow(lingeringPrompt) == FALSE,
                      std::format(L"Rename prompt should not remain open after cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closePrompt();
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestEditNewPromptFiltersEditorComboAndCreatesFile(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"edit_new_prompt_" + NewGuidText());
    const std::filesystem::path newFile    = root / L"alpha.editnew";
    const std::filesystem::path markerPath = root / L"edit-new-marker.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Edit New prompt test folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};

    Common::Settings::FileActionDefinition editNewEditor{};
    editNewEditor.id                = L"edit-new-editor";
    editNewEditor.displayName       = L"Edit New Marker Editor";
    editNewEditor.enabled           = true;
    editNewEditor.kind              = Common::Settings::FileActionKind::ExternalProgram;
    editNewEditor.executablePath    = ResolveCommandProcessorPath();
    editNewEditor.arguments         = L"/C echo {Filename}>\"{Path}\\edit-new-marker.txt\"";
    editNewEditor.workingDirectory  = L"{Path}";
    editNewEditor.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Extension, L".editnew"}};
    g_settings.fileActions.editors.actions.push_back(std::move(editNewEditor));

    Common::Settings::FileActionDefinition filteredEditor{};
    filteredEditor.id                = L"filtered-editor";
    filteredEditor.displayName       = L"Filtered Marker Editor";
    filteredEditor.enabled           = true;
    filteredEditor.kind              = Common::Settings::FileActionKind::ExternalProgram;
    filteredEditor.executablePath    = ResolveCommandProcessorPath();
    filteredEditor.arguments         = L"/C echo filtered>\"{Path}\\filtered-edit-new-marker.txt\"";
    filteredEditor.workingDirectory  = L"{Path}";
    filteredEditor.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Extension, L".othereditnew"}};
    g_settings.fileActions.editors.actions.push_back(std::move(filteredEditor));

    Common::Settings::FileActionDefinition unavailableEditor{};
    unavailableEditor.id                = L"missing-executable-editor";
    unavailableEditor.displayName       = L"Missing Executable Editor";
    unavailableEditor.enabled           = true;
    unavailableEditor.kind              = Common::Settings::FileActionKind::ExternalProgram;
    unavailableEditor.executablePath    = L"";
    unavailableEditor.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Extension, L".editnew"}};
    g_settings.fileActions.editors.actions.push_back(std::move(unavailableEditor));

    Common::Settings::FileActionDefinition computerFilteredEditor{};
    computerFilteredEditor.id                      = L"other-computer-editor";
    computerFilteredEditor.displayName             = L"Other Computer Editor";
    computerFilteredEditor.enabled                 = true;
    computerFilteredEditor.kind                    = Common::Settings::FileActionKind::ExternalProgram;
    computerFilteredEditor.executablePath          = ResolveCommandProcessorPath();
    computerFilteredEditor.arguments               = L"/C echo other>\"{Path}\\other-computer-edit-new-marker.txt\"";
    computerFilteredEditor.workingDirectory        = L"{Path}";
    computerFilteredEditor.appliesTo.matches       = {{Common::Settings::FileActionMatchKind::Extension, L".editnew"}};
    computerFilteredEditor.appliesTo.computerNames = {L"red-salamander-selftest-other-computer"};
    g_settings.fileActions.editors.actions.push_back(std::move(computerFilteredEditor));

    Common::Settings::EditorAssociationRule editNewRule{};
    editNewRule.match.kind      = Common::Settings::FileActionMatchKind::Extension;
    editNewRule.match.value     = L".editnew";
    editNewRule.editNewActionId = L"edit-new-editor";
    g_settings.fileActions.editors.associations.push_back(std::move(editNewRule));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Edit New prompt test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Edit New prompt test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeResult final
    {
        bool sawPrompt              = false;
        bool ownedByMainWindow      = false;
        bool capturedInitial        = false;
        bool setName                = false;
        bool capturedAfterExtension = false;
        bool selectedEditor         = false;
        bool actionTriggered        = false;
        bool closed                 = false;
        FolderViewEditNewPromptDebugSnapshot initialSnapshot{};
        FolderViewEditNewPromptDebugSnapshot extensionSnapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
    } probe{};

    RunEditNewPromptModalCycle(mainWindow,
                               [&](const HWND prompt) noexcept
    {
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        probe.ownedByMainWindow      = IsOwnedBy(prompt, mainWindow);
        probe.capturedInitial        = DebugGetFolderViewEditNewPromptSnapshot(probe.initialSnapshot);
        probe.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
        probe.setName                = DebugSetFolderViewEditNewPromptText(L"alpha.editnew");
        probe.capturedAfterExtension = DebugGetFolderViewEditNewPromptSnapshot(probe.extensionSnapshot);
        probe.selectedEditor         = DebugSelectFolderViewEditNewPromptEditor(L"edit-new-editor");
        probe.actionTriggered        = DebugConfirmFolderViewEditNewPrompt();
        probe.closed                 = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(probe.sawPrompt, L"Edit New prompt did not open.");
    state.Require(probe.ownedByMainWindow, L"Edit New prompt should be owned by the main window.");
    state.Require(probe.capturedInitial, L"Failed to capture initial Edit New prompt snapshot.");
    state.Require(probe.initialSnapshot.usesDxUiHost, L"Edit New prompt should use a DxUi host.");
    state.Require(probe.initialSnapshot.visibleChildWindowCount <= 1u, L"Edit New prompt should not expose native child-control fallback.");
    state.Require(probe.initialSnapshot.nameFieldFocused, L"Edit New prompt should focus the file-name field.");
    state.Require(probe.uiaPatternStats.has_value(), L"Failed to collect UI Automation stats for Edit New prompt.");
    if (probe.uiaPatternStats.has_value())
    {
        state.Require(probe.uiaPatternStats->editControlCount > 0u, L"Edit New prompt should expose a file-name edit field.");
        state.Require(probe.uiaPatternStats->comboBoxControlCount > 0u, L"Edit New prompt should expose an Editor combo.");
        state.Require(probe.uiaPatternStats->buttonControlCount > 0u, L"Edit New prompt should expose command buttons.");
    }
    state.Require(probe.setName, L"Failed to set Edit New prompt file name.");
    state.Require(probe.capturedAfterExtension, L"Failed to capture Edit New prompt after typing an extension.");
    state.Require(probe.extensionSnapshot.fileNameText == L"alpha.editnew", L"Edit New prompt should keep the typed file name.");
    state.Require(probe.extensionSnapshot.editorComboEnabled, L"Edit New prompt Editor combo should be enabled when an editor applies.");
    state.Require(probe.extensionSnapshot.editorActionIds.size() == 1u, L"Edit New prompt should list only extension-matched editors.");
    if (probe.extensionSnapshot.editorActionIds.size() == 1u)
    {
        state.Require(probe.extensionSnapshot.editorActionIds.front() == L"edit-new-editor",
                      L"Edit New prompt should list the editor matching the typed extension.");
    }
    state.Require(probe.extensionSnapshot.selectedEditorActionId == L"edit-new-editor",
                  L"Edit New prompt should select the primary editor configured for the typed extension.");
    state.Require(probe.selectedEditor, L"Failed to select the Edit New prompt editor.");
    state.Require(probe.actionTriggered, L"Failed to confirm the Edit New prompt.");
    state.Require(probe.closed, L"Edit New prompt did not close after confirmation.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::string markerText;
    bool markerReady    = false;
    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        ec.clear();
        markerReady = false;
        markerText.clear();
        if (std::filesystem::exists(markerPath, ec))
        {
            std::ifstream input(markerPath);
            std::getline(input, markerText);
            markerReady = markerText == "\"alpha.editnew\"";
        }
        ec.clear();
        if (std::filesystem::exists(newFile, ec) && markerReady && g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"alpha.editnew"))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(std::filesystem::exists(newFile, ec), L"Edit New should create the requested file.");
    ec.clear();
    state.Require(std::filesystem::exists(markerPath, ec), L"Edit New should launch the selected external editor action.");
    ec.clear();
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"alpha.editnew"),
                  L"Edit New should refresh the pane so the new file is visible.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"alpha.editnew",
                  L"Edit New should focus the new file after refresh.");
    state.Require(markerReady,
                  std::format(L"Edit New selected editor should receive the quoted {{Filename}} macro; actual marker='{}'.",
                              std::wstring(markerText.begin(), markerText.end())));

    return state.failure.empty();
}

[[nodiscard]] bool TestEditNewPromptRejectsInvalidAndExistingNames(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root         = suiteRoot / L"work" / (L"edit_new_validation_" + NewGuidText());
    const std::filesystem::path existingFile = root / L"taken.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Edit New validation test folder.");
    state.Require(SelfTest::WriteTextFile(existingFile, "already here"), L"Failed to create existing Edit New validation file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Edit New validation test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Edit New validation test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct ValidationProbe final
    {
        bool sawPrompt = false;
        bool cancelled = false;
        std::vector<std::wstring> missingValidationForNames;
        std::vector<std::wstring> closedForNames;
        std::vector<std::wstring> unfocusedForNames;
    } probe{};

    constexpr std::array<std::wstring_view, 6> kInvalidNames{{
        L"",
        L".",
        L"bad\\name.txt",
        L"bad:name.txt",
        L"NUL.txt",
        L"taken.txt",
    }};

    RunEditNewPromptModalCycle(mainWindow,
                               [&](const HWND prompt) noexcept
    {
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        for (const std::wstring_view name : kInvalidNames)
        {
            const std::wstring nameText{name};
            if (! DebugSetFolderViewEditNewPromptText(nameText) || ! DebugConfirmFolderViewEditNewPrompt())
            {
                probe.missingValidationForNames.push_back(nameText);
                continue;
            }

            FolderViewEditNewPromptDebugSnapshot snapshot{};
            if (! DebugGetFolderViewEditNewPromptSnapshot(snapshot) || snapshot.validationText.empty())
            {
                probe.missingValidationForNames.push_back(nameText);
            }
            if (IsWindow(prompt) == FALSE)
            {
                probe.closedForNames.push_back(nameText);
                return;
            }
            if (! snapshot.nameFieldFocused)
            {
                probe.unfocusedForNames.push_back(nameText);
            }
        }

        probe.cancelled = DebugCancelFolderViewEditNewPrompt();
    });

    state.Require(probe.sawPrompt, L"Edit New validation prompt did not open.");
    state.Require(probe.missingValidationForNames.empty(), L"Edit New should show a validation message for every invalid file name.");
    state.Require(probe.closedForNames.empty(), L"Edit New should keep the dialog open after invalid file names.");
    state.Require(probe.unfocusedForNames.empty(), L"Edit New should refocus the file-name field after invalid file names.");
    state.Require(probe.cancelled, L"Failed to cancel Edit New validation prompt.");

    ec.clear();
    state.Require(! std::filesystem::exists(root / L"bad", ec), L"Edit New should not create a folder/file from a path separator name.");
    ec.clear();
    state.Require(! std::filesystem::exists(root / L"NUL.txt", ec), L"Edit New should not create a reserved device-name file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestEditNewPromptCreatesFileWithoutApplicableEditor(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root    = suiteRoot / L"work" / (L"edit_new_no_editor_" + NewGuidText());
    const std::filesystem::path newFile = root / L"plain.noeditor";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Edit New no-editor test folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                             = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftBefore           = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const Common::Settings::EditorFileActionsSettings editorsBefore = g_settings.fileActions.editors;
    const auto restoreState                                         = wil::scope_exit([&]
    {
        g_settings.fileActions.editors = editorsBefore;
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    g_settings.fileActions.editors = Common::Settings::EditorFileActionsSettings{};
    Common::Settings::FileActionDefinition unrelatedEditor{};
    unrelatedEditor.id                = L"unrelated-editor";
    unrelatedEditor.displayName       = L"Unrelated Editor";
    unrelatedEditor.enabled           = true;
    unrelatedEditor.kind              = Common::Settings::FileActionKind::ExternalProgram;
    unrelatedEditor.executablePath    = ResolveCommandProcessorPath();
    unrelatedEditor.arguments         = L"/C exit /B 0";
    unrelatedEditor.workingDirectory  = L"{Path}";
    unrelatedEditor.appliesTo.matches = {{Common::Settings::FileActionMatchKind::Extension, L".different"}};
    g_settings.fileActions.editors.actions.push_back(std::move(unrelatedEditor));

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for Edit New no-editor test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Edit New no-editor test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeResult final
    {
        bool sawPrompt       = false;
        bool setName         = false;
        bool captured        = false;
        bool actionTriggered = false;
        bool closed          = false;
        FolderViewEditNewPromptDebugSnapshot snapshot{};
    } probe{};

    RunEditNewPromptModalCycle(mainWindow,
                               [&](const HWND prompt) noexcept
    {
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        probe.setName         = DebugSetFolderViewEditNewPromptText(L"plain.noeditor");
        probe.captured        = DebugGetFolderViewEditNewPromptSnapshot(probe.snapshot);
        probe.actionTriggered = DebugConfirmFolderViewEditNewPrompt();
        probe.closed          = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(probe.sawPrompt, L"Edit New no-editor prompt did not open.");
    state.Require(probe.setName, L"Failed to set Edit New no-editor file name.");
    state.Require(probe.captured, L"Failed to capture Edit New no-editor snapshot.");
    state.Require(! probe.snapshot.editorComboEnabled, L"Edit New Editor combo should be disabled when no editor applies.");
    state.Require(probe.snapshot.editorActionIds.empty(), L"Edit New Editor combo should not list unrelated editors.");
    state.Require(probe.snapshot.selectedEditorActionId.empty(), L"Edit New should not select an editor when none applies.");
    state.Require(probe.actionTriggered, L"Failed to confirm Edit New no-editor prompt.");
    state.Require(probe.closed, L"Edit New no-editor prompt did not close after confirmation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        ec.clear();
        if (std::filesystem::exists(newFile, ec) && g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"plain.noeditor"))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    ec.clear();
    state.Require(std::filesystem::exists(newFile, ec), L"Edit New should create the requested file even when no editor applies.");
    state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, L"plain.noeditor"), L"Edit New no-editor creation should refresh the pane.");
    state.Require(g_folderWindow.DebugGetFocusedItemDisplayName(FolderWindow::Pane::Left) == L"plain.noeditor",
                  L"Edit New no-editor creation should focus the new file.");

    return state.failure.empty();
}

[[nodiscard]] bool TestChangeAttributesOptionsPromptUsesDxUiSurfaceAndReturnsOptions(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root      = suiteRoot / L"work" / (L"change_attributes_options_prompt_" + NewGuidText());
    const std::filesystem::path alphaPath = root / L"alpha.txt";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    const auto clearReadOnlyBits = wil::scope_exit([&] { static_cast<void>(::SetFileAttributesW(alphaPath.c_str(), FILE_ATTRIBUTE_ARCHIVE)); });

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Change Attributes prompt test root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create alpha.txt for Change Attributes prompt test.");
    state.Require(::SetFileAttributesW(alphaPath.c_str(), FILE_ATTRIBUTE_ARCHIVE | FILE_ATTRIBUTE_HIDDEN) != FALSE,
                  L"Failed to set initial hidden attribute for Change Attributes prompt test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePane                                    = wil::scope_exit([&]
    {
        g_folderWindow.DebugSetNextChangeAttributesOptions(std::nullopt);
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    g_folderWindow.DebugSetNextChangeAttributesOptions(std::nullopt);
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for Change Attributes prompt test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Change Attributes prompt test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Change Attributes prompt test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view name) noexcept { return name == L"alpha.txt"; }, true);
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Expected alpha.txt selected before Change Attributes prompt test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeResult final
    {
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool capturedInitial   = false;
        bool setState          = false;
        bool cycledArchive     = false;
        bool capturedEdited    = false;
        bool capturedCycled    = false;
        bool actionTriggered   = false;
        bool closed            = false;
        ChangeAttributesOptionsPromptDebugSnapshot initialSnapshot{};
        ChangeAttributesOptionsPromptDebugSnapshot editedSnapshot{};
        ChangeAttributesOptionsPromptDebugSnapshot cycledSnapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
    } probe{};

    RunChangeAttributesOptionsPromptModalCycle(mainWindow,
                                               [&](const HWND prompt) noexcept
    {
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        probe.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        probe.capturedInitial   = DebugGetChangeAttributesOptionsPromptSnapshot(probe.initialSnapshot);
        probe.uiaPatternStats   = CollectVisibleUiaDescendantPatternStats(prompt);
        probe.setState          = DebugSetChangeAttributesOptionsPromptState(static_cast<uint8_t>(FolderWindow::AttributeChangeState::Set),
                                                                             static_cast<uint8_t>(FolderWindow::AttributeChangeState::Clear),
                                                                             static_cast<uint8_t>(FolderWindow::AttributeChangeState::LeaveUnchanged),
                                                                             static_cast<uint8_t>(FolderWindow::AttributeChangeState::Clear),
                                                                             true);
        probe.cycledArchive     = DebugCycleChangeAttributesOptionsPromptArchive();
        probe.capturedCycled    = DebugGetChangeAttributesOptionsPromptSnapshot(probe.cycledSnapshot);
        probe.setState = probe.setState && DebugSetChangeAttributesOptionsPromptState(static_cast<uint8_t>(FolderWindow::AttributeChangeState::Set),
                                                                                      static_cast<uint8_t>(FolderWindow::AttributeChangeState::Clear),
                                                                                      static_cast<uint8_t>(FolderWindow::AttributeChangeState::LeaveUnchanged),
                                                                                      static_cast<uint8_t>(FolderWindow::AttributeChangeState::Set),
                                                                                      true);
        probe.capturedEdited  = DebugGetChangeAttributesOptionsPromptSnapshot(probe.editedSnapshot);
        probe.actionTriggered = DebugConfirmChangeAttributesOptionsPrompt();
        probe.closed          = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(probe.sawPrompt, L"Change Attributes options prompt did not open.");
    state.Require(probe.ownedByMainWindow, L"Change Attributes options prompt should be owned by the main window.");
    state.Require(probe.capturedInitial, L"Failed to capture initial Change Attributes options prompt snapshot.");
    state.Require(probe.initialSnapshot.usesDxUiHost, L"Change Attributes options prompt should use a DxUi host.");
    state.Require(probe.initialSnapshot.visibleChildWindowCount <= 1u, L"Change Attributes options prompt should not expose native child-control fallback.");
    state.Require(probe.initialSnapshot.visibleNativeChildControlCount <= 1u,
                  L"Change Attributes options prompt should not expose a Win32 dialog-template form.");
    state.Require(probe.initialSnapshot.dialogClassName == L"RedSalamander.ChangeAttributesOptionsPrompt",
                  L"Change Attributes options prompt should use the DxUi host window class.");
    state.Require(probe.initialSnapshot.dialogClassName != L"#32770", L"Change Attributes options prompt should not be an app-owned Win32 dialog template.");
    state.Require(probe.initialSnapshot.dateTimeSectionVisible, L"Change Attributes options prompt should expose Change Date and Time.");
    state.Require(probe.initialSnapshot.modifiedTimeVisible, L"Change Attributes options prompt should expose the Modified time row.");
    state.Require(probe.initialSnapshot.createdTimeVisible, L"Change Attributes options prompt should expose the Created time row.");
    state.Require(probe.initialSnapshot.accessedTimeVisible, L"Change Attributes options prompt should expose the Accessed time row.");
    state.Require(probe.initialSnapshot.includeSubdirectoriesVisible, L"Change Attributes options prompt should expose Include subdirectories.");
    state.Require(! probe.initialSnapshot.includeSubdirectoriesEnabled,
                  L"Change Attributes Include subdirectories should be disabled when only a file is selected.");
    state.Require(probe.uiaPatternStats.has_value(), L"Failed to collect UI Automation stats for Change Attributes options prompt.");
    if (probe.uiaPatternStats.has_value())
    {
        state.Require(probe.uiaPatternStats->comboBoxControlCount == 0u,
                      L"Change Attributes options prompt should use compact checkboxes instead of attribute combo boxes.");
        state.Require(probe.uiaPatternStats->checkBoxControlCount >= 11u,
                      L"Change Attributes options prompt should expose attributes, date/time toggles, recurse, and stream-removal checkboxes.");
        state.Require(probe.uiaPatternStats->togglePatternCount >= 11u,
                      L"Change Attributes options prompt should expose toggle automation for attributes, date/time, recurse, and stream removal.");
        state.Require(probe.uiaPatternStats->buttonControlCount > 0u, L"Change Attributes options prompt should expose command buttons.");
    }
    state.Require(probe.setState, L"Failed to set Change Attributes options prompt state.");
    state.Require(probe.cycledArchive, L"Failed to cycle the Change Attributes archive checkbox.");
    state.Require(probe.capturedCycled, L"Failed to capture cycled Change Attributes options prompt snapshot.");
    state.Require(probe.cycledSnapshot.archive == static_cast<uint8_t>(FolderWindow::AttributeChangeState::LeaveUnchanged),
                  L"Change Attributes tri-state checkbox should be able to cycle back to leave-unchanged.");
    state.Require(probe.capturedEdited, L"Failed to capture edited Change Attributes options prompt snapshot.");
    state.Require(probe.editedSnapshot.readOnly == static_cast<uint8_t>(FolderWindow::AttributeChangeState::Set),
                  L"Change Attributes prompt should store the edited read-only state.");
    state.Require(probe.editedSnapshot.hidden == static_cast<uint8_t>(FolderWindow::AttributeChangeState::Clear),
                  L"Change Attributes prompt should store the edited hidden state.");
    state.Require(probe.editedSnapshot.archive == static_cast<uint8_t>(FolderWindow::AttributeChangeState::Set),
                  L"Change Attributes prompt should store the edited archive state.");
    state.Require(probe.editedSnapshot.removeAlternateDataStreams, L"Change Attributes prompt should expose the remove alternate data streams option.");
    state.Require(probe.actionTriggered, L"Failed to confirm the Change Attributes options prompt.");
    state.Require(probe.closed, L"Change Attributes options prompt did not close after confirmation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const DWORD alphaAttrs = ::GetFileAttributesW(alphaPath.c_str());
    state.Require(alphaAttrs != INVALID_FILE_ATTRIBUTES && (alphaAttrs & FILE_ATTRIBUTE_READONLY) != 0,
                  L"Confirmed Change Attributes prompt should set the read-only bit.");
    state.Require(alphaAttrs != INVALID_FILE_ATTRIBUTES && (alphaAttrs & FILE_ATTRIBUTE_HIDDEN) == 0,
                  L"Confirmed Change Attributes prompt should clear the hidden bit.");
    state.Require(alphaAttrs != INVALID_FILE_ATTRIBUTES && (alphaAttrs & FILE_ATTRIBUTE_ARCHIVE) != 0,
                  L"Confirmed Change Attributes prompt should set the archive bit.");

    const std::optional<FolderWindow::ChangeAttributesReport> report = g_folderWindow.DebugGetLastChangeAttributesReport();
    state.Require(report.has_value(), L"Change Attributes prompt should lead to a report.");
    if (report.has_value())
    {
        state.Require(report->itemsProcessed == 1u, L"Change Attributes prompt should process the selected item.");
        state.Require(report->timesChanged == 0u, L"Change Attributes prompt should not report date/time changes when no date/time row is enabled.");
        state.Require(report->failures == 0u, L"Change Attributes prompt should complete without failures.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestMakeFileListOptionsPromptUsesDxUiSurfaceAndSavesOptions(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root       = suiteRoot / L"work" / (L"make_file_list_options_prompt_" + NewGuidText());
    const std::filesystem::path alphaPath  = root / L"alpha.txt";
    const std::filesystem::path outputPath = root / L"prompt-list.json";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create Make File List prompt test root.");
    state.Require(SelfTest::WriteTextFile(alphaPath, "alpha"), L"Failed to create alpha.txt for Make File List prompt test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto makeFileListSettingsBefore                     = g_settings.makeFileList;
    const auto restoreState                                   = wil::scope_exit([&]
    {
        g_settings.makeFileList = makeFileListSettingsBefore;
        DebugClearMakeFileListAutomation();
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    g_settings.makeFileList = Common::Settings::MakeFileListSettings{};
    DebugClearMakeFileListAutomation();
    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for Make File List prompt test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Make File List prompt test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Make File List prompt test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeResult final
    {
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool capturedInitial   = false;
        bool setState          = false;
        bool capturedEdited    = false;
        bool actionTriggered   = false;
        bool closed            = false;
        MakeFileListOptionsPromptDebugSnapshot initialSnapshot{};
        MakeFileListOptionsPromptDebugSnapshot editedSnapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
    } probe{};

    RunMakeFileListOptionsPromptModalCycle(mainWindow,
                                           [&](const HWND prompt) noexcept
    {
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        probe.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        probe.capturedInitial   = DebugGetMakeFileListOptionsPromptSnapshot(probe.initialSnapshot);
        probe.uiaPatternStats   = CollectVisibleUiaDescendantPatternStats(prompt);
        probe.setState          = DebugSetMakeFileListOptionsPromptState(static_cast<uint8_t>(Common::Settings::MakeFileListSourceMode::CurrentFolder),
                                                                         false,
                                                                         static_cast<uint8_t>(Common::Settings::MakeFileListFormat::Json),
                                                                         static_cast<uint8_t>(Common::Settings::MakeFileListOutputTarget::File),
                                                                         L"{filename}",
                                                                         outputPath.wstring(),
                                                                         true,
                                                                         false,
                                                                         false,
                                                                         false,
                                                                         true,
                                                                         false);
        probe.capturedEdited    = DebugGetMakeFileListOptionsPromptSnapshot(probe.editedSnapshot);
        probe.actionTriggered   = DebugConfirmMakeFileListOptionsPrompt();
        probe.closed            = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(probe.sawPrompt, L"Make File List options prompt did not open.");
    state.Require(probe.ownedByMainWindow, L"Make File List options prompt should be owned by the main window.");
    state.Require(probe.capturedInitial, L"Failed to capture initial Make File List options prompt snapshot.");
    state.Require(probe.initialSnapshot.usesDxUiHost, L"Make File List options prompt should use a DxUi host.");
    state.Require(probe.initialSnapshot.visibleChildWindowCount <= 1u, L"Make File List options prompt should not expose native child-control fallback.");
    state.Require(probe.initialSnapshot.visibleNativeChildControlCount <= 1u, L"Make File List options prompt should not expose a Win32 dialog-template form.");
    state.Require(probe.initialSnapshot.dialogClassName == L"RedSalamander.MakeFileListOptionsPrompt",
                  L"Make File List options prompt should use the DxUi host window class.");
    state.Require(probe.initialSnapshot.dialogClassName != L"#32770", L"Make File List options prompt should not be an app-owned Win32 dialog template.");
    state.Require(probe.uiaPatternStats.has_value(), L"Failed to collect UI Automation stats for Make File List options prompt.");
    if (probe.uiaPatternStats.has_value())
    {
        state.Require(probe.uiaPatternStats->comboBoxControlCount >= 3u, L"Make File List options prompt should expose source, format, and output combos.");
        state.Require(probe.uiaPatternStats->editControlCount >= 2u, L"Make File List options prompt should expose text macro and output file fields.");
        state.Require(probe.uiaPatternStats->buttonControlCount > 0u, L"Make File List options prompt should expose command buttons.");
    }
    state.Require(probe.setState, L"Failed to set Make File List options prompt state.");
    state.Require(probe.capturedEdited, L"Failed to capture edited Make File List options prompt snapshot.");
    state.Require(probe.editedSnapshot.sourceMode == static_cast<uint8_t>(Common::Settings::MakeFileListSourceMode::CurrentFolder),
                  L"Make File List prompt should store the edited source mode.");
    state.Require(probe.editedSnapshot.format == static_cast<uint8_t>(Common::Settings::MakeFileListFormat::Json),
                  L"Make File List prompt should store the edited format.");
    state.Require(probe.editedSnapshot.outputTarget == static_cast<uint8_t>(Common::Settings::MakeFileListOutputTarget::File),
                  L"Make File List prompt should store the edited output target.");
    state.Require(probe.editedSnapshot.textMacro == L"{filename}", L"Make File List prompt should store the edited text macro.");
    state.Require(probe.editedSnapshot.outputFileText == outputPath.wstring(), L"Make File List prompt should store the edited output file.");
    state.Require(probe.editedSnapshot.includeName && ! probe.editedSnapshot.includeFullPath && probe.editedSnapshot.includeAttributes,
                  L"Make File List prompt should store the edited field selection.");
    state.Require(probe.editedSnapshot.outputFileFieldEnabled, L"Make File List prompt should enable the output file field for file output.");
    state.Require(probe.actionTriggered, L"Failed to confirm the Make File List options prompt.");
    state.Require(probe.closed, L"Make File List options prompt did not close after confirmation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (SelfTest::PathExists(outputPath))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(SelfTest::PathExists(outputPath), L"Make File List prompt should create the requested JSON output file.");
    const std::wstring jsonText = ReadUtf8TextFileForCommandSelfTest(outputPath);
    state.Require(jsonText.find(L"\"format\":\"json\"") != std::wstring::npos, L"Make File List prompt JSON should declare its format.");
    state.Require(jsonText.find(L"alpha.txt") != std::wstring::npos, L"Make File List prompt JSON should include current-folder files.");
    state.Require(g_settings.makeFileList.has_value(), L"Make File List prompt should save the last options in settings.");
    if (g_settings.makeFileList.has_value())
    {
        state.Require(g_settings.makeFileList->sourceMode == Common::Settings::MakeFileListSourceMode::CurrentFolder,
                      L"Make File List prompt should save source mode.");
        state.Require(g_settings.makeFileList->format == Common::Settings::MakeFileListFormat::Json, L"Make File List prompt should save format.");
        state.Require(g_settings.makeFileList->outputTarget == Common::Settings::MakeFileListOutputTarget::File,
                      L"Make File List prompt should save output target.");
        state.Require(g_settings.makeFileList->outputFile == outputPath, L"Make File List prompt should save output file.");
        state.Require(g_settings.makeFileList->includeName && ! g_settings.makeFileList->includeFullPath && g_settings.makeFileList->includeAttributes,
                      L"Make File List prompt should save field choices.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestArchivePackPromptUsesDxUiUniqueArchiveAndPackerExtensions(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root          = suiteRoot / L"work" / (L"archive_pack_prompt_" + NewGuidText());
    const std::filesystem::path payloadFolder = root / L"payload";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(payloadFolder), L"Failed to create Pack prompt payload folder.");
    state.Require(SelfTest::WriteTextFile(payloadFolder / L"alpha.txt", "alpha"), L"Failed to create Pack prompt payload file.");
    state.Require(SelfTest::WriteTextFile(root / L"payload.zip", "existing zip"), L"Failed to create existing ZIP conflict fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restoreState                                   = wil::scope_exit([&]
    {
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for Pack prompt test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Pack prompt test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"payload"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for Pack prompt test.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left,
                                                          [](std::wstring_view displayName) noexcept { return displayName == L"payload"; });
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Pack prompt test should select the payload folder.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeResult final
    {
        bool sawPrompt           = false;
        bool ownedByMainWindow   = false;
        bool capturedInitial     = false;
        bool setSecondPacker     = false;
        bool capturedSecond      = false;
        bool setDeleteAfter      = false;
        bool capturedDeleteAfter = false;
        bool actionTriggered     = false;
        bool closed              = false;
        ArchivePackPromptDebugSnapshot initialSnapshot{};
        ArchivePackPromptDebugSnapshot secondPackerSnapshot{};
        ArchivePackPromptDebugSnapshot deleteAfterSnapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
    } probe{};

    RunArchivePackPromptModalCycle(mainWindow,
                                   [&](const HWND prompt) noexcept
    {
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        probe.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        probe.capturedInitial   = DebugGetArchivePackPromptSnapshot(probe.initialSnapshot);
        probe.uiaPatternStats   = CollectVisibleUiaDescendantPatternStats(prompt);
        if (probe.capturedInitial && probe.initialSnapshot.packerCount > 1u)
        {
            probe.setSecondPacker = DebugSetArchivePackPromptPackerIndex(1u);
            probe.capturedSecond  = DebugGetArchivePackPromptSnapshot(probe.secondPackerSnapshot);
        }
        probe.setDeleteAfter      = DebugSetArchivePackPromptDeleteAfter(true);
        probe.capturedDeleteAfter = DebugGetArchivePackPromptSnapshot(probe.deleteAfterSnapshot);
        probe.actionTriggered     = DebugCancelArchivePackPrompt();
        probe.closed              = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(probe.sawPrompt, L"Pack prompt did not open.");
    state.Require(probe.ownedByMainWindow, L"Pack prompt should be owned by the main window.");
    state.Require(probe.capturedInitial, L"Failed to capture initial Pack prompt snapshot.");
    state.Require(probe.initialSnapshot.usesDxUiHost, L"Pack prompt should use a DxUi host.");
    state.Require(probe.initialSnapshot.visibleChildWindowCount <= 1u, L"Pack prompt should not expose native child-control fallback.");
    state.Require(probe.initialSnapshot.visibleNativeChildControlCount <= 1u, L"Pack prompt should not expose a Win32 dialog-template form.");
    state.Require(probe.initialSnapshot.dialogClassName == L"RedSalamander.ArchivePackPrompt", L"Pack prompt should use the DxUi Pack window class.");
    state.Require(probe.initialSnapshot.dialogClassName != L"#32770", L"Pack prompt should not be the stock save-file dialog.");
    state.Require(probe.initialSnapshot.archivePathText.find(L"payload (2).zip") != std::wstring::npos,
                  std::format(L"Pack prompt should suggest a conflict-free ZIP path; saw '{}'.", probe.initialSnapshot.archivePathText));
    state.Require(probe.initialSnapshot.packerDisplayName.find(L"ZIP") != std::wstring::npos,
                  std::format(L"Pack prompt should select ZIP first; saw '{}'.", probe.initialSnapshot.packerDisplayName));
    state.Require(probe.initialSnapshot.packerExtension == L"zip",
                  std::format(L"Pack prompt initial extension should be zip; saw '{}'.", probe.initialSnapshot.packerExtension));
    state.Require(probe.initialSnapshot.packerCount > 1u, L"Pack prompt should enumerate available 7-Zip packer formats in addition to ZIP.");
    state.Require(! probe.initialSnapshot.deleteAfterPacking, L"Pack prompt should leave delete-after-packing off by default.");
    state.Require(probe.initialSnapshot.commandButtonsFitInClient, L"Pack prompt command buttons should fit inside the client area.");
    state.Require(probe.uiaPatternStats.has_value(), L"Failed to collect UI Automation stats for Pack prompt.");
    if (probe.uiaPatternStats.has_value())
    {
        state.Require(probe.uiaPatternStats->comboBoxControlCount >= 2u, L"Pack prompt should expose archive-name and packer combo boxes.");
        state.Require(probe.uiaPatternStats->togglePatternCount > 0u, L"Pack prompt should expose the delete-after-packing checkbox.");
        state.Require(probe.uiaPatternStats->buttonControlCount >= 2u, L"Pack prompt should expose OK and Cancel buttons.");
    }
    state.Require(probe.setSecondPacker, L"Failed to select a non-default Pack prompt packer.");
    state.Require(probe.capturedSecond, L"Failed to capture Pack prompt after changing packer.");
    if (probe.capturedSecond)
    {
        const std::wstring expectedExtension = L"." + probe.secondPackerSnapshot.packerExtension;
        state.Require(probe.secondPackerSnapshot.selectedPackerIndex == 1u, L"Pack prompt should store the edited packer index.");
        state.Require(std::filesystem::path(probe.secondPackerSnapshot.archivePathText).extension().wstring() == expectedExtension,
                      std::format(L"Pack prompt archive extension should follow the selected packer; saw '{}'.", probe.secondPackerSnapshot.archivePathText));
    }
    state.Require(probe.setDeleteAfter, L"Failed to toggle delete-after-packing in Pack prompt.");
    state.Require(probe.capturedDeleteAfter, L"Failed to capture Pack prompt after toggling delete-after-packing.");
    if (probe.capturedDeleteAfter)
    {
        state.Require(probe.deleteAfterSnapshot.deleteAfterPacking, L"Pack prompt should store the delete-after-packing checkbox state.");
    }
    state.Require(probe.actionTriggered, L"Failed to cancel the Pack prompt through its DX action path.");
    state.Require(probe.closed, L"Pack prompt did not close after cancellation.");
    state.Require(! SelfTest::PathExists(root / L"payload (2).zip"), L"Canceling the Pack prompt should not create an archive.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct ConfirmResult final
    {
        bool sawPrompt       = false;
        bool found7zPacker   = false;
        bool captured7z      = false;
        bool actionTriggered = false;
        bool closed          = false;
        ArchivePackPromptDebugSnapshot packer7zSnapshot{};
    } confirm{};

    RunArchivePackPromptModalCycle(mainWindow,
                                   [&](const HWND prompt) noexcept
    {
        confirm.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! confirm.sawPrompt)
        {
            return;
        }

        ArchivePackPromptDebugSnapshot snapshot{};
        if (! DebugGetArchivePackPromptSnapshot(snapshot))
        {
            return;
        }

        for (size_t index = 0u; index < snapshot.packerCount; ++index)
        {
            if (! DebugSetArchivePackPromptPackerIndex(index))
            {
                continue;
            }

            ArchivePackPromptDebugSnapshot candidate{};
            if (! DebugGetArchivePackPromptSnapshot(candidate))
            {
                continue;
            }

            if (candidate.packerExtension == L"7z")
            {
                confirm.found7zPacker    = true;
                confirm.captured7z       = true;
                confirm.packer7zSnapshot = std::move(candidate);
                break;
            }
        }

        confirm.actionTriggered = confirm.found7zPacker ? DebugConfirmArchivePackPrompt() : DebugCancelArchivePackPrompt();
        confirm.closed          = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(confirm.sawPrompt, L"Pack prompt did not reopen for 7z creation.");
    state.Require(confirm.found7zPacker, L"Pack prompt should expose the 7z packer from 7zip.dll.");
    state.Require(confirm.captured7z, L"Failed to capture the 7z Pack prompt state.");
    state.Require(confirm.actionTriggered, L"Failed to confirm the Pack prompt for 7z creation.");
    state.Require(confirm.closed, L"Pack prompt did not close after 7z confirmation.");
    if (confirm.captured7z)
    {
        state.Require(std::filesystem::path(confirm.packer7zSnapshot.archivePathText).extension() == L".7z",
                      std::format(L"7z Pack prompt path should use .7z; saw '{}'.", confirm.packer7zSnapshot.archivePathText));
        state.Require(SelfTest::PathExists(confirm.packer7zSnapshot.archivePathText),
                      std::format(L"Confirmed 7z Pack prompt should create '{}'.", confirm.packer7zSnapshot.archivePathText));
    }

    const std::optional<FolderWindow::ArchiveCommandDebugResult> packResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(packResult.has_value(), L"7z Pack prompt should record a debug pack result.");
    if (packResult.has_value())
    {
        state.Require(SUCCEEDED(packResult->hr), std::format(L"7z Pack prompt should succeed; hr=0x{:08X}.", static_cast<unsigned long>(packResult->hr)));
        state.Require(packResult->entryCount >= 2u, L"7z Pack prompt should include the folder and payload file.");
        state.Require(packResult->bytesProcessed >= 5u, L"7z Pack prompt should report payload bytes.");
    }
    state.Require(SelfTest::PathExists(payloadFolder), L"Pack prompt should not delete sources unless the checkbox is enabled.");

    return state.failure.empty();
}

[[nodiscard]] bool WriteDeflatedZipFixtureForUnpackSelfTest(const std::filesystem::path& archivePath) noexcept
{
    static constexpr std::array<unsigned char, 153u> kArchiveBytes{{
        0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00, 0x6D, 0xBB, 0xA3, 0x5C, 0x01, 0xD0, 0x9C, 0x03, 0x25, 0x00, 0x00, 0x00,
        0x23, 0x00, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x61, 0x6C, 0x70, 0x68, 0x61, 0x2E, 0x74, 0x78, 0x74, 0x4B, 0xCC, 0x29, 0xC8, 0x48,
        0x54, 0x48, 0x2B, 0xCA, 0xCF, 0x55, 0x48, 0x54, 0x48, 0xCE, 0xCF, 0x2D, 0x28, 0x4A, 0x2D, 0x2E, 0x4E, 0x4D, 0x51, 0xA8, 0xCA, 0x2C,
        0x50, 0x48, 0xCB, 0xAC, 0x28, 0x29, 0x2D, 0x4A, 0x05, 0x00, 0x50, 0x4B, 0x01, 0x02, 0x14, 0x00, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00,
        0x6D, 0xBB, 0xA3, 0x5C, 0x01, 0xD0, 0x9C, 0x03, 0x25, 0x00, 0x00, 0x00, 0x23, 0x00, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x61, 0x6C, 0x70, 0x68, 0x61, 0x2E, 0x74, 0x78, 0x74, 0x50,
        0x4B, 0x05, 0x06, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x37, 0x00, 0x00, 0x00, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00,
    }};

    std::ofstream output(archivePath, std::ios::binary | std::ios::trunc);
    if (! output)
    {
        return false;
    }

    output.write(reinterpret_cast<const char*>(kArchiveBytes.data()), static_cast<std::streamsize>(kArchiveBytes.size()));
    return output.good();
}

[[nodiscard]] bool TestArchiveUnpackPromptUsesDxUiDestinationUnpackerAndMask(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root                  = suiteRoot / L"work" / (L"archive_unpack_prompt_" + NewGuidText());
    const std::filesystem::path sourceRoot            = root / L"source";
    const std::filesystem::path outputRoot            = root / L"archives";
    const std::filesystem::path extractRoot           = root / L"extract";
    const std::filesystem::path compressedExtractRoot = root / L"compressed-extract";
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    ec.clear();

    state.Require(SelfTest::EnsureDirectory(sourceRoot / L"nested"), L"Failed to create Unpack prompt source folder.");
    state.Require(SelfTest::EnsureDirectory(outputRoot), L"Failed to create Unpack prompt archive folder.");
    state.Require(SelfTest::WriteTextFile(sourceRoot / L"alpha.txt", "alpha"), L"Failed to create alpha archive fixture.");
    state.Require(SelfTest::WriteTextFile(sourceRoot / L"gamma.log", "gamma"), L"Failed to create gamma archive fixture.");
    state.Require(SelfTest::WriteTextFile(sourceRoot / L"nested" / L"beta.txt", "beta"), L"Failed to create nested archive fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring leftPluginBefore                       = std::wstring(g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left));
    const std::optional<std::filesystem::path> leftPathBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restoreState                                   = wil::scope_exit([&]
    {
        g_folderWindow.DebugClearArchiveCommandOptionsForTest();
        static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, leftPluginBefore));
        if (leftPathBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftPathBefore.value());
        }
    });

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to activate builtin file-system for Unpack prompt test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, sourceRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, sourceRoot, SelfTest::Scale(3000ms)), L"Failed to set left pane path for Unpack prompt setup.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"gamma.log", L"nested"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for Unpack prompt setup.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left, [](std::wstring_view displayName) noexcept {
        return displayName == L"alpha.txt" || displayName == L"gamma.log" || displayName == L"nested";
    });
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 3u, L"Unpack prompt setup should select three top-level items.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path archivePath = outputRoot / L"selected.zip";
    FolderWindow::ArchiveCommandDebugOptions packOptions{};
    packOptions.archivePath = archivePath;
    packOptions.overwrite   = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(packOptions);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/pack"), L"cmd/pane/pack should create the Unpack prompt fixture.");
    PumpPendingMessages();
    const std::optional<FolderWindow::ArchiveCommandDebugResult> packResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(packResult.has_value() && SUCCEEDED(packResult->hr), L"Unpack prompt fixture archive should be created.");
    state.Require(SelfTest::PathExists(archivePath), L"Unpack prompt fixture archive should exist.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, outputRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, outputRoot, SelfTest::Scale(3000ms)), L"Failed to navigate to Unpack prompt archive folder.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"selected.zip"}, SelfTest::Scale(3000ms)),
                  L"Unpack prompt archive folder did not show selected.zip.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"selected.zip"),
                  L"Failed to focus selected.zip for Unpack prompt test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct ProbeResult final
    {
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool capturedInitial   = false;
        bool setDestination    = false;
        bool setMask           = false;
        bool setDeleteAfter    = false;
        bool capturedEdited    = false;
        bool actionTriggered   = false;
        bool closed            = false;
        ArchiveUnpackPromptDebugSnapshot initialSnapshot{};
        ArchiveUnpackPromptDebugSnapshot editedSnapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
    } probe{};

    RunArchiveUnpackPromptModalCycle(mainWindow,
                                     [&](const HWND prompt) noexcept
    {
        probe.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! probe.sawPrompt)
        {
            return;
        }

        probe.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        probe.capturedInitial   = DebugGetArchiveUnpackPromptSnapshot(probe.initialSnapshot);
        probe.uiaPatternStats   = CollectVisibleUiaDescendantPatternStats(prompt);
        probe.setDestination    = DebugSetArchiveUnpackPromptDestinationPath(extractRoot.wstring());
        probe.setMask           = DebugSetArchiveUnpackPromptMask(L"alpha.txt");
        probe.setDeleteAfter    = DebugSetArchiveUnpackPromptDeleteAfter(true);
        probe.capturedEdited    = DebugGetArchiveUnpackPromptSnapshot(probe.editedSnapshot);
        probe.actionTriggered   = DebugConfirmArchiveUnpackPrompt();
        probe.closed            = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(probe.sawPrompt, L"Unpack prompt did not open.");
    state.Require(probe.ownedByMainWindow, L"Unpack prompt should be owned by the main window.");
    state.Require(probe.capturedInitial, L"Failed to capture initial Unpack prompt snapshot.");
    state.Require(probe.initialSnapshot.usesDxUiHost, L"Unpack prompt should use a DxUi host.");
    state.Require(probe.initialSnapshot.visibleChildWindowCount <= 1u, L"Unpack prompt should not expose native child-control fallback.");
    state.Require(probe.initialSnapshot.visibleNativeChildControlCount <= 1u, L"Unpack prompt should not expose a Win32 dialog-template form.");
    state.Require(probe.initialSnapshot.dialogClassName == L"RedSalamander.ArchiveUnpackPrompt", L"Unpack prompt should use the DxUi Unpack window class.");
    state.Require(probe.initialSnapshot.dialogClassName != L"#32770", L"Unpack prompt should not be the stock pick-folder dialog.");
    state.Require(
        probe.initialSnapshot.destinationPathText.find(L"selected") != std::wstring::npos,
        std::format(L"Unpack prompt should suggest a destination derived from the archive name; saw '{}'.", probe.initialSnapshot.destinationPathText));
    state.Require(probe.initialSnapshot.unpackerDisplayName.find(L"ZIP") != std::wstring::npos,
                  std::format(L"Unpack prompt should select ZIP first; saw '{}'.", probe.initialSnapshot.unpackerDisplayName));
    state.Require(probe.initialSnapshot.unpackerExtension == L"zip",
                  std::format(L"Unpack prompt initial extension should be zip; saw '{}'.", probe.initialSnapshot.unpackerExtension));
    state.Require(probe.initialSnapshot.unpackerCount == 1u, L"Unpack prompt should expose the currently supported ZIP unpacker.");
    state.Require(! probe.initialSnapshot.deleteAfterUnpacking, L"Unpack prompt should leave delete-after-unpacking off by default.");
    state.Require(probe.initialSnapshot.maskText == L"*.*", L"Unpack prompt should default the file mask to *.*.");
    state.Require(! probe.initialSnapshot.maskHelpVisible, L"Unpack prompt should keep mask hints collapsed initially.");
    state.Require(probe.initialSnapshot.commandButtonsFitInClient, L"Unpack prompt command buttons should fit inside the client area.");
    state.Require(probe.uiaPatternStats.has_value(), L"Failed to collect UI Automation stats for Unpack prompt.");
    if (probe.uiaPatternStats.has_value())
    {
        state.Require(probe.uiaPatternStats->comboBoxControlCount >= 2u, L"Unpack prompt should expose destination and unpacker combo boxes.");
        state.Require(probe.uiaPatternStats->editControlCount > 0u, L"Unpack prompt should expose an unpack-file mask edit field.");
        state.Require(probe.uiaPatternStats->togglePatternCount > 0u, L"Unpack prompt should expose the delete-after-unpacking checkbox.");
        state.Require(probe.uiaPatternStats->buttonControlCount >= 3u, L"Unpack prompt should expose OK, Cancel, and Help actions.");
    }
    state.Require(probe.setDestination, L"Failed to edit the Unpack prompt destination path.");
    state.Require(probe.setMask, L"Failed to edit the Unpack prompt mask.");
    state.Require(probe.setDeleteAfter, L"Failed to toggle delete-after-unpacking in Unpack prompt.");
    state.Require(probe.capturedEdited, L"Failed to capture edited Unpack prompt snapshot.");
    if (probe.capturedEdited)
    {
        state.Require(probe.editedSnapshot.destinationPathText == extractRoot.wstring(), L"Unpack prompt should store the edited destination path.");
        state.Require(probe.editedSnapshot.maskText == L"alpha.txt", L"Unpack prompt should store the edited file mask.");
        state.Require(probe.editedSnapshot.deleteAfterUnpacking, L"Unpack prompt should store the delete-after-unpacking checkbox state.");
    }
    state.Require(probe.actionTriggered, L"Failed to confirm the Unpack prompt through its DX action path.");
    state.Require(probe.closed, L"Unpack prompt did not close after confirmation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (SelfTest::PathExists(extractRoot / L"alpha.txt") && ! SelfTest::PathExists(archivePath))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    const std::optional<FolderWindow::ArchiveCommandDebugResult> unpackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(unpackResult.has_value(), L"Unpack prompt should record an unpack debug result.");
    if (unpackResult.has_value())
    {
        state.Require(SUCCEEDED(unpackResult->hr),
                      std::format(L"Unpack prompt extraction should succeed; hr=0x{:08X}.", static_cast<unsigned long>(unpackResult->hr)));
        state.Require(unpackResult->destinationPath == extractRoot, L"Unpack prompt extraction should use the confirmed destination.");
        state.Require(unpackResult->entryCount == 1u, L"Unpack prompt mask should extract only the matching file.");
    }
    state.Require(ReadUtf8TextFileForCommandSelfTest(extractRoot / L"alpha.txt") == L"alpha", L"Unpack prompt should extract the masked file.");
    state.Require(! SelfTest::PathExists(extractRoot / L"gamma.log"), L"Unpack prompt mask should skip unmatched root files.");
    state.Require(! SelfTest::PathExists(extractRoot / L"nested" / L"beta.txt"), L"Unpack prompt mask should skip unmatched nested files.");
    state.Require(! SelfTest::PathExists(archivePath), L"Delete-after-unpacking should remove the archive after successful extraction.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path compressedArchivePath = outputRoot / L"compressed.zip";
    state.Require(WriteDeflatedZipFixtureForUnpackSelfTest(compressedArchivePath), L"Failed to create compressed ZIP unpack fixture.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, outputRoot);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, outputRoot, SelfTest::Scale(3000ms)),
                  L"Failed to refresh Unpack prompt archive folder for compressed ZIP fixture.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"compressed.zip"}, SelfTest::Scale(3000ms)),
                  L"Unpack prompt archive folder did not show compressed.zip.");
    g_folderWindow.SetPaneSelectionByDisplayNamePredicate(FolderWindow::Pane::Left,
                                                          [](std::wstring_view displayName) noexcept { return displayName == L"compressed.zip"; });
    state.Require(g_folderWindow.DebugGetSelectedCount(FolderWindow::Pane::Left) == 1u, L"Compressed ZIP unpack should select the compressed fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    FolderWindow::ArchiveCommandDebugOptions unpackOptions{};
    unpackOptions.destinationPath = compressedExtractRoot;
    unpackOptions.overwrite       = true;
    g_folderWindow.DebugSetNextArchiveCommandOptionsForTest(unpackOptions);
    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/unpack"), L"cmd/pane/unpack should run for a compressed ZIP fixture.");
    PumpPendingMessages();

    const auto compressedDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < compressedDeadline)
    {
        PumpPendingMessages();
        if (SelfTest::PathExists(compressedExtractRoot / L"alpha.txt"))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    const std::optional<FolderWindow::ArchiveCommandDebugResult> compressedUnpackResult = g_folderWindow.DebugGetLastArchiveCommandResultForTest();
    state.Require(compressedUnpackResult.has_value(), L"Compressed ZIP unpack should record an unpack debug result.");
    if (compressedUnpackResult.has_value())
    {
        state.Require(SUCCEEDED(compressedUnpackResult->hr),
                      std::format(L"Compressed ZIP extraction should succeed; hr=0x{:08X}.", static_cast<unsigned long>(compressedUnpackResult->hr)));
        state.Require(compressedUnpackResult->destinationPath == compressedExtractRoot, L"Compressed ZIP extraction should use the requested destination.");
        state.Require(compressedUnpackResult->entryCount == 1u, L"Compressed ZIP extraction should extract the fixture file.");
    }
    state.Require(ReadUtf8TextFileForCommandSelfTest(compressedExtractRoot / L"alpha.txt") == L"alpha from a compressed zip fixture",
                  L"Compressed ZIP extraction should preserve the fixture payload.");

    return state.failure.empty();
}

[[nodiscard]] bool TestCreateDirectoryPromptUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"create_directory_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create create-directory test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path created  = root / L"created_via_dxui";
    const std::filesystem::path canceled = root / L"should_not_exist";

    const std::optional<std::filesystem::path> leftBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const auto restorePath                                = wil::scope_exit([&]
    {
        if (leftBefore.has_value())
        {
            g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, leftBefore.value());
        }
    });

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewCreateDirectoryPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for create-directory test.");

    closePrompt();

    struct CreateDirectorySurfaceCycleResult final
    {
        HWND prompt             = nullptr;
        bool sawPrompt          = false;
        bool ownedByMainWindow  = false;
        bool exposesUiaProvider = false;
        bool capturedSnapshot   = false;
        FolderViewCreateDirectoryPromptDebugSnapshot snapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
        std::optional<UiaValuePatternState> valueState;
        std::optional<UiaNamedElementState> buttonState;
        bool setText                = false;
        bool capturedEditedSnapshot = false;
        FolderViewCreateDirectoryPromptDebugSnapshot editedSnapshot{};
        std::optional<UiaValuePatternState> editedValueState;
        bool actionTriggered = false;
        bool closed          = false;
    };

    const auto requireDxSurface =
        [&](const CreateDirectorySurfaceCycleResult& result, std::wstring_view expectedText, std::wstring_view failureContext) noexcept
    {
        state.Require(result.sawPrompt, std::format(L"Create-directory prompt did not open for {}.", failureContext));
        if (! result.sawPrompt)
        {
            return;
        }

        state.Require(result.ownedByMainWindow, std::format(L"Create-directory prompt should be owned by the main window during {}.", failureContext));
        state.Require(result.exposesUiaProvider,
                      std::format(L"Create-directory prompt should answer WM_GETOBJECT with a UI Automation provider during {}.", failureContext));
        state.Require(result.capturedSnapshot, std::format(L"Failed to capture create-directory prompt snapshot during {}.", failureContext));
        state.Require(result.snapshot.usesDxUiHost, std::format(L"Create-directory prompt should use the shared DxUi host during {}.", failureContext));
        state.Require(result.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Create-directory prompt should not expose more than one visible child window during {}; saw {}.",
                                  failureContext,
                                  result.snapshot.visibleChildWindowCount));
        state.Require(result.snapshot.text == expectedText,
                      std::format(L"Create-directory prompt should expose '{}' through the live snapshot during {}.", expectedText, failureContext));
        state.Require(! result.snapshot.createInPath.empty(),
                      std::format(L"Create-directory prompt should expose the destination path during {}.", failureContext));
        state.Require(result.snapshot.nameFieldFocused,
                      std::format(L"Create-directory prompt should keep keyboard focus on the folder-name field during {}.", failureContext));
        state.Require(result.snapshot.selectionStart == 0u,
                      std::format(L"Create-directory prompt should select from the beginning of the suggested folder name during {}.", failureContext));
        state.Require(result.snapshot.selectionEnd == expectedText.size(),
                      std::format(L"Create-directory prompt should select the full suggested folder name during {}; saw selection end {} for '{}'.",
                                  failureContext,
                                  result.snapshot.selectionEnd,
                                  expectedText));

        state.Require(result.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation stats for create-directory prompt during {}.", failureContext));
        if (result.uiaPatternStats.has_value())
        {
            state.Require(result.uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Create-directory prompt should expose visible UI Automation descendants during {}.", failureContext));
            state.Require(result.uiaPatternStats->editControlCount > 0u,
                          std::format(L"Create-directory prompt should expose a visible UI Automation edit descendant during {}.", failureContext));
            state.Require(result.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Create-directory prompt should expose ValuePattern for the folder-name field during {}.", failureContext));
            state.Require(result.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Create-directory prompt should expose visible UI Automation command buttons during {}.", failureContext));
            state.Require(result.uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Create-directory prompt should expose InvokePattern for its visible DX command button during {}.", failureContext));
        }

        state.Require(result.valueState.has_value(),
                      std::format(L"Failed to collect UI Automation ValuePattern state for the create-directory prompt field during {}.", failureContext));
        if (result.valueState.has_value())
        {
            state.Require(! result.valueState->isReadOnly, std::format(L"Create-directory prompt field should remain editable during {}.", failureContext));
            state.Require(! result.valueState->name.empty(),
                          std::format(L"Create-directory prompt editable DX field should expose a stable accessible name during {}.", failureContext));
            state.Require(result.valueState->value == result.snapshot.text,
                          std::format(L"Create-directory prompt field should expose the live folder name through ValuePattern during {}; saw '{}'.",
                                      failureContext,
                                      result.valueState->value));
        }

        state.Require(result.buttonState.has_value(),
                      std::format(L"Create-directory prompt should expose a visible DX command button during {}.", failureContext));
        if (result.buttonState.has_value())
        {
            state.Require(! result.buttonState->name.empty(),
                          std::format(L"Create-directory prompt visible DX command button should expose a stable accessible name during {}.", failureContext));
        }
    };

    CreateDirectorySurfaceCycleResult confirmCycle{};
    RunCreateDirectoryPromptModalCycle(mainWindow,
                                       [&](const HWND prompt) noexcept
    {
        confirmCycle.prompt    = prompt;
        confirmCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! confirmCycle.sawPrompt)
        {
            return;
        }

        confirmCycle.ownedByMainWindow      = IsOwnedBy(prompt, mainWindow);
        confirmCycle.exposesUiaProvider     = WindowExposesUiaProvider(prompt);
        confirmCycle.capturedSnapshot       = DebugGetFolderViewCreateDirectoryPromptSnapshot(confirmCycle.snapshot);
        confirmCycle.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
        confirmCycle.valueState             = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        confirmCycle.buttonState            = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
        confirmCycle.setText                = DebugSetFolderViewCreateDirectoryPromptText(L"  created_via_dxui  ");
        confirmCycle.capturedEditedSnapshot = DebugGetFolderViewCreateDirectoryPromptSnapshot(confirmCycle.editedSnapshot);
        confirmCycle.editedValueState       = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        confirmCycle.actionTriggered        = DebugConfirmFolderViewCreateDirectoryPrompt();
        confirmCycle.closed                 = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    requireDxSurface(confirmCycle, LoadStringResource(nullptr, IDS_NEW_FOLDER_DEFAULT_NAME), L"initial confirm pass");
    state.Require(confirmCycle.setText, L"Failed to set create-directory prompt text during initial confirm pass.");
    state.Require(confirmCycle.capturedEditedSnapshot, L"Failed to capture edited create-directory prompt snapshot during initial confirm pass.");
    state.Require(confirmCycle.editedSnapshot.text == L"  created_via_dxui  ",
                  std::format(L"Create-directory prompt should keep the edited text before confirm during the initial confirm pass; saw '{}'.",
                              confirmCycle.editedSnapshot.text));
    state.Require(confirmCycle.editedValueState.has_value(),
                  L"Failed to collect edited UI Automation ValuePattern state for the create-directory prompt field during initial confirm pass.");
    if (confirmCycle.editedValueState.has_value())
    {
        state.Require(confirmCycle.editedValueState->value == confirmCycle.editedSnapshot.text,
                      std::format(L"Create-directory prompt ValuePattern should track the edited text during the initial confirm pass; saw '{}'.",
                                  confirmCycle.editedValueState->value));
    }
    state.Require(confirmCycle.actionTriggered, L"Failed to confirm create-directory prompt during initial confirm pass.");
    state.Require(confirmCycle.closed, L"Create-directory prompt did not close after confirm during initial confirm pass.");

    const auto createDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < createDeadline)
    {
        PumpPendingMessages();
        ec.clear();
        if (std::filesystem::exists(created, ec))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(std::filesystem::exists(created, ec), L"Create-directory command did not create the directory.");

    CreateDirectorySurfaceCycleResult cancelCycle{};
    RunCreateDirectoryPromptModalCycle(mainWindow,
                                       [&](const HWND prompt) noexcept
    {
        cancelCycle.prompt    = prompt;
        cancelCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! cancelCycle.sawPrompt)
        {
            return;
        }

        cancelCycle.ownedByMainWindow      = IsOwnedBy(prompt, mainWindow);
        cancelCycle.exposesUiaProvider     = WindowExposesUiaProvider(prompt);
        cancelCycle.capturedSnapshot       = DebugGetFolderViewCreateDirectoryPromptSnapshot(cancelCycle.snapshot);
        cancelCycle.uiaPatternStats        = CollectVisibleUiaDescendantPatternStats(prompt);
        cancelCycle.valueState             = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        cancelCycle.buttonState            = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
        cancelCycle.setText                = DebugSetFolderViewCreateDirectoryPromptText(L"should_not_exist");
        cancelCycle.capturedEditedSnapshot = DebugGetFolderViewCreateDirectoryPromptSnapshot(cancelCycle.editedSnapshot);
        cancelCycle.editedValueState       = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        cancelCycle.actionTriggered        = DebugCancelFolderViewCreateDirectoryPrompt();
        cancelCycle.closed                 = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    requireDxSurface(cancelCycle, LoadStringResource(nullptr, IDS_NEW_FOLDER_DEFAULT_NAME), L"reopened cancel pass");
    state.Require(cancelCycle.setText, L"Failed to set create-directory prompt cancel text during reopened cancel pass.");
    state.Require(cancelCycle.capturedEditedSnapshot, L"Failed to capture edited create-directory prompt snapshot during reopened cancel pass.");
    state.Require(cancelCycle.editedValueState.has_value(),
                  L"Failed to collect edited UI Automation ValuePattern state for the create-directory prompt field during reopened cancel pass.");
    if (cancelCycle.editedValueState.has_value())
    {
        state.Require(cancelCycle.editedValueState->value == cancelCycle.editedSnapshot.text,
                      std::format(L"Create-directory prompt ValuePattern should track the edited text during the reopened cancel pass; saw '{}'.",
                                  cancelCycle.editedValueState->value));
    }
    state.Require(cancelCycle.actionTriggered, L"Failed to cancel create-directory prompt during reopened cancel pass.");
    state.Require(cancelCycle.closed, L"Create-directory prompt did not close after cancel during reopened cancel pass.");

    state.Require(! std::filesystem::exists(canceled, ec), L"Cancel should not create the directory.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCreateDirectoryPromptLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"create_directory_live_dx_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create create-directory live interaction test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path created  = root / L"created_via_live_dx";
    const std::filesystem::path canceled = root / L"should_not_exist_live";
    const std::wstring defaultFolderName = LoadStringResource(nullptr, IDS_NEW_FOLDER_DEFAULT_NAME);
    state.Require(! defaultFolderName.empty(), L"Failed to resolve the default folder name for create-directory prompt.");
    if (defaultFolderName.empty())
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewCreateDirectoryPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for create-directory live interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto mutateCreateDirectoryPromptText = [&](const HWND prompt, std::wstring_view expectedValue, std::wstring_view context) noexcept
    {
        FolderViewCreateDirectoryPromptDebugSnapshot snapshot{};
        state.Require(DebugGetFolderViewCreateDirectoryPromptSnapshot(snapshot),
                      std::format(L"Failed to capture create-directory prompt snapshot during {}.", context));
        state.Require(snapshot.usesDxUiHost, std::format(L"Create-directory prompt should use the shared DxUi host during {}.", context));
        state.Require(snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Create-directory prompt should not expose more than one visible child window during {}; saw {}.",
                                  context,
                                  snapshot.visibleChildWindowCount));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto initialValueState = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        state.Require(initialValueState.has_value(), std::format(L"Create-directory prompt should expose a visible editable DX field during {}.", context));
        if (! initialValueState.has_value() || ! state.failure.empty())
        {
            return false;
        }

        state.Require(! initialValueState->name.empty(),
                      std::format(L"Create-directory prompt DX edit should expose a stable accessible name during {}.", context));
        state.Require(! initialValueState->isReadOnly, std::format(L"Create-directory prompt DX edit should remain editable during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto waitForEditValue = [&](std::wstring_view value) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                FolderViewCreateDirectoryPromptDebugSnapshot updatedSnapshot{};
                const bool capturedSnapshot = DebugGetFolderViewCreateDirectoryPromptSnapshot(updatedSnapshot);
                const auto valueState       = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
                if (capturedSnapshot && updatedSnapshot.text == value && valueState.has_value() && valueState->value == value)
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            FolderViewCreateDirectoryPromptDebugSnapshot updatedSnapshot{};
            const bool capturedSnapshot = DebugGetFolderViewCreateDirectoryPromptSnapshot(updatedSnapshot);
            const auto valueState       = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
            return capturedSnapshot && updatedSnapshot.text == value && valueState.has_value() && valueState->value == value;
        };

        state.Require(DebugSetFolderViewCreateDirectoryPromptText(expectedValue),
                      std::format(L"Failed to set create-directory prompt text during {}.", context));
        state.Require(waitForEditValue(expectedValue), std::format(L"Create-directory prompt edit did not update after live interaction during {}.", context));

        return state.failure.empty();
    };

    struct CreateDirectoryLiveCycleResult final
    {
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool mutated           = false;
        std::optional<UiaValuePatternState> valueState;
        bool actionTriggered = false;
        bool closed          = false;
    };

    closePrompt();

    CreateDirectoryLiveCycleResult cancelCycle{};
    RunCreateDirectoryPromptModalCycle(mainWindow,
                                       [&](const HWND prompt) noexcept
    {
        cancelCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! cancelCycle.sawPrompt)
        {
            return;
        }

        cancelCycle.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        cancelCycle.mutated           = mutateCreateDirectoryPromptText(prompt, L"should_not_exist_live", L"live create-directory cancel");
        cancelCycle.actionTriggered   = DebugCancelFolderViewCreateDirectoryPrompt();
        cancelCycle.closed            = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(cancelCycle.sawPrompt, L"Create-directory prompt did not open for live cancel interaction.");
    state.Require(cancelCycle.ownedByMainWindow, L"Create-directory prompt should be owned by the main window during live cancel interaction.");
    state.Require(cancelCycle.mutated, L"Create-directory prompt live DX edit mutation failed before cancel.");
    state.Require(cancelCycle.actionTriggered, L"Create-directory prompt did not close through the debug cancel path after live DX interaction.");
    state.Require(cancelCycle.closed, L"Create-directory prompt did not close after live DX cancel.");
    state.Require(! std::filesystem::exists(canceled, ec), L"Cancel should not create the directory after live interaction.");

    CreateDirectoryLiveCycleResult confirmCycle{};
    RunCreateDirectoryPromptModalCycle(mainWindow,
                                       [&](const HWND prompt) noexcept
    {
        confirmCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! confirmCycle.sawPrompt)
        {
            return;
        }

        confirmCycle.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
        confirmCycle.valueState        = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
        confirmCycle.mutated           = mutateCreateDirectoryPromptText(prompt, L"created_via_live_dx", L"live create-directory confirm after cancel reopen");
        confirmCycle.actionTriggered   = DebugConfirmFolderViewCreateDirectoryPrompt();
        confirmCycle.closed            = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(confirmCycle.sawPrompt, L"Create-directory prompt did not reopen for live confirm interaction.");
    state.Require(confirmCycle.ownedByMainWindow, L"Create-directory prompt should be owned by the main window during live confirm interaction.");
    state.Require(confirmCycle.valueState.has_value(), L"Create-directory prompt should expose a visible editable DX field after live cancel reopen.");
    state.Require(confirmCycle.valueState.has_value() && confirmCycle.valueState->value == defaultFolderName,
                  std::format(L"Create-directory prompt should restore the default folder name '{}' after live DX cancel reopen.", defaultFolderName));
    state.Require(confirmCycle.mutated, L"Create-directory prompt live DX edit mutation failed before confirm after cancel reopen.");
    state.Require(confirmCycle.actionTriggered, L"Create-directory prompt did not close through the debug confirm path after live DX interaction.");
    state.Require(confirmCycle.closed, L"Create-directory prompt did not close after the debug confirm path after cancel reopen.");

    const auto createDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < createDeadline)
    {
        PumpPendingMessages();
        ec.clear();
        if (std::filesystem::exists(created, ec))
        {
            break;
        }
        std::this_thread::sleep_for(20ms);
    }

    state.Require(std::filesystem::exists(created, ec), L"Create-directory live interaction did not create the directory after cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestCreateDirectoryPromptSuggestsUniqueSelectedDefault(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"create_directory_unique_default_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create unique-default create-directory test root.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring defaultFolderName = LoadStringResource(nullptr, IDS_NEW_FOLDER_DEFAULT_NAME);
    state.Require(! defaultFolderName.empty(), L"Failed to resolve the default folder name for unique create-directory prompt validation.");
    if (defaultFolderName.empty())
    {
        return false;
    }

    const std::wstring suggestedOne = std::format(L"{} ({})", defaultFolderName, 1);
    const std::wstring suggestedTwo = std::format(L"{} ({})", defaultFolderName, 2);
    state.Require(SelfTest::EnsureDirectory(root / defaultFolderName), L"Failed to create the baseline default folder.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewCreateDirectoryPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for unique create-directory prompt validation.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {defaultFolderName}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for unique create-directory prompt validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto captureSettledSnapshot = [&](std::wstring_view context, FolderViewCreateDirectoryPromptDebugSnapshot& snapshot, const HWND prompt) noexcept
    {
        state.Require(prompt != nullptr && IsWindow(prompt) != FALSE, std::format(L"Create-directory prompt did not open for {}.", context));
        if (! prompt || IsWindow(prompt) == FALSE)
        {
            return false;
        }

        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            snapshot = {};
            if (DebugGetFolderViewCreateDirectoryPromptSnapshot(snapshot) && snapshot.usesDxUiHost && snapshot.visibleChildWindowCount <= 1u &&
                snapshot.createInPath == root.wstring() && ! snapshot.text.empty())
            {
                return true;
            }

            PumpPendingMessages();
            std::this_thread::sleep_for(10ms);
        }

        state.Require(false,
                      std::format(L"Create-directory prompt did not settle for {}; usesDxUiHost={}, visibleChildren={}, createInPath='{}', text='{}'.",
                                  context,
                                  snapshot.usesDxUiHost ? L"yes" : L"no",
                                  snapshot.visibleChildWindowCount,
                                  snapshot.createInPath,
                                  snapshot.text));
        return false;
    };

    const auto requireSuggestedSelection =
        [&](const FolderViewCreateDirectoryPromptDebugSnapshot& snapshot, std::wstring_view expectedText, std::wstring_view context) noexcept
    {
        state.Require(snapshot.text == expectedText,
                      std::format(L"Create-directory prompt should suggest '{}' during {}; saw '{}'.", expectedText, context, snapshot.text));
        state.Require(snapshot.nameFieldFocused, std::format(L"Create-directory prompt should focus the name field during {}.", context));
        state.Require(snapshot.selectionStart == 0u,
                      std::format(L"Create-directory prompt should select from the beginning of the suggestion during {}.", context));
        state.Require(snapshot.selectionEnd == expectedText.size(),
                      std::format(L"Create-directory prompt should select the full suggested name during {}; saw selection [{}..{}] for '{}'.",
                                  context,
                                  snapshot.selectionStart,
                                  snapshot.selectionEnd,
                                  snapshot.text));
        state.Require(snapshot.validationText.empty(), std::format(L"Create-directory prompt should not show a validation error during {}.", context));
    };

    struct UniqueDefaultCycleResult final
    {
        bool sawPrompt        = false;
        bool capturedSnapshot = false;
        FolderViewCreateDirectoryPromptDebugSnapshot snapshot{};
        bool actionTriggered = false;
        bool closed          = false;
    };

    UniqueDefaultCycleResult confirmCycle{};
    RunCreateDirectoryPromptModalCycle(mainWindow,
                                       [&](const HWND prompt) noexcept
    {
        confirmCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! confirmCycle.sawPrompt)
        {
            return;
        }

        confirmCycle.capturedSnapshot = captureSettledSnapshot(L"the initial unique-default confirm pass", confirmCycle.snapshot, prompt);
        if (! confirmCycle.capturedSnapshot)
        {
            return;
        }

        confirmCycle.actionTriggered = DebugConfirmFolderViewCreateDirectoryPrompt();
        confirmCycle.closed          = confirmCycle.actionTriggered && WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(confirmCycle.sawPrompt, L"Create-directory prompt did not open for the initial unique-default confirm pass.");
    state.Require(confirmCycle.capturedSnapshot, L"Failed to capture the create-directory prompt during the initial unique-default confirm pass.");
    requireSuggestedSelection(confirmCycle.snapshot, suggestedOne, L"the initial unique-default confirm pass");
    state.Require(confirmCycle.actionTriggered, L"Failed to confirm the initial unique-default create-directory prompt.");
    state.Require(confirmCycle.closed, L"Create-directory prompt did not close after the initial unique-default confirm pass.");

    const auto createDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < createDeadline)
    {
        PumpPendingMessages();
        ec.clear();
        if (std::filesystem::exists(root / suggestedOne, ec))
        {
            break;
        }

        std::this_thread::sleep_for(20ms);
    }

    state.Require(std::filesystem::exists(root / suggestedOne, ec),
                  std::format(L"Create-directory prompt did not create the suggested folder '{}'.", suggestedOne));
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {defaultFolderName, suggestedOne}, SelfTest::Scale(3000ms)),
                  L"Pane contents did not refresh after accepting the suggested unique folder name.");
    if (! state.failure.empty())
    {
        return false;
    }

    UniqueDefaultCycleResult reopenCycle{};
    RunCreateDirectoryPromptModalCycle(mainWindow,
                                       [&](const HWND prompt) noexcept
    {
        reopenCycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
        if (! reopenCycle.sawPrompt)
        {
            return;
        }

        reopenCycle.capturedSnapshot = captureSettledSnapshot(L"the reopened unique-default cancel pass", reopenCycle.snapshot, prompt);
        if (! reopenCycle.capturedSnapshot)
        {
            return;
        }

        reopenCycle.actionTriggered = DebugCancelFolderViewCreateDirectoryPrompt();
        reopenCycle.closed          = reopenCycle.actionTriggered && WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
    });

    state.Require(reopenCycle.sawPrompt, L"Create-directory prompt did not reopen for the unique-default cancel pass.");
    state.Require(reopenCycle.capturedSnapshot, L"Failed to capture the reopened create-directory prompt for unique-default validation.");
    requireSuggestedSelection(reopenCycle.snapshot, suggestedTwo, L"the reopened unique-default cancel pass");
    state.Require(reopenCycle.actionTriggered, L"Failed to cancel the reopened unique-default create-directory prompt.");
    state.Require(reopenCycle.closed, L"Create-directory prompt did not close after the reopened unique-default cancel pass.");
    state.Require(! std::filesystem::exists(root / suggestedTwo, ec),
                  std::format(L"Cancel should not create the next suggested unique folder '{}'.", suggestedTwo));

    return state.failure.empty();
}

[[nodiscard]] bool TestCreateDirectoryPromptEnterAndEscapeRouteOk(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"create_directory_prompt_enter_escape_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create create-directory prompt Enter/Escape test root.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewCreateDirectoryPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for create-directory prompt Enter/Escape test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct CreateDirectoryKeyCycleResult final
    {
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool setText           = false;
        bool closed            = false;
    };

    const auto runCloseKey = [&](const WPARAM key,
                                 std::wstring_view requestedName,
                                 const std::filesystem::path& expectedPath,
                                 const bool shouldCreate,
                                 std::wstring_view context) noexcept -> bool
    {
        std::filesystem::remove_all(expectedPath, ec);

        CreateDirectoryKeyCycleResult cycle{};
        RunCreateDirectoryPromptModalCycle(mainWindow,
                                           [&](const HWND prompt) noexcept
        {
            cycle.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
            if (! cycle.sawPrompt)
            {
                return;
            }

            cycle.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
            cycle.setText           = DebugSetFolderViewCreateDirectoryPromptText(requestedName);
            PostMessageW(prompt, WM_KEYDOWN, key, 0);
            PostMessageW(prompt, WM_KEYUP, key, 0);
            cycle.closed = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        });

        state.Require(cycle.sawPrompt, std::format(L"Create-directory prompt did not open for {}.", context));
        state.Require(cycle.ownedByMainWindow, std::format(L"Create-directory prompt should be owned by the main window during {}.", context));
        state.Require(cycle.setText, std::format(L"Failed to seed create-directory prompt text during {}.", context));
        state.Require(cycle.closed, std::format(L"Create-directory prompt did not close after {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            ec.clear();
            if (std::filesystem::exists(expectedPath, ec) == shouldCreate)
            {
                break;
            }
            std::this_thread::sleep_for(20ms);
        }

        ec.clear();
        const bool exists = std::filesystem::exists(expectedPath, ec);
        state.Require(
            exists == shouldCreate,
            std::format(
                L"Create-directory prompt {} path '{}' after {}.", shouldCreate ? L"should create" : L"should not create", expectedPath.native(), context));
        return state.failure.empty();
    };

    if (! runCloseKey(VK_RETURN, L"created_via_enter", root / L"created_via_enter", true, L"Enter default-button routing"))
    {
        return false;
    }

    if (! runCloseKey(VK_ESCAPE, L"should_not_exist_escape", root / L"should_not_exist_escape", false, L"Escape cancel routing"))
    {
        return false;
    }

    state.Require(GetFolderViewCreateDirectoryPromptHandle() == nullptr || IsWindow(GetFolderViewCreateDirectoryPromptHandle()) == FALSE,
                  L"Create-directory prompt should not remain open after Enter/Escape validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestCreateDirectoryPromptLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root = suiteRoot / L"work" / (L"create_directory_prompt_churn_" + NewGuidText());
    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create create-directory prompt churn test root.");
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

    const auto closePrompt = [&]() noexcept
    {
        if (const HWND prompt = GetFolderViewCreateDirectoryPromptHandle(); prompt && IsWindow(prompt) != FALSE)
        {
            PostMessageW(prompt, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prompt, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupPrompt = wil::scope_exit([&]() noexcept { closePrompt(); });

    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for create-directory prompt churn test.");
    if (! state.failure.empty())
    {
        return false;
    }

    struct CreateDirectoryChurnCycleResult final
    {
        bool sawPrompt         = false;
        bool ownedByMainWindow = false;
        bool capturedSnapshot  = false;
        FolderViewCreateDirectoryPromptDebugSnapshot snapshot{};
        std::optional<UiaDescendantPatternStats> uiaPatternStats;
        std::optional<UiaValuePatternState> valueState;
        std::optional<UiaNamedElementState> buttonState;
        bool setText         = false;
        bool actionTriggered = false;
        bool closed          = false;
    };

    std::vector<std::wstring> createdNames;
    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const bool accept                         = (cycle % 2u) == 0u;
        const std::wstring requestedName          = accept ? std::format(L"created_dir_{:02}", cycle) : std::format(L"should_not_exist_{:02}", cycle);
        const std::filesystem::path requestedPath = root / requestedName;

        closePrompt();

        CreateDirectoryChurnCycleResult cycleResult{};
        RunCreateDirectoryPromptModalCycle(mainWindow,
                                           [&](const HWND prompt) noexcept
        {
            cycleResult.sawPrompt = prompt != nullptr && IsWindow(prompt) != FALSE;
            if (! cycleResult.sawPrompt)
            {
                return;
            }

            cycleResult.ownedByMainWindow = IsOwnedBy(prompt, mainWindow);
            cycleResult.capturedSnapshot  = DebugGetFolderViewCreateDirectoryPromptSnapshot(cycleResult.snapshot);
            cycleResult.uiaPatternStats   = CollectVisibleUiaDescendantPatternStats(prompt);
            cycleResult.valueState        = CollectVisibleDescendantValuePatternState(prompt, UIA_EditControlTypeId);
            cycleResult.buttonState       = CollectVisibleDescendantNamedElementState(prompt, UIA_ButtonControlTypeId);
            cycleResult.setText           = DebugSetFolderViewCreateDirectoryPromptText(requestedName);
            cycleResult.actionTriggered   = accept ? DebugConfirmFolderViewCreateDirectoryPrompt() : DebugCancelFolderViewCreateDirectoryPrompt();
            cycleResult.closed            = WaitForWindowClosed(prompt, SelfTest::Scale(3000ms));
        });

        state.Require(cycleResult.sawPrompt, std::format(L"Create-directory prompt did not open during cycle {}.", cycle));
        state.Require(cycleResult.ownedByMainWindow, std::format(L"Create-directory prompt should be owned by the main window during cycle {}.", cycle));
        state.Require(cycleResult.capturedSnapshot, std::format(L"Failed to capture create-directory prompt snapshot during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.usesDxUiHost, std::format(L"Create-directory prompt should use a DxUi host during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Create-directory prompt should not expose more than one visible child window during cycle {}; saw {}.",
                                  cycle,
                                  cycleResult.snapshot.visibleChildWindowCount));
        state.Require(cycleResult.snapshot.text == LoadStringResource(nullptr, IDS_NEW_FOLDER_DEFAULT_NAME),
                      std::format(L"Create-directory prompt should start with the default folder name during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.nameFieldFocused,
                      std::format(L"Create-directory prompt should keep focus on the folder-name field during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.selectionStart == 0u,
                      std::format(L"Create-directory prompt should select from the beginning of the default folder name during cycle {}.", cycle));
        state.Require(cycleResult.snapshot.selectionEnd == cycleResult.snapshot.text.size(),
                      std::format(L"Create-directory prompt should select the full default folder name during cycle {}; saw selection [{}..{}] for '{}'.",
                                  cycle,
                                  cycleResult.snapshot.selectionStart,
                                  cycleResult.snapshot.selectionEnd,
                                  cycleResult.snapshot.text));

        state.Require(cycleResult.uiaPatternStats.has_value(),
                      std::format(L"Failed to collect UI Automation stats for create-directory prompt during cycle {}.", cycle));
        if (cycleResult.uiaPatternStats.has_value())
        {
            state.Require(cycleResult.uiaPatternStats->visibleElementCount >= 3u,
                          std::format(L"Create-directory prompt should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->editControlCount > 0u,
                          std::format(L"Create-directory prompt should expose a visible editable field during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Create-directory prompt should expose ValuePattern during cycle {}.", cycle));
            state.Require(cycleResult.uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Create-directory prompt should expose visible command buttons during cycle {}.", cycle));
        }

        state.Require(cycleResult.valueState.has_value(),
                      std::format(L"Create-directory prompt should expose a visible editable ValuePattern field during cycle {}.", cycle));
        if (cycleResult.valueState.has_value())
        {
            state.Require(! cycleResult.valueState->isReadOnly, std::format(L"Create-directory prompt field should remain editable during cycle {}.", cycle));
            state.Require(cycleResult.valueState->value == cycleResult.snapshot.text,
                          std::format(L"Create-directory prompt ValuePattern should start with '{}' during cycle {}.", cycleResult.snapshot.text, cycle));
            state.Require(! cycleResult.valueState->name.empty(),
                          std::format(L"Create-directory prompt field should expose a stable accessible name during cycle {}.", cycle));
        }

        state.Require(cycleResult.buttonState.has_value(),
                      std::format(L"Create-directory prompt should expose a visible DX command button during cycle {}.", cycle));
        if (cycleResult.buttonState.has_value())
        {
            state.Require(! cycleResult.buttonState->name.empty(),
                          std::format(L"Create-directory prompt visible DX command button should expose a stable accessible name during cycle {}.", cycle));
        }

        state.Require(cycleResult.setText, std::format(L"Failed to set create-directory prompt text during cycle {}.", cycle));
        state.Require(cycleResult.actionTriggered,
                      std::format(L"Failed to close the create-directory prompt through the DX action path during cycle {}.", cycle));
        state.Require(cycleResult.closed, std::format(L"Create-directory prompt did not close after cycle {}.", cycle));

        if (accept)
        {
            const auto createDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
            while (std::chrono::steady_clock::now() < createDeadline)
            {
                PumpPendingMessages();
                ec.clear();
                if (std::filesystem::exists(requestedPath, ec))
                {
                    break;
                }
                std::this_thread::sleep_for(20ms);
            }

            state.Require(std::filesystem::exists(requestedPath, ec), std::format(L"Create-directory cycle {} did not create '{}'.", cycle, requestedName));
            createdNames.push_back(requestedName);

            for (const std::wstring& createdName : createdNames)
            {
                state.Require(g_folderWindow.DebugHasItemDisplayName(FolderWindow::Pane::Left, createdName),
                              std::format(L"Pane contents should show '{}' after create-directory cycle {}.", createdName, cycle));
            }
        }
        else
        {
            state.Require(! std::filesystem::exists(requestedPath, ec),
                          std::format(L"Create-directory cancel cycle {} should not create '{}'.", cycle, requestedName));
        }

        for (const std::wstring& createdName : createdNames)
        {
            state.Require(std::filesystem::exists(root / createdName, ec),
                          std::format(L"Previously created directory '{}' should remain after cycle {}.", createdName, cycle));
        }

        const HWND lingeringPrompt = GetFolderViewCreateDirectoryPromptHandle();
        state.Require(lingeringPrompt == nullptr || IsWindow(lingeringPrompt) == FALSE,
                      std::format(L"Create-directory prompt should not remain open after cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closePrompt();
            return false;
        }
    }

    return state.failure.empty();
}

[[nodiscard]] bool WaitForItemPropertiesLoadedSnapshot(ItemPropertiesWindowDebugSnapshot& outSnapshot, std::chrono::milliseconds timeout) noexcept;

[[nodiscard]] bool TestPaneItemPropertiesUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";
    constexpr uint64_t kFileSizeBytes    = 1536u;
    const std::string filePayload(static_cast<size_t>(kFileSizeBytes), 'x');

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties test root.");
    state.Require(SelfTest::WriteTextFile(filePath, filePayload), L"Failed to create item-properties test file.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupWindow = wil::scope_exit([&]() noexcept { closeWindow(); });

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for item-properties test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)), L"Pane contents not ready for item-properties test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"), L"Failed to focus alpha.txt.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto openPropertiesWindow = [&](std::wstring_view context) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/openProperties during {}.", context));

        const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(properties != nullptr && IsWindow(properties) != FALSE, std::format(L"Item Properties window did not open during {}.", context));
        if (properties && IsWindow(properties) != FALSE)
        {
            state.Require(! IsOwnedBy(properties, mainWindow),
                          std::format(L"Item Properties window should not stay owned above the main window during {}.", context));
            state.Require(WaitForWindowExposesUiaProvider(properties, SelfTest::Scale(3000ms)),
                          std::format(L"Item Properties window should answer WM_GETOBJECT during {}.", context));
        }
        return properties;
    };

    const auto validatePropertiesWindow = [&](const HWND properties, std::wstring_view context) noexcept
    {
        ItemPropertiesWindowDebugSnapshot snapshot{};
        state.Require(WaitForItemPropertiesLoadedSnapshot(snapshot, SelfTest::Scale(5000ms)),
                      std::format(L"Failed to capture loaded Item Properties snapshot during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(! snapshot.loadFailed, std::format(L"Item Properties load should not fail during {}.", context));
        state.Require(snapshot.usesDxUiHost, std::format(L"Item Properties window should use the shared DxUi host during {}.", context));
        state.Require(snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Item Properties window should expose at most the shared DX text-bridge child during {}; saw {}.",
                                  context,
                                  snapshot.visibleChildWindowCount));
        state.Require(snapshot.layoutOverflowRightDip <= 1.0f,
                      std::format(L"Item Properties layout should fit the initial viewport during {}; right overflow={:.1f} DIP.",
                                  context,
                                  snapshot.layoutOverflowRightDip));
        state.Require(snapshot.sectionCount > 0u, std::format(L"Item Properties window should expose at least one parsed section during {}.", context));
        state.Require(snapshot.fieldCount > 0u, std::format(L"Item Properties window should expose at least one parsed field during {}.", context));
        state.Require(snapshot.contentText.find(L"alpha.txt") != std::wstring::npos,
                      std::format(L"Item Properties window text should include the selected file name during {}.", context));
        const std::wstring exactSizeText    = std::format(L"{} bytes", kFileSizeBytes);
        const std::wstring expectedSizeText = std::format(L"{} ({})", FormatBytesCompact(kFileSizeBytes), exactSizeText);
        state.Require(snapshot.contentText.find(expectedSizeText) != std::wstring::npos,
                      std::format(L"Item Properties window should include compact and exact file size '{}' during {}.", expectedSizeText, context));
        state.Require(snapshot.contentText.find(L"Parent: ") == std::wstring::npos,
                      std::format(L"Item Properties window should omit duplicate Parent row during {}.", context));
        state.Require(snapshot.contentText.find(L"Root: ") == std::wstring::npos,
                      std::format(L"Item Properties window should omit duplicate Root row during {}.", context));
        state.Require(snapshot.contentText.find(L"Extension: ") == std::wstring::npos,
                      std::format(L"Item Properties window should omit duplicate Extension row during {}.", context));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(properties);
        state.Require(uiaPatternStats.has_value(), std::format(L"Failed to collect live UI Automation stats for Item Properties window during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Item Properties window should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->textControlCount > 0u,
                          std::format(L"Item Properties window should expose visible text descendants for the card rows during {}.", context));
            state.Require(uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Item Properties window should expose a visible command button during {}.", context));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Item Properties window should expose InvokePattern for the close button during {}.", context));
        }

        wil::com_ptr<IUIAutomationElement> fileNameElement;
        state.Require(FindMatchingVisibleDescendantElement(properties, UIA_TextControlTypeId, L"alpha.txt", fileNameElement.put()) && fileNameElement,
                      std::format(L"Item Properties visible card rows should include the selected file name during {}.", context));
        return state.failure.empty();
    };

    const auto closePropertiesWindow = [&](const HWND properties, std::wstring_view context) noexcept
    {
        PostMessageW(properties, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)), std::format(L"Item Properties window did not close during {}.", context));
        return state.failure.empty();
    };

    const HWND properties = openPropertiesWindow(L"initial baseline surface probe");
    if (! properties || IsWindow(properties) == FALSE)
    {
        return false;
    }

    state.Require(validatePropertiesWindow(properties, L"initial baseline surface probe"), L"Initial Item Properties baseline DX surface validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePropertiesWindow(properties, L"initial baseline surface probe"), L"Initial Item Properties baseline close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedProperties = openPropertiesWindow(L"reopened baseline surface probe");
    if (! reopenedProperties || IsWindow(reopenedProperties) == FALSE)
    {
        return false;
    }

    state.Require(validatePropertiesWindow(reopenedProperties, L"reopened baseline surface probe"),
                  L"Reopened Item Properties baseline DX surface validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePropertiesWindow(reopenedProperties, L"reopened baseline surface probe"), L"Reopened Item Properties baseline close validation failed.");
    return state.failure.empty();
}

[[nodiscard]] HRESULT WriteAlternateStreamForItemPropertiesTest(const std::filesystem::path& path,
                                                                std::wstring_view streamName,
                                                                std::string_view payload) noexcept
{
    if (path.empty() || streamName.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring streamPath = path.wstring();
    streamPath.push_back(L':');
    streamPath.append(streamName);

    wil::unique_handle stream(CreateFileW(
        streamPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! stream)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0u ? lastError : ERROR_GEN_FAILURE);
    }

    if (payload.empty())
    {
        return S_OK;
    }
    if (payload.size() > static_cast<size_t>((std::numeric_limits<DWORD>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    DWORD written            = 0u;
    const DWORD bytesToWrite = static_cast<DWORD>(payload.size());
    if (WriteFile(stream.get(), payload.data(), bytesToWrite, &written, nullptr) == 0 || written != bytesToWrite)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0u ? lastError : ERROR_WRITE_FAULT);
    }

    return S_OK;
}

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfind has deleted copy operations by design.
[[nodiscard]] std::optional<uint64_t> FindAlternateStreamSizeForItemPropertiesTest(const std::filesystem::path& path, std::wstring_view streamName) noexcept
{
    if (path.empty() || streamName.empty())
    {
        return std::nullopt;
    }

    std::wstring expected;
    expected.reserve(streamName.size() + 7u);
    expected.push_back(L':');
    expected.append(streamName);
    expected.append(L":$DATA");

    WIN32_FIND_STREAM_DATA streamData{};
    wil::unique_hfind findHandle(FindFirstStreamW(path.c_str(), FindStreamInfoStandard, &streamData, 0));
    if (! findHandle)
    {
        return std::nullopt;
    }

    for (;;)
    {
        if (expected == streamData.cStreamName && streamData.StreamSize.QuadPart >= 0)
        {
            return static_cast<uint64_t>(streamData.StreamSize.QuadPart);
        }

        streamData = {};
        if (FindNextStreamW(findHandle.get(), &streamData) == 0)
        {
            return std::nullopt;
        }
    }
}
#pragma warning(pop)

[[nodiscard]] bool WaitForItemPropertiesLoadedSnapshot(ItemPropertiesWindowDebugSnapshot& outSnapshot, std::chrono::milliseconds timeout) noexcept
{
    using namespace std::chrono_literals;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (DebugGetItemPropertiesWindowSnapshot(outSnapshot) && ! outSnapshot.loading)
        {
            return true;
        }
        std::this_thread::sleep_for(20ms);
    }

    PumpPendingMessages();
    return DebugGetItemPropertiesWindowSnapshot(outSnapshot) && ! outSnapshot.loading;
}

[[nodiscard]] bool TestPaneItemPropertiesTimestampsFollowGeneral([[maybe_unused]] HWND mainWindow, CaseState& state) noexcept
{
    constexpr std::string_view kOutOfOrderPropertiesJson =
        R"({"version":1,"title":"properties","sections":[{"title":"general","fields":[{"key":"name","value":"remote-item.txt"},{"key":"type","value":"file"}]},{"title":"remote","fields":[{"key":"remotePath","value":"/home/test/remote-item.txt"}]},{"title":"timestamps","fields":[{"key":"modified","value":"133455072000000000"}]},{"title":"connection","fields":[{"key":"protocol","value":"FTP"}]}]})";

    const std::wstring content = DebugBuildItemPropertiesContentTextFromJson(kOutOfOrderPropertiesJson);
    state.Require(! content.empty(), L"Item Properties parser should build content text for an out-of-order properties document.");

    const size_t generalPos    = content.find(L"General\r\n");
    const size_t timestampsPos = content.find(L"Timestamps\r\n");
    const size_t remotePos     = content.find(L"Remote\r\n");
    const size_t connectionPos = content.find(L"Connection\r\n");

    state.Require(generalPos != std::wstring::npos, L"Item Properties content should include a normalized General section.");
    state.Require(timestampsPos != std::wstring::npos, L"Item Properties content should include a normalized Timestamps section.");
    state.Require(remotePos != std::wstring::npos, L"Item Properties content should include a normalized Remote section.");
    state.Require(connectionPos != std::wstring::npos, L"Item Properties content should include a normalized Connection section.");
    state.Require(generalPos < timestampsPos && timestampsPos < remotePos && remotePos < connectionPos,
                  L"Item Properties should place Timestamps immediately after General before later provider-specific sections.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesShowsLoadingWhilePropertiesLoad(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_loading_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties loading test root.");
    state.Require(SelfTest::WriteTextFile(filePath, "hello from async item properties loading validation"),
                  L"Failed to create item-properties loading test file.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupWindow = wil::scope_exit([&]() noexcept { closeWindow(); });

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties loading test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for item-properties loading test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for item-properties loading test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for item-properties loading test.");
    if (! state.failure.empty())
    {
        return false;
    }

    DebugSetNextItemPropertiesLoadDelayMs(450u);
    const auto clearLoadDelay = wil::scope_exit([]() noexcept { DebugSetNextItemPropertiesLoadDelayMs(0u); });
    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                  L"Shortcut dispatch failed for cmd/pane/openProperties during loading validation.");

    const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(properties != nullptr && IsWindow(properties) != FALSE, L"Item Properties window did not open during loading validation.");
    if (! properties || IsWindow(properties) == FALSE)
    {
        return false;
    }

    ItemPropertiesWindowDebugSnapshot loadingSnapshot{};
    state.Require(DebugGetItemPropertiesWindowSnapshot(loadingSnapshot), L"Failed to capture Item Properties loading snapshot.");
    state.Require(loadingSnapshot.loading, L"Item Properties should expose a loading state while the properties provider is still running.");
    state.Require(loadingSnapshot.contentText.find(LoadStringResource(nullptr, IDS_PROPERTIES_LOADING)) != std::wstring::npos,
                  L"Item Properties loading state should expose the loading message in copy text.");

    const UINT dpi                = GetDpiForWindow(properties);
    const LPARAM loadingHoverSpot = MAKELPARAM(MulDiv(40, static_cast<int>(dpi), 96), MulDiv(90, static_cast<int>(dpi), 96));
    SendMessageW(properties, WM_MOUSEMOVE, 0, loadingHoverSpot);
    PumpPendingMessages();

    ItemPropertiesWindowDebugSnapshot loadedSnapshot{};
    state.Require(WaitForItemPropertiesLoadedSnapshot(loadedSnapshot, SelfTest::Scale(5000ms)), L"Item Properties window did not complete the delayed load.");
    SendMessageW(properties, WM_MOUSEMOVE, 0, loadingHoverSpot);
    PumpPendingMessages();
    state.Require(! loadedSnapshot.loadFailed, L"Item Properties delayed load should complete successfully.");
    state.Require(loadedSnapshot.contentText.find(L"alpha.txt") != std::wstring::npos,
                  L"Item Properties loaded content should include the selected file after the async load.");

    closeWindow();
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesStreamsCanRemove(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_streams_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";
    const std::filesystem::path dirPath  = root / L"beta";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties streams test root.");
    state.Require(SelfTest::EnsureDirectory(dirPath), L"Failed to create item-properties stream directory.");
    state.Require(SelfTest::WriteTextFile(filePath, "base file for stream properties validation"), L"Failed to create item-properties stream file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HRESULT hrFileZone   = WriteAlternateStreamForItemPropertiesTest(filePath, L"Zone.Identifier", "zone-id");
    const HRESULT hrFileNotes  = WriteAlternateStreamForItemPropertiesTest(filePath, L"notes", "stream-notes");
    const HRESULT hrFolderNote = WriteAlternateStreamForItemPropertiesTest(dirPath, L"folder-note", "folder-stream");
    if (FAILED(hrFileZone) || FAILED(hrFileNotes) || FAILED(hrFolderNote))
    {
        if (hrFileZone == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) || hrFileZone == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ||
            hrFileNotes == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) || hrFileNotes == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ||
            hrFolderNote == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) || hrFolderNote == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        {
            return state.Skip(L"Alternate data streams are not supported by the temporary filesystem.");
        }

        state.Require(false,
                      std::format(L"Failed to create alternate streams for item-properties validation (fileZone=0x{0:08X}, fileNotes=0x{1:08X}, "
                                  L"folderNote=0x{2:08X}).",
                                  static_cast<unsigned long>(hrFileZone),
                                  static_cast<unsigned long>(hrFileNotes),
                                  static_cast<unsigned long>(hrFolderNote)));
        return false;
    }

    state.Require(FindAlternateStreamSizeForItemPropertiesTest(filePath, L"Zone.Identifier").value_or(0u) == 7u,
                  L"Zone.Identifier stream should exist with the expected size before properties removal.");
    state.Require(FindAlternateStreamSizeForItemPropertiesTest(filePath, L"notes").value_or(0u) == 12u,
                  L"notes stream should exist with the expected size before properties removal.");
    state.Require(FindAlternateStreamSizeForItemPropertiesTest(dirPath, L"folder-note").value_or(0u) == 13u,
                  L"folder stream should exist with the expected size before properties removal.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupWindow = wil::scope_exit([&]() noexcept { closeWindow(); });

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties streams test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for stream properties test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt", L"beta"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for stream properties test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto openPropertiesForItem = [&](std::wstring_view displayName) noexcept -> HWND
    {
        closeWindow();
        state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, displayName),
                      std::format(L"Failed to focus '{}' for stream properties test.", displayName));
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/openProperties on '{}'.", displayName));
        const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(properties != nullptr && IsWindow(properties) != FALSE,
                      std::format(L"Item Properties window did not open for stream item '{}'.", displayName));
        return properties;
    };

    const auto waitForStreamCount = [&](const size_t expectedStreamCount, ItemPropertiesWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (DebugGetItemPropertiesWindowSnapshot(outSnapshot) && ! outSnapshot.loading && outSnapshot.streamCount == expectedStreamCount)
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        return DebugGetItemPropertiesWindowSnapshot(outSnapshot) && ! outSnapshot.loading && outSnapshot.streamCount == expectedStreamCount;
    };

    HWND properties = openPropertiesForItem(L"alpha.txt");
    if (! properties || IsWindow(properties) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    ItemPropertiesWindowDebugSnapshot snapshot{};
    state.Require(waitForStreamCount(2u, snapshot), L"Failed to capture loaded file stream properties snapshot.");
    state.Require(snapshot.streamCount == 2u, std::format(L"File properties should list two streams; saw {}.", snapshot.streamCount));
    state.Require(snapshot.removableStreamCount == 2u,
                  std::format(L"File properties should expose two removable streams; saw {}.", snapshot.removableStreamCount));
    state.Require(snapshot.viewableStreamCount == 2u,
                  std::format(L"File properties should expose two viewable streams; saw {}.", snapshot.viewableStreamCount));
    state.Require(snapshot.contentText.find(L"Zone.Identifier: 7 bytes") != std::wstring::npos,
                  L"File properties should include Zone.Identifier stream name and size.");
    state.Require(snapshot.contentText.find(L"notes: 12 bytes") != std::wstring::npos, L"File properties should include notes stream name and size.");
    if (! state.failure.empty())
    {
        return false;
    }

    g_folderWindow.CloseAllViewers();
    const auto closeViewers          = wil::scope_exit([&]() noexcept { g_folderWindow.CloseAllViewers(); });
    const size_t baselineViewerCount = g_folderWindow.DebugGetViewerInstanceCount();
    state.Require(SUCCEEDED(DebugOpenItemPropertiesStream(L"notes")), L"Failed to open alternate stream from Item Properties in ViewerText.");
    const auto viewerDeadline = std::chrono::steady_clock::now() + SelfTest::Scale(5000ms);
    while (std::chrono::steady_clock::now() < viewerDeadline &&
           ! (g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount + 1u && g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-text")))
    {
        PumpPendingMessages();
        std::this_thread::sleep_for(20ms);
    }
    state.Require(g_folderWindow.DebugGetViewerInstanceCount() == baselineViewerCount + 1u,
                  L"Opening an alternate stream from Item Properties should create a ViewerText instance.");
    state.Require(g_folderWindow.DebugHasViewerPluginId(L"builtin/viewer-text"), L"Opening an alternate stream from Item Properties should use ViewerText.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SUCCEEDED(DebugRemoveItemPropertiesStream(L"Zone.Identifier")), L"Failed to remove Zone.Identifier stream from file properties.");
    state.Require(FindAlternateStreamSizeForItemPropertiesTest(filePath, L"Zone.Identifier") == std::nullopt,
                  L"Zone.Identifier stream should be gone after properties removal.");
    state.Require(waitForStreamCount(1u, snapshot), L"Failed to capture file stream properties snapshot after stream removal.");
    state.Require(snapshot.streamCount == 1u, std::format(L"File properties should refresh to one stream after removal; saw {}.", snapshot.streamCount));
    state.Require(snapshot.contentText.find(L"Zone.Identifier") == std::wstring::npos,
                  L"Removed stream should no longer be present in refreshed file properties.");
    state.Require(snapshot.contentText.find(L"notes: 12 bytes") != std::wstring::npos,
                  L"Remaining stream should still be present in refreshed file properties.");
    if (! state.failure.empty())
    {
        return false;
    }

    closeWindow();
    properties = openPropertiesForItem(L"beta");
    if (! properties || IsWindow(properties) == FALSE || ! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForStreamCount(1u, snapshot), L"Failed to capture loaded folder stream properties snapshot.");
    state.Require(snapshot.streamCount == 1u, std::format(L"Folder properties should list one stream; saw {}.", snapshot.streamCount));
    state.Require(snapshot.removableStreamCount == 1u,
                  std::format(L"Folder properties should expose one removable stream; saw {}.", snapshot.removableStreamCount));
    state.Require(snapshot.viewableStreamCount == 1u,
                  std::format(L"Folder properties should expose one viewable stream; saw {}.", snapshot.viewableStreamCount));
    state.Require(snapshot.contentText.find(L"folder-note: 13 bytes") != std::wstring::npos, L"Folder properties should include folder stream name and size.");
    state.Require(InvokeVisibleDescendantByName(properties, UIA_ButtonControlTypeId, LoadStringResource(nullptr, IDS_PROPERTIES_STREAM_REMOVE)),
                  L"Failed to invoke the visible stream Remove button from properties.");
    state.Require(waitForStreamCount(0u, snapshot), L"Folder properties did not refresh to zero streams after invoking the Remove button.");
    state.Require(FindAlternateStreamSizeForItemPropertiesTest(dirPath, L"folder-note") == std::nullopt,
                  L"Folder stream should be gone after properties removal.");
    state.Require(snapshot.streamCount == 0u, std::format(L"Folder properties should refresh to zero streams after removal; saw {}.", snapshot.streamCount));
    state.Require(snapshot.contentText.find(L"folder-note") == std::wstring::npos,
                  L"Removed folder stream should no longer be present in refreshed properties.");
    state.Require(snapshot.contentText.find(LoadStringResource(nullptr, IDS_PROPERTIES_STREAMS_TITLE)) == std::wstring::npos,
                  L"Properties text should omit the Streams section after the last stream is removed.");

    closeWindow();
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_live_dx_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties live interaction test root.");
    state.Require(SelfTest::WriteTextFile(filePath, "hello from dxui item properties"), L"Failed to create item-properties live interaction test file.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupWindow = wil::scope_exit([&]() noexcept { closeWindow(); });

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties live interaction test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for item-properties live interaction test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for item-properties live interaction test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for item-properties live interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto openPropertiesWindow = [&](std::wstring_view context) noexcept
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/openProperties during {}.", context));

        const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(properties != nullptr && IsWindow(properties) != FALSE, std::format(L"Item Properties window did not open for {}.", context));
        if (properties && IsWindow(properties) != FALSE)
        {
            state.Require(! IsOwnedBy(properties, mainWindow),
                          std::format(L"Item Properties window should not stay owned above the main window during {}.", context));
            state.Require(WaitForWindowExposesUiaProvider(properties, SelfTest::Scale(3000ms)),
                          std::format(L"Item Properties window should answer WM_GETOBJECT during {}.", context));
        }
        return properties;
    };

    const auto validatePropertiesWindow = [&](const HWND properties, std::wstring_view context) noexcept
    {
        ItemPropertiesWindowDebugSnapshot snapshot{};
        state.Require(WaitForItemPropertiesLoadedSnapshot(snapshot, SelfTest::Scale(5000ms)),
                      std::format(L"Failed to capture loaded Item Properties snapshot during {}.", context));
        state.Require(! snapshot.loadFailed, std::format(L"Item Properties load should not fail during {}.", context));
        state.Require(snapshot.usesDxUiHost, std::format(L"Item Properties window should use the shared DxUi host during {}.", context));
        state.Require(snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Item Properties window should expose at most the shared DX text-bridge child during {}; saw {}.",
                                  context,
                                  snapshot.visibleChildWindowCount));
        state.Require(snapshot.layoutOverflowRightDip <= 1.0f,
                      std::format(L"Item Properties layout should fit the initial viewport during {}; right overflow={:.1f} DIP.",
                                  context,
                                  snapshot.layoutOverflowRightDip));
        state.Require(snapshot.contentText.find(L"alpha.txt") != std::wstring::npos,
                      std::format(L"Item Properties window text should include the selected file name during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(properties);
        state.Require(uiaPatternStats.has_value(), std::format(L"Failed to collect live UI Automation stats for Item Properties during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->textControlCount > 0u,
                          std::format(L"Item Properties window should expose visible text descendants during {}.", context));
            state.Require(uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Item Properties window should expose a visible close button during {}.", context));
            state.Require(uiaPatternStats->invokePatternCount > 0u, std::format(L"Item Properties window should expose InvokePattern during {}.", context));
        }

        wil::com_ptr<IUIAutomationElement> fileNameElement;
        state.Require(FindMatchingVisibleDescendantElement(properties, UIA_TextControlTypeId, L"alpha.txt", fileNameElement.put()) && fileNameElement,
                      std::format(L"Item Properties visible card rows should include the selected file name during {}.", context));
        return state.failure.empty();
    };

    const std::wstring closeButtonText = LoadStringResource(nullptr, IDS_PROPERTIES_BTN_CLOSE);
    state.Require(! closeButtonText.empty(), L"Item Properties Close button caption should resolve for live UIA interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto closePropertiesWindow = [&](const HWND properties, std::wstring_view context) noexcept
    {
        state.Require(InvokeVisibleDescendantByName(properties, UIA_ButtonControlTypeId, closeButtonText),
                      std::format(L"Failed to invoke the visible DX close button on Item Properties during {}.", context));
        state.Require(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)),
                      std::format(L"Item Properties window did not close after live UIA InvokePattern interaction during {}.", context));
        return state.failure.empty();
    };

    const HWND properties = openPropertiesWindow(L"initial live interaction");
    if (! properties || IsWindow(properties) == FALSE)
    {
        return false;
    }

    state.Require(validatePropertiesWindow(properties, L"initial live interaction"), L"Initial Item Properties live DX validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePropertiesWindow(properties, L"initial live interaction"), L"Initial Item Properties live DX close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedProperties = openPropertiesWindow(L"reopened live interaction");
    if (! reopenedProperties || IsWindow(reopenedProperties) == FALSE)
    {
        return false;
    }

    state.Require(validatePropertiesWindow(reopenedProperties, L"reopened live interaction"), L"Reopened Item Properties live DX validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePropertiesWindow(reopenedProperties, L"reopened live interaction"), L"Reopened Item Properties live DX close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesCtrlAAndCopyExportsFullText(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_copy_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties copy test root.");
    state.Require(SelfTest::WriteTextFile(filePath, "hello from dxui item properties copy validation"), L"Failed to create item-properties copy test file.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupWindow = wil::scope_exit([&]() noexcept { closeWindow(); });

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties copy test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for item-properties copy test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for item-properties copy test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for item-properties copy test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                  L"Shortcut dispatch failed for cmd/pane/openProperties during item-properties copy validation.");

    const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(properties != nullptr && IsWindow(properties) != FALSE, L"Item Properties window did not open for copy validation.");
    if (! properties || IsWindow(properties) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(properties, mainWindow), L"Item Properties window should not stay owned above the main window during copy validation.");
    state.Require(WaitForWindowExposesUiaProvider(properties, SelfTest::Scale(3000ms)),
                  L"Item Properties window should answer WM_GETOBJECT during copy validation.");
    wil::com_ptr<IUIAutomationElement> copyHintElement;
    state.Require(
        FindMatchingVisibleDescendantElement(properties, UIA_TextControlTypeId, LoadStringResource(nullptr, IDS_PROPERTIES_COPY_HINT), copyHintElement.put()) &&
            copyHintElement,
        L"Item Properties window should show the discrete Ctrl+C copy hint at the bottom.");

    ItemPropertiesWindowDebugSnapshot snapshot{};
    state.Require(WaitForItemPropertiesLoadedSnapshot(snapshot, SelfTest::Scale(5000ms)),
                  L"Failed to capture loaded Item Properties snapshot before copy validation.");
    state.Require(snapshot.usesDxUiHost, L"Item Properties window should use the shared DxUi host during copy validation.");
    state.Require(snapshot.visibleChildWindowCount <= 1u,
                  std::format(L"Item Properties window should expose at most the shared DX text-bridge child during copy validation; saw {}.",
                              snapshot.visibleChildWindowCount));
    state.Require(! snapshot.contentText.empty(), L"Item Properties content text should be populated before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SetFocus(properties);
    PumpPendingMessages();
    std::this_thread::sleep_for(20ms);

    ClearClipboardContents(properties);
    SendMessageW(properties, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(properties, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(properties, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(properties, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedText;
    for (size_t retry = 0u; retry < 20u && copiedText.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedText = ReadClipboardUnicodeText(properties);
        if (copiedText.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedText.empty(), L"Item Properties Ctrl+C should copy the full read-only DX body text to the clipboard.");
    state.Require(NormalizeComparisonNewlines(copiedText) == NormalizeComparisonNewlines(snapshot.contentText),
                  L"Item Properties clipboard export should match the full read-only DX body text after Ctrl+C.");

    ItemPropertiesWindowDebugSnapshot postCopySnapshot{};
    state.Require(DebugGetItemPropertiesWindowSnapshot(postCopySnapshot), L"Failed to capture Item Properties snapshot after copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(postCopySnapshot.visibleChildWindowCount <= 1u,
                  std::format(L"Item Properties window should keep visible child windows bounded to the shared DX text bridge after copy validation; saw {}.",
                              postCopySnapshot.visibleChildWindowCount));
    state.Require(postCopySnapshot.resizeFailureCount == 0u, L"Item Properties window should stay resize-clean after copy validation.");
    state.Require(postCopySnapshot.contentText == snapshot.contentText, L"Item Properties content text should remain stable after copy validation.");

    PostMessageW(properties, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)), L"Item Properties window did not close after copy validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesLongRunScrollingStaysBounded(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"w" / (L"ips_" + NewGuidText()) / L"alpha" / L"beta" / L"gamma";
    const std::filesystem::path filePath = root / L"item_props_scroll_payload_long_name_segments_and_extra_wrap_pressure.txt";

    std::error_code ec;
    std::filesystem::remove_all(root.parent_path(), ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties scroll test root.");
    state.Require(SelfTest::WriteTextFile(filePath, "hello from dxui item properties scroll validation"),
                  L"Failed to create item-properties scroll test file.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupWindow = wil::scope_exit([&]() noexcept { closeWindow(); });

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties scroll test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for item-properties scroll test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {filePath.filename().wstring()}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for item-properties scroll test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, filePath.filename().wstring()),
                  L"Failed to focus item-properties scroll test file.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                  L"Shortcut dispatch failed for cmd/pane/openProperties during long-run scroll validation.");

    const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(properties != nullptr && IsWindow(properties) != FALSE, L"Item Properties window did not open for long-run scrolling validation.");
    if (! properties || IsWindow(properties) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](auto&& predicate, ItemPropertiesWindowDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            if (DebugGetItemPropertiesWindowSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        PumpPendingMessages();
        return DebugGetItemPropertiesWindowSnapshot(outSnapshot) && predicate(outSnapshot);
    };
    const auto waitForWindowSize = [&](const int expectedWidthPx, const int expectedHeightPx) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            RECT currentRect{};
            if (GetWindowRect(properties, &currentRect) != FALSE)
            {
                const int widthPx  = (std::max)(1, static_cast<int>(currentRect.right - currentRect.left));
                const int heightPx = (std::max)(1, static_cast<int>(currentRect.bottom - currentRect.top));
                if (widthPx == expectedWidthPx && heightPx == expectedHeightPx)
                {
                    return true;
                }
            }

            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);
        }

        RECT currentRect{};
        if (GetWindowRect(properties, &currentRect) == FALSE)
        {
            return false;
        }

        const int widthPx  = (std::max)(1, static_cast<int>(currentRect.right - currentRect.left));
        const int heightPx = (std::max)(1, static_cast<int>(currentRect.bottom - currentRect.top));
        return widthPx == expectedWidthPx && heightPx == expectedHeightPx;
    };
    struct WindowSizePx
    {
        int width  = 0;
        int height = 0;
    };
    const auto waitForWindowSizeChangedFrom = [&](const int previousWidthPx, const int previousHeightPx, WindowSizePx& outSize) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            RECT currentRect{};
            if (GetWindowRect(properties, &currentRect) != FALSE)
            {
                const int widthPx  = (std::max)(1, static_cast<int>(currentRect.right - currentRect.left));
                const int heightPx = (std::max)(1, static_cast<int>(currentRect.bottom - currentRect.top));
                if (widthPx != previousWidthPx || heightPx != previousHeightPx)
                {
                    outSize = WindowSizePx{.width = widthPx, .height = heightPx};
                    return true;
                }
            }

            PumpPendingMessages();
            std::this_thread::sleep_for(20ms);
        }

        RECT currentRect{};
        if (GetWindowRect(properties, &currentRect) == FALSE)
        {
            return false;
        }

        const int widthPx  = (std::max)(1, static_cast<int>(currentRect.right - currentRect.left));
        const int heightPx = (std::max)(1, static_cast<int>(currentRect.bottom - currentRect.top));
        outSize            = WindowSizePx{.width = widthPx, .height = heightPx};
        return widthPx != previousWidthPx || heightPx != previousHeightPx;
    };

    ItemPropertiesWindowDebugSnapshot snapshot{};
    state.Require(WaitForItemPropertiesLoadedSnapshot(snapshot, SelfTest::Scale(5000ms)), L"Failed to capture initial loaded Item Properties snapshot.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT windowRect{};
    state.Require(GetWindowRect(properties, &windowRect) != FALSE, L"Failed to capture Item Properties window rect for long-run scrolling validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const UINT dpi             = GetDpiForWindow(properties);
    const int currentWidthPx   = (std::max)(1, static_cast<int>(windowRect.right - windowRect.left));
    const int currentHeightPx  = (std::max)(1, static_cast<int>(windowRect.bottom - windowRect.top));
    const int expandedWidthPx  = (std::max)(currentWidthPx, MulDiv(640, static_cast<int>(dpi), 96));
    const int expandedHeightPx = (std::max)(currentHeightPx, MulDiv(420, static_cast<int>(dpi), 96));

    if (expandedWidthPx != currentWidthPx || expandedHeightPx != currentHeightPx)
    {
        state.Require(SetWindowPos(properties, nullptr, 0, 0, expandedWidthPx, expandedHeightPx, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE) != FALSE,
                      L"Failed to normalize Item Properties window size before narrow-width scrolling validation.");
        state.Require(waitForWindowSize(expandedWidthPx, expandedHeightPx),
                      L"Item Properties window did not apply the normalization resize before narrow-width scrolling validation.");
        state.Require(DebugGetItemPropertiesWindowSnapshot(snapshot), L"Failed to capture Item Properties snapshot after normalization resize.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const int narrowWidthPx  = MulDiv(140, static_cast<int>(dpi), 96);
    const int narrowHeightPx = MulDiv(110, static_cast<int>(dpi), 96);
    state.Require(SetWindowPos(properties, nullptr, 0, 0, narrowWidthPx, narrowHeightPx, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE) != FALSE,
                  L"Failed to narrow Item Properties window for long-run scrolling validation.");
    WindowSizePx appliedNarrowSize{};
    state.Require(waitForWindowSizeChangedFrom(expandedWidthPx, expandedHeightPx, appliedNarrowSize),
                  L"Item Properties window did not apply any narrower resize for long-run scrolling validation.");
    state.Require(appliedNarrowSize.width <= expandedWidthPx && appliedNarrowSize.height <= expandedHeightPx,
                  std::format(L"Item Properties window narrow resize grew unexpectedly; applied={}x{}, expanded={}x{}, requested={}x{}.",
                              appliedNarrowSize.width,
                              appliedNarrowSize.height,
                              expandedWidthPx,
                              expandedHeightPx,
                              narrowWidthPx,
                              narrowHeightPx));
    PumpPendingMessages();
    std::this_thread::sleep_for(50ms);
    state.Require(DebugGetItemPropertiesWindowSnapshot(snapshot), L"Failed to capture Item Properties snapshot after narrow-width scrolling validation.");
    state.Require(snapshot.bodyVisibleLineCount > 0u && snapshot.bodyTotalLineCount > 0u,
                  std::format(L"Item Properties window did not expose a readable DX read-only surface after narrow-width validation; current={}x{}, "
                              L"requested={}x{}, applied={}x{}, "
                              L"visibleLines={}, totalLines={}, firstVisibleLine={}, canScroll={}, resizeCount={}, renderCount={}, textLength={}.",
                              currentWidthPx,
                              currentHeightPx,
                              narrowWidthPx,
                              narrowHeightPx,
                              appliedNarrowSize.width,
                              appliedNarrowSize.height,
                              snapshot.bodyVisibleLineCount,
                              snapshot.bodyTotalLineCount,
                              snapshot.bodyFirstVisibleLine,
                              snapshot.bodyCanScrollVertically ? 1 : 0,
                              snapshot.resizeCount,
                              snapshot.renderCount,
                              snapshot.contentText.size()));
    if (! state.failure.empty())
    {
        return false;
    }

    if (snapshot.bodyFirstVisibleLine > 0u)
    {
        size_t previousFirstVisibleLine = snapshot.bodyFirstVisibleLine;
        for (size_t chunk = 0u; chunk < 16u && previousFirstVisibleLine > 0u; ++chunk)
        {
            state.Require(DebugScrollItemPropertiesWindowByWheelDetents(4),
                          std::format(L"Failed to normalize Item Properties window back toward the top before long-run scrolling chunk {}.", chunk));
            state.Require(waitForSnapshot([&](const ItemPropertiesWindowDebugSnapshot& value) noexcept
            { return value.bodyFirstVisibleLine < previousFirstVisibleLine; },
                                          snapshot),
                          std::format(L"Item Properties window did not move back toward the top while normalizing pre-scroll state in chunk {}.", chunk));
            previousFirstVisibleLine = snapshot.bodyFirstVisibleLine;
        }

        state.Require(waitForSnapshot([](const ItemPropertiesWindowDebugSnapshot& value) noexcept { return value.bodyFirstVisibleLine == 0u; }, snapshot),
                      L"Item Properties window did not return to the top before long-run scrolling validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const uint64_t baselineResizeFailureCount = snapshot.resizeFailureCount;
    const uint64_t baselineRenderCount        = snapshot.renderCount;
    size_t previousFirstVisibleLine           = snapshot.bodyFirstVisibleLine;
    size_t successfulScrollChunks             = 0u;

    for (size_t chunk = 0u; chunk < 4u; ++chunk)
    {
        state.Require(DebugScrollItemPropertiesWindowByWheelDetents(-4),
                      std::format(L"Failed to scroll Item Properties window during long-run scroll chunk {}.", chunk));

        ItemPropertiesWindowDebugSnapshot advancedSnapshot{};
        if (waitForSnapshot([&](const ItemPropertiesWindowDebugSnapshot& value) noexcept { return value.bodyFirstVisibleLine > previousFirstVisibleLine; },
                            advancedSnapshot))
        {
            snapshot = advancedSnapshot;
            state.Require(snapshot.usesDxUiHost, std::format(L"Item Properties window lost its DX host during scroll chunk {}.", chunk));
            state.Require(snapshot.visibleChildWindowCount <= 1u,
                          std::format(L"Item Properties window exposed more than the shared DX text-bridge child during scroll chunk {}; saw {}.",
                                      chunk,
                                      snapshot.visibleChildWindowCount));
            state.Require(snapshot.resizeFailureCount == baselineResizeFailureCount,
                          std::format(L"Item Properties window resize failures changed during scroll chunk {}; saw {} vs baseline {}.",
                                      chunk,
                                      snapshot.resizeFailureCount,
                                      baselineResizeFailureCount));
            previousFirstVisibleLine = snapshot.bodyFirstVisibleLine;
            ++successfulScrollChunks;
            continue;
        }

        state.Require(DebugGetItemPropertiesWindowSnapshot(snapshot),
                      std::format(L"Failed to capture Item Properties snapshot after saturated scroll chunk {}.", chunk));
        const size_t maxFirstVisibleLine =
            snapshot.bodyTotalLineCount > snapshot.bodyVisibleLineCount ? snapshot.bodyTotalLineCount - snapshot.bodyVisibleLineCount : 0u;
        state.Require(snapshot.bodyFirstVisibleLine >= maxFirstVisibleLine,
                      std::format(L"Item Properties window did not advance its scroll position during chunk {}; firstVisibleLine={}, maxFirstVisibleLine={}, "
                                  L"visibleLines={}, totalLines={}.",
                                  chunk,
                                  snapshot.bodyFirstVisibleLine,
                                  maxFirstVisibleLine,
                                  snapshot.bodyVisibleLineCount,
                                  snapshot.bodyTotalLineCount));
        break;
    }

    state.Require(successfulScrollChunks > 0u, L"Item Properties window should scroll away from the top during long-run validation.");
    previousFirstVisibleLine = snapshot.bodyFirstVisibleLine;
    state.Require(waitForSnapshot([&](const ItemPropertiesWindowDebugSnapshot& value) noexcept { return value.renderCount > baselineRenderCount; }, snapshot),
                  L"Item Properties window should repaint while the DX read-only surface scrolls.");

    state.Require(snapshot.contentText.find(filePath.filename().wstring()) != std::wstring::npos,
                  L"Item Properties exported content text should still include the selected file name after scroll churn.");

    for (size_t chunk = 0u; chunk < 8u && previousFirstVisibleLine > 0u; ++chunk)
    {
        state.Require(DebugScrollItemPropertiesWindowByWheelDetents(4),
                      std::format(L"Failed to scroll Item Properties window back toward the top during chunk {}.", chunk));
        state.Require(waitForSnapshot([&](const ItemPropertiesWindowDebugSnapshot& value) noexcept
        { return value.bodyFirstVisibleLine < previousFirstVisibleLine; },
                                      snapshot),
                      std::format(L"Item Properties window did not move back toward the top during chunk {}.", chunk));
        previousFirstVisibleLine = snapshot.bodyFirstVisibleLine;
    }

    state.Require(waitForSnapshot([](const ItemPropertiesWindowDebugSnapshot& value) noexcept { return value.bodyFirstVisibleLine == 0u; }, snapshot),
                  L"Item Properties window did not return to the top after reverse scroll churn.");
    state.Require(snapshot.resizeFailureCount == baselineResizeFailureCount,
                  std::format(L"Item Properties window resize failures changed after reverse scroll churn; saw {} vs baseline {}.",
                              snapshot.resizeFailureCount,
                              baselineResizeFailureCount));

    PostMessageW(properties, WM_CLOSE, 0, 0);
    state.Require(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)), L"Item Properties window did not close after long-run scrolling validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_open_close_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties churn test root.");
    state.Require(SelfTest::WriteTextFile(filePath, "hello from dxui item properties open/close validation"),
                  L"Failed to create item-properties churn test file.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties churn test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)), L"Failed to set left pane path for item-properties churn test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for item-properties churn test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for item-properties churn test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);

    constexpr size_t kCycles = 12u;
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/openProperties during cycle {}.", cycle + 1u));
        if (! state.failure.empty())
        {
            return false;
        }

        const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(properties != nullptr && IsWindow(properties) != FALSE, std::format(L"Item Properties window did not open during cycle {}.", cycle + 1u));
        if (! properties || IsWindow(properties) == FALSE)
        {
            return false;
        }

        state.Require(! IsOwnedBy(properties, mainWindow),
                      std::format(L"Item Properties window should not stay owned above the main window during cycle {}.", cycle + 1u));
        state.Require(WaitForWindowExposesUiaProvider(properties, SelfTest::Scale(3000ms)),
                      std::format(L"Item Properties window should answer WM_GETOBJECT during cycle {}.", cycle + 1u));

        ItemPropertiesWindowDebugSnapshot snapshot{};
        state.Require(WaitForItemPropertiesLoadedSnapshot(snapshot, SelfTest::Scale(5000ms)),
                      std::format(L"Failed to capture loaded Item Properties snapshot during cycle {}.", cycle + 1u));
        if (! state.failure.empty())
        {
            closeWindow();
            return false;
        }

        state.Require(snapshot.usesDxUiHost, std::format(L"Item Properties window should use the shared DxUi host during cycle {}.", cycle + 1u));
        state.Require(snapshot.visibleChildWindowCount <= 1u,
                      std::format(L"Item Properties window should keep visible child windows bounded to the shared DX text bridge during cycle {}; saw {}.",
                                  cycle + 1u,
                                  snapshot.visibleChildWindowCount));
        state.Require(snapshot.layoutOverflowRightDip <= 1.0f,
                      std::format(L"Item Properties layout should fit the viewport during cycle {}; right overflow={:.1f} DIP.",
                                  cycle + 1u,
                                  snapshot.layoutOverflowRightDip));
        state.Require(snapshot.sectionCount > 0u, std::format(L"Item Properties window should expose parsed sections during cycle {}.", cycle + 1u));
        state.Require(snapshot.fieldCount > 0u, std::format(L"Item Properties window should expose parsed fields during cycle {}.", cycle + 1u));
        state.Require(snapshot.contentText.find(L"alpha.txt") != std::wstring::npos,
                      std::format(L"Item Properties content should include the selected file name during cycle {}.", cycle + 1u));

        const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(properties);
        state.Require(uiaPatternStats.has_value(), std::format(L"Failed to collect live UI Automation stats for Item Properties during cycle {}.", cycle + 1u));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Item Properties window should expose visible UI Automation descendants during cycle {}.", cycle + 1u));
            state.Require(uiaPatternStats->textControlCount > 0u,
                          std::format(L"Item Properties window should expose visible text descendants during cycle {}.", cycle + 1u));
            state.Require(uiaPatternStats->buttonControlCount > 0u,
                          std::format(L"Item Properties window should expose a visible command button during cycle {}.", cycle + 1u));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Item Properties window should expose InvokePattern during cycle {}.", cycle + 1u));
        }

        wil::com_ptr<IUIAutomationElement> fileNameElement;
        state.Require(FindMatchingVisibleDescendantElement(properties, UIA_TextControlTypeId, L"alpha.txt", fileNameElement.put()) && fileNameElement,
                      std::format(L"Item Properties visible card rows should include the selected file name during cycle {}.", cycle + 1u));

        const auto buttonState = CollectVisibleDescendantNamedElementState(properties, UIA_ButtonControlTypeId);
        state.Require(buttonState.has_value(), std::format(L"Failed to collect Item Properties visible DX command-button state during cycle {}.", cycle + 1u));
        if (buttonState.has_value())
        {
            state.Require(! buttonState->name.empty(),
                          std::format(L"Item Properties visible DX command button should expose a stable accessible name during cycle {}.", cycle + 1u));
        }

        PostMessageW(properties, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)),
                      std::format(L"Item Properties window did not close cleanly during cycle {}.", cycle + 1u));
    }

    state.Require(GetItemPropertiesWindowHandle() == nullptr || IsWindow(GetItemPropertiesWindowHandle()) == FALSE,
                  L"Item Properties window should not remain open after repeated churn.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesAccessKeyRoutesOk(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_access_key_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties access-key test root.");
    state.Require(SelfTest::WriteTextFile(filePath, "hello from dxui item properties access-key validation"),
                  L"Failed to create item-properties access-key test file.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupWindow = wil::scope_exit([&]() noexcept { closeWindow(); });

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties access-key test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for item-properties access-key test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for item-properties access-key test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for item-properties access-key test.");
    if (! state.failure.empty())
    {
        return false;
    }

    FocusFolderViewPane(FolderWindow::Pane::Left);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                  L"Shortcut dispatch failed for cmd/pane/openProperties during access-key validation.");

    const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
    state.Require(properties != nullptr && IsWindow(properties) != FALSE, L"Item Properties window did not open for access-key validation.");
    if (! properties || IsWindow(properties) == FALSE)
    {
        return false;
    }

    state.Require(! IsOwnedBy(properties, mainWindow), L"Item Properties window should not stay owned above the main window during access-key validation.");
    const size_t accessKeyVisibleChildren = CountVisibleChildWindows(properties);
    state.Require(accessKeyVisibleChildren <= 1u,
                  std::format(L"Item Properties window should expose at most the shared DX text-bridge child during access-key validation; saw {}.",
                              accessKeyVisibleChildren));
    state.Require(WaitForWindowExposesUiaProvider(properties, SelfTest::Scale(3000ms)),
                  L"Item Properties window should answer WM_GETOBJECT during access-key validation.");

    SendMessageW(properties, WM_SYSCHAR, static_cast<WPARAM>(L'o'), 0);
    state.Require(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)), L"Item Properties window did not close after the DX OK access key.");
    state.Require(GetItemPropertiesWindowHandle() == nullptr || IsWindow(GetItemPropertiesWindowHandle()) == FALSE,
                  L"Item Properties window should not remain open after the DX OK access key.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPaneItemPropertiesEnterAndEscapeRouteOk(HWND mainWindow, CaseState& state) noexcept
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

    const std::filesystem::path root     = suiteRoot / L"work" / (L"item_properties_enter_escape_" + NewGuidText());
    const std::filesystem::path filePath = root / L"alpha.txt";

    std::error_code ec;
    std::filesystem::remove_all(root, ec);
    state.Require(SelfTest::EnsureDirectory(root), L"Failed to create item-properties Enter/Escape test root.");
    state.Require(SelfTest::WriteTextFile(filePath, "hello from dxui item properties enter/escape validation"),
                  L"Failed to create item-properties Enter/Escape test file.");
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

    const auto closeWindow = []() noexcept
    {
        if (const HWND properties = GetItemPropertiesWindowHandle(); properties && IsWindow(properties) != FALSE)
        {
            PostMessageW(properties, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)));
        }
    };
    const auto cleanupWindow = wil::scope_exit([&]() noexcept { closeWindow(); });

    closeWindow();

    state.Require(SUCCEEDED(g_folderWindow.SetFileSystemPluginForPane(FolderWindow::Pane::Left, L"builtin/file-system")),
                  L"Failed to set local file-system plugin for item-properties Enter/Escape test.");
    g_folderWindow.SetFolderPath(FolderWindow::Pane::Left, root);
    state.Require(WaitForPanePath(FolderWindow::Pane::Left, root, SelfTest::Scale(3000ms)),
                  L"Failed to set left pane path for item-properties Enter/Escape test.");
    state.Require(WaitForPaneItems(FolderWindow::Pane::Left, {L"alpha.txt"}, SelfTest::Scale(3000ms)),
                  L"Pane contents not ready for item-properties Enter/Escape test.");
    state.Require(g_folderWindow.DebugFocusItemByDisplayName(FolderWindow::Pane::Left, L"alpha.txt"),
                  L"Failed to focus alpha.txt for item-properties Enter/Escape test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto openPropertiesWindow = [&](std::wstring_view context) noexcept -> HWND
    {
        FocusFolderViewPane(FolderWindow::Pane::Left);
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/pane/openProperties"),
                      std::format(L"Shortcut dispatch failed for cmd/pane/openProperties during {}.", context));

        const HWND properties = WaitForWindow([]() noexcept { return GetItemPropertiesWindowHandle(); }, SelfTest::Scale(5000ms));
        state.Require(properties != nullptr && IsWindow(properties) != FALSE, std::format(L"Item Properties window did not open for {}.", context));
        if (properties && IsWindow(properties) != FALSE)
        {
            state.Require(! IsOwnedBy(properties, mainWindow),
                          std::format(L"Item Properties window should not stay owned above the main window during {}.", context));
            const size_t visibleChildCount = CountVisibleChildWindows(properties);
            state.Require(
                visibleChildCount <= 1u,
                std::format(L"Item Properties window should expose at most the shared DX text-bridge child during {}; saw {}.", context, visibleChildCount));
            state.Require(WaitForWindowExposesUiaProvider(properties, SelfTest::Scale(3000ms)),
                          std::format(L"Item Properties window should answer WM_GETOBJECT during {}.", context));
        }
        return properties;
    };

    const auto runCloseKey = [&](WPARAM key, std::wstring_view context) noexcept -> bool
    {
        const HWND properties = openPropertiesWindow(context);
        if (! properties || IsWindow(properties) == FALSE)
        {
            return false;
        }

        SendMessageW(properties, WM_KEYDOWN, key, 0);
        SendMessageW(properties, WM_KEYUP, key, 0);
        state.Require(WaitForWindowClosed(properties, SelfTest::Scale(3000ms)), std::format(L"Item Properties window did not close during {}.", context));
        return state.failure.empty();
    };

    if (! runCloseKey(VK_RETURN, L"Enter default-button validation"))
    {
        return false;
    }

    if (! runCloseKey(VK_ESCAPE, L"Escape cancel validation"))
    {
        return false;
    }

    state.Require(GetItemPropertiesWindowHandle() == nullptr || IsWindow(GetItemPropertiesWindowHandle()) == FALSE,
                  L"Item Properties window should not remain open after Enter/Escape validation.");
    return state.failure.empty();
}

} // namespace (tests)

void RunDialogsCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(
        options, suite, L"cmd_app_about_uses_dxui_surface", [=](CaseState& state) noexcept { return TestAboutDialogUsesDxUiSurface(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_app_about_live_dx_interaction", [=](CaseState& state) noexcept { return TestAboutDialogLiveDxInteraction(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_app_about_access_key_routes_ok", [=](CaseState& state) noexcept { return TestAboutDialogAccessKeyRoutesOk(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_about_enter_and_escape_route_ok", [=](CaseState& state) noexcept {
        return TestAboutDialogEnterAndEscapeRouteOk(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_about_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestAboutDialogLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_prompt_uses_alert_overlay_window", [=](CaseState& state) noexcept {
        return TestAppPromptUsesAlertOverlayWindow(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_prompt_access_keys_route_expected_actions", [=](CaseState& state) noexcept {
        return TestAppPromptAccessKeysRouteExpectedActions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_prompt_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestAppPromptLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_fatal_error_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestFatalErrorDialogUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_fatal_error_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestFatalErrorDialogLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_fatal_error_access_key_routes_ok", [=](CaseState& state) noexcept {
        return TestFatalErrorDialogAccessKeyRoutesOk(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_fatal_error_enter_and_escape_route_ok", [=](CaseState& state) noexcept {
        return TestFatalErrorDialogEnterAndEscapeRouteOk(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_fatal_error_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestFatalErrorDialogLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_fatal_error_long_run_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestFatalErrorDialogLongRunScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_app_fatal_error_theme_matrix_keeps_message_legible", [=](CaseState& state) noexcept {
        return TestFatalErrorDialogThemeMatrixKeepsMessageLegible(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_app_splash_uses_dxui_surface", [=](CaseState& state) noexcept { return TestSplashScreenUsesDxUiSurface(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_app_splash_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestSplashScreenLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_app_splash_live_dx_text_update", [=](CaseState& state) noexcept { return TestSplashScreenLiveDxTextUpdate(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_changeCase_prompt_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestPaneChangeCasePromptUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_changeCase_prompt_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPaneChangeCasePromptLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_changeCase_prompt_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestPaneChangeCasePromptLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_changeCase_dialog", [=](CaseState& state) noexcept { return TestChangeCaseDialogAndMultiSelection(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_changeCase", [](CaseState& state) noexcept { return TestChangeCaseCore(state); });
    SelfTest::RunCase(options, suite, L"mask_syntax_wildcards", [](CaseState& state) noexcept { return TestMaskSyntaxWildcardMatching(state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_filter_prompt_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestPaneFilterPromptUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_filter_prompt_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPaneFilterPromptLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_filter_prompt_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestPaneFilterPromptLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"folderView_filter_watermark_empty_state", [=](CaseState& state) noexcept {
        return TestFilterWatermarkBadgeForEmptyState(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_mask_prompt_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestSelectionMaskPromptUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_mask_prompt_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestSelectionMaskPromptLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_selection_mask_prompt_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestSelectionMaskPromptLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_filter_dialog_filtering", [=](CaseState& state) noexcept { return TestPaneFilterDialogFiltering(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_filter_history_restore", [=](CaseState& state) noexcept { return TestPaneFilterHistoryRestore(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"cmd_pane_selection_mask_dialogs", [=](CaseState& state) noexcept { return TestSelectionMaskDialogs(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"cmd_pane_rename_prompt_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestPaneRenamePromptUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_rename_prompt_long_initial_selection_stays_clipped", [=](CaseState& state) noexcept {
        return TestPaneRenamePromptLongInitialSelectionStaysClipped(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_rename_prompt_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPaneRenamePromptLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_rename_prompt_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestPaneRenamePromptLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_editNew_prompt_filters_editor_combo_and_creates_file", [=](CaseState& state) noexcept {
        return TestEditNewPromptFiltersEditorComboAndCreatesFile(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_editNew_prompt_rejects_invalid_and_existing_names", [=](CaseState& state) noexcept {
        return TestEditNewPromptRejectsInvalidAndExistingNames(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_editNew_prompt_creates_file_without_applicable_editor", [=](CaseState& state) noexcept {
        return TestEditNewPromptCreatesFileWithoutApplicableEditor(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_changeAttributes_options_dialog_uses_dxui_not_win32_template", [=](CaseState& state) noexcept {
        return TestChangeAttributesOptionsPromptUsesDxUiSurfaceAndReturnsOptions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_makeFileList_options_dialog_uses_dxui_not_win32_template", [=](CaseState& state) noexcept {
        return TestMakeFileListOptionsPromptUsesDxUiSurfaceAndSavesOptions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions", [=](CaseState& state) noexcept {
        return TestArchivePackPromptUsesDxUiUniqueArchiveAndPackerExtensions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_unpack_prompt_uses_dxui_destination_unpacker_and_mask", [=](CaseState& state) noexcept {
        return TestArchiveUnpackPromptUsesDxUiDestinationUnpackerAndMask(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_createDirectory_prompt_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestCreateDirectoryPromptUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_createDirectory_prompt_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestCreateDirectoryPromptLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_createDirectory_prompt_suggests_unique_selected_default", [=](CaseState& state) noexcept {
        return TestCreateDirectoryPromptSuggestsUniqueSelectedDefault(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_createDirectory_prompt_enter_and_escape_route_ok", [=](CaseState& state) noexcept {
        return TestCreateDirectoryPromptEnterAndEscapeRouteOk(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_createDirectory_prompt_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestCreateDirectoryPromptLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_uses_dxui_surface", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesUsesDxUiSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_shows_loading", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesShowsLoadingWhilePropertiesLoad(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_timestamps_follow_general", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesTimestampsFollowGeneral(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_streams_can_remove", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesStreamsCanRemove(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_ctrlA_copy_exports_full_text", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesCtrlAAndCopyExportsFullText(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_long_run_scrolling_stays_bounded", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesLongRunScrollingStaysBounded(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_access_key_routes_ok", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesAccessKeyRoutesOk(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_pane_itemProperties_window_enter_and_escape_route_ok", [=](CaseState& state) noexcept {
        return TestPaneItemPropertiesEnterAndEscapeRouteOk(mainWindow, state);
    });
}

namespace
{
