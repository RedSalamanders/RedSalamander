namespace
{

[[nodiscard]] bool TestPreferencesDialogCategoryTreeUsesDxUiHost(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{2000})),
                      L"Existing Preferences window did not close before tree-host test.");
    }

    const auto openPreferencesWindow = [&](std::wstring_view context) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(2000ms));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto closePreferencesWindow = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)), std::format(L"Preferences window did not close during {}.", context));
        return state.failure.empty();
    };

    const auto validateCategoryTreeHost = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        const auto requireTreeSelection = [&](const HWND categoryTreeHost, const UINT expectedTitleId, std::wstring_view selectionContext) noexcept
        {
            const auto selectionState = CollectUiaSelectionPatternState(categoryTreeHost);
            state.Require(selectionState.has_value(),
                          std::format(L"Failed to collect UI Automation selection state for the Preferences category tree during {}.", selectionContext));
            if (! selectionState.has_value())
            {
                return false;
            }

            const std::wstring expectedName = LoadStringResource(nullptr, expectedTitleId);
            state.Require(selectionState->rootControlType == UIA_TreeControlTypeId,
                          std::format(L"Preferences category host should expose the UIA tree control type during {}.", selectionContext));
            state.Require(selectionState->hasSelectionPattern,
                          std::format(L"Preferences category host should expose SelectionPattern during {}.", selectionContext));
            state.Require(selectionState->selectionCount == 1u,
                          std::format(L"Preferences category host should expose exactly one selected UIA item during {}; saw {}.",
                                      selectionContext,
                                      selectionState->selectionCount));
            state.Require(selectionState->selectedControlType == UIA_TreeItemControlTypeId,
                          std::format(L"Preferences selected UIA item should be a tree item during {}.", selectionContext));
            state.Require(selectionState->selectedHasSelectionItemPattern,
                          std::format(L"Preferences selected UIA tree item should expose SelectionItemPattern during {}.", selectionContext));
            state.Require(selectionState->selectedName == expectedName,
                          std::format(L"Preferences selected UIA tree item should be '{}' during {} but was '{}'.",
                                      expectedName,
                                      selectionContext,
                                      selectionState->selectedName));
            return state.failure.empty();
        };

        RedrawWindow(prefs, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.categoryTreeUsesDxUiHost, std::format(L"Preferences category navigation is not using the shared DxUi host during {}.", context));
        state.Require(snapshot.visibleLegacyTreeViewCount == 0u,
                      std::format(L"Preferences still exposes a visible legacy SysTreeView32 category control during {}.", context));

        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        const HWND pageHost         = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, std::format(L"Preferences page host control missing during {}.", context));
        if (! categoryTreeHost || ! pageHost)
        {
            return false;
        }

        std::array<wchar_t, 32> className{};
        state.Require(GetClassNameW(categoryTreeHost, className.data(), static_cast<int>(className.size())) > 0 && _wcsicmp(className.data(), L"Static") == 0,
                      std::format(L"Preferences category host should be a plain host surface during {}.", context));
        const LONG_PTR style = GetWindowLongPtrW(categoryTreeHost, GWL_STYLE);
        state.Require((style & SS_NOTIFY) != 0, std::format(L"Preferences category host must use SS_NOTIFY during {}.", context));
        state.Require(WindowExposesUiaProvider(categoryTreeHost), std::format(L"Preferences category host should answer WM_GETOBJECT during {}.", context));
        state.Require(requireTreeSelection(categoryTreeHost, IDS_PREFS_CAT_GENERAL, std::format(L"{} on initial General selection", context)),
                      std::format(L"Preferences category tree should expose live UIA tree selection on the initial General page during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const int categoryDpi   = static_cast<int>(GetDpiForWindow(categoryTreeHost));
        const LPARAM clickPanes = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(36, categoryDpi, USER_DEFAULT_SCREEN_DPI));
        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, clickPanes);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, clickPanes);
        PumpPendingMessages();

        snapshot = {};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot),
                      std::format(L"Failed to capture Preferences snapshot after Panes click during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryPanes,
                      std::format(L"Clicking the Preferences category tree should select Panes during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PANES),
                      std::format(L"Preferences page title did not track Panes selection during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Panes page should expose exactly one shared visible page-host child during {}; saw {}.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.shellDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences shell should stay resize-failure free after Panes click during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.shellDxHostResizeFailureCount));
        state.Require(requireTreeSelection(categoryTreeHost, IDS_PREFS_CAT_PANES, std::format(L"{} after Panes click", context)),
                      std::format(L"Preferences category tree should expose live UIA tree selection for Panes during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const LPARAM clickViewers = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(60, categoryDpi, USER_DEFAULT_SCREEN_DPI));
        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, clickViewers);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, clickViewers);
        PumpPendingMessages();

        snapshot = {};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot),
                      std::format(L"Failed to capture Preferences snapshot after Viewers click during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryViewers,
                      std::format(L"Clicking the Preferences category tree should select Viewers during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                      std::format(L"Preferences page title did not track Viewers selection during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount == 1u,
                      std::format(L"Preferences Viewers page should expose exactly one shared visible page-host child during {}; saw {}.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.shellDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences shell should stay resize-failure free after Viewers click during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.shellDxHostResizeFailureCount));
        state.Require(requireTreeSelection(categoryTreeHost, IDS_PREFS_CAT_VIEWERS, std::format(L"{} after Viewers click", context)),
                      std::format(L"Preferences category tree should expose live UIA tree selection for Viewers during {}.", context));
        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial category-tree baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateCategoryTreeHost(prefs, L"initial category-tree baseline probe"), L"Initial Preferences category-tree baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial category-tree baseline probe"), L"Initial Preferences category-tree close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened category-tree baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateCategoryTreeHost(reopenedPrefs, L"reopened category-tree baseline probe"),
                  L"Reopened Preferences category-tree baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened category-tree baseline probe"),
                  L"Reopened Preferences category-tree close validation failed.");

    return state.failure.empty();
}

} // namespace

[[nodiscard]] Localization::LanguagePreference GetCurrentLocalizationPreferenceForSelfTest() noexcept
{
    Localization::LanguagePreference preference{};
    if (g_settings.ui.has_value())
    {
        const std::wstring& language = g_settings.ui.value().language;
        if (! language.empty() && ! OrdinalString::EqualsNoCase(language, L"system"))
        {
            preference.kind    = Localization::LanguagePreferenceKind::Culture;
            preference.culture = language;
        }
    }
    return preference;
}

[[nodiscard]] bool TestPreferencesDialogOpensWithFrenchSatelliteResources(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before French satellite dialog-open validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const Localization::LanguagePreference previousPreference = GetCurrentLocalizationPreferenceForSelfTest();
    const auto restoreLanguage = wil::scope_exit([&]() noexcept { static_cast<void>(Localization::ApplyLanguagePreference(previousPreference)); });

    Localization::LanguagePreference frenchPreference{};
    frenchPreference.kind    = Localization::LanguagePreferenceKind::Culture;
    frenchPreference.culture = L"fr-FR";
    const HRESULT applyHr    = Localization::ApplyLanguagePreference(frenchPreference);
    state.Require(SUCCEEDED(applyHr), L"Failed to apply fr-FR language preference before Preferences dialog-open validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open with fr-FR satellite resources active.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closePreferences = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    std::array<wchar_t, 128> caption{};
    const int captionChars             = GetWindowTextW(prefs, caption.data(), static_cast<int>(caption.size()));
    const std::wstring expectedCaption = LoadStringResource(nullptr, IDS_PREFS_CAPTION);
    state.Require(! expectedCaption.empty(), L"French Preferences caption resource should resolve.");
    state.Require(captionChars > 0 && std::wstring_view(caption.data(), static_cast<size_t>(captionChars)) == expectedCaption,
                  L"Preferences window caption should come from the active fr-FR satellite resource.");
    state.Require(GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST) != nullptr, L"French Preferences dialog should create the category host custom dialog child.");
    state.Require(GetDlgItem(prefs, IDC_PREFS_PAGE_HOST) != nullptr, L"French Preferences dialog should create the page host custom dialog child.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(DebugGetPreferencesDialogSnapshot(snapshot), L"Failed to capture Preferences snapshot with fr-FR satellite resources active.");
    state.Require(snapshot.currentCategory == kPrefCategoryGeneral, L"French Preferences dialog should initialize on the General page.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"French Preferences page title should come from the active fr-FR satellite resource.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogCategoryTreeExposesLiveUiaSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before category-tree UIA selection test.");
    }

    const auto openPreferencesWindow = [&](std::wstring_view context) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(2000ms));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto closePreferencesWindow = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)), std::format(L"Preferences window did not close during {}.", context));
        return state.failure.empty();
    };

    const auto verifySelection =
        [&](const HWND categoryTreeHost, const PrefCategory expectedCategory, const UINT expectedTitleId, std::wstring_view context) noexcept
    {
        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(snapshot.currentCategory == expectedCategory,
                      std::format(L"Preferences category-tree UIA selection test saw the wrong active category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, expectedTitleId),
                      std::format(L"Preferences page title did not match the expected category during {}.", context));

        const auto selectionState = CollectUiaSelectionPatternState(categoryTreeHost);
        state.Require(selectionState.has_value(),
                      std::format(L"Failed to collect UI Automation selection state for the Preferences category tree host during {}.", context));
        if (! selectionState.has_value())
        {
            return;
        }

        const std::wstring expectedName = LoadStringResource(nullptr, expectedTitleId);

        state.Require(
            selectionState->rootControlType == UIA_TreeControlTypeId,
            std::format(L"Preferences category host should expose UIA tree control type during {}; saw {}.", context, selectionState->rootControlType));
        state.Require(selectionState->hasSelectionPattern,
                      std::format(L"Preferences category host should expose a live UIA SelectionPattern during {}.", context));
        state.Require(
            selectionState->selectionCount == 1u,
            std::format(L"Preferences category host should expose exactly one selected UIA item during {}; saw {}.", context, selectionState->selectionCount));
        state.Require(
            selectionState->selectedControlType == UIA_TreeItemControlTypeId,
            std::format(L"Preferences selected UIA item should be a tree item during {}; saw control type {}.", context, selectionState->selectedControlType));
        state.Require(selectionState->selectedHasSelectionItemPattern,
                      std::format(L"Preferences selected UIA tree item should expose SelectionItemPattern during {}.", context));
        state.Require(
            selectionState->selectedName == expectedName,
            std::format(
                L"Preferences selected UIA tree item name should be '{}' during {} but was '{}'.", expectedName, context, selectionState->selectedName));
    };

    const auto validateSelectionCycle = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        const HWND pageHost         = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, std::format(L"Preferences page host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        verifySelection(categoryTreeHost, kPrefCategoryGeneral, IDS_PREFS_CAT_GENERAL, std::format(L"{} on initial General selection", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const int categoryDpi     = static_cast<int>(GetDpiForWindow(categoryTreeHost));
        const LPARAM clickViewers = MAKELPARAM(MulDiv(24, categoryDpi, USER_DEFAULT_SCREEN_DPI), MulDiv(60, categoryDpi, USER_DEFAULT_SCREEN_DPI));
        SendMessageW(categoryTreeHost, WM_LBUTTONDOWN, MK_LBUTTON, clickViewers);
        SendMessageW(categoryTreeHost, WM_LBUTTONUP, 0, clickViewers);
        PumpPendingMessages();

        verifySelection(categoryTreeHost, kPrefCategoryViewers, IDS_PREFS_CAT_VIEWERS, std::format(L"{} after Viewers click", context));
        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial category-tree UIA selection probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateSelectionCycle(prefs, L"initial category-tree UIA selection probe"),
                  L"Initial Preferences category-tree UIA selection validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial category-tree UIA selection probe"),
                  L"Initial Preferences category-tree UIA selection close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened category-tree UIA selection probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateSelectionCycle(reopenedPrefs, L"reopened category-tree UIA selection probe"),
                  L"Reopened Preferences category-tree UIA selection validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened category-tree UIA selection probe"),
                  L"Reopened Preferences category-tree UIA selection close validation failed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogShellUsesDxUiChrome(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before shell-host test.");
    }

    const auto openPreferencesWindow = [&](std::wstring_view context) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(2000ms));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto closePreferencesWindow = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)), std::format(L"Preferences window did not close during {}.", context));
        return state.failure.empty();
    };

    const auto validateShellChrome = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        RedrawWindow(prefs, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.shellUsesDxUiHost,
                      std::format(L"Preferences shell header/footer is not using the shared DxUi host surfaces during {}.", context));
        state.Require(snapshot.visibleLegacyShellStaticCount == 0u,
                      std::format(L"Preferences still exposes visible legacy shell static controls during {}.", context));
        state.Require(snapshot.visibleLegacyFooterButtonCount == 0u,
                      std::format(L"Preferences still exposes visible legacy footer buttons during {}.", context));
        state.Require(snapshot.shellDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences shell should stay resize-failure free during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.shellDxHostResizeFailureCount));
        state.Require(snapshot.shellHostClientWidthPx > 0 && snapshot.shellHostClientHeightPx > 0,
                      std::format(L"Preferences shell host should expose a non-empty client during {}; saw {}x{}.",
                                  context,
                                  snapshot.shellHostClientWidthPx,
                                  snapshot.shellHostClientHeightPx));
        state.Require(snapshot.shellFooterButtonsInsideHost,
                      std::format(L"Preferences DX footer buttons should stay inside the shell host during {}; host={}x{}, OK=({},{}-{},{}), "
                                  L"Cancel=({},{}-{},{}), Apply=({},{}-{},{}).",
                                  context,
                                  snapshot.shellHostClientWidthPx,
                                  snapshot.shellHostClientHeightPx,
                                  snapshot.shellOkButtonBoundsPx.left,
                                  snapshot.shellOkButtonBoundsPx.top,
                                  snapshot.shellOkButtonBoundsPx.right,
                                  snapshot.shellOkButtonBoundsPx.bottom,
                                  snapshot.shellCancelButtonBoundsPx.left,
                                  snapshot.shellCancelButtonBoundsPx.top,
                                  snapshot.shellCancelButtonBoundsPx.right,
                                  snapshot.shellCancelButtonBoundsPx.bottom,
                                  snapshot.shellApplyButtonBoundsPx.left,
                                  snapshot.shellApplyButtonBoundsPx.top,
                                  snapshot.shellApplyButtonBoundsPx.right,
                                  snapshot.shellApplyButtonBoundsPx.bottom));
        state.Require(snapshot.shellFooterButtonsInsideClip,
                      std::format(L"Preferences DX footer buttons should stay inside the shell host clipping region during {}.", context));
        state.Require(snapshot.shellOkButtonInteriorSampled,
                      std::format(L"Preferences DX footer OK button should expose a rendered sample point during {}.", context));
        state.Require(snapshot.shellOkButtonInteriorLooksPainted,
                      std::format(L"Preferences DX footer OK button rendered sample should not be raw white during {}; bgra=0x{:08X}.",
                                  context,
                                  snapshot.shellOkButtonInteriorBgra));
        state.Require(snapshot.shellFooterBackgroundSampled,
                      std::format(L"Preferences DX footer background should expose a rendered sample point during {}.", context));
        state.Require(snapshot.shellFooterBackgroundLooksThemed,
                      std::format(L"Preferences DX footer background should match the active theme during {}; bgra=0x{:08X}.",
                                  context,
                                  snapshot.shellFooterBackgroundBgra));
        state.Require(snapshot.categoryTreeTopClientPx > 0 && snapshot.categoryTreeBottomGapPx > 0,
                      std::format(L"Preferences category tree should expose client-edge geometry during {}; top={} bottomGap={} dialogBottom={}.",
                                  context,
                                  snapshot.categoryTreeTopClientPx,
                                  snapshot.categoryTreeBottomGapPx,
                                  snapshot.dialogClientBottomPx));
        state.Require(std::abs(snapshot.categoryTreeBottomGapPx - snapshot.categoryTreeTopClientPx) <= 2,
                      std::format(L"Preferences category tree should keep the same bottom gap as its top gap during {}; top={} bottomGap={}.",
                                  context,
                                  snapshot.categoryTreeTopClientPx,
                                  snapshot.categoryTreeBottomGapPx));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                      std::format(L"Preferences shell title did not initialize to the active General page during {}.", context));
        state.Require(! snapshot.pageDescription.empty(),
                      std::format(L"Preferences shell description did not initialize for the active General page during {}.", context));

        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      std::format(L"Failed to resolve the active Preferences shell host surface during {}.", context));
        const auto uiaPatternStats = (shellHost && IsWindow(shellHost) != FALSE) ? CollectVisibleUiaDescendantPatternStats(shellHost) : std::nullopt;
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences shell during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences shell should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->buttonControlCount >= 3u,
                          std::format(L"Preferences shell should expose at least the three footer DX buttons through UI Automation during {}; saw {} buttons.",
                                      context,
                                      uiaPatternStats->buttonControlCount));
            state.Require(uiaPatternStats->invokePatternCount >= 3u,
                          std::format(L"Preferences shell should expose live UI Automation InvokePattern support for the DX footer buttons during {}; saw {} "
                                      L"invoke-pattern descendants.",
                                      context,
                                      uiaPatternStats->invokePatternCount));
        }

        const auto footerButtonState =
            (shellHost && IsWindow(shellHost) != FALSE) ? CollectVisibleDescendantNamedElementState(shellHost, UIA_ButtonControlTypeId) : std::nullopt;
        state.Require(footerButtonState.has_value(), std::format(L"Preferences shell should expose a visible DX footer button during {}.", context));
        if (footerButtonState.has_value())
        {
            state.Require(! footerButtonState->name.empty(),
                          std::format(L"Preferences shell visible DX footer button should expose a stable accessible name during {}.", context));
        }
        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial shell baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateShellChrome(prefs, L"initial shell baseline probe"), L"Initial Preferences shell baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial shell baseline probe"), L"Initial Preferences shell close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened shell baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateShellChrome(reopenedPrefs, L"reopened shell baseline probe"), L"Reopened Preferences shell baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened shell baseline probe"), L"Reopened Preferences shell close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPageHostUsesDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{2000})),
                      L"Existing Preferences window did not close before page-host test.");
    }

    const auto openPreferencesWindow = [&](std::wstring_view context) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(std::chrono::milliseconds{2000}));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto validatePreferencesPageHost = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(snapshot.pageHostUsesDxUiHost, std::format(L"Preferences page host surface is not using the shared DxUi host during {}.", context));
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences should not create any dedicated page container window during {}; saw {} pane window(s).",
                                  context,
                                  snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences should not show any dedicated page container window during {}; saw {} visible pane window(s).",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(CountVisibleDescendantWindowsExposingUiaProviders(prefs) >= 2u,
                      std::format(L"Preferences should expose at least two visible WM_GETOBJECT/UIA providers during {}.", context));

        const HWND pageHost = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
        state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, std::format(L"Preferences page host control missing during {}.", context));
        if (pageHost)
        {
            std::array<wchar_t, 64> className{};
            state.Require(GetClassNameW(pageHost, className.data(), static_cast<int>(className.size())) > 0 &&
                              _wcsicmp(className.data(), L"RedSalamanderPrefsPageHost") == 0,
                          std::format(L"Preferences page host should remain the custom scroll host during {}.", context));

            const auto uiaPatternStats = CollectVisibleUiaDescendantPatternStats(pageHost);
            state.Require(uiaPatternStats.has_value(),
                          std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences page host during {}.", context));
            if (uiaPatternStats.has_value())
            {
                state.Require(uiaPatternStats->visibleElementCount > 0u,
                              std::format(L"Preferences page host should expose visible UI Automation descendants during {}.", context));
                state.Require(uiaPatternStats->buttonControlCount + uiaPatternStats->textControlCount + uiaPatternStats->togglePatternCount +
                                      uiaPatternStats->valuePatternCount >
                                  0u,
                              std::format(L"Preferences page host should expose visible DX page content through UI Automation during {}.", context));
                state.Require(uiaPatternStats->togglePatternCount > 0u,
                              std::format(L"Preferences page host should expose a visible UI Automation toggle descendant during {}.", context));
            }

            const auto toggleState = CollectVisibleDescendantTogglePatternState(pageHost);
            state.Require(toggleState.has_value(),
                          std::format(L"Preferences page host should expose a visible DX toggle on the active page during {}.", context));
            if (toggleState.has_value())
            {
                state.Require(! toggleState->name.empty(),
                              std::format(L"Preferences page host visible DX toggle should expose a stable accessible name during {}.", context));
            }
        }

        return state.failure.empty();
    };

    const auto closePreferencesWindow = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(std::chrono::milliseconds{2000})),
                      std::format(L"Preferences window did not close during {}.", context));
        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial page-host baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validatePreferencesPageHost(prefs, L"initial page-host baseline probe"), L"Initial Preferences page-host baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial page-host baseline probe"), L"Initial Preferences page-host close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened page-host baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validatePreferencesPageHost(reopenedPrefs, L"reopened page-host baseline probe"),
                  L"Reopened Preferences page-host baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened page-host baseline probe"), L"Reopened Preferences page-host close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersPageUsesDxUiChrome(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Viewers page DX chrome test.");
    }

    const auto openPreferencesWindow = [&](std::wstring_view context) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(2000ms));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto closePreferencesWindow = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)), std::format(L"Preferences window did not close during {}.", context));
        return state.failure.empty();
    };

    const auto validateViewersPageChrome = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        const HWND pageHost         = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, std::format(L"Preferences page host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, std::format(L"Failed to focus the Preferences category host during {}.", context));
        PumpPendingMessages();

        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryViewers,
                      std::format(L"Preferences navigation did not move to the Viewers category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                      std::format(L"Preferences page title did not switch to Viewers during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount <= 1u,
                      std::format(L"Preferences Viewers page should avoid child-host fanout during {}; saw {} visible child windows.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageRenderedDxHostCount <= 1u,
                      std::format(L"Preferences Viewers page should not fan out multiple DX page hosts during {}; saw {} rendered hosts.",
                                  context,
                                  snapshot.currentPageRenderedDxHostCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Viewers page should navigate without DX page-host resize failures during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.currentPageDxHostResizeFailureCount));
        state.Require(snapshot.shellDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Viewers page should navigate without shell DX host resize failures during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.shellDxHostResizeFailureCount));
        state.Require(snapshot.viewersUsesDxUiTypographyContext && snapshot.viewersUsesDxUiTypographyMetrics,
                      std::format(L"Preferences Viewers page should use shared DirectWrite typography metrics during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Viewers page is not using shared DxUi label/hint chrome during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Viewers page is not using shared DxUi text-field chrome during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Viewers page is not using shared DxUi combo/button chrome during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Viewers page is not using the shared DxUi grid surface during {}.", context));
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Viewers page should not create a pane-host child window during {}; saw {} created pane windows.",
                                  context,
                                  snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Viewers page should not expose a visible pane-host child window during {}; saw {} visible pane windows.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Viewers page still exposes visible legacy static chrome during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Viewers page still exposes a visible legacy listview during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Viewers page still exposes visible legacy edit controls during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u, std::format(L"Preferences Viewers page still exposes a visible legacy combo during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Viewers page still exposes visible legacy command buttons during {}.", context));

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      std::format(L"Failed to resolve the active Preferences page surface during {}.", context));
        const auto uiaPatternStats = (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleUiaDescendantPatternStats(activePage) : std::nullopt;
        state.Require(uiaPatternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the Preferences Viewers page during {}.", context));
        if (uiaPatternStats.has_value())
        {
            state.Require(uiaPatternStats->visibleElementCount > 0u,
                          std::format(L"Preferences Viewers page should expose visible UI Automation descendants during {}.", context));
            state.Require(uiaPatternStats->valuePatternCount > 0u,
                          std::format(L"Preferences Viewers page should expose a visible DX editable value-pattern descendant during {}.", context));
            state.Require(uiaPatternStats->invokePatternCount > 0u,
                          std::format(L"Preferences Viewers page should expose a visible DX invoke-pattern descendant during {}.", context));
        }
        const auto viewersValueState =
            (activePage && IsWindow(activePage) != FALSE) ? CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId) : std::nullopt;
        state.Require(viewersValueState.has_value(), std::format(L"Preferences Viewers page should expose a visible DX edit descendant during {}.", context));
        if (viewersValueState.has_value())
        {
            state.Require(! viewersValueState->name.empty(),
                          std::format(L"Preferences Viewers page edit descendant should expose a stable accessible name during {}.", context));
        }

        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial Viewers page baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validateViewersPageChrome(prefs, L"initial Viewers page baseline probe"), L"Initial Preferences Viewers page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial Viewers page baseline probe"), L"Initial Preferences Viewers page close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened Viewers page baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validateViewersPageChrome(reopenedPrefs, L"reopened Viewers page baseline probe"),
                  L"Reopened Preferences Viewers page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened Viewers page baseline probe"),
                  L"Reopened Preferences Viewers page close validation failed.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsPageUsesDxUiShellChrome(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins page DX shell-chrome test.");
    }

    const auto openPreferencesWindow = [&](std::wstring_view context) noexcept
    {
        SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
        const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(2000ms));
        state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, std::format(L"Preferences window did not open during {}.", context));
        return prefs;
    };

    const auto closePreferencesWindow = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        PostMessageW(prefs, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)), std::format(L"Preferences window did not close during {}.", context));
        return state.failure.empty();
    };

    const auto validatePluginsPageChrome = [&](const HWND prefs, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        const HWND pageHost         = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing during {}.", context));
        state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, std::format(L"Preferences page host control missing during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      std::format(L"Failed to select the Preferences Plugins category during {}.", context));
        PumpPendingMessages();

        PreferencesDebugSnapshot snapshot{};
        state.Require(DebugGetPreferencesDialogSnapshot(snapshot), std::format(L"Failed to capture Preferences snapshot during {}.", context));
        state.Require(snapshot.currentCategory == kPrefCategoryPlugins,
                      std::format(L"Preferences navigation did not move to the Plugins category during {}.", context));
        state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                      std::format(L"Preferences page title did not switch to Plugins during {}.", context));
        state.Require(! snapshot.pluginItemSelected, std::format(L"Plugins page test unexpectedly navigated into a plugin child item during {}.", context));
        state.Require(snapshot.pluginsTreeChildCount != 0u,
                      std::format(L"Preferences Plugins tree should expose at least one plugin child item during {}.", context));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Plugins category should keep zero visible pane-host windows during {}; saw {}.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(snapshot.pluginsPaneVisible, std::format(L"Preferences Plugins category should show the Plugins pane host during {}.", context));
        state.Require(! snapshot.generalPaneVisible, std::format(L"Preferences Plugins category should hide the General pane host during {}.", context));
        state.Require(snapshot.visibleCurrentPageChildWindowCount != 0u,
                      std::format(L"Preferences Plugins page should expose a visible active DX page surface during {}; saw {} visible child windows.",
                                  context,
                                  snapshot.visibleCurrentPageChildWindowCount));
        state.Require(snapshot.currentPageRenderedDxHostCount <= 1u,
                      std::format(L"Preferences Plugins page should not fan out multiple rendered DX page hosts during {}; saw {} rendered hosts.",
                                  context,
                                  snapshot.currentPageRenderedDxHostCount));
        state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u,
                      std::format(L"Preferences Plugins page should open without DX host resize failures during {}; saw {} failing hosts.",
                                  context,
                                  snapshot.currentPageDxHostResizeFailureCount));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Plugins page is not using shared DxUi statics for the visible list-mode shell text during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Plugins page is not using shared DxUi buttons for the visible list-mode command rows during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Plugins page is not using a shared DxUi text field for the visible search input during {}.", context));
        state.Require(true /* Phase 8: removed field */,
                      std::format(L"Preferences Plugins page is not using the shared DxUi grid for the visible main plugins list during {}.", context));
        state.Require(snapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Plugins page should not create a pane-host child window during {}; saw {} created pane windows.",
                                  context,
                                  snapshot.createdPaneWindowCount));
        state.Require(snapshot.visiblePaneWindowCount == 0u,
                      std::format(L"Preferences Plugins page should not expose a visible pane-host child window during {}; saw {} visible pane windows.",
                                  context,
                                  snapshot.visiblePaneWindowCount));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Plugins page still exposes visible legacy list-mode statics during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Plugins page still exposes visible legacy list-mode command buttons during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Plugins page still exposes a visible legacy search edit or frame during {}.", context));
        state.Require(0u /* Phase 8: removed field */ == 0u,
                      std::format(L"Preferences Plugins page still exposes a visible legacy main plugins listview during {}.", context));
        state.Require(snapshot.createdLegacyPluginsListBridgeCount == 0u,
                      std::format(L"Preferences Plugins page should not create hidden legacy listview bridges during {}; saw {} created list bridges.",
                                  context,
                                  snapshot.createdLegacyPluginsListBridgeCount));
        state.Require(snapshot.createdLegacyPluginsButtonBridgeCount == 0u,
                      std::format(L"Preferences Plugins page should not create hidden legacy command-button bridges during {}; saw {} created button bridges.",
                                  context,
                                  snapshot.createdLegacyPluginsButtonBridgeCount));
        state.Require(snapshot.createdLegacyPluginsInputBridgeCount == 0u,
                      std::format(L"Preferences Plugins page should not create a hidden legacy search bridge during {}; saw {} created input bridges.",
                                  context,
                                  snapshot.createdLegacyPluginsInputBridgeCount));
        return state.failure.empty();
    };

    const HWND prefs = openPreferencesWindow(L"initial Plugins page baseline probe");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    state.Require(validatePluginsPageChrome(prefs, L"initial Plugins page baseline probe"), L"Initial Preferences Plugins page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(prefs, L"initial Plugins page baseline probe"), L"Initial Preferences Plugins page close validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedPrefs = openPreferencesWindow(L"reopened Plugins page baseline probe");
    if (! reopenedPrefs || IsWindow(reopenedPrefs) == FALSE)
    {
        return false;
    }

    state.Require(validatePluginsPageChrome(reopenedPrefs, L"reopened Plugins page baseline probe"),
                  L"Reopened Preferences Plugins page baseline validation failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(closePreferencesWindow(reopenedPrefs, L"reopened Plugins page baseline probe"),
                  L"Reopened Preferences Plugins page close validation failed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Plugins round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins round-trip test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins round-trip test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto focusCategoryTreeHost = [&](std::wstring_view context) noexcept
    {
        if (FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)))
        {
            return true;
        }

        PreferencesDebugSnapshot focusSnapshot{};
        static_cast<void>(DebugGetPreferencesDialogSnapshot(focusSnapshot));
        state.Require(false,
                      std::format(L"Failed to focus the Preferences category host {}; nativeFocus=0x{:X}, categoryHost=0x{:X}, activePage=0x{:X}, "
                                  L"shellHost=0x{:X}, currentCategory={}, pageTitle='{}'.",
                                  context,
                                  reinterpret_cast<uintptr_t>(GetFocus()),
                                  reinterpret_cast<uintptr_t>(categoryTreeHost),
                                  reinterpret_cast<uintptr_t>(DebugGetPreferencesActivePageHandle()),
                                  reinterpret_cast<uintptr_t>(DebugGetPreferencesShellHostHandle()),
                                  static_cast<int>(focusSnapshot.currentCategory),
                                  focusSnapshot.pageTitle));
        return false;
    };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto collectActivePagePatternStats = [&]() noexcept -> std::optional<UiaDescendantPatternStats>
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences page pane during Plugins round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active Plugins page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active Plugins page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    PreferencesDebugSnapshot snapshot{};
    const auto hasPluginsPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.visiblePaneWindowCount == 0u && value.pluginsPaneVisible && ! value.generalPaneVisible && value.currentPageDxHostResizeFailureCount == 0u;
    };

    const auto hasStablePluginsPageState = [&](const PreferencesDebugSnapshot& value) noexcept
    { return hasPluginsPageSurfaceState(value) && value.pluginsMainListRowCount > 0u && value.pluginsSearchText.empty(); };

    const auto navigateToPluginsPage = [&](PreferencesDebugSnapshot& outSnapshot, std::wstring_view context) noexcept
    {
        const auto tryUnwindPluginsRoot = [&](const PreferencesDebugSnapshot& candidateSnapshot) noexcept
        {
            if (candidateSnapshot.currentCategory != kPrefCategoryPlugins ||
                (! candidateSnapshot.pluginItemSelected && ! candidateSnapshot.pluginsDetailsActive))
            {
                return false;
            }

            for (int i = 0; i < 3; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_LEFT, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_LEFT, 0);
                PumpPendingMessages();
                if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
                {
                    if (! hasStablePluginsPageState(outSnapshot))
                    {
                        state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
                        state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                                      std::format(L"Preferences Plugins page did not restore a cleared list baseline before {}.", context));
                    }
                    return state.failure.empty();
                }
            }

            return false;
        };

        if (! focusCategoryTreeHost(std::format(L"for {}", context)))
        {
            return false;
        }

        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            if (! hasStablePluginsPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
                state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                              std::format(L"Preferences Plugins page did not restore a cleared list baseline before {}.", context));
            }
            return state.failure.empty();
        }

        PreferencesDebugSnapshot candidate{};
        if (DebugGetPreferencesDialogSnapshot(candidate) && tryUnwindPluginsRoot(candidate))
        {
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      std::format(L"Failed to select the Preferences Plugins category before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();

        state.Require(waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot), std::format(L"Preferences Plugins page did not settle before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStablePluginsPageState(outSnapshot))
        {
            state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
            state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                          std::format(L"Preferences Plugins page did not restore a cleared list baseline before {}.", context));
        }
        return state.failure.empty();
    };

    state.Require(navigateToPluginsPage(snapshot, L"round-trip validation"),
                  L"Preferences Plugins page did not settle to the stabilized one-host DxUi list surface before round-trip validation.");
    state.Require(waitForSnapshot(hasStablePluginsPageState, snapshot),
                  L"Preferences Plugins page did not settle to the stabilized one-host DxUi list surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Preferences Plugins page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS_DESC),
                  L"Preferences Plugins page description did not settle before round-trip validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Plugins page still exposes visible legacy statics before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Plugins page still exposes visible legacy buttons before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Plugins page still exposes visible legacy edit chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Plugins page still exposes a visible legacy list before round-trip navigation.");
    state.Require(snapshot.createdLegacyPluginsListBridgeCount == 0u,
                  L"Preferences Plugins page should not recreate hidden legacy listview bridges before round-trip navigation.");
    state.Require(snapshot.createdLegacyPluginsButtonBridgeCount == 0u,
                  L"Preferences Plugins page should not recreate hidden legacy command-button bridges before round-trip navigation.");
    state.Require(snapshot.createdLegacyPluginsInputBridgeCount == 0u,
                  L"Preferences Plugins page should not recreate a hidden legacy search bridge before round-trip navigation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Plugins direct-host reset path should not leave a pane-host child window alive on the settled Plugins page; saw {} "
                              L"created pane hosts.",
                              snapshot.createdPaneWindowCount));
    const auto pluginsPagePatternStats = collectActivePagePatternStats();
    if (! pluginsPagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(pluginsPagePatternStats->valuePatternCount > 0u,
                  L"Preferences Plugins page should expose a visible DX editable value-pattern descendant before round-trip navigation.");
    state.Require(pluginsPagePatternStats->invokePatternCount > 0u,
                  L"Preferences Plugins page should expose a visible DX invoke-pattern descendant before round-trip navigation.");
    const auto pluginsValueState = CollectVisibleDescendantValuePatternState(DebugGetPreferencesActivePageHandle(), UIA_EditControlTypeId);
    state.Require(pluginsValueState.has_value(), L"Preferences Plugins page should expose a visible DX edit descendant before round-trip navigation.");
    if (pluginsValueState.has_value())
    {
        state.Require(! pluginsValueState->name.empty(),
                      L"Preferences Plugins page edit descendant should expose a stable accessible name before round-trip navigation.");
    }

    if (! focusCategoryTreeHost(L"before leaving Plugins for General"))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral), L"Failed to select the Preferences General category while leaving Plugins.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */

               && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences did not restore the General one-host DxUi page while leaving Plugins.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving Plugins.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving Plugins.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences should restore General without recreating a pane-host child window after leaving Plugins; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));

    snapshot = {};
    state.Require(navigateToPluginsPage(snapshot, L"returning from General to Plugins"),
                  L"Preferences Plugins page did not repaint and restore the stabilized one-host DxUi list surface after returning from General.");
    state.Require(waitForSnapshot(hasStablePluginsPageState, snapshot),
                  L"Preferences Plugins page did not repaint and restore the stabilized one-host DxUi list surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Preferences Plugins page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS_DESC),
                  L"Preferences Plugins page description did not restore after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Plugins page still exposes visible legacy statics after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Plugins page still exposes visible legacy buttons after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Plugins page still exposes visible legacy edit chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Plugins page still exposes a visible legacy list after returning from General.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Plugins direct-host reset path should restore without recreating a pane-host child window after returning from "
                              L"General; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));

    const auto restoredPluginsPatternStats = collectActivePagePatternStats();
    if (! restoredPluginsPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredPluginsPatternStats->valuePatternCount > 0u,
                  L"Preferences Plugins page should restore a visible DX editable value-pattern descendant after returning from General.");
    state.Require(restoredPluginsPatternStats->invokePatternCount > 0u,
                  L"Preferences Plugins page should restore a visible DX invoke-pattern descendant after returning from General.");
    const auto restoredPluginsValueState = CollectVisibleDescendantValuePatternState(DebugGetPreferencesActivePageHandle(), UIA_EditControlTypeId);
    state.Require(restoredPluginsValueState.has_value(), L"Preferences Plugins page should restore a visible DX edit descendant after returning from General.");
    if (restoredPluginsValueState.has_value())
    {
        state.Require(! restoredPluginsValueState->name.empty(),
                      L"Preferences Plugins page edit descendant should expose a stable accessible name after returning from General.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins theme-cycle validation.");
    }

    const auto waitForPreferencesWindow = [&]() noexcept
    { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins theme-cycle validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (prefs && IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto navigateToPluginsPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto hasPluginsPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
                   value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        };

        const auto hasStablePluginsPageState = [&](const PreferencesDebugSnapshot& value) noexcept
        { return hasPluginsPageSurfaceState(value) && value.pluginsMainListRowCount > 1u && value.pluginsSearchText.empty(); };

        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins theme-cycle validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Plugins theme-cycle validation.");
        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            if (! hasStablePluginsPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesPluginsSearchText(L""),
                              L"Failed to clear the Plugins search field before establishing the theme-cycle baseline.");
                state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                              L"Preferences Plugins page did not restore a cleared multi-row baseline before theme-cycle validation.");
            }
            return state.failure.empty();
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category before theme-cycle validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();
        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            if (! hasStablePluginsPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesPluginsSearchText(L""),
                              L"Failed to clear the Plugins search field before establishing the theme-cycle baseline.");
                state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                              L"Preferences Plugins page did not restore a cleared multi-row baseline before theme-cycle validation.");
            }
            return state.failure.empty();
        }

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i <= 7; ++i)
        {
            PreferencesDebugSnapshot candidate{};
            if (DebugGetPreferencesDialogSnapshot(candidate) && hasPluginsPageSurfaceState(candidate))
            {
                outSnapshot = std::move(candidate);
                break;
            }

            if (i != 7)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }
        }

        state.Require(waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot), L"Preferences Plugins page did not settle before theme-cycle validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStablePluginsPageState(outSnapshot))
        {
            state.Require(DebugSetPreferencesPluginsSearchText(L""), L"Failed to clear the Plugins search field before establishing the theme-cycle baseline.");
            state.Require(waitForSnapshot(hasStablePluginsPageState, outSnapshot),
                          L"Preferences Plugins page did not restore a cleared multi-row baseline before theme-cycle validation.");
        }
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToPluginsPage(snapshot))
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.pluginsMainListRowCount;

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-plugins-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.pluginsMainListRowCount == baselineRowCount && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    size_t loadableRowIndex = 0u;
    state.Require(DebugFindPreferencesPluginsLoadableMainListRow(loadableRowIndex),
                  L"Preferences Plugins theme-cycle validation could not find a loadable DX main-list row.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(loadableRowIndex),
                  L"Failed to select the loadable Preferences Plugins DX row before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsSelectedPluginIdText.empty() &&
               ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the selected loadable DX plugin row before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselinePluginId = snapshot.pluginsSelectedPluginIdText;

    state.Require(DebugFocusPreferencesPluginsSearchField(), L"Preferences Plugins DX search field did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::SearchField &&
               value.pluginItemSelected && value.pluginsSelectedPluginIdText == baselinePluginId && value.pluginsMainListRowCount == baselineRowCount &&
               ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Plugins search field did not settle as the focused control before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Plugins page surface during theme-cycle validation.");
        return activePage;
    };

    const auto initialValueState = CollectVisibleDescendantValuePatternState(getActivePage(), UIA_EditControlTypeId);
    state.Require(initialValueState.has_value(), L"Preferences Plugins page should expose a visible DX edit descendant before theme-cycle validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Plugins page visible DX edit descendant should remain editable before theme-cycle validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Plugins page visible DX edit descendant should expose a stable accessible name before theme-cycle validation.");
    if (initialValueState->isReadOnly || initialValueState->name.empty())
    {
        return false;
    }

    const std::wstring baselineEditName  = initialValueState->name;
    const std::wstring baselineEditValue = initialValueState->value;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.pluginsMainListRowCount == baselineRowCount && value.pluginItemSelected &&
                   value.pluginsSelectedPluginIdText == baselinePluginId && ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Plugins page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesPluginsSearchField(),
                      std::format(L"Preferences Plugins search field did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::SearchField &&
                   value.pluginsMainListRowCount == baselineRowCount && ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u &&
                   value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences Plugins focus target did not return to the search field after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences Plugins UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences Plugins page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Preferences Plugins page should keep a visible DX edit descendant after the {} theme update.", label));
            state.Require(stats->invokePatternCount > 0u || stats->buttonControlCount > 0u,
                          std::format(L"Preferences Plugins page should keep visible DX command actions after the {} theme update.", label));
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, baselineEditName);
        state.Require(valueState.has_value(), std::format(L"Preferences Plugins visible DX search edit disappeared after the {} theme update.", label));
        if (valueState.has_value())
        {
            state.Require(valueState->value == baselineEditValue,
                          std::format(L"Preferences Plugins search value changed unexpectedly after the {} theme update.", label));
            state.Require(! valueState->isReadOnly, std::format(L"Preferences Plugins search edit became read-only after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Preferences Plugins rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences Plugins high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-plugins-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-plugins-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-plugins-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-plugins-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersRoundTripRestoresDxUiSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before Viewers round-trip test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers round-trip test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Viewers round-trip test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto collectActivePagePatternStats = [&]() noexcept -> std::optional<UiaDescendantPatternStats>
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences page pane during Viewers round-trip validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(), L"Failed to collect live UI Automation pattern statistics for the active Viewers page subtree.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u, L"Active Viewers page subtree should expose visible UI Automation descendants.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    PreferencesDebugSnapshot snapshot{};
    const auto hasViewersPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryViewers && value.currentPageDxHostResizeFailureCount == 0u; };
    const auto hasStableViewersPageState = [hasViewersPageSurfaceState](const PreferencesDebugSnapshot& value) noexcept
    { return hasViewersPageSurfaceState(value) && value.viewersListRowCount >= 2u && value.viewersSearchText.empty(); };
    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot, std::wstring_view context) noexcept
    {
        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, std::format(L"Failed to focus the Preferences category host before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();

        if (waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            if (! hasStableViewersPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesViewersSearchText(L""), std::format(L"Failed to clear the Viewers search field before {}.", context));
                state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                              std::format(L"Preferences Viewers page did not restore a cleared list baseline before {}.", context));
            }
            return state.failure.empty();
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                      std::format(L"Failed to select the Preferences Viewers category before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();

        if (waitForSnapshot(hasViewersPageSurfaceState, outSnapshot))
        {
            if (! hasStableViewersPageState(outSnapshot))
            {
                state.Require(DebugSetPreferencesViewersSearchText(L""), std::format(L"Failed to clear the Viewers search field before {}.", context));
                state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                              std::format(L"Preferences Viewers page did not restore a cleared list baseline before {}.", context));
            }
            return state.failure.empty();
        }

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(hasViewersPageSurfaceState, outSnapshot), std::format(L"Preferences Viewers page did not settle before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStableViewersPageState(outSnapshot))
        {
            state.Require(DebugSetPreferencesViewersSearchText(L""), std::format(L"Failed to clear the Viewers search field before {}.", context));
            state.Require(waitForSnapshot(hasStableViewersPageState, outSnapshot),
                          std::format(L"Preferences Viewers page did not restore a cleared list baseline before {}.", context));
        }
        return state.failure.empty();
    };
    const auto navigateToGeneralPage = [&](PreferencesDebugSnapshot& outSnapshot, std::wstring_view context) noexcept
    {
        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, std::format(L"Failed to focus the Preferences category host before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral),
                      std::format(L"Failed to select the Preferences General category before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();

        if (waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot))
        {
            return true;
        }

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        return waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryGeneral && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot);
    };

    state.Require(navigateToViewersPage(snapshot, L"round-trip validation"),
                  L"Preferences Viewers page did not settle to the stabilized one-host DxUi surface before round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                  L"Preferences Viewers page title did not settle before round-trip validation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS_DESC),
                  L"Preferences Viewers page description did not settle before round-trip validation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes visible legacy statics before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes a visible legacy listview before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes visible legacy edit chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes visible legacy combo chrome before round-trip navigation.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes visible legacy buttons before round-trip navigation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Viewers direct-host reset path should not leave a pane-host child window alive on the settled Viewers page; saw {} "
                              L"created pane hosts.",
                              snapshot.createdPaneWindowCount));
    const auto viewersPagePatternStats = collectActivePagePatternStats();
    if (! viewersPagePatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(viewersPagePatternStats->editControlCount + viewersPagePatternStats->comboBoxControlCount > 0u,
                  L"Preferences Viewers page should expose visible edit or combo descendants before round-trip navigation.");

    snapshot = {};
    state.Require(navigateToGeneralPage(snapshot, L"leaving Viewers for General"),
                  L"Preferences did not restore the General one-host DxUi page while leaving Viewers.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL),
                  L"Preferences page title did not switch back to General while leaving Viewers.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_GENERAL_DESC),
                  L"Preferences page description did not switch back to General while leaving Viewers.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Preferences should restore General without recreating a pane-host child window after leaving Viewers; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));

    snapshot = {};
    state.Require(navigateToViewersPage(snapshot, L"returning from General to Viewers"),
                  L"Preferences Viewers page did not repaint and restore the stabilized one-host DxUi surface after returning from General.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                  L"Preferences Viewers page title did not restore after returning from General.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS_DESC),
                  L"Preferences Viewers page description did not restore after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes visible legacy statics after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes a visible legacy listview after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes visible legacy edit chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes visible legacy combo chrome after returning from General.");
    state.Require(0u /* Phase 8: removed field */ == 0u, L"Preferences Viewers page still exposes visible legacy buttons after returning from General.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Viewers direct-host reset path should restore without recreating a pane-host child window after returning from "
                              L"General; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));

    const auto restoredViewersPatternStats = collectActivePagePatternStats();
    if (! restoredViewersPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredViewersPatternStats->editControlCount + restoredViewersPatternStats->comboBoxControlCount > 0u,
                  L"Preferences Viewers page should restore visible edit or combo descendants after returning from General.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers theme-cycle validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    const auto waitForPreferencesWindow = [&]() noexcept
    { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers theme-cycle validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (prefs && IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto navigateToViewersPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE, L"Preferences category host control missing for Viewers theme-cycle validation.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Viewers theme-cycle validation.");
        PumpPendingMessages();

        SendMessageW(treeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(treeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(treeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot(
                          [](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          outSnapshot),
                      L"Preferences Viewers page did not settle before theme-cycle validation.");
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(snapshot))
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.viewersListRowCount;

    const AppTheme initialTheme = ResolveAppTheme(ThemeMode::Dark, L"preferences-viewers-selftest-theme-cycle-initial");
    UpdatePreferencesWindowsTheme(initialTheme);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.themeDark && ! value.themeHighContrast && ! value.themeRainbow &&
               value.viewersListRowCount == baselineRowCount && value.currentPageDxHostResizeFailureCount == 0u &&
               value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Viewers page did not settle to the baseline dark theme-cycle state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Preferences Viewers DX row before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.viewersListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the selected DX mapping row before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselineSelectedExtension = snapshot.viewersSelectedExtensionText;

    state.Require(DebugFocusPreferencesViewersSearchField(), L"Preferences Viewers search field did not accept focus before theme-cycle validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::SearchField &&
               value.viewersSelectedExtensionText == baselineSelectedExtension && value.viewersListRowCount == baselineRowCount &&
               value.currentPageDxHostResizeFailureCount == 0u && value.visibleCurrentPageChildWindowCount <= 1u;
    },
                      snapshot),
                  L"Preferences Viewers focus target did not settle to the search field before theme-cycle validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto getActivePage = [&]() noexcept -> HWND
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences Viewers page surface during theme-cycle validation.");
        return activePage;
    };

    const auto initialValueState = CollectVisibleDescendantValuePatternState(getActivePage(), UIA_EditControlTypeId);
    state.Require(initialValueState.has_value(), L"Preferences Viewers page should expose a visible DX edit descendant before theme-cycle validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Viewers page visible DX edit descendant should remain editable before theme-cycle validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Viewers page visible DX edit descendant should expose a stable accessible name before theme-cycle validation.");
    if (initialValueState->isReadOnly || initialValueState->name.empty())
    {
        return false;
    }

    const std::wstring baselineEditValue          = initialValueState->value;
    const std::wstring baselineEditAccessibleName = initialValueState->name;

    const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
    {
        UpdatePreferencesWindowsTheme(theme);
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast &&
                   value.themeRainbow == theme.menu.rainbowMode && value.viewersSelectedExtensionText == baselineSelectedExtension &&
                   value.viewersListRowCount == baselineRowCount && value.currentPageDxHostResizeFailureCount == 0u &&
                   value.visibleCurrentPageChildWindowCount <= 1u;
        },
                          snapshot),
                      std::format(L"Preferences Viewers page did not settle after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        state.Require(DebugFocusPreferencesViewersSearchField(),
                      std::format(L"Preferences Viewers search field did not reacquire focus after the {} theme update.", label));
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            return value.currentCategory == kPrefCategoryViewers && value.viewersFocusTarget == PreferencesViewersDebugFocusTarget::SearchField &&
                   value.viewersSelectedExtensionText == baselineSelectedExtension && value.viewersListRowCount == baselineRowCount &&
                   value.currentPageDxHostResizeFailureCount == 0u;
        },
                          snapshot),
                      std::format(L"Preferences Viewers focus target did not return to the search field after the {} theme update.", label));
        if (! state.failure.empty())
        {
            return;
        }

        const HWND activePage = getActivePage();
        const auto stats      = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(stats.has_value(), std::format(L"Failed to collect Preferences Viewers UIA pattern stats after the {} theme update.", label));
        if (stats.has_value())
        {
            state.Require(stats->visibleElementCount > 0u,
                          std::format(L"Preferences Viewers page should keep visible UIA descendants after the {} theme update.", label));
            state.Require(stats->valuePatternCount > 0u,
                          std::format(L"Preferences Viewers page should keep a visible value-pattern descendant after the {} theme update.", label));
        }

        const auto valueState = CollectVisibleDescendantValuePatternState(activePage, UIA_EditControlTypeId);
        state.Require(valueState.has_value(), std::format(L"Preferences Viewers visible DX edit descendant disappeared after the {} theme update.", label));
        if (valueState.has_value())
        {
            state.Require(valueState->value == baselineEditValue,
                          std::format(L"Preferences Viewers visible DX edit value changed unexpectedly after the {} theme update.", label));
            state.Require(valueState->name == baselineEditAccessibleName,
                          std::format(L"Preferences Viewers visible DX edit accessible name changed unexpectedly after the {} theme update.", label));
            state.Require(! valueState->isReadOnly, std::format(L"Preferences Viewers visible DX edit became read-only after the {} theme update.", label));
        }

        state.Require(snapshot.themeRainbow == expectRainbow,
                      std::format(L"Preferences Viewers rainbow-theme flag mismatch after the {} theme update.", label));
        state.Require(snapshot.themeHighContrast == expectHighContrast,
                      std::format(L"Preferences Viewers high-contrast flag mismatch after the {} theme update.", label));
    };

    requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"preferences-viewers-selftest-theme-cycle-dark"), false, false);
    requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"preferences-viewers-selftest-theme-cycle-light"), false, false);
    requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"preferences-viewers-selftest-theme-cycle-rainbow"), true, false);
    requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"preferences-viewers-selftest-theme-cycle-high-contrast"), false, true);

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersSearchRoundTripPreservesRetainedState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers search round-trip test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers search round-trip test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Viewers search round-trip test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Viewers search round-trip test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                  L"Failed to select the Preferences Viewers category for Viewers search round-trip test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryViewers && true /* Phase 8: removed field */ && value.viewersListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Viewers page did not settle before retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount           = snapshot.viewersListRowCount;
    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    state.Require(DebugSetPreferencesViewersSearchText(kSearchText), L"Failed to set the retained Viewers search text through the shared DX page host.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == L"__codex_no_match__" && value.viewersListRowCount == 0u; },
                                  snapshot),
                  L"Preferences Viewers page did not retain the filtered zero-row search state after applying the DX search text.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral),
                  L"Failed to select the Preferences General category during retained-search round-trip validation.");
    PumpPendingMessages();
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */; },
                                  snapshot),
                  L"Preferences did not leave Viewers for General during retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                  L"Failed to reselect the Preferences Viewers category during retained-search round-trip validation.");
    PumpPendingMessages();

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == kSearchText && value.viewersListRowCount == 0u &&
               value.createdPaneWindowCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not restore the retained search/filter state after leaving and re-entering the page.");
    state.Require(baselineRowCount > snapshot.viewersListRowCount,
                  L"Preferences Viewers retained-search test did not reduce the visible row set from its baseline.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersSearchActionUpdatesDxSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers deferred-search test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers deferred-search test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Viewers deferred-search test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Viewers deferred-search test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers), L"Failed to select the Preferences Viewers category for deferred-search test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Viewers page did not settle before deferred-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    state.Require(DebugSetPreferencesViewersSearchText(kSearchText), L"Failed to set the Viewers search text through the shared DX page host.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == L"__codex_no_match__" && value.viewersListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers deferred search action did not settle to the filtered DxUi state.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersLiveSearchDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers live search interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    SelfTest::AppendSelfTestTrace(L"Viewers live-search: seeding viewer associations");
    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SelfTest::AppendSelfTestTrace(L"Viewers live-search: posting Preferences open command");
    state.Require(PostMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0) != FALSE,
                  L"Failed to post Preferences open command for Viewers live search interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }
    prefs = waitForPreferencesWindow();
    SelfTest::AppendSelfTestTrace(L"Viewers live-search: Preferences window wait returned");
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers live search interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto waitForEditValue = [&](std::wstring_view expectedName, std::wstring_view expectedValue) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Viewers live search interaction validation.");
        return shellHost;
    };

    const auto navigateToViewersPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND targetCategoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(targetCategoryTreeHost != nullptr && IsWindow(targetCategoryTreeHost) != FALSE,
                      L"Preferences category host control missing while navigating to the Viewers page.");
        if (! targetCategoryTreeHost || IsWindow(targetCategoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(targetCategoryTreeHost) == targetCategoryTreeHost,
                      L"Failed to focus the Preferences category host while navigating to the Viewers page.");
        if (! state.failure.empty())
        {
            return false;
        }

        SendMessageW(targetCategoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(targetCategoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(targetCategoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(targetCategoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        return waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    SelfTest::AppendSelfTestTrace(L"Viewers live-search: navigating to Viewers page");
    if (! navigateToViewersPage(prefs, snapshot))
    {
        return false;
    }
    SelfTest::AppendSelfTestTrace(L"Viewers live-search: Viewers page ready");

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_VIEWERS),
                  L"Preferences Viewers page title did not settle before live search interaction validation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Viewers page should not recreate a pane host before live search interaction validation; saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));
    state.Require(
        snapshot.visiblePaneWindowCount == 0u,
        std::format(L"Preferences Viewers page should not expose a visible pane host before live search interaction validation; saw {} visible pane hosts.",
                    snapshot.visiblePaneWindowCount));

    const size_t baselineRowCount = snapshot.viewersListRowCount;
    const HWND activePage         = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers page surface during live search interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }
    state.Require(DebugFocusPreferencesViewersSearchField(),
                  L"Preferences Viewers search field did not accept focus before reset interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesViewersSearchField(),
                  L"Preferences Viewers search field did not accept focus before live search interaction validation.");
    const std::wstring searchEditName = LoadStringResource(nullptr, IDS_PREFS_COMMON_SEARCH);
    state.Require(! searchEditName.empty(), L"Preferences Viewers search caption should resolve before live search interaction validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto initialValueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, searchEditName);
    state.Require(initialValueState.has_value(),
                  L"Preferences Viewers page should expose a visible DX edit descendant during live search interaction validation.");
    if (! initialValueState.has_value())
    {
        return false;
    }

    state.Require(! initialValueState->isReadOnly,
                  L"Preferences Viewers page visible DX edit descendant should remain editable during live search interaction validation.");
    state.Require(! initialValueState->name.empty(),
                  L"Preferences Viewers page edit descendant should expose a stable accessible name during live search interaction validation.");
    if (initialValueState->isReadOnly || initialValueState->name.empty())
    {
        return false;
    }

    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    const std::wstring editName             = searchEditName;
    const std::wstring initialEditValue     = initialValueState->value;
    state.Require(SetVisibleDescendantValueByNameWithMessagePump(
                      activePage, UIA_EditControlTypeId, editName, kSearchText, L"Preferences Viewers initial live search SetValue"),
                  L"Preferences Viewers page visible DX search edit did not accept live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(editName, kSearchText),
                  L"Preferences Viewers page visible DX search edit did not settle to the edited value after live UIA mutation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == L"__codex_no_match__" && value.viewersListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not settle to the filtered DX state after live UIA search mutation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring shellCancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! shellCancelButtonText.empty(),
                  L"Preferences shell Cancel caption should resolve for live UIA InvokePattern validation during Viewers live search interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDescendantByNameWithMessagePump(
            getShellHost(), UIA_ButtonControlTypeId, shellCancelButtonText, L"Preferences Viewers shell Cancel"),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Viewers live search discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Viewers live search "
                  L"discard validation.");
    prefs = nullptr;

    SelfTest::AppendSelfTestTrace(L"Viewers live-search: posting Preferences reopen command");
    state.Require(PostMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0) != FALSE,
                  L"Failed to post Preferences reopen command for Viewers live search discard validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    prefs = waitForPreferencesWindow();
    SelfTest::AppendSelfTestTrace(L"Viewers live-search: Preferences reopen wait returned");
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Viewers live search discard validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToViewersPage(prefs, snapshot))
    {
        return false;
    }
    state.Require(DebugFocusPreferencesViewersSearchField(),
                  L"Preferences Viewers search field did not accept focus after shell Cancel reopened the page.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Viewers page visible DX search edit did not discard the pending search value after shell Cancel reopened the page.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == initialEditValue && value.viewersListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not restore its baseline row set after shell Cancel discarded the pending search filter.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Viewers page surface during live search revalidation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }
    state.Require(DebugFocusPreferencesViewersSearchField(),
                  L"Preferences Viewers search field did not accept focus after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetVisibleDescendantValueByNameWithMessagePump(
                      reopenedActivePage, UIA_EditControlTypeId, editName, kSearchText, L"Preferences Viewers reopened live search SetValue"),
                  L"Preferences Viewers page visible DX search edit did not accept live UIA ValuePattern mutation after shell Cancel reopen.");
    state.Require(waitForEditValue(editName, kSearchText),
                  L"Preferences Viewers page visible DX search edit did not settle to the edited value after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == L"__codex_no_match__" && value.viewersListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not settle to the filtered DX state after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetVisibleDescendantValueByNameWithMessagePump(DebugGetPreferencesActivePageHandle(),
                                                                 UIA_EditControlTypeId,
                                                                 editName,
                                                                 initialEditValue,
                                                                 L"Preferences Viewers restore live search SetValue"),
                  L"Preferences Viewers page visible DX search edit did not accept restoration through live UIA ValuePattern.");
    state.Require(waitForEditValue(editName, initialEditValue),
                  L"Preferences Viewers page visible DX search edit did not restore its original value after live UIA mutation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == initialEditValue && value.viewersListRowCount == baselineRowCount &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not restore its baseline row set after live UIA search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    return VerifyPreferencesGridSelectionPattern(
        prefs, state, L"Viewers", snapshot.viewersListRowCount, [](const size_t rowIndex) noexcept { return DebugSelectPreferencesViewersListRow(rowIndex); });
}

[[nodiscard]] bool TestPreferencesDialogViewersRemoveLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers remove interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers remove interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Viewers remove interaction validation.");
        return shellHost;
    };

    const auto navigateToViewersPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Viewers remove interaction test.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(FocusWindowAndWait(categoryTreeHost, SelfTest::Scale(1000ms)),
                      L"Failed to focus the Preferences category host for Viewers remove interaction test.");
        state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                      L"Failed to select the Preferences Viewers category for remove interaction validation.");
        PumpPendingMessages();

        const bool settled = waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot);
        state.Require(settled,
                      std::format(L"Preferences Viewers page did not settle before remove interaction validation; category={}, rows={}, "
                                  L"selected='{}', childWindows={}, renderedDxHosts={}, resizeFailures={}, pageTitle='{}'.",
                                  static_cast<int>(outSnapshot.currentCategory),
                                  outSnapshot.viewersListRowCount,
                                  outSnapshot.viewersSelectedExtensionText,
                                  outSnapshot.visibleCurrentPageChildWindowCount,
                                  outSnapshot.currentPageRenderedDxHostCount,
                                  outSnapshot.currentPageDxHostResizeFailureCount,
                                  outSnapshot.pageTitle));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(outSnapshot.createdPaneWindowCount == 0u,
                      std::format(L"Preferences Viewers page should not recreate a pane host before remove interaction validation; saw {} created pane hosts.",
                                  outSnapshot.createdPaneWindowCount));
        state.Require(
            outSnapshot.visiblePaneWindowCount == 0u,
            std::format(L"Preferences Viewers page should not expose a visible pane host before remove interaction validation; saw {} visible pane hosts.",
                        outSnapshot.visiblePaneWindowCount));

        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(prefs, snapshot))
    {
        return false;
    }

    const std::wstring removedExtension = L".selftest-viewers-001";
    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row for live Remove interaction validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == removedExtension &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the DX-selected extension before invoking the visible Remove action.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers page surface during remove interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring removeButtonText = LoadStringResource(nullptr, IDS_PREFS_VIEWERS_BUTTON_REMOVE);
    state.Require(! removeButtonText.empty(), L"Preferences Viewers Remove button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Viewers remove interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, removeButtonText),
                  L"Failed to invoke the visible Preferences Viewers Remove button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 2u && value.viewersSelectedExtensionText.empty() &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers visible DX Remove action did not remove the selected mapping and restore the shared page state.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Viewers remove discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Viewers remove discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Viewers remove commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToViewersPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to reselect the first Viewers DX row for live remove commit validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == removedExtension &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not restore the removed mapping after shell Cancel discarded the pending delete.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Viewers page surface during remove commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, removeButtonText),
                  L"Failed to invoke the visible Preferences Viewers Remove button through live UIA interaction after shell Cancel reopen.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 2u && value.viewersSelectedExtensionText.empty() &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers visible DX Remove action did not remove the selected mapping after shell Cancel reopen.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersResetLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers reset interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    const size_t defaultRowCount = TestVisibleViewerAssociationRowCount(Common::Settings::DefaultViewerFileActionsSettings());

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers reset interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto formatViewersResetSnapshot = [&](const PreferencesDebugSnapshot& value, const size_t expectedRowCount) noexcept
    {
        return std::format(L"currentCategory={}, viewersListRowCount={} (expected {}), viewersSelectedExtensionText='{}', viewersSearchText='{}', "
                           L"createdPaneWindowCount={}, visiblePaneWindowCount={}, currentPageDxHostResizeFailureCount={}",
                           static_cast<int>(value.currentCategory),
                           value.viewersListRowCount,
                           expectedRowCount,
                           value.viewersSelectedExtensionText,
                           value.viewersSearchText,
                           value.createdPaneWindowCount,
                           value.visiblePaneWindowCount,
                           value.currentPageDxHostResizeFailureCount);
    };

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Viewers reset interaction validation.");
        return shellHost;
    };

    const auto navigateToViewersPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing for Viewers reset interaction test during {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      std::format(L"Failed to focus the Preferences category host for Viewers reset interaction test during {}.", context));
        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        const bool settled = waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept {
            return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.currentPageDxHostResizeFailureCount == 0u;
        }, outSnapshot);
        if (! settled)
        {
            PreferencesDebugSnapshot actualSnapshot = outSnapshot;
            static_cast<void>(DebugGetPreferencesDialogSnapshot(actualSnapshot));
            state.Require(false,
                          std::format(L"Preferences Viewers page did not settle before reset interaction validation during {}; {}.",
                                      context,
                                      formatViewersResetSnapshot(actualSnapshot, 3u)));
        }
        if (! state.failure.empty())
        {
            return false;
        }

        return true;
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(prefs, snapshot, L"initial open"))
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers page surface during reset interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const auto ensureAssociationsTab = [&](HWND pageHost, std::wstring_view context) noexcept
    {
        size_t selectedTabIndex = 0u;
        if (DebugGetPreferencesViewersSelectedTabIndex(selectedTabIndex) && selectedTabIndex == 1u)
        {
            return true;
        }

        RECT associationsTabRect{};
        state.Require(DebugGetPreferencesViewersTabClientRect(1u, associationsTabRect),
                      std::format(L"Failed to capture the Preferences Viewers Associations tab rect during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        const LONG clickX = associationsTabRect.left + ((associationsTabRect.right - associationsTabRect.left) / 2);
        const LONG clickY = associationsTabRect.top + ((associationsTabRect.bottom - associationsTabRect.top) / 2);
        SendMouseClickToResolvedPointWindow(pageHost, MAKELPARAM(clickX, clickY));

        PreferencesDebugSnapshot tabSnapshot{};
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            selectedTabIndex = 0u;
            return value.currentCategory == kPrefCategoryViewers && DebugGetPreferencesViewersSelectedTabIndex(selectedTabIndex) &&
                   selectedTabIndex == 1u && value.currentPageDxHostResizeFailureCount == 0u;
        },
                          tabSnapshot),
                      std::format(L"Preferences Viewers Associations tab did not settle before reset interaction during {}.", context));
        return state.failure.empty();
    };

    if (! ensureAssociationsTab(activePage, L"initial open"))
    {
        return false;
    }

    const std::wstring resetButtonText = LoadStringResource(nullptr, IDS_PREFS_FILE_ACTION_BUTTON_RESET_DEFAULTS);
    state.Require(! resetButtonText.empty(), L"Preferences Viewers Reset Defaults button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Viewers reset interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(activePage, UIA_ButtonControlTypeId, resetButtonText),
                  L"Failed to invoke the visible Preferences Viewers Reset Defaults button through live UIA interaction.");
    const bool firstResetRestoredDefaults = waitForSnapshot(
        [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == defaultRowCount && value.viewersSelectedExtensionText.empty() &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
        snapshot);
    if (! firstResetRestoredDefaults)
    {
        PreferencesDebugSnapshot actualSnapshot = snapshot;
        static_cast<void>(DebugGetPreferencesDialogSnapshot(actualSnapshot));
        state.Require(false,
                      std::format(L"Preferences Viewers visible DX Reset Defaults action did not restore the default mappings and shared page state; {}.",
                                  formatViewersResetSnapshot(actualSnapshot, defaultRowCount)));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
                  L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Viewers reset discard validation.");
    state.Require(
        WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
        L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Viewers reset discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Viewers reset commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToViewersPage(prefs, snapshot, L"cancel reopen"))
    {
        return false;
    }

    const std::wstring baselineSelectedExtension = L".selftest-viewers-001";
    const bool cancelRestoredBaselineMappings    = waitForSnapshot(
        [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u &&
               value.viewersSelectedExtensionText == baselineSelectedExtension && value.currentPageDxHostResizeFailureCount == 0u;
    },
        snapshot);
    if (! cancelRestoredBaselineMappings)
    {
        PreferencesDebugSnapshot actualSnapshot = snapshot;
        static_cast<void>(DebugGetPreferencesDialogSnapshot(actualSnapshot));
        state.Require(
            false,
            std::format(
                L"Preferences Viewers page did not restore the baseline mappings and first-row selection after shell Cancel discarded the pending reset; {}.",
                formatViewersResetSnapshot(actualSnapshot, 3u)));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Viewers page surface during reset commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }
    state.Require(DebugFocusPreferencesViewersSearchField(),
                  L"Preferences Viewers search field did not accept focus before reset commit validation.");
    if (! state.failure.empty())
    {
        return false;
    }
    if (! ensureAssociationsTab(reopenedActivePage, L"cancel reopen"))
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, resetButtonText),
                  L"Failed to invoke the visible Preferences Viewers Reset Defaults button through live UIA interaction after shell Cancel reopen.");
    const bool secondResetRestoredDefaults = waitForSnapshot(
        [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == defaultRowCount && value.viewersSelectedExtensionText.empty() &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
        snapshot);
    if (! secondResetRestoredDefaults)
    {
        PreferencesDebugSnapshot actualSnapshot = snapshot;
        static_cast<void>(DebugGetPreferencesDialogSnapshot(actualSnapshot));
        state.Require(false,
                      std::format(L"Preferences Viewers visible DX Reset Defaults action did not restore the default mappings after shell Cancel reopen; {}.",
                                  formatViewersResetSnapshot(actualSnapshot, defaultRowCount)));
    }
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kRemovedSearchText = L".selftest-viewers-001";
    state.Require(DebugSetPreferencesViewersSearchText(kRemovedSearchText), L"Failed to set the Viewers search text while validating the post-reset DX state.");
    const bool resetRemovedSelfTestMapping = waitForSnapshot(
        [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSearchText == L".selftest-viewers-001" && value.viewersListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
        snapshot);
    if (! resetRemovedSelfTestMapping)
    {
        PreferencesDebugSnapshot actualSnapshot = snapshot;
        static_cast<void>(DebugGetPreferencesDialogSnapshot(actualSnapshot));
        state.Require(false,
                      std::format(L"Preferences Viewers reset validation did not remove the self-test extension mapping from the visible DX list; {}.",
                                  formatViewersResetSnapshot(actualSnapshot, 0u)));
    }
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersAddUpdateLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers Add / Update interaction test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    auto waitForPreferencesWindow = [&]() noexcept { return WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms)); };

    auto prefs = HWND{};
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers Add / Update interaction test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto getShellHost = [&]() noexcept -> HWND
    {
        const HWND shellHost = DebugGetPreferencesShellHostHandle();
        state.Require(shellHost != nullptr && IsWindow(shellHost) != FALSE,
                      L"Failed to resolve the active Preferences shell host surface during Viewers Add / Update interaction validation.");
        return shellHost;
    };

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto waitForEditValue = [&](std::wstring_view expectedName, std::wstring_view expectedValue) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            const HWND activePage = DebugGetPreferencesActivePageHandle();
            if (activePage && IsWindow(activePage) != FALSE)
            {
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    return true;
                }
            }

            std::this_thread::sleep_for(20ms);
        }

        const HWND activePage = DebugGetPreferencesActivePageHandle();
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return false;
        }

        const auto valueState = CollectVisibleDescendantValuePatternStateByName(activePage, UIA_EditControlTypeId, expectedName);
        return valueState.has_value() && valueState->value == expectedValue;
    };

    const auto navigateToViewersPage = [&](HWND targetPrefs, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND treeHost = GetDlgItem(targetPrefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(treeHost != nullptr && IsWindow(treeHost) != FALSE,
                      L"Preferences category host control missing for Viewers Add / Update interaction test.");
        if (! treeHost || IsWindow(treeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(treeHost) == treeHost, L"Failed to focus the Preferences category host for Viewers Add / Update interaction test.");
        SendMessageW(treeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(treeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i < 2; ++i)
        {
            SendMessageW(treeHost, WM_KEYDOWN, VK_DOWN, 0);
            SendMessageW(treeHost, WM_KEYUP, VK_DOWN, 0);
            PumpPendingMessages();
        }

        state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
        { return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u && value.currentPageDxHostResizeFailureCount == 0u; },
                                      outSnapshot),
                      L"Preferences Viewers page did not settle before Add / Update interaction validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(
            outSnapshot.createdPaneWindowCount == 0u,
            std::format(L"Preferences Viewers page should not recreate a pane host before Add / Update interaction validation; saw {} created pane hosts.",
                        outSnapshot.createdPaneWindowCount));
        state.Require(
            outSnapshot.visiblePaneWindowCount == 0u,
            std::format(
                L"Preferences Viewers page should not expose a visible pane host before Add / Update interaction validation; saw {} visible pane hosts.",
                outSnapshot.visiblePaneWindowCount));
        return state.failure.empty();
    };

    PreferencesDebugSnapshot snapshot{};
    if (! navigateToViewersPage(prefs, snapshot))
    {
        return false;
    }

    constexpr std::wstring_view kOriginalExtension = L".selftest-viewers-001";
    constexpr std::wstring_view kUpdatedExtension  = L".selftest-viewers-updated";
    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to select the first Viewers DX row for live Add / Update interaction validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not retain the DX-selected extension before editing the visible Add / Update form.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Viewers page surface during Add / Update interaction validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring matchValueLabel = LoadStringResource(nullptr, IDS_PREFS_FILE_ACTION_LABEL_MATCH_VALUE);
    state.Require(! matchValueLabel.empty(), L"Preferences Viewers match-value label should resolve for live UIA ValuePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    state.Require(! cancelButtonText.empty(),
                  L"Preferences Cancel caption should resolve for live UIA InvokePattern validation during Viewers Add / Update interaction.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(SetVisibleDescendantValueByName(activePage, UIA_EditControlTypeId, matchValueLabel, kUpdatedExtension),
                  L"Preferences Viewers visible DX match-value edit did not accept live UIA ValuePattern mutation.");
    state.Require(waitForEditValue(matchValueLabel, kUpdatedExtension),
                  L"Preferences Viewers visible DX match-value edit did not settle to the edited value after live UIA mutation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers edited DX extension value should not commit the mapping before the Add / Update action runs.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(
        InvokeVisibleDescendantByName(getShellHost(), UIA_ButtonControlTypeId, cancelButtonText),
        L"Preferences shell visible DX Cancel action did not expose live UIA InvokePattern interaction during Viewers Add / Update discard validation.");
    state.Require(WaitForWindowClosed(prefs, SelfTest::Scale(3000ms)),
                  L"Preferences dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action during Viewers Add / Update "
                  L"discard validation.");
    prefs = nullptr;

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    prefs = waitForPreferencesWindow();
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not reopen for Viewers Add / Update commit validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    if (! navigateToViewersPage(prefs, snapshot))
    {
        return false;
    }

    state.Require(DebugSelectPreferencesViewersListRow(0u), L"Failed to reselect the first Viewers DX row for live Add / Update commit validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers page did not restore the DX-selected extension after shell Cancel discarded the uncommitted edit.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND reopenedActivePage = DebugGetPreferencesActivePageHandle();
    state.Require(reopenedActivePage != nullptr && IsWindow(reopenedActivePage) != FALSE,
                  L"Failed to resolve the reopened Preferences Viewers page surface during Add / Update commit validation.");
    if (! reopenedActivePage || IsWindow(reopenedActivePage) == FALSE)
    {
        return false;
    }

    state.Require(waitForEditValue(matchValueLabel, kOriginalExtension),
                  L"Preferences Viewers shell Cancel path did not restore the visible DX match-value edit to its baseline value.");
    state.Require(SetVisibleDescendantValueByName(reopenedActivePage, UIA_EditControlTypeId, matchValueLabel, kUpdatedExtension),
                  L"Preferences Viewers visible DX match-value edit did not accept live UIA ValuePattern mutation during Add / Update commit validation.");
    state.Require(waitForEditValue(matchValueLabel, kUpdatedExtension),
                  L"Preferences Viewers visible DX match-value edit did not settle to the edited value during Add / Update commit validation.");
    state.Require(
        waitForSnapshot(
            [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == L".selftest-viewers-001" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Viewers edited DX extension value should not commit the mapping before the Add / Update action runs after shell Cancel reopen.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring saveAssociationButtonText = LoadStringResource(nullptr, IDS_PREFS_FILE_ACTION_BUTTON_SAVE_ASSOCIATION);
    state.Require(! saveAssociationButtonText.empty(),
                  L"Preferences Viewers Save Association button caption should resolve for live UIA InvokePattern validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(InvokeVisibleDescendantByName(reopenedActivePage, UIA_ButtonControlTypeId, saveAssociationButtonText),
                  L"Failed to invoke the visible Preferences Viewers Save Association button through live UIA interaction.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryViewers && value.viewersListRowCount == 3u &&
               value.viewersSelectedExtensionText == L".selftest-viewers-updated" && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Viewers visible DX Add / Update action did not commit the edited mapping and preserve the shared page state.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogViewersSelectionSurvivesLegacyListClear(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Viewers retained-selection test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    TestSetViewerAssociationRows({{L".selftest-viewers-001", L"builtin/viewer-text"},
                                  {L".selftest-viewers-002", L"builtin/viewer-pe"},
                                  {L".selftest-viewers-003", L"builtin/viewer-text"}});

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Viewers retained-selection test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Viewers retained-selection test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Viewers retained-selection test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers), L"Failed to select the Preferences Viewers category for retained-selection test.");
    PumpPendingMessages();

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryViewers && true /* Phase 8: removed field */ && value.viewersListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Viewers page did not settle before retained-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t targetRowIndex = snapshot.viewersListRowCount > 1u ? 1u : 0u;
    state.Require(DebugSelectPreferencesViewersListRow(targetRowIndex), L"Failed to select a Viewers DX row before retained-selection validation.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryViewers && ! value.viewersSelectedExtensionText.empty(); },
                                  snapshot),
                  L"Preferences Viewers page did not expose a retained selected extension after DX row selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring retainedExtension = snapshot.viewersSelectedExtensionText;
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == retainedExtension; },
                                  snapshot),
                  L"Preferences Viewers retained selected extension should remain stable.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryGeneral),
                  L"Failed to select the Preferences General category during retained-selection validation.");
    PumpPendingMessages();
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */; },
                                  snapshot),
                  L"Preferences did not leave Viewers for General during retained-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryViewers),
                  L"Failed to reselect the Preferences Viewers category during retained-selection validation.");
    PumpPendingMessages();

    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryViewers && value.viewersSelectedExtensionText == retainedExtension && value.createdPaneWindowCount == 0u; },
                                  snapshot),
                  L"Preferences Viewers page did not restore the retained selected extension after leaving and re-entering the page.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginTreeSelectionKeepsSingleVisiblePane(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)), L"Existing Preferences window did not close before plugin-tree switch test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for plugin-tree switch test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    const HWND pageHost         = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, L"Preferences page host control missing for Editors/Mouse navigation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        do
        {
            PumpPendingMessages();
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(SelfTest::Scale(25ms));
        } while (std::chrono::steady_clock::now() < deadline);

        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto collectActivePagePatternStats = [&]() noexcept -> std::optional<UiaDescendantPatternStats>
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Failed to resolve the active Preferences page pane during plugin-tree UIA validation.");
        if (! activePage || IsWindow(activePage) == FALSE || ! state.failure.empty())
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(),
                      L"Failed to collect live UI Automation pattern statistics for the active Preferences page subtree during plugin-tree validation.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u,
                      L"Active Preferences page subtree should expose visible UI Automation descendants during plugin-tree validation.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for plugin-tree switch test.");
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins), L"Failed to select the Preferences Plugins category for plugin-tree switch test.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for plugin-tree switch test.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsExpanded && value.visiblePaneWindowCount == 0u &&
               value.pluginsPaneVisible && ! value.generalPaneVisible && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins category did not settle to a single visible Plugins pane before child navigation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pluginsTreeChildCount != 0u, L"Preferences Plugins tree should expose at least one plugin child item for tree-switch validation.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Preferences Plugins category should still show the Plugins page title before child navigation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS_DESC),
                  L"Preferences Plugins category should still show the Plugins page description before child navigation.");
    state.Require(snapshot.shellDxHostResizeFailureCount == 0u, L"Preferences shell reported DX resize failures before tree child navigation.");
    state.Require(true /* Phase 8: removed field */, L"Preferences Plugins category should expose the shared DxUi list surface before child navigation.");
    state.Require(snapshot.createdPaneWindowCount == 0u,
                  std::format(L"Preferences Plugins root should not keep a pane-host child window alive before child navigation on the direct-host reset path; "
                              L"saw {} created pane hosts.",
                              snapshot.createdPaneWindowCount));

    const auto pluginsRootPatternStats = collectActivePagePatternStats();
    if (! pluginsRootPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(pluginsRootPatternStats->editControlCount + pluginsRootPatternStats->comboBoxControlCount > 0u,
                  L"Preferences Plugins root page should expose the visible DX search/filter input descendants before child navigation.");
    state.Require(pluginsRootPatternStats->visibleElementCount > 0u,
                  L"Preferences Plugins root page should expose visible UI Automation descendants before child navigation.");

    state.Require(DebugSelectPreferencesPluginsTreeChild(0u), L"Failed to select the first Preferences plugin child for plugin-tree switch test.");
    PumpPendingMessages();

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && value.visiblePaneWindowCount == 0u && value.pluginsPaneVisible &&
               ! value.generalPaneVisible && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Selecting a plugin child from the Preferences category tree did not settle to a single visible Plugins details pane.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! snapshot.pageTitle.empty(), L"Selecting a plugin child from the Preferences category tree should expose a non-empty shell title.");
    state.Require(snapshot.shellDxHostResizeFailureCount == 0u, L"Preferences shell reported DX resize failures after tree child navigation.");
    state.Require(true /* Phase 8: removed field */,
                  L"Selecting a plugin child from the Preferences category tree should expose the shared DxUi config statics surface.");
    state.Require(true /* Phase 8: removed field */,
                  L"Selecting a plugin child from the Preferences category tree should expose the shared DxUi config input surface.");
    state.Require(0u /* Phase 8: removed field */ == 0u,
                  L"Selecting a plugin child from the Preferences category tree should not expose visible legacy Plugins detail statics.");
    state.Require(
        snapshot.createdLegacyPluginsConfigStaticBridgeCount == 0u,
        L"Selecting a plugin child from the Preferences category tree should not recreate hidden legacy Plugins config statics on the active DX path.");
    state.Require(
        snapshot.createdLegacyPluginsConfigInputBridgeCount == 0u,
        L"Selecting a plugin child from the Preferences category tree should not recreate hidden legacy Plugins config inputs on the active DX path.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(
            L"Preferences Plugins child navigation should not recreate a pane-host child window on the direct-host reset path; saw {} created pane hosts.",
            snapshot.createdPaneWindowCount));

    const auto pluginDetailsPatternStats = collectActivePagePatternStats();
    if (! pluginDetailsPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(pluginDetailsPatternStats->visibleElementCount > 0u,
                  L"Preferences Plugins details page should expose visible UI Automation descendants after child navigation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to restore the Preferences Plugins root from the plugin child during plugin-tree switch test.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after returning to the Plugins root during plugin-tree switch test.");

    snapshot = {};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginsDetailsActive && value.visiblePaneWindowCount == 0u &&
               value.pluginsPaneVisible && ! value.generalPaneVisible && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Navigating back to the Plugins root from a plugin child did not restore a single visible Plugins list pane.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Navigating back to the Plugins root from a plugin child did not restore the Plugins page title.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS_DESC),
                  L"Navigating back to the Plugins root from a plugin child did not restore the Plugins page description.");
    state.Require(snapshot.shellDxHostResizeFailureCount == 0u, L"Preferences shell reported DX resize failures after returning to the Plugins root.");
    state.Require(true /* Phase 8: removed field */, L"Navigating back to the Plugins root from a plugin child should restore the shared DxUi list surface.");
    state.Require(
        snapshot.createdPaneWindowCount == 0u,
        std::format(L"Returning to the Preferences Plugins root should restore without recreating a pane-host child window; saw {} created pane hosts.",
                    snapshot.createdPaneWindowCount));

    const auto restoredPluginsRootPatternStats = collectActivePagePatternStats();
    if (! restoredPluginsRootPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredPluginsRootPatternStats->editControlCount + restoredPluginsRootPatternStats->comboBoxControlCount > 0u,
                  L"Returning to the Preferences Plugins root should restore the DX search/filter input descendants.");
    state.Require(restoredPluginsRootPatternStats->visibleElementCount > 0u,
                  L"Returning to the Preferences Plugins root should restore visible UI Automation descendants.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginTreeLeftRightNavigationStaysOnDxTreePath(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before plugin-tree left/right navigation test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for plugin-tree left/right navigation test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE, L"Preferences category host control missing.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        do
        {
            PumpPendingMessages();
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(SelfTest::Scale(25ms));
        } while (std::chrono::steady_clock::now() < deadline);

        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto collectActivePagePatternStats = [&]() -> std::optional<UiaDescendantPatternStats>
    {
        const HWND activePage = DebugGetPreferencesActivePageHandle();
        state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                      L"Preferences should expose an active page handle during plugin-tree left/right navigation validation.");
        if (! activePage || IsWindow(activePage) == FALSE)
        {
            return std::nullopt;
        }

        const auto pagePatternStats = CollectVisibleUiaDescendantPatternStats(activePage);
        state.Require(pagePatternStats.has_value(),
                      L"Failed to collect live UI Automation pattern statistics for the active Preferences page subtree during plugin-tree left/right "
                      L"navigation validation.");
        if (! pagePatternStats.has_value() || ! state.failure.empty())
        {
            return std::nullopt;
        }

        state.Require(pagePatternStats->visibleElementCount > 0u,
                      L"Active Preferences page subtree should expose visible UI Automation descendants during plugin-tree left/right navigation validation.");
        if (! state.failure.empty())
        {
            return std::nullopt;
        }

        return pagePatternStats;
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for plugin-tree left/right navigation test.");
    PumpPendingMessages();

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for plugin-tree left/right navigation test.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for plugin-tree left/right navigation test.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsExpanded && value.pluginsPaneVisible &&
               ! value.generalPaneVisible && value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u &&
               value.categoryTreeHasSelectedItem;
    },
                      snapshot),
                  L"Preferences Plugins root did not settle before plugin-tree left/right navigation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pluginsTreeChildCount != 0u, L"Preferences Plugins tree should expose at least one plugin child item for left/right navigation.");
    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Preferences Plugins root should expose the Plugins page title before left/right navigation.");
    state.Require(snapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS_DESC),
                  L"Preferences Plugins root should expose the Plugins page description before left/right navigation.");
    const size_t pluginsRootIndex = snapshot.categoryTreeSelectedVisibleIndex;

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_RIGHT, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_RIGHT, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot childSnapshot{};
    state.Require(waitForSnapshot(
                      [pluginsRootIndex](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && value.pluginsExpanded && value.pluginsPaneVisible &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u && value.categoryTreeHasSelectedItem &&
               value.categoryTreeSelectedVisibleIndex == pluginsRootIndex + 1u;
    },
                      childSnapshot),
                  L"Preferences plugin-tree VK_RIGHT did not move from the Plugins root to the first child on the shared DX tree path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(! childSnapshot.pageTitle.empty(), L"Selecting the first Plugins child through VK_RIGHT should expose a non-empty shell title.");
    state.Require(true /* Phase 8: pluginsPageUsesDxUiConfigStatics removed */ && true /* Phase 8: pluginsPageUsesDxUiConfigInputs removed */,
                  L"Selecting the first Plugins child through VK_RIGHT should expose the shared DxUi config detail surface.");
    const auto childPatternStats = collectActivePagePatternStats();
    if (! childPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_LEFT, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_LEFT, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot restoredRootSnapshot{};
    state.Require(waitForSnapshot(
                      [pluginsRootIndex](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginsDetailsActive && value.pluginsExpanded && value.pluginsPaneVisible &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u && value.categoryTreeHasSelectedItem &&
               value.categoryTreeSelectedVisibleIndex == pluginsRootIndex;
    },
                      restoredRootSnapshot),
                  L"Preferences plugin-tree VK_LEFT from the first child did not return to the expanded Plugins root on the shared DX tree path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredRootSnapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Returning to the Plugins root through VK_LEFT did not restore the Plugins page title.");
    state.Require(restoredRootSnapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS_DESC),
                  L"Returning to the Plugins root through VK_LEFT did not restore the Plugins page description.");
    state.Require(true /* Phase 8: pluginsPageUsesDxUiList removed */,
                  L"Returning to the Plugins root through VK_LEFT should restore the shared DxUi list surface.");
    const auto restoredRootPatternStats = collectActivePagePatternStats();
    if (! restoredRootPatternStats.has_value() || ! state.failure.empty())
    {
        return false;
    }

    state.Require(restoredRootPatternStats->editControlCount + restoredRootPatternStats->comboBoxControlCount > 0u,
                  L"Returning to the Preferences Plugins root from a child should restore the DX search/filter input descendants.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_LEFT, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_LEFT, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot collapsedSnapshot{};
    state.Require(waitForSnapshot(
                      [pluginsRootIndex](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginsDetailsActive && ! value.pluginsExpanded && value.pluginsPaneVisible &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u && value.categoryTreeHasSelectedItem &&
               value.categoryTreeSelectedVisibleIndex == pluginsRootIndex;
    },
                      collapsedSnapshot),
                  L"Preferences plugin-tree VK_LEFT on the Plugins root did not collapse the plugin subtree on the shared DX tree path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(collapsedSnapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Collapsing the Plugins root through VK_LEFT should keep the Plugins page title.");
    state.Require(collapsedSnapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS_DESC),
                  L"Collapsing the Plugins root through VK_LEFT should keep the Plugins page description.");
    state.Require(true /* Phase 8: pluginsPageUsesDxUiList removed */,
                  L"Collapsing the Plugins root through VK_LEFT should keep the shared DxUi list surface visible.");

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_RIGHT, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_RIGHT, 0);
    PumpPendingMessages();

    PreferencesDebugSnapshot expandedAgainSnapshot{};
    state.Require(waitForSnapshot(
                      [pluginsRootIndex](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginsDetailsActive && value.pluginsExpanded && value.pluginsPaneVisible &&
               value.visiblePaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u && value.categoryTreeHasSelectedItem &&
               value.categoryTreeSelectedVisibleIndex == pluginsRootIndex;
    },
                      expandedAgainSnapshot),
                  L"Preferences plugin-tree VK_RIGHT on the collapsed Plugins root did not re-expand the plugin subtree on the shared DX tree path.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(expandedAgainSnapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Re-expanding the Plugins root through VK_RIGHT should restore the Plugins page title.");
    state.Require(expandedAgainSnapshot.pageDescription == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS_DESC),
                  L"Re-expanding the Plugins root through VK_RIGHT should restore the Plugins page description.");
    state.Require(true /* Phase 8: pluginsPageUsesDxUiList removed */,
                  L"Re-expanding the Plugins root through VK_RIGHT should keep the shared DxUi list surface visible.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsSearchRoundTripPreservesRetainedState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins search round-trip test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins search round-trip test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins search round-trip test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }
            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Plugins search round-trip test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins search round-trip test.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins search round-trip test.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsMainListRowCount > 0u; },
                                  snapshot),
                  L"Preferences Plugins page did not settle before retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount           = snapshot.pluginsMainListRowCount;
    constexpr std::wstring_view kSearchText = L"__codex_no_match__";
    state.Require(DebugSetPreferencesPluginsSearchText(kSearchText), L"Failed to set the retained Plugins search text through the shared DX page host.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText == L"__codex_no_match__" && value.pluginsMainListRowCount == 0u; },
                                  snapshot),
                  L"Preferences Plugins page did not retain the filtered zero-row search state after applying the DX search text.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
    PumpPendingMessages();
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */; },
                                  snapshot),
                  L"Preferences did not leave Plugins for General during retained-search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to reselect the Preferences Plugins category after leaving it during Plugins search round-trip test.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after reselecting Plugins during Plugins search round-trip test.");

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsSearchText == kSearchText &&
               value.pluginsMainListRowCount == 0u && value.createdPaneWindowCount == 0u && value.createdLegacyPluginsButtonBridgeCount == 0u &&
               value.createdLegacyPluginsInputBridgeCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not restore the retained search/filter state after leaving and re-entering the page.");
    state.Require(baselineRowCount > snapshot.pluginsMainListRowCount,
                  L"Preferences Plugins retained-search test did not reduce the visible row set from its baseline.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsCustomPathsSelectionRoundTripPreservesRetainedState(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins custom-path retained-selection round-trip test.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.plugins.customPluginPaths.clear();
    g_settings.plugins.customPluginPaths.emplace_back(LR"(C:\SelfTest\Plugins\CustomPath001)");
    g_settings.plugins.customPluginPaths.emplace_back(LR"(C:\SelfTest\Plugins\CustomPath002)");
    g_settings.plugins.customPluginPaths.emplace_back(LR"(C:\SelfTest\Plugins\CustomPath003)");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins custom-path retained-selection round-trip test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins custom-path retained-selection round-trip test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host for Plugins custom-path retained-selection round-trip test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins custom-path retained-selection round-trip test.");
    PumpPendingMessages();
    state.Require(
        SetFocus(categoryTreeHost) == categoryTreeHost,
        L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins custom-path retained-selection round-trip test.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && true /* Phase 8: removed field */
               && value.pluginsCustomPathsListRowCount >= 3u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle before custom-path retained-selection round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::wstring_view kExpectedPath = LR"(C:\SelfTest\Plugins\CustomPath003)";
    state.Require(DebugSelectPreferencesPluginsCustomPathsListRow(2u),
                  L"Failed to select the third Plugins custom-path row through the shared DX list surface.");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && value.pluginsSelectedCustomPathText == LR"(C:\SelfTest\Plugins\CustomPath003)"; },
                                  snapshot),
                  L"Preferences Plugins page did not retain the DX-selected custom-path text after selecting a custom-path row.");
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
    SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
    PumpPendingMessages();
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryGeneral && true /* Phase 8: removed field */ && true /* Phase 8: removed field */; },
                                  snapshot),
                  L"Preferences did not leave Plugins for General during custom-path retained-selection round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to reselect the Preferences Plugins category after leaving it during custom-path retained-selection round-trip validation.");
    PumpPendingMessages();
    state.Require(
        SetFocus(categoryTreeHost) == categoryTreeHost,
        L"Failed to restore focus to the Preferences category host after reselecting Plugins during custom-path retained-selection round-trip validation.");

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsCustomPathsListRowCount == 3u &&
               value.pluginsSelectedCustomPathText == kExpectedPath && value.createdPaneWindowCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not restore the retained custom-path selection after leaving and re-entering the page.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsMainSelectionSurvivesLegacyListClear(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins retained main-selection validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins retained main-selection validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins retained main-selection validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host for Plugins retained main-selection validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins retained main-selection validation.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins retained main-selection validation.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && true /* Phase 8: removed field */
               && value.pluginsMainListRowCount > 1u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle before retained main-selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(1u), L"Failed to select the second Plugins main DX row for retained main-selection validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsSelectedPluginIdText.empty() &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins details page did not retain the DX-selected plugin id after main-row selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPluginId = snapshot.pluginsSelectedPluginIdText;
    state.Require(snapshot.createdLegacyPluginsListBridgeCount == 0u,
                  L"Preferences Plugins details page should not recreate hidden legacy listview bridges during retained main-selection validation.");
    state.Require(snapshot.createdLegacyPluginsButtonBridgeCount == 0u,
                  L"Preferences Plugins details page should not recreate hidden legacy button bridges during retained main-selection validation.");
    state.Require(snapshot.createdLegacyPluginsInputBridgeCount == 0u,
                  L"Preferences Plugins details page should not recreate hidden legacy input bridges during retained main-selection validation.");
    state.Require(snapshot.createdLegacyPluginsConfigInputBridgeCount == 0u,
                  L"Preferences Plugins details page should not recreate hidden legacy config inputs during retained main-selection validation.");

    RECT client{};
    if (GetClientRect(prefs, &client))
    {
        SendMessageW(prefs, WM_SIZE, SIZE_RESTORED, MAKELPARAM(std::max(0l, client.right - client.left), std::max(0l, client.bottom - client.top)));
    }
    else
    {
        SendMessageW(prefs, WM_SIZE, SIZE_RESTORED, 0);
    }

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && value.pluginsSelectedPluginIdText == selectedPluginId &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins details page did not preserve the retained selected plugin after clearing the hidden legacy list selection.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsCustomPathsSelectionSurvivesLegacyListClear(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins retained custom-path selection validation.");
    }

    const Common::Settings::Settings baselineSettings = g_settings;
    const auto restoreSettings                        = wil::scope_exit([&]() noexcept { g_settings = baselineSettings; });

    g_settings.plugins.customPluginPaths.clear();
    g_settings.plugins.customPluginPaths.emplace_back(LR"(C:\SelfTest\Plugins\CustomPath001)");
    g_settings.plugins.customPluginPaths.emplace_back(LR"(C:\SelfTest\Plugins\CustomPath002)");
    g_settings.plugins.customPluginPaths.emplace_back(LR"(C:\SelfTest\Plugins\CustomPath003)");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins retained custom-path selection validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins retained custom-path selection validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host for Plugins retained custom-path selection validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins retained custom-path selection validation.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins retained custom-path selection validation.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && true /* Phase 8: removed field */
               && value.pluginsCustomPathsListRowCount > 1u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle before retained custom-path selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsCustomPathsListRow(1u),
                  L"Failed to select the second Plugins custom-path DX row for retained custom-path selection validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsSelectedCustomPathText == LR"(C:\SelfTest\Plugins\CustomPath002)" &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the DX-selected custom-path after custom-path row selection.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedCustomPath = snapshot.pluginsSelectedCustomPathText;

    RECT client{};
    if (GetClientRect(prefs, &client))
    {
        SendMessageW(prefs, WM_SIZE, SIZE_RESTORED, MAKELPARAM(std::max(0l, client.right - client.left), std::max(0l, client.bottom - client.top)));
    }
    else
    {
        SendMessageW(prefs, WM_SIZE, SIZE_RESTORED, 0);
    }

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.pluginsSelectedCustomPathText == selectedCustomPath &&
               value.createdPaneWindowCount == 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not preserve the retained selected custom-path after clearing the hidden legacy custom-paths list selection.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsMainCheckboxSurvivesLegacyRowStateChange(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins retained checkbox validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins retained checkbox validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins retained checkbox validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Plugins retained checkbox validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins retained checkbox validation.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins retained checkbox validation.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && true /* Phase 8: removed field */
               && value.pluginsMainListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle before retained checkbox validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    size_t rowIndex      = 0u;
    bool originalEnabled = false;
    state.Require(DebugFindPreferencesPluginsToggleableMainListRow(rowIndex, originalEnabled),
                  L"Failed to locate a toggleable Plugins DX main-list row for retained checkbox validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const bool toggledEnabled = ! originalEnabled;
    state.Require(DebugTogglePreferencesPluginsMainListCheckbox(rowIndex),
                  L"Failed to toggle a Plugins DX main-list checkbox during retained checkbox validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        bool currentEnabled = false;
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.currentPageDxHostResizeFailureCount == 0u &&
               DebugGetPreferencesPluginsMainListRowEnabled(rowIndex, currentEnabled) && currentEnabled == toggledEnabled;
    },
                      snapshot),
                  L"Preferences Plugins DX main-list checkbox did not update retained enabled state after toggle.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT client{};
    if (GetClientRect(prefs, &client))
    {
        SendMessageW(prefs, WM_SIZE, SIZE_RESTORED, MAKELPARAM(std::max(0l, client.right - client.left), std::max(0l, client.bottom - client.top)));
    }
    else
    {
        SendMessageW(prefs, WM_SIZE, SIZE_RESTORED, 0);
    }

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        bool currentEnabled = false;
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && value.createdPaneWindowCount == 0u &&
               value.currentPageDxHostResizeFailureCount == 0u && DebugGetPreferencesPluginsMainListRowEnabled(rowIndex, currentEnabled) &&
               currentEnabled == toggledEnabled;
    },
                      snapshot),
                  L"Preferences Plugins DX main-list checkbox state did not survive hidden legacy row-state change plus refresh.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsMainCheckboxSpaceTogglesLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins live checkbox-space validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins live checkbox-space validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins live checkbox-space validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Plugins live checkbox-space validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins live checkbox-space validation.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins live checkbox-space validation.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && value.pluginsMainListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u; },
                                  snapshot),
                  L"Preferences Plugins page did not settle before live checkbox-space validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    size_t rowIndex      = 0u;
    bool originalEnabled = false;
    state.Require(DebugFindPreferencesPluginsToggleableMainListRow(rowIndex, originalEnabled),
                  L"Failed to locate a toggleable Plugins DX main-list row for live checkbox-space validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(rowIndex), L"Failed to select a Plugins DX main-list row before live checkbox-space validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        bool currentEnabled = false;
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.currentPageDxHostResizeFailureCount == 0u &&
               DebugGetPreferencesPluginsMainListRowEnabled(rowIndex, currentEnabled) && currentEnabled == originalEnabled;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the selected DX main-list row before live checkbox-space validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring baselinePluginId = snapshot.pluginsSelectedPluginIdText;
    const uint64_t baselineRenderCount  = snapshot.pluginsMainListRenderCount;
    const uint64_t baselineResizeCount  = snapshot.pluginsMainListResizeCount;
    const size_t baselineVisibleRows    = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineVisibleColumns = snapshot.pluginsMainListVisibleColumnCount;

    state.Require(DebugFocusPreferencesPluginsMainList(), L"Failed to focus the Preferences Plugins DX main list before live checkbox-space validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::MainList &&
               value.pluginsSelectedPluginIdText == baselinePluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX main list did not take focus before live checkbox-space validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX page host before live checkbox-space validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto sendSpace = [&]() noexcept
    {
        SendMessageW(activePage, WM_KEYDOWN, VK_SPACE, 0);
        SendMessageW(activePage, WM_KEYUP, VK_SPACE, 0);
        PumpPendingMessages();
    };

    const auto requireEnabledState = [&](const bool expectedEnabled, std::wstring_view label) noexcept
    {
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            bool currentEnabled = false;
            return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::MainList &&
                   value.pluginsSelectedPluginIdText == baselinePluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
                   value.currentPageDxHostResizeFailureCount == 0u && value.pluginsMainListResizeCount == baselineResizeCount &&
                   value.pluginsMainListVisibleRowCount == baselineVisibleRows && value.pluginsMainListVisibleColumnCount == baselineVisibleColumns &&
                   DebugGetPreferencesPluginsMainListRowEnabled(rowIndex, currentEnabled) && currentEnabled == expectedEnabled &&
                   value.pluginsMainListRenderCount >= baselineRenderCount;
        },
                          snapshot),
                      std::format(L"Preferences Plugins DX main-list checkbox did not reach the expected enabled state after {}.", label));
    };

    sendSpace();
    requireEnabledState(! originalEnabled, L"first Space toggle");
    if (! state.failure.empty())
    {
        return false;
    }

    sendSpace();
    requireEnabledState(originalEnabled, L"second Space toggle");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsMainCheckboxClickTogglesLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    SelfTest::AppendSelfTestTrace(L"Plugins click toggle: begin");

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins live checkbox-click validation.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    SelfTest::AppendSelfTestTrace(std::format(L"Plugins click toggle: prefs hwnd={:#x}", reinterpret_cast<uintptr_t>(prefs)));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins live checkbox-click validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    const HWND pageHost         = GetDlgItem(prefs, IDC_PREFS_PAGE_HOST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins live checkbox-click validation.");
    state.Require(pageHost != nullptr && IsWindow(pageHost) != FALSE, L"Preferences page host control missing for Plugins live checkbox-click validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE || ! pageHost || IsWindow(pageHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Plugins live checkbox-click validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins live checkbox-click validation.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins live checkbox-click validation.");

    PreferencesDebugSnapshot snapshot{};
    SelfTest::AppendSelfTestTrace(L"Plugins click toggle: waiting for plugins page settle");
    state.Require(waitForSnapshot([](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && value.pluginsMainListRowCount > 0u && value.currentPageDxHostResizeFailureCount == 0u; },
                                  snapshot),
                  L"Preferences Plugins page did not settle before live checkbox-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    size_t rowIndex      = 0u;
    bool originalEnabled = false;
    SelfTest::AppendSelfTestTrace(L"Plugins click toggle: locating toggleable row");
    state.Require(DebugFindPreferencesPluginsToggleableMainListRow(rowIndex, originalEnabled),
                  L"Failed to locate a toggleable Plugins DX main-list row for live checkbox-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSelfTestTrace(std::format(L"Plugins click toggle: selected row={} originalEnabled={}", rowIndex, originalEnabled));
    state.Require(DebugSelectPreferencesPluginsMainListRow(rowIndex), L"Failed to select a Plugins DX main-list row before live checkbox-click validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        bool currentEnabled = false;
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.currentPageDxHostResizeFailureCount == 0u &&
               DebugGetPreferencesPluginsMainListRowEnabled(rowIndex, currentEnabled) && currentEnabled == originalEnabled;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the selected DX main-list row before live checkbox-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    SelfTest::AppendSelfTestTrace(L"Plugins click toggle: capturing checkbox rect");
    RECT checkboxRect{};
    state.Require(DebugGetPreferencesPluginsMainListCheckboxClientRect(rowIndex, checkboxRect),
                  L"Failed to capture a visible Plugins DX checkbox cell rect for live checkbox-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const int clickX        = checkboxRect.left + ((checkboxRect.right - checkboxRect.left) / 2);
    const int clickY        = checkboxRect.top + ((checkboxRect.bottom - checkboxRect.top) / 2);
    const LPARAM clickPoint = MAKELPARAM(clickX, clickY);

    const std::wstring baselinePluginId = snapshot.pluginsSelectedPluginIdText;
    const uint64_t baselineRenderCount  = snapshot.pluginsMainListRenderCount;
    const uint64_t baselineResizeCount  = snapshot.pluginsMainListResizeCount;
    const size_t baselineVisibleRows    = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineVisibleColumns = snapshot.pluginsMainListVisibleColumnCount;

    state.Require(DebugFocusPreferencesPluginsMainList(), L"Failed to focus the Preferences Plugins DX main list before live checkbox-click validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::MainList &&
               value.pluginsSelectedPluginIdText == baselinePluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX main list did not take focus before live checkbox-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX page host before live checkbox-click validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto clickCheckbox = [&]() noexcept
    {
        SelfTest::AppendSelfTestTrace(
            std::format(L"Plugins click toggle: clicking point=({}, {}) pageHost={:#x}", clickX, clickY, reinterpret_cast<uintptr_t>(activePage)));
        SendMessageW(activePage, WM_MOUSEMOVE, 0, clickPoint);
        SendMessageW(activePage, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
        SendMessageW(activePage, WM_LBUTTONUP, 0, clickPoint);
        PumpPendingMessages();
    };

    const auto requireEnabledState = [&](const bool expectedEnabled, std::wstring_view label) noexcept
    {
        state.Require(waitForSnapshot(
                          [&](const PreferencesDebugSnapshot& value) noexcept
        {
            bool currentEnabled = false;
            return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::MainList &&
                   value.pluginsSelectedPluginIdText == baselinePluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
                   value.currentPageDxHostResizeFailureCount == 0u && value.pluginsMainListResizeCount == baselineResizeCount &&
                   value.pluginsMainListVisibleRowCount == baselineVisibleRows && value.pluginsMainListVisibleColumnCount == baselineVisibleColumns &&
                   DebugGetPreferencesPluginsMainListRowEnabled(rowIndex, currentEnabled) && currentEnabled == expectedEnabled &&
                   value.pluginsMainListRenderCount >= baselineRenderCount;
        },
                          snapshot),
                      std::format(L"Preferences Plugins DX main-list checkbox did not reach the expected enabled state after {}.", label));
    };

    clickCheckbox();
    SelfTest::AppendSelfTestTrace(L"Plugins click toggle: first click sent");
    requireEnabledState(! originalEnabled, L"first checkbox click");
    if (! state.failure.empty())
    {
        return false;
    }

    clickCheckbox();
    SelfTest::AppendSelfTestTrace(L"Plugins click toggle: second click sent");
    requireEnabledState(originalEnabled, L"second checkbox click");
    SelfTest::AppendSelfTestTrace(L"Plugins click toggle: complete");
    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsPageExposesLiveGridSelection(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins grid UIA selection test.");
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins grid UIA selection test.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins grid UIA selection test.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Plugins grid UIA selection test.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins grid UIA selection test.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins grid UIA selection test.");

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && true /* Phase 8: removed field */
               && value.pluginsMainListRowCount > 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not expose its DX grid surface for UIA selection validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(snapshot.pageTitle == LoadStringResource(nullptr, IDS_PREFS_CAT_PLUGINS),
                  L"Preferences page title did not switch to Plugins before UIA selection validation.");
    state.Require(snapshot.currentPageDxHostResizeFailureCount == 0u, L"Preferences Plugins page reported DX resize failures before UIA selection validation.");

    return VerifyPreferencesGridSelectionPattern(prefs, state, L"Plugins", snapshot.pluginsMainListRowCount, [](const size_t rowIndex) noexcept {
        return DebugSelectPreferencesPluginsMainListRow(rowIndex);
    });
}

[[nodiscard]] bool TestPreferencesDialogPluginsMainListHeaderDragReordersColumnsWithoutSort(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins header-reorder validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins header-reorder validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins header-reorder validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, L"Failed to focus the Preferences category host for Plugins header-reorder validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins header-reorder validation.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins header-reorder validation.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsMainListRowCount > 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle before main-list header-reorder validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(0u), L"Failed to select the first Plugins DX main-list row before header-reorder validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the baseline selected plugin row before header-reorder validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT nameHeaderRect{};
    RECT typeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, nameHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Name header rect before reorder validation.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, typeHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Type header rect before reorder validation.");
    state.Require(nameHeaderRect.right > nameHeaderRect.left && nameHeaderRect.bottom > nameHeaderRect.top,
                  L"Preferences Plugins Name header rect should be non-empty before reorder validation.");
    state.Require(typeHeaderRect.right > typeHeaderRect.left && typeHeaderRect.bottom > typeHeaderRect.top,
                  L"Preferences Plugins Type header rect should be non-empty before reorder validation.");
    state.Require(nameHeaderRect.left < typeHeaderRect.left, L"Preferences Plugins should start with Name before Type in the visible main-list header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX page host for header-reorder validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedPluginId = snapshot.pluginsSelectedPluginIdText;
    const size_t baselineVisibleRows            = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.pluginsMainListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.pluginsMainListVisibleCellCount;
    const uint64_t baselineResizeCount          = snapshot.pluginsMainListResizeCount;
    const LONG dragStartX                       = typeHeaderRect.left + ((typeHeaderRect.right - typeHeaderRect.left) / 2);
    const LONG dragY                            = typeHeaderRect.top + ((typeHeaderRect.bottom - typeHeaderRect.top) / 2);
    const LONG dragTargetX                      = nameHeaderRect.left + 12;

    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(dragStartX, dragY), MAKELPARAM(dragTargetX, dragY));

    const auto waitForReorderedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentNameHeaderRect{};
            RECT currentTypeHeaderRect{};
            const bool haveNameHeader = DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect);
            const bool haveTypeHeader = DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect);
            const bool haveSnapshot   = DebugGetPreferencesDialogSnapshot(snapshot);
            if (haveNameHeader && haveTypeHeader && haveSnapshot && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
                snapshot.currentCategory == kPrefCategoryPlugins && snapshot.pluginsSelectedPluginIdText == baselineSelectedPluginId &&
                snapshot.pluginItemSelected && ! snapshot.pluginsDetailsActive && snapshot.pluginsMainListVisibleRowCount == baselineVisibleRows &&
                snapshot.pluginsMainListVisibleColumnCount == baselineVisibleColumns && snapshot.pluginsMainListVisibleCellCount == baselineVisibleCells &&
                snapshot.pluginsMainListResizeCount == baselineResizeCount && snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u &&
                snapshot.visibleCurrentPageChildWindowCount == 1u && snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForReorderedHeaders(),
                  std::format(L"Dragging the Preferences Plugins Type header did not reorder the visible DX main-list columns without losing retained "
                              L"selection or bounded visible work; selected='{}', rows={}, cols={}, cells={}, resizeCount={}, pageResizeFailures={}.",
                              snapshot.pluginsSelectedPluginIdText,
                              snapshot.pluginsMainListVisibleRowCount,
                              snapshot.pluginsMainListVisibleColumnCount,
                              snapshot.pluginsMainListVisibleCellCount,
                              snapshot.pluginsMainListResizeCount,
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsMainListHeaderResizeChangesVisibleWidth(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins main-list header-resize validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins main-list header-resize validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
    state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                  L"Preferences category host control missing for Plugins main-list header-resize validation.");
    if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
    {
        return false;
    }

    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to focus the Preferences category host for Plugins main-list header-resize validation.");
    state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                  L"Failed to select the Preferences Plugins category for Plugins main-list header-resize validation.");
    PumpPendingMessages();
    state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                  L"Failed to restore focus to the Preferences category host after selecting Plugins for Plugins main-list header-resize validation.");

    PreferencesDebugSnapshot snapshot{};
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsMainListRowCount > 1u && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not settle before main-list header-resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(0u), L"Failed to select the first Plugins DX main-list row before header-resize validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the baseline selected plugin row before header-resize validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT nameHeaderRect{};
    RECT typeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, nameHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Name header rect before resize validation.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, typeHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Type header rect before resize validation.");
    state.Require(nameHeaderRect.right > nameHeaderRect.left && nameHeaderRect.bottom > nameHeaderRect.top,
                  L"Preferences Plugins Name header rect should be non-empty before header-resize validation.");
    state.Require(typeHeaderRect.right > typeHeaderRect.left && typeHeaderRect.bottom > typeHeaderRect.top,
                  L"Preferences Plugins Type header rect should be non-empty before header-resize validation.");
    state.Require(nameHeaderRect.left < typeHeaderRect.left, L"Preferences Plugins should start with Name before Type in the visible main-list header order.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX page host for header-resize validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedPluginId = snapshot.pluginsSelectedPluginIdText;
    const size_t baselineVisibleRows            = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.pluginsMainListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.pluginsMainListVisibleCellCount;
    const uint64_t baselineResizeCount          = snapshot.pluginsMainListResizeCount;
    const uint64_t baselineRenderCount          = snapshot.pluginsMainListRenderCount;
    const float baselineNameHeaderWidth         = static_cast<float>(nameHeaderRect.right - nameHeaderRect.left);
    const float baselineTypeHeaderLeft          = static_cast<float>(typeHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, nameHeaderRect);

    const auto waitForResizedHeaders = [&]() noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            snapshot = {};
            RECT currentNameHeaderRect{};
            RECT currentTypeHeaderRect{};
            const bool haveNameHeader          = DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect);
            const bool haveTypeHeader          = DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect);
            const bool haveSnapshot            = DebugGetPreferencesDialogSnapshot(snapshot);
            const float currentNameHeaderWidth = static_cast<float>(currentNameHeaderRect.right - currentNameHeaderRect.left);
            if (haveNameHeader && haveTypeHeader && haveSnapshot && currentNameHeaderWidth >= baselineNameHeaderWidth + 20.0f &&
                static_cast<float>(currentTypeHeaderRect.left) > baselineTypeHeaderLeft + 10.0f && currentNameHeaderRect.left < currentTypeHeaderRect.left &&
                snapshot.currentCategory == kPrefCategoryPlugins && snapshot.pluginsSelectedPluginIdText == baselineSelectedPluginId &&
                snapshot.pluginItemSelected && ! snapshot.pluginsDetailsActive && snapshot.pluginsMainListVisibleRowCount == baselineVisibleRows &&
                snapshot.pluginsMainListVisibleColumnCount == baselineVisibleColumns && snapshot.pluginsMainListVisibleCellCount == baselineVisibleCells &&
                snapshot.pluginsMainListResizeCount == baselineResizeCount && snapshot.pluginsMainListRenderCount >= baselineRenderCount &&
                snapshot.createdPaneWindowCount == 0u && snapshot.visiblePaneWindowCount == 0u && snapshot.visibleCurrentPageChildWindowCount == 1u &&
                snapshot.currentPageDxHostResizeFailureCount == 0u)
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        return false;
    };

    state.Require(waitForResizedHeaders(),
                  std::format(L"Dragging the Preferences Plugins Name header edge did not widen the visible DX column without losing retained selection or "
                              L"bounded visible work; selected='{}', rows={}, cols={}, cells={}, renderCount={}, resizeCount={}, pageResizeFailures={}.",
                              snapshot.pluginsSelectedPluginIdText,
                              snapshot.pluginsMainListVisibleRowCount,
                              snapshot.pluginsMainListVisibleColumnCount,
                              snapshot.pluginsMainListVisibleCellCount,
                              snapshot.pluginsMainListRenderCount,
                              snapshot.pluginsMainListResizeCount,
                              snapshot.currentPageDxHostResizeFailureCount));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsReorderedResizedColumnsSurviveSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr wchar_t kSearchText[] = L"__codex_no_match__";

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins reordered-resized/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins reordered-resized/search validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto hasPluginsPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount >= 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    };

    const auto hasStablePluginsPageState = [&](const PreferencesDebugSnapshot& value, const size_t minimumRowCount) noexcept
    { return hasPluginsPageSurfaceState(value) && value.pluginsSearchText.empty() && value.pluginsMainListRowCount >= minimumRowCount; };

    const auto navigateToPluginsPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins reordered-resized/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Plugins reordered-resized/search validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category before reordered-resized/search validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();
        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            return true;
        }

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i <= 7; ++i)
        {
            PreferencesDebugSnapshot candidate{};
            if (DebugGetPreferencesDialogSnapshot(candidate) && hasPluginsPageSurfaceState(candidate))
            {
                outSnapshot = std::move(candidate);
                return true;
            }

            if (i != 7)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }
        }

        return waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    state.Require(navigateToPluginsPage(snapshot), L"Preferences Plugins page did not settle before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! hasStablePluginsPageState(snapshot, 2u))
    {
        state.Require(DebugSetPreferencesPluginsSearchText(L""),
                      L"Failed to clear the Plugins search field before establishing the reordered-resized/search baseline.");
        state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, snapshot),
                      L"Preferences Plugins page did not restore a cleared multi-row baseline before reordered-resized/search validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const size_t baselineRowCount = snapshot.pluginsMainListRowCount;
    state.Require(DebugSelectPreferencesPluginsMainListRow(0u),
                  L"Failed to select the first Plugins DX main-list row before reordered-resized/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the baseline selected plugin row before reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT nameHeaderRect{};
    RECT typeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, nameHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Name header rect before combined reordered-resized/search validation.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, typeHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Type header rect before combined reordered-resized/search validation.");
    state.Require(nameHeaderRect.left < typeHeaderRect.left,
                  L"Preferences Plugins should start with Name before Type before combined reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX page host for reordered-resized/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedPluginId = snapshot.pluginsSelectedPluginIdText;
    const size_t baselineVisibleRows            = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.pluginsMainListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.pluginsMainListVisibleCellCount;
    const uint64_t baselineResizeCount          = snapshot.pluginsMainListResizeCount;
    const uint64_t baselineRenderCount          = snapshot.pluginsMainListRenderCount;

    const LONG reorderStartX  = typeHeaderRect.left + ((typeHeaderRect.right - typeHeaderRect.left) / 2);
    const LONG reorderY       = typeHeaderRect.top + ((typeHeaderRect.bottom - typeHeaderRect.top) / 2);
    const LONG reorderTargetX = nameHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               value.currentCategory == kPrefCategoryPlugins && value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected &&
               ! value.pluginsDetailsActive && value.pluginsMainListRowCount == baselineRowCount &&
               value.pluginsMainListVisibleRowCount == baselineVisibleRows && value.pluginsMainListVisibleColumnCount == baselineVisibleColumns &&
               value.pluginsMainListVisibleCellCount == baselineVisibleCells && value.pluginsMainListResizeCount == baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins combined reordered layout did not settle before search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedNameHeaderRect{};
    RECT reorderedTypeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, reorderedNameHeaderRect),
                  L"Failed to capture the reordered Preferences Plugins first logical header rect before resize.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, reorderedTypeHeaderRect),
                  L"Failed to capture the reordered Preferences Plugins second logical header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedTypeHeaderRect.right - reorderedTypeHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedNameHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedTypeHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsMainListRowCount == baselineRowCount && value.pluginsMainListVisibleRowCount == baselineVisibleRows &&
               value.pluginsMainListVisibleColumnCount == baselineVisibleColumns && value.pluginsMainListVisibleCellCount == baselineVisibleCells &&
               value.pluginsMainListResizeCount == baselineResizeCount && value.pluginsMainListRenderCount >= baselineRenderCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins combined reorder+resize did not settle before search round-trip validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedNameHeaderRect{};
    RECT resizedReorderedTypeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, resizedReorderedNameHeaderRect),
                  L"Failed to capture the resized reordered Preferences Plugins first logical header rect.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, resizedReorderedTypeHeaderRect),
                  L"Failed to capture the resized reordered Preferences Plugins second logical header rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedTypeHeaderRect.right - resizedReorderedTypeHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedNameHeaderRect.left);

    state.Require(DebugFocusPreferencesPluginsSearchField(),
                  L"Failed to focus the Preferences Plugins DX search field before live no-match search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText.empty() &&
               value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::SearchField && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX search field did not gain focus before live no-match search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND focusedWindow = GetFocus();
    state.Require(focusedWindow != nullptr && IsWindow(focusedWindow) != FALSE,
                  L"Preferences Plugins search field did not expose a focused Win32 input target before live no-match search validation.");
    if (! focusedWindow || IsWindow(focusedWindow) == FALSE)
    {
        return false;
    }

    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(kSearchText));
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText == kSearchText && value.pluginsMainListRowCount == 0u; },
                                  snapshot),
                  L"Preferences Plugins filtered no-match search rebuild did not settle during reordered-resized/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    focusedWindow = GetFocus();
    state.Require(focusedWindow != nullptr && IsWindow(focusedWindow) != FALSE,
                  L"Preferences Plugins search field did not expose a focused Win32 input target before live clear-back validation.");
    if (! focusedWindow || IsWindow(focusedWindow) == FALSE)
    {
        return false;
    }

    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSearchText.empty() && value.pluginsMainListRowCount == baselineRowCount;
    },
                      snapshot),
                  L"Preferences Plugins clearing the search rebuild did not restore the full row set with the combined reordered+resized layout intact.");
    state.Require(
        snapshot.pluginsSelectedPluginIdText == baselineSelectedPluginId,
        std::format(L"Preferences Plugins should restore the same selected plugin after the combined no-match search round-trip; expected='{}', actual='{}'.",
                    baselineSelectedPluginId,
                    snapshot.pluginsSelectedPluginIdText));

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr wchar_t kSearchText[] = L"__codex_no_match__";

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins reordered-resized-copy/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins reordered-resized-copy/search validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    const auto hasPluginsPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount >= 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    };

    const auto hasStablePluginsPageState = [&](const PreferencesDebugSnapshot& value, const size_t minimumRowCount) noexcept
    { return hasPluginsPageSurfaceState(value) && value.pluginsSearchText.empty() && value.pluginsMainListRowCount >= minimumRowCount; };

    const auto navigateToPluginsPage = [&](PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      L"Preferences category host control missing for Plugins reordered-resized-copy/search validation.");
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost,
                      L"Failed to focus the Preferences category host for Plugins reordered-resized-copy/search validation.");
        if (! state.failure.empty())
        {
            return false;
        }

        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            return true;
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      L"Failed to select the Preferences Plugins category before reordered-resized-copy/search validation.");
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();
        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            return true;
        }

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i <= 7; ++i)
        {
            PreferencesDebugSnapshot candidate{};
            if (DebugGetPreferencesDialogSnapshot(candidate) && hasPluginsPageSurfaceState(candidate))
            {
                outSnapshot = std::move(candidate);
                return true;
            }

            if (i != 7)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }
        }

        return waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    state.Require(navigateToPluginsPage(snapshot), L"Preferences Plugins page did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    if (! hasStablePluginsPageState(snapshot, 2u))
    {
        state.Require(DebugSetPreferencesPluginsSearchText(L""),
                      L"Failed to clear the Plugins search field before establishing the reordered-resized-copy/search baseline.");
        state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, snapshot),
                      L"Preferences Plugins page did not restore a cleared multi-row baseline before reordered-resized-copy/search validation.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    const size_t baselineRowCount = snapshot.pluginsMainListRowCount;
    state.Require(DebugSelectPreferencesPluginsMainListRow(0u),
                  L"Failed to select the first Plugins DX main-list row before reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the baseline selected plugin row before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT nameHeaderRect{};
    RECT typeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, nameHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Name header rect before reordered-resized-copy/search validation.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, typeHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Type header rect before reordered-resized-copy/search validation.");
    state.Require(nameHeaderRect.left < typeHeaderRect.left,
                  L"Preferences Plugins should start with Name before Type before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX keyboard host for reordered-resized-copy/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedPluginId = snapshot.pluginsSelectedPluginIdText;
    const size_t baselineVisibleRows            = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.pluginsMainListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.pluginsMainListVisibleCellCount;
    const uint64_t baselineResizeCount          = snapshot.pluginsMainListResizeCount;
    const uint64_t baselineRenderCount          = snapshot.pluginsMainListRenderCount;

    const std::optional<PrefsPluginListItem> selectedPlugin = PrefsPlugins::FindItemById(baselineSelectedPluginId);
    state.Require(selectedPlugin.has_value(),
                  std::format(L"Could not resolve the selected Plugins row '{}' for reordered-resized-copy/search validation.", baselineSelectedPluginId));
    if (! selectedPlugin.has_value())
    {
        return false;
    }

    const std::wstring expectedNameText = std::wstring(PrefsPlugins::GetDisplayName(selectedPlugin.value()));
    const std::wstring expectedTypeText = selectedPlugin->type == PrefsPluginType::FileSystem ? LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_FILE_SYSTEM)
                                                                                              : LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_VIEWER);

    const LONG reorderStartX  = typeHeaderRect.left + ((typeHeaderRect.right - typeHeaderRect.left) / 2);
    const LONG reorderY       = typeHeaderRect.top + ((typeHeaderRect.bottom - typeHeaderRect.top) / 2);
    const LONG reorderTargetX = nameHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               value.currentCategory == kPrefCategoryPlugins && value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected &&
               ! value.pluginsDetailsActive && value.pluginsMainListRowCount == baselineRowCount &&
               value.pluginsMainListVisibleRowCount == baselineVisibleRows && value.pluginsMainListVisibleColumnCount == baselineVisibleColumns &&
               value.pluginsMainListVisibleCellCount == baselineVisibleCells && value.pluginsMainListResizeCount == baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins reordered layout did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedNameHeaderRect{};
    RECT reorderedTypeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, reorderedNameHeaderRect),
                  L"Failed to capture the reordered Preferences Plugins first logical header rect before resize.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, reorderedTypeHeaderRect),
                  L"Failed to capture the reordered Preferences Plugins second logical header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedTypeHeaderRect.right - reorderedTypeHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedNameHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedTypeHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsMainListRowCount == baselineRowCount && value.pluginsMainListVisibleRowCount == baselineVisibleRows &&
               value.pluginsMainListVisibleColumnCount == baselineVisibleColumns && value.pluginsMainListVisibleCellCount == baselineVisibleCells &&
               value.pluginsMainListResizeCount == baselineResizeCount && value.pluginsMainListRenderCount >= baselineRenderCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins combined reorder+resize did not settle before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedNameHeaderRect{};
    RECT resizedReorderedTypeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, resizedReorderedNameHeaderRect),
                  L"Failed to capture the resized reordered Preferences Plugins first logical header rect.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, resizedReorderedTypeHeaderRect),
                  L"Failed to capture the resized reordered Preferences Plugins second logical header rect.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedTypeHeaderRect.right - resizedReorderedTypeHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedNameHeaderRect.left);

    state.Require(DebugFocusPreferencesPluginsSearchField(),
                  L"Failed to focus the Preferences Plugins DX search field before reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText.empty() &&
               value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::SearchField && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX search field did not gain focus before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    HWND focusedWindow = GetFocus();
    state.Require(focusedWindow != nullptr && IsWindow(focusedWindow) != FALSE,
                  L"Preferences Plugins search field did not expose a focused Win32 input target before live no-match search validation.");
    if (! focusedWindow || IsWindow(focusedWindow) == FALSE)
    {
        return false;
    }

    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(kSearchText));
    state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept
    { return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText == kSearchText && value.pluginsMainListRowCount == 0u; },
                                  snapshot),
                  L"Preferences Plugins filtered no-match search rebuild did not settle during reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    focusedWindow = GetFocus();
    state.Require(focusedWindow != nullptr && IsWindow(focusedWindow) != FALSE,
                  L"Preferences Plugins search field did not expose a focused Win32 input target before live clear-back validation.");
    if (! focusedWindow || IsWindow(focusedWindow) == FALSE)
    {
        return false;
    }

    SendMessageW(focusedWindow, EM_SETSEL, static_cast<WPARAM>(0), static_cast<LPARAM>(-1));
    SendMessageW(focusedWindow, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
    state.Require(
        waitForSnapshot(
            [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSearchText.empty() && value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected &&
               ! value.pluginsDetailsActive && value.pluginsMainListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u &&
               value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
            snapshot),
        L"Preferences Plugins clearing the search rebuild did not restore the full row set with the combined reordered+resized layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesPluginsMainList(),
                  L"Failed to focus the Preferences Plugins DX main list before reordered-resized-copy/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::MainList &&
               value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX main list did not restore focus before reordered-resized-copy/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    ClearClipboardContents(prefs);
    SendMessageW(activePage, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(activePage, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(prefs);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(), L"Preferences Plugins Ctrl+C should copy the combined reordered+resized row content after the search round-trip.");
    state.Require(copiedSelection.rfind((expectedTypeText + L"\t"), 0u) == 0u,
                  L"Preferences Plugins clipboard copy should start with the visible Type column after the combined search round-trip.");
    state.Require(copiedSelection.find(expectedNameText) != std::wstring::npos,
                  L"Preferences Plugins clipboard copy should still include the selected plugin name after the combined search round-trip.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr wchar_t kNoMatchSearch[] = L"__codex_no_match__";

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins reordered-resized-sort/search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins reordered-resized-sort/search validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    const auto hasPluginsPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount >= 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    };

    const auto hasStablePluginsPageState = [&](const PreferencesDebugSnapshot& value, const size_t minimumRowCount) noexcept
    { return hasPluginsPageSurfaceState(value) && value.pluginsSearchText.empty() && value.pluginsMainListRowCount >= minimumRowCount; };

    const auto navigateToPluginsPage = [&](PreferencesDebugSnapshot& outSnapshot, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing for {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, std::format(L"Failed to focus the Preferences category host for {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            if (! hasStablePluginsPageState(outSnapshot, 2u))
            {
                state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
                state.Require(
                    waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, outSnapshot),
                    std::format(L"Preferences Plugins page did not restore a cleared multi-row baseline before {}.", context));
            }
            return state.failure.empty();
        }

        PreferencesDebugSnapshot candidate{};
        if (DebugGetPreferencesDialogSnapshot(candidate) && candidate.currentCategory == kPrefCategoryPlugins &&
            (candidate.pluginItemSelected || candidate.pluginsDetailsActive))
        {
            for (int i = 0; i < 3; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_LEFT, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_LEFT, 0);
                PumpPendingMessages();
                if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
                {
                    if (! hasStablePluginsPageState(outSnapshot, 2u))
                    {
                        state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
                        state.Require(
                            waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, outSnapshot),
                            std::format(L"Preferences Plugins page did not restore a cleared multi-row baseline before {}.", context));
                    }
                    return state.failure.empty();
                }
            }
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      std::format(L"Failed to select the Preferences Plugins category before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();
        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            if (! hasStablePluginsPageState(outSnapshot, 2u))
            {
                state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
                state.Require(
                    waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, outSnapshot),
                    std::format(L"Preferences Plugins page did not restore a cleared multi-row baseline before {}.", context));
            }
            return state.failure.empty();
        }

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i <= 7; ++i)
        {
            candidate = {};
            if (DebugGetPreferencesDialogSnapshot(candidate) && hasPluginsPageSurfaceState(candidate))
            {
                outSnapshot = std::move(candidate);
                break;
            }

            if (i != 7)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }
        }

        state.Require(waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot), std::format(L"Preferences Plugins page did not settle before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStablePluginsPageState(outSnapshot, 2u))
        {
            state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
            state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, outSnapshot),
                          std::format(L"Preferences Plugins page did not restore a cleared multi-row baseline before {}.", context));
        }
        return state.failure.empty();
    };

    state.Require(navigateToPluginsPage(snapshot, L"reordered-resized-sort/search validation"),
                  L"Preferences Plugins page did not settle before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.pluginsMainListRowCount;
    state.Require(DebugSelectPreferencesPluginsMainListRow(0u),
                  L"Failed to select the first Plugins DX main-list row before reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the baseline selected plugin row before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT nameHeaderRect{};
    RECT typeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, nameHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Name header rect before reordered-resized-sort/search validation.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, typeHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Type header rect before reordered-resized-sort/search validation.");
    state.Require(nameHeaderRect.left < typeHeaderRect.left,
                  L"Preferences Plugins should start with Name before Type before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX page host for reordered-resized-sort/search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedPluginId = snapshot.pluginsSelectedPluginIdText;
    const size_t baselineVisibleRows            = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.pluginsMainListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.pluginsMainListVisibleCellCount;
    const uint64_t baselineResizeCount          = snapshot.pluginsMainListResizeCount;
    const uint64_t baselineRenderCount          = snapshot.pluginsMainListRenderCount;

    const LONG reorderStartX  = typeHeaderRect.left + ((typeHeaderRect.right - typeHeaderRect.left) / 2);
    const LONG reorderY       = typeHeaderRect.top + ((typeHeaderRect.bottom - typeHeaderRect.top) / 2);
    const LONG reorderTargetX = nameHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               value.currentCategory == kPrefCategoryPlugins && value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected &&
               ! value.pluginsDetailsActive && value.pluginsMainListRowCount == baselineRowCount &&
               value.pluginsMainListVisibleRowCount == baselineVisibleRows && value.pluginsMainListVisibleColumnCount == baselineVisibleColumns &&
               value.pluginsMainListVisibleCellCount == baselineVisibleCells && value.pluginsMainListResizeCount == baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins header reorder did not settle before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedNameHeaderRect{};
    RECT reorderedTypeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, reorderedNameHeaderRect),
                  L"Failed to capture the reordered Preferences Plugins Name header rect before resize.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, reorderedTypeHeaderRect),
                  L"Failed to capture the reordered Preferences Plugins Type header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedTypeHeaderRect.right - reorderedTypeHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedNameHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedTypeHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsMainListRowCount == baselineRowCount && value.pluginsMainListVisibleRowCount == baselineVisibleRows &&
               value.pluginsMainListVisibleColumnCount == baselineVisibleColumns && value.pluginsMainListVisibleCellCount == baselineVisibleCells &&
               value.pluginsMainListResizeCount == baselineResizeCount && value.pluginsMainListRenderCount >= baselineRenderCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins combined reorder+resize did not settle before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedNameHeaderRect{};
    RECT resizedReorderedTypeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, resizedReorderedNameHeaderRect),
                  L"Failed to capture the resized reordered Preferences Plugins Name header rect before sort/search validation.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, resizedReorderedTypeHeaderRect),
                  L"Failed to capture the resized reordered Preferences Plugins Type header rect before sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedTypeHeaderRect.right - resizedReorderedTypeHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedNameHeaderRect.left);
    const LONG sortClickX       = resizedReorderedTypeHeaderRect.left + ((resizedReorderedTypeHeaderRect.right - resizedReorderedTypeHeaderRect.left) / 2);
    const LONG sortClickY       = resizedReorderedTypeHeaderRect.top + ((resizedReorderedTypeHeaderRect.bottom - resizedReorderedTypeHeaderRect.top) / 2);
    const LPARAM sortClickPoint = MAKELPARAM(sortClickX, sortClickY);
    const HWND sortWindow       = ResolveMouseInputWindowForHostPoint(activePage, sortClickPoint);
    state.Require(sortWindow != nullptr && IsWindow(sortWindow) != FALSE,
                  L"Failed to resolve the Preferences Plugins DX mouse-input window for sort/search validation.");
    if (! sortWindow || IsWindow(sortWindow) == FALSE)
    {
        return false;
    }

    const LPARAM mappedSortClickPoint = MapClientPointLParam(activePage, sortWindow, sortClickPoint);
    for (int i = 0; i < 2; ++i)
    {
        SendMessageW(sortWindow, WM_MOUSEMOVE, 0, mappedSortClickPoint);
        SendMessageW(sortWindow, WM_LBUTTONDOWN, MK_LBUTTON, mappedSortClickPoint);
        SendMessageW(sortWindow, WM_LBUTTONUP, 0, mappedSortClickPoint);
        PumpPendingMessages();
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(0u), L"Failed to select the first visible Plugins DX row after sort cycles.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryPlugins &&
               ! value.pluginsSelectedPluginIdText.empty() && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsMainListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins sort cycles did not settle before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPluginId                     = snapshot.pluginsSelectedPluginIdText;
    const std::optional<PrefsPluginListItem> selectedPlugin = PrefsPlugins::FindItemById(selectedPluginId);
    state.Require(selectedPlugin.has_value(),
                  std::format(L"Could not resolve the selected Plugins row '{}' before the sort/search round-trip.", selectedPluginId));
    if (! selectedPlugin.has_value())
    {
        return false;
    }

    const std::wstring targetSearchText = std::wstring(PrefsPlugins::GetDisplayName(selectedPlugin.value()));
    state.Require(! targetSearchText.empty(), L"Preferences Plugins should expose a selected display name before reordered-resized-sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsSearchText(targetSearchText),
                  L"Failed to set the Plugins search text to the selected display name during reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSearchText == targetSearchText && value.pluginsMainListRowCount > 0u && value.pluginsSelectedPluginIdText == selectedPluginId &&
               value.pluginItemSelected && ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins filtered rebuild did not preserve the combined reordered-resized sorted layout.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsSearchText(kNoMatchSearch),
                  L"Failed to set the Plugins no-match search text during reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText == kNoMatchSearch && value.pluginsMainListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins no-match search did not settle before reordered-resized-sort/search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsSearchText(targetSearchText),
                  L"Failed to restore the Plugins search text to the selected display name before reordered-resized-sort/search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSearchText == targetSearchText && value.pluginsMainListRowCount > 0u && value.pluginsSelectedPluginIdText == selectedPluginId &&
               value.pluginItemSelected && ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins filtered restore did not preserve the combined reordered-resized sorted layout.");

    return state.failure.empty();
}

[[nodiscard]] bool TestPreferencesDialogPluginsReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip(HWND mainWindow,
                                                                                                                          CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr wchar_t kNoMatchSearch[] = L"__codex_no_match__";

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (const HWND existing = GetPreferencesDialogHandle(); existing && IsWindow(existing) != FALSE)
    {
        PostMessageW(existing, WM_CLOSE, 0, 0);
        state.Require(WaitForWindowClosed(existing, SelfTest::Scale(2000ms)),
                      L"Existing Preferences window did not close before Plugins reordered-resized-copy/sort-search validation.");
    }
    if (! state.failure.empty())
    {
        return false;
    }

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_FILE_PREFERENCES, 0), 0);
    const HWND prefs = WaitForWindow([] noexcept { return GetPreferencesDialogHandle(); }, SelfTest::Scale(3000ms));
    state.Require(prefs != nullptr && IsWindow(prefs) != FALSE, L"Preferences window did not open for Plugins reordered-resized-copy/sort-search validation.");
    if (! prefs || IsWindow(prefs) == FALSE)
    {
        return false;
    }

    const auto closeWindow = wil::scope_exit([&]() noexcept
    {
        if (IsWindow(prefs) != FALSE)
        {
            PostMessageW(prefs, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(prefs, SelfTest::Scale(2000ms)));
        }
    });

    const auto waitForSnapshot = [&](const auto& predicate, PreferencesDebugSnapshot& outSnapshot) noexcept
    {
        const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
        while (std::chrono::steady_clock::now() < deadline)
        {
            PumpPendingMessages();
            outSnapshot = {};
            if (DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot))
            {
                return true;
            }

            std::this_thread::sleep_for(20ms);
        }

        outSnapshot = {};
        return DebugGetPreferencesDialogSnapshot(outSnapshot) && predicate(outSnapshot);
    };

    PreferencesDebugSnapshot snapshot{};
    const auto hasPluginsPageSurfaceState = [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && ! value.pluginItemSelected && ! value.pluginsDetailsActive && value.pluginsPaneVisible &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount >= 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    };

    const auto hasStablePluginsPageState = [&](const PreferencesDebugSnapshot& value, const size_t minimumRowCount) noexcept
    { return hasPluginsPageSurfaceState(value) && value.pluginsSearchText.empty() && value.pluginsMainListRowCount >= minimumRowCount; };

    const auto navigateToPluginsPage = [&](PreferencesDebugSnapshot& outSnapshot, std::wstring_view context) noexcept
    {
        const HWND categoryTreeHost = GetDlgItem(prefs, IDC_PREFS_CATEGORY_LIST);
        state.Require(categoryTreeHost != nullptr && IsWindow(categoryTreeHost) != FALSE,
                      std::format(L"Preferences category host control missing for {}.", context));
        if (! categoryTreeHost || IsWindow(categoryTreeHost) == FALSE)
        {
            return false;
        }

        state.Require(SetFocus(categoryTreeHost) == categoryTreeHost, std::format(L"Failed to focus the Preferences category host for {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            if (! hasStablePluginsPageState(outSnapshot, 2u))
            {
                state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
                state.Require(
                    waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, outSnapshot),
                    std::format(L"Preferences Plugins page did not restore a cleared multi-row baseline before {}.", context));
            }
            return state.failure.empty();
        }

        PreferencesDebugSnapshot candidate{};
        if (DebugGetPreferencesDialogSnapshot(candidate) && candidate.currentCategory == kPrefCategoryPlugins &&
            (candidate.pluginItemSelected || candidate.pluginsDetailsActive))
        {
            for (int i = 0; i < 3; ++i)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_LEFT, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_LEFT, 0);
                PumpPendingMessages();
                if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
                {
                    if (! hasStablePluginsPageState(outSnapshot, 2u))
                    {
                        state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
                        state.Require(
                            waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, outSnapshot),
                            std::format(L"Preferences Plugins page did not restore a cleared multi-row baseline before {}.", context));
                    }
                    return state.failure.empty();
                }
            }
        }

        state.Require(DebugSelectPreferencesCategory(kPrefCategoryPlugins),
                      std::format(L"Failed to select the Preferences Plugins category before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }
        PumpPendingMessages();
        if (waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot))
        {
            if (! hasStablePluginsPageState(outSnapshot, 2u))
            {
                state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
                state.Require(
                    waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, outSnapshot),
                    std::format(L"Preferences Plugins page did not restore a cleared multi-row baseline before {}.", context));
            }
            return state.failure.empty();
        }

        SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_HOME, 0);
        SendMessageW(categoryTreeHost, WM_KEYUP, VK_HOME, 0);
        PumpPendingMessages();
        for (int i = 0; i <= 7; ++i)
        {
            candidate = {};
            if (DebugGetPreferencesDialogSnapshot(candidate) && hasPluginsPageSurfaceState(candidate))
            {
                outSnapshot = std::move(candidate);
                break;
            }

            if (i != 7)
            {
                SendMessageW(categoryTreeHost, WM_KEYDOWN, VK_DOWN, 0);
                SendMessageW(categoryTreeHost, WM_KEYUP, VK_DOWN, 0);
                PumpPendingMessages();
            }
        }

        state.Require(waitForSnapshot(hasPluginsPageSurfaceState, outSnapshot), std::format(L"Preferences Plugins page did not settle before {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        if (! hasStablePluginsPageState(outSnapshot, 2u))
        {
            state.Require(DebugSetPreferencesPluginsSearchText(L""), std::format(L"Failed to clear the Plugins search field before {}.", context));
            state.Require(waitForSnapshot([&](const PreferencesDebugSnapshot& value) noexcept { return hasStablePluginsPageState(value, 2u); }, outSnapshot),
                          std::format(L"Preferences Plugins page did not restore a cleared multi-row baseline before {}.", context));
        }
        return state.failure.empty();
    };

    state.Require(navigateToPluginsPage(snapshot, L"reordered-resized-copy/sort-search validation"),
                  L"Preferences Plugins page did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const size_t baselineRowCount = snapshot.pluginsMainListRowCount;
    state.Require(DebugSelectPreferencesPluginsMainListRow(0u),
                  L"Failed to select the first Plugins DX main-list row before reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               ! value.pluginsSelectedPluginIdText.empty() && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins page did not retain the baseline selected plugin row before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT nameHeaderRect{};
    RECT typeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, nameHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Name header rect before reordered-resized-copy/sort-search validation.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, typeHeaderRect),
                  L"Failed to capture the visible Preferences Plugins Type header rect before reordered-resized-copy/sort-search validation.");
    state.Require(nameHeaderRect.left < typeHeaderRect.left,
                  L"Preferences Plugins should start with Name before Type before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HWND activePage = DebugGetPreferencesActivePageDxHostHandle();
    state.Require(activePage != nullptr && IsWindow(activePage) != FALSE,
                  L"Failed to resolve the active Preferences Plugins DX keyboard host for reordered-resized-copy/sort-search validation.");
    if (! activePage || IsWindow(activePage) == FALSE)
    {
        return false;
    }

    const std::wstring baselineSelectedPluginId = snapshot.pluginsSelectedPluginIdText;
    const size_t baselineVisibleRows            = snapshot.pluginsMainListVisibleRowCount;
    const size_t baselineVisibleColumns         = snapshot.pluginsMainListVisibleColumnCount;
    const size_t baselineVisibleCells           = snapshot.pluginsMainListVisibleCellCount;
    const uint64_t baselineResizeCount          = snapshot.pluginsMainListResizeCount;
    const uint64_t baselineRenderCount          = snapshot.pluginsMainListRenderCount;

    const LONG reorderStartX  = typeHeaderRect.left + ((typeHeaderRect.right - typeHeaderRect.left) / 2);
    const LONG reorderY       = typeHeaderRect.top + ((typeHeaderRect.bottom - typeHeaderRect.top) / 2);
    const LONG reorderTargetX = nameHeaderRect.left + 12;
    SendMouseDragToResolvedPointWindow(activePage, MAKELPARAM(reorderStartX, reorderY), MAKELPARAM(reorderTargetX, reorderY));

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               value.currentCategory == kPrefCategoryPlugins && value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected &&
               ! value.pluginsDetailsActive && value.pluginsMainListRowCount == baselineRowCount &&
               value.pluginsMainListVisibleRowCount == baselineVisibleRows && value.pluginsMainListVisibleColumnCount == baselineVisibleColumns &&
               value.pluginsMainListVisibleCellCount == baselineVisibleCells && value.pluginsMainListResizeCount == baselineResizeCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins header reorder did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT reorderedNameHeaderRect{};
    RECT reorderedTypeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, reorderedNameHeaderRect),
                  L"Failed to capture the reordered Preferences Plugins Name header rect before resize.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, reorderedTypeHeaderRect),
                  L"Failed to capture the reordered Preferences Plugins Type header rect before resize.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float baselineFirstVisibleWidth = static_cast<float>(reorderedTypeHeaderRect.right - reorderedTypeHeaderRect.left);
    const float baselineSecondVisibleLeft = static_cast<float>(reorderedNameHeaderRect.left);
    SendScaledHeaderResizeDrag(activePage, reorderedTypeHeaderRect);

    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= baselineFirstVisibleWidth + 8.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= baselineSecondVisibleLeft + 4.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSelectedPluginIdText == baselineSelectedPluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsMainListRowCount == baselineRowCount && value.pluginsMainListVisibleRowCount == baselineVisibleRows &&
               value.pluginsMainListVisibleColumnCount == baselineVisibleColumns && value.pluginsMainListVisibleCellCount == baselineVisibleCells &&
               value.pluginsMainListResizeCount == baselineResizeCount && value.pluginsMainListRenderCount >= baselineRenderCount &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins combined reorder+resize did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    RECT resizedReorderedNameHeaderRect{};
    RECT resizedReorderedTypeHeaderRect{};
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(0u, resizedReorderedNameHeaderRect),
                  L"Failed to capture the resized reordered Preferences Plugins Name header rect before sort/search validation.");
    state.Require(DebugGetPreferencesPluginsMainListHeaderClientRect(1u, resizedReorderedTypeHeaderRect),
                  L"Failed to capture the resized reordered Preferences Plugins Type header rect before sort/search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const float resizedFirstVisibleWidth = static_cast<float>(resizedReorderedTypeHeaderRect.right - resizedReorderedTypeHeaderRect.left);
    const float resizedSecondVisibleLeft = static_cast<float>(resizedReorderedNameHeaderRect.left);
    const LONG sortClickX       = resizedReorderedTypeHeaderRect.left + ((resizedReorderedTypeHeaderRect.right - resizedReorderedTypeHeaderRect.left) / 2);
    const LONG sortClickY       = resizedReorderedTypeHeaderRect.top + ((resizedReorderedTypeHeaderRect.bottom - resizedReorderedTypeHeaderRect.top) / 2);
    const LPARAM sortClickPoint = MAKELPARAM(sortClickX, sortClickY);
    const HWND sortWindow       = ResolveMouseInputWindowForHostPoint(activePage, sortClickPoint);
    state.Require(sortWindow != nullptr && IsWindow(sortWindow) != FALSE,
                  L"Failed to resolve the Preferences Plugins DX mouse-input window for copy/sort-search validation.");
    if (! sortWindow || IsWindow(sortWindow) == FALSE)
    {
        return false;
    }

    const LPARAM mappedSortClickPoint = MapClientPointLParam(activePage, sortWindow, sortClickPoint);
    for (int i = 0; i < 2; ++i)
    {
        SendMessageW(sortWindow, WM_MOUSEMOVE, 0, mappedSortClickPoint);
        SendMessageW(sortWindow, WM_LBUTTONDOWN, MK_LBUTTON, mappedSortClickPoint);
        SendMessageW(sortWindow, WM_LBUTTONUP, 0, mappedSortClickPoint);
        PumpPendingMessages();
    }

    state.Require(DebugSelectPreferencesPluginsMainListRow(0u), L"Failed to select the first visible Plugins DX row after sort cycles.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryPlugins &&
               ! value.pluginsSelectedPluginIdText.empty() && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.pluginsMainListRowCount == baselineRowCount && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins sort cycles did not settle before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring selectedPluginId                     = snapshot.pluginsSelectedPluginIdText;
    const std::optional<PrefsPluginListItem> selectedPlugin = PrefsPlugins::FindItemById(selectedPluginId);
    state.Require(selectedPlugin.has_value(),
                  std::format(L"Could not resolve the selected Plugins row '{}' before the copy/sort-search round-trip.", selectedPluginId));
    if (! selectedPlugin.has_value())
    {
        return false;
    }

    const std::wstring targetSearchText = std::wstring(PrefsPlugins::GetDisplayName(selectedPlugin.value()));
    state.Require(! targetSearchText.empty(),
                  L"Preferences Plugins should expose a selected display name before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsSearchText(targetSearchText),
                  L"Failed to set the Plugins search text to the selected display name during reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSearchText == targetSearchText && value.pluginsMainListRowCount > 0u && value.pluginsSelectedPluginIdText == selectedPluginId &&
               value.pluginItemSelected && ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins filtered rebuild did not preserve the combined reordered-resized sorted layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsSearchText(kNoMatchSearch),
                  L"Failed to set the Plugins no-match search text during reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsSearchText == kNoMatchSearch && value.pluginsMainListRowCount == 0u &&
               value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u && value.visibleCurrentPageChildWindowCount == 1u &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins no-match search did not settle before reordered-resized-copy/sort-search restoration.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugSetPreferencesPluginsSearchText(targetSearchText),
                  L"Failed to restore the Plugins search text to the selected display name before reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        RECT currentNameHeaderRect{};
        RECT currentTypeHeaderRect{};
        return DebugGetPreferencesPluginsMainListHeaderClientRect(0u, currentNameHeaderRect) &&
               DebugGetPreferencesPluginsMainListHeaderClientRect(1u, currentTypeHeaderRect) && currentTypeHeaderRect.left + 4 < currentNameHeaderRect.left &&
               static_cast<float>(currentTypeHeaderRect.right - currentTypeHeaderRect.left) >= resizedFirstVisibleWidth - 2.0f &&
               static_cast<float>(currentNameHeaderRect.left) >= resizedSecondVisibleLeft - 2.0f && value.currentCategory == kPrefCategoryPlugins &&
               value.pluginsSearchText == targetSearchText && value.pluginsMainListRowCount > 0u && value.pluginsSelectedPluginIdText == selectedPluginId &&
               value.pluginItemSelected && ! value.pluginsDetailsActive && value.createdPaneWindowCount == 0u && value.visiblePaneWindowCount == 0u &&
               value.visibleCurrentPageChildWindowCount == 1u && value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins filtered restore did not preserve the combined reordered-resized sorted layout before copy validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(DebugFocusPreferencesPluginsMainList(),
                  L"Failed to focus the Preferences Plugins DX main list before reordered-resized-copy/sort-search validation.");
    state.Require(waitForSnapshot(
                      [&](const PreferencesDebugSnapshot& value) noexcept
    {
        return value.currentCategory == kPrefCategoryPlugins && value.pluginsFocusTarget == PreferencesPluginsDebugFocusTarget::MainList &&
               value.pluginsSelectedPluginIdText == selectedPluginId && value.pluginItemSelected && ! value.pluginsDetailsActive &&
               value.currentPageDxHostResizeFailureCount == 0u;
    },
                      snapshot),
                  L"Preferences Plugins DX main list did not restore focus before reordered-resized-copy/sort-search validation.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::wstring expectedNameText = std::wstring(PrefsPlugins::GetDisplayName(selectedPlugin.value()));
    const std::wstring expectedTypeText = selectedPlugin->type == PrefsPluginType::FileSystem ? LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_FILE_SYSTEM)
                                                                                              : LoadStringResource(nullptr, IDS_PREFS_PLUGINS_TYPE_VIEWER);
    ClearClipboardContents(prefs);
    SendMessageW(activePage, WM_KEYDOWN, VK_CONTROL, 0);
    SendMessageW(activePage, WM_KEYDOWN, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, static_cast<WPARAM>(L'C'), 0);
    SendMessageW(activePage, WM_KEYUP, VK_CONTROL, 0);

    std::wstring copiedSelection;
    for (size_t retry = 0u; retry < 20u && copiedSelection.empty(); ++retry)
    {
        PumpPendingMessages();
        copiedSelection = ReadClipboardUnicodeText(prefs);
        if (copiedSelection.empty())
        {
            std::this_thread::sleep_for(20ms);
        }
    }

    state.Require(! copiedSelection.empty(),
                  L"Preferences Plugins Ctrl+C should copy the reordered-resized visible row content after the sort/search round-trip.");
    state.Require(copiedSelection.rfind((expectedTypeText + L"\t"), 0u) == 0u,
                  L"Preferences Plugins clipboard copy should start with the visible Type column after the reordered-resized sort/search round-trip.");
    state.Require(copiedSelection.find(expectedNameText) != std::wstring::npos,
                  L"Preferences Plugins clipboard copy should still include the selected plugin name after the reordered-resized sort/search round-trip.");

    return state.failure.empty();
}
