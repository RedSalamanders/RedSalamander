[[nodiscard]] bool RunFileOpsPathDetailNoActivationRebuildChecks(SelfTestState& state) noexcept
{
    using namespace FileOperationsPopupInternal;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps path detail focus: {}", reason));
        return false;
    };
    RedSalamander::TestSupport::ScopedWindowActivationBlocker activationBlocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! activationBlocker.Start())
        return reject(L"no-activation window blocker could not be acquired");

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = folder ? TryGetFileOps(folder) : nullptr;
    if (! folder || ! operations || operations->HasActiveOperations() || operations->GetPopupHwndForSelfTest() || operations->GetIssuesPaneHwndForSelfTest())
        return reject(L"fixture requires an idle File Operations state with no existing popup or Issues pane");

    const bool previousFooterOnly = operations->GetPopupFooterOnly();
    operations->SetPopupFooterOnly(false);
    auto restoreFooterOnly = wil::scope_exit([&]() noexcept
    {
        operations->SetPopupFooterOnly(previousFooterOnly);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(10000u));
    });
    auto hostLease         = std::make_shared<int>(1);
    wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, hostLease));
    if (! popup)
        return reject(L"synthetic popup could not be created");
    auto destroyPopup = wil::scope_exit([&]() noexcept
    {
        popup.reset();
        hostLease.reset();
    });

    auto* popupState = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA));
    if (! popupState)
        return reject(L"synthetic popup state is unavailable");

    constexpr uint64_t taskId       = 0xF26A202610010021ull;
    constexpr uint64_t secondTaskId = 0xF26A202610010022ull;
    TaskSnapshot task{};
    task.taskId                       = taskId;
    task.operation                    = FILESYSTEM_COPY;
    task.lifecyclePhase               = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
    task.started                      = true;
    task.discoveryClosed              = true;
    task.hasProgressCallbacks         = true;
    task.lastProgressCallbackTick     = GetTickCount64();
    task.operationStartTick           = task.lastProgressCallbackTick > 250u ? task.lastProgressCallbackTick - 250u : 0u;
    task.currentSourcePath            = L"C:\\Users\\Émile\\Archives\\Photo-שלום-📷\\original.raw";
    task.currentDestinationPath       = L"D:\\Backup\\Archives\\Photo-שלום-📷\\original.raw";
    task.totalBytes                   = 4096u;
    task.completedBytes               = 2048u;
    task.totalItems                   = 2u;
    task.completedItems               = 1u;
    TaskSnapshot secondTask           = task;
    secondTask.taskId                 = secondTaskId;
    secondTask.currentSourcePath      = L"E:\\Other\\second.raw";
    secondTask.currentDestinationPath = L"F:\\Other\\second.raw";
    const std::wstring detailRowId    = std::format(L"FileOperations.Task.{}.Path.Source.0", taskId);
    const std::wstring textId         = L"FileOperations.PathDetail.Text";
    const std::wstring copyId         = L"FileOperations.PathDetail.Copy";
    const std::wstring closeId        = L"FileOperations.PathDetail.Close";

    popupState->DebugSetVisualScenario({task}, false, true);
    const HWND focusBefore      = GetFocus();
    const HWND foregroundBefore = GetForegroundWindow();
    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    if (! foregroundBefore || foregroundBefore == popup.get() || IsChild(popup.get(), foregroundBefore) != FALSE || focusBefore == popup.get() ||
        (focusBefore && IsChild(popup.get(), focusBefore) != FALSE) || GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
        return reject(L"showing the synthetic popup changed native focus or foreground ownership");

    // Exercise only the ownership predicate with injected observations. No SetFocus or
    // SetForegroundWindow call is made by this matrix.
    wil::unique_hwnd ownershipChild(CreateWindowExW(0, L"STATIC", L"", WS_CHILD, 0, 0, 1, 1, popup.get(), nullptr, GetModuleHandleW(nullptr), nullptr));
    if (! ownershipChild)
        return reject(L"could not create the hidden child handle used by the focus-policy matrix");
    if (! DebugCanRestoreFileOperationsPathDetailTextFocus(popup.get(), popup.get(), popup.get()) ||
        ! DebugCanRestoreFileOperationsPathDetailTextFocus(popup.get(), ownershipChild.get(), popup.get()) ||
        ! DebugCanRestoreFileOperationsPathDetailTextFocus(popup.get(), popup.get(), ownershipChild.get()) ||
        DebugCanRestoreFileOperationsPathDetailTextFocus(popup.get(), state.mainWindow, popup.get()) ||
        DebugCanRestoreFileOperationsPathDetailTextFocus(popup.get(), popup.get(), state.mainWindow) ||
        DebugCanRestoreFileOperationsPathDetailTextFocus(popup.get(), state.mainWindow, state.mainWindow) ||
        DebugCanRestoreFileOperationsPathDetailTextFocus(popup.get(), nullptr, popup.get()) ||
        DebugCanRestoreFileOperationsPathDetailTextFocus(popup.get(), popup.get(), nullptr))
        return reject(L"text-focus ownership matrix accepted a missing or external focus/foreground observation");

    popupState->DebugOpenPathDetail(detailRowId);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    auto* detail = popupState->DebugPathDetailField();
    if (! detail || detail->GetText() != task.currentSourcePath || popupState->DebugPathDetailLogicalFocusForSelfTest() != textId)
        return reject(L"opening path detail did not retain its exact text target and full path");

    const auto verifyAccessibleText = [&]() noexcept
    {
        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> root;
        root.attach(DxUi::CreateWindowHostAccessibilityProvider(popup.get()));
        wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
        if (! root || FAILED(root.query_to(fragment.put())) || ! fragment)
            return false;
        std::vector<wil::com_ptr_nothrow<IRawElementProviderFragment>> pending;
        pending.push_back(std::move(fragment));
        size_t visited = 0u;
        while (! pending.empty() && ++visited <= 512u)
        {
            auto current = std::move(pending.back());
            pending.pop_back();
            wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
            wil::unique_variant automationId;
            if (FAILED(current.query_to(simple.put())) || ! simple || FAILED(simple->GetPropertyValue(UIA_AutomationIdPropertyId, &automationId)))
                return false;
            if (automationId.vt == VT_BSTR && automationId.bstrVal && textId == std::wstring_view(automationId.bstrVal, SysStringLen(automationId.bstrVal)))
            {
                wil::unique_variant controlType;
                wil::unique_variant isOffscreen;
                wil::com_ptr_nothrow<IUnknown> textPattern;
                UiaRect bounds{};
                return SUCCEEDED(simple->GetPropertyValue(UIA_ControlTypePropertyId, &controlType)) && controlType.vt == VT_I4 &&
                       controlType.lVal == UIA_EditControlTypeId && SUCCEEDED(simple->GetPropertyValue(UIA_IsOffscreenPropertyId, &isOffscreen)) &&
                       isOffscreen.vt == VT_BOOL && isOffscreen.boolVal == VARIANT_FALSE && SUCCEEDED(current->get_BoundingRectangle(&bounds)) &&
                       bounds.width > 0.0 && bounds.height > 0.0 && SUCCEEDED(simple->GetPatternProvider(UIA_TextPatternId, textPattern.put())) && textPattern;
            }
            wil::com_ptr_nothrow<IRawElementProviderFragment> child;
            if (FAILED(current->Navigate(NavigateDirection_FirstChild, child.put())))
                return false;
            size_t siblingCount = 0u;
            while (child)
            {
                if (++siblingCount > 128u || visited + pending.size() >= 512u)
                    return false;
                wil::com_ptr_nothrow<IRawElementProviderFragment> next;
                if (FAILED(child->Navigate(NavigateDirection_NextSibling, next.put())))
                    return false;
                pending.push_back(std::move(child));
                child = std::move(next);
            }
        }
        return false;
    };
    if (! verifyAccessibleText())
        return reject(L"the visible detail editor lost its UIA edit/TextPattern bounds");

    const std::array<std::wstring, 3u> logicalTargets{{textId, copyId, closeId}};
    for (size_t index = 0u; index < logicalTargets.size(); ++index)
    {
        detail = popupState->DebugPathDetailField();
        if (! detail)
            return reject(L"path detail disappeared before its reconstruction cycle");
        const size_t selectionAnchor = task.currentSourcePath.size() / 3u;
        const size_t caret           = task.currentSourcePath.size() * 2u / 3u;
        detail->SetSelectionRange(caret, selectionAnchor);
        const auto selectionBefore = detail->GetSelectionRange();
        const size_t caretBefore   = detail->GetCaretIndex();
        if (! selectionBefore || ! popupState->DebugSetPathDetailLogicalFocusForSelfTest(popup.get(), logicalTargets[index]))
            return reject(L"test could not establish a valid logical Text, Copy, or Close focus target");

        // Alternate one and two task rows to force a retained-tree reconstruction while
        // keeping the inspected path descriptor itself unchanged.
        popupState->DebugSetVisualScenario({task}, false, true);
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        popupState->DebugSetVisualScenario({task, secondTask}, false, true);
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        detail = popupState->DebugPathDetailField();
        if (! detail || detail->GetText() != task.currentSourcePath || detail->GetSelectionRange() != selectionBefore ||
            detail->GetCaretIndex() != caretBefore || popupState->DebugPathDetailLogicalFocusForSelfTest() != logicalTargets[index])
            return reject(L"reconstruction changed the path text, selection, caret, or logical target");

        PopupLayoutDebugSnapshot layout{};
        layout.taskId = taskId;
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout))
            return reject(L"reconstructed popup layout snapshot is unavailable");
        if (index != 0u && layout.hostedFocusedAutomationId != logicalTargets[index])
            return reject(L"background reconstruction failed to restore logical Copy or Close focus");
        if (! verifyAccessibleText() || GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
            return reject(L"background reconstruction changed native ownership or lost the accessible detail editor");

        const UINT dpi = GetDpiForWindow(popup.get());
        RECT desired{0, 0, MulDiv(index == 1u ? 420 : 520, static_cast<int>(dpi), 96), MulDiv(index == 1u ? 300 : 380, static_cast<int>(dpi), 96)};
        if (AdjustWindowRectExForDpi(&desired,
                                     static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_STYLE)),
                                     FALSE,
                                     static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_EXSTYLE)),
                                     dpi) == FALSE ||
            SetWindowPos(popup.get(), nullptr, 0, 0, desired.right - desired.left, desired.bottom - desired.top, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE) ==
                FALSE)
            return reject(L"no-activation resize failed while path detail was open");
        SendMessageW(popup.get(), WM_THEMECHANGED, 0, 0);
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        detail = popupState->DebugPathDetailField();
        if (! detail || detail->GetText() != task.currentSourcePath || detail->GetSelectionRange() != selectionBefore ||
            detail->GetCaretIndex() != caretBefore || GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore || ! verifyAccessibleText())
            return reject(L"resize/theme update damaged detail state or changed native focus ownership");
    }

    constexpr LPARAM repeatedKeyData = (LPARAM{1} << 30) | 1;
    for (const UINT message : std::array<UINT, 2>{{WM_KEYDOWN, WM_SYSKEYDOWN}})
    {
        SendMessageW(popup.get(), message, VK_ESCAPE, repeatedKeyData);
        SendMessageW(popup.get(), message == WM_KEYDOWN ? WM_KEYUP : WM_SYSKEYUP, VK_ESCAPE, 0);
        if (! popupState->DebugPathDetailField() || GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
            return reject(L"a held Escape dismissed detail or changed native ownership");
    }
    SendMessageW(popup.get(), WM_KEYDOWN, VK_ESCAPE, 0);
    SendMessageW(popup.get(), WM_KEYUP, VK_ESCAPE, 0);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PopupLayoutDebugSnapshot closed{};
    closed.taskId = taskId;
    if (popupState->DebugPathDetailField() || popupState->DebugPathDetailLogicalFocusForSelfTest().size() != 0u ||
        ! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), closed) || closed.hostedFocusedAutomationId != detailRowId || ! IsWindowVisible(popup.get()) ||
        GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore || verifyAccessibleText())
        return reject(L"Escape did not close path detail safely while preserving background ownership");

    // A held Escape that dismissed detail must not immediately hide its parent.
    SendMessageW(popup.get(), WM_KEYDOWN, VK_ESCAPE, repeatedKeyData);
    SendMessageW(popup.get(), WM_KEYUP, VK_ESCAPE, 0);
    if (! IsWindowVisible(popup.get()) || GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
        return reject(L"a held detail Escape reached the parent hide action or changed native ownership");

    // Exercise the actual HWND/control dispatch against newly rebound B controls.
    // The synthetic task has no engine worker; the hosted callback counter observes
    // dispatch before the engine intentionally rejects that synthetic task identity.
    task.conflict.active                        = true;
    task.conflict.bucket                        = static_cast<uint8_t>(FileOperations::ConflictClass::RegularFileExists);
    task.conflict.generation                    = 101u;
    task.conflict.sourcePath                    = task.currentSourcePath;
    task.conflict.destinationPath               = task.currentDestinationPath;
    task.conflict.destinationMetadata.available = true;
    popupState->DebugSetVisualScenario({task}, false, true);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PopupLayoutDebugSnapshot bindingA{};
    bindingA.taskId = taskId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), bindingA) || bindingA.hostedDefaultAutomationId.empty() ||
        bindingA.hostedCancelAutomationId.empty() || ! popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), bindingA.hostedDefaultAutomationId))
        return reject(L"generation A did not publish focusable default and cancel controls");

    task.conflict.generation = 102u;
    popupState->DebugSetVisualScenario({task}, false, true);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PopupLayoutDebugSnapshot bindingB{};
    TaskSnapshot projectedB{};
    bindingB.taskId = taskId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), bindingB) || ! DebugGetFileOperationsPopupTaskSnapshot(popup.get(), taskId, projectedB))
        return reject(L"generation B's controls or prompt are unavailable after reconstruction");
    const auto actionId = [&](uint8_t action)
    {
        return std::format(
            L"FileOperations.Action.{}.{}.{}.{}", static_cast<uint32_t>(PopupHitTest::Kind::TaskConflictAction), taskId, action, task.conflict.generation);
    };
    if (projectedB.conflict.generation != task.conflict.generation || bindingB.hostedDefaultAutomationId != actionId(projectedB.conflict.defaultAction) ||
        bindingB.hostedCancelAutomationId != actionId(projectedB.conflict.escapeAction) ||
        bindingB.hostedDefaultAutomationId == bindingA.hostedDefaultAutomationId || bindingB.hostedCancelAutomationId == bindingA.hostedCancelAutomationId)
        return reject(L"default/cancel bindings were not replaced with exact generation-B identities");

    for (const UINT message : std::array<UINT, 2>{{WM_KEYDOWN, WM_SYSKEYDOWN}})
    {
        for (const UINT key : std::array<UINT, 3>{{VK_RETURN, VK_SPACE, VK_ESCAPE}})
        {
            const bool focusAction          = key == VK_SPACE || (message == WM_SYSKEYDOWN && key == VK_RETURN);
            const std::wstring_view focusId = focusAction ? std::wstring_view{bindingB.hostedDefaultAutomationId} : std::wstring_view{};
            if (! popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), focusId))
                return reject(L"could not establish B's focused/default/cancel keyboard route without activation");
            const uint64_t before = popupState->DebugHostedActionCountForSelfTest();
            SendMessageW(popup.get(), message, key, repeatedKeyData);
            SendMessageW(popup.get(), message == WM_KEYDOWN ? WM_KEYUP : WM_SYSKEYUP, key, 0);
            if (popupState->DebugHostedActionCountForSelfTest() != before || GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
                return reject(L"held Enter/Space/Escape dispatched B's action or changed native ownership");

            // DxUi has no SYSKEYDOWN cancel route. System Escape's consumer route
            // is separately exercised above while path detail remains open.
            if (message == WM_SYSKEYDOWN && key == VK_ESCAPE)
                continue;
            // Positive control: a fresh press still invokes that same current B route.
            SendMessageW(popup.get(), message, key, 1);
            SendMessageW(popup.get(), message == WM_KEYDOWN ? WM_KEYUP : WM_SYSKEYUP, key, 0);
            if (popupState->DebugHostedActionCountForSelfTest() != before + 1u || GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
                return reject(L"a fresh Enter/Space/Escape failed to invoke exactly one B action without activation");
        }
    }

    // Repeated task navigation remains available through the same HWND route.
    if (! popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), bindingB.hostedDefaultAutomationId))
        return reject(L"could not focus B's task action before repeated navigation");
    const uint64_t beforeNavigation = popupState->DebugHostedActionCountForSelfTest();
    // Cancel is physically rightmost. Held Left must navigate to its neighbor;
    // held Right must return to Cancel without invoking either action.
    SendMessageW(popup.get(), WM_KEYDOWN, VK_LEFT, repeatedKeyData);
    SendMessageW(popup.get(), WM_KEYUP, VK_LEFT, 0);
    PopupLayoutDebugSnapshot navigated{};
    navigated.taskId = taskId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), navigated) || navigated.hostedFocusedAutomationId.empty() ||
        navigated.hostedFocusedAutomationId == bindingB.hostedDefaultAutomationId || popupState->DebugHostedActionCountForSelfTest() != beforeNavigation ||
        GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
        return reject(L"repeat suppression swallowed Left navigation from rightmost Cancel, invoked an action, or changed native ownership");
    SendMessageW(popup.get(), WM_KEYDOWN, VK_RIGHT, repeatedKeyData);
    SendMessageW(popup.get(), WM_KEYUP, VK_RIGHT, 0);
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), navigated) ||
        navigated.hostedFocusedAutomationId != bindingB.hostedDefaultAutomationId ||
        popupState->DebugHostedActionCountForSelfTest() != beforeNavigation ||
        GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
        return reject(L"held Right did not return to rightmost Cancel without invoking an action or changing native ownership");

    return true;
}

// Focus on a task-owned path region is focus on that task: Escape/Enter must never
// bind to another task's sole published decision (popup-b#2). Routed messages only.
[[nodiscard]] bool RunFileOpsTaskOwnedRegionFocusEscapeChecks(SelfTestState& state) noexcept
{
    using namespace FileOperationsPopupInternal;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps task-owned focus Escape: {}", reason));
        return false;
    };
    RedSalamander::TestSupport::ScopedWindowActivationBlocker activationBlocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! activationBlocker.Start())
        return reject(L"no-activation window blocker could not be acquired");

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = folder ? TryGetFileOps(folder) : nullptr;
    if (! folder || ! operations || operations->HasActiveOperations() || operations->GetPopupHwndForSelfTest() || operations->GetIssuesPaneHwndForSelfTest())
        return reject(L"fixture requires an idle File Operations state with no existing popup or Issues pane");

    const bool previousFooterOnly = operations->GetPopupFooterOnly();
    operations->SetPopupFooterOnly(false);
    auto restoreFooterOnly = wil::scope_exit([&]() noexcept
    {
        operations->SetPopupFooterOnly(previousFooterOnly);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(10000u));
    });
    auto hostLease         = std::make_shared<int>(1);
    wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, hostLease));
    if (! popup)
        return reject(L"synthetic popup could not be created");
    auto destroyPopup = wil::scope_exit([&]() noexcept
    {
        popup.reset();
        hostLease.reset();
    });
    auto* popupState  = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA));
    if (! popupState)
        return reject(L"synthetic popup state is unavailable");

    // Task A copies without a decision; task B holds the popup's only actionable decision.
    constexpr uint64_t taskAId = 0xF26A202610080021ull;
    constexpr uint64_t taskBId = 0xF26A202610080022ull;
    TaskSnapshot taskA{};
    taskA.taskId                                 = taskAId;
    taskA.operation                              = FILESYSTEM_COPY;
    taskA.lifecyclePhase                         = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
    taskA.started                                = true;
    taskA.discoveryClosed                        = true;
    taskA.hasProgressCallbacks                   = true;
    taskA.lastProgressCallbackTick               = GetTickCount64();
    taskA.operationStartTick                     = taskA.lastProgressCallbackTick > 250u ? taskA.lastProgressCallbackTick - 250u : 0u;
    taskA.currentSourcePath                      = L"C:\\Escape\\TaskA\\source.raw";
    taskA.currentDestinationPath                 = L"D:\\Escape\\TaskA\\source.raw";
    taskA.totalBytes                             = 4096u;
    taskA.completedBytes                         = 2048u;
    taskA.totalItems                             = 2u;
    taskA.completedItems                         = 1u;
    TaskSnapshot taskB                           = taskA;
    taskB.taskId                                 = taskBId;
    taskB.currentSourcePath                      = L"E:\\Escape\\TaskB\\second.raw";
    taskB.currentDestinationPath                 = L"F:\\Escape\\TaskB\\second.raw";
    taskB.conflict.active                        = true;
    taskB.conflict.bucket                        = static_cast<uint8_t>(FileOperations::ConflictClass::RegularFileExists);
    taskB.conflict.generation                    = 201u;
    taskB.conflict.sourcePath                    = taskB.currentSourcePath;
    taskB.conflict.destinationPath               = taskB.currentDestinationPath;
    taskB.conflict.destinationMetadata.available = true;

    popupState->DebugSetVisualScenario({taskA, taskB}, false, true);
    const HWND focusBefore      = GetFocus();
    const HWND foregroundBefore = GetForegroundWindow();
    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
    // Keep both cards, A's path rows and B's decision buttons inside the viewport.
    const UINT dpi = GetDpiForWindow(popup.get());
    RECT desired{0, 0, MulDiv(560, static_cast<int>(dpi), 96), MulDiv(760, static_cast<int>(dpi), 96)};
    if (AdjustWindowRectExForDpi(&desired,
                                 static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_STYLE)),
                                 FALSE,
                                 static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_EXSTYLE)),
                                 dpi) == FALSE ||
        SetWindowPos(popup.get(), nullptr, 0, 0, desired.right - desired.left, desired.bottom - desired.top, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE) ==
            FALSE)
        return reject(L"no-activation resize of the synthetic popup failed");
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    if (GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
        return reject(L"showing the synthetic popup changed native focus or foreground ownership");

    TaskSnapshot projectedB{};
    if (! DebugGetFileOperationsPopupTaskSnapshot(popup.get(), taskBId, projectedB) || ! projectedB.conflict.active ||
        ! projectedB.conflict.buttonsPublishable || projectedB.conflict.generation != taskB.conflict.generation)
        return reject(L"task B did not publish an actionable decision");
    const std::wstring escapeIdB = std::format(L"FileOperations.Action.{}.{}.{}.{}",
                                               static_cast<uint32_t>(PopupHitTest::Kind::TaskConflictAction),
                                               taskBId,
                                               projectedB.conflict.escapeAction,
                                               projectedB.conflict.generation);
    const auto pressEscape       = [&]() noexcept
    {
        SendMessageW(popup.get(), WM_KEYDOWN, VK_ESCAPE, 1);
        SendMessageW(popup.get(), WM_KEYUP, VK_ESCAPE, 0);
    };

    // Control: with global/no focus, the sole actionable decision (B) owns Escape.
    if (! popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), {}))
        return reject(L"could not clear hosted focus for the global-focus control");
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PopupLayoutDebugSnapshot globalBinding{};
    globalBinding.taskId = taskBId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), globalBinding) || globalBinding.hostedCancelAutomationId != escapeIdB ||
        globalBinding.hostedDefaultAutomationId.empty())
        return reject(std::format(L"global focus did not bind Enter/Escape to B's sole decision (default='{}', cancel='{}', expected '{}', "
                                  L"focused='{}', actions={}, primary={}, overflow={})",
                                  globalBinding.hostedDefaultAutomationId,
                                  globalBinding.hostedCancelAutomationId,
                                  escapeIdB,
                                  globalBinding.hostedFocusedAutomationId,
                                  globalBinding.hostedActionControlCount,
                                  globalBinding.conflictPrimaryActionCount,
                                  globalBinding.conflictOverflowActionCount));
    const uint64_t beforeGlobal = popupState->DebugHostedActionCountForSelfTest();
    pressEscape();
    if (popupState->DebugHostedActionCountForSelfTest() != beforeGlobal + 1u || ! IsWindowVisible(popup.get()) || GetForegroundWindow() != foregroundBefore ||
        GetFocus() != focusBefore)
        return reject(L"Escape with global focus did not invoke exactly B's decision without hiding or activation");

    // Focus A's own path region, as returning from its path detail does.
    const std::wstring pathIdA = std::format(L"FileOperations.Task.{}.Path.Source.0", taskAId);
    if (! popupState->DebugSetHostedPathRegionFocusForSelfTest(popup.get(), pathIdA))
        return reject(L"task A's source path region is not a focusable hosted region");
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PopupLayoutDebugSnapshot regionBinding{};
    regionBinding.taskId = taskBId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), regionBinding) || regionBinding.hostedFocusedAutomationId != pathIdA)
        return reject(L"task A's path region did not keep hosted focus across synchronization");
    if (! regionBinding.hostedCancelAutomationId.empty() || ! regionBinding.hostedDefaultAutomationId.empty() ||
        popupState->DebugConflictEscapeTaskIdForSelfTest() != 0u)
        return reject(std::format(L"focus on A's path region bound Enter/Escape to another task's decision (default='{}', cancel='{}', escapeTask={})",
                                  regionBinding.hostedDefaultAutomationId,
                                  regionBinding.hostedCancelAutomationId,
                                  popupState->DebugConflictEscapeTaskIdForSelfTest()));

    // The first Escape may only dismiss the region's focus tooltip; the next hides the popup.
    const auto escapeHidesWithoutAction = [&](std::wstring_view context) noexcept
    {
        const uint64_t before = popupState->DebugHostedActionCountForSelfTest();
        for (size_t press = 0u; press < 2u && IsWindowVisible(popup.get()); ++press)
        {
            pressEscape();
            if (popupState->DebugHostedActionCountForSelfTest() != before)
                return reject(std::format(L"Escape {} submitted task B's decision", context));
        }
        if (IsWindowVisible(popup.get()) || GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
            return reject(std::format(L"Escape {} did not hide the popup, or changed native ownership", context));
        return true;
    };
    const auto reshow = [&]() noexcept
    {
        ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        return IsWindowVisible(popup.get()) && GetForegroundWindow() == foregroundBefore && GetFocus() == focusBefore;
    };
    if (! escapeHidesWithoutAction(L"on A's path region"))
        return false;

    // A's region can vanish while A stays present (stream slots, a momentarily empty
    // path). Focus keeps A's identity through the safe fallback and never turns global.
    if (! reshow())
        return reject(L"re-showing the synthetic popup failed or changed native ownership");
    if (! popupState->DebugSetHostedPathRegionFocusForSelfTest(popup.get(), pathIdA))
        return reject(L"task A's source path region could not be focused again");
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    TaskSnapshot taskAWithoutSource = taskA;
    taskAWithoutSource.currentSourcePath.clear();
    popupState->DebugSetVisualScenario({taskAWithoutSource, taskB}, false, true);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    const std::wstring toggleIdA = std::format(L"FileOperations.Action.{}.{}.0", static_cast<uint32_t>(PopupHitTest::Kind::TaskToggleCollapse), taskAId);
    PopupLayoutDebugSnapshot vanishedBinding{};
    vanishedBinding.taskId = taskBId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), vanishedBinding) || vanishedBinding.hostedFocusedAutomationId != toggleIdA)
        return reject(
            std::format(L"focus did not fall back to A's disclosure when A's path region vanished (focused='{}')", vanishedBinding.hostedFocusedAutomationId));
    if (! vanishedBinding.hostedCancelAutomationId.empty() || ! vanishedBinding.hostedDefaultAutomationId.empty() ||
        popupState->DebugConflictEscapeTaskIdForSelfTest() != 0u)
        return reject(std::format(L"A's vanished path region bound Enter/Escape to another task's decision (default='{}', cancel='{}', escapeTask={})",
                                  vanishedBinding.hostedDefaultAutomationId,
                                  vanishedBinding.hostedCancelAutomationId,
                                  popupState->DebugConflictEscapeTaskIdForSelfTest()));
    if (! escapeHidesWithoutAction(L"after A's focused path region vanished"))
        return false;

    // Closing A's path detail returns focus to A's region before the next paint. The
    // binding synchronized while detail was open (global, so B's) must not survive it.
    popupState->DebugSetVisualScenario({taskA, taskB}, false, true);
    if (! reshow() || ! popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), {}))
        return reject(L"re-showing the synthetic popup with global focus failed or changed native ownership");
    popupState->DebugOpenPathDetail(pathIdA);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    if (! popupState->DebugPathDetailField() || popupState->DebugConflictEscapeTaskIdForSelfTest() != taskBId)
        return reject(L"fixture precondition: open path detail did not leave B's escape decision synchronized");
    pressEscape();
    PopupLayoutDebugSnapshot closedBinding{};
    closedBinding.taskId = taskBId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), closedBinding) || closedBinding.hostedFocusedAutomationId != pathIdA ||
        ! IsWindowVisible(popup.get()))
        return reject(L"Escape did not close A's path detail back to A's path region");
    if (popupState->DebugConflictEscapeTaskIdForSelfTest() != 0u || ! closedBinding.hostedCancelAutomationId.empty())
        return reject(std::format(L"closing A's path detail kept B's escape decision bound (escapeTask={}, cancel='{}')",
                                  popupState->DebugConflictEscapeTaskIdForSelfTest(),
                                  closedBinding.hostedCancelAutomationId));
    // No paint runs between these routed presses, so only the reset protects B.
    if (! escapeHidesWithoutAction(L"right after closing A's path detail"))
        return false;

    // Detail closes itself when its region vanishes (a stream slot ends, a terminal card).
    // Its return focus keeps A's identity through A's disclosure, never global Options.
    popupState->DebugSetVisualScenario({taskA, taskB}, false, true);
    if (! reshow() || ! popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), {}))
        return reject(L"re-showing the synthetic popup for the vanished-detail case failed or changed native ownership");
    popupState->DebugOpenPathDetail(pathIdA);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    if (! popupState->DebugPathDetailField())
        return reject(L"fixture precondition: A's path detail did not open");
    popupState->DebugSetVisualScenario({taskAWithoutSource, taskB}, false, true);
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    if (popupState->DebugPathDetailField() || ! IsWindowVisible(popup.get()))
        return reject(L"A's path detail did not close when its region vanished");
    // The next synchronization re-derives Enter/Escape from the restored focus.
    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    PopupLayoutDebugSnapshot detailVanishedBinding{};
    detailVanishedBinding.taskId = taskBId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), detailVanishedBinding) || detailVanishedBinding.hostedFocusedAutomationId != toggleIdA)
        return reject(std::format(L"closing A's path detail with its region gone did not return focus to A's disclosure (focused='{}')",
                                  detailVanishedBinding.hostedFocusedAutomationId));
    if (! detailVanishedBinding.hostedCancelAutomationId.empty() || ! detailVanishedBinding.hostedDefaultAutomationId.empty() ||
        popupState->DebugConflictEscapeTaskIdForSelfTest() != 0u)
        return reject(std::format(L"A's vanished path detail bound Enter/Escape to another task's decision (default='{}', cancel='{}', escapeTask={})",
                                  detailVanishedBinding.hostedDefaultAutomationId,
                                  detailVanishedBinding.hostedCancelAutomationId,
                                  popupState->DebugConflictEscapeTaskIdForSelfTest()));
    return escapeHidesWithoutAction(L"after A's path detail closed because its region vanished");
}

// I30 x-ux#1, x-ux#7, popup-a#2, popup-b#4: screen readers hear Canceled, Outcome unknown and a result
// whose card first appears already Failed or Partial (also on a new popup's first synchronization),
// but not an interrupted-Move restart notice or a first-seen Done or Canceled card; the footer counts
// paused tasks in their own Paused clause; path rows are invokable Buttons named by label and path;
// Space activates a focused Cancel or Skip button instead of pausing, and a Ctrl/Alt chord is neither.
// Every part runs and logs before the verdict. Routed messages only; no activation.
[[nodiscard]] bool RunFileOpsPopupStatusAccessibilityChecks(SelfTestState& state) noexcept
{
    using namespace FileOperationsPopupInternal;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps popup status accessibility: {}", reason));
        return false;
    };
    RedSalamander::TestSupport::ScopedWindowActivationBlocker activationBlocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! activationBlocker.Start())
        return reject(L"no-activation window blocker could not be acquired");

    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = folder ? TryGetFileOps(folder) : nullptr;
    if (! folder || ! operations || operations->HasActiveOperations() || operations->GetPopupHwndForSelfTest() || operations->GetIssuesPaneHwndForSelfTest())
        return reject(L"fixture requires an idle File Operations state with no existing popup or Issues pane");

    const bool previousFooterOnly = operations->GetPopupFooterOnly();
    operations->SetPopupFooterOnly(false);
    auto restoreFooterOnly = wil::scope_exit([&]() noexcept
    {
        operations->SetPopupFooterOnly(previousFooterOnly);
        static_cast<void>(operations->DebugFlushPendingSettingsSaveForSelfTest(10000u));
    });
    auto hostLease         = std::make_shared<int>(1);
    wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, hostLease));
    if (! popup)
        return reject(L"synthetic popup could not be created");
    auto destroyPopup = wil::scope_exit([&]() noexcept
    {
        popup.reset();
        hostLease.reset();
    });
    auto* popupState  = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA));
    if (! popupState)
        return reject(L"synthetic popup state is unavailable");

    using Status                = TaskSnapshot::StatusKind;
    using Phase                 = FileOperations::TaskLifecyclePhase;
    constexpr uint64_t runId    = 0xF26A202610100041ull;
    constexpr uint64_t failId   = 0xF26A202610100042ull;
    constexpr uint64_t noticeId = 0xF26A202610100043ull;
    constexpr uint64_t doneId   = 0xF26A202610100044ull;
    constexpr uint64_t unknId   = 0xF26A202610100045ull;
    constexpr uint64_t lateId   = 0xF26A202610100046ull;
    constexpr uint64_t cancId   = 0xF26A202610100047ull;
    constexpr uint64_t partId   = 0xF26A202610100048ull;
    constexpr uint64_t skipId   = 0xF26A202610100049ull;
    TaskSnapshot running{};
    running.taskId                   = runId;
    running.operation                = FILESYSTEM_COPY;
    running.lifecyclePhase           = static_cast<uint8_t>(Phase::Running);
    running.statusKind               = Status::Running;
    running.started                  = true;
    running.discoveryClosed          = true;
    running.hasProgressCallbacks     = true;
    running.lastProgressCallbackTick = GetTickCount64();
    running.operationStartTick       = running.lastProgressCallbackTick > 250u ? running.lastProgressCallbackTick - 250u : 0u;
    running.currentSourcePath        = L"C:\\Status\\Running\\source.raw";
    running.currentDestinationPath   = L"D:\\Status\\Running\\source.raw";
    running.totalBytes               = 4096u;
    running.completedBytes           = 2048u;
    running.totalItems               = 2u;
    running.completedItems           = 1u;
    const auto finishedAs            = [&](uint64_t taskId, Status status, HRESULT hr) noexcept
    {
        TaskSnapshot task      = running;
        task.taskId            = taskId;
        task.lifecyclePhase    = static_cast<uint8_t>(Phase::Terminal);
        task.statusKind        = status;
        task.finished          = true;
        task.resultHr          = hr;
        task.currentSourcePath = std::format(L"C:\\Status\\{}\\source.raw", taskId);
        task.currentDestinationPath.clear();
        task.errorCount             = status == Status::Failed ? 1u : 0u;
        task.warningCount           = status == Status::Partial ? 1u : 0u;
        task.hasIndeterminateResult = status == Status::Indeterminate;
        return task;
    };

    const HWND focusBefore      = GetFocus();
    const HWND foregroundBefore = GetForegroundWindow();
    const auto present          = [&](std::vector<TaskSnapshot> tasks) noexcept
    {
        popupState->DebugSetVisualScenario(std::move(tasks), false, true);
        ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    };
    const auto layoutFor = [&](uint64_t taskId, PopupLayoutDebugSnapshot& layout) noexcept
    {
        layout        = {};
        layout.taskId = taskId;
        return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout);
    };
    const auto notifications = [&]() noexcept -> std::optional<uint64_t>
    {
        // The snapshot needs a present task; the fast failure stays in every announcement step.
        PopupLayoutDebugSnapshot layout{};
        if (! layoutFor(failId, layout))
            return std::nullopt;
        return layout.hostedAccessibilityNotificationCount;
    };

    // x-ux#1. The popup's first synchronization shows a Failed card that ended before its reveal, an
    // interrupted-Move restart notice and a Done card: only the failure is announced.
    const auto announcements = [&]() noexcept
    {
        // Not read before the first synchronization: the new popup has announced nothing yet.
        const std::optional<uint64_t> initial = uint64_t{0};
        TaskSnapshot notice                   = finishedAs(noticeId, Status::Indeterminate, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE));
        notice.operation                      = FILESYSTEM_MOVE;
        notice.interruptedMoveNotice          = true;
        notice.interruptedOperationId         = 41u;
        std::vector<TaskSnapshot> tasks{finishedAs(failId, Status::Failed, E_ACCESSDENIED), notice, finishedAs(doneId, Status::Done, S_OK)};
        present(tasks);
        const UINT dpi = GetDpiForWindow(popup.get());
        RECT desired{0, 0, MulDiv(560, static_cast<int>(dpi), 96), MulDiv(900, static_cast<int>(dpi), 96)};
        if (AdjustWindowRectExForDpi(&desired,
                                     static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_STYLE)),
                                     FALSE,
                                     static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_EXSTYLE)),
                                     dpi) == FALSE ||
            SetWindowPos(popup.get(), nullptr, 0, 0, desired.right - desired.left, desired.bottom - desired.top, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE) ==
                FALSE)
            return reject(L"no-activation resize of the synthetic popup failed");
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        std::optional<uint64_t> previous = notifications();
        bool passed                      = true;
        const auto expect                = [&](std::wstring_view step, uint64_t expected) noexcept
        {
            const std::optional<uint64_t> now = notifications();
            const bool matched                = previous && now && *now == *previous + expected;
            AppendLog(std::format(L"FileOps popup status accessibility: {} raised {} notification(s), expected {}",
                                  step,
                                  previous && now ? *now - *previous : uint64_t{0},
                                  expected));
            if (! matched)
                passed = reject(std::format(L"x-ux#1: {} raised the wrong number of announcements", step));
            previous = now;
        };
        if (! previous)
            return reject(L"x-ux#1: the notification counter is unavailable");
        previous = initial;
        expect(L"the first synchronization with a fast Failed card, a restart notice and a Done card", 1u);

        tasks.push_back(running);
        present(tasks);
        expect(L"a running card appearing", 0u);
        TaskSnapshot& live         = tasks.back();
        live.cancellationRequested = true;
        live.lifecyclePhase        = static_cast<uint8_t>(Phase::Stopping);
        present(tasks);
        expect(L"Running to Stopping", 1u);
        tasks.back() = finishedAs(runId, Status::Canceled, HRESULT_FROM_WIN32(ERROR_CANCELLED));
        present(tasks);
        expect(L"Stopping to Canceled", 1u);

        TaskSnapshot unknown = running;
        unknown.taskId       = unknId;
        tasks.push_back(unknown);
        present(tasks);
        expect(L"a second running card appearing", 0u);
        tasks.back() = finishedAs(unknId, Status::Indeterminate, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE));
        present(tasks);
        expect(L"Running to Outcome unknown", 1u);

        tasks.push_back(finishedAs(lateId, Status::Failed, E_ACCESSDENIED));
        present(tasks);
        expect(L"a card first appearing Failed in a synchronized popup", 1u);
        tasks.push_back(finishedAs(cancId, Status::Canceled, HRESULT_FROM_WIN32(ERROR_CANCELLED)));
        tasks.push_back(finishedAs(partId, Status::Partial, HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)));
        present(tasks);
        expect(L"first-seen Canceled and Partial cards", 1u);
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        present(tasks);
        expect(L"repaints without a state change", 0u);
        return passed;
    };

    // x-ux#7. Paused tasks have their own clause and are not counted as running.
    const auto footer = [&]() noexcept
    {
        TaskSnapshot paused          = running;
        paused.taskId                = failId;
        paused.statusKind            = Status::Paused;
        paused.paused                = true;
        const std::wstring rc        = FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_RUNNING_COUNT, 1ul);
        const std::wstring pc        = FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_PAUSED_COUNT, 1ul);
        const std::wstring separator = LoadStringResource(nullptr, IDS_FILEOPS_GLOBAL_STATUS_SEPARATOR);
        bool passed                  = ! pc.empty() && pc != rc;
        if (! passed)
            passed = reject(L"x-ux#7: the Paused clause resource is missing");
        PopupLayoutDebugSnapshot layout{};
        present({paused});
        if (! layoutFor(failId, layout) || ! layout.globalSummaryText.starts_with(pc) ||
            layout.globalSummaryText.find(FormatStringResource(nullptr, IDS_FMT_FILEOPS_GLOBAL_RUNNING_COUNT, 1ul)) != std::wstring::npos)
            passed = reject(std::format(L"x-ux#7: a paused-only footer reads '{}', expected '{}'", layout.globalSummaryText, pc));
        present({running, paused});
        const std::wstring both = rc + separator + pc;
        if (! layoutFor(runId, layout) || ! layout.globalSummaryText.starts_with(both))
            passed = reject(std::format(L"x-ux#7: a running and a paused task read '{}', expected '{}'", layout.globalSummaryText, both));
        return passed;
    };

    // popup-a#2. A path row is an invokable Button named by its label and full path.
    const auto pathRows = [&]() noexcept
    {
        present({running});
        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        const std::wstring rowId = std::format(L"FileOperations.Task.{}.Path.Source.0", runId);
        wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> root;
        root.attach(DxUi::CreateWindowHostAccessibilityProvider(popup.get()));
        wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
        if (! root || FAILED(root.query_to(fragment.put())) || ! fragment)
            return reject(L"popup-a#2: the popup has no UIA root");
        std::vector<wil::com_ptr_nothrow<IRawElementProviderFragment>> pending;
        pending.push_back(std::move(fragment));
        size_t visited = 0u;
        while (! pending.empty() && ++visited <= 512u)
        {
            auto current = std::move(pending.back());
            pending.pop_back();
            wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
            wil::unique_variant automationId;
            if (FAILED(current.query_to(simple.put())) || ! simple || FAILED(simple->GetPropertyValue(UIA_AutomationIdPropertyId, &automationId)))
                return reject(L"popup-a#2: a UIA element could not be read");
            if (automationId.vt == VT_BSTR && automationId.bstrVal && rowId == std::wstring_view(automationId.bstrVal, SysStringLen(automationId.bstrVal)))
            {
                wil::unique_variant controlType;
                wil::unique_variant focusable;
                wil::unique_variant name;
                wil::com_ptr_nothrow<IUnknown> invoke;
                const bool read             = SUCCEEDED(simple->GetPropertyValue(UIA_ControlTypePropertyId, &controlType)) && controlType.vt == VT_I4 &&
                                              SUCCEEDED(simple->GetPropertyValue(UIA_IsKeyboardFocusablePropertyId, &focusable)) && focusable.vt == VT_BOOL &&
                                              SUCCEEDED(simple->GetPropertyValue(UIA_NamePropertyId, &name)) && name.vt == VT_BSTR && name.bstrVal;
                const std::wstring nameText = read ? std::wstring(name.bstrVal, SysStringLen(name.bstrVal)) : std::wstring{};
                const bool invokable        = SUCCEEDED(simple->GetPatternProvider(UIA_InvokePatternId, invoke.put())) && invoke;
                AppendLog(std::format(L"FileOps popup status accessibility: path row control type {} focusable {} invokable {} name '{}'",
                                      read ? controlType.lVal : 0,
                                      read && focusable.boolVal == VARIANT_TRUE,
                                      invokable,
                                      nameText));
                if (! read || controlType.lVal != UIA_ButtonControlTypeId || focusable.boolVal != VARIANT_TRUE || ! invokable ||
                    ! nameText.ends_with(running.currentSourcePath) || nameText.size() <= running.currentSourcePath.size() + 1u)
                    return reject(L"popup-a#2: the source path row is not a focusable, invokable Button named by its label and path");
                return true;
            }
            wil::com_ptr_nothrow<IRawElementProviderFragment> child;
            if (FAILED(current->Navigate(NavigateDirection_FirstChild, child.put())))
                return reject(L"popup-a#2: UIA navigation failed");
            size_t siblingCount = 0u;
            while (child)
            {
                if (++siblingCount > 128u || visited + pending.size() >= 512u)
                    return reject(L"popup-a#2: the UIA tree exceeded its traversal bound");
                wil::com_ptr_nothrow<IRawElementProviderFragment> next;
                if (FAILED(child->Navigate(NavigateDirection_NextSibling, next.put())))
                    return reject(L"popup-a#2: UIA navigation failed");
                pending.push_back(std::move(child));
                child = std::move(next);
            }
        }
        return reject(L"popup-a#2: the source path row is missing from the UIA tree");
    };

    // popup-b#4. Routing matrix, then routed key messages on retained buttons.
    const auto spaceKeys = [&]() noexcept
    {
        using Kind  = PopupHitTest::Kind;
        using Route = HostedTaskSpaceKeyRoute;
        struct RouteCase final
        {
            Kind kind;
            bool enabled;
            bool modified;
            Route expected;
        };
        constexpr std::array<RouteCase, 13> routes{{
            {Kind::TaskCancel, true, false, Route::ControlActivation},
            {Kind::TaskSkip, true, false, Route::ControlActivation},
            {Kind::TaskSpeedLimit, true, false, Route::ControlActivation},
            {Kind::TaskStartNow, true, false, Route::ControlActivation},
            {Kind::TaskQueueMoveUp, true, false, Route::ControlActivation},
            {Kind::TaskPause, true, false, Route::PauseOrResume},
            {Kind::TaskCancel, false, false, Route::PauseOrResume},
            {Kind::TaskCancel, true, true, Route::Ignore},
            {Kind::TaskPause, true, true, Route::Ignore},
            {Kind::TaskCancel, false, true, Route::Ignore},
            {Kind::TaskToggleCollapse, true, true, Route::ControlActivation},
            {Kind::TaskConflictAction, true, false, Route::ControlActivation},
            {Kind::TaskConflictToggleApplyToAll, true, true, Route::ControlActivation},
        }};
        bool passed = true;
        for (const RouteCase& route : routes)
        {
            const Route actual = ResolveHostedTaskSpaceKeyRoute(route.kind, route.enabled, route.modified);
            if (actual != route.expected)
                passed = reject(std::format(L"popup-b#4: Space on kind {} (enabled {}, chord {}) routes {}, expected {}",
                                            static_cast<uint32_t>(route.kind),
                                            route.enabled,
                                            route.modified,
                                            static_cast<uint32_t>(actual),
                                            static_cast<uint32_t>(route.expected)));
        }

        TaskSnapshot discovering           = running;
        discovering.taskId                 = skipId;
        discovering.discoveryAheadActive   = true;
        discovering.discoveryClosed        = false;
        discovering.currentSourcePath      = L"C:\\Status\\Discovering\\source.raw";
        discovering.discoveredFileCount    = 12u;
        discovering.discoveredTotalBytes   = 8192u;
        discovering.currentDestinationPath = L"D:\\Status\\Discovering\\source.raw";
        present({running, discovering});
        const auto actionId = [](Kind kind, uint64_t taskId) { return std::format(L"FileOperations.Action.{}.{}.0", static_cast<uint32_t>(kind), taskId); };
        const auto press    = [&](std::wstring_view id, UINT message, LPARAM keyData, bool ctrl) noexcept -> std::optional<uint64_t>
        {
            if (! popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), id))
                return std::nullopt;
            std::array<BYTE, 256> saved{};
            if (ctrl && GetKeyboardState(saved.data()) == FALSE)
                return std::nullopt;
            auto restoreKeys = wil::scope_exit([&]() noexcept
            {
                if (ctrl)
                    static_cast<void>(SetKeyboardState(saved.data()));
            });
            if (ctrl)
            {
                std::array<BYTE, 256> chord = saved;
                chord[VK_CONTROL]           = 0x80u;
                chord[VK_LCONTROL]          = 0x80u;
                if (SetKeyboardState(chord.data()) == FALSE)
                    return std::nullopt;
            }
            const uint64_t before = popupState->DebugHostedActionCountForSelfTest();
            SendMessageW(popup.get(), message, VK_SPACE, keyData);
            SendMessageW(popup.get(), message == WM_SYSKEYDOWN ? WM_SYSKEYUP : WM_KEYUP, VK_SPACE, LPARAM{1} | (LPARAM{1} << 30) | (LPARAM{1} << 31));
            return popupState->DebugHostedActionCountForSelfTest() - before;
        };
        const auto expectPress = [&](std::wstring_view what, std::wstring_view id, UINT message, LPARAM keyData, bool ctrl, uint64_t expected) noexcept
        {
            const std::optional<uint64_t> actions = press(id, message, keyData, ctrl);
            AppendLog(std::format(L"FileOps popup status accessibility: {} ran {} hosted action(s), expected {}",
                                  what,
                                  actions.has_value() ? std::format(L"{}", *actions) : std::wstring(L"no"),
                                  expected));
            if (! actions || *actions != expected)
                passed = reject(std::format(L"popup-b#4: {} did not run exactly {} hosted action(s)", what, expected));
        };
        constexpr LPARAM altContext = LPARAM{1} | (LPARAM{1} << 29);
        const std::wstring cancelId = actionId(Kind::TaskCancel, runId);
        PopupLayoutDebugSnapshot discoveryLayout{};
        if (! layoutFor(skipId, discoveryLayout) || ! discoveryLayout.taskSkipVisible)
            passed = reject(L"popup-b#4: fixture precondition: the discovering task shows no Skip button");
        else
            expectPress(L"Space on a focused Skip", actionId(Kind::TaskSkip, skipId), WM_KEYDOWN, 1, false, 1u);
        expectPress(L"Space on a focused Cancel", cancelId, WM_KEYDOWN, 1, false, 1u);
        expectPress(L"Ctrl+Space on a focused Cancel", cancelId, WM_KEYDOWN, 1, true, 0u);
        expectPress(L"Alt+Space on a focused Cancel", cancelId, WM_SYSKEYDOWN, altContext, false, 0u);
        expectPress(L"Space on a focused Cancel after the chords", cancelId, WM_KEYDOWN, 1, false, 1u);
        return passed;
    };

    const bool announced  = announcements();
    const bool footerOk   = footer();
    const bool pathRowsOk = pathRows();
    const bool keysOk     = spaceKeys();
    if (GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
        return reject(L"the scenario changed native focus or foreground ownership");
    return announced && footerOk && pathRowsOk && keysOk;
}

// I30 popup-b#1/#3/#5, x-perf#4/#7 (owner decision Q7): the retained hosted tree is reconciled by
// identity, never rebuilt while its controls survive. A copy whose active stream count changes on
// every tick keeps one Source/Destination pair and a held button press still clicks; scrolling never
// pulls focus back or moves it to a global control; a conflict body without comparison children is
// stable; live text and scan counters do not republish or remeasure on every frame. Every check runs
// and logs its counters before the verdict, so a failing run also records the baseline evidence.
// Routed messages only; no activation.
[[nodiscard]] bool RunFileOpsPopupStableHostedTreeChecks(SelfTestState& state) noexcept
{
    using namespace FileOperationsPopupInternal;
    using namespace std::chrono_literals;
    bool passed     = true;
    const auto fail = [&passed](std::wstring_view reason)
    {
        AppendLog(std::format(L"Popup stable hosted tree: FAILED: {}", reason));
        passed = false;
    };
    const auto note = [](std::wstring_view text) { AppendLog(std::format(L"Popup stable hosted tree: {}", text)); };

    auto* folder = TryGetFolderWindow(state.mainWindow);
    if (! folder)
    {
        fail(L"the live folder window is unavailable");
        return false;
    }
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
    {
        fail(L"could not acquire the no-activation window guard");
        return false;
    }

    // A private engine keeps the application's popup and tasks out of this witness.
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    auto owner = std::make_unique<FolderWindow>();
    owner->SetSettings(&settings);
    FolderWindow::FileOperationState operations(*owner);
    if (operations.GetPopupFooterOnly())
    {
        fail(L"the private engine must default to the card list, not footer-only mode");
        return false;
    }
    auto lifetime = std::make_shared<int>(0);
    wil::unique_hwnd popup(FileOperationsPopup::Create(&operations, folder, state.mainWindow, lifetime));
    auto closePopup  = wil::scope_exit([&]() noexcept
    {
        if (popup && GetCapture() == popup.get())
            ReleaseCapture();
        // Private popup: expire its host token before WM_NCDESTROY, as the terminal-facts witness does.
        lifetime.reset();
        popup.reset();
    });
    auto* popupState = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA)) : nullptr;
    if (! popupState)
    {
        fail(L"could not create the private production popup");
        return false;
    }

    const HWND focusBefore      = GetFocus();
    const HWND foregroundBefore = GetForegroundWindow();
    const UINT dpi              = GetDpiForWindow(popup.get());
    const auto resizeTo         = [&](int widthDip, int heightDip) noexcept
    {
        RECT desired{0, 0, MulDiv(widthDip, static_cast<int>(dpi), 96), MulDiv(heightDip, static_cast<int>(dpi), 96)};
        static_cast<void>(AdjustWindowRectExForDpi(&desired,
                                                   static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_STYLE)),
                                                   FALSE,
                                                   static_cast<DWORD>(GetWindowLongPtrW(popup.get(), GWL_EXSTYLE)),
                                                   dpi));
        SetWindowPos(popup.get(), nullptr, 0, 0, desired.right - desired.left, desired.bottom - desired.top, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    };
    const auto paint    = [&]() noexcept { RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN); };
    const auto counters = [&]() noexcept { return popupState->DebugHostedTreeCountersForSelfTest(); };
    const auto actionId = [](PopupHitTest::Kind kind, uint64_t taskId)
    { return std::format(L"FileOperations.Action.{}.{}.0", static_cast<uint32_t>(kind), taskId); };
    const auto findButton = [&](PopupHitTest::Kind kind, uint64_t taskId) noexcept -> std::optional<D2D1_RECT_F>
    {
        for (const PopupButton& button : popupState->DebugVisualButtons())
            if (button.hit.kind == kind && button.hit.taskId == taskId)
                return button.bounds;
        return std::nullopt;
    };
    const auto centerOf = [](const D2D1_RECT_F& bounds) noexcept
    {
        return MAKELPARAM(static_cast<int>(std::lround((bounds.left + bounds.right) * 0.5f)),
                          static_cast<int>(std::lround((bounds.top + bounds.bottom) * 0.5f)));
    };
    const auto layoutFor = [&](uint64_t taskId) noexcept
    {
        PopupLayoutDebugSnapshot layout{};
        layout.taskId = taskId;
        static_cast<void>(DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout));
        return layout;
    };
    const auto makeCopy = [](uint64_t taskId, std::wstring_view root)
    {
        TaskSnapshot task{};
        task.taskId                   = taskId;
        task.operation                = FILESYSTEM_COPY;
        task.lifecyclePhase           = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
        task.started                  = true;
        task.discoveryClosed          = true;
        task.hasProgressCallbacks     = true;
        task.lastProgressCallbackTick = GetTickCount64();
        task.operationStartTick       = task.lastProgressCallbackTick > 250u ? task.lastProgressCallbackTick - 250u : 0u;
        task.currentSourcePath        = std::format(L"{}\\Source\\first.bin", root);
        task.currentDestinationPath   = std::format(L"{}\\Destination\\first.bin", root);
        task.destinationFolder        = std::format(L"{}\\Destination", root);
        task.totalBytes               = 64ull << 20;
        task.completedBytes           = 16ull << 20;
        task.totalItems               = 400u;
        task.completedItems           = 100u;
        return task;
    };
    // One active in-flight row per stream; the generation changes every path.
    const auto setStreams = [](TaskSnapshot& task, size_t count, size_t generation)
    {
        task.autoConcurrencyUsed        = true;
        task.autoTunedConcurrency       = static_cast<unsigned int>(count);
        task.effectiveConcurrencyBudget = 8u;
        task.inFlightFileCount          = count;
        for (size_t index = 0u; index < count; ++index)
        {
            auto& entry          = task.inFlightFiles[index];
            entry                = {};
            entry.sourcePath     = std::format(L"C:\\Stable\\Streams\\file-{}-{}.bin", generation, index);
            entry.totalBytes     = 1ull << 20;
            entry.completedBytes = 256ull << 10;
            entry.lastUpdateTick = GetTickCount64();
        }
    };
    // Perf evidence: one fixed workload is timed with each tree strategy in this build. The baseline
    // sets debugRebuildOnStructuralChange (the whole tree replaced on every structural change, as
    // before popup-b#1); the candidate reconciles. Only synchronizations that changed the tree count.
    struct WorkloadTiming final
    {
        size_t structuralFrames  = 0u;
        uint64_t structuralUs    = 0u;
        uint64_t maxStructuralUs = 0u;
        uint64_t created         = 0u;
    };
    const auto timedPaint = [&](WorkloadTiming& timing) noexcept
    {
        const auto before = counters();
        paint();
        const auto after = counters();
        timing.created += after.controlsCreated - before.controlsCreated;
        if (after.reconciles == before.reconciles) // A rebuild reconciles too.
            return;
        const uint64_t us = (after.rebuildUs + after.reconcileUs) - (before.rebuildUs + before.reconcileUs);
        ++timing.structuralFrames;
        timing.structuralUs += us;
        timing.maxStructuralUs = std::max(timing.maxStructuralUs, us);
    };
    const auto describeTiming = [](const WorkloadTiming& timing)
    {
        return std::format(L"structuralFrames={} totalUs={} meanUs={} maxUs={} created={}",
                           timing.structuralFrames,
                           timing.structuralUs,
                           timing.structuralFrames == 0u ? 0u : timing.structuralUs / timing.structuralFrames,
                           timing.maxStructuralUs,
                           timing.created);
    };

    // 1. Q7 / popup-b#1: a multi-stream copy publishes one stable source region whose help text names
    // the other active streams, and a press held while the active stream count changes on every tick
    // still clicks: the hosted tree is never rebuilt and the pressed and hovered control survives.
    // A second card appears and disappears on alternate ticks, so the tree is also reconciled (and
    // that card's controls retired) while the press is held.
    constexpr uint64_t copyId = 0xF26A2026100A0011ull;
    TaskSnapshot copy         = makeCopy(copyId, L"C:\\Stable");
    setStreams(copy, 4u, 0u);
    popupState->DebugSetVisualScenario({copy}, false, true);
    popupState->debugFixedGeometry = true;
    resizeTo(760, 900);
    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
    paint();
    paint();
    {
        const std::wstring sourcePrefix = std::format(L"FileOperations.Task.{}.Path.Source", copyId);
        const size_t sourceRegions      = popupState->DebugHostedRegionCountWithPrefixForSelfTest(sourcePrefix);
        std::wstring sourceName;
        std::wstring sourceHelp;
        const bool sourceRead = popupState->DebugHostedRegionAccessibleTextForSelfTest(sourcePrefix + L".0", sourceName, sourceHelp);
        bool helpNamesStreams = sourceRead && sourceName.find(copy.inFlightFiles[0].sourcePath) != std::wstring::npos;
        for (size_t index = 1u; index < copy.inFlightFileCount; ++index)
            helpNamesStreams = helpNamesStreams && sourceHelp.find(copy.inFlightFiles[index].sourcePath) != std::wstring::npos;
        note(std::format(L"Q7 stable pair: sourceRegions={} read={} helpNamesStreams={} name='{}'", sourceRegions, sourceRead, helpNamesStreams, sourceName));
        if (sourceRegions != 1u || ! helpNamesStreams)
            fail(std::format(L"a 4-stream copy published {} source regions; Q7 requires one stable region whose help text names the other streams",
                             sourceRegions));
    }

    PopupHitTest::Kind pressKind           = PopupHitTest::Kind::TaskPause;
    std::optional<D2D1_RECT_F> pressBounds = findButton(pressKind, copyId);
    if (! pressBounds)
    {
        pressKind   = PopupHitTest::Kind::TaskCancel;
        pressBounds = findButton(pressKind, copyId);
    }
    if (! pressBounds)
    {
        fail(L"the running copy published neither Pause nor Cancel");
        return false;
    }
    const std::wstring pressId    = actionId(pressKind, copyId);
    const uintptr_t pressIdentity = popupState->DebugHostedControlIdentityForSelfTest(pressId);
    const auto churnBefore        = counters();
    const uint64_t actionsBefore  = popupState->DebugHostedActionCountForSelfTest();
    SendMessageW(popup.get(), WM_MOUSEMOVE, 0, centerOf(*pressBounds));
    SendMessageW(popup.get(), WM_LBUTTONDOWN, MK_LBUTTON, centerOf(*pressBounds));
    bool pressSurvived      = pressIdentity != 0u;
    bool pressHovered       = popupState->DebugVisualButtonHovered(pressKind);
    const auto churnStarted = std::chrono::steady_clock::now();
    TaskSnapshot other      = makeCopy(copyId + 5u, L"C:\\Stable\\Other");
    setStreams(other, 1u, 0u);
    const auto churnScenario = [&](size_t tick)
    {
        setStreams(copy, 1u + tick % 4u, tick);
        std::vector<TaskSnapshot> churnTasks{copy};
        if (tick % 2u == 1u)
            churnTasks.push_back(other);
        popupState->DebugSetVisualScenario(std::move(churnTasks), false, true);
    };
    WorkloadTiming churnReconcile{};
    for (size_t tick = 1u; tick <= 12u; ++tick)
    {
        churnScenario(tick);
        {
            Debug::Perf::Scope frame(L"FileOps.Popup.StableTree.StreamChurnFrameUs");
            timedPaint(churnReconcile);
        }
        pressSurvived = pressSurvived && popupState->DebugHostedControlIdentityForSelfTest(pressId) == pressIdentity;
        pressHovered  = pressHovered && popupState->DebugVisualButtonHovered(pressKind);
    }
    const uint64_t churnUs = Debug::Perf::ElapsedUs(churnStarted);
    // The card height follows the stream count; release over the button's current position.
    const std::optional<D2D1_RECT_F> releaseBounds = findButton(pressKind, copyId);
    SendMessageW(popup.get(), WM_LBUTTONUP, 0, centerOf(releaseBounds.value_or(*pressBounds)));
    const auto churnAfter = counters();
    const uint64_t clicks = popupState->DebugHostedActionCountForSelfTest() - actionsBefore;
    Debug::Perf::Emit(L"FileOps.Popup.StableTree.StreamChurn",
                      L"12 frames, 1-4 active streams, press held",
                      churnUs,
                      churnAfter.rebuilds - churnBefore.rebuilds,
                      churnAfter.accessibleTextUpdates - churnBefore.accessibleTextUpdates);
    note(std::format(L"stream churn: clicks={} survived={} hovered={} syncs={} rebuilds={} reconciles={} created={} revived={} retired={} "
                     L"accessibleUpdates={} deferred={} frameTotalUs={}",
                     clicks,
                     pressSurvived,
                     pressHovered,
                     churnAfter.syncs - churnBefore.syncs,
                     churnAfter.rebuilds - churnBefore.rebuilds,
                     churnAfter.reconciles - churnBefore.reconciles,
                     churnAfter.controlsCreated - churnBefore.controlsCreated,
                     churnAfter.controlsRevived - churnBefore.controlsRevived,
                     churnAfter.controlsRetired - churnBefore.controlsRetired,
                     churnAfter.accessibleTextUpdates - churnBefore.accessibleTextUpdates,
                     churnAfter.accessibleTextDeferrals - churnBefore.accessibleTextDeferrals,
                     churnUs));
    if (clicks != 1u)
        fail(std::format(L"a press held while the active stream count changed did not click exactly once (clicks={})", clicks));
    if (churnAfter.rebuilds != churnBefore.rebuilds)
        fail(std::format(L"stream-count changes rebuilt the whole hosted tree {} times", churnAfter.rebuilds - churnBefore.rebuilds));
    if (! pressSurvived || ! pressHovered)
        fail(L"the pressed control was recreated or lost its hover while only stream rows and another card changed");
    if (churnAfter.reconciles == churnBefore.reconciles || churnAfter.controlsRetired == churnBefore.controlsRetired)
        fail(L"fixture: the appearing and disappearing card never reconciled or retired a hosted control");
    if (churnAfter.controlsRevived == churnBefore.controlsRevived)
        fail(L"the card that reappeared while the press was held was recreated instead of shown again");
    if (GetCapture() == popup.get())
        fail(L"the popup kept mouse capture after the release");
    // The controls retired while the press was held (the last tick dropped the second card) leave
    // the tree at the first synchronization after the release, and leaving does not count them again.
    const size_t retainedAtRelease     = popupState->DebugRetiredHostedControlCountForSelfTest();
    const size_t bodyChildrenAtRelease = popupState->DebugHostedBodyChildCountForSelfTest();
    paint();
    const auto settled               = counters();
    const size_t retainedAfterSettle = popupState->DebugRetiredHostedControlCountForSelfTest();
    const size_t bodyChildrenSettled = popupState->DebugHostedBodyChildCountForSelfTest();
    note(std::format(L"after release: retained={}->{} bodyChildren={}->{} retiredCounted={}->{} reconciles={}",
                     retainedAtRelease,
                     retainedAfterSettle,
                     bodyChildrenAtRelease,
                     bodyChildrenSettled,
                     churnAfter.controlsRetired - churnBefore.controlsRetired,
                     settled.controlsRetired - churnBefore.controlsRetired,
                     settled.reconciles - churnAfter.reconciles));
    if (retainedAtRelease == 0u)
        fail(L"fixture: no control was retired while the press was held");
    if (retainedAfterSettle != 0u || bodyChildrenSettled >= bodyChildrenAtRelease)
        fail(L"controls retired while the press was held stayed in the tree after it ended");
    if (settled.controlsRetired != churnAfter.controlsRetired)
        fail(std::format(L"removing the controls retired during the press counted them again ({} more)", settled.controlsRetired - churnAfter.controlsRetired));

    // Perf baseline for the same workload: the press is held while the same 12 frames are replayed
    // with the whole tree rebuilt on every structural change. The press cannot survive that strategy
    // (it was the High defect); only the timing is recorded.
    setStreams(copy, 4u, 0u);
    popupState->DebugSetVisualScenario({copy}, false, true);
    paint();
    paint();
    WorkloadTiming churnRebuild{};
    popupState->debugRebuildOnStructuralChange = true;
    SendMessageW(popup.get(), WM_MOUSEMOVE, 0, centerOf(*pressBounds));
    SendMessageW(popup.get(), WM_LBUTTONDOWN, MK_LBUTTON, centerOf(*pressBounds));
    for (size_t tick = 1u; tick <= 12u; ++tick)
    {
        churnScenario(tick);
        timedPaint(churnRebuild);
    }
    SendMessageW(popup.get(), WM_LBUTTONUP, 0, centerOf(findButton(pressKind, copyId).value_or(*pressBounds)));
    popupState->debugRebuildOnStructuralChange = false;
    Debug::Perf::Emit(L"FileOps.Popup.StableTree.ChurnSyncUs",
                      L"rebuild baseline: 12 frames, held press, 1-4 streams, second card on odd frames",
                      churnRebuild.structuralUs,
                      churnRebuild.structuralFrames,
                      churnRebuild.maxStructuralUs);
    Debug::Perf::Emit(L"FileOps.Popup.StableTree.ChurnSyncUs",
                      L"reconcile: 12 frames, held press, 1-4 streams, second card on odd frames",
                      churnReconcile.structuralUs,
                      churnReconcile.structuralFrames,
                      churnReconcile.maxStructuralUs);
    note(std::format(L"stream churn timing: rebuild baseline {}; reconcile {}", describeTiming(churnRebuild), describeTiming(churnReconcile)));
    if (churnReconcile.created >= churnRebuild.created)
        fail(std::format(L"reconciling created {} controls for the workload a rebuild serves with {}", churnReconcile.created, churnRebuild.created));

    // 2. popup-b#3: scrolling never undoes itself through a focus restore and never moves keyboard
    // focus to a global control. The focused card keeps its hosted controls off screen.
    std::vector<TaskSnapshot> cards;
    for (uint64_t index = 0u; index < 3u; ++index)
    {
        TaskSnapshot card = makeCopy(copyId + 1u + index, std::format(L"C:\\Stable\\Card{}", index));
        setStreams(card, 1u, index);
        cards.push_back(std::move(card));
    }
    const uint64_t firstCardId = cards.front().taskId;
    popupState->DebugSetVisualScenario(cards, false, true);
    resizeTo(560, 360);
    paint();
    paint();
    const std::wstring firstToggleId = actionId(PopupHitTest::Kind::TaskToggleCollapse, firstCardId);
    if (! popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), firstToggleId))
        fail(L"fixture: the first card's disclosure is not a focusable hosted control");
    paint();
    const auto scrollBefore = counters();
    WorkloadTiming scrollReconcile{};
    for (size_t line = 0u; line < 3u; ++line)
        SendMessageW(popup.get(), WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0), 0);
    const int lineScrolled = layoutFor(firstCardId).scrollPosition;
    timedPaint(scrollReconcile);
    timedPaint(scrollReconcile);
    const PopupLayoutDebugSnapshot afterLines = layoutFor(firstCardId);
    SendMessageW(popup.get(), WM_VSCROLL, MAKEWPARAM(SB_BOTTOM, 0), 0);
    const int bottomScrolled = layoutFor(firstCardId).scrollPosition;
    timedPaint(scrollReconcile);
    timedPaint(scrollReconcile);
    const PopupLayoutDebugSnapshot afterBottom = layoutFor(firstCardId);
    const auto scrollAfter                     = counters();
    note(std::format(L"scroll: lines={}->{} bottom={}->{} focusedAfterLines='{}' focusedAfterBottom='{}' firstToggleHosted={} rebuilds={} "
                     L"reconciles={}",
                     lineScrolled,
                     afterLines.scrollPosition,
                     bottomScrolled,
                     afterBottom.scrollPosition,
                     afterLines.hostedFocusedAutomationId,
                     afterBottom.hostedFocusedAutomationId,
                     popupState->DebugHostedHitCountForSelfTest(PopupHitTest::Kind::TaskToggleCollapse, firstCardId),
                     scrollAfter.rebuilds - scrollBefore.rebuilds,
                     scrollAfter.reconciles - scrollBefore.reconciles));
    if (lineScrolled == 0 || bottomScrolled <= lineScrolled)
        fail(L"fixture: the three-card list did not scroll");
    if (afterLines.scrollPosition != lineScrolled || afterBottom.scrollPosition != bottomScrolled)
        fail(L"a paint after scrolling moved the list back to the focused control");
    if (afterLines.hostedFocusedAutomationId != firstToggleId || afterBottom.hostedFocusedAutomationId != firstToggleId)
        fail(std::format(L"scrolling moved keyboard focus from the first card's disclosure to '{}'", afterBottom.hostedFocusedAutomationId));
    if (scrollAfter.rebuilds != scrollBefore.rebuilds)
        fail(L"scrolling rebuilt the whole hosted tree");
    // Perf baseline for the same scroll from the top with the same focus, rebuilt on every
    // structural change. A rebuild restores focus with a reveal, so its later positions can differ.
    SendMessageW(popup.get(), WM_VSCROLL, MAKEWPARAM(SB_TOP, 0), 0);
    paint();
    paint();
    static_cast<void>(popupState->DebugSetHostedButtonFocusForSelfTest(popup.get(), firstToggleId));
    paint();
    WorkloadTiming scrollRebuild{};
    popupState->debugRebuildOnStructuralChange = true;
    for (size_t line = 0u; line < 3u; ++line)
        SendMessageW(popup.get(), WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0), 0);
    timedPaint(scrollRebuild);
    timedPaint(scrollRebuild);
    SendMessageW(popup.get(), WM_VSCROLL, MAKEWPARAM(SB_BOTTOM, 0), 0);
    timedPaint(scrollRebuild);
    timedPaint(scrollRebuild);
    popupState->debugRebuildOnStructuralChange = false;
    Debug::Perf::Emit(L"FileOps.Popup.StableTree.ScrollSyncUs",
                      L"rebuild baseline: 3 cards, 3 lines then bottom, focus on the first card",
                      scrollRebuild.structuralUs,
                      scrollRebuild.structuralFrames,
                      scrollRebuild.maxStructuralUs);
    Debug::Perf::Emit(L"FileOps.Popup.StableTree.ScrollSyncUs",
                      L"reconcile: 3 cards, 3 lines then bottom, focus on the first card",
                      scrollReconcile.structuralUs,
                      scrollReconcile.structuralFrames,
                      scrollReconcile.maxStructuralUs);
    note(std::format(L"scroll timing: rebuild baseline {}; reconcile {}", describeTiming(scrollRebuild), describeTiming(scrollReconcile)));

    // 3. popup-b#5: a visible conflict body with no comparison children (no paths, no last note)
    // reaches a stable hosted tree; repainting it changes nothing.
    TaskSnapshot bare = makeCopy(copyId + 10u, L"C:\\Stable\\Bare");
    bare.currentSourcePath.clear();
    bare.currentDestinationPath.clear();
    bare.conflict.active     = true;
    bare.conflict.generation = 501u;
    bare.conflict.bucket     = static_cast<uint8_t>(FileOperations::ConflictClass::RegularFileExists);
    bare.conflict.status     = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    popupState->DebugSetVisualScenario({bare}, false, true);
    resizeTo(760, 900);
    paint();
    paint();
    const auto bareBefore = counters();
    for (size_t pass = 0u; pass < 4u; ++pass)
        paint();
    const auto bareAfter = counters();
    note(std::format(L"bare conflict body: visible={} syncs={} rebuilds={} reconciles={}",
                     popupState->DebugConflictBodyVisibleForSelfTest(bare.taskId),
                     bareAfter.syncs - bareBefore.syncs,
                     bareAfter.rebuilds - bareBefore.rebuilds,
                     bareAfter.reconciles - bareBefore.reconciles));
    if (! popupState->DebugConflictBodyVisibleForSelfTest(bare.taskId))
        fail(L"fixture: the bare conflict body is not visible");
    if (bareAfter.rebuilds != bareBefore.rebuilds || bareAfter.reconciles != bareBefore.reconciles)
        fail(L"repainting a conflict body without comparison children changed the hosted tree on every frame");

    // 4. x-perf#4: live text (paths changing on every telemetry frame) republishes accessible text at
    // most once per control per interval, and the latest value is still published afterwards.
    TaskSnapshot live = makeCopy(copyId + 20u, L"C:\\Stable\\Live");
    setStreams(live, 3u, 0u);
    popupState->DebugSetVisualScenario({live}, false, true);
    paint();
    paint();
    const PopupLayoutDebugSnapshot liveLayout = layoutFor(live.taskId);
    const size_t liveControls = liveLayout.hostedTooltipRegionCount + liveLayout.hostedProgressControlCount + liveLayout.hostedGraphControlCount;
    const auto liveBefore     = counters();
    const auto liveStarted    = std::chrono::steady_clock::now();
    for (size_t tick = 1u; tick <= 10u; ++tick)
    {
        // A live tick, not a new fixture: only the paths change.
        setStreams(live, 3u, 100u + tick);
        popupState->debugVisualScenario = std::vector<TaskSnapshot>{live};
        Debug::Perf::Scope frame(L"FileOps.Popup.StableTree.LiveTextFrameUs");
        paint();
    }
    const uint64_t liveUs  = Debug::Perf::ElapsedUs(liveStarted);
    const auto liveAfter   = counters();
    const uint64_t applied = liveAfter.accessibleTextUpdates - liveBefore.accessibleTextUpdates;
    const uint64_t allowed = static_cast<uint64_t>(liveControls) * (1u + liveUs / 1'000'000u);
    Debug::Perf::Emit(L"FileOps.Popup.StableTree.LiveText", L"10 frames, 3 streams, paths change", liveUs, applied, liveControls);
    note(std::format(L"live text: controls={} applied={} allowed={} deferred={} elapsedUs={}",
                     liveControls,
                     applied,
                     allowed,
                     liveAfter.accessibleTextDeferrals - liveBefore.accessibleTextDeferrals,
                     liveUs));
    if (applied > allowed)
        fail(std::format(L"live text republished accessible text {} times in 10 frames (bound {})", applied, allowed));
    // The popup timer keeps painting while text is pending; the latest path is published within the interval.
    const std::wstring liveSourceId = std::format(L"FileOperations.Task.{}.Path.Source.0", live.taskId);
    const auto latestPublished      = [&]() noexcept
    {
        std::wstring name;
        std::wstring help;
        return popupState->DebugHostedRegionAccessibleTextForSelfTest(liveSourceId, name, help) &&
               name.find(live.inFlightFiles[0].sourcePath) != std::wstring::npos;
    };
    if (! RedSalamander::TestSupport::PumpMessagesUntil(latestPublished, {.timeout = SelfTest::Scale(5000ms), .operationName = L"latest live text publication"})
              .conditionMet)
        fail(L"the latest live source path was never published to UI Automation");
    // A progress bar's Name is its task status, not live text: two status changes within one
    // publication interval each publish on their own frame.
    using StatusChange = std::pair<TaskSnapshot::StatusKind, UINT>;
    for (const auto& [statusKind, statusText] : {StatusChange{TaskSnapshot::StatusKind::Paused, IDS_FILEOPS_GRAPH_PAUSED},
                                                 StatusChange{TaskSnapshot::StatusKind::Stopping, IDS_FILEOPS_STATUS_STOPPING}})
    {
        live.statusKind                 = statusKind;
        live.paused                     = statusKind == TaskSnapshot::StatusKind::Paused;
        popupState->debugVisualScenario = std::vector<TaskSnapshot>{live};
        paint();
        std::wstring published;
        const std::wstring expected = LoadStringResource(nullptr, statusText);
        const bool found            = popupState->DebugHostedProgressAccessibleNameForSelfTest(live.taskId, published);
        note(std::format(L"progress status name: found={} expected='{}' published='{}'", found, expected, published));
        if (! found || published != expected)
            fail(std::format(L"the progress bar published '{}' instead of the new status '{}' on the frame it changed", published, expected));
    }

    // 5. x-perf#7: scan counters change on every frame; the path truncation is measured once.
    TaskSnapshot scan{};
    scan.kind                              = TaskSnapshot::Kind::Informational;
    scan.taskId                            = copyId + 30u;
    scan.informational.taskId              = scan.taskId;
    scan.informational.kind                = FolderWindow::InformationalTaskUpdate::Kind::CompareDirectories;
    scan.informational.title               = L"Compare";
    scan.informational.leftRoot            = L"C:\\Stable\\Left";
    scan.informational.rightRoot           = L"D:\\Stable\\Right";
    scan.informational.scanActive          = true;
    scan.informational.scanFolderCount     = 100u;
    scan.informational.scanEntryCount      = 1000u;
    scan.informational.scanCurrentRelative = L"Stable\\Scan\\current-entry.raw";
    popupState->DebugSetVisualScenario({scan}, false, true);
    paint();
    paint();
    const uint64_t measuredBefore = layoutFor(scan.taskId).scanStatusMeasurementCount;
    bool scanTextCurrent          = true;
    for (uint64_t tick = 1u; tick <= 8u; ++tick)
    {
        scan.informational.scanFolderCount = 100u + tick;
        scan.informational.scanEntryCount  = 1000u + tick;
        popupState->DebugSetVisualScenario({scan}, false, true);
        {
            Debug::Perf::Scope frame(L"FileOps.Popup.StableTree.ScanStatusFrameUs");
            paint();
        }
        const std::wstring expected = FormatStringResource(nullptr,
                                                           IDS_FMT_COMPARE_SCAN_STATUS,
                                                           scan.informational.scanCurrentRelative.native(),
                                                           scan.informational.scanFolderCount,
                                                           scan.informational.scanEntryCount);
        scanTextCurrent             = scanTextCurrent && layoutFor(scan.taskId).taskScanStatusText == expected;
    }
    const uint64_t remeasured = layoutFor(scan.taskId).scanStatusMeasurementCount - measuredBefore;
    Debug::Perf::Emit(L"FileOps.Popup.StableTree.ScanStatus", L"8 frames, counters change, same path", 0u, remeasured, 8u);
    note(std::format(L"scan status: remeasured={} textCurrent={}", remeasured, scanTextCurrent));
    if (! scanTextCurrent)
        fail(L"the scan status row did not show the current counters");
    if (remeasured > 1u)
        fail(std::format(L"changing scan counters remeasured the path truncation {} times in 8 frames", remeasured));

    // 6. popup-b#1 review (High): a press on a path region opens path detail on mouse-down, so the
    // body's ScrollPanel holds the region as its captured (and hovered) child. DxUi can drop the
    // host capture without telling that panel: ControlHost::ReleaseMouseCapture is silent, as is the
    // prune a SetFocusControl runs. (In the pinned DxUi, hiding the captured body happens to notify
    // it, because the accessibility publication re-enters the host with WM_GETTEXT; the popup must
    // not rely on that.) The capture is dropped silently here, the next paint opens detail and hides
    // the body, the region's descriptor vanishes so the region leaves the tree, detail closes, and
    // the pointer moves over the list again. Removed controls are kept alive and detached here, so
    // input the panel still routes to the region is counted instead of reading freed memory. The
    // second pass freezes the region's own hover flag (false) while the panel records it as the
    // hovered child: the panel's pointer must be dropped without trusting that flag.
    for (const bool hoverFlagFrozen : {false, true})
    {
        TaskSnapshot pressedPath = makeCopy(copyId + (hoverFlagFrozen ? 41u : 40u), hoverFlagFrozen ? L"C:\\Stable\\PressedFrozen" : L"C:\\Stable\\Pressed");
        popupState->DebugSetVisualScenario({pressedPath}, false, true);
        paint();
        paint();
        const std::wstring pressedRegionId            = std::format(L"FileOperations.Task.{}.Path.Source.0", pressedPath.taskId);
        const std::optional<D2D1_RECT_F> regionBounds = popupState->DebugHostedRegionBoundsForSelfTest(pressedRegionId);
        const uintptr_t regionIdentity                = popupState->DebugHostedControlIdentityForSelfTest(pressedRegionId);
        if (! regionBounds || regionIdentity == 0u)
            fail(L"fixture: the copy published no hosted source path region");
        else if (hoverFlagFrozen && ! popupState->DebugFreezeHostedRegionHoverFlagForSelfTest(pressedRegionId))
            fail(L"fixture: the source path region's hover flag could not be frozen");
        else
        {
            popupState->DebugSetRetainRetiredControlsForSelfTest(true);
            auto releaseRetained        = wil::scope_exit([&]() noexcept { popupState->DebugSetRetainRetiredControlsForSelfTest(false); });
            const uint64_t routedBefore = popupState->DebugDetachedControlInputCountForSelfTest();
            const LPARAM regionPoint    = centerOf(*regionBounds);
            SendMessageW(popup.get(), WM_MOUSEMOVE, 0, regionPoint);
            SendMessageW(popup.get(), WM_LBUTTONDOWN, MK_LBUTTON, regionPoint);
            const bool capturedOnPress = popupState->DebugHostedCaptureActiveForSelfTest();
            popupState->DebugReleaseHostCaptureSilentlyForSelfTest();
            const bool capturePruned = ! popupState->DebugHostedCaptureActiveForSelfTest();
            paint(); // Detail opens and hides the body before the release.
            const bool detailOpened = popupState->DebugPathDetailField() != nullptr;
            pressedPath.currentSourcePath.clear();
            popupState->DebugSetVisualScenario({pressedPath}, false, true);
            paint(); // The region retires without a capture; detail closes because its row is gone.
            paint();
            const bool regionRemoved =
                popupState->DebugHostedControlIdentityForSelfTest(pressedRegionId) == 0u && popupState->DebugRetiredControlRetainedForSelfTest(regionIdentity);
            const bool detailClosed = popupState->DebugPathDetailField() == nullptr;
            SendMessageW(popup.get(), WM_MOUSEMOVE, 0, regionPoint);
            SendMessageW(popup.get(),
                         WM_MOUSEMOVE,
                         0,
                         centerOf(D2D1::RectF(regionBounds->left + 8.0f, regionBounds->top + 24.0f, regionBounds->right + 8.0f, regionBounds->bottom + 24.0f)));
            SendMessageW(popup.get(), WM_LBUTTONUP, 0, regionPoint);
            const uint64_t routed = popupState->DebugDetachedControlInputCountForSelfTest() - routedBefore;
            note(std::format(L"pressed region removed: hoverFlagFrozen={} capturedOnPress={} capturePruned={} detailOpened={} removed={} detailClosed={} "
                             L"routedAfterRemoval={}",
                             hoverFlagFrozen,
                             capturedOnPress,
                             capturePruned,
                             detailOpened,
                             regionRemoved,
                             detailClosed,
                             routed));
            if (! capturedOnPress || ! capturePruned || ! detailOpened || ! regionRemoved || ! detailClosed)
                fail(L"fixture: the press did not open detail, drop its capture, and remove the region before detail closed");
            if (routed != 0u)
                fail(std::format(L"a ScrollPanel routed {} pointer events to a path region after it left the tree (hoverFlagFrozen={}; use-after-free in "
                                 L"production)",
                                 routed,
                                 hoverFlagFrozen));
        }
    }

    if (GetForegroundWindow() != foregroundBefore || GetFocus() != focusBefore)
        fail(L"the scenario changed native focus or foreground ownership");
    return passed;
}

// Wave 5 conflict inspection on a real Local Copy conflict in the active pane, without activation:
// a refresh of the target folder before its enumeration settles is followed (folderview#1), a user
// cancel of that enumeration ends the request as Canceled with the previous folder restored and
// frees the next request while clearing the pane path or preparing a disconnect ends it Stale
// (folderview#2), the prompt's item kind breaks a same-name tie in the pane and through the host
// request (folderview#4), and destroying the popup or answering the prompt ends a pending request
// as Stale at once (x-lifetime#1).
[[nodiscard]] bool RunFileOpsConflictInspectionSettlementChecks(SelfTestState& state) noexcept
{
    using namespace std::chrono_literals;
    using Status      = FileOperations::ConflictInspectionNavigationStatus;
    using Result      = FileOperations::ConflictInspectionNavigationResult;
    using Task        = FolderWindow::FileOperationState::Task;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps conflict inspection: {}", reason));
        return false;
    };
    RedSalamander::TestSupport::ScopedWindowActivationBlocker activationBlocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! activationBlocker.Start())
        return reject(L"no-activation window blocker could not be acquired");

    constexpr auto pane = FolderWindow::Pane::Left;
    auto* folder        = TryGetFolderWindow(state.mainWindow);
    auto* operations    = folder ? TryGetFileOps(folder) : nullptr;
    auto* view          = TryGetFolderView(folder, pane);
    if (! folder || ! operations || ! view || ! state.fsLocal || operations->HasActiveOperations() || operations->GetPopupHwndForSelfTest())
        return reject(L"fixture requires an idle File Operations state with no popup and the live left pane");

    const auto root      = state.tempRoot / L"conflict-inspection";
    const auto home      = root / L"home";
    const auto sourceDir = root / L"source";
    const auto destDir   = root / L"destination";
    const auto source    = sourceDir / L"item.txt";
    const auto dest      = destDir / L"item.txt";
    std::error_code ec;
    if (! RecreateEmptyDirectory(root) || ! std::filesystem::create_directories(home, ec) || ! std::filesystem::create_directories(sourceDir, ec) ||
        ! std::filesystem::create_directories(destDir, ec) || ! WriteTestFile(source, 64u) || ! WriteTestFile(dest, 32u) ||
        ! WriteTestFile(destDir / L"neighbor.txt", 16u) || ! WriteTestFile(home / L"home.txt", 16u))
        return reject(L"could not seed the home, source and destination folders");

    const auto activeBefore = folder->GetActivePane();
    const auto pathBefore   = folder->GetCurrentPath(pane);
    const std::wstring pluginBefore(folder->GetFileSystemPluginId(pane));
    std::optional<uint64_t> taskId;
    std::optional<Result> result;
    const auto& identity = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem();
    std::optional<FolderView::DebugKindTwinRows> twin;
    std::optional<bool> twinCurrentIsDirectory;
    const auto complete = [&](const Result& value)
    {
        result = value;
        if (twin)
        {
            // Host-level tie-break check: record the row the request made current, then remove
            // the synthetic twin before the pane model can change.
            const auto current     = view->GetFocusedItemDetails();
            twinCurrentIsDirectory = current ? std::optional<bool>(current->isDirectory) : std::nullopt;
            view->DebugRemoveKindTwinRowForSelfTest(twin.value());
            twin.reset();
        }
    };
    const auto wait = [](const auto& condition, std::wstring_view operation) noexcept
    { return RedSalamander::TestSupport::PumpMessagesUntil(condition, {.timeout = SelfTest::Scale(10s), .operationName = operation}).conditionMet; };
    auto restore = wil::scope_exit([&]() noexcept
    {
        folder->SetPaneEnumerationSettledCallback(pane, {});
        operations->CancelPendingConflictInspection();
        if (twin)
            view->DebugRemoveKindTwinRowForSelfTest(twin.value());
        if (taskId)
        {
            if (operations->FindTask(taskId.value()))
                operations->CancelAll();
            static_cast<void>(
                wait([&] { return ! operations->FindTask(taskId.value()) && ! operations->HasActiveOperations(); }, L"inspection Copy retirement"));
            operations->DismissCompletedTask(taskId.value());
            operations->DebugRemoveDiagnosticsForTask(taskId.value());
            state.completedTasks.erase(taskId.value());
        }
        if (operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
        if (! pluginBefore.empty() && ! EqualsIgnoreCase(pluginBefore, kPluginIdLocal))
            static_cast<void>(folder->SetFileSystemPluginForPane(pane, pluginBefore));
        if (pathBefore)
            folder->SetFolderPath(pane, pathBefore.value());
        folder->SetActivePane(activeBefore);
    });
    if (FAILED(folder->SetFileSystemPluginForPane(pane, kPluginIdLocal)))
        return reject(L"the left pane could not switch to the Local provider");
    folder->SetActivePane(pane);
    const auto settleAt = [&](const std::filesystem::path& path) noexcept
    {
        folder->SetFolderPath(pane, path);
        return wait([&] { return CurrentPanePathEquals(folder, pane, path) && view->IsCurrentFolderEnumerated(); }, L"inspection pane navigation");
    };
    const auto focusedOnDest = [&]()
    {
        const auto focused = view->GetFocusedPath();
        return focused.has_value() && EqualsIgnoreCase(focused->wstring(), dest.wstring());
    };
    const auto statusText = [&]() { return result ? std::to_wstring(static_cast<unsigned>(result->status)) : std::wstring(L"none"); };
    if (! settleAt(home))
        return reject(L"the left pane did not settle in the home folder");

    taskId = StartFileOperationAndGetId(
        operations, FILESYSTEM_COPY, pane, FolderWindow::Pane::Right, state.fsLocal, {source}, destDir, FILESYSTEM_FLAG_RECURSIVE, false);
    if (! taskId)
        return reject(L"the Local Copy admission failed");
    uint64_t generation = 0u;
    std::shared_ptr<const FileOperations::ConflictInspectionContext> context;
    if (! wait(
            [&]
    {
        const auto prompt = TryGetConflictPromptCopy(operations->FindTask(taskId.value()));
        if (! prompt || ! prompt->pathDisplay.inspection || ! prompt->pathDisplay.inspection->destination)
            return false;
        generation = prompt->generation;
        context    = prompt->pathDisplay.inspection;
        return true;
    },
            L"inspectable Copy conflict"))
        return reject(L"the Copy never published an actionable, inspectable destination conflict");
    const auto begin = [&]() noexcept
    {
        result.reset();
        return operations->BeginPendingConflictInspection(taskId.value(), generation, context, true, complete);
    };
    const auto waitResult = [&]() noexcept { return wait([&] { return result.has_value(); }, L"inspection result"); };

    // The resolved destination target carries the prompt's no-follow item kind for the tie-break.
    FileOperations::ConflictInspectionTarget target;
    if (FAILED(operations->ResolvePendingConflictInspectionTarget(taskId.value(), generation, context, true, target)) ||
        target.itemIsDirectory != std::optional<bool>(false))
        return reject(L"the destination target did not carry the prompt's known file kind");

    // folderview#1: a watcher or cache refresh of the target before its enumeration settles.
    if (FAILED(begin()) || result.has_value() || ! CurrentPanePathEquals(folder, pane, destDir))
        return reject(L"the refresh inspection did not navigate and stay pending before its enumeration settled");
    view->ForceRefresh();
    if (! waitResult() || result->status != Status::Selected || ! focusedOnDest())
        return reject(std::format(L"a refresh before settlement ended the inspection with status {}, not by selecting the item", statusText()));

    // folderview#4: a file and a folder listed under one name.
    if (const std::wstring tieBreak = view->DebugValidateInspectionKindTieBreakForSelfTest(identity, dest.native()); ! tieBreak.empty())
        return reject(tieBreak);
    if (! settleAt(home))
        return reject(L"the left pane did not return home after the tie-break check");

    // folderview#4 through the host: the destination settles listing a folder twin of the file
    // (added ahead of the inspection listener); the request must pass the prompt's file kind.
    bool injectTwin = true;
    folder->SetPaneEnumerationSettledCallback(pane,
                                              [&](const FolderView::InspectionVisit&, const std::filesystem::path&, HRESULT status)
    {
        if (injectTwin && SUCCEEDED(status) && ! twin)
        {
            twin       = view->DebugAppendKindTwinRowForSelfTest(identity, dest.native()); // Absent outside the destination.
            injectTwin = ! twin.has_value();
        }
    });
    twinCurrentIsDirectory.reset();
    const bool hostTieBegun = SUCCEEDED(begin());
    const bool hostTieDone  = hostTieBegun && waitResult();
    folder->SetPaneEnumerationSettledCallback(pane, {});
    if (! hostTieDone || injectTwin || result->status != Status::Selected || twinCurrentIsDirectory != std::optional<bool>(false) || ! focusedOnDest())
        return reject(std::format(L"a destination listing a same-name folder ended the inspection with status {} (twin added {}), not by "
                                  L"selecting the file row",
                                  statusText(),
                                  ! injectTwin));
    if (! settleAt(home))
        return reject(L"the left pane did not return home after the host tie-break check");

    // folderview#2: the user cancels the busy enumeration of the inspected folder.
    if (FAILED(begin()) || result.has_value() || ! CurrentPanePathEquals(folder, pane, destDir))
        return reject(L"the canceled inspection did not navigate and stay pending");
    view->CancelEnumerationByUser();
    if (! waitResult() || result->status != Status::Canceled || result->error != HRESULT_FROM_WIN32(ERROR_CANCELLED))
        return reject(std::format(L"a user-canceled enumeration ended the inspection with status {}, not Canceled", statusText()));
    if (! wait([&] { return CurrentPanePathEquals(folder, pane, home) && view->IsCurrentFolderEnumerated(); }, L"canceled inspection rollback"))
        return reject(L"the canceled inspection did not restore the previous folder");
    if (FAILED(begin()) || ! waitResult() || result->status != Status::Selected || ! focusedOnDest())
        return reject(std::format(L"the inspection after a canceled one ended with status {}, not by selecting the item", statusText()));
    if (! settleAt(home))
        return reject(L"the left pane did not return home after the canceled inspection");

    // x-lifetime#1: the popup that owns the pending request is destroyed.
    operations->DebugEnsurePopupVisibleForSelfTest();
    if (! operations->GetPopupHwndForSelfTest())
        return reject(L"the File Operations popup could not be shown");
    if (FAILED(begin()) || result.has_value())
        return reject(L"the popup-owned inspection did not stay pending");
    operations->DebugResetPopupForSelfTest();
    if (! result || result->status != Status::Stale)
        return reject(std::format(L"destroying the popup ended its pending inspection with status {}, not Stale at once", statusText()));
    if (FAILED(begin()) || ! waitResult() || result->status != Status::Selected)
        return reject(std::format(L"the inspection after the popup was destroyed ended with status {}, not Selected", statusText()));
    if (! settleAt(home))
        return reject(L"the left pane did not return home after the popup check");

    // folderview#2: clearing the inspected pane's path, or preparing it for a network drive
    // disconnect, cancels its enumeration with nothing to settle; the request ends Stale at once.
    if (FAILED(begin()) || result.has_value())
        return reject(L"the inspection before the path was cleared did not stay pending");
    folder->DebugClearPanePathForSelfTest(pane);
    if (! result || result->status != Status::Stale)
        return reject(std::format(L"clearing the pane path ended its pending inspection with status {}, not Stale at once", statusText()));
    if (! settleAt(home))
        return reject(L"the left pane did not return home after its path was cleared");
    if (FAILED(begin()) || result.has_value())
        return reject(L"the inspection before the disconnect preparation did not stay pending");
    folder->PrepareForNetworkDriveDisconnect(pane);
    if (! result || result->status != Status::Stale)
        return reject(std::format(L"preparing a disconnect ended the pending inspection with status {}, not Stale at once", statusText()));
    if (! settleAt(home))
        return reject(L"the left pane did not return home after the disconnect preparation");

    // x-lifetime#1: the prompt is answered while its inspection is pending.
    if (FAILED(begin()) || result.has_value())
        return reject(L"the inspection before the answer did not stay pending");
    auto* task        = operations->FindTask(taskId.value());
    const auto prompt = TryGetConflictPromptCopy(task);
    if (! task || ! prompt || prompt->generation != generation || ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) ||
        ! task->SubmitConflictDecision(generation, Task::ConflictAction::Skip, false))
        return reject(L"the pending prompt could not be answered with Skip");
    if (! result || result->status != Status::Stale)
        return reject(std::format(L"answering the prompt ended its pending inspection with status {}, not Stale at once", statusText()));
    if (! wait([&] { return ! operations->FindTask(taskId.value()) && ! operations->HasActiveOperations(); }, L"skipped Copy completion"))
        return reject(L"the skipped Copy did not finish");
    return true;
}
