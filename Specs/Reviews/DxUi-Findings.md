> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# DxUi - verified findings (compact)

- **[accessibility/CONFIRMED]** `Common/DxUi/DxUi.Accessibility.cpp:2459` - ExpandToEnclosingUnit does not expand for Character/Word/Line units (degenerate text range)
- **[contract/CONFIRMED]** `Common/DxUi/DxUi.TextInput.cpp:1020` - ReplaceSelectionAndNotify mutates the text buffer without syncing the native TSF/IME state cache, desyncing ACP offsets
- **[contract/CONFIRMED]** `Common/DxUi/DxUi.Controls.cpp:3321` - Slider mouse drag/click ignores _step and tick marks, so the value is set to an unquantized continuous value (contract violation vs keyboard nav)
- **[dos-oom/CONFIRMED]** `Common/DxUi/DxUi.TextStoreACP.cpp:110` - TSF GetTextExt resolves a range rect by measuring a caret rect for every code unit in the range (O(n) layout work on the UI thread inside a TSF lock)
- **[dos-oom/CONFIRMED]** `Common/DxUi/DxUi.TextStoreACP.cpp:110` - GetTextExt multiline path performs O(text length) caret-rect/layout computations per call with no upper bound
- **[dos-oom/CONFIRMED]** `Common/DxUi/DxUi.SingleLineTextEditing.cpp:213` - O(n^2) grapheme counting on masked-field repaint freezes UI thread (StepToNextTextElement restarts scan from index 0 every call)
- **[dos-oom/PLAUSIBLE]** `Common/DxUi/DxUi.Menu.cpp:2516` - GetMenuItemLayoutRects is O(n) per item during paint, making the paint pass O(n^2) for long/dynamic menus on the UI thread
- **[incorrect-behavior/CONFIRMED]** `Common/DxUi/DxUi.TextInput.cpp:1499` - Masked multiline TextField paints/selects using control-text offsets against a shorter masked display string
- **[incorrect-behavior/CONFIRMED]** `Common/DxUi/DxUi.NativeTextInput.cpp:593` - IME composition preview left committed in the control when the session is torn down on focus loss instead of being abandoned
- **[incorrect-behavior/CONFIRMED]** `Common/DxUi/DxUi.SingleLineTextEditing.cpp:797` - Double-click word selection uses raw UTF-16 code-unit indices and can split a surrogate pair, corrupting text on delete
- **[incorrect-behavior/CONFIRMED]** `Common/DxUi/DxUi.Accessibility.cpp:3818` - Tree item runtime IDs use the volatile visible index, not a stable id (unstable runtime-id, mis-targeted operations)
- **[robustness/CONFIRMED]** `Common/DxUi/DxUi.WindowHost.cpp:3278` - Control virtuals (Paint/OnMouseDown/OnKeyDown/OnChar/Tick) called from noexcept boundaries can std::terminate on any thrown exception
- **[robustness/CONFIRMED]** `Common/DxUi/DxUi.WindowHost.cpp:3329` - Device-removed/reset recovery never re-invalidates: window stays blank until an unrelated paint trigger
- **[robustness/PLAUSIBLE]** `Common/DxUi/DxUi.Controls.cpp:4322` - MenuBar::ActivateItem calls RequestInvalidate() (member _host read) after the open-item callback may have destroyed the MenuBar
- **[robustness/PLAUSIBLE]** `Common/DxUi/DxUi.Controls.cpp:1301` - PageHost::SetPage destroys the outgoing page synchronously in the non-animated path, causing a use-after-free when navigation is triggered from within that page
- **[robustness/PLAUSIBLE]** `Common/DxUi/DxUi.Tree.cpp:799` - OnMouseDoubleClick reuses a stale hit-test index after a reentrant model mutation (OOB / null deref)
- **[robustness/PLAUSIBLE]** `Common/DxUi/DxUi.Tree.cpp:532` - Tree expansion-animation index mapping assumes node-id stability across model rebuild
- **[use-after-free/CONFIRMED]** `Common/DxUi/DxUiNativeMenuInterop.h:751` - Null-pointer dereference of _menuBar after ContextMenu::Show nested loop tears down the host window
- **[use-after-free/CONFIRMED]** `Common/DxUi/DxUi.Controls.cpp:2817` - RadioButton::OnMouseUp reads _onSelected after the group selection callback may have destroyed the button (reentrancy use-after-free)
- **[use-after-free/CONFIRMED]** `Common/DxUi/DxUi.Controls.cpp:6611` - ScrollPanel stores raw captured-child pointer AFTER dispatching the child's mouse-down, re-arming a dangling pointer if the handler clears children
- **[use-after-free/CONFIRMED]** `Common/DxUi/DxUi.WindowHost.cpp:2457` - Pointer-down handler dereferences hit-test target after OnMouseDown, which can synchronously destroy that control (use-after-free)
- **[use-after-free/CONFIRMED]** `Common/DxUi/DxUi.WindowHost.cpp:2534` - Pointer-up diagnostics trace dereferences hit-test target after OnMouseUp may have destroyed it
- **[use-after-free/CONFIRMED]** `Common/DxUi/DxUi.Accessibility.cpp:3500` - UIA provider read methods race against UI-thread control-tree/model mutation (data race, dangling GetChildren() span)
- **[use-after-free/CONFIRMED]** `Common/DxUi/DxUi.WindowHost.cpp:2457` - Pointer-down dispatch dereferences hit-test target after its OnMouseDown handler may have destroyed it (UAF)
- **[use-after-free/PLAUSIBLE]** `Common/DxUi/DxUi.WindowHost.cpp:2503` - WM_LBUTTONUP/RBUTTONUP routes to captured-or-hit target and dereferences it after OnMouseUp without re-validation
