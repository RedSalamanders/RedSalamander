# DxUI Find Menu Input Routing Repair Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` to execute this plan.

## Goal

Make the Find dialog split-button menu behave like a normal Windows menu under real input:

- hover highlights the item under the pointer,
- clicking an enabled item invokes immediately,
- clicking outside closes the menu immediately,
- visible hover and close feedback flows through popup invalidation/`WM_PAINT` without waiting for a later owner-window/title-bar message,
- the Find dialog remains responsive while the menu is open,
- existing column sizing and list behavior regressions stay fixed.

## Root Cause Assessment

The original menu implementation had a split-brain input model. Menu pointer state was mutated from both the popup window procedure and the modal menu loop, with additional idle cursor polling layered on top. That made real Win32 input order fragile and let tests pass through one route while the running Find dialog failed through another.

Specific risks found in the current code:

- `Common/DxUi/DxUi.Menu.cpp` handles popup mouse messages in both `MenuWndProc` / `ProcessMenuPopupMouseMessage(...)` and `RunMenuModalLoop(...)`.
- Idle cursor polling ran every 16 ms and competed with queued mouse messages, so it could overwrite authoritative message-driven state.
- The Find split-button opens the menu synchronously from the button mouse-up callback in `RedSalamander/FindFilesWindow.cpp`, entering a nested modal menu loop before the original button dispatch has unwound.
- Existing tests simulate or fall back to posted messages, so they do not prove the failing real desktop route: physical hover, click, outside-dismiss, and immediate command return.

## Impact Scope

This is not a Find-only fix. `ContextMenu::Show(...)` is shared DxUI infrastructure, so the router change affects every DxUI menu caller.

Review these production callers during implementation:

- `RedSalamander/FindFilesWindow.cpp`: Find split-button action menu. Confirmed caller-specific deferral is required because it currently opens the menu directly from the split-button mouse-up callback.
- `RedSalamander/NavigationView.Menus.cpp`: navigation menu, drive menu, disk info, history, and related dropdowns. This already uses posted dropdown messages for some entrypoints; preserve that model and verify root switching.
- `RedSalamander/NavigationView.FullPathPopup.cpp`: full-path popup menu hosted by the navigation popup.
- `RedSalamander/CompareDirectoriesWindow.Menu.cpp`: Compare Directories sort popup and Dx menu bar popup.
- `RedSalamander/RedSalamander.cpp`: main window sort popup, main menu bar Dx popup, and user/native-menu interop paths.
- `RedSalamander/FolderView.Menus.cpp`: folder view context menu opened from right-click or keyboard context-menu command.
- `RedSalamander/FolderWindow.FileOperations.Popup.cpp`: file operation speed-limit and destination menus.
- `RedSalamander/FolderWindow.FileSystem.cpp`: prompt history menu opened from a DxUI button.
- `RedSalamander/FolderWindow.Layout.cpp`: filter-bar history menu.
- `Common/DxUi/DxUiNativeMenuInterop.h`: native `HMENU` to DxUI menu conversion and menu bar adapter paths.

For each caller, decide whether it only needs shared router validation or also needs a caller-specific deferred-open change. Do not blindly defer every caller; defer only menus opened from inside an input callback where the nested modal loop can interfere with the dispatch stack.

## Architecture Decision

Replace duplicated popup/modal pointer handling with one explicit menu input router. All pointer events must enter the same state machine, tagged by source.

The popup window procedure and the modal loop may still both receive Win32 messages, but neither should directly mutate hover, keyboard focus, button-down, invoke, dismiss, submenu, or root-switch state outside the router.

Idle polling must not be a second state machine. The final design keeps delivered popup/menu messages authoritative and permits only a bounded idle input resync that routes through the same menu controller. That resync is a watchdog for starved owner-window queues and fast click edges; it must not switch menu-bar roots or force repainting behind delivered messages.

Keyboard handling is part of the same repair. The final design must make keyboard routing first-class, not leave it as another independent state machine beside pointer routing.

Keyboard contract:

- `VK_UP` / `VK_DOWN` move to the previous or next navigable item and skip separators, info rows, and disabled entries.
- `VK_HOME` / `VK_END` move to the first or last navigable item, except focused slider items may use them for min/max stops.
- `VK_RETURN` / `VK_SPACE` invoke the focused item, or open its submenu, or commit the focused slider stop.
- `VK_ESCAPE` closes the top submenu; if only the root popup is open, it dismisses the whole menu.
- `VK_LEFT` closes a submenu, switches menu-bar roots when supported, or decrements a focused slider.
- `VK_RIGHT` opens a submenu, switches menu-bar roots when supported, or increments a focused slider.
- `VK_TAB`, `VK_F10`, and bare `VK_MENU` dismiss the menu and restore focus according to the existing session contract.
- mnemonic letters select, open, or invoke the matching enabled item.
- pointer movement may clear keyboard focus only when the movement comes from a real pointer message inside a popup.

The Find split-button menu should open from a posted private message, not directly inside the button mouse-up callback.

## Implementation Tasks

### 1. Add Diagnostic Trace Before Refactoring

Add a small diagnostic ring buffer around menu routing in `Common/DxUi/DxUi.Menu.cpp`.

Capture at least:

- sequence number,
- source: popup window proc, modal loop, Find menu open, invoke, dismiss,
- message id,
- message hwnd,
- capture hwnd,
- focus hwnd,
- active hwnd,
- screen point,
- hit popup index,
- hit item index,
- hovered index,
- keyboard index,
- pressed item state,
- running/result state,
- timestamp.

Expose the trace to `DxUiTests` and optionally dump it to `OutputDebugStringW` on failed menu selftests. This is required because the previous fixes were made without seeing the real failing event sequence.

### 2. Write Failing Real-Path Regression Tests

Add tests before changing production behavior.

Required coverage:

- Open a `ContextMenu` from a split-button-style mouse-up callback and verify the menu can highlight, invoke, and return without an extra activation/title-bar click.
- Open the Find dialog menu through the real Find split-button path and click `Find Now`; verify the command starts promptly.
- After results exist, hover an enabled `Refine ...` item and verify visual hover state changes.
- Click outside the menu and verify the menu closes promptly.
- Open a menu with the keyboard and verify `Up`, `Down`, `Home`, `End`, `Enter`, `Space`, `Escape`, `Left`, `Right`, `Tab`, `F10`, bare `Alt`, and mnemonic keys still behave as normal menu keys.
- Open a menu with the mouse, then use the keyboard before moving the pointer; verify stationary cursor state does not erase the keyboard highlight.
- Open a submenu with keyboard, move the pointer inside the existing submenu chain, and verify submenu ownership and highlight remain stable.

Do not silently pass these tests by falling back to posted messages when real `SendInput` is expected. If the environment cannot deliver desktop input, mark that variant skipped with an explicit reason and keep a separate deterministic posted-message test.

### 3. Review Shared Caller Impact

Before production changes, audit every production `ContextMenu::Show(...)` caller listed in `Impact Scope`.

For each caller, record:

- how the menu opens: mouse up, right-click, keyboard, posted message, command callback, or custom popup host,
- whether opening happens inside another DxUI control callback,
- whether the caller uses `ignoreInitialLeftButtonUp`, `focusFirstNavigableItem`, root switching, native `HMENU` conversion, sliders, or submenus,
- what focus should be restored when the menu closes,
- which selftest or manual validation covers that caller.

Expected caller decisions:

- Find split-button: defer menu opening with a posted private message.
- NavigationView dropdowns: preserve the existing posted-message opening model and regression-test the shared router.
- Main/Compare menu bars: regression-test root switching and keyboard left/right behavior.
- FolderView and file-operation context menus: regression-test pointer hover, outside-dismiss, keyboard context-menu opening, and immediate invocation.
- Prompt/filter history buttons: audit for the same synchronous-button-callback problem as Find and defer if the trace proves nested dispatch is involved.

### 4. Remove Accidental Complexity Before Adding More Code

Before introducing the final router, audit the touched menu and Find dialog code for complexity that only exists because of earlier failed fixes.

Remove or simplify:

- duplicate hover state variables that represent the same concept,
- fallback branches that silently convert one input route into another,
- test-only behavior that changes production timing or message routing,
- idle-poll state that overlaps with real mouse-message state,
- redundant root-switch pointer tracking if the router can express it directly,
- debug helpers that mutate behavior instead of only observing it,
- special cases added for one test but not required by the UI contract,
- synchronous menu-open paths that can be replaced by posted UI-thread messages,
- dead branches left after popup and modal input handling are unified.

The cleanup pass must preserve intentional behavior only when it is documented by a test, a spec, or a clear Win32 requirement. Anything else should be treated as suspicious until proven useful.

### 5. Introduce A Single Menu Input Router

In `Common/DxUi/DxUi.Menu.cpp`, introduce a single internal routing API, for example:

```cpp
enum class MenuInputSource
{
    ModalMessage,
    PopupWndProc,
};

enum class MenuPointerKind
{
    Move,
    LeftDown,
    LeftUp,
    RightDown,
    RightUp,
    Wheel,
    Cancel,
};

struct MenuPointerEvent
{
    MenuInputSource source;
    MenuPointerKind kind;
    HWND hwnd;
    POINT screenPoint;
    WPARAM wParam;
    LPARAM lParam;
    bool mayInvoke;
    bool mayDismiss;
    bool maySwitchRoot;
};

enum class MenuKeyboardKind
{
    Up,
    Down,
    Left,
    Right,
    Home,
    End,
    Enter,
    Space,
    Escape,
    Tab,
    F10,
    Alt,
    Mnemonic,
};

struct MenuKeyboardEvent
{
    MenuInputSource source;
    MenuKeyboardKind kind;
    HWND hwnd;
    UINT virtualKey;
    wchar_t mnemonic;
    LPARAM lParam;
};

enum class MenuInputDisposition
{
    Ignored,
    Consumed,
    HoverChanged,
    Dismissed,
    Invoked,
    SwitchedRoot,
};
```

Then route all popup and modal pointer input through one function:

```cpp
MenuInputDisposition RouteMenuPointerEvent(MenuController& controller, const MenuPointerEvent& event);
MenuInputDisposition RouteMenuKeyboardEvent(MenuController& controller, const MenuKeyboardEvent& event);
```

Rules:

- only the router changes hovered item, keyboard item, pressed item, submenu timers, invoke result, dismiss result, and root switch state,
- popup window procedure events are authoritative for popup hover, click, invoke, dismiss, and keyboard routing,
- the modal loop must dispatch popup messages rather than pre-consuming them, while keeping global coordination such as activation loss, menu-bar root switching, and submenu timers,
- idle poll events are positive-only hover refreshes,
- keyboard events are authoritative for `keyboardIndex` and clear `hoveredIndex` only when a navigation key successfully moves focus,
- disabled items can hover if the existing UI contract allows it, but cannot invoke.

### 6. Defer The Find Split-Button Menu Open

Change `RedSalamander/FindFilesWindow.cpp` so the drop-down callback posts a private window message instead of calling `ShowFindActionMenu(...)` synchronously from the button mouse-up stack.

Expected shape:

- drop-down callback computes/stores the anchor point,
- callback posts `WM_APP`-range Find-window message,
- Find window message handler calls `ShowFindActionMenu(...)` after the original button input dispatch has returned.

This removes the nested modal loop from inside `Button::OnMouseUp` and should address the symptom where the clicked action only happens after another title-bar click.

### 7. Defer Other Callback-Opened Menus Only If The Caller Audit Requires It

Use the diagnostic trace and caller audit to decide whether the prompt history menu, filter-bar history menu, or file-operation popup menus also need deferred opening.

If a caller opens a menu from inside a DxUI button/control input callback and shows the same nested-dispatch symptoms as Find, change it to the same posted-message model:

- add a unique `WM_APP`-range message in `Common/WindowMessages.h`,
- store or encode the anchor data needed to open the menu,
- return from the control callback,
- open the menu from the window message handler,
- keep command execution after `ContextMenu::Show(...)` returns.

Do not add deferral to right-click context menus or already-posted dropdowns unless the trace shows they need it.

### 8. Remove The Old Competing Paths

After the router is in place:

- remove direct hover/invoke/dismiss mutation from `ProcessMenuPopupMouseMessage(...)`,
- remove equivalent direct pointer and keyboard mutation from the modal loop,
- keep only conversion from Win32 message to `MenuPointerEvent`,
- keep only conversion from Win32 key messages to `MenuKeyboardEvent`,
- remove the old continuous idle cursor polling path instead of keeping it as a hidden fallback state machine,
- keep only the bounded modal input resync that feeds cursor/button state into the same router and does not perform root switching.

This step is not optional. Leaving both old and new paths active would preserve the class of bug that caused this regression.

### 9. Final Simplification Review

After behavior is green, do one final readability and ownership pass.

Check for:

- one owner for menu lifetime,
- one owner for hover/keyboard/pressed state,
- one place where outside clicks dismiss the menu,
- one place where item invocation is committed,
- no hidden input fallback that makes tests pass through a different route than the real app,
- no polling path that can undo message-driven state or switch roots behind delivered menu-bar messages,
- no column/list changes entangled with menu fixes unless directly required.

If a helper, flag, timer, or state field cannot be explained in terms of the final menu contract, remove it.

### 10. Verify

Run:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=Menu
.\.build\x64\Debug\DxUiTests.exe --suite=NewControls
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

Also run the Find dialog command selftests that cover menu enablement and command dispatch.

Manual validation is required before closing this plan:

- open Find dialog,
- open the split-button menu,
- move over every item and confirm hover feedback,
- click `Find Now` and confirm the action starts immediately,
- create results and confirm `Refine ...` items hover and invoke,
- click outside the menu and confirm it closes immediately,
- reopen the Find menu and verify `Down`, `Up`, `Enter`, `Escape`, and mnemonic activation,
- verify NavigationView menus, main menu bar, Compare Directories menu bar, FolderView context menu, file-operation popup menus, prompt history, and filter history still hover, invoke, outside-dismiss, and keyboard-navigate correctly,
- repeat after resizing list columns and sorting headers.

### 11. Update Documentation And Specs

Update the authoritative documentation after the final behavior is implemented and verified. Do not leave the durable contract only in this WIP plan.

Required documentation updates:

- `Specs/UI/UI_DxUiWinUIDesign.md`: document the shared DxUI menu input contract, including one router for pointer and keyboard input, bounded idle input resync, focus restoration, outside-dismiss, and popup/modal message ownership.
- `Specs/UI/UI_DxUiSharedGrid.md`: document the Find results list expectations that were part of the original report, including stable resized column widths across sort changes, path/subfolder display behavior, and file/folder icon expectations if this fix touches or validates them.
- Find dialog spec, or create one under `Specs/UI/` if no Find-specific spec exists: document the Find split-button menu behavior, deferred-open requirement, action enablement, keyboard support, and expected responsiveness.
- NavigationView/menu bar specs: document that existing posted dropdown opening remains intentional, and that root switching must work by mouse hover and keyboard left/right.
- FolderView/context menu specs: document that DxUI context menus opened from right-click or keyboard context-menu command share the same hover, invoke, outside-dismiss, and keyboard rules.
- File operation popup specs: document speed-limit and destination menu behavior if any caller-specific changes are made there.
- `Common/DxUi` developer documentation, if present; otherwise add a short DxUI menu section to the closest existing DxUI spec rather than creating duplicate guidance.

Documentation must name both the behavior and the ownership rule: callers choose when to open a menu, but `ContextMenu::Show(...)` owns all menu hover, keyboard, invocation, dismissal, submenu, and root-switch state while the menu is active.

## Spec Closeout

Before marking complete:

- complete the documentation updates in Task 11,
- move this plan from `Specs/Plans/WIP/` to `Specs/Plans/Done/` only after tests and manual validation pass.

## Closeout Notes

Completed on 2026-05-31.

- Implemented the shared menu input router with popup-window-procedure ownership for popup mouse and keyboard messages; the modal loop now dispatches popup messages and keeps only global coordination such as activation loss, menu-bar root switching, and submenu timers.
- Removed the startup live-cursor sample and the old competing idle cursor polling state machine. The remaining bounded modal input resync handles starved hover/click/outside-dismiss input and fast click edges by routing through the same controller, and it cannot switch menu-bar roots behind delivered messages.
- Removed the forced hover repaint helper. Popup state changes now invalidate and repaint through the normal popup `WM_PAINT` path; the only remaining menu `UpdateWindow(...)` is the first-show render path for newly created popup windows.
- Deferred the Find split-button action menu through `WndMsg::kFindShowActionMenu`, so the nested `ContextMenu::Show(...)` loop no longer starts inside the split-button mouse-up callback.
- Covered the live route with `TestSplitButtonContextMenuLivePointerHoverAndOutsideDismiss` and stricter Find split-menu live input selftests; covered the deterministic route with `TestSplitButtonContextMenuSentMouseMessagesHoverAndInvokeImmediately`.
- Updated the durable contracts in `Specs/UI/UI_DxUiWinUIDesign.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/UI/UI_DxUiSharedGrid.md`, `Specs/UI/UI_FindFilesWindow.md`, `Specs/UI/UI_NavigationView.md`, and `Specs/UI/UI_FolderView.md`.
- The diagnostic trace requirement was satisfied with routed debug state/debug paint seams and `DXUI_MENU_TRACE` instrumentation rather than a persistent ring buffer; no test-only production timing hooks were kept.
- Manual hand validation was not performed in this agent run; automated live pointer, sent-message, DxUi, and Find-window command selftests covered the reported menu, keyboard, outside-dismiss, column, path, icon, and Open Parent scenarios.
