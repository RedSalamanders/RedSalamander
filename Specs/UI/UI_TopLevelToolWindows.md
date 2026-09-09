# Top-Level Tool Windows

## Scope

This spec defines the normative windowing behavior for long-lived RedSalamander tool windows such as:

- Preferences
- Connection Manager (single-canvas DxUi; `RedSalamander/ConnectionManagerWindow.{h,cpp}`; modal-facade synchronous wrapper for plugin-host callers preserves the legacy `S_OK`/`S_FALSE` ABI)
- Find Files and Directories
- Shortcuts
- Compare Directories
- About / Help-style informational windows
- viewer windows such as Space Viewer
- the singleton floating Terminal window
- app-owned captioned utility/dialog windows when a domain spec explicitly includes them in the shared tool-window chrome contract

Native OS dialogs remain out of scope here. Transient app-owned prompts, confirmations, credential editors, and popup-owned helper surfaces may remain explicitly owned or modal when their own specs require it, but any such window that opts into the shared tool-window chrome contract MUST follow the backdrop policy below.

Transient alert/help overlays are not long-lived tool windows. They MAY be implemented as owned top-level popup windows to preserve reliable Win32 input routing and z-order over their anchor/parent, but they MUST remain short-lived, hidden/destroyed by their owner, and follow the transient overlay routing rules in `Specs/UI/UI_CommandMenuKeyboard.md`. Owner-level keyboard dismissal such as `Escape` MUST enumerate these owned top-level popups as well as child windows, because they are intentionally not children of the owner root.

## Normative Rules

- Long-lived tool windows MUST be modeless.
- Long-lived tool windows MUST be independent top-level windows.
- Long-lived tool windows MUST NOT be created as owned windows of the main folder window or any other app window.
- Long-lived tool windows MUST NOT call `EnableWindow(owner, FALSE)` as part of their open path.
- The main folder window MUST remain interactive while a tool window is open.
- A tool window MAY use the invoking window only as a placement reference for initial size, monitor choice, centering, or DPI selection.
- A tool window MAY be single-instance and reuse its existing window instead of opening duplicates.
- When a modeless tool window closes and its invoking owner is still the main
  folder window, the close path MUST request active-pane `FolderView` focus
  restoration instead of leaving keyboard focus on the root main-window HWND.
- RedSalamander tool windows MUST register their large and small window-class icons from the app icon resources so captions and Alt-Tab/task-switch UI do not fall back to the generic system icon.
- Resizable app-owned top-level windows MUST expose a DPI-scaled minimum track size via `WM_GETMINMAXINFO` so the window cannot be resized into an unusable layout. New implementations SHOULD use the shared `Common/WindowSizing.h` helper rather than ad-hoc pixel constants.
- On application shutdown, unowned RedSalamander top-level tool windows MUST be closed during shutdown teardown so graphics resources are released before process exit.

## Backdrop Policy

- Supported tool windows MUST apply the persisted `ui.windowBackdrop` preference through the shared window-backdrop policy/helper path instead of ad-hoc per-window DWM calls.
- Supported app-owned captioned tool/dialog/utility windows MUST use tool-window backdrop target semantics unless their domain spec explicitly defines another target.
- Applying a requested backdrop MUST gracefully fall back to `None` when the OS, window state, or accessibility mode does not support the requested material.
- High contrast remains authoritative over backdrop preference.
- Preferences MAY preview a pending backdrop choice immediately on the live Preferences window before `Apply`, but persisted app-wide changes to tool windows still happen through the normal `Apply` / `OK` or settings hot-reload path.
- The main folder window is outside this spec's shared tool-window backdrop acceptance scope. Its title bar and whole-window backdrop behavior are owned by the main-window implementation/specs, not by tool-window coverage.

## Connection Manager

The Connection Manager is a top-level modeless tool window for normal application commands. Connect validates and saves the current profile, posts a typed payload to the owner, and closes only after the owner notification is queued. Close validates and saves dirty edits or leaves the window open when validation fails.

## Floating Terminal

`cmd/terminal/openFloatingWindow` (`Ctrl+Shift+N`) creates or activates one
independent, modeless floating Terminal top-level window. RedSalamander MUST
never create a second floating Terminal root. Invoking the command while the
root exists adds and selects a new tab at the invoking trusted path. The one
root may contain as many tabs as the user opens; each tab owns an independent
Terminal instance and shell session.

The window persists one placement record plus the ordered live-tab set: stable
tab ID, profile ID, provider ID, canonical trusted path, and active tab ID.
Move/resize and tab reorder state is coalesced before save. A clean shutdown
marks whether the root was open; startup restores one root and fresh shells in
the exact saved tab order, then selects the saved active ID. Closed tabs are
removed immediately and no unbounded closed-session ledger is retained.

The normal tab commands create, switch, select, reorder, and close floating
tabs. `Ctrl+Shift+T` adds a tab at the active trusted path;
`Ctrl+Tab`/`Ctrl+Shift+Tab` cycle live tabs; `Ctrl+Alt+1` through
`Ctrl+Alt+8` select those indices; `Ctrl+Alt+9` selects the last tab; and
`Ctrl+Shift+W` closes the selected tab. A real root-shell exit removes only its
matching tab after final terminal output is drained. Closing or exiting the
last tab closes the floating root. Manual close, shell exit, app shutdown, and
late/duplicate callback races converge on the same idempotent teardown path.

The floating Terminal follows the shared top-level icon, minimum-size, DPI,
backdrop, placement clamping, and shutdown rules above. Window placement is
associated with the singleton, while path/profile state belongs to individual
tabs; it is not a per-path collection of floating windows.

## Rationale

Owned modeless windows stay stacked above their owner and make the main UI feel blocked even when technically still enabled. RedSalamander tool windows are intended to support side-by-side work with the main folder window, so they must behave like independent application windows rather than attached popups.
