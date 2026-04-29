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
- app-owned captioned utility/dialog windows when a domain spec explicitly includes them in the shared tool-window chrome contract

Native OS dialogs remain out of scope here. Transient app-owned prompts, confirmations, credential editors, and popup-owned helper surfaces may remain explicitly owned or modal when their own specs require it, but any such window that opts into the shared tool-window chrome contract MUST follow the backdrop policy below.

## Normative Rules

- Long-lived tool windows MUST be modeless.
- Long-lived tool windows MUST be independent top-level windows.
- Long-lived tool windows MUST NOT be created as owned windows of the main folder window or any other app window.
- Long-lived tool windows MUST NOT call `EnableWindow(owner, FALSE)` as part of their open path.
- The main folder window MUST remain interactive while a tool window is open.
- A tool window MAY use the invoking window only as a placement reference for initial size, monitor choice, centering, or DPI selection.
- A tool window MAY be single-instance and reuse its existing window instead of opening duplicates.
- RedSalamander tool windows MUST register their large and small window-class icons from the app icon resources so captions and Alt-Tab/task-switch UI do not fall back to the generic system icon.
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

## Rationale

Owned modeless windows stay stacked above their owner and make the main UI feel blocked even when technically still enabled. RedSalamander tool windows are intended to support side-by-side work with the main folder window, so they must behave like independent application windows rather than attached popups.
