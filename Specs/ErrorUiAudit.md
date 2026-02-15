# Error UI Audit (No MessageBox)

Goal: eliminate Win32 `MessageBox*` usage in shipping binaries (**RedSalamander.exe**, **RedSalamanderMonitor.exe**, and in-tree plugins) and consolidate user-facing error/info/confirmation UI into shared host-rendered surfaces (overlay alerts, blocking overlay prompts, and inline validation).

Status: **DONE** — there are no remaining `MessageBox*` call sites under `RedSalamander/`, `RedSalamanderMonitor/`, or `Plugins/`.

## Replacements used

### Blocking prompts (replaces `MB_OKCANCEL` / `MB_YESNO*`)

- **DONE**: Copy/Move confirmation now uses `IHostPrompts::ShowPrompt` (overlay prompt) instead of `MessageBox*` (`RedSalamander/FolderViewInternal.h`, `RedSalamander/FolderWindow.FileOperations.State.cpp`).
- **DONE**: Cancel-all confirmations now use `IHostPrompts::ShowPrompt` (`RedSalamander/FolderWindow.FileOperations.Popup.cpp`, `RedSalamander/FolderWindow.FileOperations.cpp`).

### Fatal/startup (replaces modal `MessageBoxResource` / `MessageBoxW`)

- **DONE (RedSalamander)**: Startup and fatal crash paths show `IDD_FATAL_ERROR` (modal dialog, not `MessageBox*`) via `ShowFatalErrorDialog(...)` in `RedSalamander/RedSalamander.cpp`.
- **DONE (Monitor)**: Startup/fatal paths show `IDD_MODAL_MESSAGE` (modal dialog, not `MessageBox*`) via `ShowModalMessageDialog(...)` in `RedSalamanderMonitor/RedSalamanderMonitor.cpp`.
- **DONE**: FolderWindow creation now fails cleanly when WM_CREATE initialization fails (`RedSalamander/FolderWindow.cpp`), allowing the fatal path to be centralized in the caller.

### Inline validation (no modal UI)

- **DONE**: Navigation invalid path validation uses edit balloon tips (`EM_SHOWBALLOONTIP`) (`RedSalamander/NavigationView.Edit.cpp`, `RedSalamander/NavigationView.FullPathPopup.cpp`).
- **DONE**: Preferences invalid shortcut command id uses an edit balloon tip (`RedSalamander/PreferencesDialog.cpp`).

### Non-blocking alerts / info (overlay)

- **DONE**: “Command not implemented” uses `IHostAlerts::ShowAlert` (overlay alert) (`RedSalamander/RedSalamander.cpp`).
- **DONE**: Plugins manager errors/info use `IHostAlerts::ShowAlert` (overlay alert) (`RedSalamander/ManagePluginsDialog.cpp`).
- **DONE**: FolderView and ViewerText already render in-window alerts using `RedSalamander::Ui::AlertOverlay`.

## Follow-ups (optional)

- `Common/Helpers.h` still contains legacy `MessageBox*` helper implementations. Shipping code should keep avoiding these; they can be removed once there are no remaining consumers.
- PoC projects are not part of this audit and may still use `MessageBoxW`.
