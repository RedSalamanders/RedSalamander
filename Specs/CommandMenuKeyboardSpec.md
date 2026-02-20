# Command, Menu, and Keyboard Specification

This document defines the canonical **command catalog** (`cmd/*`), the RedSalamander **main menu bar** structure, and **keyboard routing + shortcut defaults** for RedSalamander (main app) and RedSalamanderMonitor.

## Goals

- Provide a single source of truth for **commands**, **menus**, and **shortcuts** so `Specs/*` and implementation stay aligned.
- Make pane navigation and file operations fully usable from the keyboard.
- Preserve existing selection/navigation semantics (Arrow keys, Shift/Ctrl modifiers, Page Up/Down behavior).
- Provide a keyboard-driven **Function Bar** that reflects current shortcut configuration.
- Ensure all main-menu entries map to a stable command ID and a localized display name.

## Terminology

- **Focused pane**: the pane that contains the current keyboard focus (either its `NavigationView` or `FolderView`).
- **Active pane**: the pane targeted by “apply to active pane” commands when focus is not inside either pane (implementation uses `FolderWindow::_activePane` as fallback).
- **Target pane**: the pane a `cmd/pane/*` command is applied to for this invocation (resolved from focused/active pane rules or from an explicit Left/Right menu origin).
- **Current item**: the item with the caret in a `FolderView` (the one that moves with Arrow keys).
- **Selected items**: the multi-selection set.
- **Incremental search mode**: a transient mode within `FolderView` entered by typing printable characters.
- **Shortcut chord**: a key + modifier set (Ctrl/Alt/Shift) bound to a command ID (e.g. `cmd/...`).

## Key Routing Model (Normative)

### Priority order

1. **Menu loop**: when the Win32 menu loop is active, it owns the keyboard until it exits.
2. **Configurable shortcuts**: `settings.shortcuts` bindings are evaluated (Function Bar + FolderView).
3. **Accelerators**: application accelerators (TranslateAccelerator) run next and send `WM_COMMAND`.
4. **Focused control handlers**: `FolderView`, `NavigationView` and any edit controls handle remaining messages.

When the **menu loop** is active, pressing `Tab` (or `Shift+Tab`) MUST exit menu mode and return focus to the active pane (or the previously focused pane control).

When an **edit control** is focused, the host MUST bypass (2) and (3) and MUST NOT execute application-level accelerators or configurable shortcut bindings (text-edit safety).

### Scope and focus

- Shortcuts that target “the active pane” MUST resolve to the **focused pane** when focus is inside a pane; otherwise use the **active pane**.
- Shortcuts that act on file selection MUST prefer `Selected items`; if none are selected, they act on the `Current item` (and MAY implicitly select it for the operation).
- While an **edit control** is active (NavigationView address edit, rename edit, dialogs), the edit control owns the keyboard: application-level accelerators and configurable shortcut bindings MUST NOT execute (text-edit safety).

### Command resolution (normative)

- **Shortcuts** and **menu items** map to commands identified by stable IDs (example shape: `cmd/...`).
- Command IDs MUST be in one of these namespaces:
  - `cmd/app/*`: application-global commands.
  - `cmd/pane/*`: pane-targeted commands; these MUST resolve to the **focused pane** when focus is inside a pane, otherwise the **active pane** (see rules above).
- Command display names MUST be localized resource strings (`.rc` STRINGTABLE). UI (Function Bar, settings dialog, tooltips) MUST NOT hardcode user-facing command names.
- If a shortcut is bound to a command that is not implemented at runtime, invoking it MUST show a localized message box stating it is not yet implemented and MUST do nothing else.
- Some menu items are **parameterized** (example: History/Hot Path entries, plugin/theme entries). Parameterized menu entries MUST still use a stable `cmd/*` ID; the parameter is carried by the menu item payload (not encoded into the string ID unless explicitly specified).

### Canonical Command IDs

This section is the single source of truth for the command ID catalog.

**Application commands (`cmd/app/*`)**
- `cmd/app/about` *(planned)*
- `cmd/app/exit`
- `cmd/app/openLeftDriveMenu`
- `cmd/app/openRightDriveMenu`
- `cmd/app/compare`
- `cmd/app/fullScreen`
- `cmd/app/openFileExplorerKnownFolder` *(planned, parameterized: knownFolderId)*
- `cmd/app/preferences` *(planned)*
- `cmd/app/showShortcuts`
- `cmd/app/swapPanes`
- `cmd/app/toggleFunctionBar` *(planned)*
- `cmd/app/toggleMenuBar` *(planned)*
- `cmd/app/viewWidth`
- `cmd/app/rereadAssociations` *(planned)*
- `cmd/app/theme/select` *(planned, parameterized: themeId)*
- `cmd/app/theme/systemHighContrastIndicator` *(planned)*
- `cmd/app/plugins/manage` *(planned)*
- `cmd/app/plugins/toggleEnabled` *(planned, parameterized: pluginId)*
- `cmd/app/plugins/configure` *(planned, parameterized: pluginId)*

**Pane commands (`cmd/pane/*`)**
- `cmd/pane/historyBack` *(planned)*
- `cmd/pane/historyForward` *(planned)*
- `cmd/pane/goDriveRoot` *(parameterized: driveLetter)*
- `cmd/pane/goRootDirectory` *(planned)*
- `cmd/pane/setPathFromOtherPane` *(planned)*
- `cmd/pane/navigatePath` *(planned, parameterized: path)*
- `cmd/pane/selectFileSystemPlugin` *(planned, parameterized: pluginId)*
- `cmd/pane/bringCurrentDirToCommandLine`
- `cmd/pane/bringFilenameToCommandLine`
- `cmd/pane/clipboardCut` *(planned)*
- `cmd/pane/clipboardCopy`
- `cmd/pane/clipboardPaste`
- `cmd/pane/clipboardPasteShortcut` *(planned)*
- `cmd/pane/copyNameAsText`
- `cmd/pane/copyPathAndFileName`
- `cmd/pane/copyPathAndNameAsText`
- `cmd/pane/copyPathAsText`
- `cmd/pane/executeOpen`
- `cmd/pane/moveToRecycleBin`
- `cmd/pane/openCurrentFolder`
- `cmd/pane/openProperties`
- `cmd/pane/openSecurity` *(planned)*
- `cmd/pane/permanentDeleteWithValidation`
- `cmd/pane/quickSearch`
- `cmd/pane/selectCalculateDirectorySizeNext`
- `cmd/pane/selectNext`
- `cmd/pane/switchPaneFocus`
- `cmd/pane/upOneDirectory`
- `cmd/pane/windowMenu`
- `cmd/pane/alternateView`
- `cmd/pane/calculateDirectorySizes`
- `cmd/pane/changeAttributes`
- `cmd/pane/changeCase`
- `cmd/pane/changeDirectory`
- `cmd/pane/connect`
- `cmd/pane/contextMenu`
- `cmd/pane/contextMenuCurrentDirectory`
- `cmd/pane/disconnect`
- `cmd/pane/edit`
- `cmd/pane/editWith` *(planned, parameterized: editorId)*
- `cmd/pane/editNew`
- `cmd/pane/editWidth`
- `cmd/pane/filter`
- `cmd/pane/find`
- `cmd/pane/hotPaths`
- `cmd/pane/listOpenedFiles`
- `cmd/pane/showFoldersHistory`
- `cmd/pane/loadSelection`
- `cmd/pane/makeFileList` *(planned)*
- `cmd/pane/menu`
- `cmd/pane/pack`
- `cmd/pane/permanentDelete`
- `cmd/pane/refresh`
- `cmd/pane/saveSelection`
- `cmd/pane/shares`
- `cmd/pane/unpack`
- `cmd/pane/userMenu`
- `cmd/pane/zoomPanel`
- `cmd/pane/copyToOtherPane`
- `cmd/pane/createDirectory`
- `cmd/pane/delete`
- `cmd/pane/display/brief`
- `cmd/pane/display/detailed`
- `cmd/pane/display/extraDetailed`
- `cmd/pane/moveToOtherPane`
- `cmd/pane/rename`
- `cmd/pane/sort/none`
- `cmd/pane/sort/attributes` *(planned)*
- `cmd/pane/sort/extension`
- `cmd/pane/sort/name`
- `cmd/pane/sort/size`
- `cmd/pane/sort/time`
- `cmd/pane/view`
- `cmd/pane/viewWith` *(planned, parameterized: viewerId)*
- `cmd/pane/viewSpace`
- `cmd/pane/newFromShellTemplate` *(planned, parameterized: templateId)*
- `cmd/pane/selection/selectDialog` *(planned)*
- `cmd/pane/selection/unselectDialog` *(planned)*
- `cmd/pane/selection/invert` *(planned)*
- `cmd/pane/selection/selectAll` *(planned)*
- `cmd/pane/selection/unselectAll` *(planned)*
- `cmd/pane/selection/restore` *(planned)*
- `cmd/pane/selection/selectSameExtension` *(planned)*
- `cmd/pane/selection/unselectSameExtension` *(planned)*
- `cmd/pane/selection/selectSameName` *(planned)*
- `cmd/pane/selection/unselectSameName` *(planned)*
- `cmd/pane/selection/hideSelectedNames` *(planned)*
- `cmd/pane/selection/hideUnselectedNames` *(planned)*
- `cmd/pane/selection/showHiddenNames` *(planned)*
- `cmd/pane/selection/goToPreviousSelectedName` *(planned)*
- `cmd/pane/selection/goToNextSelectedName` *(planned)*
- `cmd/pane/goToShortcutOrLinkTarget` *(planned)*
- `cmd/pane/openCommandShell`
- `cmd/pane/viewOptions/toggleHiddenFiles` *(planned)*
- `cmd/pane/viewOptions/toggleSystemFiles` *(planned)*
- `cmd/pane/viewOptions/toggleFileExtensions` *(planned)*
- `cmd/pane/viewOptions/toggleThumbnails` *(planned)*
- `cmd/pane/viewOptions/togglePreviewPane` *(planned)*
- `cmd/pane/viewOptions/toggleFilterBar` *(planned)*
- `cmd/pane/viewOptions/toggleNavigationBar` *(planned)*
- `cmd/pane/viewOptions/toggleStatusBar` *(planned)*

## Main Menu Bar (Target)

### Requirements (Normative)

- The menu bar structure and static labels MUST be defined in `.rc` resources (`RedSalamander/RedSalamander.rc`) to support localization (see `Specs/LocalizationSpec.md`).
- Each menu item that triggers application behavior MUST map to a `cmd/*` command ID (shown in brackets below).
  - If the menu item is dynamic and requires a parameter (history path, hot path, plugin ID, theme ID), the menu item MUST still map to a stable `cmd/*` command ID; the parameter is carried in the menu item payload.
- The displayed shortcut text (when present) MUST reflect the effective current bindings (default or user-customized).
- Top-level menu order MUST be:
  - `Left`, `Files`, `Edit`, `Commands`, `Plugins`, `View`, `Right`, `Help`
- `Help` MUST be right-justified (appear at the right edge of the menu bar).

### Placement rationale (Non-normative)

- **Left/Right**: pane-scoped navigation/view commands for the corresponding pane.
- **Files**: operations on the selection/current item (view/edit/copy/move/delete/properties).
- **Edit**: clipboard + selection set manipulation + “copy as text” utilities.
- **Commands**: directory utilities, lists, network/connect, shell, association refresh, user menu, Explorer jump list.
- **Plugins**: plugin management and plugin selection/configuration.
- **View**: UI/layout/theme preferences and view toggles.
- **Help**: help menu (documentation, about, etc ...)

### Menu structure (Target)

Notation:
- `[cmd/...]` suffix links the menu entry to the command system.
- `(shortcut)` shows the current default shortcut when one exists; `⊘` means none by default.
- `[td]` suffix in the label means the command/menu entry is not implemented yet (TODO).
- `[dbg]` suffix in the label means the menu entry is debug-only.
- `…` indicates a modal dialog or picker is expected.

#### Left (pane menu: targets Left pane)

- Change Drive (`Alt+F1`) *(opens file-system drive menu; when pane is in a non-`file` plugin, the NavigationView menu also exposes a bottom “Change Drive” submenu)* `[cmd/app/openLeftDriveMenu]`
- Go to >
  - Back [td] (`Alt+Left`) *(History Back)* `[cmd/pane/historyBack]`
  - Forward [td] (`Alt+Right`) *(History Forward)* `[cmd/pane/historyForward]`
  - Parent Directory (`Backspace`) `[cmd/pane/upOneDirectory]`
  - Root Directory [td] (`⊘`) `[cmd/pane/goRootDirectory]`
  - Path from Other Panel (`⊘`) `[cmd/pane/setPathFromOtherPane]`
  - ---
  - Hot Paths… [td] (`Shift+F9`) `[cmd/pane/hotPaths]`
  - *(Hot Paths section, dynamic)*
    - `<Hot Path Name>` [td] (`⊘`; parameterized: path) `[cmd/pane/navigatePath]`
  - ---
  - *(History section, dynamic)*
    - `<History Path>` [td] (`⊘`; parameterized: path) `[cmd/pane/navigatePath]`
- ---
- Brief (`Alt+2`) `[cmd/pane/display/brief]`
- Detailed (`Alt+3`) `[cmd/pane/display/detailed]`
- Extra Detailed (`Alt+4`) `[cmd/pane/display/extraDetailed]`
- ---
- Sort By >
  - None (`Ctrl+F2`) `[cmd/pane/sort/none]`
  - Name (`Ctrl+F3`) `[cmd/pane/sort/name]`
  - Extension (`Ctrl+F4`) `[cmd/pane/sort/extension]`
  - Time (`Ctrl+F5`) `[cmd/pane/sort/time]`
  - Size (`Ctrl+F6`) `[cmd/pane/sort/size]`
  - Attributes (`⊘`) `[cmd/pane/sort/attributes]`
- ---
- Maximize/Restore Pane (`Ctrl+F11`) *(toggle: move splitter to edge; restore only if splitter wasn't dragged while maximized; state persisted in settings)* `[cmd/pane/zoomPanel]`
- Swap Panes (`Ctrl+U`) *(swap Left/Right pane file system + current folder; view options stay with the pane; global history unaffected)* `[cmd/app/swapPanes]`
- ---
- Filter… [td] (`Ctrl+F12`) `[cmd/pane/filter]`
- Refresh (`Ctrl+F9`) *(invalidate directory cache + re-enumerate current folder)* `[cmd/pane/refresh]`

#### Files (targets Focused pane unless explicitly stated)

- Rename… (`F2`) `[cmd/pane/rename]`
- Open / Execute (`Enter`) `[cmd/pane/executeOpen]`
- View (`F3`) `[cmd/pane/view]`
- View Width… (`Ctrl+Shift+F3`) `[cmd/app/viewWidth]`
- Alternate View [td] (`Alt+F3`) `[cmd/pane/alternateView]`
- View With > [td]
  - *(Viewer list, dynamic)*
    - `<Viewer Name>` [td] (`⊘`; parameterized: viewerId) `[cmd/pane/viewWith]`
- Edit [td] (`F4`) `[cmd/pane/edit]`
- Edit Width… [td] (`Ctrl+Shift+F4`) `[cmd/pane/editWidth]`
- Edit With > [td]
  - *(Editor list, dynamic)*
    - `<Editor Name>` [td] (`⊘`; parameterized: editorId) `[cmd/pane/editWith]`
- Edit New File… [td] (`Shift+F4`) `[cmd/pane/editNew]`
- Copy… (`F5`) `[cmd/pane/copyToOtherPane]`
- Move/Rename… (`F6`) `[cmd/pane/moveToOtherPane]`
- Delete >
  - Delete (`F8`) `[cmd/pane/delete]`
  - Move to Recycle Bin (`Del`) `[cmd/pane/moveToRecycleBin]`
- Permanent Delete [td] (`Shift+F8`) `[cmd/pane/permanentDelete]`
- Permanent Delete (With Validation) [td] (`Shift+Del`) `[cmd/pane/permanentDeleteWithValidation]`
- Properties (`Alt+Enter`) `[cmd/pane/openProperties]`
- Context Menu (`Shift+F10`) `[cmd/pane/contextMenu]`
- Context Menu (Current Directory) [td] (`Alt+Shift+F10`) `[cmd/pane/contextMenuCurrentDirectory]`
- Security… [td] (`⊘`) `[cmd/pane/openSecurity]`
- ---
- Change Attributes… [td] (`Ctrl+F8`) `[cmd/pane/changeAttributes]`
- Change Case… (`Ctrl+F7`) `[cmd/pane/changeCase]`
- Pack… [td] (`Alt+F5`) `[cmd/pane/pack]`
- Unpack… [td] (`Alt+F6`) `[cmd/pane/unpack]`
- New >
  - Folder… (`F7`) `[cmd/pane/createDirectory]`
  - ---
  - *(Shell “New” templates, dynamic)*
    - `<Template Name>` [td] (`⊘`; parameterized: templateId) `[cmd/pane/newFromShellTemplate]`
- ---
- Exit (`Alt+F4`) `[cmd/app/exit]`

#### Edit (targets Focused pane unless explicitly stated)

- Cut [td] (`Ctrl+X`) `[cmd/pane/clipboardCut]`
- Copy (`Ctrl+C` target; also `Ctrl+Insert` default binding) `[cmd/pane/clipboardCopy]`
- Paste (`Ctrl+V` target; also `Shift+Insert` default binding) `[cmd/pane/clipboardPaste]`
- Paste Shortcut [td] (`⊘`) `[cmd/pane/clipboardPasteShortcut]`
- ---
- Copy Path + Name as Text [td] (`Alt+Insert`) `[cmd/pane/copyPathAndNameAsText]`
- Copy Name as Text [td] (`Alt+Shift+Insert`) `[cmd/pane/copyNameAsText]`
- Copy Path as Text [td] (`Ctrl+Alt+Insert`) `[cmd/pane/copyPathAsText]`
- Copy Path + File Name as Text [td] (`Ctrl+Shift+Insert`) `[cmd/pane/copyPathAndFileName]`
- ---
- Select… [td] (`⊘`) `[cmd/pane/selection/selectDialog]`
- Unselect… [td] (`⊘`) `[cmd/pane/selection/unselectDialog]`
- Invert Selection [td] (`⊘`) `[cmd/pane/selection/invert]`
- Select All (`Ctrl+A` target) `[cmd/pane/selection/selectAll]`
- Unselect All (`⊘`) `[cmd/pane/selection/unselectAll]`
- Restore Selection [td] (`⊘`) `[cmd/pane/selection/restore]`
- Select Next (`Insert`) `[cmd/pane/selectNext]`
- Select + Calculate Directory Size + Next (`Space`) `[cmd/pane/selectCalculateDirectorySizeNext]`
- Advanced >
  - Save Selection [td] (`Ctrl+Shift+F2`) `[cmd/pane/saveSelection]`
  - Load Selection… [td] (`Ctrl+Shift+F6`) `[cmd/pane/loadSelection]`
  - ---
  - Select Same Extensions [td] (`⊘`) `[cmd/pane/selection/selectSameExtension]`
  - Unselect Same Extensions [td] (`⊘`) `[cmd/pane/selection/unselectSameExtension]`
  - ---
  - Select Same Names [td] (`⊘`) `[cmd/pane/selection/selectSameName]`
  - Unselect Same Names [td] (`⊘`) `[cmd/pane/selection/unselectSameName]`
  - ---
  - Hide Selected Names [td] (`⊘`) `[cmd/pane/selection/hideSelectedNames]`
  - Hide Unselected Names [td] (`⊘`) `[cmd/pane/selection/hideUnselectedNames]`
  - Show Hidden Names [td] (`⊘`) `[cmd/pane/selection/showHiddenNames]`
  - ---
  - Go to Previous Selected Name [td] (`⊘`) `[cmd/pane/selection/goToPreviousSelectedName]`
  - Go to Next Selected Name [td] (`⊘`) `[cmd/pane/selection/goToNextSelectedName]`

#### Commands (targets Focused pane unless explicitly stated)

- Create Directory… (`F7`) `[cmd/pane/createDirectory]`
- Change Directory… (`Shift+F7`) `[cmd/pane/changeDirectory]` *(opens NavigationView address edit; mounted: `<instanceContext>|/path`)*
- Compare Directories… [td] (`Ctrl+F10`) `[cmd/app/compare]`
- Calculate Occupied Space (`Alt+F10`) `[cmd/pane/viewSpace]`
- Calculate Directory Sizes (`Ctrl+Shift+F10`) `[cmd/pane/calculateDirectorySizes]`
- Find Files and Directories… [td] (`Alt+F7`) `[cmd/pane/find]`
- Make File List… [td] (`⊘`) `[cmd/pane/makeFileList]`
- Go to Shortcut or Link Target [td] (`⊘`) `[cmd/pane/goToShortcutOrLinkTarget]`
- ---
- List of Opened Files [td] (`Alt+F11`) `[cmd/pane/listOpenedFiles]`
- Show Folders History (`Alt+F12`) `[cmd/pane/showFoldersHistory]` *(opens NavigationView history dropdown)*
- ---
- Connect Network Drive… (`F11`) `[cmd/pane/connect]` *(opens the Windows dialog; remote path is editable; when focused pane is File System browsing an UNC path (`\\\\...`), prefill remote name with the current path; otherwise open with no prefill; on success, if a new logical drive appears, navigate the focused pane to the new drive root)*
- Disconnect… (`F12`) `[cmd/pane/disconnect]` *(opens the Windows dialog; before opening, cancel any pending enumeration and clear DirectoryInfoCache (stops folder watchers) for the focused pane; if focused pane is a mapped network drive, preselect it; if the focused pane drive is removed, navigate to the default file system root)*
- Shared Directories… [td] (`Ctrl+Shift+F9`) `[cmd/pane/shares]`
- ---
- Command Shell (`⊘`) `[cmd/pane/openCommandShell]` *(opens system shell at focused pane path; mounted: opens at mount backing folder)*
- Quick Search / Command Line Input [td] (`Shift+Space`) `[cmd/pane/quickSearch]`
- Bring Current Directory to Command Line [td] (`Ctrl+Space`) `[cmd/pane/bringCurrentDirToCommandLine]`
- Bring Filename to Command Line [td] (`Ctrl+Enter`) `[cmd/pane/bringFilenameToCommandLine]`
- Pane Menu (`F10`) `[cmd/pane/menu]`
- Reread Associations [td] (`⊘`) `[cmd/app/rereadAssociations]`
- ---
- User Menu > [td]
  - *(User menu items, dynamic)*
    - `<User Menu Item>` [td] (`F9` opens the user menu root) `[cmd/pane/userMenu]`
- Open File Explorer >
  - Current Folder (`Shift+F3`) `[cmd/pane/openCurrentFolder]`
  - *(Known folders, fixed list; menu labels + icons come from Shell (localized display names + system icons))*
    - Desktop (`⊘`) `[cmd/app/openFileExplorerKnownFolder]`
    - Documents (`⊘`) `[cmd/app/openFileExplorerKnownFolder]`
    - Downloads (`⊘`) `[cmd/app/openFileExplorerKnownFolder]`
    - Pictures (`⊘`) `[cmd/app/openFileExplorerKnownFolder]`
    - Music (`⊘`) `[cmd/app/openFileExplorerKnownFolder]`
    - Videos (`⊘`) `[cmd/app/openFileExplorerKnownFolder]`
    - OneDrive (`⊘`) `[cmd/app/openFileExplorerKnownFolder]` *(disabled when not present)*

#### Plugins

- Plugin Manager… [td] (`⊘`) `[cmd/app/plugins/manage]`
- ---
- *(Installed plugins, dynamic)*
  - `<File System Plugin Name>` [td] (`⊘`; parameterized: pluginId) `[cmd/pane/selectFileSystemPlugin]`
  - Enable/Disable `<Plugin>` [td] (`⊘`; parameterized: pluginId) `[cmd/app/plugins/toggleEnabled]`
  - Configure `<Plugin>`… [td] (`⊘`; parameterized: pluginId) `[cmd/app/plugins/configure]`

#### View

- Theme > [td]
  - *(System high contrast indicator, dynamic system state)*
    - High Contrast (System) [td] (`⊘`; disabled when not active) `[cmd/app/theme/systemHighContrastIndicator]`
  - ---
  - System [td] (`⊘`; parameterized: `builtin/system`) `[cmd/app/theme/select]`
  - Light [td] (`⊘`; parameterized: `builtin/light`) `[cmd/app/theme/select]`
  - Dark [td] (`⊘`; parameterized: `builtin/dark`) `[cmd/app/theme/select]`
  - Rainbow [td] (`⊘`; parameterized: `builtin/rainbow`) `[cmd/app/theme/select]`
  - High Contrast (App) [td] (`⊘`; parameterized: `builtin/highContrast`) `[cmd/app/theme/select]`
  - ---
  - *(Theme files and user themes, dynamic)*
    - `<Theme Name>` [td] (`⊘`; parameterized: themeId) `[cmd/app/theme/select]`
- ---
- Toggle Fullscreen (`Ctrl+Shift+F11`) `[cmd/app/fullScreen]`
- Window Menu [td] (`Alt+Space`) `[cmd/pane/windowMenu]`
- Switch Pane Focus (`Tab`) `[cmd/pane/switchPaneFocus]`
- Pane > [td]
  - Show Hidden Files [td] (`⊘`; checkable) `[cmd/pane/viewOptions/toggleHiddenFiles]`
  - Show System Files [td] (`⊘`; checkable) `[cmd/pane/viewOptions/toggleSystemFiles]`
  - Show File Extensions [td] (`⊘`; checkable) `[cmd/pane/viewOptions/toggleFileExtensions]`
  - Show Thumbnails [td] (`⊘`; checkable) `[cmd/pane/viewOptions/toggleThumbnails]`
  - Show Preview Pane [td] (`⊘`; checkable) `[cmd/pane/viewOptions/togglePreviewPane]`
  - Show Filter Bar [td] (`⊘`; checkable) `[cmd/pane/viewOptions/toggleFilterBar]`
  - ---
  - Show Navigation Bar (Left) [td] (`⊘`; checkable, targets Left pane) `[cmd/pane/viewOptions/toggleNavigationBar]`
  - Show Navigation Bar (Right) [td] (`⊘`; checkable, targets Right pane) `[cmd/pane/viewOptions/toggleNavigationBar]`
  - ---
  - Show Status Bar (Left) [td] (`⊘`; checkable, targets Left pane) `[cmd/pane/viewOptions/toggleStatusBar]`
  - Show Status Bar (Right) [td] (`⊘`; checkable, targets Right pane) `[cmd/pane/viewOptions/toggleStatusBar]`
- Show Function Bar [td] (`⊘`; checkable) `[cmd/app/toggleFunctionBar]`
- Show Menu [td] (`⊘`; checkable) `[cmd/app/toggleMenuBar]`
- ---
- Preferences… (`⊘`) `[cmd/app/preferences]`

#### Right (pane menu: targets Right pane)

Right menu is identical to Left menu, except:
- Change Drive (`Alt+F2`) *(opens file-system drive menu; when pane is in a non-`file` plugin, the NavigationView menu also exposes a bottom “Change Drive” submenu)* `[cmd/app/openRightDriveMenu]`
- All `cmd/pane/*` entries target the Right pane.

#### Help (right-justified)

- Display Shortcuts… (`F1`) `[cmd/app/showShortcuts]`
- ---
- About… (`Alt+?`) `[cmd/app/about]`

##### Shortcuts window (`cmd/app/showShortcuts`)

- The window includes a **Search** edit at the top.
- Search is **case-insensitive** and filters rows by **command name**, **description**, or **shortcut text**.
- Matching substrings are highlighted in the list.

### Command details (Implemented)

#### Toggle Fullscreen (`cmd/app/fullScreen`)

- Invoking the command MUST toggle borderless fullscreen for the main window (hide title bar, cover the current monitor including taskbar).
- While fullscreen is active, pressing `Esc` MUST exit fullscreen.
- Invoking the command again MUST exit fullscreen.

#### View Width (`cmd/app/viewWidth`)

- Invoking the command MUST enter “view width adjust” mode for the main pane splitter.
- While active:
  - `Left` / `Right` arrows MUST nudge the splitter.
  - `Enter` MUST commit the new width.
  - `Esc` MUST cancel and restore the splitter ratio captured when the mode started.
- Invoking the command again while active MUST commit (same as `Enter`).

#### Calculate Directory Sizes (`cmd/pane/calculateDirectorySizes`)

- Invoking the command MUST open Space Viewer for the target pane’s **current folder**.
- The selected/current item MUST NOT change the target; the current folder is always used.
- If Space Viewer cannot be opened, the command MUST show a localized error and MUST do nothing else.

#### Change Case (`cmd/pane/changeCase`)

- Invoking the command MUST show a modal dialog:
  ```text
  Title: Change Case
  Change Case to
    o Lower case
    o Upper case
    o Partially mixed case (name in mixed, extension in lower)
    o Mixed case
  Change
    o Whole filename
    o Only name
    o Only extension
  Options
    [ ] Include subdirectories

    [ OK ] [ Cancel ]
  ```
- Scope: apply to `Selected items`; if no items are selected, apply to the `Current item`.
- “Include subdirectories”:
  - MUST be available for any filesystem plugin.
  - When enabled, traversal MUST be **non-recursive** (iterative) to avoid stack overflow on deep directory hierarchies.
  - Traversal MUST use the active plugin’s directory enumeration semantics (it MUST work with non-Windows / plugin-specific paths).

### Command/menu mapping status (Current implementation)

- Main menu structure is implemented in `RedSalamander/RedSalamander.rc` and follows the target top-level layout.
- Shortcut text in menus is dynamic (reflects the effective current bindings).

**Menu items still using pane-specific `WM_COMMAND` IDs (not `CommandRegistry`-mapped yet):**
- `Left/Right → Go to → *` (`IDM_LEFT_GO_TO_*` / `IDM_RIGHT_GO_TO_*`) including dynamic Hot Paths/History ranges.
- `Left/Right → Display → *` (`IDM_LEFT_DISPLAY_*` / `IDM_RIGHT_DISPLAY_*`)
- `Left/Right → Sort by → *` (`IDM_LEFT_SORT_*` / `IDM_RIGHT_SORT_*`)
- `Left/Right → Maximize/Restore Pane` (`IDM_LEFT_ZOOM_PANEL` / `IDM_RIGHT_ZOOM_PANEL`)
- `Left/Right → Filter…` (`IDM_LEFT_FILTER` / `IDM_RIGHT_FILTER`)
- `Left/Right → Refresh` (`IDM_LEFT_REFRESH` / `IDM_RIGHT_REFRESH`)
- `Commands → Open File Explorer → Known folders` (`IDM_APP_OPEN_FILE_EXPLORER_*`) — labels are Shell-localized; planned to route to parameterized `cmd/app/openFileExplorerKnownFolder`.
- Debug-only overlay sample entries: `Left/Right → Overlay Sample [dbg] → *` and `FolderView context → Overlay Sample [dbg] → *`.

**`cmd/*` commands whose non-zero `wmCommandId` does not appear in `RedSalamander/RedSalamander.rc` today (equivalent UI exists via pane-specific IDs or popups):**
- `cmd/pane/sort/none` (`IDM_PANE_SORT_NONE`) — main menu uses `IDM_LEFT_SORT_NONE` / `IDM_RIGHT_SORT_NONE`.
- `cmd/pane/sort/attributes` (`IDM_PANE_SORT_ATTRIBUTES`) — main menu uses `IDM_LEFT_SORT_ATTRIBUTES` / `IDM_RIGHT_SORT_ATTRIBUTES`.
- `cmd/pane/userMenu` (`IDM_PANE_USER_MENU`) — the Commands menu exposes a `User Menu` popup root; items are dynamic.

## Canonical Shortcut Map (Target)

This section documents the intended default bindings; the implementation may temporarily differ while shortcut customization is being built.

### Function Bar (Command Bar UI)

The application window includes a bottom **Function Bar** to make the current shortcut configuration discoverable:

- Height: `24 DIP`, full window width.
- Layout: `12` equal-width zones (F1..F12).
- Each zone displays:
  - A small key glyph (rounded rectangle) containing the function key (e.g. `F1`).
  - The localized command display name bound to that key for the **currently active modifier set**.
- Typography (default):
  - Function key glyph text: `7 DIP`
  - Command label text: `11 DIP`
- Modifier behavior:
  - While the user holds `Ctrl`, `Alt`, `Shift`, or their supported combinations, the Function Bar updates to show the bindings for that modifier set.
  - When a function key is pressed, its zone is highlighted.
- Optional modifier indicator:
  - A right-aligned indicator shows the currently held modifiers (e.g. `Ctrl`, `Shift`, `Ctrl+Shift`).
  - If there is not enough horizontal space, the modifier indicator is hidden.
- If the window is too small to display all content, text is truncated (no wrap).
- Mouse:
  - Hover highlights the zone.
  - Clicking a zone invokes the binding for the current modifier set.

All Function Bar bindings MUST be configurable in settings.

### Default Function Bar Bindings

`⊘` means “no shortcut assigned”.

| Key  | None            | Ctrl                     | Alt                         | Shift                | Ctrl+Shift                      | Alt+Shift                                 |
|------|-----------------|--------------------------|-----------------------------|----------------------|----------------------------------|-------------------------------------------|
| F1   | Shortcuts       | ⊘                        | Open Left Drive Menu        | ⊘                    | ⊘                                | ⊘                                         |
| F2   | Rename          | Sort None                | Open Right Drive Menu       | ⊘                    | Save Selection                   | ⊘                                         |
| F3   | View            | Sort by Name             | Alternate View              | Open Current Folder  | View Width                       | ⊘                                         |
| F4   | Edit            | Sort by Extension        | Exit                        | Edit New             | Edit Width                       | ⊘                                         |
| F5   | Copy            | Sort by Time             | Pack                        | ⊘                    | Save Selection                   | ⊘                                         |
| F6   | Move            | Sort by Size             | Unpack                      | ⊘                    | Load Selection                   | ⊘                                         |
| F7   | Make Directory  | Change Case              | Find                        | Change Directory     | ⊘                                | ⊘                                         |
| F8   | Delete          | Change Attributes        | ⊘                           | Permanent Delete     | ⊘                                | ⊘                                         |
| F9   | User Menu       | Refresh                  | Unpack                      | Hot Paths            | Shares                           | ⊘                                         |
| F10  | Menu            | Compare                  | Space View                  | Context Menu         | Calculate Directory Sizes        | Context Menu (Current Directory)          |
| F11  | Connect         | Zoom Panel               | List of Opened Files        | ⊘                    | Full Screen                      | ⊘                                         |
| F12  | Disconnect      | Filter                   | Show Folders History        | ⊘                    | ⊘                                | ⊘                                         |

### FolderView Configurable Shortcuts (Non-Function Bar)

These chords apply while focus is inside the `FolderWindow` (either `FolderView` or `NavigationView`).

Resolution order (normative):
- When a `FolderView` has focus: if the chord is bound in `settings.shortcuts.folderView`, the host MUST execute the bound command and MUST consume the key message.
- When focus is inside the `FolderWindow` but not in a `FolderView`: chords in `settings.shortcuts.folderView` are evaluated only when at least one modifier (Ctrl/Alt/Shift) is down and the key is not `Tab`; if bound, the host MUST execute the bound command and MUST consume the key message.
- Otherwise, the message continues through the normal routing pipeline (accelerators, then `FolderView`’s built-in key handling).

This means any key listed as a valid `vk` in `Specs/SettingsStoreSpec.md` (including Arrow keys / PageUp / PageDown / Home / End / `0`-`9`) can be made configurable by adding a binding entry; unbound chords keep their built-in behavior.

`⊘` means “no shortcut assigned”.

#### Default `shortcuts.folderView` bindings (implemented)

| Key       | None                               | Ctrl                             | Alt                      | Shift                              | Ctrl+Shift                        | Ctrl+Alt          | Alt+Shift             |
|-----------|------------------------------------|----------------------------------|--------------------------|------------------------------------|-----------------------------------|-------------------|-----------------------|
| Backspace | Up One Directory                   | ⊘                                | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| Tab       | Switch Pane Focus                  | ⊘                                | ⊘                        | Switch Pane Focus                  | ⊘                                 | ⊘                 | ⊘                     |
| U         | ⊘                                  | Swap Panes                        | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| A         | ⊘                                  | Select All                       | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| C         | ⊘                                  | Clipboard Copy                   | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| V         | ⊘                                  | Clipboard Paste                  | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| L         | ⊘                                  | Focus Address Bar                | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| D         | ⊘                                  | ⊘                                | Focus Address Bar        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| Up        | ⊘                                  | ⊘                                | Up One Directory         | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| Down      | ⊘                                  | ⊘                                | Show Folders History     | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| /         | ⊘                                  | ⊘                                | About                    | ⊘                                  | ⊘                                 | ⊘                 | About                 |
| 2         | ⊘                                  | ⊘                                | Display as Brief         | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| 3         | ⊘                                  | ⊘                                | Display as Detailed      | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| 4         | ⊘                                  | ⊘                                | Display as Extra Detailed | ⊘                                 | ⊘                                 | ⊘                 | ⊘                     |
| A..Z      | ⊘                                  | ⊘                                | ⊘                        | Go to Drive Root (`<drive>:\\`)    | ⊘                                 | ⊘                 | ⊘                     |
| Enter     | Execute / Open                     | Bring Filename to Command Line   | Open Properties          | ⊘                                  | Bring Filename to Command Line    | ⊘                 | ⊘                     |
| Space     | Select + Calc Dir Size + Next      | Bring Current Dir to Command Line | Window Menu              | Quick Search / Command Line Input  | Bring Current Dir to Command Line | ⊘                 | ⊘                     |
| Insert    | Select + Next                      | Clipboard Copy                   | Copy Path + Name as Text | Clipboard Paste                    | Copy Path + File Name as Text     | Copy Path as Text | Copy Name as Text     |
| Delete    | Move to Recycle Bin                | ⊘                                | ⊘                        | Permanent Delete (With Validation) | Permanent Delete (With Validation) | ⊘                 | ⊘                     |

Notes:
- Unmodified digit keys (`0`-`9`) and unmodified letter keys are unbound by default so they can be used for incremental search typing; planned Hot Path digit bindings are future work.

### Shortcut Customization UI (Preferences)

- The main menu includes `View → Preferences...` (near the bottom, separated).
- The settings dialog includes a `Shortcuts` tab with sections:
  - Function Bar shortcuts (F1..F12)
  - Folder view shortcuts (all supported keys)
- Settings are loaded at application startup; shortcut bindings are restored and applied before the first main window interaction.
- Editing model (example):

| Command Name              | Key | CTRL | ALT | SHIFT |
|---------------------------|-----|------|-----|-------|
| Rename                    | F2  |  X   |     |       |
| Alternate View            | F3  |      |  X  |       |
| Edit New                  | F4  |      |     |   X   |

- Conflicts MUST be detected and shown with a warning icon + tooltip (example: `Conflict with command 'Rename' (Ctrl + F2)`).
- A `Restore defaults` button resets shortcut configuration to the canonical defaults documented above.

### RedSalamander (main app)

**Menu bar**
- `Alt` (alone) temporarily shows the menu bar when hidden and starts menu interaction (see `Specs/SettingsStoreSpec.md`).

**Accelerators** (see `RedSalamander/RedSalamander.rc`)
- None. (Reserved for legacy / future use.)

**FolderView** (see `RedSalamander/FolderView.Interaction.cpp`)
- Arrow keys / Home / End: move `Current item` without changing selection state (focused item may be selected or not).
- `Page Up` / `Page Down`: horizontal paging by **visible columns** (layout is column-based)
- `Shift+Arrow`: range selection from anchor (existing)
- `Space`: select `Current item`, request folder subtree size computation (if folder), and advance to the next item
- `Insert`: select `Current item` and advance to the next item (no folder subtree size computation)
- `Ctrl+A`: select all
- `Ctrl+C`: copy `Selected items` to clipboard (or `Current item` when selection is empty)
- `Ctrl+V`: paste from clipboard
- `Enter`: open `Current item`
- `Backspace`: go to parent folder. When you are on a mount file system root, go to the parent folder of the mount point.
- `Delete`: delete `Selected items` (or `Current item` when selection is empty)
- `Shift + Delete`: do a permanent delete `Selected items` (or `Current item` when selection is empty) without going to recycle bin(user confirmation needed)
- `F2`: rename `Current item`
- `Tab` / `Shift+Tab`: move focus   between Pane `FolderView`s (no longer enters NavigationView)
- `Alt+D` / `Ctrl+L`: focus NavigationView address edit
- `Alt+Down`: open history dropdown
- `Alt+Up`: navigate to parent folder
- `Tab`: switch focus to the other pane’s `FolderView`.
  - `Shift+Tab`: same as `Tab` (two-pane toggle) unless later extended.

**NavigationView** (see `RedSalamander/NavigationView.Interaction.cpp`)
- `Tab` / `Shift+Tab`: cycle focus between visible regions (Menu → Path → History → Disk Info), then hand off to FolderView
- `Alt+D` / `Ctrl+L`: enter edit mode / focus address edit
- `Alt+Down`: open history dropdown
- `Enter` / `Space`: activate focused region

### RedSalamanderMonitor

**ColorTextView** (see `Specs/RedSalamanderMonitorSpec.md`)
- `Ctrl+F`: open find UI
- `F3`: find next
- `Ctrl+C`, `Ctrl+A`: selection/copy
- `Page Up` / `Page Down`: scroll


### Pane switching and top UI access


- `NavigationView` keeps **Tab** traversal inside its regions (existing), but `FolderView` no longer uses Tab to enter `NavigationView` (replaced by Alt+D/Ctrl+L for keyboard access to the address bar).

### Focused vs unfocused pane selection visuals

- In the **focused pane**:
  - `Current item` draws **border + background**.
  - `Selected items` use the active selection palette (`FolderViewTheme.itemBackgroundSelected` + `FolderViewTheme.textSelected`), plus the `Current item` border.
- In the **unfocused pane**:
  - `Current item` draws **border only** (no background fill).
  - `Selected items` remain visibly selected using a subtle inactive selection palette (`FolderViewTheme.itemBackgroundSelectedInactive` + `FolderViewTheme.textSelectedInactive`).
- If the `Current item` is also selected, the focus border must remain visible on top of the selection background (use a contrasting stroke).

### Function key operations (global to FolderWindow)

These keys target the **focused pane** as the source (unless stated otherwise):

- `F2`: Rename
- `F3`: View (open focused file in viewer; for folders behave like Enter)
- `F5`: Copy from focused pane → other pane
- `F6`: Move from focused pane → other pane
- `F7`: Create directory in focused pane
  - MUST show a modal dialog centered on the main window that prompts for the new folder name.
  - MUST display the destination path where the new folder will be created.
  - MUST validate the typed folder name and show a localized warning if it contains invalid characters (`\\ / : * ? " < > |`).
  - After a successful create, the newly created directory MUST become the `Current item` (focused/active) in the focused pane’s `FolderView` and be scrolled into view (so `Enter` opens it).
  - If directory creation is not supported for the current file system/plugin, the host MUST show a localized error message.
  - If the plugin `CreateDirectory` method returns `E_NOTIMPL`, the host MUST show a localized error message that includes the plugin display name.
- `F8`: Delete (equivalent to Delete key)

### Sorting shortcuts (existing)

- `Ctrl+F2`: Sort None (restore initial order)
- `Ctrl+F3`: Sort by Name
- `Ctrl+F4`: Sort by Extension
- `Ctrl+F5`: Sort by Time
- `Ctrl+F6`: Sort by Size

**Sort None semantics (normative)**
- Selecting **None** MUST set `view.sortBy` to `"none"` and restore the list to the initial order as it was presented for the current directory snapshot.
- In `"none"` mode, the host MUST apply no sort key; it MAY still keep stable grouping rules (e.g., directories-first) as long as switching back to None deterministically restores the initial order for that snapshot.

### Space selection + folder size accumulation

- **Space** (in `FolderView`):
  - Toggle selection state of the `Current item`
  - If the `Current item` is a folder, request folder subtree size computation for it.
  - Move `Current item` to the next item (Down; wraps or clamps per current navigation rules).
  - Update the pane status bar “selected bytes” to include:
    - File sizes directly.
    - Folder sizes computed by traversing all descendant folders (see below).
- Moving `Current item` with Arrow/Home/End/Page keys MUST NOT clear existing selections.
 
- **Insert** (in `FolderView`):
  - Toggle selection state of the `Current item`
  - Move `Current item` to the next item (Down; wraps or clamps per current navigation rules).
- Moving `Current item` with Arrow/Home/End/Page keys MUST NOT clear existing selections.
- For responsiveness, **folder subtree size computation is triggered only by the Space workflow**. Other selection changes (mouse selection, Insert, `Ctrl`/`Shift` range selection) MUST update selection counts immediately but MUST NOT start folder subtree size computation.

**Folder subtree traversal**
- Folder sizes MUST be computed asynchronously (background thread) and be cancelable when selection changes.
- While computing, the status bar MUST show a “calculating” state (exact text must be in `.rc` resources).
- While computing, the status bar MUST also display the **current bytes computed so far** and update periodically as the total increases.
- If folder sizes are not currently computed (because the user did not trigger Space, or because a computation was canceled), the status bar MUST show a localized **unknown size** placeholder for folder bytes (exact text must be in `.rc` resources).
- the size computation MUST:
  - Traverse all subfolders of the selected folder(s) using an **iterative** algorithm (explicit stack/queue; no call-stack recursion).
  - Sum file sizes only (ignore folder metadata size).
  - Handle access errors gracefully (skip inaccessible files/folders, log if needed).
- the size display is in italic when computation pending with an animated icon to be clear to computation is in progress. After computation completes, the size display returns to normal font.
- the size display MUST update incrementally as each folder’s size becomes available.
- the size computation MUST NOT block UI interaction.
- the size computation MUST re use cache where possible to avoid redundant work.
- Cancellation behavior:
  - If the selection changes (items added/removed), any ongoing computations for deselected items MUST be aborted.
  - New computations for newly selected folders MUST start promptly.
- the size display MUST remain accurate if the selection changes during computation:
  - If an item is deselected before its size is computed, its result MUST be discarded.
  - If an item is newly selected, its size MUST be computed and added to the total.
- The result MUST contribute to the selection’s total byte count once available.

### Incremental search (FolderView)

**Enter mode**
- When `FolderView` has focus and the user types a printable character, enter incremental search mode and append the character to the query.
- Incremental search mode exits when `FolderView` loses focus or its folder contents are refreshed.

**Search semantics**
- Match against the item display name (case-insensitive by default).
- Match is “contains” (substring), not prefix-only.
- If the `Current item` still matches the query, it stays.
- Otherwise, selecting a match moves the `Current item` to the first matching item after the current position (wrap allowed).

**Highlight**
- All **visible** items whose display name matches the query highlight the matched substring with a **selection-style background** (and selection text color) while in this mode (no font-weight change).
- Arrow keys navigate between matches without clearing the query (exact cycling rules are below).

**Keys while in mode**
- Printable character: extend query; if the `Current item` no longer matches, jump to the next match
- `Backspace`: remove last character; if query becomes empty, exit mode
- `Esc`: exit mode and clear highlight
- `Up` / `Left`: move to previous match (wrap allowed)
- `Down` / `Right`: move to next match (wrap allowed)
- Any “command/navigation” key (e.g., `Tab`, `Enter`, `Delete`, `F2`, `Home/End`, `Page Up/Down`) exits incremental search first, then performs the command

## Implementation Plan (Proposed)

1. **Command catalog + resources**
   - Add all planned `cmd/*` IDs to the command registry and ensure each has:
     - Localized display name + description in `.rc` STRINGTABLE.
     - A stable `WM_COMMAND` ID when it is invokable from the Win32 main menu.
   - Keep the registry sorted; do not introduce new command namespaces outside `cmd/app/*` and `cmd/pane/*`.

2. **Menu resource update (localization-first)**
   - Replace the main `MENUEX` definition with the target top-level order (`Left, Files, Edit, Commands, Plugins, View, Right, Help`).
   - Keep all static menu structure in `.rc`; runtime code only fills dynamic sections (history/hotpaths/themes/plugins/shell-driven lists).
   - Ensure `Help` is right-justified.

3. **Command routing + pane targeting**
   - Route `WM_COMMAND` to a single command executor that resolves `cmd/pane/*` target pane based on:
     - Explicit pane menu origin (Left/Right menus), else
     - Focused pane, else active pane.
   - Keep `WndProc` cases minimal and route to `On*` handlers (per AGENTS.md).

4. **Dynamic menus (safe + RAII)**
   - Implement dynamic menu rebuild for:
     - Left/Right `Go to` (Hot Paths + History)
     - `View With` / `Edit With`
     - `New` templates
     - Plugins list
     - Theme list
   - Use WIL RAII for all Win32 resources (HMENU, HBITMAP, HICON, etc.); no manual cleanup.

5. **Shortcut system alignment**
   - Ensure `.rc` accelerators and `ShortcutDefaults` match the canonical defaults in this spec.
   - Enforce text-edit safety (no app-level shortcuts while an edit control is focused).
   - Update Preferences → Shortcuts UI to expose all configurable bindings and detect conflicts.

6. **Validation**
   - Build with `/W4` and keep warnings at zero (except explicitly allowed infrastructure warnings).
   - Manual smoke-check: menu structure, right-justification, dynamic menus, shortcut display text, and correct pane targeting.

## Open Questions (Resolved)

1. **Tab inside NavigationView**
   - `FolderView`: `Tab` / `Shift+Tab` switches focus between Left/Right panes.
   - `NavigationView`: `Tab` / `Shift+Tab` cycles within NavigationView regions; when reaching the end, focus returns to the pane.
   - `Menu bar` (menu loop): `Tab` / `Shift+Tab` exits menu mode and returns focus to the active pane.

2. **Inactive selection visuals**
   - Keep the current behavior: `Selected items` in the unfocused pane retain a dim background fill using the inactive selection theme tokens.

3. **Copy/Move across plugins**
   - Default rule remains “same effective `IFileSystem` context” for cross-pane Copy/Move.
