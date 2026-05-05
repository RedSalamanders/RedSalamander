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
- **Incremental search mode**: a transient mode within `FolderView` entered by the Quick Search command or by typing printable characters.
- **Shortcut chord**: a key + modifier set (Ctrl/Alt/Shift) bound to a command ID (e.g. `cmd/...`).

## Key Routing Model (Normative)

### Priority order

1. **Menu loop**: when the Win32 menu loop is active, it owns the keyboard until it exits.
2. **Configurable shortcuts**: `settings.shortcuts` bindings are evaluated (Function Bar + FolderView).
3. **Accelerators**: application accelerators (TranslateAccelerator) run next and send `WM_COMMAND`.
4. **Focused control handlers**: `FolderView`, `NavigationView` and any edit controls handle remaining messages.

When the **menu loop** is active, pressing `Tab` (or `Shift+Tab`) MUST exit menu mode and return focus to the active pane (or the previously focused pane control).
When the **menu loop** is active, pressing `Alt` or `F10` MUST exit menu mode and return focus to the active pane (or the previously focused pane control).

When an **edit control** is focused, the host MUST bypass (2) and (3) and MUST NOT execute application-level accelerators or configurable shortcut bindings (text-edit safety).

Within the DxUi menu loop, keyboard-owned top-level and cascading popups MUST follow standard Windows directional behavior:
- Opening a root popup or submenu from the keyboard MUST move the keyboard highlight to the first navigable item in that newly opened popup.
- `Right` on a highlighted item with a submenu MUST open that submenu and move the keyboard highlight to its first navigable item.
- `Right` on a highlighted leaf item inside a top-level menu session MUST switch to the next enabled top-level menu and move the keyboard highlight to that popup's first navigable item.
- `Left` inside a submenu MUST close only the current submenu and restore the highlight to the parent item that opened it.
- `Left` in the root popup of a top-level menu session MUST switch to the previous enabled top-level menu and move the keyboard highlight to that popup's first navigable item.
- A stationary mouse pointer MUST NOT steal root switching or highlighted-item ownership from keyboard navigation; pointer-driven root switching and hover takeover require actual mouse movement.
- Opening a root popup by mouse MUST NOT synthesize a keyboard or hover selection from the checked item; checked, radio, and toggle state MUST be shown only by the item glyph until actual pointer movement or keyboard navigation selects an item.
- When a submenu is already open and the pointer moves back onto the parent item that opened that submenu, the submenu MUST remain open and any pending child-close timer MUST be canceled.
- When a submenu is already open and the pointer settles on a different sibling item that does not keep that submenu active, the existing child submenu chain MUST close after the standard cascade hover delay unless a replacement submenu opens instead.
- Exiting a DxUi top-level menu session without transferring focus to another control MUST restore keyboard focus to the pane/control that owned focus before menu mode started.
- Pressing `Escape` while a top-level menu bar, menu popup, or pane-owned context menu has keyboard ownership MUST dismiss that transient UI first, then restore keyboard focus to the active pane's `FolderView` unless the chosen command intentionally opens another focus-owning surface.

### Scope and focus

- Outside explicit child-window focus (edit controls, open dropdowns, or the Win32 menu loop), the active pane’s `FolderView` is the default keyboard owner.
- When the application or `FolderWindow` regains focus and no child window already has an intentional keyboard claim, the host MUST restore keyboard focus to the active pane’s most recent pane child; if none, restore the active pane’s `FolderView`.
- Shortcuts that target “the active pane” MUST resolve to the **focused pane** when focus is inside a pane; otherwise use the **active pane**.
- Shortcuts that act on file selection MUST prefer `Selected items`; if none are selected, they act on the `Current item` (and MAY implicitly select it for the operation).
- While an **edit control** is active (NavigationView address edit, rename edit, dialogs), the edit control owns the keyboard: application-level accelerators and configurable shortcut bindings MUST NOT execute (text-edit safety).
- Standard text-edit commands remain local to the control, including `Ctrl+Backspace` deleting the previous word/segment instead of invoking any pane/app shortcut.
- Mouse interaction with passive pane chrome MUST NOT permanently move keyboard focus away from the pane’s `FolderView`; only explicit keyboard entry into `NavigationView`, entering an edit control, or opening a keyboard-owned popup may take focus.
- `Escape` is the focus-reclaim key for the main file-manager window. If keyboard focus is in main-window chrome or transient menu UI instead of a `FolderView`, the first `Escape` MUST move focus to the active pane's `FolderView` and MUST NOT clear the pane selection. If focus is already inside the `FolderView`, the existing `FolderView` `Escape` behavior applies (for example, cancel incremental search or clear selection).
- `NavigationView` edit, suggestion, history, drive/menu, and full-path popup states MAY handle `Escape` locally, but their completed cancel/dismiss path MUST end with focus returned to the owning pane's `FolderView`.

### Command resolution (normative)

- **Shortcuts** and **menu items** map to commands identified by stable IDs (example shape: `cmd/...`).
- Command IDs MUST be in one of these namespaces:
  - `cmd/app/*`: application-global commands.
  - `cmd/pane/*`: pane-targeted commands; these MUST resolve to the **focused pane** when focus is inside a pane, otherwise the **active pane** (see rules above).
- Command display names MUST be localized resource strings (`.rc` STRINGTABLE). UI (Function Bar, settings dialog, tooltips) MUST NOT hardcode user-facing command names.
- If a shortcut is bound to a command that is not implemented at runtime, invoking it MUST show a localized message box stating it is not yet implemented and MUST do nothing else.
- Some commands/menu entries are **parameterized** (drive roots, hot paths, history paths, plugin/theme entries). For shortcuts the parameter is encoded in the command ID (e.g. `cmd/pane/goDriveRoot/C`, `cmd/pane/hotPath/1`) and is canonicalized for display/lookup; for menus the parameter is carried by dynamic menu-item ranges/payloads (or encoded in a command ID suffix for shortcut-like commands).

### Shortcut And Link Target Navigation

`cmd/pane/goToShortcutOrLinkTarget` operates on the target pane's current item. It is implemented for the built-in local file system when the current item is a `.lnk`, a `.url` whose URL resolves to a local path, a junction, a mount point, or a directory symbolic link.

- A target directory opens directly in the same pane.
- A target file opens its parent folder and restores focus to the target file when it is visible.
- Broken links, missing targets, non-local `.url` targets, unsupported reparse tags, and unsupported file-system plugins keep the pane in place and show localized pane feedback.
- The command records `shell.go_to_shortcut_target_us` in command selftests so shortcut resolution and navigation cost stay visible.

### Shell New Templates

`cmd/pane/newFromShellTemplate` is the stable command family for entries under **Files -> New** after **Folder**. The menu is populated at popup time from the Windows ShellNew registry view for local built-in file-system folders. The dynamic menu item carries the same template id that can be used by shortcuts as `cmd/pane/newFromShellTemplate/<templateId>`.

- Template ids are stable, sanitized ids derived from the extension and ShellNew template kind.
- The implementation supports ShellNew `NullFile`, `Data`, and `FileName` templates. ShellNew `Command` entries are intentionally not invoked by this safe implementation.
- Choosing a template prompts for the new file name in the current folder, prefilled with the template default name. The same filename validation rules used by Edit New apply.
- On success, the command creates the file, refreshes the pane, and focuses the created item when it is visible.
- If no local folder is active, no templates are available, a direct template id is stale, or creation fails, the pane remains in place and shows localized feedback.
- The command records `shellnew.enumerate_us`, `shellnew.menu_populate_us`, `shellnew.create_us`, and `shellnew.feedback_us` in command selftests.

### Clipboard File Commands

`cmd/pane/clipboardCopy`, `cmd/pane/clipboardCut`, `cmd/pane/clipboardPaste`, and `cmd/pane/clipboardPasteShortcut` share the standard Windows file-drop clipboard contract for local built-in file-system paths.

- Text edit controls keep ownership of ordinary text clipboard commands. When a navigation edit owns focus, `Ctrl+C`, `Ctrl+X`, and `Ctrl+V` MUST copy, cut, and paste text in that edit control before pane file commands are considered.
- `cmd/pane/clipboardCopy` writes selected or focused local file-system items as `CF_HDROP` with Preferred DropEffect `DROPEFFECT_COPY`.
- `cmd/pane/clipboardCut` writes selected or focused local file-system items as `CF_HDROP` with Preferred DropEffect `DROPEFFECT_MOVE`. It does not delete or move files immediately.
- `cmd/pane/clipboardPaste` copies clipboard file-drop paths into the current local folder using the file-operation copy path.
- `cmd/pane/clipboardPasteShortcut` reads clipboard file-drop paths and creates `.lnk` shortcuts in the current local folder. Shortcut names MUST be unique in the destination folder, the pane MUST refresh after creation, and the last created shortcut SHOULD become the focused item when visible.
- Unsupported providers, empty selections, clipboard contents without file paths, and shortcut creation failures keep the pane in place and show localized pane feedback instead of falling through to a generic not-implemented command.
- Command selftests MUST keep correctness and responsiveness visible with `clipboard.cut_us`, `clipboard.paste_shortcut_us`, and `clipboard.feedback_us` metrics.

### Quick Search

`cmd/pane/quickSearch` activates the target pane's integrated incremental search mode. It is not the persistent filter bar and it is separate from the command-line input commands.

- Invoking the command focuses the target pane's `FolderView`, enters search mode, clears any previous quick-search query, and shows the transient search indicator.
- Printable typing appends to the query. Matching is case-insensitive.
- The initial focused item prefers the first item whose name starts with the query. If no prefix match exists, the first item containing the query is focused.
- Rendering highlights the matching range for every visible item whose name contains the query.
- `Up`/`Left` and `Down`/`Right` navigate through all matching items in folder order while search mode is active.
- `Escape` exits search mode and clears the query. `Enter` exits search mode and activates the current focused item.
- No-match state remains non-modal: the query stays visible in the transient indicator and focus remains in the folder view.
- Command selftests MUST keep responsiveness visible with `quicksearch.activate_us`, `quicksearch.update_us`, `quicksearch.navigate_us`, and `render.incremental_search_effect_updates` metrics.

### Pane View Options

Pane view option commands target the focused pane, or the active pane when focus is outside both panes, unless a left/right menu item names a pane explicitly.

- `cmd/pane/viewOptions/toggleFileExtensions` toggles extension display in the target pane only. This is display-only: file operations, command-line insertion, clipboard actions, and plugin calls continue to use real item names and full paths.
- `cmd/pane/viewOptions/toggleThumbnails` is a legacy command id that selects the exclusive Thumbnails display mode in the target pane. The pane switches away from Brief/Detailed/Extra Detailed, uses larger DPI-aware item visuals, schedules bounded asynchronous thumbnail work for visible items, uses shell thumbnails when available, and renders the normal file/folder icon as fallback without blocking navigation. Repeating the command leaves the pane in Thumbnails; selecting another display mode leaves thumbnail mode and cancels stale work.
- `cmd/pane/viewOptions/togglePreviewPane` toggles preview mode for the active source pane and hosts the preview in the opposite pane. Opening preview shows compact themed DxUi Folder/Preview tabs at the top of the host pane, selects Preview, hides that pane's folder view while Preview is selected, and updates the embedded viewer preview when the source pane focus or selection changes. The tabs must behave as real pointer targets without stealing keyboard focus from the source pane. Preview tabs use attached, Visual Studio-like chrome: inactive tabs have no border, selected tabs blend into the pane below with square lower corners, the Folder tab tooltip displays the host pane path after the standard hover delay, and the Preview tab close glyph is visible when Preview is selected or hovered and closes preview mode. Preview resolves the configured viewer plugin for the focused item and uses it when the plugin supports embedded hosting; when saved viewer associations are missing or only resolve the default text viewer, preview uses the built-in embedded viewer defaults before falling back to the embedded text viewer or localized placeholder. If a focus change resolves to the same embedded viewer plugin already hosted by Preview, the host reuses that viewer instance and refreshes it with the new open context; it replaces the preview window only when resolution chooses a different plugin or refresh fails. Embedded preview viewers MUST NOT take keyboard focus from the source pane. Closing or replacing an embedded viewer persists changed plugin configuration, including ViewerVLC volume/mute state, and preview resolution/fallback choices are logged for monitor diagnostics. Switching back to Folder keeps preview mode open with the host folder view visible. Closing preview removes the tabs and restores the host pane. The preview area extends to the function bar, or to the bottom of the window when the function bar is hidden.
- `cmd/pane/viewOptions/toggleFilterBar` toggles a persistent themed DxUi filter bar for the target pane. The bar reflects the current `cmd/pane/filter` state, follows restored per-history filters, and never replaces Quick Search.
- `cmd/pane/viewOptions/toggleNavigationBar` toggles the target pane navigation/address bar. Left/right menu entries target their named pane; shortcut routing targets the active pane. Commands that focus the address bar MUST show the bar first, then focus the address edit.
- `cmd/pane/viewOptions/toggleStatusBar` routes shortcut invocation to the active pane and shares the existing left/right `Show` menu status-bar implementation.
- These visibility states are persisted per pane through `folders.items[].view.*` settings and MUST keep menu check marks synchronized with the current pane state.
- Command selftests MUST keep correctness and responsiveness visible for setting round-trip, menu labels, active/explicit pane routing, focus fallback, restored filters, and pane-view-option toggle latency.

### Command-Line Input

`cmd/pane/bringCurrentDirToCommandLine` and `cmd/pane/bringFilenameToCommandLine` open a pane-scoped command-line input that is separate from Quick Search and the navigation address edit.

- The command-line input appears above the function bar and below the pane area, receives keyboard focus, and is associated with the pane that invoked it.
- `cmd/pane/bringCurrentDirToCommandLine` appends the active local folder path using command-line quoting.
- `cmd/pane/bringFilenameToCommandLine` appends the focused item display name when no explicit selection exists. When one or more items are selected, it appends full local item paths; if the focused item is part of the selection, that focused path is first and the rest stay in pane order.
- Insertions happen at the current caret/selection and add a single separating space when adjacent text would otherwise touch.
- Pressing `Enter` executes the current text through the system command processor with the pane's current local folder as working directory, then clears and hides the input after a successful launch.
- Pressing `Escape` hides the input and restores folder-view focus without changing pane selection.
- Unsupported providers, missing local folders, and empty item scope keep the pane in place and show localized pane feedback.
- Command selftests MUST keep responsiveness visible with `commandline.focus_to_visible_us`, `commandline.insert_current_dir_us`, `commandline.insert_filename_us`, `commandline.launch_us`, and `commandline.feedback_us` when an error path is exercised.

### Reread Associations

`cmd/app/rereadAssociations` reloads settings-backed associations and action menus on demand without restarting the application.

- The command is application-scoped and is available from **Commands -> Reread Associations**.
- It reads the current settings file through the non-destructive hot-reload path. Invalid JSON, unsupported schema, or unreadable settings keep the current runtime settings in place and show the localized invalid-reload alert.
- Disk settings are authoritative for `fileActions` viewer/editor actions and associations, User Menu actions, file-system extension mappings, plugin settings, shortcuts, and other persisted preference sections, but current pane folders and already-open window placement are preserved.
- After a successful reload, the app rebuilds dynamic View With, Edit With, User Menu, ShellNew, and file-system-plugin menus, clears normal and association icon caches, refreshes both panes, preserves the active pane, and notifies settings-reload participants.
- Stale dynamic menu ids from before the reload MUST NOT launch old actions after the menus are rebuilt.
- Command selftests MUST keep correctness and responsiveness visible with `rereadAssociations.total_us`, menu-rebuild assertions, pane-refresh assertions, icon-association cache clearing, and preserved live pane paths.

### Make File List

`cmd/pane/makeFileList` opens a pane-scoped options dialog that generates a list from the focused pane.

- The command is available for local file-system folders. Unsupported providers keep the pane in place and show localized pane feedback.
- Source options are selected/focused items or the current folder. When no item is selected, the focused item is used. Current-folder mode enumerates the active folder contents.
- The recursive option descends into folders. `includeDirectories` controls whether directory rows appear in the generated list.
- Output formats are JSON, CSV, and text. JSON and CSV use the selected field flags (`includeName`, `includeFullPath`, `includeSize`, `includeModified`, and `includeAttributes`). Text uses `textMacro`.
- Text macros are case-insensitive and include `{filename}`, `{name}`, `{fullPath}`, `{path}`, `{size}`, `{modified}`, `{attributes}`, and `{isDirectory}`. `{{` and `}}` emit literal braces.
- Output targets are the clipboard or a UTF-8 file. File output requires a non-empty output path.
- Generated entries are deterministic and sorted by full path. CSV quotes commas, quotes, and newlines. JSON includes `format`, `count`, and `entries`.
- The last selected options are persisted in `settings.makeFileList`.
- Command selftests MUST keep correctness and responsiveness visible with `makeFileList.collect_us`, `makeFileList.generate_us`, `makeFileList.output_us`, `makeFileList.total_us`, and an archived JSON/CSV/text output artifact set.

### List Opened Files

`cmd/pane/listOpenedFiles` opens a modeless application dialog listing files RedSalamander currently has open.

- The command is available from **Commands -> List of Opened Files** and defaults to `Alt+F11`.
- Rows include internal viewer windows, external viewer/editor launches started by View/Edit/User Menu action plumbing, and the active Preview pane item.
- Each row shows the display file name, source (`Viewer`, `Editor`, or `Preview Pane`), opener/action name when known, and the full path.
- External process rows with captured handles are pruned once the process exits. Viewer rows are removed when their viewer window closes. Preview rows follow the current preview item and disappear when Preview is closed.
- When no rows exist, the dialog remains visible and shows a localized empty state.
- Double-click and **Focus Item** navigate the owning pane to the row path when the item is still reachable, then focus/select it in the folder view.
- The dialog MUST be a modeless DxUi-hosted app window, inherit the active app theme and themed window chrome, and avoid visible native dialog-template controls in DxUi mode.
- Command selftests MUST keep correctness and responsiveness visible with themed DxUi-host coverage, viewer/editor/preview source coverage, closed-process pruning, focus navigation, `listOpenedFiles.open_us`, and an archived `list_opened_files_metrics.json` artifact.

### Shared Directories

`cmd/pane/shares` opens a modeless application dialog listing local Windows disk shares from the focused pane context.

- The command is available from **Commands -> Shared Directories** and defaults to `Ctrl+Shift+F9`.
- Rows are sorted by share name and show share name, local path, share type, and remark.
- Only disk-tree shares are listed. Reachable local paths enable **Open Path**; unreachable paths stay visible but cannot be opened.
- **Open Path** switches the target pane to the built-in local file-system provider if needed, then navigates to the selected share's local path.
- **Manage** launches the Windows Shared Folders management console. Launch failure shows localized nonfatal pane feedback.
- Access denied while enumerating shares keeps the dialog open, clears stale rows, and shows a localized access-denied empty/error state.
- Command selftests MUST keep correctness and responsiveness visible with synthetic provider rows, sorted display, open-path navigation, access-denied state coverage, `sharedDirectories.open_us`, and an archived `shared_directories_metrics.json` artifact.

### Archive Pack And Unpack

`cmd/pane/pack` creates an archive from the active pane selection. `cmd/pane/unpack` extracts selected or focused ZIP archives to a destination selected in the app-owned Unpack prompt.

- Both commands are available only from local file-system folders. Unsupported providers keep the pane in place and show localized pane feedback.
- Pack uses selected items, or the focused item when nothing is selected. The built-in `ZIP (Plugin)` packer writes a deterministic stored ZIP archive. Directory entries use `/` separators, selected empty directories are preserved, and entries are sorted by archive path.
- In interactive use, Pack MUST use the app-owned DxUi Pack prompt rather than the stock save-file dialog. The prompt MUST suggest a non-conflicting archive path in the current folder, MUST list ZIP plus update-capable formats discovered from the bundled `7zip.dll`, and MUST update the suggested archive extension when the selected packer changes.
- When a 7-Zip packer is selected, Pack creates the archive through `IOutArchive::UpdateItems`. The delete-after-packing option MUST be off by default and MUST remove selected local sources only after archive creation succeeds and after the user accepts the permanent-delete confirmation prompt.
- Test/debug automation may supply a ZIP path and overwrite policy directly.
- Unpack supports stored ZIP entries through the built-in reader and delegates compressed ZIP entries, 7-Zip archives, and other formats supported by the bundled `7zip.dll` to the 7-Zip extraction path while preserving the same destination, overwrite, mask, and safe-entry-path contract. Unsupported/encrypted methods fail with localized pane feedback.
- Unpack MUST reject archives whose declared or observed decompressed payload exceeds 4 GiB for a single entry or 8 GiB total for the selected extraction set. Built-in ZIP central-directory validation and the 7-Zip extraction path both fail this case with `ERROR_FILE_TOO_LARGE`, before committing any oversized output file.
- Stored ZIP filename decoding MUST honor the UTF-8 general-purpose bit and otherwise use CP437 for legacy ZIP names. Compressed ZIP filename decoding is delegated to the bundled 7-Zip path.
- In interactive use, Unpack MUST use the app-owned DxUi Unpack prompt rather than the stock pick-folder dialog. The prompt MUST suggest a non-conflicting destination folder derived from the focused archive name, MUST expose the currently supported `ZIP (Plugin)` unpacker, MUST default the file mask to `*.*`, MUST show mask syntax help from the shared wildcard-mask contract, and MUST leave delete-after-unpacking off by default.
- The Unpack file mask uses the same wildcard syntax as Select/Unselect and pane Filter. `*.*`, `*`, or an empty mask extracts all safe entries. Other masks match either the archive entry path or the entry filename; unmatched files and empty directory entries are skipped.
- When enabled, delete-after-unpacking MUST remove the selected archive files only after extraction succeeds and after the user accepts the permanent-delete confirmation prompt.
- Test/debug automation may supply the destination and overwrite policy directly.
- Existing archive outputs or extracted files are preserved when overwrite is disabled. Failures report the relevant HRESULT and do not fall through to the generic "not implemented" message.
- Command selftests MUST keep correctness and responsiveness visible with Pack prompt coverage, Unpack prompt/mask/delete-after coverage, 7z creation coverage, stored-ZIP round-trip coverage, compressed-ZIP extraction coverage, CP437 non-ASCII filename coverage, decompressed-size rejection coverage, overwrite validation, invalid destination/path validation, unsupported-provider feedback, `archive.pack_us`, `archive.unpack_us`, `archive.feedback_us`, and an archived `archive_commands_metrics.json` artifact.

### Canonical Command IDs

This section is the single source of truth for the command ID catalog.

**Application commands (`cmd/app/*`)**
- `cmd/app/about`
- `cmd/app/exit`
- `cmd/app/externalHelp`
- `cmd/app/openLeftDriveMenu`
- `cmd/app/openRightDriveMenu`
- `cmd/app/compare`
- `cmd/app/fullScreen`
- `cmd/app/openFileExplorerKnownFolder` *(planned, parameterized: knownFolderId)*
- `cmd/app/preferences`
- `cmd/app/showShortcuts`
- `cmd/app/swapPanes`
- `cmd/app/toggleFunctionBar`
- `cmd/app/toggleMenuBar`
- `cmd/app/viewWidth`
- `cmd/app/rereadAssociations`
- `cmd/app/theme/select` *(parameterized: themeId)*
- `cmd/app/theme/selectNext`
- `cmd/app/theme/selectPrev`
- `cmd/app/theme/systemHighContrastIndicator`
- `cmd/app/plugins/manage`
- `cmd/app/plugins/toggleEnabled` *(planned, parameterized: pluginId)*
- `cmd/app/plugins/configure` *(planned, parameterized: pluginId)*

**Pane commands (`cmd/pane/*`)**
- `cmd/pane/historyBack`
- `cmd/pane/historyForward`
- `cmd/pane/goDriveRoot` *(parameterized: driveLetter)*
- `cmd/pane/hotPath` *(parameterized: digit `1..9` and `0` for slot 10)*
- `cmd/pane/setHotPath` *(parameterized: digit `1..9` and `0` for slot 10)*
- `cmd/pane/goRootDirectory`
- `cmd/pane/setPathFromOtherPane`
- `cmd/pane/navigatePath` *(planned, parameterized: path)*
- `cmd/pane/selectFileSystemPlugin` *(parameterized: pluginId)*
- `cmd/pane/bringCurrentDirToCommandLine`
- `cmd/pane/bringFilenameToCommandLine`
- `cmd/pane/clipboardCut`
- `cmd/pane/clipboardCopy`
- `cmd/pane/clipboardPaste`
- `cmd/pane/clipboardPasteShortcut`
- `cmd/pane/copyNameAsText`
- `cmd/pane/copyUncPathAndNameAsText`
- `cmd/pane/copyPathAndNameAsText`
- `cmd/pane/copyPathAsText`
- `cmd/pane/executeOpen`
- `cmd/pane/moveToRecycleBin`
- `cmd/pane/openCurrentFolder`
- `cmd/pane/openProperties`
- `cmd/pane/openSecurity`
- `cmd/pane/quickSearch`
- `cmd/pane/selectCalculateDirectorySizeNext`
- `cmd/pane/selectNext`
- `cmd/pane/switchPaneFocus`
- `cmd/pane/upOneDirectory`
- `cmd/pane/windowMenu`
- `cmd/pane/alternateView`
- `cmd/pane/changeAttributes`
- `cmd/pane/changeCase`
- `cmd/pane/changeDirectory`
- `cmd/pane/connect`
- `cmd/pane/contextMenu`
- `cmd/pane/contextMenuCurrentDirectory`
- `cmd/pane/disconnect`
- `cmd/pane/edit`
- `cmd/pane/editWith` *(parameterized: editorId)*
- `cmd/pane/editNew`
- `cmd/pane/alternateEdit`
- `cmd/pane/filter`
- `cmd/pane/find`
- `cmd/pane/hotPaths`
- `cmd/pane/hotPath` *(parameterized: digit `1..9` and `0`)*
- `cmd/pane/setHotPath` *(parameterized: digit `1..9` and `0`)*
- `cmd/pane/listOpenedFiles`
- `cmd/pane/showFoldersHistory`
- `cmd/pane/makeFileList`
- `cmd/pane/menu`
- `cmd/pane/pack`
- `cmd/pane/permanentDelete`
- `cmd/pane/refresh`
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
- `cmd/pane/sort/attributes`
- `cmd/pane/sort/extension`
- `cmd/pane/sort/name`
- `cmd/pane/sort/size`
- `cmd/pane/sort/time`
- `cmd/pane/view`
- `cmd/pane/viewWith` *(parameterized: viewerId)*
- `cmd/pane/viewSpace`
- `cmd/pane/newFromShellTemplate` *(parameterized: templateId)*
- `cmd/pane/selection/selectDialog`
- `cmd/pane/selection/unselectDialog`
- `cmd/pane/selection/invert`
- `cmd/pane/selection/selectAll`
- `cmd/pane/selection/unselectAll`
- `cmd/pane/selection/restore`
- `cmd/pane/selection/save`
- `cmd/pane/selection/selectSameExtension`
- `cmd/pane/selection/unselectSameExtension`
- `cmd/pane/selection/selectSameName`
- `cmd/pane/selection/unselectSameName`
- `cmd/pane/selection/hideSelectedNames`
- `cmd/pane/selection/hideUnselectedNames`
- `cmd/pane/selection/showHiddenNames`
- `cmd/pane/selection/goToPreviousSelectedName`
- `cmd/pane/selection/goToNextSelectedName`
- `cmd/pane/goToShortcutOrLinkTarget`
- `cmd/pane/openCommandShell`
- `cmd/pane/viewOptions/toggleHiddenFiles`
- `cmd/pane/viewOptions/toggleSystemFiles`
- `cmd/pane/viewOptions/toggleFileExtensions`
- `cmd/pane/viewOptions/toggleThumbnails`
- `cmd/pane/viewOptions/togglePreviewPane`
- `cmd/pane/viewOptions/toggleFilterBar`
- `cmd/pane/viewOptions/toggleNavigationBar`
- `cmd/pane/viewOptions/toggleStatusBar`

## Main Menu Bar (Target)

### Requirements (Normative)

- The menu bar structure and static labels MUST be defined in `.rc` resources (`RedSalamander/RedSalamander.rc`) to support localization (see `Specs/Core/Core_Localization.md`).
- Each menu item that triggers application behavior MUST map to a `cmd/*` command ID (shown in brackets below).
  - If the menu item is dynamic and requires a parameter (history path, hot path, plugin ID, theme ID), the menu item MUST still map to a stable `cmd/*` command ID; the parameter is carried in the menu item payload.
- The displayed shortcut text (when present) MUST reflect the effective current bindings (default or user-customized).
- Top-level menu order MUST be:
  - `Left`, `Files`, `Edit`, `Commands`, `Plugins`, `View`, `Right`, `Help`
- `Help` MUST be right-justified (appear at the right edge of the menu bar).

### Debug Self-Test Contract

- The default Commands self-test suite MUST prefer deterministic, local-only scenarios over environment-dependent integration.
- Command registry coverage MUST validate canonical command IDs only; removed command IDs are not preserved as aliases.
- Shortcut-default coverage MUST assert fixed high-value bindings directly, including the full `Insert` row for copy/paste/copy-as-text commands and the `Ctrl+F2..F6` sort bindings.
- Menu-contract coverage MUST assert the `Edit` menu copy-text group order, labels, separator boundaries, and text-only icon policy.
- Command behavior coverage for copy-text commands MUST run in a temp local folder with clipboard assertions and MUST stay separate from selection save/restore scenarios.
- The global dispatch smoke test remains a smoke test: it verifies that commands do not wedge the UI or leak transient windows, but it is not a substitute for behavior assertions.

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
  - Back (`Alt+Left`) *(History Back)* `[cmd/pane/historyBack]`
  - Forward (`Alt+Right`) *(History Forward)* `[cmd/pane/historyForward]`
  - Parent Directory (`Backspace`) `[cmd/pane/upOneDirectory]`
  - Root Directory (`Shift+Backspace`) `[cmd/pane/goRootDirectory]`
  - Path from Other Panel (`Ctrl+.`) `[cmd/pane/setPathFromOtherPane]`
  - ---
  - Hot Paths… (`Shift+F9`) `[cmd/pane/hotPaths]`
  - *(Hot Paths section, dynamic — from settings; navigates to the stored slot path)*
    - `<Hot Path>` (`⊘`)
  - ---
  - *(History section, dynamic)*
    - `<History Path>` (`⊘`)
- ---
- Brief (`Alt+2`) `[cmd/pane/display/brief]`
- Detailed (`Alt+3`) `[cmd/pane/display/detailed]`
- Extra Detailed (`Alt+4`) `[cmd/pane/display/extraDetailed]`
- Thumbnails (`Alt+5`; radio display mode, targets Left pane) `[cmd/pane/viewOptions/toggleThumbnails]`
- Preview Pane (`Alt+6`; checkable, source is Left pane, preview host is Right pane) `[cmd/pane/viewOptions/togglePreviewPane]`
- ---
- Sort By >
  - None (`Ctrl+F2`) `[cmd/pane/sort/none]`
  - Name (`Ctrl+F3`) `[cmd/pane/sort/name]`
  - Extension (`Ctrl+F4`) `[cmd/pane/sort/extension]`
  - Time (`Ctrl+F5`) `[cmd/pane/sort/time]`
  - Size (`Ctrl+F6`) `[cmd/pane/sort/size]`
  - Attributes (`⊘`) `[cmd/pane/sort/attributes]`
- Show >
  - Hidden Files (`⊘`; checkable, global visibility setting) `[cmd/pane/viewOptions/toggleHiddenFiles]`
  - System Files (`⊘`; checkable, global visibility setting) `[cmd/pane/viewOptions/toggleSystemFiles]`
  - File Extensions (`⊘`; checkable, targets Left pane) `[cmd/pane/viewOptions/toggleFileExtensions]`
  - ---
  - Filter Bar (`⊘`; checkable, targets Left pane) `[cmd/pane/viewOptions/toggleFilterBar]`
  - Navigation Bar (`⊘`; checkable, targets Left pane) `[cmd/pane/viewOptions/toggleNavigationBar]`
  - Status Bar (`⊘`; checkable, targets Left pane) `[cmd/pane/viewOptions/toggleStatusBar]`
- ---
- Maximize/Restore Pane (`Ctrl+F11`) *(toggle: move splitter to edge; restore only if splitter wasn't dragged while maximized; state persisted in settings)* `[cmd/pane/zoomPanel]`
- Swap Panes (`Ctrl+U`) *(swap Left/Right pane file system + current folder; view options stay with the pane; global history unaffected)* `[cmd/app/swapPanes]`
- Path from Other Panel (`Ctrl+.`) `[cmd/pane/setPathFromOtherPane]`
- Refresh (`Ctrl+F9`) *(invalidate directory cache + re-enumerate current folder)* `[cmd/pane/refresh]`
- Filter… (`Ctrl+F12`) *(open pane filter dialog; wildcard mask syntax shared with Select/Unselect; history saved; active filter shows a subtle background watermark; filter state restored when navigating to a path from history)* `[cmd/pane/filter]`

#### Files (targets Focused pane unless explicitly stated)

- Rename… (`F2`) `[cmd/pane/rename]`
- Open / Execute (`Enter`) `[cmd/pane/executeOpen]`
- View (`F3`) `[cmd/pane/view]`
- View Width… (`Ctrl+Shift+F3`) `[cmd/app/viewWidth]`
- Alternate View (`Alt+F3`) `[cmd/pane/alternateView]`
- View With >
  - *(Viewer list, dynamic)*
    - `<Viewer Name>` (`⊘`; parameterized: viewerId) `[cmd/pane/viewWith]`
- Edit (`F4`) `[cmd/pane/edit]`
- Alternate Edit (`Ctrl+Shift+F4`) `[cmd/pane/alternateEdit]`
- Edit With >
  - *(Editor list, dynamic)*
    - `<Editor Name>` (`⊘`; parameterized: editorId) `[cmd/pane/editWith]`
- Edit New File… (`Shift+F4`) `[cmd/pane/editNew]`
- Copy… (`F5`) `[cmd/pane/copyToOtherPane]`
- Move/Rename… (`F6`) `[cmd/pane/moveToOtherPane]`
- Delete >
  - Delete (`F8`) `[cmd/pane/delete]`
  - Move to Recycle Bin (`Del`) `[cmd/pane/moveToRecycleBin]`
- Permanent Delete (`Shift+F8` / `Shift+Del`) `[cmd/pane/permanentDelete]`
- Properties (`Alt+Enter`) `[cmd/pane/openProperties]`
- Context Menu (`Shift+F10`) `[cmd/pane/contextMenu]`
- Context Menu (Current Directory) (`Alt+Shift+F10`) `[cmd/pane/contextMenuCurrentDirectory]`
- Security… (`⊘`) `[cmd/pane/openSecurity]`
- ---
- Change Attributes… (`Ctrl+F8`) `[cmd/pane/changeAttributes]`
- Change Case… (`Ctrl+F7`) `[cmd/pane/changeCase]`
- Pack… (`Alt+F5`) `[cmd/pane/pack]`
- Unpack… (`Alt+F6`) `[cmd/pane/unpack]`
- New >
  - Folder… (`F7`) `[cmd/pane/createDirectory]`
  - ---
  - *(Shell “New” templates, dynamic)*
    - `<Template Name>` (`⊘`; parameterized: templateId) `[cmd/pane/newFromShellTemplate]`
- ---
- Exit (`Alt+F4`) `[cmd/app/exit]`

#### Edit (targets Focused pane unless explicitly stated)

- Cut (`Ctrl+X`) `[cmd/pane/clipboardCut]`
- Copy (`Ctrl+C` target; also `Ctrl+Insert` default binding) `[cmd/pane/clipboardCopy]`
- Paste (`Ctrl+V` target; also `Shift+Insert` default binding) `[cmd/pane/clipboardPaste]`
- Paste Shortcut (`⊘`) `[cmd/pane/clipboardPasteShortcut]`
- ---
- Copy Path + Name as Text (`Alt+Insert`) `[cmd/pane/copyPathAndNameAsText]`
- Copy Name as Text (`Alt+Shift+Insert`) `[cmd/pane/copyNameAsText]`
- Copy Path as Text (`Ctrl+Alt+Insert`) `[cmd/pane/copyPathAsText]`
- Copy UNC Path + Name as Text (`Ctrl+Shift+Insert`) `[cmd/pane/copyUncPathAndNameAsText]`
- Note: this resolves mapped drives to their provider UNC path and local file-system paths to `\\<machine>\<drive>$\...` when available.
- Note: `Name` means filename plus extension.
- ---
- Select… (`Ctrl+<key left of Backspace>`) `[cmd/pane/selection/selectDialog]`
- Unselect… (`Ctrl+<key right of 0>`) `[cmd/pane/selection/unselectDialog]`
- Invert Selection (`⊘`) `[cmd/pane/selection/invert]`
- Select All (`Ctrl+A` target) `[cmd/pane/selection/selectAll]`
- Unselect All (`Esc`) `[cmd/pane/selection/unselectAll]`
- Restore Selection (`Ctrl+Shift+F6`) `[cmd/pane/selection/restore]`
- Select Next (`Insert`) `[cmd/pane/selectNext]`
- Select + Calculate Directory Size + Next (`Space`) `[cmd/pane/selectCalculateDirectorySizeNext]`
- Advanced >
  - Save Selection (`Ctrl+Shift+F5`) `[cmd/pane/selection/save]`
  - Load Selection… (`Ctrl+Shift+F6`) `[cmd/pane/selection/restore]`
  - ---
  - Select Same Extensions (`Ctrl+Shift+<key left of Backspace>`) `[cmd/pane/selection/selectSameExtension]`
  - Unselect Same Extensions (`Ctrl+Shift+<key right of 0>`) `[cmd/pane/selection/unselectSameExtension]`
  - ---
  - Select Same Names (`⊘`) `[cmd/pane/selection/selectSameName]`
  - Unselect Same Names (`⊘`) `[cmd/pane/selection/unselectSameName]`
  - ---
  - Hide Selected Names (`⊘`) `[cmd/pane/selection/hideSelectedNames]`
  - Hide Unselected Names (`⊘`) `[cmd/pane/selection/hideUnselectedNames]`
  - Show Hidden Names (`⊘`) `[cmd/pane/selection/showHiddenNames]`
  - ---
  - Go to Previous Selected Name (`Alt+Up`) `[cmd/pane/selection/goToPreviousSelectedName]`
  - Go to Next Selected Name (`Alt+Down`) `[cmd/pane/selection/goToNextSelectedName]`

#### Commands (targets Focused pane unless explicitly stated)

- Create Directory… (`F7`) `[cmd/pane/createDirectory]`
- Change Directory… (`Shift+F7`) `[cmd/pane/changeDirectory]` *(opens NavigationView address edit; mounted: `<instanceContext>|/path`)*
- Compare Directories… (`Ctrl+F10`) `[cmd/app/compare]`
- Calculate Occupied Space (`Alt+F10`) `[cmd/pane/viewSpace]`
- Find Files and Directories… (`Alt+F7`) `[cmd/pane/find]`
- Make File List… (`⊘`) `[cmd/pane/makeFileList]`
- Go to Shortcut or Link Target (`⊘`) `[cmd/pane/goToShortcutOrLinkTarget]`
- ---
- List of Opened Files (`Alt+F11`) `[cmd/pane/listOpenedFiles]`
- Show Folders History (`Alt+F12`) `[cmd/pane/showFoldersHistory]` *(opens NavigationView history dropdown)*
- ---
- Connect Network Drive… (`F11`) `[cmd/pane/connect]` *(opens the Windows dialog; remote path is editable; when focused pane is File System browsing an UNC path (`\\\\...`), prefill remote name with the current path; otherwise open with no prefill; on success, if a new logical drive appears, navigate the focused pane to the new drive root)*
- Disconnect… (`F12`) `[cmd/pane/disconnect]` *(opens the Windows dialog; before opening, cancel any pending enumeration and clear DirectoryInfoCache (stops folder watchers) for the focused pane; if focused pane is a mapped network drive, preselect it; if the focused pane drive is removed, navigate to the default file system root)*
- Shared Directories… (`Ctrl+Shift+F9`) `[cmd/pane/shares]`
- ---
- Command Shell (`⊘`) `[cmd/pane/openCommandShell]` *(opens system shell at focused pane path; mounted: opens at mount backing folder)*
- Quick Search (`Shift+Space`) `[cmd/pane/quickSearch]`
- Bring Current Directory to Command Line (`Ctrl+Space`) `[cmd/pane/bringCurrentDirToCommandLine]`
- Bring Filename to Command Line (`Ctrl+Enter`) `[cmd/pane/bringFilenameToCommandLine]`
- Pane Menu (`F10`) `[cmd/pane/menu]`
- Reread Associations (`⊘`) `[cmd/app/rereadAssociations]`
- ---
- User Menu >
  - *(User menu items, dynamic)*
    - `<User Menu Item>` (`F9` opens the user menu root) `[cmd/pane/userMenu]`
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

- Plugin Manager… (`⊘`) `[cmd/app/plugins/manage]`
- ---
- *(Installed plugins, dynamic)*
  - `<File System Plugin Name>` (`⊘`; parameterized: pluginId) `[cmd/pane/selectFileSystemPlugin]`
  - Enable/Disable `<Plugin>` [td] (`⊘`; parameterized: pluginId) `[cmd/app/plugins/toggleEnabled]`
  - Configure `<Plugin>`… [td] (`⊘`; parameterized: pluginId) `[cmd/app/plugins/configure]`

#### View

- Theme >
  - *(System high contrast indicator, dynamic system state)*
    - High Contrast (System) (`⊘`; read-only indicator) `[cmd/app/theme/systemHighContrastIndicator]`
  - ---
  - System (`⊘`; parameterized: `builtin/system`) `[cmd/app/theme/select]`
  - Light (`⊘`; parameterized: `builtin/light`) `[cmd/app/theme/select]`
  - Dark (`⊘`; parameterized: `builtin/dark`) `[cmd/app/theme/select]`
  - Rainbow (`⊘`; parameterized: `builtin/rainbow`) `[cmd/app/theme/select]`
  - High Contrast (App) (`⊘`; parameterized: `builtin/highContrast`) `[cmd/app/theme/select]`
  - ---
  - *(Theme files and user themes, dynamic)*
    - `<Theme Name>` (`⊘`; parameterized: themeId) `[cmd/app/theme/select]`
  - ---
  - Previous Theme (`Shift+F11`) `[cmd/app/theme/selectPrev]`
  - Next Theme (`Shift+F12`) `[cmd/app/theme/selectNext]`
- ---
- Toggle Fullscreen (`Ctrl+Shift+F11`) `[cmd/app/fullScreen]`
- Window Menu (`Alt+Space`) `[cmd/pane/windowMenu]`
- Switch Pane Focus (`Tab`) `[cmd/pane/switchPaneFocus]`
- Show Function Bar (`⊘`; checkable) `[cmd/app/toggleFunctionBar]`
- Show Menu (`⊘`; checkable) `[cmd/app/toggleMenuBar]`
- ---
- Preferences… (`⊘`) `[cmd/app/preferences]`

#### Right (pane menu: targets Right pane)

Right menu is identical to Left menu, except:
- Change Drive (`Alt+F2`) *(opens file-system drive menu; when pane is in a non-`file` plugin, the NavigationView menu also exposes a bottom “Change Drive” submenu)* `[cmd/app/openRightDriveMenu]`
- All `cmd/pane/*` entries target the Right pane.

#### Help (right-justified)

- Display Shortcuts… (`F1`) `[cmd/app/showShortcuts]`
- External Help (`⊘`) `[cmd/app/externalHelp]`
- ---
- About… (`Alt+?`) `[cmd/app/about]`

##### Shortcuts window (`cmd/app/showShortcuts`)

- The window includes a **Search** edit at the top.
- Search is **case-insensitive** and filters rows by **command name**, **description**, or **shortcut text**.
- Matching substrings are highlighted in the list.

##### External Help (`cmd/app/externalHelp`)

- Invoking the command MUST open the external RedSalamander documentation URL in the default browser:
  `https://github.com/RedSalamanders/RedSalamander/tree/main/Docs#readme`
- The command has no default keyboard shortcut.

### Command details (Implemented)

#### Go Root Directory (`cmd/pane/goRootDirectory`)

- Invoking the command MUST navigate the target pane to the “effective root” for the current file system:
  - **Win32 file system (`file`)**: the drive root (e.g. `C:\`) or UNC share root (e.g. `\\server\share\`).
  - **Plugins**: the plugin root (`/`), except when the current plugin path is under a Connection Manager root (`/@conn:<name>/...`), in which case the effective root is `/@conn:<name>/`.
  - **Mounted plugins** (`<shortId>:<instanceContext>|<pluginPath>`): the effective root MUST preserve the mount context and set the plugin path to `/` (or the Connection Manager root when applicable).

#### Toggle Hidden Files (`cmd/pane/viewOptions/toggleHiddenFiles`)

- Invoking the command MUST toggle `folders.showHiddenFiles` (see `Specs/Core/Core_SettingsStore.md`).
- Default is `true` (shown).
- When `true`, hidden items (Windows `FILE_ATTRIBUTE_HIDDEN`) MUST be visible and MUST display a dimmed icon.

#### Toggle System Files (`cmd/pane/viewOptions/toggleSystemFiles`)

- Invoking the command MUST toggle `folders.showSystemFiles` (see `Specs/Core/Core_SettingsStore.md`).
- Default is `true` (shown).

#### Toggle Status Bar (`cmd/pane/viewOptions/toggleStatusBar`)

- Invoking the generic command MUST toggle the active pane's status bar.
- Explicit `Left` and `Right` menu entries MUST target their named pane.
- Menu checks MUST refresh after the command changes visibility.

#### Toggle Fullscreen (`cmd/app/fullScreen`)

- Invoking the command MUST toggle borderless fullscreen for the main window (hide title bar, cover the current monitor including taskbar).
- While fullscreen is active, pressing `Esc` MUST exit fullscreen.
- Invoking the command again MUST exit fullscreen.

#### Theme Selection (`cmd/app/theme/select`, `cmd/app/theme/selectNext`, `cmd/app/theme/selectPrev`)

- `cmd/app/theme/select` MUST apply the requested built-in, file, or user theme and update the checked item in the Theme menu.
- `cmd/app/theme/selectNext` and `cmd/app/theme/selectPrev` MUST cycle through selectable themes in the same order shown by the Theme menu.
- The cycle order MUST be: System, Light, Dark, Rainbow, High Contrast (App), theme files sorted by name/id, then user settings themes sorted by name/id.
- Cycling MUST wrap at either end and MUST skip the High Contrast (System) read-only indicator.

#### External Action Macros

Settings-driven external viewer/editor/user-menu launch strings MUST support these macro tokens:

| Macro | Expands to |
| --- | --- |
| `{Path}` | Current item parent path, or the explicitly supplied current directory. |
| `{FullPath}` | Current item full path, including filename. |
| `{PathAndFilename}` | Alias for `{FullPath}`. |
| `{Filename}` | Current item filename only. |
| `{SelectedPathsFile}` | Temporary file containing selected item paths, when supplied by the command. |
| `{OppositePanePath}` | Opposite pane current path. |
| `{ComputerName}` | Current computer name used for settings filters. |

- Literal braces MUST be escaped as `{{` and `}}`.
- Unknown macros, unclosed macros, and required macros with missing context MUST fail validation before any process is launched.
- The launch-plan builder MUST be deterministic and testable without starting a process.

#### Viewer and Editor Commands

- `cmd/pane/view` and `cmd/pane/edit` MUST target the focused item in the active pane and resolve the primary action from `fileActions.viewers.associations` or `fileActions.editors.associations`.
- `cmd/pane/alternateView` and `cmd/pane/alternateEdit` MUST resolve the alternate action from `fileActions`. If no applicable alternate action exists, the command MUST show a localized pane alert instead of opening the primary action or doing nothing.
- `cmd/pane/viewWith` and `cmd/pane/editWith` MUST populate their dynamic menus from applicable configured actions for the focused item. Parameterized forms (`cmd/pane/viewWith/<viewerId>` and `cmd/pane/editWith/<editorId>`) MUST launch the configured action whose ID matches case-insensitively.
- External viewer/editor actions MUST use the macro contract above, including creating and later cleaning up `{SelectedPathsFile}` only when the launch string requests it.
- Disabled, filtered, missing, invalid-id, macro-validation, and process-launch failures MUST report precise localized feedback with enough context for the user to fix the action.
- `cmd/pane/editNew` MUST create a new file in the active pane's current directory after validating the requested filename. Its Editor combo MUST be filtered from `editNewActionId` associations by extension/pattern/default row, current computer, action applicability, and executable availability; creating the file is allowed even when no applicable editor is available.

#### View Width (`cmd/app/viewWidth`)

- Invoking the command MUST enter “view width adjust” mode for the main pane splitter.
- While active:
  - `Left` / `Right` arrows MUST nudge the splitter.
  - `Enter` MUST commit the new width.
  - `Esc` MUST cancel and restore the splitter ratio captured when the mode started.
- Invoking the command again while active MUST commit (same as `Enter`).

#### Windows Shell Actions

- `cmd/pane/contextMenuCurrentDirectory` MUST open the Windows shell context menu for the active pane's current local folder.
- `cmd/pane/openSecurity` MUST open the Windows Security property page for the active pane's focused local item.
- These commands MUST only run against the built-in Windows file-system provider. If the active pane is a remote/plugin provider, or if the required local folder/item cannot be resolved, the command MUST show localized pane feedback and MUST NOT fall through to the generic "not implemented" message.
- Shell COM and menu resources MUST be owned by RAII wrappers. The context-menu command MUST treat a canceled popup as normal control flow.
- Deterministic command selftests MUST cover command routing with a shell-action probe and archive `shell.context_menu_current_directory_us` / `shell.open_security_us` timing metrics.

#### Change Attributes (`cmd/pane/changeAttributes`)

- The command MUST target the selected items in the active pane, or the focused item when no selection is present.
- It MUST show a tri-state dialog for read-only, hidden, system, and archive attributes. Checked sets an attribute, clear removes it, and mixed leaves it unchanged. Repeated keyboard or pointer activation MUST cycle each editable attribute through set, clear, and leave-unchanged so the user can return to "no changes" without closing the dialog.
- It MUST include a Change Date and Time section with opt-in rows for Modified, Created, and Accessed timestamps. Editing a row's date or time MUST automatically enable that row; unchecked rows MUST leave the corresponding timestamp unchanged. Invalid enabled date/time input MUST keep the dialog open.
- It MUST include an Include subdirectories option. The option MUST be visible but disabled when the selection contains no folders, and enabled when at least one selected/focused target is a folder.
- It MUST include an option to remove alternate data streams from the targeted selection when the active file-system provider exposes removable named streams.
- It MUST apply attribute changes, timestamp changes, and stream removal per item, refresh the pane when anything changed, and show a localized operation report with processed items, changed attribute count, changed date/time count, removed stream count, failure count, and first failure HRESULT when failures occur.
- When Include subdirectories is enabled, the command MUST run as a File Operations informational task. The task MUST show enumeration/apply status, including the current path and item counts, and MUST finish with the same localized summary shown by the pane feedback overlay.
- Recursive Change Attributes MUST include each selected folder itself and its descendants. It MUST enumerate descendants through the active provider's directory API and MUST NOT follow child directories marked as reparse points, so mount points and other link-like folders are changed only as selected items and are not traversed.
- Unsupported providers, empty selections, canceled dialogs, and no-op dialogs MUST report or return without falling through to the generic "not implemented" message.
- Deterministic command selftests MUST cover selected-item scope, attribute set/clear/leave-unchanged cycling, date/time rows, Include subdirectories enablement, recursive date/time application with File Operations progress, alternate data stream removal, report contents, and archived `fileattrs.*` timing metrics.

#### Calculate Occupied Space (`cmd/pane/viewSpace`)

- Invoking the command MUST open Space Viewer for the target pane.
- If exactly one selected item is a directory, that directory is the Space Viewer target.
- Otherwise, the target pane’s current folder is the Space Viewer target.

#### Find Files and Directories (`cmd/pane/find`)

- Invoking the command MUST open or activate the modeless host-owned `Find Files and Directories` window.
- The command MUST target the focused pane when focus is inside a pane; otherwise it MUST target the active pane.
- The initial search scope MUST come from the target pane's current plugin, instance context, and current path.
- If the Find window is already open, invoking the command again MUST reuse that window and refresh its context instead of opening a duplicate.
- Search execution MUST remain off the UI thread, and the dialog MUST support `Find`, `Append`, `Intersect`, `Subtract`, and `Cancel`.
- See `Specs/Core/Core_Search.md` for backend selection, result-set behavior, and persistence rules.

#### Rename (`cmd/pane/rename`)

- Invoking the command MUST show a modal dialog:
  ```text
  Title: Rename Item

  New name:
  [ <name> ]   (editable input)

                       [ OK ] [ Cancel ]
  ```
- The dialog MUST be theme-aware (title bar + background + input + buttons) and follow `Specs/UI/UI_VisualStyle.md` for framed inputs and owner-draw buttons (skip custom drawing in high contrast).
- Layout:
  - The dialog SHOULD default to a wide size (same width as the Select/Unselect mask dialog) to accommodate long names.
  - The dialog MUST be centered on the main window.
- Default value MUST be the focused item’s current display name.
- Initial text selection:
  - For files, the dialog SHOULD select the name without its extension (up to the last `.`).
  - For folders, the dialog SHOULD select the full name.
- `Enter` commits; `Escape` cancels.
- While the name field is focused, standard edit navigation and clipboard keys MUST stay local to the field: arrows, Home/End, Backspace/Delete,
  `Ctrl+Left/Right`, `Ctrl+Backspace/Delete`, `Ctrl+A/C/X/V/Z/Y`, `Ctrl+Insert`, `Shift+Insert`, and `Shift+Delete` MUST edit, select,
  copy, cut, paste, undo, or redo the proposed name instead of dispatching pane shortcuts.
- The new name MUST be trimmed; empty input MUST be rejected (warning beep) and the dialog MUST remain open.

#### Change Case (`cmd/pane/changeCase`)

- Invoking the command MUST show a modal dialog:
  ```text
  Title: Change Case

  [ ] Include subdirectories (apply to selected folders recursively)

  Change Case to                    Change
    o Lower case                      o Whole filename
    o Upper case                      o Only name
    o Partially mixed case (...)      o Only extension
    o Mixed case (...)

    [ OK ] [ Cancel ]
  ```
- Layout:
  - When there is enough horizontal space, the dialog SHOULD use a two-column layout with **balanced** column widths (avoid an oversized left column).
  - When space is constrained, the dialog MUST fall back to a single-column stacked layout.
  - When DPI changes while the dialog is open, it MUST recompute layout and typography so the two-column split remains balanced and controls do not overlap.
  - The dialog SHOULD size its height to content so the button row does not sit far below the last option card.
- Scope: apply to `Selected items`; if no items are selected, apply to the `Current item`.
- “Include subdirectories”:
  - MUST be available for any filesystem plugin.
  - When enabled, traversal MUST be **non-recursive** (iterative) to avoid stack overflow on deep directory hierarchies.
  - Traversal MUST use the active plugin’s directory enumeration semantics (it MUST work with non-Windows / plugin-specific paths).
 - Execution MUST be asynchronous (no long UI-thread stalls) and MAY surface progress as an informational task in the File Operations popup for long runs.
 - Case-only renames on case-insensitive file systems SHOULD be supported (use a temp rename where required).

#### Select / Unselect (Mask Dialog) (`cmd/pane/selection/selectDialog`, `cmd/pane/selection/unselectDialog`)

- Invoking `selectDialog` MUST show a modal dialog:
  ```text
  Title: Select

  Select
  [ <mask> ]   (editable combo box)

  Mask syntax ▸ (toggle)
  (help text)

                               [ OK ] [ Cancel ]
  ```
- Invoking `unselectDialog` MUST show a modal dialog with the same layout, but title/label “Unselect”.
- History:
  - The mask combo MUST persist history entries (most-recent-first, max 10).
  - `selectDialog` history key: `settings.selectionMasks.selectHistory`
  - `unselectDialog` history key: `settings.selectionMasks.unselectHistory`
- Mask matching:
  - `*` matches any number of characters; `?` matches any single character.
  - Masks are separated by `;` (use `;;` for a literal `;`).
  - Excluded masks come after `|` (examples: `|*.tmp`, `*.txt;*.doc|~*`).
  - Matching is case-insensitive.
- Behavior:
  - `selectDialog` MUST add matching items to the current selection (it MUST NOT unselect non-matching items).
  - `unselectDialog` MUST remove matching items from the current selection (it MUST NOT select anything).

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
- The visible Function Bar text surface (key glyphs, command labels, and the optional modifier indicator) MUST render and measure through the shared `DxUi.Typography` / DirectWrite path; do not reintroduce plain GDI `DrawTextW` text paint or `HFONT` text measurement on this surface.
- Function Bar Direct2D drawing is DPI-aware. The control receives Win32 pixel rectangles, but drawing coordinates, text layout rectangles, rounded glyph radii, and separator stroke widths MUST be converted to DIPs before rendering so enabled chrome is painted correctly at every DPI scale.
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
| F2   | Rename          | Sort None                | Open Right Drive Menu       | ⊘                    | ⊘                                | ⊘                                         |
| F3   | View            | Sort by Name             | Alternate View              | Open Current Folder  | View Width                       | ⊘                                         |
| F4   | Edit            | Sort by Extension        | Exit                        | Edit New             | Alternate Edit                   | ⊘                                         |
| F5   | Copy            | Sort by Time             | Pack                        | ⊘                    | Save Selection                   | ⊘                                         |
| F6   | Move            | Sort by Size             | Unpack                      | ⊘                    | Restore Selection                | ⊘                                         |
| F7   | Make Directory  | Change Case              | Find                        | Change Directory     | ⊘                                | ⊘                                         |
| F8   | Delete          | Change Attributes        | ⊘                           | Permanent Delete     | ⊘                                | ⊘                                         |
| F9   | User Menu       | Refresh                  | Unpack                      | Hot Paths            | Shares                           | ⊘                                         |
| F10  | Menu            | Compare                  | Space View                  | Context Menu         | ⊘                                | Context Menu (Current Directory)          |
| F11  | Connect         | Zoom Panel               | List of Opened Files        | Previous Theme       | Full Screen                      | ⊘                                         |
| F12  | Disconnect      | Filter                   | Show Folders History        | Next Theme           | ⊘                                | ⊘                                         |

### FolderView Configurable Shortcuts (Non-Function Bar)

These chords apply while focus is inside the `FolderWindow` (either `FolderView` or `NavigationView`).

Resolution order (normative):
- When a `FolderView` has focus: if the chord is bound in `settings.shortcuts.folderView`, the host MUST execute the bound command and MUST consume the key message.
- When focus is inside the `FolderWindow` but not in a `FolderView`: chords in `settings.shortcuts.folderView` are evaluated only when at least one modifier (Ctrl/Alt/Shift) is down and the key is not `Tab`; if bound, the host MUST execute the bound command and MUST consume the key message.
- Otherwise, the message continues through the normal routing pipeline (accelerators, then `FolderView`’s built-in key handling).

This means any key listed as a valid `vk` in `Specs/Core/Core_SettingsStore.md` (including Arrow keys / PageUp / PageDown / Home / End / `0`-`9`) can be made configurable by adding a binding entry; unbound chords keep their built-in behavior.

`⊘` means “no shortcut assigned”.

#### Default `shortcuts.folderView` bindings (implemented)

| Key       | None                               | Ctrl                             | Alt                      | Shift                              | Ctrl+Shift                        | Ctrl+Alt          | Alt+Shift             |
|-----------|------------------------------------|----------------------------------|--------------------------|------------------------------------|-----------------------------------|-------------------|-----------------------|
| Backspace | Up One Directory                   | ⊘                                | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| Tab       | Switch Pane Focus                  | ⊘                                | ⊘                        | Switch Pane Focus                  | ⊘                                 | ⊘                 | ⊘                     |
| U         | ⊘                                  | Swap Panes                        | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| A         | ⊘                                  | Select All                       | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| =         | ⊘                                  | Select...                        | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| -         | ⊘                                  | Unselect...                      | ⊘                        | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
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
| 5         | ⊘                                  | ⊘                                | Display as Thumbnails    | ⊘                                  | ⊘                                 | ⊘                 | ⊘                     |
| 0..9      | ⊘                                  | Go to Hot Path (`Ctrl+<digit>`)  | ⊘                        | ⊘                                  | Set Hot Path (`Ctrl+Shift+<digit>`) | ⊘               | ⊘                     |
| A..Z      | ⊘                                  | ⊘                                | ⊘                        | Go to Drive Root (`<drive>:\\`)    | ⊘                                 | ⊘                 | ⊘                     |
| Enter     | Execute / Open                     | Bring Filename to Command Line   | Open Properties          | ⊘                                  | Bring Filename to Command Line    | ⊘                 | ⊘                     |
| Space     | Select + Calc Dir Size + Next      | Bring Current Dir to Command Line | Window Menu              | Quick Search                       | Bring Current Dir to Command Line | ⊘                 | ⊘                     |
| Insert    | Select + Next                      | Clipboard Copy                   | Copy Path + Name as Text | Clipboard Paste                    | Copy UNC Path + Name as Text      | Copy Path as Text | Copy Name as Text     |
| Delete    | Move to Recycle Bin                | ⊘                                | ⊘                        | Permanent Delete                    | Permanent Delete                    | ⊘                 | ⊘                     |

Notes:
- Unmodified digit keys (`0`-`9`) and unmodified letter keys are unbound by default so they can be used for incremental search typing; Hot Paths use `Ctrl+<digit>` / `Ctrl+Shift+<digit>` and do not interfere with typing.

### Shortcut Customization UI (Preferences)

- The main menu includes `View → Preferences...` (near the bottom, separated).
- The settings dialog includes a `Shortcuts` tab with sections:
  - Function Bar shortcuts (F1..F12)
  - Folder view shortcuts (all supported keys)
- Settings are loaded at application startup; shortcut bindings are restored and applied before the first main window interaction.
- In the Shortcuts window, clicking the `Key` column MUST sort by semantic key identity rather than the rendered chord text. Ascending order is: function keys (`F1`..`F24`, numeric order), digit keys (`0`..`9`), letter keys (`A`..`Z`), then other keys by localized key display text. Rows with the same base key MUST stay together and compare by displayed modifier phrase alphabetically (unmodified first), then by command text as a stable tie-breaker. Persisted `Key` sort state MUST use the same semantic order when the window is reopened.
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
- `Alt` (alone) temporarily shows the menu bar when hidden and starts menu interaction (see `Specs/Core/Core_SettingsStore.md`).

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
- `Shift + Delete`: ask for confirmation, then permanently delete `Selected items` (or `Current item` when selection is empty) without using the Recycle Bin.
- `F2`: rename `Current item`
- `Tab` / `Shift+Tab`: move focus   between Pane `FolderView`s (no longer enters NavigationView)
- `Alt+D` / `Ctrl+L`: focus NavigationView address edit
- `Alt+Down`: go to next selected name
- `Alt+Up`: go to previous selected name
- `Tab`: switch focus to the other pane’s `FolderView`.
  - `Shift+Tab`: same as `Tab` (two-pane toggle) unless later extended.

**NavigationView** (see `RedSalamander/NavigationView.Interaction.cpp`)
- `Tab` / `Shift+Tab`: cycle focus between visible regions (Menu → Path → History → Disk Info), then hand off to FolderView
- `Alt+D` / `Ctrl+L`: enter edit mode / focus address edit
- `Enter` / `Space`: activate focused region
- Mouse or menu activation in an unfocused pane's `NavigationView` MUST first make that pane active and return keyboard focus to its `FolderView` before or as part of the navigation action. This covers the drive/menu button, history and disk-info dropdowns, breadcrumb segment clicks, sibling menus, and menu-selected path changes.

### RedSalamanderMonitor

**ColorTextView** (see `Specs/Core/Core_RedSalamanderMonitor.md`)
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
