# Preferences

Open Preferences from:

- **View → Preferences…**
- Tip: **Hot Paths** can also be opened directly with `Shift+F9`.

![Preferences](res/preferences.png)

## How Preferences works

- The dialog is modeless (you can keep using the main window behind it).
- **OK**: apply changes and close.
- **Apply**: apply changes without closing.
- **Cancel**: discard pending changes and close.

## External settings changes

- If the main settings file changes on disk and Preferences has no unsaved edits, the dialog reloads from the live app settings automatically.
- If Preferences has unsaved edits, it warns you and lets you choose **Reload from disk** or **Keep editing**.
- If you keep editing, the dialog becomes stale; the next **OK** or **Apply** asks whether to **Overwrite disk**, **Reload from disk**, or **Cancel**.
- If RedSalamander finds an unsupported settings schema version, such as v15 when v16 is required, it backs up that settings file, restores defaults, and shows a warning that the old file was not migrated automatically.
- Theme previews are cancelled when an external reload is applied, and the dialog itself switches to the newly applied theme.
- The modal plugin configuration editor follows the same reload/conflict rules.

## Pages

The left tree contains:

- **General**: common app behavior.
- **Panes**: default pane display/sort behavior, visibility options, and history size.
- **Viewers**: viewer Actions and Associations for `F3`, `Alt+F3`, and View With.
- **Editors**: editor Actions and Associations for `F4`, `Ctrl+Shift+F4`, `Shift+F4`, and Edit With.
- **User Menu**: external command entries shown by `F9` and the Commands -> User Menu popup.
- **Keyboard**: shortcut bindings for Function Bar and FolderView commands.
- **Mouse**: placeholder (not implemented yet).
- **Themes**: theme selection, user-theme editing, and theme file management.
- **Plugins**: enable/disable plugins, configure plugins, and run plugin tests.
- **File Operations**: host-wide defaults for pre-calculation, default copy/move speed limit, and cross-filesystem bridge buffering.
- **Compare Directories**: default options used by the Compare Directories window.
- **Hot Paths**: bookmark folder paths for quick access (`Ctrl+1`..`Ctrl+0`).
- **Advanced**: expert settings such as diagnostics, monitor filters, and cache limits.

Tip: Plugins also appear as child nodes under **Plugins** when a plugin exposes configurable fields.

## Common workflows

### Viewers

- Use **Associations** to choose which viewer action `F3` and `Alt+F3` use for an extension, pattern, default `*` row, or computer-specific override.
- Use **Actions** to configure internal viewer-plugin actions or external viewer programs.
- Test a file path on the page to see the resolved action and why it was chosen.
- View With lists the configured viewer actions that apply to the focused file and current computer.

### Editors

- Use **Associations** to choose which editor action `F4`, `Ctrl+Shift+F4`, and `Shift+F4` use for an extension, pattern, default `*` row, or computer-specific override.
- Use **Actions** to configure external editor programs.
- Edit New shows an Editor combo filtered by file extension, computer name, configured action availability, and executable availability.
- Apply or OK saves viewer/editor Actions and Associations to the main settings file.

### User Menu

- Configure the ordered external command entries shown by `F9` and **Commands -> User Menu**.
- Each entry can define an executable, arguments, working directory, enabled state, extension filter, and computer-name filter.
- Commands use the same macros and applicability rules as viewer and editor actions, including selected-paths file support for multi-selection workflows.

### Keyboard

- Search bindings by command name or shortcut text.
- Assign or remove shortcuts.
- Import/export shortcut bindings as JSON.
- Reset shortcuts to defaults.
- Resolve conflicts directly in the page by replacing or swapping bindings when prompted.

### Themes

- Switch between built-in, file-based, and user themes.
- Create or duplicate a user theme from an existing theme.
- Edit color overrides while still seeing inherited values.
- Preview a theme immediately with **Apply Temporarily**.
- Load/save `*.theme.json5` theme files.

### Plugins

- Search installed plugins by name, id, or type.
- Enable/disable plugins.
- Add/remove custom plugin DLL paths.
- Run **Test** for the selected plugin or **Test All** for the whole set.
- Use the **Plugins** root page for the plugin list, then open child nodes for per-plugin settings.

### File Operations

- Turn the copy/move pre-calculation scan on or off.
- Choose the pre-calculation worker budget (`1` to `8`) used before copy/move starts.
- Set the default per-task speed limit for new copy/move operations using presets or a custom value.
- Tune the cross-filesystem bridge buffer size used when the host copies between different file-system implementations.
- Use **Plugins → File System** for plugin-owned concurrency, recycle-bin batching, and search-walker settings; those are not duplicated on the File Operations page.

### Hot Paths and Advanced

- **Hot Paths** lets you edit the label, target path, and **Show in Change Drive menu** flag for each of the 10 slots.
- **Advanced** exposes file-operations diagnostics, monitor defaults, and cache-related settings.
- Some rarely used settings still remain JSON-only; see [Settings File & Advanced Configuration](SettingsFile.md).

## Screenshots (key pages)

### Viewers

![Preferences → Viewers](res/preferences-viewers.png)

### Keyboard

![Preferences → Keyboard](res/preferences-keyboard.png)

### Themes

![Preferences → Themes](res/preferences-themes.png)

### Plugins

![Preferences → Plugins](res/preferences-plugins.png)
