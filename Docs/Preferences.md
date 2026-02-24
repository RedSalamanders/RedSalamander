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

## Pages

The left tree contains:

- **General**: common app settings.
- **Panes**: left/right pane behavior and visibility options (including hidden/system file visibility and history size).
- **Viewers**: file extension → viewer mapping (used by `F3`).
- **Editors**: placeholder (not implemented yet).
- **Keyboard**: shortcut bindings (Function Bar + FolderView).
- **Mouse**: placeholder (not implemented yet).
- **Themes**: theme selection and theme file management.
- **Plugins**: enable/disable plugins, configure plugins, and run plugin tests.
- **Compare Directories**: default options used by the Compare Directories window.
- **Hot Paths**: bookmark folder paths for quick access (`Ctrl+1`..`Ctrl+0`).
- **Advanced**: expert settings and diagnostics-related options.

Tip: Plugins also appear as child nodes under **Plugins** when a plugin exposes configurable fields.
