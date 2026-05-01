# Screenshots (`Docs/res/`)

The user documentation in `Docs/` references screenshots in this folder.

Screenshots should be captured from the running application whenever possible. The current user guide intentionally uses multiple themes (Light, Dark, Rainbow, and High Contrast App) to illustrate the theme system.

Current placeholders (1×1 PNG):

- `placeholder.png` (template placeholder image)

## Capture guidelines

- Prefer Windows 11.
- Use the **Light** theme for neutral workflow screenshots unless the screenshot is explicitly about theming or the guide needs a theme mix.
- Do not use open dropdown/flyout menus in theme-comparison screenshots. Capture theme examples with the menus closed, and capture menu/dropdown behavior as its own focused screenshot.
- Viewer screenshots must show the viewer plugin opened on a representative file or folder. Do not use Preferences -> Plugins or Preferences -> Viewers screenshots as substitutes for viewer-in-action examples.
- Crop tightly to the relevant UI.
- Remove or blur sensitive paths, hostnames, usernames, bucket names, and emails.

## Files (what each screenshot should show)

- `main-window.png`: RedSalamander dual-pane main window (menu bar + function bar visible).
- `navigation-view.png`: Navigation bar sections (menu/drive, breadcrumb, history, disk info).
- `folder-view.png`: FolderView in Brief vs Detailed (or a view that clearly shows items + selection).
- `drive-menu.png`: NavigationView drive/menu dropdown (known folders + WSL + drives).
- `history-dropdown.png`: NavigationView history dropdown open.
- `file-operations-popup.png`: File operations popup with at least one active task.
- `file-operations-issues.png`: File operations “Failed Items / Issues” pane with an example issue.
- `file-properties.png`: Properties dialog showing general metadata, timestamps, and attributes.
- `compare-directories.png`: Compare Directories window, options, or scan result.
- `preferences.png`: Preferences dialog overview.
- `preferences-plugins.png`: Preferences → Plugins (plugin list + a plugin details subpage).
- `preferences-viewers.png`: Preferences → Viewers (extension → viewer mapping).
- `preferences-keyboard.png`: Preferences → Keyboard (shortcut bindings).
- `preferences-themes.png`: Preferences → Themes (theme selection + edit controls).
- `connections-manager.png`: Connection Manager dialog (profile list + editor pane).
- `plugins.png`: Plugin Manager entry point (menu item or Preferences → Plugins root page).
- `viewer-text.png`: Text viewer showing Text/Hex toggle and status bar.
- `viewer-space.png`: ViewerSpace treemap during or after scanning.
- `viewer-imgraw.png`: Image viewer (fit-to-window, zoom controls, or Exif overlay).
- `viewer-pe.png`: PE viewer showing parsed info and export options.
- `viewer-sqlite.png`: SQLite viewer showing a table/query preview on a sample database.
- `viewer-vlc.png`: VLC viewer playing a media file (or the “VLC required” message).
- `viewer-web.png`: Web viewer rendering a sample HTML/PDF-style document.
- `viewer-json.png`: JSON viewer rendering structured JSON.
- `viewer-markdown.png`: Markdown viewer rendering a Markdown document.
- `monitor.png`: RedSalamanderMonitor main window with live/loaded logs and filter menu visible.
- `theme-light.png`, `theme-dark.png`, `theme-rainbow.png`, `theme-high-contrast-app.png`: main application captured under each built-in theme.
