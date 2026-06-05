# Screenshots (`docs/res/`)

The user documentation in `docs/` references screenshots in this folder.

Screenshots should be captured from the running application whenever possible. The current user guide intentionally uses multiple themes (Light, Dark, Rainbow, and High Contrast App) to illustrate the theme system.

Current placeholders (1×1 PNG):

- `placeholder.png` (template placeholder image)

## Latest capture pass

The 2026-05-21 documentation review refreshed the public theme screenshots from
the current Debug build and added sanitized running-app captures for:

- `folder-view-brief-detailed.png`: Brief and Detailed pane layouts.
- `folder-view-extra-thumbnails.png`: Extra Detailed and Thumbnails pane layouts.
- `preview-pane-text.png`: Preview Pane hosted in the opposite pane with text content.
- `theme-light.png`, `theme-dark.png`, `theme-rainbow.png`, `theme-high-contrast-app.png`: same generated demo folders under different built-in app themes.
- `theme-controls-*.png` and `theme-button-states-after-fix.png`: regenerated from `DxUiTests`.

The running-app captures used generated files under `.build/docs-screenshot-demo`
and a temporary sanitized settings file that was restored after capture. They
must not show user-profile paths, real folder history, remotes, emails, tokens,
bucket names, or connection credentials.

## Missing screenshot backlog

The docs are now organized so missing images are tracked explicitly instead of being implied by prose. Capture these next when a running build is available:

- `selection-mask.png`: Select / Unselect by mask dialog from the **Edit** menu.
- `pane-filter.png`: pane filter dialog and the resulting Filter Bar summary.
- `quick-search.png`: Quick Search active in a folder pane with highlighted matches.
- `folder-view-sort-menu.png`: pane sort/status popup, including sort direction and thumbnail-size slider.
- `folder-view-hidden-system.png`: hidden/system file visibility, using generated non-sensitive fixtures.
- `command-line.png`: pane command-line input after **Bring Current Directory/Filename to Command Line**.
- `shortcuts-window.png`: **Help -> Display Shortcuts**.
- `preferences-general.png`: Preferences -> General with display, language, compact mode, animation, and backdrop options visible.
- `preferences-panes.png`: Preferences -> Panes with left/right display, sort, thumbnail size, hidden/system visibility, and history settings.
- `preferences-editors.png`: Preferences -> Editors actions/associations.
- `preferences-user-menu.png`: Preferences -> User Menu entry editor.
- `preferences-mouse.png`: Preferences -> Mouse placeholder page.
- `preferences-file-operations.png`: Preferences -> File Operations with pre-calculation, speed limit, and bridge buffer settings.
- `preferences-compare-directories.png`: Preferences -> Compare Directories default options.
- `preferences-hot-paths.png`: Preferences -> Hot Paths slot editor.
- `preferences-advanced.png`: Preferences -> Advanced diagnostics, monitor, connections, and cache settings.
- `preferences-plugin-child.png`: one plugin-specific child page under Preferences -> Plugins.
- `find-files.png`: Find Files and Directories window with options and live or completed results.
- `compare-options.png`: Compare Directories options panel.
- `view-width.png`: **Files -> View Width...** dialog.
- `change-attributes.png`: **Files -> Change Attributes...** dialog.
- `change-case.png`: **Files -> Change Case...** dialog.
- `make-file-list.png`: Make File List options dialog.
- `opened-files.png`: List of Opened Files dialog.
- `shared-directories.png`: Shared Directories dialog.
- `shell-new.png`: **Files -> New** with generated/safe ShellNew templates.
- `user-menu.png`: configured **Commands -> User Menu** with generated safe entries.
- `network-drive.png`: Connect/Disconnect Network Drive flow, cropped to avoid real network names.
- `viewer-web-menu.png`: Web/JSON/Markdown viewer with a menu open, preferably **Tools**.
- `viewer-text-diff.png`: Text viewer in diff side-by-side or inline mode.

## Capture guidelines

- Prefer Windows 11.
- Use the **Light** theme for neutral workflow screenshots unless the screenshot is explicitly about theming or the guide needs a theme mix.
- Do not use open dropdown/flyout menus in theme-comparison screenshots. Capture theme examples with the menus closed, and capture menu/dropdown behavior as its own focused screenshot.
- Viewer screenshots must show the viewer plugin opened on a representative file or folder. Do not use Preferences -> Plugins or Preferences -> Viewers screenshots as substitutes for viewer-in-action examples.
- Crop tightly to the relevant UI.
- Remove or blur sensitive paths, hostnames, usernames, bucket names, and emails.
- Prefer generated files/folders under `.build/` or a dedicated disposable demo root. Do not capture `%UserProfile%`, OneDrive, live remote profiles, local machine names, secrets, or real command history.
- For native app captures, prefer handle-based capture (`PrintWindow` or a tested in-app capture helper) over desktop-region capture so background windows cannot leak into the PNG.

## Files (what each screenshot should show)

- `main-window.png`: RedSalamander dual-pane main window (menu bar + function bar visible).
- `navigation-view.png`: Navigation bar sections (menu/drive, breadcrumb, history, disk info).
- `folder-view.png`: FolderView in Brief vs Detailed (or a view that clearly shows items + selection).
- `folder-view-brief-detailed.png`: running app with generated files, left pane in Brief mode and right pane in Detailed mode.
- `folder-view-extra-thumbnails.png`: running app with generated files, left pane in Extra Detailed mode and right pane in Thumbnails mode.
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
- `theme-controls-*.png`: generated DxUi control galleries, one file per built-in or shipped JSON5 theme. Regenerate with `.build\x64\Debug\DxUiTests.exe --suite=Gallery --gallery-output-directory=docs\res`.
- `theme-button-states-after-fix.png`: generated DxUi button-state contrast audit across built-in and shipped JSON5 themes. Regenerate with `.build\x64\Debug\DxUiTests.exe --suite=ButtonContrast --button-audit-output=docs\res\theme-button-states-after-fix.png`.
- `connections-manager.png`: Connection Manager dialog (profile list + editor pane).
- `plugins.png`: Plugin Manager entry point (menu item or Preferences → Plugins root page).
- `viewer-text.png`: Text viewer showing Text/Hex toggle and status bar.
- `preview-pane-text.png`: Preview Pane in the opposite pane, showing generated text content.
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
