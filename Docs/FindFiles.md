# Find Files and Directories

Use **Find Files and Directories** to search from the current pane without leaving your current location.

Open it with:
- **Commands → Find Files and Directories…**
- `Alt+F7`

## Scope

The dialog starts from the focused pane:
- the current filesystem plugin,
- the current mounted context, if any,
- the current path as the search root.

If the window is already open, RedSalamander reuses it and refreshes it to the current pane context.

## Search options

You can search by:
- **Name** using wildcard, literal, or regex matching
- **Content** using text literal or text regex matching

Common options:
- recursive search,
- include files,
- include directories,
- follow symlinks,
- case-sensitive name matching,
- case-sensitive content matching,
- prefer indexed search,
- request text snippets.

If both name and content are enabled, a result must match both.

## Result actions

The result list updates while the search is running.

Actions:
- **Find Now** replaces the current result set
- The **Find Now** split-button menu offers **Append**, **Intersect**, and **Subtract** once the result list contains items
- **Append** adds new matches
- **Intersect** keeps only matches that are already in the result set
- **Subtract** removes newly found matches from the current result set
- **Cancel** stops the current search

When recursive search finds matches below the root, the Path column shows the containing subfolder relative to the search root. Result rows use the shared grid density, so compact mode reduces the row spacing for one-line results.

Double-click behavior:
- file: open it with the normal host flow
- directory: navigate to that directory
- parent action: open the parent and focus the matched item when possible

## Backends

RedSalamander chooses the best available backend for the current pane:
- local `file:` locations prefer the Windows search service, then the local index, then a direct scan
- other plugins may provide their own native search
- when native search is unavailable, the host falls back to a direct scan when possible

The status line shows the active backend and any degradation warning, such as content search not being available for the current plugin.

## Saved state

The dialog remembers:
- recent roots,
- recent name and content patterns,
- the last-used options,
- window placement.

To reset this history manually, remove the `search` section from the settings file and optionally remove `windows.FindFilesWindow`. See [Settings File & Advanced Configuration](SettingsFile.md).
