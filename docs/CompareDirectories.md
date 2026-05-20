# Compare Directories

**Compare Directories** opens a dedicated window that compares two folder trees (Left vs Right) and highlights differences.

![Compare Directories window](res/compare-directories.png)

## Open

- Menu: **Commands → Compare Directories…**
- Shortcut (default): `Ctrl+F10`

## Workflow

1. In the Compare Directories window, navigate the **Left** pane and **Right** pane to the folders you want to compare.
2. Click **Options…** to choose the comparison criteria and display options, then click **OK** to start a scan.
3. Use **Rescan** to re-establish the compare roots from the panes’ current folders and start a fresh scan.
   - While a scan is running, **Rescan** becomes **Cancel**.

Progress is shown in the banner, and long-running scans may also appear as an informational card in the [File Operations popup](FileOperations.md).

## Matching rules

Items are matched by **name** under the same **relative folder** (directory-oriented comparison).

## Options (high level)

- **Compare files with same name by**: Size, Date/Time, Attributes, Content
- **Subdirectories**: whether to compare subdirectories (and their attributes)
- **Ignore patterns**: file and directory ignore pattern lists
- **Display**: **Show Identical Items** (when off, only differences are shown)

## Window commands

The Compare Directories window has its own command surface:

- **Options...** opens the compare options panel and saves accepted changes as the new default compare settings.
- **Rescan** restarts the comparison using the current left/right roots; while scanning it changes to **Cancel**.
- **Show Identical Items** toggles identical rows when they were retained by the scan.
- **Restore Differences Selection** returns selection to the detected difference set.
- **Invert Differences Selection** flips the current difference selection.
- **Brief**, **Detailed**, and **Extra Detailed** change the result-pane display mode.
- **Close** exits the compare window.

If the main settings file changes while the **Options…** panel is open:

- A clean options panel reloads from disk automatically.
- A dirty options panel prompts **Reload from disk** or **Keep editing**.
- If you keep editing, pressing **OK** later asks whether to **Overwrite disk**, **Reload from disk**, or **Cancel**.
- **Options → OK** saves the updated compare defaults to the main settings file immediately before starting the new scan.

## Sync operations (copy/move)

When Compare mode is active, copying/moving a **directory** to the other pane acts like a sync:

- Only items that differ (or exist on only one side) are copied/moved.
- Identical descendants are skipped.
