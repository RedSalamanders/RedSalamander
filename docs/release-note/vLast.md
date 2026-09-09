# RedSalamander vLast — Notes since v7.0.48

Last published tag: [v7.0.48](https://github.com/RedSalamanders/RedSalamander/releases/tag/v7.0.48) (July 20, 2026).

This is the working note for the **next** release; `vLast` is not a version number. It describes everything that changed after v7.0.48. The published [v7.0.48 notes](v7.0.48.md) remain the last official changelog.

v7.0.48 was a reliability and data-safety release. This drop is a product release: RedSalamander gains an **embedded terminal**, a **command palette**, a **file-operations engine that renames instead of copying** when source and destination live in the same place, **writable cloud and remote destinations**, and a consistent **current-item / selection model** in the folder panes.

## Highlights

- A built-in **Terminal** in either pane and in one floating window with tabs. The old bottom command line is gone.
- A global **Command Palette** (`Ctrl+Shift+P`) that works from panes, previews, and terminals.
- **Move within the same volume, share, server, or cloud drive is now a rename.** Moving a large folder on one disk finishes at once instead of copying every byte.
- **Copy and Move no longer open a dialog first.** `F5` / `F6` start immediately; `Shift+F5` / `Shift+F6` open the options surface.
- Optional **content verification after copy**, with the source kept when verification fails.
- **FTP, SFTP, SCP, Amazon S3, Google Drive, Microsoft Drive and SMB shares are now full destinations**, and Cancel works even when a server stops answering.
- Each pane keeps a persistent **current item** that is independent from selection. `Space`, `Insert`, and `Esc` follow that model.
- Reorganized menus, split item/background context menus, and a searchable encoding picker in Viewer Text.
- A brief **theme overlay** when cycling themes, and optional pointer-driven pane focus.

---

## Embedded Terminal

RedSalamander now hosts a real terminal: ConPTY for the process, Ghostty's `libghostty-vt` for the screen model, and a Direct2D renderer. It ships as the built-in plugin `Plugins\Terminal.dll` with its private engine under `Plugins\TerminalRuntime\`, in both the x64 and ARM64 packages.

### Opening a terminal

- **Terminal Pane** (`Alt+7`, also `Ctrl+Shift+T` from a folder, or **Left / Right → Terminal Pane**) opens or reuses the terminal in the **opposite** pane, rooted at the active folder. A live session is reused rather than starting a new process.
- **Command Shell Window** (`Ctrl+Alt+T`, `Ctrl+Shift+N`, or **Commands → Terminal**) opens or focuses one floating terminal window and adds a tab. It remembers size, position, tab order, active tab and paths, and starts **fresh shells** after a restart; it does not resurrect running processes.
- **Edit** (`F4`) on a **folder** opens the same opposite-pane terminal. Edit on a file still opens the configured editor.
- Typing `exit` closes that session or tab once output has drained. The floating window closes with its last tab.

### Which shell starts

- By default RedSalamander picks PowerShell 7, then Windows PowerShell, then Command Prompt, always resolved from their registered install locations rather than from `PATH` or the current directory. **Preferences → Plugins → Terminal** can pin PowerShell or Command Prompt instead.
- A folder under `\\wsl.localhost\<Distro>` or `\\wsl$\<Distro>` starts **that** distribution at the matching Linux path, with its configured default user and shell.
- Locations with no real Windows path (archives, cloud, MTP, other plugin providers) explain why a terminal cannot start there. They never silently open a shell somewhere else.

### Sending paths to the terminal

These commands insert text and **never press Enter**:

- **Insert Current Directory** (`Ctrl+Space` or `Ctrl+Shift+Space`) inserts the pane's folder path.
- **Insert Focused Item** (`Ctrl+Enter`) inserts just the file name when the terminal has proven that an idle PowerShell is sitting in that exact parent folder; otherwise it inserts the full path.
- **Insert Full Path** (`Ctrl+Shift+Enter`) always inserts the full path.

Quoting is shell-aware. Command Prompt sessions start with AutoRun and delayed expansion disabled, so a path containing `%NAME%` cannot expand.

### Working inside a terminal

- `Ctrl+Shift+F` searches the visible text and scrollback without sending anything to the shell.
- `Ctrl+Shift+.` searches this session's history plus your PowerShell history file and **inserts** the chosen line without running it.
- `Ctrl+Tab` / `Ctrl+Shift+Tab` and `Ctrl+Alt+1` … `Ctrl+Alt+9` move between tabs and pane content; `Ctrl+Shift+W` closes the selected session.
- `Ctrl+Shift+C` / `Ctrl+Insert` copy. `Enter` and `Ctrl+C` copy and clear an active selection; with no selection they reach the running program unchanged, exactly once.
- `Ctrl++` / `Ctrl+-` / `Ctrl+0` change that session's font size. The main-keyboard bindings follow physical key positions, so they work on any keyboard layout.
- Selection supports drag, `Alt`+drag for a rectangle, double-click for a word, and `Ctrl+Shift+A` for everything. Hold `Shift` to select while an application is using the mouse.
- Terminal colors follow the active RedSalamander theme, and an unfocused terminal dims like an unfocused folder pane. Windows Terminal profiles and color schemes are not imported.
- Programs that speak the Kitty graphics protocol can draw inline images.

### Terminal settings and keyboard

**Preferences → Plugins → Terminal** carries the default shell, font family (`Cascadia Mono`) and size (14), maximum paste size (4 MiB), the unsafe-paste warning, optional follow-when-idle, the hyperlink policy and the clipboard policy.

Three defaults are deliberately cautious and can be relaxed:

- A multiline or control-character paste asks for confirmation first.
- `Ctrl`+click on an `http`, `https` or `mailto` link asks before opening. Ordinary clicks never follow a link.
- A program writing your clipboard (OSC 52) asks. Programs can never *read* the clipboard.

**Preferences → Keyboard → Terminal** lists all 39 terminal commands. Any chord can be reassigned, set to **Pass through to terminal**, or set to **No action**; a terminal binding that shadows a global one is labelled **Overrides global shortcut**. While a terminal has focus, `F11`, `Ctrl+,`, `Ctrl+Shift+,`, `Ctrl+Shift+P`, `Alt+Enter`, `Alt+F4` and `Alt+Space` keep their global meaning.

Follow-when-idle (off by default) lets a PowerShell session change directory when you navigate the pane, only at an idle prompt with nothing typed. A `cd` in the shell never moves the pane.

---

## Command Palette, menus, and shortcuts

- `Ctrl+Shift+P` opens a searchable, command-centric palette showing each command's icon, name and current shortcut. Unavailable commands are shown disabled instead of failing silently.
- `F1` / **Help → Display Shortcuts** remains the binding-centric list, one row per key. Use the palette to find a command, `F1` to inspect keys.
- `Ctrl+,` opens Preferences; `Ctrl+Shift+,` opens the active settings file in your editor.
- Menus were reorganized so that each command has one obvious home:
  - **Preferences** moved from View to the end of **Commands**. **View Width** moved from Files to **View**.
  - **Create Directory** now lives only under **Files → New → Folder...** (`F7`).
  - The four copy-as-text commands moved into **Edit → Copy as Text**, and the selection tools into **Edit → Advanced Selection** (including Save and Restore Selection).
  - **Commands → Terminal** holds the floating window and the path-insertion commands; each pane menu gained **Terminal Pane**.
  - The three delete commands are now flat rows instead of a submenu.
- Folder panes have **separate context menus** for an item and for the background. The background menu offers Paste, New, Refresh and Calculate Occupied Space instead of a list of greyed-out item commands. The item menu gained **Cut**.
- Find Files results use a single action list that targets either the selection or the row you clicked, instead of duplicating every command twice.
- **Batch Rename** is present in the Czech, French, Japanese and Slovak menus again.

---

## Folder panes: current item and selection

Every non-empty pane now has exactly one **current item** (the keyboard cursor). Selection is a separate set of zero or more items.

- Arrow keys, `Home`, `End` and `PageUp` / `PageDown` move the current item **without** changing selection.
- `Esc` or a click on empty background **clears the selection** and **keeps** the current item.
- `Space` toggles the current item, updates folder-size work from what remains selected, and moves to the next item **without wrapping**. A second `Space` on a selected folder deselects it and cancels its size calculation.
- `Insert` toggles and advances without starting folder-size work.
- `Shift`+click and `Shift`+arrows move the current item to the end of the range and select the whole range; `Ctrl+Shift` adds to the selection.
- Back / Forward and other navigation restore the remembered **current item** when it is still listed. They never restore selection; use **Save Selection** and **Restore Selection** for that.
- Selection commands act on what is **displayed when you run them**. Clearing a filter afterwards does not silently reselect items that were hidden; run Restore Selection again.
- When the current item disappears (deleted elsewhere, filtered, hidden), the pane moves to the first surviving item after it in the current sort order, otherwise to the nearest one before it.
- A genuinely empty folder shows a **Go to parent** action for `Enter`, double-click and `Backspace`. It is not a file, and a filter that matches nothing does not show it.
- A drag now starts once the pointer leaves the system drag threshold, instead of requiring the pointer to cross the whole row.

---

## File operations

### Move is a rename where it can be

- Moving inside one volume, one SMB share, one FTP/SFTP/SCP connection, one Google Drive or one Microsoft Drive is a **single rename** — for files, folders, junctions and symbolic links alike. Large same-disk folder moves complete immediately instead of copying the tree and deleting the source.
- Moving a folder **onto an existing folder** relocates each non-colliding child by rename and only prompts for the children that actually collide. If a child had to stay behind, the result reads **"Moved; source folder kept"** instead of re-copying anything.
- Cross-volume and cross-provider moves still copy, verify, and remove the source only after success.

### Starting an operation

- `F5` (Copy), `F6` (Move), Paste, and drag-and-drop **start straight away**. Use **Copy with Options** (`Shift+F5`) and **Move/Rename with Options** (`Shift+F6`) to choose link handling, verification, queue versus parallel, and a bandwidth limit before starting.
- Every task shows a cancellable **Preparing…** phase before it touches anything, so a dead share or slow device no longer freezes the window.
- If part of a Move can only be copied, the count and the reason are shown before you start, and the result says **"Copied; source kept"**.

### Verification

- New setting **Preferences → File Operations → Verify copied files** (off by default) re-reads copied files and compares content. Cloud destinations use the provider's own checksum instead of downloading the file again.
- With verification on, a Move deletes the source only after the copy is proven. Otherwise it reports "Copied; source kept because content verification failed" (or "…canceled", "…unavailable").
- Progress shows verification as a second bar segment and a second graph series, and includes it in the estimate.

### Progress that tells the truth

- The separate pre-calculation phase is gone. One pass discovers and transfers at the same time, so bytes start moving immediately.
- While a folder is still being scanned the card shows real counters and an indeterminate bar rather than a percentage, with a provisional "still discovering" estimate. The Windows taskbar stays indeterminate until totals are known. A one-way **Skip discovery** button releases full concurrency at once.
- Results use plain words: Copied, Moved, Copied; source kept, Moved; source folder kept, Skipped, Outcome unknown, Verified, Completed with skipped items, Completed; cleanup item retained. Panes, Find Files and Compare Directories remove a source row only when its removal was proven.

### Conflicts and consent

- Conflict prompts offer **Overwrite**, **Keep both**, **Replace read-only**, **Replace link**, **Skip item**, **Retry**, and **Skip all** under **More…**, with Cancel as the default. A card first shows "Waiting for your decision" and keeps destructive buttons disabled until it has read the facts.
- A folder meeting a folder always merges and never asks to overwrite. A file meeting a folder offers only Keep both or Skip.
- Retry is genuinely repeatable when nothing was committed; the prompt shows how many attempts failed instead of quietly turning into Skip.
- Mid-operation consent prompts appear, each with the safe choice preselected: not enough free space, encrypted data that would be written as plaintext, sparse files that would inflate, cloud placeholders that must be downloaded, a Move that would lose Mark-of-the-Web, alternate data streams or extended attributes, and a Recycle Bin deletion that failed.
- If two tasks touch the same place, RedSalamander now says so: **"Another task is working in the same place"** offers Queue after these tasks, Run at the same time, or Don't start. Clearly separate folders never prompt.

### Renaming and artifacts

- Inline rename (`F2`) runs through the file-operations engine on a worker thread. A quick success is invisible; a slow or conflicting rename shows a task card.
- Batch Rename runs as an ordinary task with Cancel, and now **refuses rename cycles up front** (for example swapping two names) with an explanation, instead of executing half of them.
- Interrupted-operation leftovers (`.rs_tmp_`, `.rs_bak_`, `.rs_ren_`, `.rs_copy_tmp_`, `.~rs-write-`) are no longer hidden from panes, Find, Search or Compare. They carry a "Possible interrupted-operation artifact" badge and a Properties section explaining them. Viewing, copying and exporting them is unrestricted; renaming, deleting or overwriting one asks first. Nothing is deleted automatically.
- An interrupted Move leaves a note that is shown once on the next start, with Open source and Open destination.

### The File Operations popup

- Escape cancels the decision you are looking at and never means Cancel All. The window's close button only hides the popup; **View → File Operations** or `Ctrl+Shift+J` brings it back, and it reappears by itself when a decision needs you.
- Closing RedSalamander with work in progress asks whether to cancel everything, and drains the cancellation without blocking the window.
- Layout, colors and tooltips were reworked: per-file bars match their graph band, the throughput graph is no longer multicolored for a single transfer, counts read "Running / Waiting / Needs attention: N", and screen readers get the popup in visual order.

---

## Remote, cloud, and archive locations

- **FTP, SFTP and SCP**: full read/write destinations, with same-connection rename for files and folders. Uploads stage first and only then replace the target, so an interrupted upload can no longer destroy the file it was replacing. A server that cannot report a file's size can no longer be used as a Copy or Move destination, because a truncated upload could not be detected there.
- **Amazon S3**: creating a folder writes a real prefix marker; recursive delete re-checks until the prefix is empty; deletes and rollbacks are pinned to the exact object version observed, so a file replaced by another client is never overwritten or removed. Versioned buckets get delete markers, and public buckets can be browsed anonymously.
- **Microsoft Drive / OneDrive**: same-drive Move is a server-side move by item ID. Delete is a Recycle Bin operation. Ranged downloads validate the returned range, so a shifted response can no longer corrupt a copy.
- **Google Drive**: full destination with server-side copy, native move and rename. A retried request can no longer create a duplicate folder or copy, and duplicate names are disambiguated rather than guessed.
- **MTP devices**: object identity is now exact, so recovery and overwrite can never act on the wrong object. Move remains copy-and-keep-source.
- **SMB shares**: same-share moves are server-side renames, and a hung share no longer blocks Cancel or exit — cancellation returns in about a second.
- **IMAP** stays read-only by design, as do S3 Tables.

---

## Viewers

- **Viewer Text**: the long encoding cascade is replaced by a searchable **More Encodings...** picker over the full catalog, filtered by codepage number, name or alias. The View and Encoding menus were regrouped, and choosing a display encoding is now separate from the Save Encoding policy.
- **Viewer Web**: **Toggle DevTools** (`F12`) moved from View to a new **Tools** menu.
- **Viewer Image/RAW**: "Other Files" moved under **File**, and the brightness, contrast and gamma commands are now named instead of shown as `+` and `-`.
- **Viewer PE** and **Viewer Space** menus were normalized to the same shape.
- Files opened over FTP, SFTP or SCP can no longer be shown as complete when the transfer ended early.

---

## Themes, mouse, and windows

- Cycling themes with `Shift+F11` / `Shift+F12`, choosing one from **View → Theme**, or clicking a theme command on the Function Bar shows a centered overlay: previous theme, current theme in large type, next theme. It stays readable for a moment, updates in place when you press again, never takes focus, and disappears when clicked. Choosing the theme that is already active does nothing.
- **Preferences → Mouse** gained two independent options, both off by default: focus follows the pointer, and focus follows the pointer only while a pane is showing its terminal.
- Windows reopen on the monitor they were closed on when it is still connected, and are never restored smaller than their minimum size.

---

## RedSalamander Monitor

- **Save As** runs in the background with byte and line progress, can be cancelled, and writes an exact snapshot of the moment you invoked it. Choosing Save As again while a save is running cancels it rather than starting a second one.
- **Open** decodes large files incrementally, so peak memory is far lower, and a file ending in a newline no longer gains a phantom empty last line. Open followed by Save As is now byte-identical.
- The **Options** menu was reordered and the message filters renamed to Error, Warning, Information, Performance, Debug and Trace, with clearer presets. Duplicate access keys were fixed in all four translations.
- Find (`F3`) stays in order when old lines are dropped by the retention limit, and toggling process and thread IDs no longer leaves the selection pointing at the wrong characters.

---

## RedConfigure and the search service

- RedConfigure's localization **paste matrix** now follows what you can actually see: it pastes into the visible, filtered, reordered rows and cultures, validates the whole rectangle first, and applies as one undoable step. A rectangle with any invalid cell is rejected instead of being partly applied.
- The search service serializes per index store instead of process-wide, so two different stores no longer block each other and a local user can no longer prevent the service from starting. `--run-foreground` and `--compact` now use the same store as the installed service.

---

## Accessibility, input, and localization

- Terminal contents, selection and caret are exposed to screen readers, and the terminal's security confirmations are readable in light and high-contrast themes.
- Screen-reader word, line and paragraph navigation moves by the requested unit and never splits a surrogate pair.
- Text fields treat combining marks as part of the preceding character, so caret movement and Backspace behave correctly in Vietnamese, Thai, Indic, Hebrew and Arabic text.
- Progress bars, status strips and the throughput graph expose proper roles, and tooltips work on non-interactive elements.
- New Terminal, palette, menu, conflict, consent and Monitor strings are resource-backed, and the Czech, French, Japanese and Slovak satellites carry real translations, including four new Terminal plugin satellites. Every menu row has a unique access key in all five languages.

---

## Reliability and security fixes worth knowing

- Closing a terminal tab or the application can no longer hang waiting for a stuck console process.
- Shells are resolved from validated system locations, so a program of the same name in the current folder cannot be launched instead.
- Opening a terminal from a location with no real path no longer opens a shell in an unrelated folder.
- File lists handed to external tools are written to a private folder and cleaned up by verified file identity, and file names containing control characters are rejected instead of corrupting the list.
- Settings recovery now works on the exact file it read: a valid file written by another instance can no longer be replaced by defaults, and an unreadable settings file is never treated as missing.
- A damaged shortcut, terminal or settings section is recovered on its own instead of resetting everything else.
- Interrupted archive, cloud and remote operations leave the original file intact.

---

## Important upgrade notes

- **There is no bottom command line.** `Ctrl+Enter`, `Ctrl+Space` and `Ctrl+Shift+Enter` insert into a terminal and never run anything. Use the terminal or a User Menu entry to execute commands.
- **`Ctrl+Alt+T` now opens the floating terminal window.** The embedded pane terminal is `Alt+7` (or `Ctrl+Shift+T`). If you had customized `Ctrl+Alt+T`, your binding is kept; only the default is migrated, once.
- **`Ctrl+Shift+Enter` is now Insert Full Path**, no longer a duplicate of Insert Focused Item.
- **`Space`, `Insert` and `Esc` follow the current-item model.** `Space` on an already selected item now deselects it, `Insert` no longer calculates folder sizes, and `Esc` keeps the cursor where it is.
- **`F5` and `F6` no longer ask before starting.** Use `Shift+F5` / `Shift+F6` when you want the options.
- **Menu locations moved**: Preferences is under Commands, View Width under View, Create Directory only under Files → New, Restore Selection under Edit → Advanced Selection, and the copy-path commands under Edit → Copy as Text.
- **Link handling changed.** The Preserve/Skip choice applies to Copy only; every Move relocates links as they are. "Follow targets" no longer exists and old settings become Skip. Preserve now copies a link exactly as written, so an absolute link inside a copied tree still points at the original location.
- **A cut list is consumed when the Move is queued** and is not restored if something fails. The task card offers "Cut retained items again" when the remaining set is known exactly.
- **Batch Rename no longer keeps a recovery journal**, and its Resume and Roll back commands are gone; refresh the folder after an interruption.
- **A renamed folder tree keeps its existing permissions** and does not recompute inherited ones, exactly as File Explorer behaves. An absolute link inside a moved tree keeps pointing at the old location.
- **NTFS compression is no longer carried across a copy**; the destination inherits its parent folder's setting, like File Explorer.
- **`fileOperations.preCalcEnabled` and `fileOperations.preCalcMaxWorkers` are retired** and ignored. The new key is `fileOperations.verifyAfterCopy` (default `false`).
- **Interrupted-operation files are now visible** in panes, Find and Search results, badged rather than hidden.
- **Settings schema is still 16.** Existing files load unchanged. New sections (`mouse`, `terminal.floatingWindow`, and the Application and Terminal shortcut scopes) appear with defaults until you change them.
- **Third-party plugins must be rebuilt.** The host and viewer plugin interfaces moved to size-prefixed records; plugins built against v7.0.48 headers are not binary compatible.
- Terminal hyperlinks, clipboard writes by programs, and unsafe pastes ask for confirmation until you change the Terminal plugin settings.

## Not in this drop

These are specified or under way but deliberately not part of this release:

- **Preferences → Confirmation** and the related "completed with warnings" outcomes.
- Undo for file operations, and a persistent searchable operation history.
- Retargeting links so they point inside a copied tree.
- More than one embedded terminal per pane, a shell chooser per session, remembered shells per folder, terminal profiles, and per-folder terminal history.
- SSH, Git Bash, MSYS2 and Cygwin sessions. WSL sessions work but are terminal-only: no follow, no history suggestions, no short-name insertion.
- Native Move on MTP devices, and same-drive server-side Copy on Microsoft Drive.

## Availability and requirements

- Windows 11 build 22000.2600 or later, x64 or ARM64.
- Portable ZIP packages for x64 and ARM64 with a SHA-256 checksum manifest, plus the usual Winget package. Both packages include the Terminal plugin and its engine, and every release is now gated on a native ARM64 qualification run.
- Third-party notices for the terminal engine ship in `Plugins\TerminalRuntime\`.
- The next GitHub release will name the actual version and attach the packages. Until then, treat this file as the changelog for the unreleased tree after v7.0.48.
