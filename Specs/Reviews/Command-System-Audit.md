> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


Based on the verified findings, here is the final audit report:

---

# RedSalamander Command System — Correctness & State-Safety Audit: Final Report

## 1. Verdict

The command system is **structurally sound in its registry/routing invariants but dangerously thin on precondition guards at the handler layer**, and it leaks irreversible filesystem mutations through paths the UI does not gate. The single most dangerous theme is **"enable-state is the only guard"**: destructive handlers (delete, permanent-delete, rename, change-attributes, change-case, pack/unpack) are reached directly by accelerators/WM_COMMAND with no precondition re-check, so a shortcut or a stale menu can hit the wrong target, the wrong pane, or an unconfirmed irreversible action. The second theme is **silent overwrite of existing files** — Unpack and Make-File-List clobber on-disk data with no confirmation and no recycle-bin fallback, and Pack's "delete sources" can destroy its own output. The third is **TOCTOU on the selection/target** across async enumeration and modal message pumps (copy/move-to-other-pane uses the *requested* not *enumerated* folder; Find's context menu acts on a positional row index captured before the menu pump; ChangeCase snapshots selection *after* its dialog). A fourth, lower-severity but pervasive theme is a **dispatch gap**: five parameterized command families are advertised as keyboard-assignable yet have no string-dispatch handler, so binding them yields a guaranteed "not implemented" popup. Shell-argument handling has one real defect (cmd `/K pushd` allows `%VAR%` expansion of the working dir). Most data-loss findings are real and code-confirmed; the residual risk is concentrated in modal-reentrancy and focus-race scenarios that need interactive testing.

---

## 2. Findings by Severity

### TIER 1 — Data-loss / Crash / Wrong-target-or-pane / Security

---

**Unpack silently overwrites existing destination files (overwrite hard-coded `true`)**
`RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp:11702` — **data-loss / CONFIRMED**
`CommandUnpack` initializes `bool overwrite = true;` and the interactive prompt branch (11718–11726) populates destination/unpacker/deleteArchive/maskText but **never assigns `overwrite`** — `ArchiveUnpackPromptResult` (3801) has no overwrite member and the dialog has no replace control. `overwrite==true` flows to `ExtractStoredZipArchive`/`ExtractSevenZipArchive` (11791/11794), which use `MOVEFILE_REPLACE_EXISTING` (7638/7911), bypassing the `if (! overwrite && exists) return ERROR_FILE_EXISTS` guard (7575). **Scenario:** select `archive.zip`, press Alt+F6/Alt+F9, accept destination = current folder; any colliding existing file (e.g. an edited `report.docx`) is silently replaced with no recycle-bin recovery. Violates `Specs/UI/UI_CommandMenuKeyboard.md:242`. **Fix:** mirror `CommandPack` (which sets `overwrite = false;` at 11543) — default false in the prompt branch, or add a real overwrite/skip/rename choice to the struct and dialog, surfacing a collision confirmation. *(Three separate audit units independently confirmed this same defect.)*

---

**Pack "delete sources after" can delete the freshly-created archive**
`RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp:8826` — **data-loss / CONFIRMED**
`CommandPack` writes the archive to the user-chosen `archivePath` (11601) then calls `DeletePackedSources(selectedPaths)` (11622), which runs `std::filesystem::remove_all` (8826) on each selected directory — **with no check that `archivePath` is not equal to or contained within any selected path**. `ReadResultFromUi` (5007–5026) only trims/uniquifies the name; no canonicalize/ancestor comparison exists anywhere. **Scenario:** user selects folder `Foo`, enables "delete sources after packing", and types an archive path resolving under the selection (e.g. `Foo\Foo.zip`); pack succeeds, then `remove_all(Foo)` permanently deletes both the originals and the new archive (bypasses recycle bin). The default path is always a sibling, so this requires deliberate path entry — but the irreversible-delete mechanism is entirely unguarded. **Fix:** canonicalize and reject the delete (or the prompt) when `archivePath` is equal to or a descendant of any selected path.

---

**`makeFileList` silently overwrites an existing output file (CREATE_ALWAYS, no confirmation)**
`RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp:2893` — **data-loss / CONFIRMED**
`WriteMakeFileListUtf8File` opens the destination with `CreateFileW(... CREATE_ALWAYS ...)`, truncating any existing file; the only precondition is `path.empty()`. The prompt is a plain text field with no Save-As dialog and no existence check, and `outputFile` is persisted to settings (11448) so the same path is re-truncated every run. **Scenario:** the output-file field is pre-filled (prior run / pasted) with the path of a real document; clicking OK irreversibly replaces its contents with the listing, no prompt. **Fix:** `std::filesystem::exists` check + confirmation (or `IFileSaveDialog` with its standard overwrite prompt), matching the pack/unpack overwrite convention already in this file.

---

**Inline rename accepts path separators and `..`, relocating the item outside the displayed folder**
`RedSalamander/FolderView.FileOps.cpp:931` — **wrong-target-or-pane / CONFIRMED**
`RenameFocusedItem` builds `target = fullPath.parent_path() / prompt.text` and passes it straight to `RenameItem`. The only sanitization is whitespace-trim + empty-check; nothing rejects `\`, `/`, `.`, or `..`. `operator/` appends separators as sub-path components, and the builtin plugin runs the destination through `GetFullPathNameW` (`FileSystem.Path.cpp:32`), collapsing `..` and resolving absolute strings — so renaming `report.txt` to `..\sibling` **succeeds**, moving the file to the grandparent; a typed drive/UNC path targets anywhere. The sibling **batch**-rename path already guards this via `ContainsPathSeparator` (`BatchRenameEngine.cpp:600`), proving the single-item omission. **Scenario:** F2 on a pre-selected name, type `..\report.txt`, Enter → file silently leaves the folder; user believes it vanished. **Fix:** validate `prompt.text` at the command layer — reject separators, `.`/`..`, Win32-illegal chars, reserved device names, trailing dot/space; keep result leaf within `parent_path()`.

---

**ChangeCase captures the selection AFTER its modal dialog (selection TOCTOU)**
`RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp:12965` — **wrong-target-or-pane / CONFIRMED**
`CommandChangeCase` opens the modal `PromptForChangeCase` (12965) and only **afterward** reads `GetSelectedOrFocusedPaths()` (12971). Pack/Unpack/ChangeAttributes all snapshot **before** their dialog (11504/11685/12736). The modal's own message loop (`ShowModal`, 3564–3579) `EnableWindow(owner,FALSE)` blocks *user* input but not file-watcher-driven selection updates dispatched to the folder view. **Scenario:** user selects 3 files, opens Change Case; a folder refresh clears the selection while the dialog is up; on OK the irreversible case-rename applies to the focused fallback item (or no-ops with the dialog work lost). **Fix:** capture `paths` and bail-on-empty before showing the dialog, like the sibling commands.

---

**Copy/Move-to-other-pane writes into the other pane's REQUESTED (not enumerated) folder**
`RedSalamander/FolderWindow.FileOperations.cpp:1149` — **wrong-target-or-pane / CONFIRMED**
`SanityCheckBothPanes` requires only `dest.folderView.GetFolderPath().has_value()`, then passes `GetFolderPath().value()` as the destination (1238/1279). `GetFolderPath()` returns `_currentFolder` — the *requested* folder set synchronously in `SetFolderPath` **before** async enumeration completes or succeeds. The dedicated predicate `IsCurrentFolderEnumerated()` (`FolderView.cpp:291`) exists precisely to confirm displayed contents match `_currentFolder`, but the copy/move check never calls it. **Scenario:** the other pane is mid-navigation into a slow/offline UNC or access-denied dir; F5/F6 from the active pane passes the sanity check and the op lands files in (or, for MOVE, sources are touched before a late failure against) an unconfirmed destination the user never saw loaded. **Fix:** require `dest.IsCurrentFolderEnumerated()` (and ideally the same for src) before proceeding, rejecting with a localized "destination still loading" message.

---

**Same-folder self-move guard uses `_currentFolder`, not the items' actual parent**
`RedSalamander/FolderWindow.FileOperations.cpp:1161` — **data-loss / PLAUSIBLE**
The same-folder rejection compares `src._currentFolder` vs `dest._currentFolder`, but the moved paths are built from `_itemsFolder` (`GetItemFullPath`, `FolderView.cpp:124`) — the folder the *displayed* items belong to, which diverges from `_currentFolder` mid-navigation. **Scenario:** source pane navigates away from folder X (where the still-selected, X-rooted items live) toward the destination's folder; the guard sees `_currentFolder`(Y) ≠ dest(X) and passes, but the X-rooted selection is moved into X itself — a self-move whose overwrite handling can destroy the originals. Requires a specific mid-enumeration timing window plus a matching pane config, hence PLAUSIBLE. **Fix:** gate on `src.IsCurrentFolderEnumerated()` and/or derive the same-folder comparison from the items' actual parent folder.

---

**HideSelectedNames leaves hidden files selected; a follow-up Delete/Attr/Copy hits invisible files**
`RedSalamander/FolderView.Selection.cpp:531` — **data-loss / PLAUSIBLE**
`HideSelectedNames()` adds names to `_hiddenNames` and calls the **asynchronous** `RequestRefreshFromCache()`, but never clears the selection or recomputes stats (unlike the sibling `DeselectSameNameWithoutExtension`). Between the hide gesture and the async re-enumeration, `_items` still holds the hidden entries with `item.selected==true`, and destructive resolvers (`GetSelectedOrFocusedPaths`) iterate by `item.selected` with no `_hiddenNames` filter. **Scenario:** user hides selected files (rows disappear → believes them excluded), immediately presses Delete (key-repeat / fast two-key) before re-enumeration posts → recycles/deletes files they can no longer see. Timing-dependent, hence PLAUSIBLE. **Fix:** clear selection synchronously for hidden items (set `selected=false`, `RecomputeSelectionStats`, `NotifySelectionChanged`) before the async refresh, or have `GetSelectedPaths` skip names in the active `_hiddenNames` filter.

---

**CommandSelectionRestore restores by name into any pane/folder with no source-identity check**
`RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp:12541` — **wrong-target-or-pane / CONFIRMED**
`CommandSelectionSave` records `sourcePluginId`/`sourceInstanceContext`/`sourceFolder` (11150–11155) but `CommandSelectionRestore` reads **only** `displayNames`, matching by name against whatever pane is focused, with `clearExistingSelection=true`. No check that the active pane is the saved source. **Scenario:** user saves a selection in folder A, navigates/switches to folder B (a backup/sibling with same-named files), presses the default **Ctrl+Shift+F6** (`ShortcutDefaults.cpp:193`) → B's selection is unconditionally cleared and same-named files in B are selected; a subsequent destructive command targets the wrong files. The clear also fires when zero names match. **Fix:** compare active-pane folder/plugin/instance against the captured source identity; confirm or warn on mismatch, and skip the clear when nothing matches.

---

**Batch-rename allows trailing-space / edge-space / trailing-dot targets (Warning only) → un-addressable files**
`RedSalamander/BatchRenameEngine.cpp:644` — **data-loss / CONFIRMED**
`HasEdgeSpaceOrTrailingDot` is classified `IssueSeverity::Warning`, which does **not** increment `errorRows`, so the Rename button stays enabled (`BatchRenameWindow.cpp:3358`) and `ExecuteRename` proceeds (3653 blocks only on `errorRows`). The plugin moves via `\\?\`-prefixed extended paths (`FileSystem.FileOps.cpp:4943/4956`), which **suppress** Win32's normal trailing-space/dot stripping. **Scenario:** template `{stem} ` (trailing space) shows only a yellow warning; clicking Rename produces on-disk leaves like `report ` that ordinary Explorer/Win32 open/rename/delete cannot address (recoverable only via extended-path tooling). **Fix:** promote `HasEdgeSpaceOrTrailingDot` to `Error` (matching the existing `name_reserved_device` treatment), or gate execution behind explicit confirmation for such rows.

---

**Find results context-menu acts on a positional row index captured before the modal menu pump (TOCTOU)**
`RedSalamander/FindFilesWindow.cpp:4085` — **wrong-target-or-pane / CONFIRMED**
`OnGridContextMenu` passes a positional grid index into `ShowResultContextMenu`, which calls the modal `ContextMenu::Show` (4060) — a nested message pump. During the pump, a posted file-operation-completed message fires `OnFolderWindowFileOperationCompleted` → `RemoveKeysFromResults` (4676), which rebuilds `_results` and shifts every index below a removed row. On return, `DispatchResultContextMenuCommand` re-checks only out-of-bounds (4077) and sets `_resultCommandIndexOverride = {clickedRowIndex}` (4085), consumed positionally against the mutated list. **Scenario:** user has an in-flight find-result move/delete; right-clicks a different row; the task completes during the open menu, shifting rows up; user picks "Permanently delete" → an irreversible op hits a file they never selected. **Fix:** capture the clicked item's stable key/`fullPath` before showing the menu, re-resolve the index from that key on return, and abort if the key is gone.

---

**Find copy/move-to-other-pane resolves the destination live; displayed destination can diverge from the actual target**
`RedSalamander/FindFilesWindow.cpp:3416` — **wrong-target-or-pane / PLAUSIBLE**
With no explicit destination set, `CopyOrMoveSelectedResultsToOtherPane` defers to `StartFileOperationForResolvedPathsToOtherPane`, which resolves the destination **live** at execute time (`dest.folderView.GetFolderPath()`, `FolderWindow.FileOperations.cpp:835`) and starts the op with `requireConfirmation=false` (860–872). The "Destination:" status text is computed separately and refreshed only on a few events — **not** on main-window pane navigation. Since the Find window is modeless, the user can navigate the destination pane between reading the status and triggering the move. (The finding's stronger "focus-flip flips source/dest" narrative is guarded by `pathFitsPane`; the real, reachable risk is stale-display + live-resolve for an irreversible MOVE with no confirmation.) **Fix:** freeze and pass the resolved destination explicitly (the `_explicitDestinationFolder` path already exists), or confirm the move using the exact resolved path so displayed and executed cannot diverge.

---

**Destructive pane commands are enabled while a compare run is still scanning and auto-rewriting the selection**
`RedSalamander/CompareDirectoriesWindow.cpp:871` — **data-loss / CONFIRMED**
The `IDM_PANE_DELETE/PERMANENT_DELETE/RENAME` block gates only on `_compareStarted` (latched true once), not on `_compareRunPending`/`scanActiveScans`/`contentPendingCompares`. Meanwhile `ApplySelectionForFolder` (2410) **programmatically rewrites** the pane selection to the current diff set every time a pane finishes enumerating or a decision-refresh fires mid-run. The keyboard path bypasses menu state entirely (`SendMessage(WM_COMMAND)`, 1173). **Scenario:** user starts a compare of two large trees, sees an initial highlighted set, presses Delete while the run is still pending; a content-compare completion expands the selection between the user's glance and `CommandDelete`'s snapshot → they confirm a delete for files they never intended (irreversible with `IDM_PANE_PERMANENT_DELETE`, which omits the recycle-bin flag). **Fix:** gate the destructive cases on the run being idle (`!_compareRunPending && scanActiveScans==0 && contentPendingCompares==0`) and disable the menu items in `UpdateViewMenuChecks` while a run is pending.

---

**Connection Manager modal facade ignores the modeless single-instance → concurrent editors clobber saves**
`RedSalamander/ConnectionManagerWindow.cpp:3696` — **data-loss / CONFIRMED**
`ShowDialog()` (modal facade, used by plugin connect prompts via `HostServices.cpp:1419`) never consults or registers `g_singleInstance` (skipped when `_isModalFacade`), and disables only `effectiveOwner` — not the already-open modeless window. Each `WindowImpl` snapshots its own `_connections = _settings->connections->items` (2963) and `SaveConnectionsSettings` overwrites `_settings->connections` wholesale (2660) plus deletes secrets for removed connections. The in-process stale-save guard (`ResolveStaleSettingsBeforeSave`) never fires for in-process facade saves because the save's own stamp is ignored by the directory watcher. **Scenario:** user edits connection A in the open modeless manager; a plugin opens a second modal-facade manager and saves connection X; the user returns to the stale modeless window and clicks Connect/Close → X's profile is silently deleted from settings while X's credential is orphaned in Credential Manager. **Fix:** make the modal facade participate in single-instance discipline (refuse/surface the existing window, or register its HWND), and re-load `_connections` from `g_settings` on focus.

---

**Monitor-settings save failure discards already-applied main settings on close (partial apply)**
`RedSalamander/Preferences.Dialog.cpp:2486` — **data-loss / CONFIRMED**
`SaveSettingsFromDialog` writes main settings to disk first (2474), then returns `SaveMonitorSettingsFromDialog` (a separate file). If only the monitor write fails, `CommitAndApply` treats the whole thing as failure and returns early, so `*state.settings`/`baselineSettings` are never updated — in-memory still holds OLD values while disk has NEW. On close, `WM_NCDESTROY` unconditionally re-saves `*state.settings` (5418), **overwriting the freshly-written main-settings file with the stale pre-apply copy**. **Scenario:** user changes a General setting AND a Monitor setting, clicks Apply; the monitor file is locked (RedSalamanderMonitor running / AV / read-only); main settings persist but the user sees only a monitor error; on close the applied main change is silently lost. **Fix:** commit the main-settings success (update in-memory + baseline + post `kSettingsApplied`) before attempting the monitor save, and don't let `WM_NCDESTROY` overwrite a newer on-disk file with a stale in-memory copy.

---

**ResetShortcutsToDefaults replaces the entire user keymap with no confirmation**
`RedSalamander/Preferences.Keyboard.cpp:2697` — **data-loss / CONFIRMED**
A single click on the Keyboard pane "Restore defaults" button (which sits next to Assign/Remove) calls `ResetShortcutsToDefaults`, which unconditionally does `state.workingSettings.shortcuts.emplace(CreateDefaultShortcuts())` (2704), discarding every custom binding. The only guard (2699) is a `keyboardCaptureActive` precondition, **not** a confirmation — and `ShowDialogAlert` is used elsewhere in this very file for error reporting, proving the facility was available. **Scenario:** user has dozens of custom shortcuts, accidentally clicks "Restore defaults"; entire keymap wiped in one click; only recovery is to Cancel the whole dialog (losing all other pending edits) — Apply/OK persists the wipe permanently. **Fix:** confirmation alert (Yes/No, existing IDS_ resource) before emplacing defaults.

---

**`cmd.exe /K pushd` quotes with argv rules but cmd performs `%VAR%` expansion of the working dir**
`RedSalamander/FolderWindow.FileSystem.Navigation.Part.cpp:150` — **wrong-target-or-pane / CONFIRMED**
`BuildCmdCommandShellLaunchPlan` builds `L"/K pushd "` plus the output of
`Common::Process::QuoteWindowsCommandLineArgument(workingDirText)`. The shared
helper correctly implements Windows argv quoting (escapes `\` and `"`, never
`%`), but cmd.exe performs environment-variable substitution on the `/K` tail
even inside double quotes. **Scenario:** user navigates into a UNC share whose
component literally contains a defined token (e.g. `\\server\%TEMP%share`) and
invokes the legacy external-shell selftest seam with no Windows Terminal
installed; cmd expands `%TEMP%` and `pushd` lands in the wrong directory or
fails. Quotes still neutralize spaces/`&`/`^`, so impact is wrong-target, not
injection. The normal product `cmd/pane/openCommandShell` path now opens the
embedded Terminal and does not reach this legacy fallback. **Fix for the
retained seam:** launch `cmd /V:OFF` and/or double `%` to `%%` in the pushd
argument, or prefer the `lpDirectory` mechanism used in the non-UNC branch.

---

**Edit/AlternateEdit launches an external editor on a focused DIRECTORY (View rejects directories, Edit does not)**
`RedSalamander/FolderWindow.Viewers.cpp:1248` — **wrong-target-or-pane / CONFIRMED**
`RequestViewFocusedItem` explicitly refuses directories (`FolderView.FileOps.cpp:388`), but the Edit path (`CommandEdit` → `TryEditFocusedFileWithEditor` → `GetFocusedPath`) has no such guard, and `FileActionResolver::ActionAppliesToContext` returns true for any path under a catch-all editor action with no `is_directory` check. The Edit menu item is statically `MFS_ENABLED` with no `EnableMenuItem` gate. **Scenario:** user has a catch-all editor association (common — text editor mapped to all files), focuses a sub-folder, presses F4 (or Ctrl+Shift+F4) → the editor is launched with the directory path as argument, while View on the same folder is correctly refused. **Fix:** mirror the View guard — reject (return false / show the unavailable overlay) when the focused item is a directory.

---

### TIER 2 — State-corruption / Reentrancy-race / Dispatch-gap

---

**[MERGED] Five parameterized command families have no string-dispatch handler → "command not implemented" on any bound shortcut**
`RedSalamander/RedSalamander.cpp:6527` (fall-through) / `:6535` (`ShowCommandNotImplementedMessage`) — **dispatch-gap / CONFIRMED**
Affected commands: **`cmd/pane/navigatePath/<path>`, `cmd/app/openFileExplorerKnownFolder/<folder>`, `cmd/app/plugins/configure/<id>`, `cmd/app/plugins/toggleEnabled/<id>`, `cmd/pane/selectFileSystemPlugin/<id>`** (plus the bare canonical `cmd/app/theme/select`, `cmd/pane/viewWith`, `cmd/pane/editWith`, `cmd/pane/userMenu` when bound without a suffix, and `theme/systemHighContrastIndicator`).
**Root cause (single):** these are registered in `kCommands` with `wmCommandId==0` (`CommandRegistry.cpp:41,44,46,51,122,137`) and are implemented *only* via transient dynamic-menu id maps (`g_navigatePathMenuTargets`, `g_pluginMenuIdToPluginId`, `IDM_APP_OPEN_FILE_EXPLORER_*`), never as a `commandId == ...` branch in `ExecuteCommandById`. `CanonicalizeCommandId` strips the suffix; `TryGetWmCommandId` returns `nullopt` (wmId 0); control falls to `ShowCommandNotImplementedMessage`. The Keyboard Preferences UI enumerates **every** command from `GetAllCommands()` with no filter for these (`Preferences.Keyboard.cpp:2100`), so they are presented as fully assignable. `ShortcutManager` stores/returns the bound id verbatim. **Scenario:** user binds a key to e.g. "Open File Explorer Known Folder" or "Select File System Plugin" or "Navigate Path" and presses it → a modal/modeless "command not implemented" popup instead of the action, even though the identical command works from the menu. The dispatch smoke self-test masks this by auto-closing the transient alert (`Commands.SelfTest.ViewCommands.cpp:1429`) and skipping `openFileExplorerKnownFolder`. **Fix (one):** either add explicit suffix-parsing handlers in `ExecuteCommandById` for each family (capture the suffix *before* `CanonicalizeCommandId`, validate, call the same code the menu uses — `SetFolderPath`/known-folder `ShellExecute`/`SetActivePlugin`+`SetFileSystemPluginForPane`/configure-toggle), **or** exclude pure-menu/parameterized-base commands from the keyboard-assignable surface so they cannot be bound. *(This single root cause was reported as 7 separate findings across registry, dispatch-core, wmcommand-routing, navigation, clipboard-text-explorer, and plugins-themes units — all the same defect.)*

---

**OK with a live theme preview but no net change leaks the un-committed preview to disk**
`RedSalamander/Preferences.Dialog.cpp:5135` — **state-corruption / CONFIRMED**
`ApplyThemeTemporarily` writes the preview into `*state.settings` (==`g_settings`) and sets `previewApplied=true`, but reverting the combo to the original (`ThemesThemeChanged`) only touches `workingSettings`, making `IsDirty` false **without** re-syncing `*state.settings.theme`. The IDOK handler skips both `CommitAndApply` and `RestorePreviewAppliedPreferencesOnCancel` when `dirty==false`, then closes; `WM_NCDESTROY` unconditionally persists `*state.settings` (5418). **Scenario:** preview theme B, dislike it, reselect theme A, click OK → theme B is kept in runtime and written to disk despite the user backing out. **Fix:** call `RestorePreviewAppliedPreferencesOnCancel` on the IDOK path whenever `previewApplied` and commit is skipped, and add a preview-restore safety net in teardown.

---

**SwapCapturedShortcut leaves a stale same-chord unassigned placeholder → two bindings on one chord**
`RedSalamander/Preferences.Keyboard.cpp:2610` — **state-corruption→incorrect-behavior / CONFIRMED**
`ApplyCapturedShortcut`'s conflict scan **skips** unassigned placeholders (2434), so a placeholder on the captured chord doesn't count; `SwapCapturedShortcut` then swaps only target↔conflict and erases nothing — unlike `CommitCapturedShortcut`, which erases all same-chord bindings (its skip test is `.empty()`, and the placeholder id is non-empty, so it *is* removed). **Scenario:** a real binding X and an unassigned placeholder both hold chord C (reachable via Remove-on-default or unvalidated Import); capturing C for Y triggers the single-conflict swap, leaving the placeholder on C → redundant/stale entry polluting working settings, export, and displayed rows. *(Downstream runtime misrouting is NOT realized — `ShortcutManager::LoadBindings` tolerates real+unassigned pairs and the real binding deterministically wins; the defect is the leftover placeholder, not a load-order conflict.)* **Fix:** erase any other binding (including placeholders) whose chord equals the captured chord before swapping, or recompute `conflictIndex` with the same skip rule commit uses.

---

**CommitAndApply has no reentrancy guard; a nested overlay pump allows double-commit**
`RedSalamander/Preferences.Dialog.cpp:2710` — **reentrancy-race / PLAUSIBLE**
The modeless dialog's IDOK/Apply handlers gate only on `state->dirty`, which is cleared only *after* all work completes (2737), and the Apply button stays enabled during the work. `HostShowPrompt` runs a real nested message loop (`HostServices.cpp:2059`) **without** disabling the owner, so a queued WM_COMMAND (key-repeat/double-click/accelerator) is dispatched while a prompt is up. That prompt is reachable from `CommitAndApply` via `ResolvePreferencesStaleSaveConflict` (2717) *before* the save, while `dirty` is still set → genuine re-entry → duplicate disk writes and redundant `kSettingsApplied`/`kPluginsChanged` (full plugin reload). *(Note: the finding's primary cited trigger — the snapshot `SendMessageW` "pumps messages" — is a misread; same-thread `SendMessage` does not pump. Real reentrancy requires an overlay prompt: stale-save conflict, reset-all, or dirty-close. The bug class — missing guard + non-disabling overlay pump — is real but the trigger is narrower than stated.)* **Fix:** add a `commitInProgress` flag set at the top of `CommitAndApply` via `wil::scope_exit`; early-return from the command handlers while set; disable OK/Apply for the commit duration.

---

**Non-recursive ChangeAttributes runs while a recursive ChangeAttributes worker is in flight on the same pane**
`RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp:12904` — **reentrancy-race / CONFIRMED**
The recursive branch guards re-entry with `if (changeAttributesThread.joinable()) { MessageBeep; return; }` (12796); the synchronous non-recursive branch (12904) has **no** such check and runs `ApplyChangeAttributesToPath` on the UI thread regardless of an outstanding background recursive run on the same pane. Both operate via `IFileSystemIO` on overlapping paths. **Scenario:** user starts a recursive Change Attributes on a large tree, then re-invokes choosing non-recursive on the same/overlapping selection → two unsynchronized writers produce interleaved/last-writer-wins attribute & timestamp state, and the completion overlay reflects only the worker's view. **Fix:** hoist the `joinable()` in-flight guard above the recursive/non-recursive branch split.

---

**Banner Rescan / `IDM_COMPARE_RESCAN` not disabled during an in-progress copy/move sync**
`RedSalamander/CompareDirectoriesWindow.cpp:1028` — **state-corruption / CONFIRMED**
The rescan handler is busy-aware only of scan/content phases, not of an outstanding async `StartCompareSyncToOtherPane` file op (which it doesn't even track — `taskIdOut=nullptr`). Once the compare run is done but the copy/move is still executing, the else-branch runs `ScheduleBeginOrRescanCompare` → `SetRoots`/`StartScan` over trees being concurrently written. **Scenario:** user invokes Copy-to-other-pane, then clicks Rescan while the async copy runs → the session is invalidated and a fresh scan starts over a destination tree being written, producing transient/incorrect decisions; the subsequent completion-invalidation races the new scan. **Fix:** track an operation-in-progress flag and disable `IDM_COMPARE_RESCAN` (and pane copy/move/delete) until the file op completes, or early-return from the rescan handler while a sync is outstanding.

---

### TIER 3 — Incorrect-behavior / Contract / Robustness (compressed)

**Incorrect-behavior:**
- **Clipboard Copy lacks Cut's builtin/local-FS guard** — `FolderView.FileOps.cpp:676` (CONFIRMED). Ctrl+C on a 7z/cloud pane puts a virtual path into a `CF_HDROP`; Explorer/in-app paste then operates on a non-existent path or `PasteShortcut` builds a broken `.lnk`. Fix: add the same `IsBuiltinFileSystemPlugin` guard Cut uses (702).
- **Clipboard Paste has no destination-FS guard** — `FolderView.FileOps.cpp:725` (CONFIRMED, effectively wrong-target). Pasting local `CF_HDROP` paths into a virtual pane hands `C:\...` to the virtual plugin's `CopyItems` as its own namespace (gate `CanSameFileSystemOperation` passes for any copy-capable virtual FS). Fix: mirror `PasteShortcut`'s `IsBuiltinFileSystemPlugin` requirement or set `sourceContextSpecified` so the cross-FS gate is exercised.
- **Move-paste does not clear the cut clipboard** — `FolderView.FileOps.cpp:798` (CONFIRMED). After a successful move, the `DROPEFFECT_MOVE` `CF_HDROP` persists; a second Ctrl+V re-attempts a move on now-deleted paths (or re-moves a recreated same-named file). Fix: `EmptyClipboard` / drop cut-state after a committed move; guard against in-flight double-invoke.
- **External edit/view launches the raw virtual-FS path** — `FolderWindow.Viewers.cpp:1206` (CONFIRMED). On 7z/cloud panes the external program gets an archive-internal/virtual path that doesn't exist locally (internal viewers work because they read through `state.fileSystem`). Fix: materialize to a temp local copy or refuse with an overlay, like `CommandContextMenuCurrentDirectory`'s builtin-FS guard.
- **Copy-UNC-as-text calls `WNetGetUniversalNameW` synchronously on the UI thread, once per item** — `FolderWindow.FileSystem.cpp:226` (CONFIRMED). Ctrl+Shift+Insert over files on an offline mapped drive freezes the UI for the full network timeout per item. Fix: skip resolution when `GetDriveTypeW != DRIVE_REMOTE`, resolve off-thread, and cap/timeout.
- **`makeFileList` collects entries synchronously on the UI thread (recursive walk)** — `FolderWindow.FileSystem.Commands.Part.cpp:11384` (CONFIRMED, robustness). "Whole folder"+recursive over a deep/network tree blocks the message pump with no cancel. Fix: run collection/render/write off-thread with cancellation.
- **`toggleThumbnails` never toggles off** — `RedSalamander.cpp:6276` (CONFIRMED). Both dispatch sites unconditionally `SetDisplayMode(Thumbnails)`; Alt+5 pressed twice is a no-op (the menu is a radio group, so the user isn't truly stuck — the real defect is the asymmetric misnamed "toggle"). Fix: remember and restore the prior list mode.
- **New-from-template drops the template extension when the typed name contains a dot** — `FolderWindow.FileSystem.Commands.Part.cpp:585` (CONFIRMED). `EnsureShellNewExtension` uses `path.extension().empty()`, so `config v1.2` (extension `.2`) skips the `.txt` append → file created without the template extension. Fix: use `EndsWithNoCaseText` like `BuildShellNewDefaultFileName` (351).
- **Shell-New template lookup by collapsed id can select the wrong template on a punctuation-collision** — `FolderWindow.FileSystem.Commands.Part.cpp:10395` (PLAUSIBLE). Lossy `LowerShellNewTemplateId` (every non-`[a-z0-9_-]` → `_`) plus no de-dup means two extensions collapsing to the same id resolve to the first via `find_if`; clicking the second creates the first's type. (The finding's `.h++`/`.h--` example is wrong — `-` is allowed; a real collision needs e.g. `.a+`/`.a#` → `a_`, config-dependent.) Fix: route by original extension or a stable index, and de-dup ids.
- **Rename does not normalize/reject trailing dot** — `FolderViewInternal.h:841` (CONFIRMED). `data.` is preserved through trim but Win32/`GetFullPathNameW` strips it, so on-disk name disagrees with typed name. Fix: strip or reject trailing `.` in `Confirm()`.
- **Pack/unpack delete-after-source runs a recursive permanent delete synchronously on the UI thread** — `FolderWindow.FileSystem.Commands.Part.cpp:8826` (CONFIRMED). `remove_all` on a large/network tree freezes the window with no progress/cancel (consistent with the rest of CommandPack being synchronous, but still blocks). Fix: route through the async `FileOperationState` pipeline.
- **ChangeCase has no local/read-only FS guard** — `FolderWindow.FileSystem.Commands.Part.cpp:12948` (CONFIRMED). Only `if (! state.fileSystem)`, missing the `IsFilePluginShortId` check that Pack/Unpack/ChangeAttributes have; on a cloud/archive pane it runs a full dialog + background recursive enumeration/rename that can leave the remote tree half-cased. Fix: add the sibling precondition + "requires a local folder" overlay.
- **Connect/Close persists an out-of-range TCP port** — `ConnectionManagerWindow.cpp:1954` (CONFIRMED). `wcstoul` cast to `uint32_t` with no upper bound; `99999` is stored verbatim (`TryValidateAndNormalizeConnectionProfiles` checks only the name). Fix: clamp to 1..65535 on parse and validate before save.
- **goDriveRoot navigates to not-ready drives** — `CompareDirectoriesWindow.cpp:1325` (PLAUSIBLE). The guard rejects only `DRIVE_NO_ROOT_DIR`, so an empty card reader (`DRIVE_REMOVABLE`) or disconnected mapped drive still drives the pane out of compare scope. (The "inconsistency" framing is refuted — the main-window handler is identical; and whether navigation commits to a not-ready drive can't be proven statically.) Fix: verify readiness (`GetVolumeInformationW`/`PathFileExistsW` on the root) before `SetFolderPath`.
- **After a partial/failed batch, `_targets` is rewritten to renamed paths; re-clicking Rename re-applies the rule** — `BatchRenameWindow.cpp:3916` (CONFIRMED). Non-idempotent rule (e.g. prefix `v_`) double-applies to the 60 successes when retrying 40 failures (`v_v_<name>`). Fix: mark renamed rows committed and exclude from the recomputed plan, or reset rule inputs.

**Robustness (terse list):**
- **hotPath/setHotPath suffix parser reads only `suffix[0]`** — `RedSalamander.cpp:6114` (PLAUSIBLE). `setHotPath/12` → slot 1; `hotPath/0extra` → slot 9; malformed suffix targets the wrong slot or silently no-ops. Only reachable via hand-edited/imported settings (defaults emit single digits). Fix: require exactly one `0-9` digit (validate whole substring with `from_chars`).
- **`kCommands` lacks a compile-time `wmCommandId`-uniqueness `static_assert`** — `CommandRegistry.cpp:224` (PLAUSIBLE). Sorted-id and prefix-mapping are statically enforced; uniqueness is only a runtime self-test (`Commands.SelfTest.Settings.cpp:6129`). A future macro collision would route via first-match `FindCommandInfoByWmCommandId`. Fix: add the consteval uniqueness assert.
- **Dynamic history id range is exactly adjacent to the hot-path base (zero gap)** — `RedSalamander.cpp:3885` (PLAUSIBLE). `IDM_*_HISTORY_BASE + 50 == IDM_*_HOT_PATH_BASE`; both write the same id-keyed map. Collision-free today but a `kHistoryMenuMaxItems` bump silently overlaps. Fix: `static_assert(HISTORY_BASE + kHistoryMenuMaxItems <= HOT_PATH_BASE)` and reserve headroom.
- **`cmd/pane/hotPaths` omits the `g_hFolderWindow` init guard every sibling has** — `RedSalamander.cpp:6519` (CONFIRMED). Calls `GetActivePane()`/`ResyncNavigationShellFromFolderView` unconditionally; benign (global object, harmless no-op on default state) but the sole sibling missing the guard. Fix: add the `g_hFolderWindow.load(acquire)` check.
- **`CommandDelete` lacks the `EnsureFileOperations()` guard `CommandPermanentDelete` has, and `FolderView::CommandDelete` re-posts `IDM_PANE_DELETE`** — `FolderWindow.FileOperations.cpp:1046` (PLAUSIBLE). A null `_fileOperations` path self-re-posts; reachable only during a teardown race that can't be deterministically triggered. Fix: call `EnsureFileOperations()` at the top and have `FolderView::CommandDelete` delete directly rather than re-post.
- **`theme/select/<name>` persists an unvalidated theme id** — `RedSalamander.cpp:8185` (two findings, CONFIRMED). `ApplyThemeId` writes `currentThemeId` before resolving; a bound shortcut to a since-deleted custom theme stores a dangling id (applied theme falls back to a valid mode, so no crash/data-loss — a phantom-string hygiene defect). Fix: validate via `FindThemeById` before assigning.
- **Disabled/greyed FS plugin can be activated and re-enabled via posted WM_COMMAND** — `RedSalamander.cpp:10566` (PLAUSIBLE). `g_pluginMenuIdToPluginId` is populated for disabled plugins and the handler doesn't re-check `disabled`/`loadable`; `SetActivePlugin` silently un-disables and persists. Grey is the only gate — reachable only by injected/automation WM_COMMAND (the command route is itself unimplemented). Fix: re-validate with `FindPluginById`, or don't map grayed entries.
- **Import accepts a `vk:0` binding (no lower-bound `vk` validation)** — `Preferences.Keyboard.cpp:3060` (PLAUSIBLE). Numeric JSON `0` stores a non-dispatchable chord-0 entry that still participates in conflict detection. Hand-edited/buggy-exporter file only. Fix: reject vk not round-trippable by `TryParseVkFromText`.
- **ImportShortcuts wholesale-replaces the keymap with an unvalidated file, no confirmation** — `Preferences.Keyboard.cpp:3173` (CONFIRMED, downgraded to robustness). Replace targets `workingSettings` (draft, Cancel-revertable), and `ParseShortcutsImportJson` accepts any `cmd/`-prefixed id without registry validation → keymap of dead bindings. Fix: confirm before replace; validate each id via `FindCommandInfo` and drop/report unknowns.
- **COM0/LPT0 falsely flagged as reserved device names** — `BatchRenameEngine.cpp:587` (CONFIRMED). `folded[3] >= L'0'` blocks the legal `COM0`/`LPT0`; Windows reserves only COM1-9/LPT1-9. Fix: change to `>= L'1'`.
- **Hardcoded English fallback strings** (contract/i18n, all CONFIRMED): ShellNew generic `"<ext> File"`/`"File"` label & caption — `FolderWindow.FileSystem.Commands.Part.cpp:409`; clipboard "Copied N items to clipboard." fallback with non-localizable `s`-suffix — `:11290`. Both violate the no-hardcoded-strings rule; only fire when the resource table is broken (resources exist in all shipped satellites). Fix: format from a dedicated IDS_ resource; treat a missing string as a build error.
- **Shell context menu for current directory omits IContextMenu2/3 forwarding** — `FolderWindow.cpp:622` (PLAUSIBLE, contract). No `WM_INITMENUPOPUP`/`WM_DRAWITEM`/`WM_MENUCHAR` forwarding to `HandleMenuMsg`/`HandleMenuMsg2`; lazily-populated owner-draw submenus (New / Send To / Open With) from shell extensions come up empty/garbled. Actual breakage depends on installed extensions. Fix: query `IContextMenu3` (fallback 2), store for the `TrackPopupMenuEx` lifetime, forward the messages.
- **`listOpenedFiles` focus acts on a stale snapshot row** — `FolderWindow.FileSystem.Commands.Part.cpp:12133` (downgraded to robustness). Rows refresh only on dialog re-open; a since-closed editor's row still navigates. Not wrong-target — the user explicitly asked to be taken there; worst case the focus item isn't found. Fix: re-validate the entry against live state at focus time.
- **`OnFunctionBarInvoke` (main window) routes without the `cmd/app/` scope assertion the keydown paths apply** — `RedSalamander.cpp:6544` (PLAUSIBLE). Safe today only because function-bar clicks post to `GA_ROOT`; a future compare/pane-only binding would `SendMessage` an unhandled WM_COMMAND → benign "not implemented". Fix: funnel both keyboard and mouse invocation through one resolver with canonicalize+unassigned+scope checks.
- **Shell-New registry enumeration runs synchronously on the UI thread** (menu build + command) — `FolderWindow.FileSystem.Commands.Part.cpp:10385` (CONFIRMED). Walks all of HKCR (thousands of subkeys) with no cache on every New-submenu open / template shortcut. Fix: cache with HKCR change-notification/TTL and/or enumerate off-thread.

---

## 3. Cross-cutting Themes

1. **Enable-state / menu-grey is the only guard, and the keyboard path bypasses it.** Destructive commands (`delete`, `permanentDelete`, `rename`, `changeAttributes`, `changeCase`, copy/move-to-other-pane) are dispatched directly from accelerators/`SendMessage(WM_COMMAND)` with no precondition re-check at the handler. Compare-window destructive commands (gated only on `_compareStarted`), the disabled-plugin re-enable (grey is the sole gate), and the static-`MFS_ENABLED` Edit-on-directory all share this shape. **The fix pattern is uniform: every destructive/state-mutating handler must re-validate its preconditions at entry, not rely on UI enablement.**

2. **Selection / target TOCTOU.** The selection or target is captured at one time and acted on at another, across an async enumeration or a nested modal pump: copy/move uses *requested* (`_currentFolder`) not *enumerated* folder; the same-folder guard vs `_itemsFolder`; ChangeCase snapshots *after* its dialog; HideSelectedNames leaves hidden items selected during async refresh; Find's context menu uses a positional row index captured before the menu pump; compare auto-rewrites the selection mid-run. **`IsCurrentFolderEnumerated()` and stable-key resolution exist but are not used in these paths.**

3. **Parameterized-suffix handling is incomplete and unbounded.** Five families have no string-dispatch handler at all (the dispatch-gap cluster); hotPath/setHotPath/goDriveRoot parse only the first char or only the upper bound; theme/select and import accept unvalidated ids that round-trip to disk; the Keyboard UI advertises every command as assignable with no filter. **Suffixes need exact, bounds-checked, allowlist-validated parsing, and the assignable surface must exclude pure-menu commands.**

4. **Irreversible writes without confirmation or recycle-bin.** Unpack overwrites (`MOVEFILE_REPLACE_EXISTING`), makeFileList truncates (`CREATE_ALWAYS`), Pack's `remove_all` can eat its own output, batch-rename produces un-addressable `\\?\` names, keymap reset/import wipe bindings — none confirm. The codebase *has* the right primitives (`overwrite=false` in Pack, `ConfirmPermanentDeletePaths`, `ShowDialogAlert`) but doesn't apply them consistently.

5. **Synchronous I/O on the UI thread** in command handlers: HKCR enumeration, `WNetGetUniversalNameW`, recursive `remove_all`, recursive makeFileList walk — all violate the CLAUDE.md no-blocking rule.

6. **Shell-argument quoting** is a narrow but real theme: argv-style quoting applied to a `cmd /K` tail that cmd re-parses for `%VAR%` expansion.

7. **Asymmetric guards between sibling handlers** — Cut vs Copy (local-FS), CommandDelete vs CommandPermanentDelete (`EnsureFileOperations`), recursive vs non-recursive ChangeAttributes (`joinable`), Swap vs Commit (placeholder erase), Pack vs Unpack (`overwrite`). One branch is hardened, the parallel branch is not. **A diff-the-siblings review pass would catch most of these.**

---

## 4. Coverage & Residual Risk

**Audited (static, code-confirmed):** the full registry & dispatch core (`CommandRegistry.cpp`, `ExecuteCommandById`, canonicalization, `TryGetWmCommandId`); WM_COMMAND/accelerator routing and the function-bar invoke paths; every data-mutating handler (delete/recycle, copy/move-pane, rename, clipboard cut/copy/paste, create-dir/new-from-template/editNew, attributes/case/pack/unpack, makeFileList); shell-execute/context-menu/properties; edit/view/open-with; navigation, hot-paths/user-menu, selection commands, sort/display/view options; app-level commands; and the preferences (apply/persist), keyboard-mapping, find/search, batch-rename, connections, compare-directories, and plugins/themes dialogs and dynamic menus. Routing was traced end-to-end (registry → WM_COMMAND/shortcut → handler) and default shortcut bindings were verified against `ShortcutDefaults.cpp`.

**Still needs a human / interactive test:**
- **Modal-reentrancy & nested-pump races** (`CommitAndApply` double-commit, Find context-menu TOCTOU, ChangeCase selection-during-modal, Connection Manager concurrent editors): the *mechanisms* are code-confirmed, but the win probability of the race needs live double-invoke / key-repeat / external-file-change testing.
- **Async-enumeration timing windows** (copy/move requested-vs-enumerated folder, same-folder self-move, HideSelectedNames-then-Delete): need a slow/offline UNC or a watcher-driven refresh to land the window reliably; PLAUSIBLE verdicts hinge on this.
- **Real shortcut/automation/imported-profile paths** for the parameterized commands and malformed suffixes: the standard UI doesn't generate multi-digit/garbage suffixes or bind the unimplemented families *automatically*, but the Keyboard UI **does** let a user bind them; confirm the actual "not implemented" popup and any imported-profile behavior interactively.
- **Shell-extension-dependent behavior** (IContextMenu2/3 empty submenus, ShellNew punctuation-collision, cmd `%VAR%` expansion, not-ready-drive navigation, batch-rename `\\?\` names): outcomes depend on the installed extensions, registry state, drive hardware, and live filesystem — verify on a real machine.
- **`SetFolderPath` commit semantics** for not-ready/failed navigation (goDriveRoot) and the teardown-ordering of `_fileOperations` reset (CommandDelete re-post) could not be fully resolved by static reading and need debugger/instrumented confirmation.

The highest-confidence, ship-blocking items are the **Tier-1 CONFIRMED data-loss/wrong-target findings** (Unpack overwrite, Pack self-delete, makeFileList truncate, inline-rename `..` escape, ChangeCase TOCTOU, copy/move requested-folder, SelectionRestore wrong-pane, batch-rename edge-space, compare-while-scanning delete, Connection Manager concurrent saves, monitor partial-apply, no-confirm keymap reset, cmd `%VAR%`, Edit-on-directory) — these are reachable by ordinary gestures or default shortcuts and should be fixed before the dispatch-gap and robustness clusters.
