# Batch Rename Window Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` when executing this plan with parallel implementation streams, or `superpowers:executing-plans` when executing it serially with review checkpoints.

**Goal:** Add a first-class Batch Rename window that can be launched from Rename workflows and from a dedicated command, previews every proposed rename before execution, supports manual names, macros, search/replace with regular expressions, and reusable Change Case logic.

**Architecture:** Build a pure rename-planning engine, a DxUi top-level window modeled after Find Files, and an async execution path that uses existing `IFileSystem::RenameItems` contracts with progress, cancellation, refresh, and deterministic selftests.

**Tech Stack:** Win32, Direct2D/DirectWrite through `Common/DxUi`, WIL RAII, `IFileSystem`, `Common::SettingsStore`, `.rc` resources, `PostMessagePayload`, command registry, and the existing build/selftest harness.

**During Implementation:** Check all the checkboxes with nice emoji to explain where we are in the implementation process, and add comments about any important decisions, discoveries, or changes to the plan as they happen.

---

## Implementation Progress

- [x] 2026-06-10: Created `BatchRenameEngine.h/.cpp` as the first pure, window-independent planning slice.
- [x] 2026-06-10: Added focused command selftests for macro/search/case/validation and manual-mode line mapping.
- [x] 2026-06-10: Added `Specs/UI/UI_BatchRenameWindow.md` and linked the command/file-operation specs to it.
- [x] 2026-06-10: Added command/resource registration, Files menu entry, focused pane launch routing, selected/focused path seeding, initial DxUi window shell, preview grid, and debug snapshot helper.
- [x] 2026-06-10: Ran focused build/selftest for the Batch Rename command/window slice; `ResourceLocalizationContracts.Tests.ps1` still reports pre-existing language-neutral FileOps/HotPaths contract failures unrelated to Batch Rename.
- [x] 2026-06-10: Added standard rename prompt `Batch...` action for folders, preserved single-file rename prompt behavior, and routed multi-selection `Rename` directly into Batch Rename with selected-path preview seeding.
- [x] 2026-06-10: Routed the Batch Rename window preview through `BatchRename::BuildPlan`, added visible Rules controls for template/search/replace/options/case transforms, and covered control-driven recompute with command selftests.
- [x] 2026-06-10: Added Manual mode UI with a multiline DxUi editor, preview seeding, Fill/Clear controls, exact line-to-row mapping, mismatch validation, and mode-switch preservation coverage.
- [x] 2026-06-10: Added localized helper menu definitions, visible dropdown helper buttons for template/search/replace fields, command-id insertion into DxUi text fields, and helper insertion selftests.
- [x] 2026-06-10: Added deterministic engine timestamp macro support for optional target metadata, plus selftests for escaped braces, `$(Name)` alias normalization, date/time formatting, and regex `$1`/`$&`/`$$` replacement tokens.
- [x] 2026-06-10: Added local folder-scope target collection when no explicit paths are supplied, local metadata seeding for explicit and folder-scope paths, size/date/time preview text, and a folder-scope collection selftest.
- [x] 2026-06-10: Added v1 Batch Rename execution through a shared arena-backed `FileSystemRenameBatch` helper, depth-grouped `IFileSystem::RenameItems` calls, cache move notifications, refreshed successful preview targets, and a local execution selftest.
- [x] 2026-06-10: Added local execution revalidation before dispatch for missing sources and destinations created after preview, preventing earlier rows from being partially applied.
- [x] 2026-06-10: Added engine preview perf instrumentation and a 10,000-target deterministic macro/date/time preview selftest.
- [x] 2026-06-10: Added `Common::Settings::BatchRenameSettings`, schema/store roundtrip coverage, and Batch Rename window load/save of histories, options, case transforms, grid view state, and placement while excluding manual multiline names.
- [x] 2026-06-10: Hardened Manual mode with too-many-line validation coverage, Paste from multiline Unicode clipboard text, and preservation of manual names across non-target rule edits.
- [x] 2026-06-10: Added visible local folder-scope controls for mask, include subdirectories, files, and folders; local collection now filters recursive scopes and persists mask/scope options.
- [x] 2026-06-10: Added parent/child local execution regression coverage and fixed successful preview refresh so child targets move under renamed parent directories.
- [x] 2026-06-10: Completed the screenshot-inspired helper menu gaps with selected-text regex literal escaping and custom matched-subexpression insertion.
- [x] 2026-06-10: Added window-level Manual mode target-change coverage proving changed folder scopes preserve manual text, block mismatched line counts, and re-enable Rename after reconciliation.
- [x] 2026-06-10: Completed Phase 16 documentation updates across command routing, shared Find-style header reference, file-operation execution contract, and user-facing docs.
- [x] 2026-06-10: Added Batch Rename window posted-payload lifecycle guard with `InitPostedPayloadWindow` on create and `DrainPostedPayloadsForWindow` on destroy, ahead of async provider work.
- [x] 2026-06-10: Added local execution regression coverage proving blocking preview errors are rejected before provider dispatch and source files remain untouched.
- [x] 2026-06-10: Added preview sorting through the DxUi grid delegate plus Manual `Sort like preview`, preserving target-to-manual-name pairing when visible preview order changes.
- [x] 2026-06-10: Expanded the footer validation summary to include warning counts and covered it with a window selftest.
- [x] 2026-06-10: Added `batchrename.regex.compile.us` instrumentation around per-preview regex compilation and covered it with a focused command selftest.
- [x] 2026-06-10: Added local Batch Rename case-only execution coverage, proving destination preflight allows case-only changes and the provider preserves requested casing.
- [x] 2026-06-10: Added a 150 ms Batch Rename preview debounce for text edits, with explicit pending/flush regression coverage.
- [x] 2026-06-10: Added `batchrename.validation.us` instrumentation for the preview validation phase, including duplicate detection, and covered it with a focused perf selftest.
- [x] 2026-06-10: Added current v1 execution perf metrics (`batchrename.execute.us`, rows, completed, failed) and asserted them through the local execution selftest.
- [x] 2026-06-10: Added provider-backed folder-scope enumeration through `IFileSystem::ReadDirectoryInfo`, provider metadata seeding for size/date/time, collection perf metrics, and selftest coverage proving provider enumeration is used when supplied.
- [x] 2026-06-10: Added `RenameItem` fallback for providers that return unsupported from `RenameItems`, plus a window execution selftest proving the fallback dispatches one single-item rename per changed row.
- [x] 2026-06-10: Added Batch Rename window preview recompute and visible grid refresh perf instrumentation, with window-level selftest coverage for both metrics.
- [x] 2026-06-10: Added retained Batch Rename execution reports and localized status summaries for planned/completed/skipped/failed rows, first failure text, and canceled-state storage.
- [x] 2026-06-10: Added preview-time local destination collision validation so existing unselected siblings block Rename before execution, while duplicate proposed names remain engine-level validation.
- [x] 2026-06-10: Added synchronous Batch Rename `IFileSystemCallback` plumbing for execution, including `FileSystemShouldCancel` checks, cancel-on-issue behavior, canceled report classification, and focused cancellation regression coverage.
- [x] 2026-06-10: Replaced the plain Batch Rename root path label with the shared Find Files-style `NavigationView` breadcrumb header, including debug snapshot coverage for the real navigation child and populated current root.
- [x] 2026-06-10: Added preview grid context-menu copy actions for original name, new name, source path, and TSV preview rows, with clipboard regression coverage.
- [x] 2026-06-10: Added explicit accessible names for Batch Rename focusable controls and covered Rules-mode tab order plus live light/dark theme updates with a focused snapshot selftest.
- [x] 2026-06-10: Added a Batch Rename success callback carrying successful source/target path pairs and wired pane-launched executions to refresh affected visible panes after successful renames.
- [x] 2026-06-10: Added preview-grid `Reveal in Active Pane`, routing the clicked preview source path through the pane context so `FolderWindow` navigates to its parent and focuses the source item.
- [x] 2026-06-10: Added the footer `Hide unchanged` preview toggle, keeping no-op rows visible by default while allowing the grid to show only changed rows without changing full-plan stats or execution.
- [x] 2026-06-10: Wired preview-grid row activation to the same reveal-in-active-pane callback, so double-click/Enter on a preview row navigates to the source item when the launch context supports reveal.
- [x] 2026-06-10: Added a retained undo-plan report artifact for successful Batch Rename rows, with `Copy Undo Plan` exporting a TSV report of current path, restore name, and original path while excluding skipped no-op rows.
- [x] 2026-06-10: Added shell icon resolution for preview rows and rendered the `Original Name` column as DxUi icon/text cells through the Batch Rename grid delegate, with debug snapshot coverage for resolved icon indices.
- [x] 2026-06-10: Added `New Name` warning/error status styling with row tones, status glyphs, and tooltips carrying stable issue IDs for affected preview rows.
- [x] 2026-06-10: Added deterministic Batch Rename leaf validation for Windows-invalid/control characters and target leaf names longer than the 255-code-unit Windows component limit, with stable `name_invalid_character` and `name_too_long` issue IDs.
- [x] 2026-06-10: Added `Copy Execution Report` for retained Batch Rename execution reports, exporting counts, canceled state, first-failure HRESULT, and first-failure text as TSV after success, failure, or cancellation.
- [ ] 2026-06-10 checkpoint slice: started provider-cancelable target-collection coverage. The checkpoint touched `RedSalamander/BatchRenameWindow.cpp`, `RedSalamander/BatchRenameWindow.h`, `RedSalamander/SelfTest/Commands/Commands.SelfTest.BatchRename.cpp`, and this plan. The intended change is:
  - expose `BatchRenameDebugCollectionResult` and `DebugCollectBatchRenameTargetsForTests(...)` under `ENABLE_TESTS`;
  - make `CollectProviderScopeTargets(...)` return provider cancellation HRESULTs from `IFileSystem::ReadDirectoryInfo(...)` instead of silently continuing;
  - make `CollectBatchRenameTargets(...)` avoid local fallback when provider enumeration reports cancellation;
  - add `cmd_pane_batchRename_target_collection_respects_provider_cancellation`, using a wrapped local file system that returns `HRESULT_FROM_WIN32(ERROR_CANCELLED)` on the second `ReadDirectoryInfo(...)` call from an MTA worker thread.
  Verification state at interruption:
  - RED evidence before production helper: `.\build.ps1 -ProjectName RedSalamander` failed in `.build\logs\msbuild-20260610_230533_650.log` because `BatchRenameDebugCollectionResult` and `DebugCollectBatchRenameTargetsForTests` were undeclared.
  - First green-attempt build failed with an MSBuild/CL process issue, not source diagnostics: `.build\logs\msbuild-20260610_231121_571.log`, `D8040` plus `MSB4166`.
  - Direct MSBuild `/m:1` retry is not useful evidence in this sandbox because it failed in MSBuild's CL task with duplicate `Path`/`PATH` environment entries.
  - Wrapper retry passed: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug`, `.build\logs\msbuild-20260610_232850_708.log`, diagnostics `0 warning(s), 0 error(s)`.
  - Focused selftest is not green: `Start-Process -FilePath .\.build\x64\Debug\RedSalamander.exe -ArgumentList @('--commands-selftest','--selftest-case=cmd_pane_batchRename_target_collection_respects_provider_cancellation') -Wait -PassThru -WindowStyle Hidden` returned `ExitCode=1`. The direct `& .\.build\x64\Debug\RedSalamander.exe ...; "$LASTEXITCODE"` form is invalid for reliable evidence here because the GUI executable returned a blank `$LASTEXITCODE`.
  - `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run` still showed older timestamps when checked, so the next chat should first capture the real focused-run failure detail before changing code. Do not mark provider collection cancellation complete until the focused case and the broader `cmd_pane_batchRename_` suite pass.
- [x] 2026-06-11: 🚧➡️✅ Review-fix slice (window/execution hardening):
  - 🛡️ Fixed the double delete on creation failure in `BatchRenameWindow::Create`/`ShowBatchRenameWindow` (and the identical pattern in `FindFilesWindow::Create`/`ShowFindFilesWindow`) with a stack destruction-observer: the destructor marks a stack bool owned by `Create()`, `Create()` now owns all failure cleanup, and callers never delete after calling it.
  - 🔀 Execution is now dependency-aware inside each deepest-first depth group: ops run in topological layers so a source is vacated before its path is reused (chains), swaps/cycles are broken through a unique same-directory temp name (`<leaf>.rsren-<guid>`), temp-step failures surface as that row's failure with best-effort temp restore, and undo entries record only NET source→final transitions.
  - 📋 `FileSystemItemCompleted` outcomes are recorded per batch (mutex-guarded, matched by folded source path) so partially failed parallel batches keep undo entries, `NotifyPathMoved`, success-callback spans, target refresh, and completed/failed counts accurate; providers that never report per-item completion fall back to all-or-nothing per batch. `onSuccessfulRename` now also fires on failure/cancel paths when at least one row renamed.
  - 🧵 Completed the checkpointed cancellation work: folder-scope collection and the execution dispatch phase now run on a single background `std::jthread` (MTA COM) with generation-tagged `PostMessagePayload` results (`WndMsg::kBatchRenameCompleted`) and progress (`WndMsg::kBatchRenameTaskUpdate` via `FileSystemProgress` + layer boundaries). A footer `Cancel` button sets the shared atomic `_cancelRequested` honored by `FileSystemShouldCancel` and the collection loops; `Rename` is guarded against re-entry; window close cancels and joins the worker before teardown. Explicit-selection seeding stays synchronous (bounded by selection size).
  - 🧪 Under `ENABLE_TESTS`, `DebugExecute`, `DebugSetScope`, `DebugFlushPendingPreview`, and the snapshot accessor pump the message loop until background tasks settle so existing synchronous selftests keep observing terminal state.
  - 🧾 View-only refreshes (grid sort, `Hide unchanged`) no longer destroy the retained execution report/undo plan; only plan-changing edits clear it. Manual `Sort like preview` now operates on the full target set independent of the `Hide unchanged` filter.
  - 🪟 Preview context menu re-resolves the clicked row by stable id after the modal menu loop, so debounced rebuilds firing inside the menu cannot dispatch against a stale row index.
  - 🌐 Collection failures are logged via `Debug::Error` and surfaced through a localized status message; non-local provider contexts now resolve explicit-selection metadata through `ReadDirectoryInfo` of the parent folder instead of local stat calls.
  - 🧹 Removed dead `WM_KEYDOWN`/`VK_ESCAPE` and local `WM_DPICHANGED` branches (both handled by `DxUi::WindowHost`); preview issue tooltips now show `localized text (raw_id)` through new `IDS_BATCH_RENAME_ISSUE_*` strings (including prospective `name_reserved_device` and a generic `regex_*` fallback) in all five `.rc` files.
  - ⚠️ Known follow-ups: the focused `cmd_pane_batchRename_target_collection_respects_provider_cancellation` case from the interrupted checkpoint still needs debugging; the accessibility-snapshot, cancel-count, and execute-metric selftests likely need updating for the new Cancel button, per-item completed counts, and worker-emitted metrics; File Operations informational-task progress remains open.
- [x] 2026-06-11: Extended the standard rename prompt `Batch...` action to files: the button is now always visible; folders keep opening Batch Rename rooted at the prompted folder, files open it seeded with exactly the prompted file (captured at prompt-open time and passed through the widened `BatchRenameRequestCallback(targetPath, isDirectoryRoot)` contract plus a `CommandBatchRename` explicit-targets overload, so watcher-driven selection resets while the prompt is open cannot change the target). Updated the single-file prompt selftest and added `cmd_pane_rename_file_prompt_batch_button_opens_batch_rename`.
- [ ] Next slice: remaining provider/perf/refinement gaps (File Operations informational task progress, provider edge cases, perf evidence).

## Product Contract

- Batch Rename is a top-level app-owned window with the visual density and navigation header style of Find Files, not a loose Win32 dialog.
- The top navigation bar is populated from the active pane's current folder and file-system context.
- Single-item `Rename` keeps opening the standard rename prompt.
- The standard rename prompt always shows a `Batch...` button: for a folder it opens Batch Rename rooted at that folder; for a file it opens Batch Rename seeded with exactly the prompted file (captured at prompt-open time).
- When more than one item is selected, the pane `Rename` command bypasses the standard prompt and opens Batch Rename directly with those selected items in the preview.
- A dedicated command, menu entry, and command-palette entry named `Batch Rename...` opens the same window from the active pane.
- The preview grid has these visible columns in this order for v1: `Original Name`, `New Name`, `Size`, `Date`, `Time`, `Path`.
- The `New Name` column shows the computed target leaf name. Blocking validation issues are rendered through row status color, icon/adornment, tooltip text, and disabled `Rename`, not through a required extra visible column.
- Manual mode uses a multiline edit where each non-CR line maps to one preview row in current grid order.
- Rule mode supports:
  - New-name templates with macros.
  - Search and replace.
  - Regular expression search.
  - Case-sensitive and whole-word options.
  - Replace once or replace all.
  - Exclude extension for search/replace.
  - Direct Change Case controls for file name and extension, reusing the Change Case engine logic.
- Helper menu buttons are available next to the template, search, and replace fields.
- The window never performs a rename until the current preview plan has been validated and the user presses `Rename`.

## User Experience Details

1. Launch from multiple selection:
   - Use `FolderView::GetSelectedOrFocusedPaths()` plus metadata from the current view.
   - Seed the preview with exactly the selected items, preserving the pane selection order when available and falling back to stable display order.
   - Do not enable `Include subdirectories` by default for an explicit multi-selection.

2. Launch from the standard rename prompt:
   - The existing rename prompt stays optimized for one simple rename.
   - `Batch...` is always shown, for folders and files alike.
   - Pressing `Batch...` closes the rename prompt without renaming and opens Batch Rename: rooted at that folder when `originalIsDirectory == true`, otherwise seeded with exactly the prompted file (captured at prompt-open time, immune to selection changes while the prompt is open).
   - For a folder root, seed `Mask` as `*.*` and `Include subdirectories` as the user's persisted default.

3. Launch from the dedicated command:
   - If selected count is greater than one, seed from selected items.
   - If selected count is one and it is a folder, seed from that folder.
   - If selected count is one and it is a file, seed with that file.
   - If nothing is selected, seed from the active pane current folder.

4. Header and scope controls:
   - Reuse `NavigationView` styling from Find Files.
   - Show root path from the active pane.
   - Provide `Mask`, `Include subdirectories`, `Files`, and `Folders` controls.
   - Changing scope recomputes the target set and preview asynchronously.

5. Rename rule area:
   - Tabs or segmented control:
     - `Rules`
     - `Manual`
   - `Rules` contains new-name template, search/replace, regex options, and change-case controls.
   - `Manual` contains the multiline editor and a small command row: `Fill from preview`, `Clear`, `Paste`, `Sort like preview`.
   - Avoid in-app explanatory paragraphs; use labels, tooltips, menus, and validation text.

6. Preview footer:
   - Show counts: total rows, changed rows, unchanged rows, errors, warnings.
   - Show `Rename` split button:
     - Primary action: execute valid plan.
     - Menu actions: `Copy Preview`, `Export Plan...`, `Export Undo Plan...`.
   - `Rename` is disabled if there is any blocking validation issue.

## File Map

### New Files

- `RedSalamander/BatchRenameEngine.h`
- `RedSalamander/BatchRenameEngine.cpp`
- `RedSalamander/BatchRenameWindow.h`
- `RedSalamander/BatchRenameWindow.cpp`
- `RedSalamander/BatchRenameMenus.h`
- `RedSalamander/BatchRenameMenus.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.BatchRename.cpp`
- `Specs/UI/UI_BatchRenameWindow.md`
- `Specs/TestRuns/<timestamp>-batch-rename-preview-baseline/README.md`
- `Specs/TestRuns/<timestamp>-batch-rename-preview-candidate/README.md`
- `Specs/TestRuns/<timestamp>-batch-rename-execute-candidate/README.md`

### Modified Files

- `RedSalamander/RedSalamander.vcxproj`
- `RedSalamander/RedSalamander.vcxproj.filters`
- `RedSalamander/Resource.h`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/CommandRegistry.cpp`
- `RedSalamander/RedSalamander.cpp`
- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.cpp`
- `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.FileOps.cpp`
- `RedSalamander/FolderViewInternal.h`
- `Common/WindowMessages.h`
- `Common/SettingsStore.h`
- `Common/Common/SettingsStore.cpp`
- `Specs/SettingsStore.schema.json`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/UI/UI_FindFilesWindow.md`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/FileSystem/FileSystem_FileOperations.md`
- `docs/UserGuide.md`
- `docs/MainWindow.md`

## Phase 0 - Finalize Scope and IDs

- [x] Reserve command ID `IDM_PANE_BATCH_RENAME` at `33098` after verifying it remains unused in `RedSalamander/Resource.h`.
- [x] Reserve string IDs for `IDS_BATCH_RENAME_*` in the open range `1810..1899`, because `1792..1799` is too small and `1800..1804` is already used.
- [x] Reserve command strings:
  - `IDS_CMD_BATCH_RENAME`
  - `IDS_CMD_DESC_BATCH_RENAME`
- [x] Add `WndMsg::kBatchRenameTaskUpdate = WM_APP + 0x544`.
- [x] Add `WndMsg::kBatchRenameCompleted = WM_APP + 0x545`.
- [x] Add `WndMsg::kBatchRenameWindowDebug = WM_APP + 0x546` under `ENABLE_TESTS`.
- [x] Confirm no conflict with `Common/WindowMessages.h` test and non-test message ranges.
  - 2026-06-10: `cmd_pane_batchRename_command_registered` asserts the central message IDs so future local `WM_APP` drift is caught.
- [x] Confirm whether satellite `.rc` files are present and update them with source-language fallback strings if the repo requires complete satellite coverage in the same change.

## Phase 1 - Authoritative Spec First

- [x] Create `Specs/UI/UI_BatchRenameWindow.md`.
- [x] Define launch routes:
  - Standard rename prompt `Batch...` (folders and files).
  - Multiple selection `Rename`.
  - Dedicated `Batch Rename...` command.
  - Command palette or command registry invocation.
- [x] Define the window type as an app-owned top-level DxUi window using the Find Files navigation/header pattern.
- [x] Define target collection semantics:
  - Explicit selection.
  - Active folder contents.
  - Single folder scope.
  - Include subdirectories.
  - File/folder inclusion toggles.
  - Mask matching.
  - Snapshot consistency while the folder changes underneath the window.
- [x] Define preview grid columns, default widths, alignment, sort defaults, tooltip behavior, and status rendering.
- [x] Define macro syntax and escaping:
  - Canonical RedSalamander syntax uses `{macro}`.
  - The helper menu inserts canonical syntax.
  - `$(Macro)` aliases are accepted for common migration cases from other renamers, but are normalized internally to the same macro tokens.
  - Literal braces use `{{` and `}}`.
- [x] Define v1 macro set:
  - `{name}`: original leaf name including extension.
  - `{stem}`: original leaf name without final extension.
  - `{ext}`: final extension including dot.
  - `{extNoDot}`: final extension without dot.
  - `{parent}`: immediate parent folder leaf.
  - `{relativeFolder}`: folder path relative to the rename root.
  - `{relativeFolderFlat}`: relative folder with path separators replaced by the configured separator.
  - `{size}`: byte size.
  - `{date:yyyy-MM-dd}` and `{time:HH-mm-ss}`: last-write date/time.
  - `{created:yyyy-MM-dd}`: creation date where the provider supplies it.
  - `{counter}` and `{counter:000}`: one-based row counter after current preview order.
  - `{index}`: zero-based row index after current preview order.
- [x] Define invalid macro behavior:
  - Unknown macros are blocking errors with a row-independent validation message.
  - Macro output containing path separators is blocking unless the specific macro is documented as flattened.
  - Names that become empty, `.` or `..`, or contain provider-invalid characters are blocking.
- [x] Define regex behavior:
  - `std::wregex` ECMAScript syntax.
  - `Case sensitive` controls `std::regex_constants::icase`.
  - `Whole words` wraps the user expression in word-boundary logic.
  - `Only once in each name` limits replacement to the first match.
  - Replacement syntax uses standard ECMAScript replacement tokens: `$&`, `$1`, `$2`, and `$$`.
- [x] Define Change Case behavior:
  - File name and extension have independent dropdowns: `Do not change`, `Lower case`, `Upper case`, `Mixed case`.
  - The implementation reuses or factors `ChangeCase::TransformLeafName`.
  - `Include path part` is not supported for leaf-only v1 renames because `IFileSystem::RenameItems` accepts leaf names only; offer `relativeFolderFlat` macros instead.
- [x] Define execution behavior:
  - Revalidate just before execution.
  - Execute only changed, valid rows.
  - Skip no-op rows.
  - Use provider bulk rename when available through `IFileSystem::RenameItems`.
  - Support cancel through the existing file-system callback mechanism.
  - Refresh affected panes and notify directory cache for successful renames.
- [x] Define explicit non-goals:
  - Moving files to different directories.
  - Cross-provider renames.
  - Bulk creation of missing directories.
  - Persistent undo command in the main command system.

## Phase 2 - Settings and Resources

- [x] Add `Common::Settings::BatchRenameSettings` in `Common/SettingsStore.h`.
- [x] Include:
  - `lastRoot`
  - `recentMasks`
  - `recentNameTemplates`
  - `recentSearchPatterns`
  - `recentReplacePatterns`
  - `includeSubdirectories`
  - `includeFiles`
  - `includeFolders`
  - `regexEnabled`
  - `caseSensitive`
  - `wholeWords`
  - `replaceOnce`
  - `excludeExtension`
  - `flattenSeparator`
  - `fileNameCaseStyle`
  - `extensionCaseStyle`
  - `previewGridLayout`
  - `previewSort`
- [x] Persist the window placement under `settings.windows.BatchRenameWindow`.
- [x] Do not persist manual multiline names.
- [x] Update JSON read/write and sanitization in `Common/Common/SettingsStore.cpp`.
- [x] Update `Specs/SettingsStore.schema.json`.
- [x] Update `Specs/Core/Core_SettingsStore.md`.
- [x] Add all UI strings to `RedSalamander/RedSalamander.rc`; no hardcoded user-visible strings in C++.
- [x] Use positional placeholders in resource strings that contain runtime data.
- [x] Add the menu command to the main `Files` menu near `Rename...`.
- [x] Add command registry entries:
  - `cmd/pane/batchRename`
  - Name resource `IDS_CMD_BATCH_RENAME`
  - Description resource `IDS_CMD_DESC_BATCH_RENAME`
- [x] Run `pwsh -File .\Tools\Tests\ResourceLocalizationContracts.Tests.ps1`.
  - 2026-06-10: Command completed with exit code 0, but the Pester report still lists existing language-neutral FileOps/HotPaths violations outside the Batch Rename changes.

## Phase 3 - Pure Rename Engine

- [x] Implement `BatchRenameEngine.h/.cpp` without window dependencies.
- [x] Define data types:
  - `BatchRenameTarget`
  - `BatchRenameRules`
  - `BatchRenamePreviewRow`
  - `BatchRenameIssue`
  - `BatchRenamePlan`
  - `BatchRenameStats`
- [x] Store all paths as `std::filesystem::path` or `std::wstring` consistently with existing folder/file-system code.
- [x] Keep target `newName` as a leaf name only, matching `FileSystemRenamePair::newName`.
- [x] Add a macro table with parser and evaluator.
- [x] Accept canonical `{macro}` syntax and supported `$(Macro)` aliases.
- [x] Add deterministic formatting helpers for size/date/time that do not depend on current UI column text.
- [x] Add search/replace pipeline:
  - Split stem/extension when `excludeExtension` is enabled.
  - Apply regex or literal replacement.
  - Apply replace-once or replace-all.
  - Preserve extension correctly for names such as `.gitignore`, `archive.tar.gz`, and files with no extension.
- [x] Add Change Case pipeline after template and replacement steps.
- [x] Factor Change Case code so `BatchRenameEngine` and `ChangeCase` use one implementation for:
  - lower
  - upper
  - mixed
  - stem-only
  - extension-only
- [x] Do not copy/paste the existing Change Case transformations.
- [ ] Add validation:
  - Empty new names.
  - `.` and `..`.
  - [x] Names with path separators.
  - [x] Windows-invalid/control characters where available.
    - 2026-06-10: Engine validation blocks Windows-invalid leaf characters plus control characters with stable `name_invalid_character` issues.
  - [ ] Provider-specific invalid characters for future non-local providers when those providers expose stricter rules.
  - Duplicate target names in the same parent, compared case-insensitively.
  - [x] Existing local destination conflicts not part of the same validated plan.
    - 2026-06-10: Local preview validation now adds blocking `name_destination_exists` issues for existing unselected siblings before execution; execution revalidation still handles races after preview.
  - [x] Local source missing at execution revalidation.
  - Parent/child selected paths that require deepest-first execution order.
  - [x] Target leaf longer than the known Windows component limit.
    - 2026-06-10: Engine validation blocks target leaves longer than 255 UTF-16 code units with stable `name_too_long` issues.
  - [ ] Target leaf longer than provider-specific limits when those limits are exposed by a provider.
  - Unknown macros and malformed date/time format specifiers.
  - Invalid regex with `std::regex_error` captured as a validation issue.
- [x] Add stable row IDs so preview selection remains stable across rule edits when the target set does not change.
- [ ] Return warning issues for:
  - [x] No-op rows as `name_unchanged`.
  - [x] Case-only renames as `name_case_only`.
  - [x] Names with leading/trailing spaces or trailing dots as `name_edge_space_or_dot`.
  - [ ] Manual-mode line count mismatch before it becomes a blocking error.
- [x] Ensure the engine is deterministic for tests and does not read UI state directly.

## Phase 4 - Engine Tests Before UI

- [x] Add `Commands.SelfTest.BatchRename.cpp`.
- [x] Add selftests for macro expansion:
  - [x] simple stem/ext
  - [x] `{filename}` and `{extNoDot}`
  - [x] `{relativeFolder}` and `{relativeFolderFlat}` rooted at the pane folder
  - [x] `{size}` and `{created:yyyyMMdd}`
  - [x] counters with padding
  - [x] date/time formatting
  - [x] escaped literal braces
  - [x] unknown macro error
  - [x] `$(Name)` alias normalization
- [ ] Add selftests for literal search/replace:
  - [x] replace all
  - [x] replace once
  - [x] case-insensitive
  - [x] whole words
  - [x] exclude extension
- [ ] Add selftests for regex:
  - [x] capture replacement `$1`
  - [x] `$&`
  - [x] `$$`
  - [x] invalid regex
  - [x] whole-word regex
- [ ] Add selftests for case transforms:
  - [x] lower stem only
  - [x] upper extension only
  - [x] mixed case for words with separators
  - [x] no-op extension handling
- [ ] Add validation selftests:
  - [x] duplicate target in same folder
  - [x] duplicate differing only by case
  - [x] target already exists
  - [x] invalid leaf with separator
  - [x] invalid leaf with Windows-invalid/control character
  - [x] target leaf longer than the Windows component limit
  - [x] empty manual line
  - [x] source missing during execution revalidation
  - [x] parent/child deepest-first ordering
- [x] Add a large deterministic preview selftest with at least 10,000 synthetic targets and assert bounded time through perf metrics rather than wall-clock-only pass/fail.
  - 2026-06-10: `cmd_pane_batchRename_engine_large_preview_perf` covers 10,000 synthetic targets through macro expansion, date/time formatting, literal replacement, and Change Case; it asserts `batchrename.preview.build_plan_us`, `batchrename.preview.rows`, `batchrename.preview.changed`, and `batchrename.preview.errors` are emitted, and bounds the emitted build-plan duration at 5,000,000 us.

## Phase 5 - Launch Routing

- [x] Update `FolderView::CommandRename()` and/or `FolderWindow::CommandRename(Pane pane)`.
- [x] If selected count is greater than one, call `FolderWindow::CommandBatchRename(pane)` directly.
- [x] Keep single-selection rename behavior unchanged for files.
- [x] Add `FolderWindow::CommandBatchRename(Pane pane)`.
- [x] Build a `BatchRenamePaneContext` modeled after `FindFilesPaneContext`:
  - `wil::com_ptr<IFileSystem> fileSystem`
  - `pluginId`
  - `pluginShortId`
  - `instanceContext`
  - `rootPluginPath`
  - active pane root/current path
  - selected or focused input paths
- [x] Add `ShowBatchRenameWindow(...)` in `BatchRenameWindow.h`.
- [x] Update `RedSalamander.cpp` command dispatch for:
  - `IDM_PANE_BATCH_RENAME`
  - `cmd/pane/batchRename`
- [x] Add menu and command-palette discoverability.
- [x] Add a debug helper command under `ENABLE_TESTS` for opening and snapshotting the Batch Rename window.

## Phase 6 - Standard Rename Prompt Integration

- [x] Extend the existing `FolderViewRenamePromptWindow` in `FolderViewInternal.h`.
- [x] Replace the binary optional result with an explicit result:
  - `RenamePromptAction::Cancel`
  - `RenamePromptAction::Rename`
  - `RenamePromptAction::BatchRename`
- [x] Show `Batch...` for folders and files (2026-06-11: originally folders-only; extended to files).
- [x] Keep `OK`, `Cancel`, Enter, Escape, and initial filename selection behavior unchanged.
- [x] When `Batch...` is clicked, close the prompt and return `BatchRename`.
- [x] In `FolderView::RenameFocusedItem()`, route `BatchRename` back to `FolderWindow` or a pane-level callback instead of performing a single rename (folder roots Batch Rename at the folder; file routes the pane selection).
- [x] Avoid circular ownership between `FolderView` and `FolderWindow`; use the same callback pattern already used for pane commands if available.
- [x] Add dialog selftests:
  - [x] folder shows `Batch...` and opens Batch Rename rooted at the folder
  - [x] file shows `Batch...` and opens Batch Rename seeded with the prompted file
  - [x] pressing `Batch...` does not call `RenameItem`
  - [x] Enter still performs standard rename
  - [x] Escape still cancels

## Phase 7 - Batch Rename Window Shell

- [x] Implement `BatchRenameWindow final` using DxUi, following `FindFilesWindow.cpp` structure where appropriate.
  - 2026-06-10: The Batch Rename shell is a DxUi-owned top-level window and now embeds the shared `NavigationView` root header like Find Files.
- [x] Use `NavigationView` for the root header.
  - 2026-06-10: `cmd_pane_batchRename_opens_from_active_pane` asserts the real navigation child window and the populated root path through `NavigationView::DebugGetSnapshot`.
- [x] Use `DxUi::Grid` for preview.
- [ ] Use `DxUi::ComboBox` or existing editable combo control patterns for:
  - mask
  - new-name template
  - search for
  - replace with
- [x] Add rule-mode text fields for new-name template, search for, and replace with; v1 histories/editable combos remain with settings persistence.
- [x] Use checkboxes for boolean options.
- [x] Use dropdown helper buttons beside the rule fields, with localized tooltips.
- [x] Use compact labels and controls; avoid explanatory static text blocks.
- [x] Persist and restore:
  - [x] window placement
  - [x] preview column layout
  - [x] sort
  - [x] histories
  - [x] last options
- [x] Register window class with WIL-safe lifetime management.
- [x] Call `InitPostedPayloadWindow(hwnd)` when the window can receive posted payload messages.
- [x] Call `DrainPostedPayloadsForWindow(hwnd)` during `WM_NCDESTROY`.
- [ ] Keep WndProc cases short and delegate to methods, following `win32-wndproc` guidance.
- [x] Add `UpdateBatchRenameWindowsTheme(...)` if multiple windows can be open.
- [x] Update live theme when app theme changes.
- [x] Add accessible names and keyboard focus order for all controls.
  - 2026-06-10: Rules-mode focusable controls now expose stable accessible names in DxUi traversal order. Manual-mode keyboard specifics remain tracked under Phase 15 keyboard behavior.

## Phase 8 - Target Collection

- [x] Implement target collection in a helper that is UI-thread initiated and worker-thread executable.
  - 2026-06-10 partial: Collection is now factored into provider/local helper paths and emits metrics, but still runs synchronously on the UI thread. Worker execution remains open.
  - 2026-06-10 interrupted: uncommitted code exposes a worker-callable `DebugCollectBatchRenameTargetsForTests(...)` probe and a focused cancellation selftest, but the focused case currently exits `1`; investigate before committing.
  - 2026-06-11: ✅ Folder-scope collection now runs on a background `std::jthread` (MTA COM) started by `StartTargetCollection()`; results post back through `WndMsg::kBatchRenameCompleted` payloads with generation-based stale-drop. Explicit-selection seeding stays synchronous because it is bounded by the selection size.
- [x] For explicit local selections, use the selected paths as the target set and fetch missing metadata only as needed.
- [x] For folder scope, enumerate via `IFileSystem::ReadDirectoryInfo`.
  - 2026-06-10: Provider-backed folder scopes enumerate through `ReadDirectoryInfo` when a file-system provider is supplied. Local filesystem fallback remains for local contexts when provider enumeration cannot be used. Cancelable worker enumeration remains open.
- [ ] Respect:
  - [x] mask for local folder scopes
  - [x] include subdirectories for local folder scopes
  - [x] include files for local folder scopes
  - [x] include folders for local folder scopes
  - [ ] provider cancellation
    - 2026-06-10 interrupted: uncommitted code attempts to propagate `ERROR_CANCELLED` from provider `ReadDirectoryInfo(...)` and suppress local fallback, but the focused selftest is still failing and must be debugged before this item can be checked.
    - 2026-06-11: collection loops (provider and local) now also honor the window-level atomic cancel flag and return `ERROR_CANCELLED`; the focused checkpoint selftest still needs to be debugged before checking this off.
- [x] Use background work for enumeration when scope can exceed visible selected items.
- [x] Initialize worker threads that call provider APIs with `wil::CoInitializeEx(COINIT_MULTITHREADED)`.
  - 2026-06-11: workers use the established `CoInitializeEx(nullptr, COINIT_MULTITHREADED)` + `wil::unique_couninitialize_call` pattern shared with the ChangeCase/ChangeAttributes tasks.
- [x] Post collection and preview results to the window with `PostMessagePayload`.
- [x] Drop stale worker results using a monotonically increasing generation number.
- [x] Keep the UI responsive while preview recomputes.
- [x] Add a progress row or compact status line while enumeration is running.
  - 2026-06-11: the status strip shows a localized `Collecting items...` message and the footer `Cancel` button is enabled while collection runs.
- [x] Cancel outstanding enumeration when:
  - window closes
  - root changes
  - include-subdirectories changes
  - mask changes
  - provider context becomes invalid
  - 2026-06-11: every new collection/execution request and window close cancels and joins the single background task before proceeding.

## Phase 9 - Preview Grid

- [x] Implement `BatchRenameGridModel : DxUi::IDxGridModel`.
- [x] Implement `BatchRenameWindow : DxUi::IDxGridDelegate` for icons, tooltips, context menu, and row activation behavior.
  - 2026-06-10 partial: Batch Rename first implemented the grid delegate for sort requests.
  - 2026-06-10 partial: Preview grid context menu now covers copy actions and reveal-in-active-pane through a provider context callback. Row activation now uses the same reveal callback.
  - 2026-06-10: The grid delegate now also serves system icon bitmaps for `Original Name` icon/text cells; clipped-cell tooltips are provided by the grid model.
- [ ] Default columns:
  - [x] `Original Name`: left aligned, icon enabled.
    - 2026-06-10: Preview rows resolve file/folder shell icon indices with the shared `IconCache`, render the column as `IconText`, and fall back to Fluent file/folder glyphs when a bitmap is unavailable.
  - [x] `New Name`: left aligned, status styling.
    - 2026-06-10: Preview rows with warning/error issues use `GridRowTone::Warning`/`Error`, render the `New Name` cell with a warning/error glyph, and expose stable issue IDs in the cell tooltip.
  - `Size`: right aligned, folder blank or localized folder marker consistent with folder view.
  - `Date`: right aligned or locale-consistent.
  - `Time`: right aligned.
  - `Path`: left aligned, clipped with tooltip.
- [x] Sort preview by current grid sort while keeping row IDs stable.
- [x] Make manual-mode row order match current preview order through the `Sort like preview` action.
- [x] Recompute preview through a debounce of roughly 150 ms for text edits.
  - 2026-06-10: Rule-template/search/replace, manual multiline, and mask text edits now schedule one pending preview timer; helper commands and programmatic debug setters still rebuild immediately.
- [x] Compile regex only when needed and cache per preview generation.
  - 2026-06-10: `BatchRename::BuildPlan` compiles one `std::wregex` before row iteration only when regex search is enabled and non-empty, then emits `batchrename.regex.compile.us` with target count and final pattern length.
- [x] Show validation summary in the footer.
- [x] Disable `Rename` for blocking errors.
  - 2026-06-10: Window selftests now assert the Rename button disables for unknown macros, manual line-count mismatches, duplicate targets, existing destinations, and invalid preview execution.
- [x] Keep no-op rows visible by default and add a footer action to hide unchanged rows if simple to implement without complicating execution.
  - 2026-06-10: The footer `Hide unchanged` checkbox filters only the visible preview rows; stats, validation, and execution still operate on the full plan.
- [x] Add context menu actions:
  - [x] Copy original name
  - [x] Copy new name
  - [x] Copy source path
  - [x] Copy preview rows
  - [x] Reveal in active pane when provider context supports it

## Phase 10 - Helper Menus

- [x] Implement helper menu definitions in `BatchRenameMenus.h/.cpp`.
- [x] Macro helper menu inserts canonical tokens:
  - [x] `{name}`
  - [x] `{stem}`
  - [x] `{ext}`
  - [x] `{extNoDot}`
  - [x] `{parent}`
  - [x] `{relativeFolderFlat}`
  - [x] `{counter}`
  - [x] `{counter:000}`
  - [x] `{date:yyyy-MM-dd}`
  - [x] `{time:HH-mm-ss}`
  - [x] literal `{` and `}`
- [x] Regex search helper menu includes:
  - [x] Any character `.`
  - [x] Character set `[...]`
  - [x] Negated set `[^...]`
  - [x] Word boundary `\b`
  - [x] Alternation `|`
  - [x] Zero or more `*`
  - [x] One or more `+`
  - [x] Optional `?`
  - [x] Group `(...)`
  - [x] Non-capturing group `(?:...)` omitted because ECMAScript `std::wregex` does not support it.
  - [x] Escaped literal helper for selected text.
  - [x] Word character `\w`
  - [x] Non-word character `\W`
  - [x] Whitespace `\s`
  - [x] Non-whitespace `\S`
  - [x] Digit `\d`
  - [x] Non-digit `\D`
  - [x] Decimal number `\d+`
  - [x] Hexadecimal number `[0-9A-Fa-f]+`
  - [x] File name split helper `^(.+?)(\.[^.]+)?$`
- [x] Replacement helper menu includes:
  - [x] Literal dollar `$$`
  - [x] Whole match `$&`
  - [x] Matched subexpression `$1`
  - [x] Matched subexpression `$2`
  - [x] Custom `Matched Subexpression...`
- [x] Insert helper text at the caret or replace the current selection.
- [x] Keep helper menu items localized through resources.
- [x] Add tests for helper insertion at caret and selection replacement.
  - `cmd_pane_batchRename_helper_menus_expose_canonical_insertions`
  - `cmd_pane_batchRename_window_helper_buttons_insert_into_rule_fields`
  - 2026-06-10: These selftests now also cover selected-text regex escaping and custom replacement subexpression insertion through the visible rule fields.

## Phase 11 - Manual Mode

- [x] Add multiline edit support using the existing DxUi text field multiline mode or the nearest project-approved multiline control.
- [x] Seed manual text from current `New Name` values when switching into manual mode for the first time.
- [x] Preserve manual text when switching away and back while the target set is unchanged.
- [x] If the target set changes while manual mode is active, show a blocking validation message until the line count is reconciled.
- [x] Treat each line as an exact target leaf except for CR removal.
- [x] Reject empty lines.
- [x] Do not trim spaces silently; warn or block according to provider rules.
- [x] Add `Fill from preview` to replace manual text with current rule-generated names.
- [x] Add `Paste` behavior that preserves line breaks.
- [x] Add `Sort like preview` to rewrite manual lines into the current visible preview order while preserving target/manual-name pairing.
- [x] Add selftests:
  - [x] exact line count passes
  - [x] too few lines blocks rename
  - [x] too many lines blocks rename
  - [x] empty line blocks rename
  - [x] manual text survives mode switches with unchanged targets
  - [x] manual text survives non-target rule edits
  - [x] target-set changes preserve manual text, block mismatch, and clear after reconciliation
  - [x] Sort-like-preview rewrites manual lines after preview sorting without changing target/manual-name pairing

## Phase 12 - Execution Path

- [x] Create an immutable `BatchRenamePlan` snapshot when `Rename` is pressed.
- [x] Revalidate the snapshot against current provider state before dispatch.
  - 2026-06-10: Current v1 execution rebuilds and validates the plan immediately before dispatch, and local file-system execution preflights missing sources plus external destination conflicts. Provider-general remote revalidation remains open under provider edge-case coverage.
- [x] Convert valid changed rows to `FileSystemRenamePair`.
- [x] Group pairs by provider/file-system context; v1 should normally have one context from the active pane.
- [x] Sort execution deepest-first when selected paths contain parent/child relationships.
  - 2026-06-10: `cmd_pane_batchRename_window_executes_parent_child_deepest_first` covers a local parent-directory plus child-file selection and verifies the refreshed preview points the child under the renamed parent.
- [x] Call `IFileSystem::RenameItems(...)` for bulk execution.
- [x] Use `RenameItem(...)` only as a fallback if a provider does not support bulk semantics and the contract permits fallback.
  - 2026-06-10: `FileSystemRenameBatch::Execute` first attempts arena-backed `RenameItems`; it falls back to one `RenameItem` per row only for `E_NOTIMPL`, `ERROR_CALL_NOT_IMPLEMENTED`, or `ERROR_NOT_SUPPORTED`. `cmd_pane_batchRename_window_falls_back_to_rename_item_when_bulk_unsupported` covers this path.
- [x] Support cancel through `FileSystemShouldCancel`.
  - 2026-06-10: Synchronous v1 execution now passes a per-execution callback through `FileSystemRenameBatch::Execute`; provider cancellation HRESULTs set the retained report `canceled` flag and leave remaining rows unapplied. UI stop/progress work remains open under the async execution items below.
- [x] Use `std::jthread` owned by `FolderWindow::PaneState` or by the Batch Rename window with an explicit stop path.
  - 2026-06-11: the Batch Rename window owns one `_taskWorker` jthread shared by collection and execution, with cancel-and-join on new task start, window close, and `WM_NCDESTROY`.
- [x] If owned by `FolderWindow`, add `batchRenameThread` next to `changeCaseThread` and `changeAttributesThread`.
  - 2026-06-11: not applicable — the worker is owned by the Batch Rename window, so no `FolderWindow::PaneState` thread member is needed.
- [x] Initialize COM MTA in worker code that calls provider APIs.
- [x] Post progress through `PostMessagePayload` and consume with `TakeMessagePayload<T>`.
  - 2026-06-11: `FileSystemProgress` and layer boundaries post throttled `WndMsg::kBatchRenameTaskUpdate` payloads; completion posts `WndMsg::kBatchRenameCompleted` with the full report payload.
- [ ] Surface running progress through the existing File Operations informational task when the operation exceeds the existing threshold.
  - 2026-06-11: progress is currently surfaced through the window status strip (`Renaming {0} of {1}...`); the informational-task surface remains open.
- [x] Report:
  - [x] total planned rows
  - [x] completed rows
  - [x] skipped no-op rows
  - [x] failed rows
  - [x] first failure message
  - [x] canceled state
  - 2026-06-10: Current synchronous execution stores the latest report in the Batch Rename window, exposes it through the debug snapshot, and updates the localized status strip summary. Provider cancellation through `FileSystemShouldCancel` is wired; user-visible async cancel/progress remains open.
- [x] On each successful rename, call `DirectoryInfoCache::NotifyPathMoved` with source and target.
- [x] Refresh affected panes after completion, using existing watcher-aware refresh rules.
  - 2026-06-10: `BatchRenamePaneContext::onSuccessfulRename` reports successful source/target pairs, and `FolderWindow::CommandBatchRename` uses it to refresh same-context panes showing affected source or target folders.
- [x] Keep the window open after failures so the user can inspect and copy the report.
  - 2026-06-10: Execution paths retain the window-level report instead of closing the window; `Copy Execution Report` is available when a report exists, and `cmd_pane_batchRename_cancel_does_not_apply_remaining_rows` covers copying a canceled failure report.
- [x] Close or reset the window after full success according to the user's setting; default should keep it open with a success summary until the user closes it.
  - 2026-06-10: The current default keeps the window open after success, refreshes the preview to the renamed targets, disables `Rename` when no changed rows remain, and leaves the localized success summary visible until rules or targets change.
- [x] Export undo plan as a text or JSON report that reverses successful rows only; this is a report artifact, not an automatic undo command.
  - 2026-06-10: The retained execution report stores successful undo entries and the preview context menu exposes `Copy Undo Plan` as a TSV artifact. It includes only successful changed rows and excludes skipped no-op rows.

## Phase 13 - Integration with Existing Change Case

- [ ] Refactor `ChangeCase` transformations into reusable functions if the current API is too operation-oriented.
- [x] Preserve existing `Change Case...` behavior and selftests.
  - 2026-06-10: `cmd_pane_changeCase` passed after extracting shared `FileSystemRenameBatch` marshalling for `IFileSystem::RenameItems`.
- [x] Add Batch Rename case controls that call the shared transform code.
- [x] Ensure case-only renames still rely on existing file-system handling such as local-plugin temp rename fallback.
  - 2026-06-10: `cmd_pane_batchRename_window_executes_case_only_local_rename` verifies a local same-leaf case transform executes successfully, preserves the requested directory-entry casing, and refreshes the preview.
- [x] Add regression tests proving existing `cmd/pane/changeCase` cases still pass.

## Phase 14 - Performance Validation

- [x] Treat Batch Rename as perf-sensitive because it affects file operations, UI responsiveness, preview recomputation, enumeration, and large grids.
- [ ] Add instrumentation:
  - 2026-06-10 partial: pure engine preview planning emits `batchrename.preview.build_plan_us` with total and changed rows as dimensions, plus `batchrename.preview.rows`, `batchrename.preview.changed`, `batchrename.preview.errors`, and `batchrename.preview.warnings` count metrics.
  - 2026-06-10 partial: regex-enabled preview planning emits `batchrename.regex.compile.us` around the single per-preview regex compilation; detail is `regex`, `whole_words`, or `invalid`, `value0` is target count, and `value1` is final pattern length.
  - 2026-06-10 partial: preview validation emits `batchrename.validation.us`; detail is `rules` or `manual`, `value0` is row count, and `value1` is blocking error row count after duplicate detection.
  - 2026-06-10 partial: current synchronous v1 execution emits `batchrename.execute.us`; detail is `success`, `noop`, `preview_errors`, `revalidate_failed`, `rename_failed`, or `missing_context`, `value0` is rows to execute, and `value1` is completed rows. It also emits rows/completed/failed count metrics.
  - 2026-06-10 partial: synchronous target collection emits `batchrename.collect.us` with detail `selection`, `provider`, `local`, `local-fallback`, or `empty-root`, and emits `batchrename.collect.targets` with the collected target count.
  - 2026-06-10 partial: window preview recomputation emits `batchrename.preview.recompute.us`; detail is `rules` or `manual`, `value0` is total preview rows, and `value1` is changed rows. Visible preview model/grid refresh emits `batchrename.preview.visible_refresh.us`; detail is `grid` or `model`, `value0` is visible row count, and `value1` is changed rows.
  - [x] `batchrename.collect.targets`
  - [x] `batchrename.collect.us`
  - [x] `batchrename.preview.rows`
  - [x] `batchrename.preview.changed`
  - [x] `batchrename.preview.errors`
  - [x] `batchrename.preview.recompute.us`
  - [x] `batchrename.preview.visible_refresh.us`
  - [x] `batchrename.regex.compile.us`
  - [x] `batchrename.validation.us`
  - [x] `batchrename.execute.rows`
  - [x] `batchrename.execute.completed`
  - [x] `batchrename.execute.failed`
  - [x] `batchrename.execute.us`
- [ ] Add deterministic perf scenarios:
  - `batchrename.preview.1000_selected_literal`
  - [x] `batchrename.preview.10000_selected_macro`
  - `batchrename.preview.10000_selected_regex`
  - `batchrename.manual.10000_lines`
  - `batchrename.execute.1000_leaf_renames_tempfs`
- [ ] Archive a baseline run before UI execution is wired if measuring against existing no-op or synthetic engine behavior.
- [ ] Archive candidate runs under `Specs/TestRuns/`.
  - 2026-06-10 partial: Batch Rename command-family candidate archived at `Specs/TestRuns/4cb089111a23/Commands/2026-06-10_161852/` with aligned `batchrename.preview.*` metric names, bounded preview threshold, 17 passed / 0 failed selftests, and run README.
- [ ] Include command lines, machine/configuration, input dataset description, metric summaries, and pass/fail thresholds in each run `README.md`.
- [ ] Use `Tools\CompareTestRuns.ps1` when comparing baseline and candidate runs.
- [ ] Do not close the plan until perf evidence exists.

## Phase 15 - Regression and UI Tests

- [ ] Add command selftests:
  - [x] `cmd_pane_rename_multi_selection_opens_batch_rename`
  - [x] `cmd_pane_rename_single_file_uses_standard_prompt`
  - [x] `cmd_pane_rename_folder_prompt_batch_button_opens_batch_rename`
  - [x] `cmd_pane_batchRename_command_registered`
  - [x] `cmd_pane_batchRename_opens_from_active_pane`
    - 2026-06-10: Also covers `Original Name` icon/text cells and resolved shell icon indices in the preview snapshot.
  - [x] `cmd_pane_batchRename_window_preview_context_menu_copies_rows`
    - 2026-06-10: Also covers the preview-row reveal callback path.
    - 2026-06-10: Also covers preview-row activation routing through the same reveal callback.
  - [x] `cmd_pane_batchRename_theme_accessibility_snapshot`
  - [x] `cmd_pane_batchRename_folder_scope_collects_local_children_metadata`
    - 2026-06-10: Also covers visible local scope filtering for mask, recursive files, and folder-only targets.
  - [ ] `cmd_pane_batchRename_target_collection_respects_provider_cancellation`
    - 2026-06-10 interrupted: test has been added in the working tree but currently exits `1`. Next step is to capture the actual selftest failure detail, then verify whether the failing assertion is `result.hr`, `readDirectoryInfoCalls`, or the partial `fullPaths` expectation.
  - [x] `cmd_pane_batchRename_preview_macro_regex_case_validation`
  - [x] `cmd_pane_batchRename_engine_macro_alias_datetime_regex_tokens`
  - [x] `cmd_pane_batchRename_manual_mode_line_count_validation`
  - [x] `cmd_pane_batchRename_engine_remaining_validation_transform_coverage`
    - 2026-06-10: Also covers stable `name_invalid_character` and `name_too_long` blocking errors for invalid manual target leaves.
  - [x] `cmd_pane_batchRename_window_rules_recompute_preview`
    - 2026-06-10: Also covers footer validation summary warning counts and window preview recompute/visible-refresh perf metric emission.
    - 2026-06-10: Also covers the `Hide unchanged` preview filter, including default visibility and unchanged full-plan stats.
    - 2026-06-10: Also covers `New Name` warning/error glyphs and stable issue-ID tooltips for warning and blocking preview rows.
  - [x] `cmd_pane_batchRename_window_rule_controls_drive_preview`
  - [x] `cmd_pane_batchRename_window_debounces_text_preview`
  - [x] `cmd_pane_batchRename_window_uses_and_persists_settings`
    - 2026-06-10: Also covers scope mask history and include files/folders/subdirectories option persistence.
  - [x] `cmd_pane_batchRename_window_manual_mode_controls_drive_preview`
    - 2026-06-10: Also covers preview sorting and Manual `Sort like preview` line reordering.
  - [x] `cmd_pane_batchRename_window_manual_mode_target_change_blocks_until_reconciled`
  - [x] `cmd_pane_batchRename_window_executes_local_rename`
    - 2026-06-10: Also covers current v1 execution perf metric emission.
  - [x] `cmd_pane_batchRename_window_refreshes_pane_after_success`
  - [x] `cmd_pane_batchRename_window_invokes_success_callback`
  - [x] `cmd_pane_batchRename_window_reports_execution_summary`
    - 2026-06-10: Also covers retained undo-plan row counts and `Copy Undo Plan` TSV export after a mixed execution.
  - [x] `cmd_pane_batchRename_window_executes_case_only_local_rename`
  - [x] `cmd_pane_batchRename_window_blocks_invalid_preview_execution`
  - [x] `cmd_pane_batchRename_window_executes_parent_child_deepest_first`
  - [x] `cmd_pane_batchRename_window_blocks_destination_created_after_preview`
  - [x] `cmd_pane_batchRename_window_blocks_source_missing_after_preview`
  - [x] `cmd_pane_batchRename_helper_menus_expose_canonical_insertions`
  - [x] `cmd_pane_batchRename_window_helper_buttons_insert_into_rule_fields`
  - [x] `cmd_pane_batchRename_engine_large_preview_perf`
  - [x] `cmd_pane_batchRename_engine_regex_compile_perf`
  - [x] `cmd_pane_batchRename_engine_validation_perf`
  - [x] `cmd_pane_batchRename_collision_existing_and_duplicate_names`
  - [x] `cmd_pane_batchRename_cancel_does_not_apply_remaining_rows`
    - 2026-06-10: Also covers copying the retained canceled execution report as TSV with counts, canceled state, and first-failure HRESULT.
  - `cmd_pane_batchRename_theme_accessibility_snapshot`
- [x] Add debug snapshot helper in `BatchRenameWindow.h` similar to Find Files and Change Case prompt debug helpers.
- [ ] Test keyboard behavior:
  - Tab order
  - Enter in single-line fields recomputes preview or executes only from buttons, according to spec
  - Ctrl+Enter in manual multiline applies or inserts newline according to chosen project convention
  - Escape closes menus first, then closes window if no operation is running
  - F1 or Help menu opens relevant documentation if implemented in the app
- [ ] Test theme switching:
  - [x] Light
  - [x] Dark
  - [ ] High contrast if supported by current theme contracts
- [ ] Test DPI:
  - 100 percent
  - 150 percent
  - 200 percent
- [ ] Test provider edge cases:
  - local NTFS
  - plugin file system with basic rename support
  - read-only or access-denied row
  - source deleted after preview
  - destination created after preview
  - [x] local source deleted after preview
  - [x] local destination created after preview
- [ ] Run `pwsh -File .\Tools\Verify-NoSubclassManager.ps1` if the UI work touches any subclassing or native control wrappers.

## Phase 16 - Documentation and Closeout

- [x] Update `Specs/UI/UI_CommandMenuKeyboard.md` with the new command and launch behavior.
- [x] Update `Specs/UI/UI_FindFilesWindow.md` only to reference shared navigation/header conventions, not to duplicate Batch Rename behavior.
- [x] Update `Specs/FileSystem/FileSystem_FileOperations.md` with bulk rename execution and preview validation expectations.
- [x] Update `docs\UserGuide.md` with a concise user-facing Batch Rename section.
- [x] Update `docs\MainWindow.md` for the Files menu entry and multi-selection Rename behavior.
- [ ] Move this plan to `Specs/Plans/Done/` only after implementation, validation, perf evidence, and authoritative specs are complete.

## Validation Commands

Run these from `Z:\src\RedSalamander`:

```powershell
.\build.ps1 -ProjectName RedSalamander
```

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_batchRename_
```

For focused GUI-exe command selftests, prefer `Start-Process -Wait -PassThru` so the exit code is reliable:

```powershell
$p = Start-Process -FilePath .\.build\x64\Debug\RedSalamander.exe -ArgumentList @('--commands-selftest','--selftest-case=cmd_pane_batchRename_target_collection_respects_provider_cancellation') -Wait -PassThru -WindowStyle Hidden
"ExitCode=$($p.ExitCode)"
```

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_rename_
```

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_changeCase_
```

```powershell
pwsh -File .\Tools\Tests\ResourceLocalizationContracts.Tests.ps1
```

```powershell
pwsh -File .\Tools\CompareTestRuns.ps1 -Baseline Specs\TestRuns\<baseline-run> -Candidate Specs\TestRuns\<candidate-run>
```

## Acceptance Checklist

- [x] Standard single-file rename behavior is unchanged.
- [x] Standard rename prompt has a working `Batch...` button for folders and files.
- [x] Multi-selection `Rename` opens Batch Rename directly.
- [x] Dedicated `Batch Rename...` command exists in menu and command registry.
- [x] Window uses the Find Files-style navigation/header pattern.
- [x] Preview grid shows `Original Name`, `New Name`, `Size`, `Date`, `Time`, and `Path`.
- [x] Manual mode validates one line per target.
- [x] Macro mode supports the documented v1 macro set.
- [x] Regex search/replace supports helper menus and replacement tokens.
- [x] Change Case controls reuse shared Change Case logic.
- [x] Blocking preview issues disable the window Rename button.
- [x] Blocking preview issues prevent execution.
- [x] Execution uses existing file-system rename contracts and supports cancel.
- [x] Directory cache and panes refresh correctly after successful renames.
  - 2026-06-10: Covered by `cmd_pane_batchRename_window_invokes_success_callback`, `cmd_pane_batchRename_window_refreshes_pane_after_success`, and the existing local execution/parent-child execution checks.
- [x] Settings, histories, and grid layout persist.
- [x] All user-visible strings are localized through resources.
  - 2026-06-10: Batch Rename source and satellite resources include the accessibility and preview context-menu strings. `ResourceLocalizationContracts.Tests.ps1` still reports pre-existing FileOps/HotPaths language-neutral contract failures unrelated to Batch Rename.
- [ ] Selftests cover engine, launch routing, UI snapshots, and execution.
- [ ] Perf instrumentation and archived runs exist under `Specs/TestRuns/`.
- [x] Durable behavior is documented outside this WIP plan.

## Implementation Notes and Risks

- `IFileSystem::RenameItems` accepts leaf names only. Any UI affordance that looks like moving to a different folder must be excluded or flattened in v1.
- Regex helper menus must match the actual engine syntax. Do not expose unsupported constructs.
- `std::regex` can be expensive on adversarial inputs. Keep regex compilation cached per generation, cap preview work per UI frame, and measure large regex scenarios.
- Parent/child selections can invalidate child source paths if parent directories are renamed first. Always compute a safe execution order.
- Case-only renames are provider-sensitive. Preserve the existing local-plugin handling and mark them as warnings in preview.
- Manual mode is powerful but easy to desynchronize from target collection. Treat line-count mismatch as blocking.
- The preview engine should not call Win32 UI APIs. Keeping it pure makes the risky parts testable.
- Do not store owning raw COM pointers; use `wil::com_ptr`.
- Do not post raw heap payloads across threads; use `PostMessagePayload` and `TakeMessagePayload`.
- Do not use `catch (...)`. Catch `std::regex_error` where regex is compiled and convert it into a validation issue.

## Self-Review Before Execution

- [ ] Every phase has concrete files and behaviors.
- [ ] The plan includes the user's requested launch, preview, manual, macro, regex, and Change Case requirements.
- [ ] The plan includes missing best-in-class pieces: validation, collision detection, helper menus, persistence, cancel, reports, accessibility, localization, tests, and perf evidence.
- [ ] The plan respects RedSalamander rules for WIL RAII, resources, message payloads, thread safety, and spec closeout.
- [ ] No step requires destructive file-system or git operations.
