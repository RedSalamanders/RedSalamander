# Batch Rename

**Batch Rename** is the preview-first surface for renaming many files and folders at once. You build a rename plan from macro templates, search/replace rules, and case transforms (or type names by hand), see every proposed name validated in a grid, and only then apply the renames through the file-system plugin layer.

This page covers the user-facing window first, then the developer-facing engine, window, and execution split. For the standalone **Change Case** command see the [Change Case](#change-case) section below; for the broader copy/move/delete surface see [FileOperations.md](FileOperations.md).

## Open

- **Rename** (`F2`): with two or more items selected, RedSalamander opens Batch Rename instead of the single-item rename prompt.
- The single-item rename prompt offers a **Batch...** action: for a folder it opens Batch Rename rooted at that folder; for a file it opens Batch Rename seeded with just that file.
- Folder scope: open Batch Rename on a folder to rename its children, optionally filtered by a mask and recursed into subdirectories.

The window is modeless, so you can keep working in the panes while it is open.

## Scope

The top row uses the same address-bar **NavigationView** as [Find Files](FindFiles.md), seeded from the launching pane's plugin, instance, and current folder.

| Control | Effect |
| --- | --- |
| **Mask:** | Wildcard include/exclude filter for folder-scope collection. A plain `*.*` mask means "all items". |
| **Include subdirectories** | Recurse into child folders when collecting targets. |
| **Files** | Include files in the target set. |
| **Folders** | Include folders in the target set. |

Masks use `*` and `?` glob wildcards, are case-insensitive, and accept a `;`-separated list. A `|` splits the text into include patterns (left) and exclude patterns (right); write `;;` for a literal semicolon. Recent masks are remembered as a most-recent-first history.

Explicit multi-selection launch uses exactly the selected paths and does **not** enable recursion by default. File rows carry byte size and last-write date/time; folder rows leave size blank.

## Rules mode

Rules mode (the default) builds each new name by applying, in order:

1. **New name** template macro expansion.
2. **Search / replace** (literal or regex).
3. **Change Case** transforms for the name and extension independently.
4. Validation and duplicate detection.

### Template macros (New name field)

Canonical macro syntax is `{macro}`. The `$(Macro)` form is also accepted for compatibility, but the helper menu always inserts canonical `{...}` tokens. Write a literal brace as `{{` or `}}`. Unknown macros are blocking errors.

| Macro | Expands to |
| --- | --- |
| `{name}` / `{filename}` | The whole original leaf name (stem + extension). |
| `{stem}` | File name without extension. |
| `{ext}` | Extension, including the leading dot. |
| `{extNoDot}` | Extension without the leading dot. |
| `{parent}` | Name of the immediate parent folder. |
| `{relativeFolder}` | Containing folder relative to the Batch Rename root. |
| `{relativeFolderFlat}` | Same relative folder with path separators replaced by the flatten separator (default ` - `). |
| `{size}` | File size in bytes. |
| `{date}` / `{date:yyyy-MM-dd}` | Last-write date. Default format `yyyy-MM-dd`. |
| `{time}` / `{time:HH-mm-ss}` | Last-write time. Default format `HH-mm-ss`. |
| `{created}` / `{created:yyyy-MM-dd}` | Created date/time, same token set. |
| `{counter}` / `{counter:000}` | 1-based row counter, optionally zero-padded. |
| `{index}` | 0-based row index (same padding rules as `{counter}`). |

Date/time format tokens are `yyyy`, `MM`, `dd`, `HH`, `mm`, `ss`; any other characters are copied through literally. Timestamps are formatted in local wall-clock time (falling back to UTC if time-zone data is unavailable). Counter/index padding accepts an all-zero pattern (`{counter:000}` pads to the pattern length) or a plain numeric width (`{counter:3}` zero-pads to three digits); other non-empty formats are blocking `macro_invalid_format` errors.

The preview row order is authoritative for `{counter}` and `{index}`, so sorting the grid changes the numbers each row receives.

### Search / replace

The **Search for** and **Replace with** fields rewrite the expanded name. Options:

| Option | Effect |
| --- | --- |
| **Regex** | Treat **Search for** as an ECMAScript `std::wregex` pattern. When off, the search is a literal substring. |
| **Case sensitive** | Match case exactly. When off, matching folds case. |
| **Whole words** | Match only at word boundaries. In regex mode the pattern is wrapped as `\b(?:...)\b`, preserving your `$1`, `$2`, ... group indexes. |
| **Only once in each name** | Replace only the first match per name. |
| **Exclude extension** | Apply search/replace to the stem only, leaving the extension untouched. |

Invalid regex patterns are blocking errors. Match-time regex failures (for example backtracking-heavy patterns that hit the engine's complexity limit) become a per-row blocking error rather than crashing the preview.

In regex mode the **Replace with** field supports ECMAScript replacement tokens such as `$&` (whole match), `$1`/`$2` (captured groups), and `$$` (a literal dollar sign).

Literal whole-word replacement never accepts a match boundary that splits a UTF-16 surrogate pair.

### Case changes (independent name vs extension)

Two dropdowns under the rule fields apply casing to the **file name** and the **extension** separately. Each offers:

| Style | Result |
| --- | --- |
| **Do not change** | Leave casing as is. |
| **Lower case** | All lowercase. |
| **Upper case** | All uppercase. |
| **Mixed case** | Capitalize the first letter of each word. |

These reuse the same casing logic as the standalone [Change Case](#change-case) command, so `report.PDF` can become, for example, name `Report` + extension `.pdf` in one pass.

### Helper menus

A drop-down helper button sits next to the template, search, and replace fields. Each command inserts canonical text at the caret (or replaces the current selection):

- **Template helper**: `{name}`, `{stem}`, `{ext}`, `{extNoDot}`, `{parent}`, `{relativeFolderFlat}`, `{counter}`, `{counter:000}`, `{date:yyyy-MM-dd}`, `{time:HH-mm-ss}`, and the literal braces `{{` / `}}`.
- **Regex helper**: only ECMAScript-supported syntax, including `.`, `[abc]`, `[^abc]`, `\b`, `|`, `*`, `+`, `?`, `()`, the character classes `\w \W \s \S \d \D`, `\d+`, `[0-9A-Fa-f]+`, the file-name split `^(.+?)(\.[^.]+)?$`, and a **Selected text as literal** command that escapes regex metacharacters in the current search-field selection. With an empty selection it inserts a two-backslash skeleton and selects the second slash for quick overwrite.
- **Replacement helper**: `$$`, `$&`, `$1`, `$2`, and a **Matched subexpression...** command that prefixes a selected run of digits with `$` (or inserts `$1` and selects the `1` for quick overwrite when nothing numeric is selected).

## Manual mode

Switch the **Manual** radio button (next to **Rules**) to type names directly, one target leaf name per line. The first switch into Manual seeds the editor from the current preview **New Name** values.

- The line count must match the target count before rename can run.
- Manual names are exact input except that carriage returns are stripped; leading and trailing spaces are **not** silently trimmed.
- Manual text is transient: it is never persisted and never written into rule histories.

Command-row buttons:

| Button | Action |
| --- | --- |
| **Fill from preview** | Replace the editor with the current preview new names. |
| **Clear** | Empty the editor. |
| **Paste** | Read Unicode clipboard text into the editor without normalizing line endings. |
| **Sort like preview** | Reorder the lines into the current preview sort order over the complete target set, keeping each target paired with its name. |

If the target set changes while Manual mode is active, your text stays visible and the preview stays blocked until the line counts line up again.

## Preview grid

The grid shows one row per target and validates every proposed name before any rename is allowed.

| Column | Contents |
| --- | --- |
| **Original Name** | Current leaf name, rendered as an icon + text cell. |
| **New Name** | Proposed name. A warning/error glyph and tooltip flag any issues. |
| **Size** | File size (blank for folders). |
| **Date** | Last-write date. |
| **Time** | Last-write time. |
| **Path** | Containing folder relative to the window root (empty for items directly under the root; absolute parent folder for items outside it). The cell tooltip shows the absolute path. |

The grid supports column resize/reorder, sorting, clipped-cell tooltips, and keyboard selection. Rows with errors use the grid error tone; rows with only warnings use the warning tone. Each **New Name** issue tooltip reads as a localized description followed by its stable ID in parentheses, for example `Unknown macro (macro_unknown)`.

- **Hide unchanged** (footer toggle): hides rows whose name will not change. This is purely a view filter — statistics, validation, and execution always use the complete plan.
- Sorting the **Path** column uses the displayed root-relative path text; equal path keys keep the existing preview/source order in either direction.
- The footer summary shows total rows, changed rows, unchanged rows, warnings, and errors. **Rename** is disabled whenever any row has a blocking error.

### Issue IDs

Blocking errors include `name_empty`, `name_dot`, `name_separator`, `name_invalid_character`, `name_reserved_device` (classic DOS device names such as `CON`, `NUL`, `CONIN$`, `CONOUT$`, `COM1`, bare or with an extension), `name_too_long` (over 255 UTF-16 code units), `name_duplicate`, `macro_unknown`, `macro_unclosed`, `macro_invalid_format`, `regex_invalid`, `regex_match_failed`, and `manual_line_count`. Warnings include `name_unchanged`, `name_case_only`, and `name_edge_space_or_dot`.

### Context menu and activation

Right-click a preview row for:

- **Copy** the clicked row's original name, new name, or source path.
- **Copy** the full visible preview as tab-separated text with localized headers in column order.
- **Reveal in Active Pane** (when a reveal callback is available): navigate the owning pane to the row's parent folder and focus the item. Double-click or Enter does the same.
- **Copy Execution Report** (after a run): a tab-separated summary with total rows, completed rows, skipped rows, failed rows, canceled state, the first-failure HRESULT, and its message text.
- **Copy Undo Plan** (after a run that renamed at least one row): a tab-separated list of current path, restore leaf name, and original path for each successfully renamed row. This is an informational artifact you can keep — it is **not** an automatic undo command.

View-only refreshes (sorting, **Hide unchanged**) preserve the retained execution report and its undo plan; the report is cleared only when the plan actually changes (rule edits, manual edits, scope or target-set changes).

## Running the rename

Press **Rename** to apply the plan. While a collection or execution task runs, **Rename** is disabled and the footer **Cancel** button is enabled; the status strip shows `Renaming {completed} of {total}...`. Cancelling stops further renames and marks the retained report `canceled` — rows already renamed stay renamed.

After a run, the status strip shows a result summary, and affected panes refresh to the new names.

---

# Developer reference

Batch Rename is split into a pure engine, a modeless DxUi window, helper-menu logic, a standalone Change Case command, and the low-level batch marshaller they all share. See [DeveloperGuide.md](DeveloperGuide.md) (section "Batch Rename, Change Case & Rename Batching") for the subsystem overview and [`Specs/UI/UI_BatchRenameWindow.md`](../Specs/UI/UI_BatchRenameWindow.md) for the authoritative contract.

## Components

| File | Role |
| --- | --- |
| `RedSalamander/BatchRenameEngine.{h,cpp}` | Pure `BatchRename::BuildPlan(targets, rules)` → `Plan`. Macros, search/replace, case, validation. No I/O, `noexcept`. |
| `RedSalamander/BatchRenameWindow.{h,cpp}` | Modeless `BatchRenameWindow` (DxUi): scope/collection, preview grid, threading envelope, settings, undo plan, debug API. |
| `RedSalamander/BatchRenameExecutionEngine.{h,cpp}` | Window-free provider-dispatch engine: dependency layers, temp-hop cycles, per-item outcomes, undo entries, directory-move path rewriting, and execution metrics. |
| `RedSalamander/BatchRenameMenus.{h,cpp}` | Template/regex/replacement helper flyouts and caret-aware insertion. |
| `RedSalamander/MaskSyntax.{h,cpp}` | `*.*`-style include/exclude wildcard masks plus MRU history. |
| `RedSalamander/ChangeCase.{h,cpp}` | `cmd/pane/changeCase` casing logic plus the recursive `ApplyToPaths` driver. |
| `RedSalamander/FileSystemRenameBatch.{h,cpp}` | Arena-backed `Execute()` that calls `IFileSystem::RenameItems`, falling back to per-item `RenameItem`. |

## The engine: BuildPlan

`BatchRename::BuildPlan` is pure and `noexcept`. It takes a `std::vector<Target>` and a `Rules` struct and returns a `Plan` (rows + `Stats`). Each row records the source path, original leaf, proposed new leaf, file/folder flag, size, timestamps, a stable `rowId`, and a vector of `Issue`s.

For each target it:

1. Expands the template via `ExpandTemplate`/`ResolveMacro`. `NormalizeAliasTokens` rewrites `$(Macro)` → `{macro}` first; `{{`/`}}` collapse to literal braces; an unclosed `{` emits `macro_unclosed` and an unknown name emits `macro_unknown`.
2. Applies search/replace via `ApplyReplacement` — either `ReplaceLiteral` (with case folding, whole-word, and once semantics) or a single `std::wregex` compiled **once per preview** before the row loop. `Exclude extension` runs the replace on the stem only and re-appends the extension.
3. Applies casing via `ApplyCaseTransforms`, which delegates to `ChangeCase::TransformLeafName` once for `OnlyName` and once for `OnlyExtension`.
4. Validates: `ValidateLeafName` (empty, dot, separator, invalid/control chars, reserved device names, length), `AddWarningIssues` (unchanged, case-only, edge space/dot), and `MarkDuplicateTargets` (same parent + same provider-identity leaf). Duplicate validation uses helper-owned path/component keys when available and verifies every key hit with `FileSystemPathIdentity` equality; if keys are not safe for a profile, it falls back to direct identity comparisons.

Manual mode skips steps 1–3 and instead pulls names from `rules.manualNames` by index, flagging `manual_line_count` when the line count does not match the target count.

The engine uses allocation-light stem/extension splitting on the hot macro path, keeping the edge-case behavior pinned to `std::filesystem::path` by selftest. The window debounces edits (~150 ms) through `RequestPreviewRebuild` → `RebuildPreview` → `BuildPlan` plus the window's contextual validation. Perf is instrumented through `batchrename.preview.build_plan_us`, `batchrename.regex.compile.us`, `batchrename.validation.us`, and the `batchrename.preview.*` count metrics.

## Execution ordering

Execution is the safety-critical part. `BatchRenameWindow::ExecuteRename` builds one `BatchRenameExecutionOp` per non-no-op row (`currentSource`, `originalSource`, `finalLeaf`, `depth`, `isDirectory`), then sorts deepest-first. The depth key is the count of path separators (`PathDepthKey`); ties break by longer path first. The worker `std::jthread` owns MTA COM initialization and completion posting, then calls the window-free `RunBatchRenameExecutionEngine` in `BatchRenameExecutionEngine.{h,cpp}`, which combines three ideas:

1. **Deepest-first depth groups.** Ops are processed one depth group at a time, deepest first, so a child is renamed before its parent directory is renamed out from under it. Within a group, all ops share the same depth.

2. **Dependency layering (chains).** Within a depth group, an op whose target leaf identifies another *pending* op's source is "blocked" until that source is vacated. Each round collects the unblocked ops into one **layer**, dispatched together as a single provider batch (which the provider may parallelize). Successive layers run sequentially. A self-rename whose target identifies only its own source (a case-only rename) is not treated as a dependency on itself. This lets chains like `a → b`, `b → c` admitted by preview validation succeed without clobbering.

3. **Temp-hop for cycles (swaps).** If a round finds no unblocked op, every pending op targets another pending source — a pure rename cycle. One member is renamed to a unique temp leaf in the same directory (`<original leaf>.rsren-<guid suffix>` from `MakeBatchRenameTempLeaf`). That vacates a name so the rest of the cycle can drain on the next round; the temp entry is renamed to its final name afterward.

Bookkeeping:

- Undo entries and the success callback record only the **net** original→final transition; temp hops are an implementation detail. A failed temp step surfaces as that row's failure, and an orphaned temp is renamed back to its original name best-effort (`restoreTempBestEffort`).
- Per-item outcomes come from `BatchRenameExecutionCallback::FileSystemItemCompleted`, so a partially failed parallel batch still records undo entries, notifies `DirectoryInfoCache::NotifyPathMoved`, and counts only the rows that actually failed or never ran. Providers that never report per-item completion fall back to all-or-nothing semantics for that batch.
- Successful directory renames are tracked as `ExecutedDirectoryMove`s and replayed (`ApplyExecutedDirectoryMoves`) so child paths in undo entries and the window's own target refresh move under the renamed parent. The `onSuccessfulRename` callback, by contrast, reports each row's path *as of its own rename* (parent moves not folded in); consumers replay the parallel source/target lists sequentially — see the contract in `BatchRenameWindow.h` and `FolderWindow::RefreshPanesAfterBatchRename`.

The engine returns a `BatchRenameExecutionResult`; the worker wrapper posts it back to the UI thread, where the retained `BatchRenameExecutionReport` (totals, first failure, canceled flag, `undoEntries`) feeds the status strip, **Copy Execution Report**, and **Copy Undo Plan**.

> **Note:** `Specs/UI/UI_BatchRenameWindow.md` mandates routing every identity-sensitive "same path?" decision through one resolved `FileSystemPathIdentity` profile. Batch Rename now uses that profile for duplicate/case-only validation, dependency layering, provider callback pairing fallback, destination probing, cache notification, target refresh, provider-selection matching, and undo path finalization. The shared helper also owns accepted-separator-aware path equality and optional map keys; same-context Copy/Move/Delete rejects missing, malformed, unsupported, or unstable provider `pathIdentity` before task creation.

## FileSystemRenameBatch

`FileSystemRenameBatch::Execute` is the shared marshaller for every multi-item leaf rename (Batch Rename and Change Case both use it). Given a span of `RenameOp` (`sourcePath`, `newLeaf`, `depth`, `isDirectory`), it:

- Computes the total buffer size, copies each source path and new leaf into a plugin-arena (`FileSystemArenaOwner`), and builds a `FileSystemRenamePair` array obeying the arena ownership contract.
- Rejects empty source/leaf (`E_INVALIDARG`) and new leaves containing path separators (`ERROR_INVALID_NAME`) — `RenameItems`/`RenameItem` take **leaf names only**.
- Calls `IFileSystem::RenameItems` once for the whole batch. If the provider reports bulk rename unsupported (`E_NOTIMPL`, `ERROR_CALL_NOT_IMPLEMENTED`, or `ERROR_NOT_SUPPORTED`), it falls back to one `RenameItem` per op, reporting `FileSystemItemCompleted` with the submitted index for every attempt (including the failing one) and honoring `FileSystemShouldCancel` between items.

## Change Case

The **Change Case** command (`cmd/pane/changeCase`, default shortcut `Ctrl+F7`) applies casing to selected items without the full preview window. It is also the engine Batch Rename calls for its case dropdowns.

`ChangeCase::TransformLeafName(leafName, options)` is the pure transform. `Options` combines a `CaseStyle` and a `ChangeTarget`:

| `CaseStyle` | Effect |
| --- | --- |
| `Lower` | All lowercase. |
| `Upper` | All uppercase. |
| `Mixed` | Capitalize the first letter of each word (via the Windows character tables, so non-ASCII initials work). |
| `PartiallyMixed` | Stem in mixed case, extension in lower case (when applicable). |

| `ChangeTarget` | Applies to |
| --- | --- |
| `WholeFilename` | Stem and extension together. |
| `OnlyName` | The stem only (extension preserved). |
| `OnlyExtension` | The extension only (stem preserved). |

`ChangeCase::ApplyToPaths` is the recursive driver behind the standalone command. With `includeSubdirs` it walks the tree non-recursively through `IFileSystem::ReadDirectoryInfo` (skipping reparse points), de-duplicates paths, then sorts deepest-first (more separators first, longer path as tiebreak) so children rename before parents. Renames are dispatched through `FileSystemRenameBatch::Execute` in 64-item batches per depth group, with cooperative cancellation via `std::stop_token` and progress through a `ProgressCallback`.

## Cross-links

- User: [UserGuide.md](UserGuide.md) · [FileOperations.md](FileOperations.md) · [FindFiles.md](FindFiles.md) · [KeyboardShortcuts.md](KeyboardShortcuts.md)
- Developer: [DeveloperGuide.md](DeveloperGuide.md)
- Specs: [`Specs/UI/UI_BatchRenameWindow.md`](../Specs/UI/UI_BatchRenameWindow.md) · [`Specs/Plugins/Plugins_VirtualFileSystem.md`](../Specs/Plugins/Plugins_VirtualFileSystem.md)
