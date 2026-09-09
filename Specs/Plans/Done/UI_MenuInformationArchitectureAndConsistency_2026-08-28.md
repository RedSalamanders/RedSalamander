# Menu information architecture, coherence, and discoverability

> **NON-NORMATIVE COMPLETED RECORD.** Moved from `Specs/Plans/WIP/` on
> 2026-08-29. **No remaining action on this file.** Durable menu behavior is
> owned by the authoritative specifications under `Specs/<Domain>/`, resource
> contracts, public command identifiers, and executable tests.

## Status

| Field | Value |
|---|---|
| State | **COMPLETE** |
| Index owner | Archived from **I16**; no live WIP owner |
| Priority | **P0** localized command existence and menu-contract coverage; **P1** main/context-menu coherence; **P2** secondary-app and viewer normalization |
| Planned at | `388f31695bf1bdf8317c4e55da88972c664b6359` (`388f31695`) |
| Implementation reconciled at | `25314bea5350e246f6ff4b9b2bc020ffabcc8794` (`25314bea5`); the remaining menu package is an uncommitted working-tree change |
| Updated | 2026-08-29 |
| Scope | All application-owned static and dynamic menu surfaces in RedSalamander, RedSalamanderMonitor, RedConfigure, and viewer plugins |
| Product goal | Make commands predictable to find, remove accidental duplication, keep context menus contextual, and preserve shortcuts/stable command IDs |
| Drift check | `git diff 388f31695bf1bdf8317c4e55da88972c664b6359 -- RedSalamander RedSalamanderMonitor RedConfigure Plugins/ViewerText Plugins/ViewerImgRaw Plugins/ViewerPE Plugins/ViewerSpace Plugins/ViewerWeb Common/Common/LocalizationManager.cpp Tools/Tests/ResourceLocalizationContracts.Tests.ps1 Specs/UI Specs/Core Specs/Plugins Specs/Testing` |

## Progress checklist

`[x]` means implemented, validated at the package's required level, and
documented. `[+]` means the active package. A commit is not required by this
plan. Do not check a parent package while a child test, translation,
specification, or evidence row remains open.

- [x] **M0. Reconcile the audit against the implementation commit**
  - [x] Run the drift command and re-inventory every surface in section 4.
  - [x] Record any changed command, resource, or active-plan ownership before editing.
  - [x] Capture the current English and satellite menu trees as the behavioral baseline.
- [x] **M1. Install complete structural and localization contracts**
  - [x] Add MENU/MENUEX structural-parity checks for every base/satellite resource pair.
  - [x] Add sibling-level access-key checks with explicit dynamic/disabled exceptions.
  - [x] Add table-driven runtime menu-tree checks for the main application.
  - [x] Add a command-surface coverage ledger that classifies every canonical command and fails on unclassified additions.
  - [x] Repair the missing localized Batch Rename entries before broad reordering.
- [x] **M2. Normalize the RedSalamander main menu**
  - [x] Apply the exact target tree in section 7 without changing stable command IDs; make only the reviewed `Ctrl+Alt+T` default migration.
  - [x] Keep both intentional Path from Other Pane placements in each pane menu.
  - [x] Add Left/Right Terminal Pane after Preview Pane with `Alt+7`, targeting the physical opposite pane.
  - [x] Make Commands > Terminal > Command Shell Window open the floating terminal surface.
  - [x] Retain Change Case and Change Attributes in Files with their existing shortcuts and behavior.
  - [x] Expose Insert Full Path in Terminal through a real static menu command ID.
  - [x] Specify the existing Connections Manager and Failed Operations Pane commands.
  - [x] Remove only accidental duplicate placements and encode intentional duplicates in a reviewed allowlist.
- [x] **M3. Split and simplify contextual menus**
  - [x] Give FolderView item and background invocations different static base menus.
  - [x] Replace the Find Files “This item” plus “Selection” duplication with one target model.
  - [x] Preserve artifact, Shell, provider, navigation, Compare, Batch Rename, and Terminal ownership boundaries.
- [x] **M4. Normalize viewer and utility-app menu families**
  - [x] Apply the common viewer File-menu skeleton where capabilities exist.
  - [x] Restructure ViewerText View and Encoding surfaces.
  - [x] Normalize ViewerImgRaw, ViewerWeb, and Monitor; retain RedConfigure File > Exit under the section 9.8 STOP gate.
  - [x] Retain intentionally minimal/no-menu viewers and document why.
- [x] **M5. Complete access keys and localized resource updates**
  - [x] Give every actionable static sibling a unique access key in each supported locale.
  - [x] Recheck all four satellites against the base trees and resolve translation review blockers.
  - [x] Verify keyboard invocation and displayed shortcut text through the DxUi menu path.
- [x] **M6. Validate responsiveness-sensitive dynamic work**
  - [x] Instrument the ViewerText searchable encoding picker before replacing the current cascade.
  - [x] Run same-machine baseline/candidate measurements and archive the Viewers evidence.
  - [x] Keep static ordering-only changes free of unrelated performance instrumentation.
- [x] **M7. Close out**
  - [x] Update every authoritative menu-owning specification listed in section 14.
  - [x] Run focused tests, full build, Fresh Full, spec/test inventories, and archive checks.
  - [x] Move this file to `Specs/Plans/Done/` and remove I16 from the WIP index.

## Closeout evidence

| Gate | Result |
|---|---|
| Source identity | Final validation ran from the uncommitted menu package over `9d375606399c109487188b32569e47c3a62e9b04`. Unrelated WIP changes remained untouched. |
| Build | Complete test-enabled x64 Debug solution build passed with 0 warnings and 0 errors; final post-toolchain artifact receipt `9de64a3e4bdf15b72fffd8ee9d65ddd27fd3c98d28bd855ba3a3c1269dd4653a`. |
| Focused contracts | Resource/localization Pester passed 9/9. Exact menu-tree, duplicate-allowlist, Terminal Pane/floating-window, Change Case, Change Attributes, command-surface, context-menu, and ViewerText picker cases passed. |
| Command coverage | `CommandSurfaceCoverage.inc` classifies all 171 canonical commands and proves 100% classification, user-facing feature-family reachability, and Folder/Terminal palette parity. Final Commands entry passed 879, failed 0, skipped 2. |
| Fresh Full | Fresh run `20260828T232206Z-2796-d3a9279ab06b4c36b49a9a300e635da2` passed all 18 entries: 2,010 total cases, 1,957 passed, 0 failed, 53 documented environment/opt-in skips; flaky/regression/isolation/unclassified counts are all zero. Validation evidence is under `.build/ValidationEvidence/runs/20260828T232206Z-2796-d3a9279ab06b4c36b49a9a300e635da2/`. |
| ViewerText performance | Five independent test-enabled x64 Release processes validated the 140-row searchable encoding picker. Open-to-ready p95 was 25,150 us under the 50,000 us ceiling; filter p95 was 12,112 us under the 16,670 us ceiling. Archive: `Specs/TestRuns/4cb089111a23/Viewers/20260828_192400_menu_encoding_picker_release/`. |
| Inventories | Focused specification-architecture Pester passed 22/22; Tools inventory classified 183/183 files with 0 blocking findings; spec inventory passed with 0 blocking findings; test inventory generated successfully; whole TestRuns archive validation passed all 1,216 files; `git diff --check` exited zero. |

## 1. Purpose

The current menu system has strong command registration, shortcut integration,
localization loading, and DxUi presentation, but the information architecture
has accumulated three kinds of drift:

1. **Existence drift:** a command is documented or registered but missing from a
   menu, or is present only in the embedded English resource.
2. **Placement drift:** one command appears in multiple locations without a
   distinct contextual reason, or appears under a category that does not match
   its effect.
3. **surface drift:** related applications and viewers expose equivalent
   capabilities under different top-level structures or overly long cascades.

This plan fixes those problems without redesigning File Operations semantics or
provider/Shell-owned command order. Stable command identity and user overrides
remain intact. The one deliberate default-shortcut change splits embedded and
floating terminal behavior as specified in sections 6 and 7. The plan separates
correctness/parity work from subjective reordering so the P0 repair can land
first and be reverted independently.

### M0 reconciliation record — 2026-08-28

- The implementation was reconciled after `HEAD` advanced from the planned
  `388f31695` baseline to `25314bea5`; the drift review kept the menu package
  scoped to the owners listed in sections 4 and 11.
- The base and satellite `.rc` trees captured at the planned commit were used as
  the pre-implementation structural baseline for the M1 contract.
- Existing dirty work overlaps only normative/testing specifications owned by
  I15/I10 (`Specs/Terminal/Terminal_EmbeddedPlugin.md`,
  `Specs/Testing/Testing_PerformanceValidation.md`,
  `Specs/Testing/Testing_SelfTests.md`, and related governance prose). I16 edits
  those files by narrow additive patches and preserves the existing bytes.
- The runtime-builder re-inventory found the planned owners plus
  `NavigationView.Interaction.cpp`; its menu-related behavior stays under the
  Navigation keep contract and adds no new target surface.

## 2. Review principles

The target follows these testable rules:

- A command has one primary menu home. A second placement is allowed only when
  it serves a materially different context and is listed in a reviewed allowlist.
- Top-level categories describe the effect: Files for object lifecycle, Edit for
  clipboard/selection, Commands for workflows/tools/settings, View for
  presentation, pane menus for pane-local navigation/display, and Help for help.
- Each separator creates a meaningful group. A group should normally contain no
  more than seven peer commands.
- A normal menu should remain at or below 25 visible rows; a context menu should
  remain at or below 15 visible rows. When a surface must exceed that, use a
  searchable picker or a clearly named submenu.
- Context menus show relevant actions; they do not present a nearly complete
  item menu with most rows disabled for background invocation.
- Primary/default actions come first, transfer actions next, object-management
  actions after that, and Properties last.
- Every actionable static item has one access key that is unique within its
  separator-delimited sibling group in every supported locale. Top-level rows
  and submenus without separators remain one group. Dynamic labels and disabled
  `(Empty)` placeholders are explicitly exempt unless the owning builder supplies
  numeric access keys.
- Static menu text stays in `.rc` resources. Runtime code creates only genuinely
  dynamic entries and uses resource-backed labels for host-injected commands.
- Stable `cmd/*` IDs, WM command IDs, persisted shortcuts, and Function Bar
  bindings do not change merely because a label or location changes.

Reference guidance used for the audit:

- Microsoft menu controls: <https://learn.microsoft.com/windows/apps/develop/ui/controls/menus>
- Microsoft commanding basics: <https://learn.microsoft.com/windows/apps/design/basics/commanding-basics>
- Microsoft classic menu guidance: <https://learn.microsoft.com/windows/win32/uxguide/cmd-menus>
- Microsoft settings placement guidance: <https://learn.microsoft.com/windows/apps/design/app-settings/guidelines-for-app-settings>

## 3. Ownership and coordination

### 3.1 This plan owns

- the menu-surface inventory and exact target ordering in this file;
- base/satellite MENU/MENUEX structural parity;
- RedSalamander main-menu and FolderView context-menu information architecture;
- Find Files result-menu target semantics;
- Monitor, RedConfigure, and viewer menu normalization;
- menu access-key and stable-command-existence regression coverage;
- authoritative spec updates required by these changes.

### 3.2 Existing owners remain authoritative

- **I1/I2 — FolderView work:** coordinate before changing
  `FolderView.Menus.cpp` or shared FolderView Commands fixtures. I16 owns only
  menu composition/targeting, not thumbnail, layout, focus, or render behavior.
- **I4 — Astrolabe:** owns repository-wide test-contract migration. I16 adds
  behavioral menu tests and one narrowly structural resource-parity contract; it
  does not create a second source-test migration queue.
- **I5 — performance measurement contract:** owns reusable measurement and
  archive infrastructure. I16 reuses it for the ViewerText encoding picker.
- **I6 — Rosetta Lantern:** owns human-reviewed RedConfigure Czech, Japanese,
  and Slovak localization. Coordinate RedConfigure menu removal/resource edits;
  do not overwrite translations being reviewed there.
- **I10 — tooling governance:** coordinate any material edit to
  `Tools/Tests/ResourceLocalizationContracts.Tests.ps1` and follow
  `Tools/README.md`, `Specs/Testing/Testing_ToolingGovernance.md`, and the
  tooling-governance skill.
- **I15 — Terminal Ghostty remediation:** coordinate edits to shared Terminal
  command, lifecycle, metric, and test files. I16 owns only menu labels/placement,
  pane-target dispatch IDs, and the reviewed default-shortcut split; it must
  reuse the existing embedded/floating implementations and not alter the admitted
  terminal engine or pin.
- **I12/I14 — File Operations:** own the File Operations popup/More-menu and
  post-closeout implementation work; I14 also owns remaining Change Case
  architecture debt. I16 may rename/reposition only the main menu entry that
  opens the Failed Operations Pane and must retain the existing Change Case menu
  contract. It must not redesign the popup, task actions, or Change Case behavior.

### 3.3 External and provider-owned boundaries

- Windows Shell context-menu entries queried through `IContextMenu` retain Shell
  ordering and labels.
- Plugin/provider `INavigationMenu` entries retain provider ordering. The host
  may keep its Connections/Go To/drive-list boundary but must not reorder the
  provider's internal commands.
- User Menu content, ShellNew entries, known-folder entries, themes, plugin
  actions, history, hot paths, and sibling folders stay dynamically ordered by
  their owning data source.
- This plan does not create menus for ViewerVLC or ViewerSqlite unless a separate
  capability review first proves user-visible commands that need a menu home.

## 4. Complete menu-surface inventory

The inventory is complete for application-owned menu construction at the
planned-at commit. Re-run the search in M0 because new `CreatePopupMenu`,
`AppendMenuW`, `InsertMenuW`, MENU, or MENUEX sites extend the scope.

| Owner / surface | Construction | Current disposition | Package |
|---|---|---|---|
| RedSalamander main bar | `RedSalamander/RedSalamander.rc`, `IDC_REDSALAMANDER` MENUEX | Reorder and repair existence/spec drift | M1-M2 |
| FolderView context | `RedSalamander/RedSalamander.rc`, `IDR_FOLDERVIEW_CONTEXT`; `FolderView.Menus.cpp` | Split item/background menus | M3 |
| Compare Directories bar | `IDR_COMPARE_DIRECTORIES_MENU`; compare runtime state | Keep order; add to whole-surface contract | M1 |
| Files > View With / Edit With | Static placeholder plus dynamic file actions | Keep dynamic order and empty-state behavior | M1 |
| Files > New | Static Folder plus ShellNew dynamic entries | Keep dynamic order; reuse in background context | M2-M3 |
| Commands > User Menu | Dynamic configured actions | Keep user-defined order | M1 |
| Commands > Open in File Explorer | Static current folder plus dynamic/known folders | Keep current folder first, known folders after separator | M1 |
| Plugins and Theme menus | Static manager/presets plus dynamic plugins/themes | Keep owner ordering; validate separator/empty states | M1 |
| Pane Go To / hot paths / history | Static pane menu plus dynamic lists | Keep Back/Forward/Parent/Root, then hot paths/history | M2 |
| Navigation drive/provider menu | `NavigationView.Menus.cpp` | Keep Connections before Go To and Go To before drives | M1 |
| Full-path sibling popup | `NavigationView.FullPathPopup.cpp` | Keep sibling ordering and current-path semantics | M1 |
| Filter and prompt history | `FolderWindow.FileSystem.cpp`, `FolderWindow.Layout.cpp` | Keep deduped recency order and current check | M1 |
| Pane sort flyout | `RedSalamander.cpp` | Keep None, Name, Extension, Time, Size, Attributes; slider after separator | M1 |
| Compare sort flyout | `CompareDirectoriesWindow.Menu.cpp` | Keep current sort order | M1 |
| Find Files result context | `FindFilesWindow.cpp` | Replace duplicate target blocks with one target model | M3 |
| Batch Rename preview context | `BatchRenameWindow.cpp` | Keep current grouped eight-item surface | M1 |
| Embedded/floating Terminal context | `FolderWindow.Interaction.cpp`, `FloatingTerminalWindow.cpp` | Keep Copy/Paste/Select All then Find/Suggestions/Close | M1 |
| File Operations More | File Operations popup code | Existing I12/I14 owner; do not redesign here | External |
| Shell context menu | `FolderWindow.cpp`, Shell `IContextMenu` | Shell-owned order; host only presents result | External |
| ViewerText | `ViewerTextResources.rc` plus dynamic Other Files/encoding behavior | Normalize File/View; replace oversized encoding cascade | M4-M6 |
| ViewerImgRaw | `ViewerImgRawResources.rc`; image commands | Move Other Files under File; normalize transform/adjust labels | M4-M5 |
| ViewerPE | `ViewerPEResources.rc` | Keep minimal structure; normalize common File order/access keys | M4-M5 |
| ViewerSpace | `ViewerSpaceResources.rc`; treemap insertions in `ViewerSpace.cpp` | Keep minimal File menu and treemap context insertion order | M4-M5 |
| ViewerWeb | `ViewerWebResources.rc` | Normalize common File order; move developer tools to Tools | M4-M5 |
| ViewerVLC / ViewerSqlite | No application-owned static menu resource | Intentional no-menu state pending capability evidence | M4 |
| RedSalamanderMonitor | `RedSalamanderMonitor.rc` plus dynamic themes | Rename/re-group Options; normalize access keys | M4-M5 |
| RedConfigure | `RedConfigure.rc`, one File > Exit menu | Remove the low-value one-item menu bar after close-path proof | M4-M5 |

## 5. Current-state evidence and findings

Line numbers are anchors at `388f31695`; executors must use symbols and rerun
the drift check rather than blindly editing by line.

| ID | Priority | Effort / risk | Evidence | Impact / target |
|---|---:|---|---|---|
| MNU-01 | P0 | S / low | English `RedSalamander.rc:182-184` contains `IDM_PANE_BATCH_RENAME`; all four `RedSalamander/Lang/*/*.rc` files place Rename directly before Open/Execute and omit it. `LocalizationManager.cpp:350-356` loads the entire satellite menu when available. | Czech, French, Japanese, and Slovak users lose Batch Rename from Files. Restore the item and prevent future structural drift. |
| MNU-02 | P0 | S / low | `cmd/pane/bringFullPathToTerminal` is registered with WM ID `0` at `CommandRegistry.cpp:167`, bound to `Ctrl+Shift+Enter` at `ShortcutDefaults.cpp:281`, dispatched at `RedSalamander.cpp:6646-6653`, and documented as a Commands menu row at `UI_CommandMenuKeyboard.md:584`, but no MENUITEM exists. | Give the existing stable command a real WM ID and Commands > Terminal row; do not change its shortcut or behavior. |
| MNU-03 | P0 | S / low | `IDM_VIEW_FILEOPS_FAILED_ITEMS` exists at `RedSalamander.rc:330` and maps to `cmd/app/toggleFileOperationsFailedItems` at `CommandRegistry.cpp:149-152`, but the current-menu specification omits it. | Rename the visible row to **Failed Operations Pane**, retain the stable ID, and specify its checked/state behavior. |
| MNU-04 | P0 | S / low | Connections Manager exists at `RedSalamander.rc:280` and maps to stable `cmd/pane/connections` at `CommandRegistry.cpp:185`, but the current-menu specification omits it; a selftest diagnostic still mentions obsolete `cmd/pane/connectionManager`. | Standardize specification/tests on `cmd/pane/connections`; retain the existing WM ID. |
| MNU-05 | Keep | S / low | Left and Right each contain Path from Other Pane once under Go To and again at the pane-menu root (`RedSalamander.rc:117/153` and `345/381`); current tests explicitly require both. Product review confirms that the duplication is intentional. | Keep both. Go To provides categorical discovery; the root row provides the fast pane-workflow placement. Encode both pane-specific IDs in the reviewed duplicate allowlist. |
| MNU-06 | P1 | S / low | Edit exposes the same `IDM_PANE_SELECTION_RESTORE` at `RedSalamander.rc:242` and again as Load Selection under Advanced at line 248. | Keep Save Selection and Restore Selection together under **Advanced Selection**; remove the root duplicate. |
| MNU-07 | P1 | S / medium | `IDM_PANE_CREATE_DIR` appears under Files > New and Commands (`RedSalamander.rc:218/266`). | Keep **New > Folder...** as the primary home; retain F7/Function Bar; remove the Commands duplicate. The risk is Commander-style menu expectation. |
| MNU-08 | P1 | S / low | View Width is under Files (`RedSalamander.rc:186`) and Preferences is under View (`:334`). | Move View Width to View and Preferences to the end of Commands. |
| MNU-09 | P1 | M / medium | Files has 25 command/submenu rows and Commands has 21, with unrelated actions sharing groups. | Apply the grouped target trees in section 7; keep each group bounded and named. |
| MNU-10 | P1 | M / high | `FolderView.Menus.cpp:415-545` loads one item menu for both target and background invocation and disables item commands after a background hit. The runtime case at `Commands.SelfTest.ViewCommands.cpp:22741-22745` locks this disabled-Open behavior. | Use separate localized item and background menus. Preserve focus/selection snapshots, stale-target rejection, and provider capability checks. |
| MNU-11 | P1 | M / high | `FindFilesWindow.cpp:3825-3907` builds one action set; multi-selection then duplicates the complete set as “This item” and “Selection (N)” around lines 4000-4031. | Use one standard target model: selected-row click targets the selection; unselected-row click selects and targets only that row. Preserve destructive target authority. |
| MNU-12 | P1 | M / medium | Current main-menu tests in `Commands.SelfTest.Settings.cpp:10372-10670` spot-check selected labels/order and encode known duplicates; copy-text order is separately locked at `Commands.SelfTest.ViewCommands.cpp:1444-1504`. | Replace spot checks with table-driven trees plus reviewed dynamic/duplicate allowlists without creating brittle label-only source tests. |
| MNU-13 | P2 | S / low | ViewerText/Web/PE place Other Files under File, while ViewerImgRaw makes Other Files a top-level category. | Use a capability-based common File skeleton and move ViewerImgRaw Other Files under File. |
| MNU-14 | P2 | L / medium | ViewerText View interleaves Text/Hex, five `Diff / ...` rows, navigation, and display toggles; Encoding mixes display encoding, a very large codepage cascade, and save-conversion policy. | Group representation, diff, navigation, and display commands; replace the codepage cascade with a searchable More Encodings picker and separate Save Encoding policy. |
| MNU-15 | P2 | S / low | ViewerWeb exposes DevTools under View while also having a Tools menu. | Move DevTools to Tools; retain document-view toggles in View. |
| MNU-16 | P2 | S / low | Monitor uses singular **Option**, mixes window/view behavior, and labels presets with repetitive `Preset:` prefixes. | Rename to **Options**, group general toggles before Active Filter, and use concise preset labels. |
| MNU-17 | P2 | S / medium | RedConfigure creates a menu bar with only File > Exit (`RedConfigure.rc:11-15`, `Main.cpp:90`), while the window already exposes close chrome and explicit Undo/Redo controls. | Remove the one-item menu bar after proving Alt+F4/window Close, dirty confirmation, accessibility, and I6-safe resource updates. |
| MNU-18 | P1 | M / medium | English static menus contain missing and duplicate sibling access keys, including terminal insert commands, copy-as-text rows, known folders, theme navigation, Monitor presets, and ViewerImgRaw adjustments. ViewerText's long codepage cascade has many unkeyed entries. | Enforce one unique static access key per actionable sibling; the searchable encoding picker removes the need to assign fragile mnemonics across the full codepage catalog. |
| MNU-19 | P0 | M / medium | `cmd/pane/openCommandShell` is currently the Commands > Command Shell row and is bound to `Alt+7`, `Ctrl+Alt+T`, and `Ctrl+Shift+T` in FolderView, but its production behavior already opens/reuses an embedded terminal in the physical opposite pane. The distinct `cmd/terminal/openFloatingWindow` command already opens a floating window and is bound to `Ctrl+Shift+N` only in Terminal scope. | Separate the two surfaces without changing their stable IDs: expose `cmd/pane/openCommandShell` as Left/Right > Terminal Pane (`Alt+7`), and map Commands > Terminal > Command Shell Window plus the FolderView `Ctrl+Alt+T` default to `cmd/terminal/openFloatingWindow`. |
| MNU-20 | Keep | S / low | English and all four RedSalamander satellites currently contain Change Attributes and Change Case. Both have stable registry entries, `Ctrl+F8` / `Ctrl+F7` defaults, WM dispatch, production implementations, detailed authoritative contracts, and deterministic dialog/mutation coverage. The first version of this target tree accidentally omitted them. | Retain both commands in Files and in all satellites. Treat any disappearance, shortcut change, or routing change as regression; coordinate implementation overlap with I14 without moving ownership. |
| MNU-21 | P0 | M / low | The global Command Palette already presents one row per context-eligible canonical command, but the repository has no single gate classifying each command as classic-menu, contextual/dynamic, shortcut/function-bar, palette-only, or internal/parameterized. Direct registry/resource counts are not comparable and cannot prove 100% feature coverage. | Add Command Palette to Commands for discoverability and add a table-driven coverage ledger. Require 100% classification, 100% user-facing feature-family reachability, and complete Folder/Terminal palette projections; do not require every leaf command in a classic menu. |

## 6. Stable-command and behavior invariants

Every package must preserve these invariants:

- Existing `cmd/*` values remain byte-for-byte stable, including
  `cmd/pane/connections`, `cmd/pane/bringFullPathToTerminal`, and
  `cmd/app/toggleFileOperationsFailedItems`, `cmd/pane/openCommandShell`, and
  `cmd/terminal/openFloatingWindow`.
- User shortcut overrides remain valid. Visible shortcut text continues to come
  from the effective ShortcutManager binding. The only default reassignment is
  FolderView `Ctrl+Alt+T`: migrate it from `cmd/pane/openCommandShell` to
  `cmd/terminal/openFloatingWindow` only when the persisted binding still equals
  the legacy default. Preserve `Alt+7` and FolderView `Ctrl+Shift+T` for the
  embedded opposite-pane terminal, and preserve Terminal-scope `Ctrl+Shift+N`
  for the floating window.
- New WM command IDs are allocated only where a menu cannot currently dispatch a
  registered command. Add `IDM_PANE_BRING_FULL_PATH_TO_TERMINAL` in the existing
  resource-ID family, map it in `CommandRegistry`, route it through the existing
  stable-command dispatcher, and add icon mapping only if the menu family uses one.
- Left/Right symmetric menus retain pane-specific IDs. Shared application commands
  such as Swap Panes may intentionally occur once in each pane menu and must be
  documented in the duplicate allowlist.
- Left/Right Terminal Pane rows use pane-specific WM IDs but map to the existing
  stable `cmd/pane/openCommandShell`. Selecting Left uses Left as the source and
  opens/reuses the physical Right terminal; selecting Right does the inverse.
- Path from Other Pane intentionally occurs twice inside each Left/Right menu:
  once under Go To and once at the pane-menu root. Tests must allow exactly those
  two occurrences and reject a third.
- Moving a command does not change its enablement, check/radio state, prompt,
  mutation authority, or undo/cancellation behavior.
- Labels that imply a pane or dialog use the established conventions:
  checkable surfaces use “Show ...” or “... Pane”; commands that open a dialog
  end in an ellipsis; immediate actions do not.
- Static labels remain resource-owned in every base and satellite `.rc` file.
- No runtime menu builder introduces hardcoded user-visible English.

## 7. Exact RedSalamander main-menu target

Top-level order remains:

`Left | Files | Edit | Commands | Plugins | View | Right | Help`

Help remains right-justified. Left and Right remain symmetric. Entries shown as
`<dynamic ...>` keep their owning runtime order.

### 7.1 Left / Right

Apply this identical shape with pane-specific command IDs:

```text
Change Drive
Go to >
  Back
  Forward
  Parent Directory
  Root Directory
  Path from Other Pane
  ---
  Hot Paths...
  <dynamic hot paths>
  ---
  <dynamic history>
---
Brief
Detailed
Extra Detailed
Thumbnails
Preview Pane
Terminal Pane
---
Sort By >
  None
  Name
  Extension
  Time
  Size
  Attributes
Show >
  Hidden Files
  System Files
  File Extensions
  ---
  Filter Bar
  Navigation Bar
  Status Bar
---
Refresh
Filter...
---
Maximize/Restore Pane
Swap Panes
Path from Other Pane
<debug overlay submenu in debug builds only>
```

Rules:

- Keep Path from Other Pane both under Go To and at the root after Swap Panes.
  This is an intentional duplicate, not cleanup debt.
- Add Terminal Pane immediately after Preview Pane in both pane menus. It maps
  to the existing `cmd/pane/openCommandShell`, displays the effective `Alt+7`
  binding, and uses the named pane as the source so the terminal opens/reuses in
  the physical opposite pane even when focus changed while the menu was open.
- Refresh and Filter stay together as content-state actions.
- Maximize/Restore and Swap stay together as pane-layout actions.
- Preserve the current sort/show radio/check state and all dynamic Go To items.

### 7.2 Files

```text
Open / Execute
View
Alternate View
View With >
  <dynamic file actions>
Edit
Alternate Edit
Edit With >
  <dynamic file actions>
---
New >
  Folder...
  Edit New File...
  ---
  <dynamic ShellNew entries>
---
Rename...
Batch Rename...
Change Case...
Change Attributes...
---
Copy...
Move/Rename...
---
Pack...
Unpack...
---
Delete...
Move to Recycle Bin
Delete Permanently...
---
Properties
Security...
Shell Context Menu >
  Selected Item...
  Current Folder...
---
Exit
```

Rules:

- Move View Width out of Files.
- Keep Create Directory only as New > Folder; F7 remains unchanged.
- Retain Change Case (`Ctrl+F7`, `cmd/pane/changeCase`) and Change Attributes
  (`Ctrl+F8`, `cmd/pane/changeAttributes`) as the metadata/transformation group
  immediately after Rename and Batch Rename. Preserve their dialogs, target
  rules, mutation behavior, cancellation, and tests.
- Flatten the current Delete submenu because it contains only two closely related
  actions and Permanent Delete already sits beside it.
- Keep Shell context invocation explicit so users can distinguish selected-item
  and current-folder targets.

### 7.3 Edit

```text
Cut
Copy
Paste
Paste Shortcut
---
Copy as Text >
  Path + Name
  Name
  Path
  UNC Path + Name
---
Select...
Unselect...
Invert Selection
Select All
Unselect All
Select Next
Select + Calculate Directory Size + Next
Advanced Selection >
  Save Selection
  Restore Selection...
  ---
  Select Same Extensions
  Unselect Same Extensions
  ---
  Select Same Names
  Unselect Same Names
  ---
  Hide Selected Names
  Hide Unselected Names
  Show Hidden Names
  ---
  Go to Previous Selected Name
  Go to Next Selected Name
```

Rules:

- Rename **Advanced** to **Advanced Selection**.
- Remove the root Restore Selection duplicate.
- The four copy-text commands retain their stable IDs and shortcuts; their submenu
  reduces the repeated root rows while keeping one predictable home.

### 7.4 Commands

```text
Change Directory...
Find Files and Directories...
Quick Search
---
Compare Directories...
Calculate Occupied Space
Make File List...
Go to Shortcut or Link Target
---
List Opened Files
Show Folder History
Open Active Pane Menu
Command Palette...
---
Connections >
  Connections Manager...
  Connect Network Drive...
  Disconnect...
  Shared Directories...
---
Terminal >
  Command Shell Window
  ---
  Insert Current Directory
  Insert Focused Item
  Insert Full Path
---
User Menu >
  <dynamic configured actions>
Open in File Explorer >
  Current Folder
  ---
  <known folders>
---
Reread Associations
Preferences...
```

Rules:

- Remove Create Directory from Commands.
- Move Quick Search next to Find and Change Directory.
- Add Command Palette (`Ctrl+Shift+P`) as the classic-menu doorway to the full
  canonical command catalog; map it to the existing `cmd/app/commandPalette`.
- Use the existing stable `cmd/pane/connections` for Connections Manager.
- Change Command Shell Window to map to the existing
  `cmd/terminal/openFloatingWindow`. From FolderView it displays and uses the new
  `Ctrl+Alt+T` default and opens/activates the floating terminal window without
  altering the embedded Terminal Pane state. Terminal-scope `Ctrl+Shift+N`
  remains an additional context-specific binding for the same stable command.
- Do not reuse `cmd/pane/openCommandShell` for the Commands row: it remains the
  embedded opposite-pane Terminal Pane command with `Alt+7` and FolderView
  `Ctrl+Shift+T`.
- Add the missing static Insert Full Path row and map it to the existing command.
- Move Preferences here as the last application-configuration command.

### 7.5 Plugins

```text
Plugin Manager...
---
<dynamic installed-plugin actions, or disabled Empty placeholder>
```

No semantic change is required.

### 7.6 View

```text
Theme >
  High Contrast (System)
  ---
  System
  Light
  Dark
  Rainbow
  High Contrast (App)
  ---
  Previous Theme
  Next Theme
  ---
  <dynamic themes>
---
Toggle Full Screen
View Width...
---
Window Menu
Switch Pane Focus
---
Failed Operations Pane
Show Function Bar
Show Menu Bar
```

Rules:

- Move View Width here.
- Use checkmarks consistently for Failed Operations Pane, Function Bar, and Menu
  Bar and use labels that describe the checked state.
- Preferences no longer appears under View.

### 7.7 Help

```text
Display Shortcuts...
External Help
---
About...
```

No semantic change is required; normalize spacing before the ellipsis in About.

### 7.8 Command and feature coverage policy

Classic menus are not required to contain every canonical leaf command. Doing so
would expose terminal navigation, numbered tab selection, parameterized commands,
provider-owned actions, and internal helpers as permanent static rows and would
make the menus less usable. The coverage target is instead:

- **100% canonical-command classification:** every registry command is classified
  as main menu, pane/context/dynamic menu, shortcut/Function Bar, palette-only,
  or intentionally internal/parameterized, with a reason for the latter two;
- **100% user-facing feature-family reachability:** every product feature family
  has at least one discoverable classic/contextual menu entry, or a reviewed
  palette-only rationale;
- **100% palette parity:** each Folder/Terminal Command Palette projection shows
  exactly one row per context-eligible palette-visible canonical command, every
  palette-visible command is reachable from at least one projection, and neither
  projection contains orphan/duplicate rows;
- **zero unclassified additions:** adding a registry command fails the Commands
  selftest until its surface classification is supplied.

Implement this as a table-driven Commands selftest/coverage report tied to
`CommandRegistry`, runtime menu traversal, Function Bar/shortcut catalogs, and
Command Palette enumeration. Do not claim 100% coverage from incomparable raw
counts of `cmd/*` strings and MENUITEM resource IDs. M1 must establish a green
baseline or mark each unresolved command `[blocked]` with its missing owner or
product decision.

## 8. Exact context-menu targets

### 8.1 FolderView item context

Create a dedicated static localized resource, proposed
`IDR_FOLDERVIEW_ITEM_CONTEXT`, with this base order:

```text
Open
Open With...
Calculate Occupied Space
---
Cut
Copy
Paste
---
Move...
Delete...
Rename...
<artifact submenu appended after a separator when applicable>
---
Properties
<debug overlay submenu in debug builds only>
```

Implementation requirements:

- Right-clicking a selected item applies selection-aware commands to the current
  selection; right-clicking an unselected item first makes it the sole target.
- Add Cut through the existing stable clipboard command and snapshot validation.
- Retain the current selection/current-target snapshot revalidation before posting
  a command; include Cut in the selection-aware set.
- Enable Paste from actual destination capability/clipboard state, not merely
  `CF_HDROP` presence if provider paste supports another advertised format.
- Keep Properties last among normal item actions. Insert artifact actions as a
  separate, clearly named submenu immediately before the Properties group and
  avoid duplicate adjacent separators.

### 8.2 FolderView background context

Create a separate static localized resource, proposed
`IDR_FOLDERVIEW_BACKGROUND_CONTEXT`:

```text
Paste
New >
  Folder...
  Edit New File...
  ---
  <dynamic ShellNew entries>
---
Refresh
Calculate Occupied Space
<debug overlay submenu in debug builds only>
```

Implementation requirements:

- Background invocation retains current-item focus ownership and clears selection
  exactly as the current input contract requires, but does not show disabled Open,
  Rename, Delete, or Properties rows.
- Reuse the existing Files > New dynamic builder; do not duplicate ShellNew
  enumeration or hardcode labels in C++.
- Keyboard context invocation remains item-targeted when a current item exists;
  pointer background invocation uses the background resource.

### 8.3 Find Files result context

Use one action list:

```text
Open
Go to Containing Folder
---
View
Alternate View
Edit
Alternate Edit
---
Copy
Cut
Copy To...
Move To...
---
Delete...
Delete Permanently...
```

Target semantics:

- Right-click on a selected row preserves and targets the selection.
- Right-click on an unselected row selects that row only and targets it.
- Keyboard invocation targets the current selection or the focused row when the
  selection is empty.
- Remove the duplicated “This item” and “Selection (N)” sections and their extra
  separators. Keep action enablement capability-aware.

### 8.4 Context surfaces intentionally retained

- Navigation provider/drive menu: retain Connections before Go To and Go To
  immediately before the drive list.
- Compare menu bar: retain **Compare** = Options, Rescan, separator, Show
  Identical Items, separator, Restore Differences Selection, Invert Differences
  Selection, separator, Close; retain **View** = Brief, Detailed, Extra Detailed.
- Compare and pane sort menus: retain None, Name, Extension, Time, Size,
  Attributes with one checked radio item.
- Batch Rename preview menu: retain Copy Original Name, Copy New Name, Copy
  Source Path, separator, Reveal in Active Pane, separator, Copy Preview Rows,
  Copy Execution Report, Copy Undo Plan, with capability/state enablement.
- Terminal menus: retain Copy, Paste, Select All, separator, Find, Suggestions,
  Close where supported; the session menu remains provider/session owned.
- ViewerSpace treemap context: retain Focus in Pane, Zoom In, Zoom Out before the
  cloned host context; normalize separators after the FolderView split.
- Shell `IContextMenu`: do not inspect or reorder Shell-owned items.

## 9. Viewer and secondary-application target

### 9.1 Common viewer File skeleton

Use this capability-based order; omit unsupported rows without leaving empty
groups:

```text
Open...                     (when supported)
Save As... / Export >       (when supported)
Refresh
Other Files >               (when supported)
  Previous File
  Next File
  First File
  Last File
---
Exit
```

Apply it as follows:

- **ViewerText:** Open, Save As, Refresh, Other Files, Exit.
- **ViewerWeb:** Save As, Refresh, Other Files, Exit.
- **ViewerPE:** Export Text, Export Markdown, Refresh, Other Files, Exit.
- **ViewerImgRaw:** Refresh, Export, Other Files, Exit; remove Other Files as a
  top-level category.
- **ViewerSpace:** retain Refresh, Up, Exit because it navigates the treemap
  hierarchy and has no sibling-file capability contract.

### 9.2 ViewerText

Target View order:

```text
Text
Hex
---
Side by Side Diff
Inline Diff
Show Unchanged Text
Next Difference
Previous Difference
---
Line Numbers
Wrap
---
Go to Top
Go to Bottom
Go to Offset...
```

Target Encoding order:

```text
Next Encoding
Previous Encoding
---
UTF-8
UTF-8 with BOM
UTF-16 LE
UTF-16 BE
Windows-1252
System ANSI
---
More Encodings...
---
Save Encoding >
  Keep Original Encoding
  UTF-8
  UTF-8 with BOM
  UTF-16 LE
  UTF-16 BE
```

Requirements:

- **More Encodings...** opens a localized searchable picker backed by the same
  canonical codepage catalog; it is not a second catalog.
- Filter by codepage number, canonical name, and localized/common alias.
- Keep the current encoding selected and visible; Enter applies, Escape cancels,
  arrow/page navigation works, and no choice changes the document until commit.
- Separate display decoding from save-conversion policy. Preserve existing
  warning/round-trip behavior and update `Plugins_ViewerText.md` before code.
- Preserve F7 or other existing diff-navigation shortcuts even if labels lose
  the repetitive `Diff /` prefix.

### 9.3 ViewerImgRaw

- Keep Fit, Actual Size, and Toggle Fit/Actual together.
- Keep zoom commands together, then Transform, Adjust, Source, and EXIF.
- Rename ambiguous `+`/`-` adjustment labels to **Increase/Decrease Brightness**,
  **Increase/Decrease Contrast**, and equivalent explicit names while retaining
  their command IDs and shortcuts.
- Use distinct access keys for Rotate Clockwise and Rotate Counterclockwise.

### 9.4 ViewerPE

- Retain the minimal File/View/Help categories and current capability set.
- Apply the common File order and unique sibling access keys.
- Do not add menus for unsupported editing or navigation behavior.

### 9.5 ViewerSpace

- Retain its minimal File menu and Up command.
- Preserve Focus in Pane, Zoom In, Zoom Out as the treemap-specific prefix to the
  host context menu.
- After FolderView context splitting, clone the correct item/background resource
  and prove no doubled or trailing separators.

### 9.6 ViewerWeb

- Keep Search commands under Search and document presentation under View.
- Move DevTools from View to Tools beside Copy HTML/Markdown and other tooling.
- Retain mode-specific enablement for Expand All, Collapse All, and Markdown
  source/display toggles; do not show unsupported rows as permanently disabled.

### 9.7 RedSalamanderMonitor

Top-level order remains:

`File | Edit | View | Options | Help`

Options target:

```text
Auto Scroll
Show IDs
Always on Top
---
Active Filter >
  Error
  Warning
  Information
  Performance
  Debug
  Trace
  ---
  Errors only
  Errors and warnings
  Errors, performance, and debug
  All message types
```

Keep Toolbar, Line Numbers, and Theme under View. Normalize access keys in View,
Theme, Options, and Active Filter.

### 9.8 RedConfigure

The removal gate did not pass during this package. Retain
`IDR_REDCONFIGURE_MAINMENU` and File > Exit because I6 still owns overlapping
RedConfigure localization work and the plan did not establish that no automated
or external caller depends on `IDM_REDCONFIGURE_EXIT`.

The retained command is not a second lifecycle: File > Exit and `WM_CLOSE` both
route through the same close/save path. A future removal package may stop passing
the one-item menu to the root window only after all of these are proven:

- system Close and Alt+F4 close the window through the same save/dirty-confirm
  contract as the current Exit command;
- Undo and Redo buttons remain keyboard-reachable and expose accurate accessible
  names, enabled state, and tooltips; if standard Undo/Redo shortcuts are added,
  the tooltips and authoritative UI spec must show them;
- no automated or external caller depends on `IDM_REDCONFIGURE_EXIT`;
- base and satellite resources remove the menu together;
- I6 has reconciled its in-flight RedConfigure localization work.

The current STOP result is therefore satisfied by retaining File > Exit. Do not
invent placeholder menus merely to satisfy symmetry.

## 10. Resource and localization contract

### 10.1 Structural parity

Extend the existing resource-localization Pester contract rather than adding a
new public Tool command. The parser must compare, for every base/satellite menu:

- resource ID and MENU versus MENUEX kind;
- popup nesting and sibling order;
- command IDs and separators;
- checked/radio/default/disabled flags that affect behavior;
- intentional build-configuration blocks;
- absence of satellite-only commands and absence of base-only commands.

Labels may differ by locale. Disabled dynamic placeholders may differ only when
explicitly reviewed. The test must report owner, resource ID, popup path, expected
token, actual token, and source line.

At minimum, cover these base owners and every matching `Lang/<culture>` tree:

- RedSalamander
- RedSalamanderMonitor
- RedConfigure
- ViewerText
- ViewerImgRaw
- ViewerPE
- ViewerSpace
- ViewerWeb

The first test-driven repair is MNU-01: all four RedSalamander satellites gain
`IDM_PANE_BATCH_RENAME` in the same Files position as English, with reviewed
localized text and a sibling-unique access key.

### 10.2 Access-key contract

For each separator-delimited static sibling group:

- require exactly one access key on every actionable enabled text row;
- require the key to be unique case-insensitively within that group;
- accept both `&X` and localized suffix conventions such as Japanese `(&X)`;
- treat escaped `&&` as a literal ampersand;
- exclude separators, disabled `(Empty)` placeholders, numeric-only rows, and
  explicitly dynamic text;
- report duplicate key, both labels, popup path, culture, and source lines.

Do not mechanically copy English access keys into translated labels. Resolve the
label and access key together, and obtain human review when the translated term
changes. The contract prevents structural drift; it does not certify translation
quality.

### 10.3 Runtime contract

Add a table-driven Commands case proposed as
`main_menu_information_architecture` that loads the effective menu and validates:

- exact top-level order and right-justified Help;
- exact static command/submenu/separator sequence from section 7;
- all dynamic insertion anchors and empty states;
- each actionable app command maps to the expected stable `cmd/*` ID;
- each stable command shown in the menu dispatches through its existing path;
- no same-command duplicates inside one top-level menu except a reviewed
  allowlist;
- exactly two Path from Other Pane placements in each Left/Right menu and no
  additional occurrence;
- Left/Right Terminal Pane maps to `cmd/pane/openCommandShell`, displays `Alt+7`,
  and uses the named source pane rather than late focus state;
- Commands > Terminal > Command Shell Window maps to
  `cmd/terminal/openFloatingWindow` and displays the effective FolderView
  `Ctrl+Alt+T` binding;
- check/radio/enablement state for pane display, View surfaces, themes, and
  capability-dependent commands;
- visible shortcut text reflects the current effective binding;
- English and each available satellite load the same command tree.

Keep exact focused cases for theme and Help order if they provide clearer failure
diagnostics, but remove assertions that require the accidental duplicates retired
by M2.

Add a second table-driven case proposed as `command_surface_coverage_contract`.
It enumerates the canonical registry and requires each entry to have one or more
reviewed surface classifications from section 7.8. It also asserts Command
Palette one-row-per-command parity and emits a reviewable summary grouped by
main menu, contextual/dynamic menu, shortcut/Function Bar, palette-only, and
internal/parameterized. A new registry entry without classification is a failure.

## 11. Implementation packages

### 11.1 M0 — baseline and drift reconciliation

1. Run the drift command.
2. Search for new menu construction with:

   ```powershell
   rg -n --glob '*.rc' '^\s*[A-Za-z0-9_]+\s+MENU(EX)?\b|^\s*POPUP\b|^\s*MENUITEM\b' RedSalamander RedSalamanderMonitor RedConfigure Plugins
   rg -n --glob '*.cpp' 'CreatePopupMenu|AppendMenuW?|InsertMenuW?|InsertMenuItemW?|TrackPopupMenu|ContextMenu::Show' RedSalamander RedSalamanderMonitor RedConfigure Plugins
   ```

3. Capture current base/satellite trees in a test artifact or review note.
4. Verify that no in-scope file has overlapping uncommitted changes. Preserve
   unrelated user edits and STOP on an unresolvable overlap.

Commit boundary: audit/baseline only; no production change.

### 11.2 M1 — contracts and P0 repairs

Files:

- `Tools/Tests/ResourceLocalizationContracts.Tests.ps1`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`
- `RedSalamander/RedSalamander.rc`
- `RedSalamander/Lang/*/*.rc`
- `RedSalamander/resource.h`
- `RedSalamander/CommandRegistry.cpp`
- `RedSalamander/RedSalamander.cpp`
- `Specs/UI/UI_CommandMenuKeyboard.md`

Order:

1. Add a failing menu structural-parity test that identifies the missing Batch
   Rename command in every RedSalamander satellite.
2. Repair all four satellite menu trees and access keys.
3. Add the WM command ID and menu/registry mapping for Insert Full Path in
   Terminal; route to the existing stable-command behavior.
4. Add Connections Manager and Failed Operations Pane to the authoritative
   command/menu contract and fix obsolete test wording.
5. Add the full runtime menu-tree test using the current tree as an intermediate
   baseline; update it with M2 in the next commit.
6. Add the command-surface coverage ledger and establish the classified baseline
   before making any 100% feature-reachability claim.

Verification after each step:

```powershell
Invoke-Pester -Path .\Tools\Tests\ResourceLocalizationContracts.Tests.ps1
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter main_menu_information_architecture
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter command_surface_coverage_contract
```

### 11.3 M2 — main-menu information architecture

Files:

- base and all RedSalamander satellite `.rc` files;
- `RedSalamander/resource.h`, `CommandRegistry.cpp`, `ShortcutDefaults.cpp`, and
  the existing shortcut-default migration owner;
- `RedSalamander/RedSalamander.cpp` for pane-specific Terminal Pane dispatch,
  floating-window command mapping, dynamic submenu anchors, and menu state;
- affected Commands tests;
- `Specs/UI/UI_CommandMenuKeyboard.md` and
  `Specs/Terminal/Terminal_EmbeddedPlugin.md`.

Implementation sequence:

1. Normalize Left/Right while retaining both Path from Other Pane placements.
2. Allocate pane-specific Left/Right Terminal Pane WM IDs. Map both to
   `cmd/pane/openCommandShell` for shortcut display, but dispatch Left with
   `Pane::Left` and Right with `Pane::Right`; do not derive the source from focus
   after a named pane-menu invocation.
3. Allocate a static Commands-menu WM ID for
   `cmd/terminal/openFloatingWindow`, route it through the existing stable-command
   dispatcher, and label it Command Shell Window in every locale.
4. Reassign the FolderView `Ctrl+Alt+T` default to the floating command. Add a
   narrow migration that changes only the legacy default binding; preserve every
   user-customized binding. Keep `Alt+7` and FolderView `Ctrl+Shift+T` on the
   embedded command and Terminal-scope `Ctrl+Shift+N` on the floating command.
5. Normalize Files, retaining Change Case/Change Attributes, flattening the
   Delete group, and keeping New as the sole Create Directory menu home.
6. Normalize Edit and create Copy as Text / Advanced Selection submenus.
7. Normalize Commands, including Command Palette, Connections, floating Terminal,
   and Preferences.
8. Normalize View labels/positions and preserve checks.
9. Normalize access keys in every touched sibling list and locale.
10. Update table-driven tree tests and individual command-state tests in the same
   commit as each resource move.

Every commit must build resources and pass the exact tree test. Never land an
English tree change without all four satellite structures in the same commit.
The terminal split reuses the existing embedded `terminal.open_to_visible_us`
instrumentation and floating-terminal lifecycle tests; it does not authorize a
terminal-engine redesign or duplicate metric family.

### 11.4 M3 — FolderView and Find context menus

Files:

- `RedSalamander/RedSalamander.rc`, satellites, `resource.h`;
- `RedSalamander/FolderView.Menus.cpp` and existing declarations only as needed;
- the existing Files > New builder owner if extraction is necessary;
- `RedSalamander/FindFilesWindow.cpp`;
- focused Commands tests;
- `Specs/UI/UI_FolderView.md`, `Specs/UI/UI_FindFilesWindow.md`, and
  `Specs/UI/UI_CommandMenuKeyboard.md`.

Implementation sequence:

1. Extract/reuse one New-menu population helper with explicit ownership and no
   hardcoded labels. Do not move it to `Common/` unless its semantics genuinely
   serve another executable and `Core_SharedHelpers.md` is updated.
2. Add static item/background resources and load the correct one at invocation.
3. Expand snapshot-safe command classification for Cut and background commands.
4. Update the existing empty-background case to assert absence of item commands,
   correct background order, and successful Paste/New/Refresh routing.
5. Update item-context tests for selected versus unselected right-click targets,
   multi-selection Cut/Copy/Move/Delete, Properties-last order, and stale-target
   rejection.
6. Replace Find's duplicated action blocks with the one-target model and add
   selected/unselected/keyboard/multi-selection cases.
7. Recheck ViewerSpace's cloned FolderView menu and artifact-separator behavior.

Proposed focused case names:

- `folder_view_item_context_menu_contract`
- `folder_view_background_context_menu_contract`
- `find_files_context_target_contract`
- `viewer_space_context_menu_composition`

### 11.5 M4 — viewers and utility apps

Land one executable/plugin per commit so localization and routing failures remain
small and revertible:

1. ViewerImgRaw File hierarchy and explicit adjustment labels.
2. ViewerWeb DevTools placement and common File order.
3. ViewerPE common File order/access keys.
4. ViewerSpace common checks and context composition.
5. ViewerText View order and Save Encoding separation, without the searchable
   picker yet.
6. Monitor Options order/preset labels.
7. RedConfigure one-item menu removal after the section 9.8 gate and I6
   coordination.

Each commit updates its base resource, all satellites, command routing tests, and
owning plugin/core specification together.

### 11.6 M5 — access-key closure

1. Run the static contract over every owner/culture.
2. Resolve all duplicate and missing actionable access keys.
3. Add runtime keyboard-navigation cases for at least the RedSalamander main bar,
   Monitor Options, ViewerText View/Encoding, and ViewerImgRaw Transform/Adjust.
4. Verify that custom DxUi conversion retains access-key behavior even when the
   display renderer hides ampersand markers.
5. Obtain human review for changed translated labels; mark a locale `[blocked]`
   if no qualified reviewer is available rather than silently shipping English
   or an unreviewed machine translation as final.

### 11.7 M6 — ViewerText searchable encoding picker

This is the only package in this plan that necessarily adds a new interactive
surface and repeated filter/render path, so the performance contract applies
from its first commit.

Protected scenario:

`viewer_text_encoding_picker_open_and_filter_full_catalog`

Metric family (one row per user-visible gesture, never per codepage):

- `viewer.encoding_picker.catalog_load_us`
- `viewer.encoding_picker.open_to_ready_us`
- `viewer.encoding_picker.filter_us`
- `viewer.encoding_picker.catalog_items`
- `viewer.encoding_picker.filtered_items`

Required sequence:

1. Because the retired encoding cascade had no comparable instrumentation,
   archive the first same-machine test-enabled Release picker run as an
   invariant/scaling baseline; do not claim historical before/after latency.
2. Build the picker from the canonical codepage catalog; do not duplicate names,
   aliases, or codepage metadata.
3. Add deterministic correctness coverage for empty query, numeric query, alias,
   no result, current-selection visibility, Enter, Escape, rapid query replacement,
   high-DPI layout, theme change, and close/unload during filtering.
4. Run the identical catalog fixture for baseline and candidate. Establish a
   baseline rather than inventing a hard percentile budget from one pair.
5. Archive `results.json`, `trace.txt`, and `perf_metrics.jsonl` under
   `Specs/TestRuns/<MachineHash>/Viewers/<RunId>/`, validate the archive, and
   report build flavor, sample counts, same-machine status, and caveats.
6. Update `Specs/Plugins/Plugins_ViewerText.md` and
   `Specs/Testing/Testing_PerformanceValidation.md` with the durable scenario and
   metrics if the surface lands.

Accepted first baseline: `Specs/TestRuns/4cb089111a23/Viewers/20260828_192400_menu_encoding_picker_release/`.
Five independent test-enabled x64 Release processes exercised 140 catalog rows.
Open-to-ready p95 was 25,150 us against the 50,000 us ceiling; filter p95 was
12,112 us against the 16,670 us ceiling. The archive contract passed for
`results.json`, `trace.txt`, and `perf_metrics.jsonl`. This is the first accepted
invariant/scaling baseline and makes no historical before/after claim.

Static resource reordering alone does not justify adding instrumentation to every
menu-open path. If a static change unexpectedly affects menu-session latency or
paint, STOP and add the smallest attributable metric and deterministic Commands
case before claiming completion.

## 12. Test plan

### 12.1 Static/resource tests

- base/satellite MENU/MENUEX structural parity for every owner;
- unique/missing access keys per sibling and culture;
- valid resource IDs with no accidental numeric collision;
- no hardcoded static user-facing menu labels in runtime builders;
- no satellite-only or base-only actionable menu commands;
- no trailing, leading, or doubled separators after conditional/dynamic removal.

### 12.2 RedSalamander runtime tests

- exact top-level and full static subtree order;
- dynamic insertion anchors and empty placeholders;
- stable command mapping, shortcut display, check/radio/enabled state;
- left/right symmetry after command-ID normalization;
- exactly two reviewed Path from Other Pane placements per pane menu, and only
  reviewed duplicate command placements elsewhere;
- Left/Right Terminal Pane uses the named source, opens/reuses the physical
  opposite pane, and retains `Alt+7` plus FolderView `Ctrl+Shift+T`;
- Commands > Terminal > Command Shell Window creates/activates the floating
  terminal, leaves embedded pane state unchanged, and uses FolderView
  `Ctrl+Alt+T`;
- legacy-default shortcut migration changes only the old `Ctrl+Alt+T` binding
  and preserves user overrides; Terminal-scope `Ctrl+Shift+N` remains valid;
- Change Case and Change Attributes remain present in every locale, display
  `Ctrl+F7` / `Ctrl+F8`, and dispatch through their existing tested behavior;
- Command Palette is present in Commands with `Ctrl+Shift+P`, and the coverage
  ledger classifies every registry entry with one palette row per canonical command;
- Batch Rename present under every loaded locale;
- Insert Full Path menu activation dispatches the same behavior as
  `Ctrl+Shift+Enter` without sending Enter;
- Connections Manager and Failed Operations Pane dispatch/state;
- FolderView item/background target semantics and menu shape;
- Find selected/unselected/keyboard target semantics;
- Compare, navigation, Terminal, Batch Rename, and ViewerSpace retained shapes.

### 12.3 Viewer/utility tests

- every File menu follows the capability-based order and omits unsupported
  groups cleanly;
- ViewerText View/Encoding exact order and state;
- ViewerImgRaw transform/adjust routing after label/access-key changes;
- ViewerWeb DevTools routes from Tools;
- Monitor Options state and filter presets;
- RedConfigure File > Exit and `WM_CLOSE` share the close/save lifecycle; retain
  Undo/Redo keyboard reachability and accessibility with the menu present.

### 12.4 Manual accessibility matrix

Run after automated coverage, not instead of it:

- keyboard only: Alt activation, arrow traversal, access keys, Enter, Escape;
- mouse and touchpad: target selection and context positioning;
- 100%, 150%, and 200% DPI where available;
- light, dark, application high contrast, and system high contrast;
- English, Czech, French, Japanese, and Slovak;
- persistent and temporarily revealed main menu bar;
- single monitor and a secondary monitor with different DPI if available;
- screen reader/UIA name, enabled, checked, and submenu states for changed
  RedConfigure/Monitor/main-menu rows.

## 13. Verification commands

Focused iteration:

```powershell
Invoke-Pester -Path .\Tools\Tests\ResourceLocalizationContracts.Tests.ps1
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter main_menu_information_architecture
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter command_surface_coverage_contract
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter folder_view_empty_background_keeps_current
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_pane_open_command_shell_prefers_windows_terminal
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter terminal_shortcut_floating_singleton_tabs_restore_and_close
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_pane_changeCase
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_pane_changeAttributes_applies_attributes_removes_streams_and_reports
```

The resource-localization Pester contract owns the exact static item/background
menu shapes; `folder_view_empty_background_keeps_current` adds the live pointer
versus keyboard target and background-shape proof. The two retired
`folder_view_*_context_menu_contract` recipe names never became registered
selftest cases and must not be used as closeout evidence.

ViewerText performance package:

```powershell
try {
    $env:RSBuildEnableTests = 'true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
```

Use the viewer test entrypoint selected by the existing viewer harness and the
exact case filter introduced for
`viewer_text_encoding_picker_open_and_filter_full_catalog`:

```powershell
.\.build\x64\Release\ViewerPETests.exe viewer_text_encoding_picker_open_and_filter_full_catalog
```

Final closeout:

```powershell
.\build.ps1 -Rebuild -MaxCpuCount 4
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Get-TestInventory.ps1 -Format Json
.\Tools\Test-TestRunArchive.ps1 -Inventory
git diff --check
```

Record the exact commit, build receipt, Full run ID, per-suite passed/failed/
skipped counts, archive path, and any environment-gated skip. A Commands-only,
Affected, or Resume run is useful iteration evidence but does not replace the
final Fresh Full.

## 14. Authoritative specification updates

Update these in the same package that changes their durable behavior:

- `Specs/UI/UI_CommandMenuKeyboard.md` — exact main tree, stable IDs, shortcut
  display, intentional duplicate policy, command-surface coverage policy, access
  keys, Command Palette, Connections/Failed Operations, and the embedded/floating
  terminal split.
- `Specs/Terminal/Terminal_EmbeddedPlugin.md` — Terminal Pane naming and named
  source/opposite-host behavior, floating-window command separation, shortcut
  scopes/default migration, retained metrics, and lifecycle tests.
- `Specs/UI/UI_FolderView.md` — item/background target semantics and composition.
- `Specs/UI/UI_FindFilesWindow.md` — one-target result context model.
- `Specs/Core/Core_CompareDirectories.md` — retained menu/sort order if the
  whole-surface test makes it durable.
- `Specs/UI/UI_BatchRenameWindow.md` — retained preview-context order if the
  whole-surface test makes it durable.
- `Specs/Plugins/Plugins_ViewerText.md` — File/View/Encoding target, picker, save
  encoding, tests, and performance scenario.
- `Specs/Plugins/Plugins_ViewerImgRaw.md` — File hierarchy and adjustment labels.
- `Specs/Plugins/Plugins_ViewerPE.md` — common File order.
- `Specs/Plugins/Plugins_ViewerSpace.md` — retained minimal File/context order.
- `Specs/Plugins/Plugins_ViewerWeb.md` — DevTools under Tools and common File order.
- `Specs/Core/Core_RedSalamanderMonitor.md` — Options/filter order and labels.
- `Specs/Core/Core_RedConfigure.md` and `Specs/UI/UI_RedConfigure.md` — no-menu
  close/Undo/Redo/accessibility contract if the removal gate passes.
- `Specs/Testing/Testing_SelfTests.md` — lasting runtime/resource menu-test
  contract if it establishes a reusable repository rule.
- `Specs/Testing/Testing_PerformanceValidation.md` — ViewerText picker metrics and
  archive path if M6 lands.

Do not leave the final menu trees or access-key/parity rules only in this WIP
file.

## 15. Git workflow and commit strategy

This plan does not authorize changing branches or moving the current dirty
checkout. An executor uses the task environment assigned by the user/runner and
must preserve unrelated working-tree changes. If the user explicitly requests a
new branch, create a `codex/menu-information-architecture` branch from the
drift-reviewed base in a clean worktree; do not transplant the current dirty
checkout or reset another owner's changes.

Use small reviewable commits in this order:

1. resource/menu contract tests;
2. P0 satellite and missing-command repairs;
3. main pane/Files tree, including Terminal Pane and retained Change commands;
4. main Edit/Commands/View tree, floating Command Shell Window, Command Palette,
   and shortcut-default migration;
5. FolderView context split;
6. Find result target simplification;
7. one viewer or utility executable per commit;
8. ViewerText picker instrumentation baseline;
9. ViewerText picker implementation/candidate evidence;
10. specs, final validation, and plan closeout.

Every production commit includes its affected base resource, all satellites,
focused tests, and owning authoritative spec. Do not batch unrelated viewer
changes into the main-menu commit. Preserve the current dirty worktree and never
overwrite unrelated user changes.

## 16. Definition of done

This plan is complete only when all of the following are true:

- every surface in section 4 is implemented as **change**, verified as **keep**,
  or explicitly marked externally owned with a tested host boundary;
- every base/satellite menu pair has identical command structure;
- Batch Rename exists in every supported RedSalamander locale;
- every documented menu command exists and every implemented stable menu command
  is documented, including Insert Full Path, Connections Manager, and Failed
  Operations Pane;
- the exact main-menu target in section 7 is green in runtime tests;
- both intentional Path from Other Pane placements remain in each pane menu;
  accidental duplicates in MNU-06 and MNU-07 are gone, with only reviewed
  duplicate placements remaining;
- Terminal Pane follows Preview Pane in Left/Right, uses `Alt+7`, and opens the
  terminal in the named pane's physical opposite;
- Commands > Terminal > Command Shell Window uses the floating-terminal stable
  command and the migrated FolderView `Ctrl+Alt+T` default without overwriting
  user shortcut customizations;
- Change Case and Change Attributes remain present, localized, shortcut-visible,
  routed, specified, and covered by their existing deterministic tests;
- Command Palette is reachable from Commands and the command-surface coverage
  ledger proves 100% classification, feature-family reachability, and palette
  parity, with no unreviewed internal/palette-only exceptions;
- FolderView background menus contain background commands rather than disabled
  item actions;
- Find Files renders one action list with deterministic targeting;
- viewer/Monitor/RedConfigure targets are implemented or an explicit STOP
  condition is recorded for the blocked package;
- all actionable static siblings have reviewed unique access keys in all
  supported locales;
- ViewerText picker correctness and same-machine performance evidence is archived
  if M6 lands;
- focused tests, full build, Fresh Full, spec/test inventory, archive inventory,
  and `git diff --check` are green;
- durable behavior is merged into the authoritative specs;
- this plan is moved to `Specs/Plans/Done/` and I16 is removed from the WIP index.

## 17. STOP conditions

Stop the affected package and request/reconcile direction if any of these occur:

- a proposed move changes stable `cmd/*` identity, user shortcut persistence, or
  Function Bar semantics, except for the explicitly reviewed legacy-default-only
  `Ctrl+Alt+T` migration;
- current authoritative behavior conflicts with this target and the conflict is
  a product decision rather than clear drift;
- a locale cannot be updated with reviewed text/access keys;
- a new static menu label would have to be hardcoded in C++;
- a context simplification weakens target snapshot/revalidation, capability
  checks, destructive confirmation, or provider authority;
- Shell/provider/user-defined command ordering would have to be copied or
  reinterpreted by the host;
- I1/I2/I4/I6/I10/I12/I14/I15 has overlapping active edits that cannot be cleanly
  sequenced;
- RedConfigure menu removal loses an accessible close, dirty-confirmation,
  Undo/Redo, or automation path;
- the ViewerText picker would duplicate the codepage catalog, block the UI on
  catalog/filter work, or lacks deterministic/archived performance validation;
- final validation requires killing an unrelated process, discarding user edits,
  weakening a test, or treating non-Fresh evidence as a Full closeout.

When a package stops, mark it `[blocked]` with the exact missing decision,
reviewer, owner, or evidence. Continue independent packages when their ownership
and validation remain sound.

## 18. Ongoing maintenance after closeout

- Run resource structural/access-key contracts in the canonical Tools Pester
  lane for every change to a base or satellite MENU/MENUEX resource.
- Require the owning authoritative spec and runtime menu-tree table to change in
  the same commit as a durable menu move.
- Require every new canonical command to declare a reviewed surface classification
  and keep each Command Palette projection at exactly one row per eligible
  palette-visible canonical command.
- Add new dynamic builders to the surface inventory and give them explicit
  insertion anchors, empty state, ownership, and separator tests.
- Review duplicate command placements through a small named allowlist; never
  weaken the contract globally to admit one exception.
- Keep the two Path from Other Pane placements and the embedded/floating terminal
  command split as explicit regression cases rather than relying on label-only
  review.
- Re-run the ViewerText picker scenario after catalog growth, picker rendering,
  localization, or filter-algorithm changes, and archive evidence when the change
  can affect responsiveness.
- Re-audit menu information architecture when a top-level menu reaches 25 rows,
  a context menu reaches 15 rows, or a new one-/two-item top-level category is
  proposed.
