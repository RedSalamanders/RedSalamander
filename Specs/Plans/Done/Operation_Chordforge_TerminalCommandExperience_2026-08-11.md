# Operation Chordforge — Terminal Command Experience

> **NON-NORMATIVE COMPLETED PLAN.** Operation Chordforge delivered a
> context-aware merge of Windows Terminal keyboard conventions with
> RedSalamander's embedded terminal, its shortcut helper, and searchable global
> command palette. Current behavior is owned by the authoritative specifications
> named below.

## Status

- **State:** COMPLETE
- **Priority:** P1 terminal input correctness and usability
- **Planned at:** `274cee199f5b192265dc837dcd3916294fd3d5c2`
- **Ownership boundary:** terminal-focused shortcut settings, input arbitration,
  physical key-position identity, the binding-centric F1 Helper, the global
  RedSalamander Command Palette, Terminal Find/history Suggestions, floating
  Terminal window/tabs, shell-exit lifecycle, and the pane/window commands required
  to provide an honest Windows Terminal-compatible experience
- **Coordination:** reconcile terminal input and ABI changes with
  `CodeReview_CrossDomain_Remediation_2026-08-10.md` before editing shared
  Terminal files; that plan owns existing correctness findings, while this plan
  owns only the new shortcut experience
- **Drift check:**

  ```powershell
  git diff 274cee199f5b192265dc837dcd3916294fd3d5c2..HEAD -- `
    Common/PlugInterfaces/Terminal.h Common/SettingsStore.h Common/Common/SettingsStore.cpp `
    Common/Keyboard* Plugins/Terminal Plugins/ViewerWeb Plugins/ViewerImgRaw `
    RedSalamander/CommandRegistry.* RedSalamander/ShortcutDefaults.cpp `
    RedSalamander/ShortcutManager.* RedSalamander/Preferences.Keyboard.* `
    RedSalamander/ShortcutsWindow.* RedSalamander/RedSalamander.cpp `
    RedSalamander/FolderWindow* RedSalamander/FloatingTerminalWindow* `
    RedSalamander/WindowPlacementPersistence.h Tests/PluginContractTests `
    RedSalamander/SelfTest/Commands Tools/Tests/TerminalPluginSourceContracts.Tests.ps1 `
    Specs/Terminal/Terminal_EmbeddedPlugin.md Specs/UI/UI_CommandMenuKeyboard.md `
    Specs/Core/Core_SettingsStore.md Specs/Core/Core_SharedHelpers.md `
    docs/KeyboardShortcuts.md docs/UserGuide.md docs/SettingsFile.md
  ```

## Continuation checkpoint — 2026-08-11

The implementation and normative documentation are saved in the current
unstaged, uncommitted worktree. Do not discard or overwrite unrelated dirty
changes. Resume from this plan; no external branch, stash, or commit was
created.

Completed and reverified at this checkpoint:

- `ITerminal::SetCallback(ITerminalEventCallback*, void*)` is on IID
  `47E1AC4C-022E-42C4-8E2C-0C2BBC687D31`, and the only callback signature is
  `OnTerminalEvent(const TerminalEvent* event, void* cookie) noexcept` with
  context last. `TerminalEvent` preserves its size while carrying the
  process-watcher's root-exit timestamp.
- The full shortcut/helper/palette/Preferences/Find/Suggestions/floating-window
  implementation described below is present. Startup restore is visible-first
  and incrementally hydrates every remembered tab in stable order.
- Floating-window `WM_NCDESTROY` now drains posted payloads before the shared
  DxUi host can consume teardown; final telemetry proves zero retained roots,
  tabs, Terminal instances, callbacks, and payloads.
- The Microsoft Terminal refresh is pinned to
  `b888cb7e4c3b0b21b7ed66c224bcdf7fa9ef6d9a`; the reviewed upstream total is
  still `48 + 12 + 8 = 68`.
- Terminal source contracts pass `27/27`. The focused Debug floating lifecycle
  case passed as runner
  `20260811T185755Z-32436-b27740023ae84451b365de4879c12d7e`.
- The definitive x64 test-enabled Release performance case passed `1/1` as
  runner `20260811T192050Z-97512-d3ee19ffa6494d449080a6f4c13556d8`.
  Its compact evidence is
  `Specs/TestRuns/4cb089111a23/Commands/20260811_192050_operation_chordforge_terminal_command_experience_release/`.
  Raw JSONL identity: 10,474,542 bytes, 28,069 rows, SHA-256
  `393b340fc5f4d238d87a08a0e1d614a5ef158b9728db3df396b291b4ea03132d`.
- `Tools/Test-TestRunArchive.ps1 -RunPath <archive>` passes the explicit
  pre-stage contract for all seven archived files.

Required resume/closeout sequence:

1. Run the eight `Show-PerfRuns.ps1 -FailOnQuality` commands recorded in the
   evidence README; the explicit archive-size/shape validation is already green.
2. Rebuild current Debug after the final event timestamp/teardown/load-metric
   changes and rerun the focused `terminal_shortcut_` Commands family plus
   Terminal source contracts, `PluginContractTests`, `SettingsSchemaTests`, and
   `DxUiTests`.
3. Run a normal production Release build without `RSBuildEnableTests`.
4. Run `Tools/Run-AllTests.ps1 -Suite Full`,
   `Tools/Get-SpecInventory.ps1 -FailOnFindings`,
   `Tools/Test-TestRunArchive.ps1 -Inventory`, and `git diff --check`.
5. Record those final run IDs/results here, change **State** to complete, and
   move this file to `Specs/Plans/Done/`. Do not move it before every gate is
   green and the authoritative specs continue to agree.

## Continuation checkpoint — 2026-08-12

Closeout advanced substantially, but the plan remains **ACTIVE** because the
repository-wide Full gate exposed two failures outside the focused Terminal
matrix. The durable continuation record is:

`Specs/TestRuns/4cb089111a23/Continuation/20260812_093152_operation_chordforge_closeout_checkpoint/README.md`

New green results:

- all eight archived performance quality analyzers passed;
- full Debug x64 rebuild passed (0 warnings, 0 errors), log
  `.build/logs/msbuild-20260812_084412_823.log`;
- source contracts 27/27, Terminal plugin debug selftests 56/56 plus unload
  quiet point, SettingsSchemaTests, and DxUiTests all passed;
- focused `terminal_shortcut_` Commands passed 10/10 as runner
  `20260812T065107Z-10160-8919ba1965f54dc3b09a51705d655eaa`;
- normal production Release x64 passed with zero errors, log
  `.build/logs/msbuild-20260812_085128_320.log`.

Full runner `20260812T065351Z-74552-6c75e0bca16346eba46902ea7be2e3af`
returned red: `Commands` lost its main window after the unrelated
`pane_view_options_preview_uses_configured_embedded_viewer_and_preserves_focus`
failure, cascading to 711 reported failures, and aggregate `ToolsPesterTests`
returned exit code 1 without retaining its individual failure list. File
Operations was investigated as an apparent hang but completed 118/0/20 with a
clean `FileOpsSelfTest: PASS`; there was no crash or new dump.

Resume with the isolated Commands preview-focus case and direct complete Pester
run recorded in the continuation README. Only after both are reconciled should
the Full gate, spec inventory, archive inventory, diff hygiene, final
requirement audit, and WIP-to-Done move proceed. No branch, stash, commit, or
stage was created; all work remains in the current dirty worktree.

## Final-closeout checkpoint — 2026-08-12 23:05 +02:00

Closeout state, green evidence, the active Commands runner, its first two
non-Terminal fixture failures, and the exact remaining Done sequence are saved
in:

`Specs/TestRuns/4cb089111a23/Continuation/20260812_230530_operation_chordforge_final_closeout_checkpoint/README.md`

This plan remains **ACTIVE**. Do not move it to `Specs/Plans/Done/` until the
complete Commands suite and the final repository Full gate are green.

## Closeout checkpoint — 2026-08-13 15:54 +02:00

Operation Chordforge is **COMPLETE**. The durable convergence record is:

`Specs/TestRuns/4cb089111a23/Continuation/20260813_155403_operation_chordforge_closeout_complete/README.md`

The final immutable-build Full runner
`20260813T124341Z-80664-e2122867560e4176aeab859cc69b5cf9` completed with
1,199 passed, 3 failed, 53 skipped, and zero disk-audit issues. Every Operation
Chordforge case was green. The three late failures were closed without
repeating the unchanged Full matrix:

- the Viewers theme-cycle fixture's ordinal category navigation was corrected
  to select `kPrefCategoryViewers` by stable identity, rebuilt with 0 warnings
  and 0 errors, and passed as focused runner
  `20260813T134259Z-43540-7185d94935ba4e46ab55b15f785ba60d`;
- the exact `RedSalamanderMonitorEtwLatency` Full command passed directly; and
- the complete tooling Pester bundle passed 695/0/5.

Broad Commands had already passed 843/0/2 on runner
`20260813T105343Z-15212-de54692d11ab4b95a594ff9c09066887`, and the focused
Terminal, File Operations, schema, localization, plugin-contract, DxUi,
performance, archive, and source-contract evidence is green in the continuation
archives referenced above. The authoritative Terminal, command/keyboard,
Preferences, settings, top-level-window, File Operations, shared-helper, user
guide, shortcut, and settings-file documents contain the lasting behavior.
Final specification-information-architecture tests passed 21/0, specification
inventory reported zero blocking findings, TestRuns archive inventory passed
for 505 files, and `git diff --check` reported no whitespace errors.

The closeout validation rule is source-impact-aware: a completed aggregate run
plus a successful rebuild and one exact revalidation of each failed gate is
sufficient when the remaining aggregate cases ran from the same immutable
binary and their sources were not changed. This avoids repeating already-green
cases without weakening changed-code coverage.

## Upstream review basis

This review uses the 68 bindings in the Microsoft-owned Windows Terminal
[`defaults.json`](https://github.com/microsoft/terminal/blob/b888cb7e4c3b0b21b7ed66c224bcdf7fa9ef6d9a/src/cascadia/TerminalSettingsModel/defaults.json)
at upstream commit `b888cb7e4c3b0b21b7ed66c224bcdf7fa9ef6d9a`,
refreshed on 2026-08-11, and the Microsoft Learn
[`actions` and key-binding contract](https://learn.microsoft.com/en-us/windows/terminal/customize-settings/actions).
The Learn contract explicitly supports unbinding a chord so the underlying
shell or terminal application receives it. RedSalamander needs that same
pass-through outcome as a first-class choice; silently consuming an unsupported
Windows Terminal chord is not compatibility.

The closeout refresh found no drift: upstream still contains exactly 68
bindings and the reviewed partition remains `48 + 12 + 8 = 68`. Microsoft Learn
still documents `unbound` as the way to let the application receive the chord.
No build or runtime path downloads or depends on the upstream file.

## Product decision

Adopt **Windows Terminal familiarity where RedSalamander has the same user
intent**, adapt it where the fixed two-pane/content-tab model has a clear
equivalent, and pass the chord through to Ghostty/the shell everywhere else.

The merge has four invariants:

1. Existing Folder/Navigation shortcuts keep their current meanings. The only
   folder-scope additions are non-conflicting pane-content navigation chords
   explicitly listed in this plan.
2. A new terminal shortcut scope applies only while a live embedded or floating
   Terminal child owns input. Duplicate chords across Folder and Terminal
   scopes are intentional and are not conflicts.
3. Terminal security/exclusive-input state wins over every new shortcut.
   Shell/TUI input wins when a mapped action is not eligible in the current
   terminal state.
4. An unsupported, deferred, unknown, disabled, or explicitly passed-through
   chord reaches the child unchanged. Do not synthesize shell text or escape
   sequences to imitate a missing host action.

This preserves the current Salamander integration commands:

| Existing context | Chord | Required behavior |
|---|---|---|
| Folder/Navigation | `Ctrl+Alt+T`, `Alt+7`, `Ctrl+Shift+T`, `Ctrl+T`,| Open or reuse the embedded terminal in the opposite physical pane. |
| Folder/Navigation | `Ctrl+Enter` | Insert the focused item using the authenticated contextual leaf/full-path rule. |
| Folder/Navigation | `Ctrl+Shift+Enter` | Insert the focused item's full path. |
| Folder/Navigation | `Ctrl+Space`, `Ctrl+Shift+Space` | Insert the current directory without Enter. |
| Terminal | `Ctrl+C` | Copy when a selection exists; otherwise deliver exactly one interrupt to the foreground process. |
| Any RedSalamander content | `F11` | Run the existing Salamander Connect command; never toggle Terminal fullscreen. |

### Non-goals

- importing or editing Windows Terminal profiles, schemes, or `settings.json`;
- adding multiple independent terminal sessions inside one physical pane,
  Quake mode, or a global hotkey;
- reserving deferred chords before their features exist; and
- implementing mark mode, profile slots, duplicate-tab semantics, or
  clear-buffer behavior inside P1; and
- scraping rendered terminal output, injecting hidden shell commands, or
  automatically executing a selected history suggestion to imitate shell
  integration.

## Current repository evidence

- `ShortcutDefaults.cpp`, `ShortcutManager`, settings, and Preferences currently
  expose only `functionBar` and `folderView` shortcut maps.
- `RedSalamander.cpp` deliberately bypasses menu mnemonics, Function Bar state,
  FolderWindow shortcut dispatch, fullscreen Escape, and static accelerators
  when `FolderWindow::IsTerminalInputTarget(...)` is true.
- `Terminal.dll` currently owns confirmation overlays, selection, copy, paste,
  `Ctrl+C` interrupt fall-through, special-key encoding, AltGr/dead-key input,
  IME commit, and Ghostty writes. It already handles `Ctrl+Shift+C`,
  `Ctrl+Shift+V`, `Shift+Insert`, and `Ctrl+Shift+A` directly.
- The plugin already has bounded `scrollViewport(...)` and
  `scrollViewportTo(...)` paths and a validated `fontSizeDip` setting in the
  `8..32` DIP range, but it has no keyboard scrollback or session-local zoom
  commands.
- `ITerminal` is a source-tree/release-lockstep interface and may be extended in
  this feature. Terminal actions remain a separate queryable capability, while
  the weak event callback registration is part of `ITerminal` itself.
- Terminal.dll already publishes `TerminalLifecycleState::Exited` internally
  after ConPTY output ends, but the wake targets only the plugin child window.
  The host receives no generation-checked exit event, so a shell `exit` can
  leave a dead Terminal content tab visible.
- There is no plugin-owned Find surface, trusted command-history provider,
  history Suggestions surface, or app-owned floating Terminal host today.
- Folder, Preview, and Terminal are stable content-tab identities, and a
  physical pane has at most one embedded Terminal session. One singleton
  floating Terminal window may host any number of tabbed Terminal instances
  and does not change that per-pane invariant.

## Remaining outcomes

1. Add a configurable terminal shortcut scope with explicit pass-through and
   context-aware conflict handling.
2. Add a narrow plugin-owned arbitration boundary that preserves security,
   shell/TUI input, international text entry, and key sequence integrity.
3. Deliver the 48 P1 Windows Terminal mappings plus Salamander's retained
   selection-aware `Ctrl+C` and `Shift+F10` behaviors without changing current
   folder meanings.
4. Deliver Terminal Find, trusted history Suggestions, and one remembered
   floating Terminal window with as many tabs as the user opens.
5. Persist the floating window's placement plus tab order, active tab, profile,
   and trusted path, and close only the affected tab after its real shell
   session exits and final output is published.
6. Prove settings, accessibility, localization, correctness, teardown, and
   performance, then merge the durable behavior into the authoritative specs.

## Shortcut helper window review and target design

### Current helper evidence

`cmd/app/showShortcuts` (`F1`) already opens `ShortcutsWindow`, an independent
top-level DxUi window with a searchable, sortable, grouped grid. The current
implementation is a strong base but is incomplete for Operation Chordforge:

- `ShortcutsWindow.cpp` defines only Function Bar and Folder View group IDs and
  builds rows only from `shortcuts.functionBar` and `shortcuts.folderView`.
- `ShortcutRow` is **binding-centric**: it stores one virtual key/modifier pair
  and one command ID. This is the required Helper model: commands with aliases
  appear as separate binding rows so every configured chord can be inspected.
- The two columns are Command and Key. Command already contains localized title
  plus description; Key contains one trailing-aligned rendered chord.
- Search already matches localized command name, description, tooltip, and
  shortcut text. Enter already invokes the selected command. Collapse, sort,
  column layout, selected-row retention, UI Automation, and bounded visible-grid
  work already have focused regression coverage.
- The only row icon today is the conflict mark. There is no command visual
  metadata and no Terminal group, so the helper cannot present all terminal
  actions cleanly.

### Required helper result

Keep the existing independent F1 window and DxUi grid **binding-centric**:

1. Add stable Application and Terminal groups
   (`kGroupStableIdApplication = 3`, `kGroupStableIdTerminal = 4`) sourced from
   `shortcuts.application` and `shortcuts.terminal`. Persist
   `applicationCollapsed` and `terminalCollapsed` beside the existing group
   states.
2. Preserve exactly one row per configured binding. Each row contains the
   command icon, localized primary name and secondary description, plus exactly
   one trailing shortcut keycap. Copy therefore has separate `Ctrl+Shift+C`
   and `Ctrl+Insert` rows; Paste has separate `Ctrl+Shift+V` and `Shift+Insert`
   rows; the merged conditional copy/pass-through command has separate `Enter`
   and `Ctrl+C` rows. The Helper must never aggregate aliases into one command
   row.
3. The Terminal group must show all 46 terminal-scope default binding rows
   defined below, backed by all 39 Terminal-scope-eligible P1 command
   definitions.
   The Application group owns three global rows: Settings (`Ctrl+,`), Open
   Settings File (`Ctrl+Shift+,`), and Command Palette (`Ctrl+Shift+P`). The
   existing Function Bar group keeps Connect (`F11`). These scopes expose 50
   relevant default binding rows while Terminal content has focus. Every
   user-added alias creates another row. Every explicit
   `cmd/shortcut/passthrough` entry creates an informational **Pass through to
   terminal** binding row; the consumed `cmd/shortcut/unassigned` sentinel
   stays hidden.
4. Render the existing Key cell as one keyboard-key-shaped chip per row. The
   conflict warning belongs to that binding row/chip. The chip retains semantic
   key ordering and exposes its full localized chord as UI Automation text.
5. Extend `CommandInfo` through typed visual metadata (for example a
   `CommandVisualId` enum resolved to Segoe Fluent Icons with a text fallback),
   terminal eligibility, palette visibility, and search keywords. Do not store
   raw font glyphs or localized display text in settings or the Terminal ABI.
6. Search evaluates each binding row independently and matches command name,
   description, stable command ID, keywords, and that row's shortcut. Preserve
   the Helper's current sort, group, selection-retention, and binding-row
   identity behavior; adding or removing an alias adds or removes a row rather
   than mutating another row's shortcut list.
7. Enter/click invokes the enabled command associated with the selected
   binding through the existing dispatch path. Pass-through rows are
   informational and never dispatch. Bindings whose commands are unavailable
   in the current context remain visible with subdued text and a localized
   reason rather than disappearing unpredictably.

Extract one app-layer `ShortcutCommandCatalog` model used by both the F1 helper
and the global RedSalamander Command Palette. This module may live under `RedSalamander/`
because it depends on `CommandRegistry`, app settings, and localized resources;
it is not a policy-neutral `Common` helper. Document that dependency-layer
difference at the definition and update `Specs/Core/Core_SharedHelpers.md` as
required by repository helper policy. The model exposes canonical command
definitions and their individual bindings. `ShortcutsWindow` projects one row
per binding; `RedSalamanderCommandPalette` alone groups active bindings by
canonical command ID into one command row with multiple shortcut chips. Share
command lookup, search terms, icon resolution, enabled-state, and keycap
formatting without forcing the two windows to share a row model.

## Complete Windows Terminal default review

Decision meanings:

- **P1 adopt/adapt:** default binding delivered by this plan.
- **Follow-up:** not captured in P1; the chord passes through until a separately
  approved implementation owns the missing feature.
- **Pass through:** deliberately left to the shell/TUI because no useful or
  honest Salamander action exists.
- **Reject:** must not be implemented in-process under this plan.

### Application actions — 12 bindings

| Count | Windows Terminal chord | Upstream intent | RedSalamander decision |
|---:|---|---|---|
| 1 | `Alt+F4` | Close window | **P1 adapt:** `cmd/app/exit`; preserve the ordinary top-level close path. |
| 1 | `Alt+Enter` | Toggle fullscreen | **P1 adopt:** existing `cmd/app/fullScreen` while terminal-focused. |
| 1 | `F11` | Toggle fullscreen | **P1 retain Salamander:** never bind F11 to Terminal fullscreen. Keep the existing Function Bar `cmd/pane/connect` action authoritative, including while Terminal has focus, after the Terminal security gate. |
| 1 | `Ctrl+Shift+Space` | New-tab dropdown | **P1 adapt:** open a localized Terminal Session menu for the focused physical pane. Folder scope retains Insert Current Dir. |
| 1 | `Ctrl+,` | Settings UI | **P1 adapt:** invoke existing global `cmd/app/preferences` in Application scope; do not add `cmd/terminal/settings` or deep-link to a Terminal-only page. |
| 1 | `Ctrl+Shift+,` | Open settings file | **P1 adapt:** global `cmd/app/openSettingsFile` opens the active RedSalamander settings file in its shell-default editor, creating the current file first when necessary. |
| 1 | `Ctrl+Alt+,` | Open default settings file | **Pass through:** no honest equivalent. |
| 1 | `Ctrl+Shift+F` | Find terminal text | **P1 adopt:** open the plugin-owned Find surface specified below and search the current Terminal buffer/scrollback without sending input to the shell. |
| 1 | `Ctrl+Shift+P` | Command palette | **P1 adopt:** open the searchable, icon-bearing global RedSalamander Command Palette specified below. |
| 1 | `Win+scan-code 41` | Quake mode | **Reject:** do not register a global hotkey or create a drop-down window mode. |
| 1 | `Alt+Space` | System menu | **P1 adopt:** show the RedSalamander top-level window system menu. |
| 1 | `Ctrl+Shift+.` | Suggestions | **P1 adapt:** open the plugin-owned command-history Suggestions surface specified below, visually following the supplied compact history popup reference. |

### Tab actions — 23 bindings

| Count | Windows Terminal chord(s) | Upstream intent | RedSalamander decision |
|---:|---|---|---|
| 1 | `Ctrl+Shift+T` | New tab | **P1 adapt:** from embedded/Folder context, open or reuse Terminal in the other physical pane and focus it; from the floating Terminal window, add a tab at the active tab's trusted path. Also add this non-conflicting binding to Folder scope. |
| 1 | `Ctrl+Shift+N` | New window | **P1 adapt:** `cmd/terminal/openFloatingWindow` creates or focuses the singleton floating Terminal window. If it already exists, add and select a new tab at the invoking trusted path. |
| 9 | `Ctrl+Shift+1` … `Ctrl+Shift+9` | New tab with profile 1…9 | **Follow-up:** profile-slot creation/replacement requires a separate lifecycle decision; pass through in P1. Folder scope retains Set Hot Path. |
| 1 | `Ctrl+Shift+D` | Duplicate tab | **Follow-up:** duplicating profile and trusted cwd into the other pane requires a separate lifecycle/security plan; pass through. |
| 1 | `Ctrl+Tab` | Next tab | **P1 adapt:** cycle available Folder/Preview/Terminal content in an embedded pane or floating tabs in the singleton window; add the pane-content variant to Folder scope too. |
| 1 | `Ctrl+Shift+Tab` | Previous tab | **P1 adapt:** reverse-cycle the applicable embedded content or floating tabs; add the pane-content variant to Folder scope too. |
| 3 | `Ctrl+Alt+1` … `Ctrl+Alt+3` | Select tab by index | **P1 adapt:** in embedded/Folder context select stable identity Folder=1, Preview=2, Terminal=3; in the floating window select tab index 1…3. An unavailable target is a no-op. Add the pane-content variants to Folder scope. AltGr never triggers these actions. |
| 5 | `Ctrl+Alt+4` … `Ctrl+Alt+8` | Select tab by index | **P1 adapt:** select floating tab index 4…8; pass the original chord through when an embedded Terminal owns input or that floating index is unavailable. AltGr never triggers it. |
| 1 | `Ctrl+Alt+9` | Select last tab | **P1 adapt:** select the last available embedded content identity or floating tab; add the pane-content variant to Folder scope. AltGr never triggers it. |

### Pane actions — 12 bindings

| Count | Windows Terminal chord | Upstream intent | RedSalamander decision |
|---:|---|---|---|
| 1 | `Ctrl+Shift+W` | Close pane | **P1 adapt:** close the focused embedded Terminal session or selected floating Terminal tab through the same lifecycle path. |
| 2 | `Alt+Shift+-`, `Alt+Shift++` | Duplicate split down/right | **Pass through:** fixed physical panes are not terminal splits. |
| 1 | `Alt+Shift+Down` | Resize pane down | **Pass through:** no horizontal splitter. |
| 1 | `Alt+Shift+Left` | Resize pane left | **P1 adapt:** move the main divider left by a DPI-scaled step, respecting current minimum widths. |
| 1 | `Alt+Shift+Right` | Resize pane right | **P1 adapt:** move the main divider right by the same step and constraints. |
| 1 | `Alt+Shift+Up` | Resize pane up | **Pass through:** no horizontal splitter. |
| 1 | `Alt+Down` | Focus pane down | **Pass through:** no lower physical pane. |
| 1 | `Alt+Left` | Focus pane left | **P1 adapt:** focus the left physical pane and its selected content. |
| 1 | `Alt+Right` | Focus pane right | **P1 adapt:** focus the right physical pane and its selected content. |
| 1 | `Alt+Up` | Focus pane up | **Pass through:** no upper physical pane. |
| 1 | `Ctrl+Alt+Left` | Focus previous pane | **P1 adapt:** toggle physical pane focus; AltGr never triggers it. |

### Clipboard and selection actions — 8 bindings

| Count | Windows Terminal chord | Upstream intent | RedSalamander decision |
|---:|---|---|---|
| 1 | `Ctrl+Shift+C` | Copy | **P1 adopt:** copy Terminal selection. |
| 1 | `Ctrl+Insert` | Copy | **P1 adopt:** same action. |
| 1 | `Enter` | Copy only when selection exists | **P1 adopt:** copy and consume only with a selection; otherwise deliver Enter unchanged. Confirmation-overlay Enter remains higher priority. |
| 1 | `Ctrl+Shift+V` | Paste | **P1 adopt:** retain the plugin-owned bounded/safe paste pipeline. |
| 1 | `Shift+Insert` | Paste | **P1 adopt:** same pipeline. |
| 1 | `Ctrl+Shift+A` | Select all | **P1 adopt:** retain Ghostty-owned selection semantics. |
| 1 | `Ctrl+Shift+M` | Mark mode | **Follow-up:** no mark-mode model; pass through in P1. |
| 1 | `Menu` | Context menu | **P1 adopt:** open the localized Terminal context menu at the caret/selection anchor. `Shift+F10` invokes the same accessible action. |

### Scrollback actions — 7 bindings

| Count | Windows Terminal chord | Upstream intent | RedSalamander decision |
|---:|---|---|---|
| 1 | `Ctrl+Shift+Down` | Scroll one line down | **P1 adopt:** use the existing viewport path on the primary buffer; pass through to an alternate-screen/TUI owner. |
| 1 | `Ctrl+Shift+PageDown` | Scroll one page down | **P1 adopt:** one visible page with the same eligibility rule. |
| 1 | `Ctrl+Shift+Up` | Scroll one line up | **P1 adopt:** primary buffer only. |
| 1 | `Ctrl+Shift+PageUp` | Scroll one page up | **P1 adopt:** primary buffer only. |
| 1 | `Ctrl+Shift+Home` | Scroll to top | **P1 adopt:** use the existing absolute-row mapping. |
| 1 | `Ctrl+Shift+End` | Scroll to bottom | **P1 adopt:** restore follow-bottom through the existing viewport model. |
| 1 | `Ctrl+Shift+K` | Clear buffer | **Follow-up:** implement only if the selected Ghostty boundary exposes a real buffer-clear operation; never inject shell text or ANSI as an imitation. Pass through in P1. |

### Visual actions — 6 bindings

| Count | Windows Terminal chord | Upstream intent | RedSalamander decision |
|---:|---|---|---|
| 1 | `Ctrl++` | Increase font size | **P1 adopt:** increase session-local font size by 1 DIP; bind the physical number-row key left of Backspace, never a `+`/`=` character. |
| 1 | `Ctrl+-` | Decrease font size | **P1 adopt:** decrease by 1 DIP; bind the physical number-row key right of `0`, never a `-` character. |
| 1 | `Ctrl+Numpad+` | Increase font size | **P1 adopt:** same action. |
| 1 | `Ctrl+Numpad-` | Decrease font size | **P1 adopt:** same action. |
| 1 | `Ctrl+0` | Reset font size | **P1 adopt:** reset to the configured Terminal base size. Folder scope retains Hot Path 10. |
| 1 | `Ctrl+Numpad0` | Reset font size | **P1 adopt:** same action. |

Coverage check: `12 application + 23 tab + 12 pane + 8 clipboard + 7
scrollback + 6 visual = 68` reviewed bindings. The decisions are `48 P1`,
`12 follow-up mappings that pass through in P1`, and `8 pass-through or
rejected`.

The canonical Terminal map also retains two RedSalamander/Windows conventions:
`Enter` and `Ctrl+C` -> `cmd/terminal/copySelectionOrPassthrough` copy when a
selection exists and otherwise pass the original Enter or interrupt input
through exactly once;
`Shift+F10` -> `cmd/terminal/contextMenu` is the keyboard alternative to the
`Menu` key. They are additional to, not part of, the 68-binding Windows Terminal
count. The Terminal scope contains 46 defaults. Application contributes three
global bindings (`Ctrl+,`, `Ctrl+Shift+,`, `Ctrl+Shift+P`), and the existing
Function Bar contributes `F11`, for 50 relevant bindings while Terminal has
focus.
`Ctrl+V` is intentionally not captured: only `Ctrl+Shift+V`, `Shift+Insert`,
and `WM_PASTE` invoke terminal paste, so shells/TUIs retain `Ctrl+V`.

### P1 command definitions reachable from Terminal context

Every row below is a normative `CommandRegistry` definition, not merely a
display label or router constant. The F1 Helper shows each listed alias as a
separate binding row. The global Command Palette groups aliases into one
command row in the order shown.

Exactly 39 definitions are Terminal-scope-eligible and selectable in
Preferences -> Keyboard -> Terminal. Three additional definitions have their
defaults in global Application scope: `cmd/app/preferences`,
`cmd/app/openSettingsFile`, and `cmd/app/commandPalette`. The existing
`cmd/pane/connect` remains in Function Bar scope. Together the table contains
43 unique definitions. Each definition supplies a stable ID, localized title
and description resources, executor owner, Terminal/context eligibility,
enabled state query, palette visibility/search keywords, visual ID, and
dispatch path.

| Command | Stable command ID | Default shortcut(s) | Visual intent |
|---|---|---|---|
| Exit RedSalamander | `cmd/app/exit` | `Alt+F4` | Close/window |
| Toggle Full Screen | `cmd/app/fullScreen` | `Alt+Enter` | Full-screen corners |
| Connect | `cmd/pane/connect` | `F11` (existing Function Bar scope) | Connection/plug |
| Terminal Session Menu | `cmd/terminal/sessionMenu` | `Ctrl+Shift+Space` | Terminal with chevron |
| Settings | `cmd/app/preferences` | `Ctrl+,` (global Application scope) | Settings gear |
| Open Settings File | `cmd/app/openSettingsFile` | `Ctrl+Shift+,` (global Application scope) | File with gear/code |
| Find Terminal Text | `cmd/terminal/find` | `Ctrl+Shift+F` | Search in document |
| RedSalamander Command Palette | `cmd/app/commandPalette` | `Ctrl+Shift+P` (global Application scope) | Search with command sparkle |
| Window System Menu | `cmd/app/systemMenu` | `Alt+Space` | App window |
| Command History Suggestions | `cmd/terminal/suggestions` | `Ctrl+Shift+.` | History with search |
| New Terminal Tab / Other Pane | `cmd/terminal/tab/new` | `Ctrl+Shift+T` | Terminal tab with plus |
| Open Floating Terminal / New Tab | `cmd/terminal/openFloatingWindow` | `Ctrl+Shift+N` | Terminal window with plus |
| Next Terminal Tab/Content | `cmd/terminal/tab/next` | `Ctrl+Tab` | Next tab |
| Previous Terminal Tab/Content | `cmd/terminal/tab/previous` | `Ctrl+Shift+Tab` | Previous tab |
| Select Terminal Tab/Content 1 | `cmd/terminal/tab/select/1` | `Ctrl+Alt+1` | Tab 1 / Folder |
| Select Terminal Tab/Content 2 | `cmd/terminal/tab/select/2` | `Ctrl+Alt+2` | Tab 2 / Preview |
| Select Terminal Tab/Content 3 | `cmd/terminal/tab/select/3` | `Ctrl+Alt+3` | Tab 3 / Terminal |
| Select Floating Terminal Tab 4 | `cmd/terminal/tab/select/4` | `Ctrl+Alt+4` | Tab 4 |
| Select Floating Terminal Tab 5 | `cmd/terminal/tab/select/5` | `Ctrl+Alt+5` | Tab 5 |
| Select Floating Terminal Tab 6 | `cmd/terminal/tab/select/6` | `Ctrl+Alt+6` | Tab 6 |
| Select Floating Terminal Tab 7 | `cmd/terminal/tab/select/7` | `Ctrl+Alt+7` | Tab 7 |
| Select Floating Terminal Tab 8 | `cmd/terminal/tab/select/8` | `Ctrl+Alt+8` | Tab 8 |
| Select Last Terminal Tab/Content | `cmd/terminal/tab/last` | `Ctrl+Alt+9` | Last tab |
| Close Terminal Tab | `cmd/terminal/close` | `Ctrl+Shift+W` | Close tab/session |
| Move Divider Left | `cmd/pane/resizeSplitter/left` | `Alt+Shift+Left` | Horizontal resize left |
| Move Divider Right | `cmd/pane/resizeSplitter/right` | `Alt+Shift+Right` | Horizontal resize right |
| Focus Left Pane | `cmd/pane/focus/left` | `Alt+Left` | Left arrow/pane |
| Focus Right Pane | `cmd/pane/focus/right` | `Alt+Right` | Right arrow/pane |
| Switch Pane Focus | `cmd/pane/switchPaneFocus` | `Ctrl+Alt+Left` | Swap focus |
| Copy Terminal Selection | `cmd/terminal/copy` | `Ctrl+Shift+C`, `Ctrl+Insert` | Copy |
| Copy Selection or Pass Through | `cmd/terminal/copySelectionOrPassthrough` | `Enter`, `Ctrl+C` | Copy/input fork |
| Paste in Terminal | `cmd/terminal/paste` | `Ctrl+Shift+V`, `Shift+Insert` | Paste |
| Select All Terminal Text | `cmd/terminal/selectAll` | `Ctrl+Shift+A` | Select all |
| Terminal Context Menu | `cmd/terminal/contextMenu` | `Menu`, `Shift+F10` | Menu/more |
| Scroll One Line Down | `cmd/terminal/scroll/lineDown` | `Ctrl+Shift+Down` | Down arrow |
| Scroll One Page Down | `cmd/terminal/scroll/pageDown` | `Ctrl+Shift+PageDown` | Page down |
| Scroll One Line Up | `cmd/terminal/scroll/lineUp` | `Ctrl+Shift+Up` | Up arrow |
| Scroll One Page Up | `cmd/terminal/scroll/pageUp` | `Ctrl+Shift+PageUp` | Page up |
| Scroll to Top | `cmd/terminal/scroll/top` | `Ctrl+Shift+Home` | Top boundary |
| Scroll to Bottom | `cmd/terminal/scroll/bottom` | `Ctrl+Shift+End` | Bottom boundary |
| Increase Terminal Font | `cmd/terminal/font/increase` | `Ctrl+<number-row plus position>`, `Ctrl+Numpad+` | Zoom in |
| Decrease Terminal Font | `cmd/terminal/font/decrease` | `Ctrl+<number-row minus position>`, `Ctrl+Numpad-` | Zoom out |
| Reset Terminal Font | `cmd/terminal/font/reset` | `Ctrl+0`, `Ctrl+Numpad0` | Reset zoom |

## Shortcut settings and Preferences contract

1. Add `ShortcutsSettings::application` and `ShortcutsSettings::terminal` plus
   optional `shortcuts.applicationCollapsed` and
   `shortcuts.terminalCollapsed`. These are optional additive fields and do
   **not** justify invalidating existing schema-v16 settings. Missing bindings
   receive canonical defaults using the same restoration rules as existing
   scopes. `Ctrl+,`, `Ctrl+Shift+,`, and `Ctrl+Shift+P` exist once in
   Application, never duplicated into Folder or Terminal defaults. `F11`
   remains in Function Bar and is never introduced as a Terminal fullscreen
   default.
   Extend shortcut key identity with an explicit physical-position kind for
   `NumberRowPlus` and `NumberRowMinus`, plus virtual-key names for
   `NumpadPlus`, `NumpadMinus`, `Numpad0`, and `Menu`; main-keyboard and numpad
   identities must remain distinct.
2. Preserve `cmd/shortcut/unassigned` as a consumed no-op in every scope. Add
   `cmd/shortcut/passthrough` as a terminal-only persisted sentinel. It means
   “do not run a RedSalamander action; deliver this exact chord to the embedded
   terminal.” Preserve it through load/save/import/export, hide it from normal
   command reverse lookup, and show it as an explicit Preferences choice.
3. An unbound non-default terminal chord naturally passes through. Removing a
   canonical terminal default must ask/offer **Pass through** and **No action**;
   it must not silently choose between shell input and consumption.
4. Extend command metadata with application/terminal eligibility and executor
   owner rather than maintaining a second ad hoc allowlist. Reuse existing
   stable IDs such as `cmd/app/exit`, `cmd/app/fullScreen`,
   `cmd/app/preferences`, `cmd/pane/connect`, and
   `cmd/pane/openCommandShell`. Add stable `cmd/app/openSettingsFile`,
   `cmd/app/commandPalette`, `cmd/pane/contentTab/*`, `cmd/pane/focus/*`,
   `cmd/pane/resizeSplitter/*`, `cmd/terminal/find`,
   `cmd/terminal/suggestions`, `cmd/terminal/openFloatingWindow`,
   `cmd/terminal/tab/*`, and the other `cmd/terminal/*` IDs in the
   command-definition table. Do not add
   `cmd/terminal/settings`.
5. The Terminal command picker is populated from `CommandRegistry`, never a
   Preferences-local list. It exposes every one of the 39
   Terminal-scope-eligible command definitions in the table, including
   host-owned `cmd/app/*` and
   `cmd/pane/*` actions that are valid from Terminal scope. All 46 factory
   bindings are editable, removable, restorable, and replaceable; users may
   add any number of aliases for any eligible command. Pass through and No
   action remain explicit pseudo-command choices. Follow-up/rejected Windows
   Terminal features do not receive fake executable commands; when a follow-up
   becomes implemented, registering its complete command definition is a
   prerequisite for making it selectable.
6. Preferences -> Keyboard gains localized **Application** and **Terminal**
   sections. A duplicate inside one scope is a conflict. Context bindings take
   precedence over Application bindings; an intentional context override is
   shown as “Overrides global shortcut” rather than a same-scope conflict. A
   Terminal `cmd/shortcut/passthrough` entry for `Ctrl+Shift+P` explicitly
   suppresses the global palette and sends the chord to the child. Search,
   sort, collapse, import/export, Restore defaults, UI Automation, and
   dirty-state behavior must match the existing groups.
7. `cmd/app/openSettingsFile` resolves the same active main settings path used
   by Preferences Advanced through `Common::Settings::GetSettingsPath(appId)`.
   If absent, it saves the current settings first, then invokes the shell-
   default editor. Failure shows a localized error and never marks Preferences
   dirty. Reuse or extract the existing settings-file-opening helper rather
   than duplicating path/create/ShellExecute policy.
8. Persist the singleton under `terminal.floatingWindow`. Store one normalized
   top-level placement (normal bounds, DPI, monitor identity, and show state),
   `wasOpenAtCleanShutdown`, the active tab ID, and an ordered `tabs` array.
   Each tab record contains a stable tab ID, Terminal profile ID, normalized
   logical location (`providerId` plus canonical trusted path), and order. Tabs
   sharing the same path remain distinct by tab ID. Closing a tab removes its
   record; never retain a list of closed sessions. Persist data only: no live
   process, HWND, plugin pointer, screen contents, command history, or shell
   input buffer.
9. Reopening the floating window restores its saved placement, ordered tabs,
   active tab, profiles, and trusted paths. If it was open at a clean app
   shutdown, restore it automatically on the next startup; an explicit user
   close keeps the remembered tab set for the next `Ctrl+Shift+N` but clears
   `wasOpenAtCleanShutdown`. On an explicit later reopen, restore the saved set;
   select an existing tab matching the invoking profile/path or add one when no
   match exists. A missing/invalid tab path falls back to the normal Terminal
   launch location with a localized diagnostic instead of dropping the
   remaining tabs.
10. Command names, descriptions, group labels, Terminal Session menu, context
   menu, pass-through/no-action explanations, and error text use `.rc`
   resources. Update every shipped satellite with the exact resource-ID set.

## Input arbitration and ABI design

Do not weaken the existing terminal-priority bypass. Add a separately queryable `ITerminalActions` IID to the
same Terminal instance. It owns shortcut arbitration plus menu action
state/execution; it is not a callback from the plugin into the host.

The new packed request records at least:

- `sizeBytes`, stable command ID span, message kind (`WM_KEYDOWN` or
  `WM_SYSKEYDOWN`), virtual key, scan code, extended/system flags, repeat count,
  previous-down state, normalized modifiers, and Terminal instance/session
  generation;
- left/right modifier state sufficient to recognize AltGr and avoid treating
  Right Alt's synthetic Ctrl+Alt as a shortcut;
- reserved zeroed fields following the current public ABI convention.

The result is one of:

- `PassThrough`: dispatch the original message through the unchanged child
  path;
- `Handled`: the plugin executed a terminal-local action; consume it;
- `InvokeHostCommand`: the plugin yielded and the host may execute the already
  resolved, context-eligible application or terminal command;
- `Blocked`: an active security/exclusive-input surface consumed the chord and
  no host command may run.

The same interface exposes typed `GetActionState` and `ExecuteAction` methods
for context-menu commands. The plugin returns enabled/checked state and a
bounded child-coordinate anchor record; the host owns menu presentation and
maps the anchor to screen coordinates. Plugin-local action execution never
travels through `WM_COMMAND`, and the plugin never receives or retains a host
callback pointer.

Required routing order:

1. A visible closable host alert retains first refusal, including Escape.
2. If the target is not the live terminal child/descendant, resolve an
   Application binding first, then use the existing context-specific
   Function Bar/Folder/accelerator paths. Owned modal dialogs and text-entry
   controls keep their existing higher-priority local handling.
3. For a terminal target, resolve an exact Terminal binding first. A
   `cmd/shortcut/passthrough` match dispatches the original message to the child
   and suppresses any same-chord Application or Function Bar binding; an
   unassigned match is consumed.
4. If no Terminal binding exists, resolve the Application scope, then an
   eligible existing Function Bar binding. This keeps `F11` mapped to
   `cmd/pane/connect` rather than Terminal fullscreen. If none of those scopes
   matches, dispatch the original message to the child unchanged.
5. For a matched command binding, synchronously call `ITerminalActions` on the UI
   thread before `TranslateMessage`.
6. The plugin first resolves confirmation/security state, then terminal-local
   conditional actions, then alternate-screen/TUI eligibility. Only an
   eligible host action returns `InvokeHostCommand`.
7. If `QueryInterface` or routing fails, fail open to the existing child input
   path. The child plugin still owns its confirmation overlay and will consume
   unsafe input there. Never fail open to a host command.

The host tracks a consumed key-down by terminal instance/session generation,
virtual key, scan code, extended flag, and system/non-system class. It consumes
the matching repeat/key-up sequence so Ghostty never sees an orphan release.
Clear stale entries on matching key-up, focus loss, terminal close, child
destruction, session-generation change, and app deactivation. A host command
may destroy or move focus away from the child, so no terminal pointer/window is
used after command dispatch.

Text-producing AltGr, dead-key, surrogate, and IME paths are never shortcut
inputs. `Ctrl+Alt+<printable>` bindings are ineligible while Right Alt/AltGr is
active. The Windows-reserved `Win` family is not added to the configurable
terminal scope in P1.

### Session-exit event boundary

Extend the release-lockstep `ITerminal` interface with `SetCallback(...)`; do
not turn `ITerminalActions` into a callback. `ITerminal::SetCallback(...)`
accepts a host-owned raw `ITerminalEventCallback*` plus an opaque cookie. The
callback is a non-COM vtable: no `IUnknown`, `AddRef`, `Release`, or
`wil::com_ptr`. Terminal.dll stores it weakly and returns the cookie verbatim.

The event record contains `sizeBytes`, event kind, instance ID, session
generation, state generation, exit-code presence/value, and
`finalSnapshotComplete`. The ConPTY reader never calls host code. It posts the
existing output-ready wake to the Terminal child; after the child UI thread has
consumed final output and observes the genuine root session as `Exited`, it
invokes one coalesced callback. Startup `Failed`, explicit host `Close`, and
intermediate child-process exits are not automatic-close events.

The host callback must not destroy the Terminal child inline from its WndProc.
It posts a typed `PostMessagePayload(...)` to the owning app window (main window
for an embedded tab, singleton floating root for a floating tab), then returns.
The receiver validates owner identity, stable floating tab ID when applicable,
Terminal instance ID, and session generation before using the existing
close-session path. A valid event removes the embedded Terminal content
identity and selects the normal Folder fallback, or closes only the matching
floating tab; the floating root closes when no tabs remain. Stale, duplicate,
or manually closed events are ignored.

`SetCallback(nullptr, nullptr)` is a hard barrier. Teardown stops producers,
clears the callback, calls `Close`, releases the Terminal instance, and unloads
only at the existing quiet point. No worker may call a cleared callback; no
posted payload may target the plugin child HWND; main/floating window
`WM_NCDESTROY` drains its registered payloads. Manual close, shell exit, app
shutdown, HWND reuse, and simultaneous output/exit races must converge on one
idempotent close.

### Physical plus/minus key-position contract

The user-visible names remain `Ctrl++` and `Ctrl+-`, but their identities are
physical positions, not characters or layout-dependent virtual keys:

- **Increase:** set-1 scan code `0x0D`, the number-row key immediately left of
  Backspace (US legends `=` / `+`).
- **Decrease:** set-1 scan code `0x0C`, the number-row key immediately right of
  `0` (US legends `-` / `_`).
- Numpad aliases remain the distinct `VK_ADD` and `VK_SUBTRACT` identities.

This follows the established repository precedent in
`Plugins/ViewerWeb/ViewerWeb.cpp` (`ZoomInVirtualKeyForLayout` /
`ZoomOutVirtualKeyForLayout`), `Plugins/ViewerImgRaw/ViewerImgRaw.cpp`, and
`ShortcutDefaults.cpp` (`DefaultSelectDialogVk` /
`DefaultUnselectDialogVk`). Do not add another local copy. Introduce or extend a
reviewed keyboard-position helper, catalog it in
`Specs/Core/Core_SharedHelpers.md`, and migrate the matching local helpers when
their dependency/ABI constraints allow. If a plugin cannot consume the shared
binary helper, keep a thin documented wrapper over the same header-level
constants rather than duplicating policy.

`ShortcutBinding` and its JSON representation accept exactly one key identity:
the existing `vk`, or a new localized-display-independent `keyPosition` token
(`numberRowPlus` or `numberRowMinus`). Chord hashing/conflict detection includes
the identity kind. Terminal defaults use `keyPosition`; they are never persisted
as `VK_OEM_PLUS`, `VK_OEM_MINUS`, `'+'`, `'='`, or `'-'`.

Use the same position identities for the existing factory Folder defaults
`cmd/pane/selection/selectDialog`, `cmd/pane/selection/unselectDialog`, and their
same-extension aliases, whose current builders already intend scan positions
`0x0D`/`0x0C`. Preserve explicit user-authored virtual-key bindings; default
restoration may migrate only a missing canonical default and must not rewrite a
custom chord during load/save.

At runtime, compare the scan code and extended bit from the original key message
against the physical binding before character translation. Do not call
`ToUnicodeEx`, and do not accept a same-character key at another position.
`GetKeyNameTextW`/the current layout may supply only the localized display label
for the helper/palette chip. A keyboard-layout change updates that label but not
the binding identity. Deterministic tests vary `wParam`/layout while retaining
scan code `0x0D` or `0x0C`, then vary the scan position while producing the same
character; only the physical-position cases may zoom.

### State policy

- Confirmation overlays consume all mapped keys except their own existing
  accept/cancel handling. A shortcut must never execute behind unsafe-paste,
  OSC 52, or hyperlink confirmation.
- Copy/paste/select-all and context-menu semantics remain plugin-owned because
  only the plugin has authoritative selection and security state. Both
  `cmd/terminal/copy` and `cmd/terminal/copySelectionOrPassthrough` clear the
  selection only after a successful clipboard copy; a failed clipboard write
  leaves the selection intact. For the conditional command, no selection means
  the original `Enter` or `Ctrl+C` message passes through exactly once.
- Scrollback actions run only against eligible primary-buffer history. They
  pass through while alternate-screen/TUI input owns the session.
- Pane/content navigation passes through while alternate-screen/TUI input owns
  the session. `F11` remains the host-level Connect command in both buffers and
  never toggles fullscreen; a Terminal-scope pass-through override is available
  for TUIs that need F11. Top-level `Alt+F4`, `Alt+Space`, and explicitly bound
  `Alt+Enter` retain window-level behavior after the plugin's security gate.
- Global `Ctrl+,` Settings and `Ctrl+Shift+,` Open Settings File are
  window-level Application commands after the Terminal security gate and have
  no Terminal-only variants.
- Global `Ctrl+Shift+P` is a window-level Application command after the Terminal
  security gate and remains available while alternate-screen/TUI input owns the
  session. A Terminal-scope pass-through override is the escape hatch for TUIs
  that need this chord.
- Font zoom is terminal-local and may run in either buffer. It changes only the
  live session, clamps to `8..32` DIP, resets to the current configured base,
  reflows through the existing resize/render path, and is not persisted.

### Menu surfaces

- The host-owned Terminal Session menu contains the stable Folder, Preview
  (when available), and Terminal identities; Open/Focus Terminal in Other Pane;
  Open Floating Terminal/New Tab; Find; Command History Suggestions; Close
  Terminal; global Settings; and Open Settings File. It reflects availability
  and the current identity and never creates a hidden embedded session.
- The host-owned DxUi Terminal context menu contains Copy, Paste, Select All,
  Find, Command History Suggestions, Scroll to Bottom, Open Floating Terminal /
  New Tab, global Settings, Open Settings File, and Close Terminal.
  Terminal.dll remains authoritative for selection/action state and executes
  plugin-local items through `ITerminalActions`; the host dispatches its own
  items only after the same gate. Copy is disabled without a selection. Labels,
  enabled states, keyboard invocation, focus return, and UI Automation names
  are deterministic and localized.

## Terminal Find (`Ctrl+Shift+F`)

`cmd/terminal/find` is a plugin-owned action available in embedded and floating
Terminal surfaces. It opens a compact DxUi Find bar inside the Terminal child;
the initial search field is focused and no key used by the bar reaches ConPTY.

- Search covers the current primary screen plus retained scrollback. In an
  alternate screen it searches only the current alternate-screen contents and
  exposes that limited scope in localized helper text; it never searches hidden
  primary history while a TUI owns the screen.
- The bar contains Search, previous/next result, case-sensitive and whole-word
  toggles, a live `current/total` result label, and Close. `Enter`/`F3` selects
  next, `Shift+Enter`/`Shift+F3` selects previous, `Esc` closes, and repeated
  `Ctrl+Shift+F` selects the current query.
- Matches are highlighted without mutating the underlying terminal selection.
  Closing Find restores the prior Terminal focus and viewport unless the user
  explicitly navigated to a match; the active match remains visible but no
  shell input is replayed.
- Build an immutable, generation-tagged searchable text/row snapshot in bounded
  batches. Filtering may run off the UI thread, is cancellable by query or
  buffer generation, and publishes only results matching the current Terminal
  instance/session/search generation. Never hold the Ghostty/session/render
  lock while filtering, painting, invoking UIA, or waiting for a worker.
- The UIA surface exposes a named Find toolbar, editable `ValuePattern`, toggle
  states, result count, and previous/next/close `InvokePattern`s. Match
  highlights remain represented through the Terminal text provider.

## Command History Suggestions (`Ctrl+Shift+.`)

`cmd/terminal/suggestions` is a plugin-owned command-history search surface,
distinct from the global RedSalamander Command Palette. It opens above the
active prompt/input area when trustworthy prompt geometry is available and
otherwise anchors to the lower center of the Terminal client.

### History data contract

- Introduce a plugin-owned `TerminalHistoryProvider` abstraction. P1 includes a
  trusted PowerShell/PSReadLine provider through the existing authenticated
  bootstrap channel and a current-session provider for shells that publish
  explicit command-boundary integration events. The provider supplies bounded
  UTF-16 command entries, recency, optional frequency, current prompt buffer,
  and trusted logical working directory.
- Never reconstruct commands from painted cells or arbitrary ConPTY output,
  inject a hidden command such as `doskey /history`, or execute a shell command
  to obtain history. A shell without a trusted provider opens the surface with
  a localized **Command history is unavailable for this shell** state rather
  than passing the shortcut through or guessing.
- The persisted history source is read-only, at most the newest 10,000 entries
  or 4 MiB per refresh, and parsed on a cancellable worker. Deduplicate exact
  commands while preserving most-recent occurrence and frequency. Do not log,
  instrument, synchronize, or persist command text in RedSalamander settings or
  performance JSONL. Clear in-memory snapshots when the Terminal session closes.
- Rank case-insensitive exact/prefix matches first, then token-prefix/fuzzy
  matches, then recency/frequency, with same-working-directory affinity only
  when the provider supplies a trusted path. Deterministic ties use normalized
  command text and original history order.

### Suggestions interaction and visual design

- Follow the supplied reference: a compact elevated surface with rounded themed
  border, a slim accent at the selected row, a History icon on every result,
  one-line command text with matched characters emphasized, and a search field
  along the bottom with a leading search/history icon and trailing Clear button.
  Show at most eight complete rows before scrolling; long commands ellipsize
  visually while UIA retains the full local text.
- Seed the query from the provider's current prompt buffer (for example `buil`)
  when available. Typing edits only the overlay query. `Up`/`Down`,
  `PageUp`/`PageDown`, pointer hover/click, and the wheel navigate; `Esc` closes
  and restores the original prompt unchanged.
- `Enter` or `Tab` closes the overlay and inserts the selected command through
  the existing bounded safe-paste/input path **without executing it**. The shell
  retains final editing and execution ownership. Revalidate instance/session,
  security state, prompt trust, and insertion size immediately before insertion;
  failure leaves the shell buffer unchanged and shows a localized nonmodal
  explanation.
- UIA exposes a popup root, editable search `ValuePattern`, results
  `SelectionPattern`, and history `DataItem`s with full command text. High
  Contrast uses system colors and no decorative elevation; theme colors resolve
  through the existing DxUi palette.

![Operation Chordforge Terminal command-history Suggestions mockup](../../Assets/OperationChordforge_HistorySuggestionsMockup.svg)

## Floating Terminal Window and Tabs (`Ctrl+Shift+N`)

`cmd/terminal/openFloatingWindow` creates or activates one independent,
modeless top-level `FloatingTerminalWindow`. It follows
`Specs/UI/UI_TopLevelToolWindows.md`, uses shared app
icon/backdrop/minimum-size/DPI behavior, appears independently in
Alt-Tab/taskbar, and is never owned by the main window. RedSalamander creates at
most one such top-level window, but the user may open as many tabs inside it as
available resources allow. Each tab owns one Terminal plugin instance; the
one-embedded-Terminal-per-physical-pane rule remains unchanged.

- With no remembered tabs, first invocation creates the window and one selected
  tab using the invoking profile and trusted logical path. With a remembered
  closed window, it restores the saved set and selects a matching profile/path
  tab or adds one when none matches. If the singleton is already open,
  `Ctrl+Shift+N` activates it, adds a new tab for the invoking profile/path, and
  selects that tab. Command Palette invocation from non-Terminal content uses
  the active pane's logical path and configured default Terminal profile.
- The tab strip exposes New Tab, Close Tab, pointer selection, reordering, and
  a per-tab context menu. `Ctrl+Shift+T` adds a tab when focus is inside the
  floating window; from embedded/Folder context it retains the other-physical-
  pane behavior. `Ctrl+Tab`/`Ctrl+Shift+Tab` cycle floating tabs while the
  floating window owns focus. These host-context variants use the same stable
  registered commands and remain customizable; they do not create a second
  shortcut registry.
- Each tab owns a stable tab ID, Terminal COM instance, child HWND, event
  registration, profile, trusted logical path, title/status projection, and
  close lifecycle. Closing/reordering one tab never closes or retargets another
  floating or embedded session. Closing the last tab closes the floating root;
  explicit root close first snapshots the ordered live-tab restore records,
  then closes all live instances through the normal plugin quiet point without
  treating them as individually closed tabs.
- Persistence follows the Settings contract above. Restore one normalized
  window placement, ordered tabs, selected tab, each tab's profile, and each
  tab's last trusted provider/path. A trusted cwd change updates only that tab's
  saved path after prompt confirmation. Untrusted escape, relative, or malformed
  paths never replace a persisted location. Duplicate paths are valid because
  tab IDs remain distinct.
- Coalesce `WM_MOVE`/`WM_SIZE`, tab reorder/selection, and trusted-path churn;
  persist settled state and the final clean-shutdown snapshot rather than
  writing every message. Restore sessions by creating fresh shells at the saved
  paths; never attempt to resurrect prior processes, screen contents, command
  history snapshots, or shell input buffers.

## Shell-exit close behavior

When the genuine root shell session reaches `Exited` after final output is
consumed, the event boundary above must close its UI automatically:

- an embedded session removes the Terminal content tab from its physical pane
  and selects/focuses Folder through the same idempotent path as Close Terminal;
- a floating session removes only its matching tab and persisted tab record;
  the singleton window closes only when that was the last tab; and
- exit code zero or nonzero does not leave a dead tab/window. Startup `Failed`
  remains visible as the existing diagnostic surface because it is not a normal
  shell exit.

The close must happen once even if shell exit races with manual close, app
shutdown, pane destruction, or a second queued event. Final output/state is
published before close, but the plan does not add a linger screen or confirmation
prompt.

## Global RedSalamander Command Palette (`Ctrl+Shift+P`)

`Ctrl+Shift+P` is no longer deferred. It binds
`cmd/app/commandPalette` in the global Application shortcut scope and opens a
host-owned transient DxUi palette over the active RedSalamander window. It is
the RedSalamander command launcher, not a Terminal-owned palette. Folder,
Navigation, Preview, and Terminal are invocation contexts that affect ranking
and enabled state, never separate palette implementations. When invoked from
Terminal, Terminal.dll must first pass the same confirmation/security gate as
every other host action. After that gate, the palette is available in both the
primary and alternate screen; users who need the chord in a TUI can add the
Terminal-scope pass-through override.

### Information and visual design

- Reuse `ShortcutCommandCatalog`; the palette is a filtered presentation, not a
  second command registry or search implementation.
- The catalogue includes every `CommandRegistry` entry whose typed metadata
  allows palette presentation. With an empty query, rank commands valid for the
  active context first, then global Application commands, then disabled/other
  context commands. Category labels such as Application, Pane, Terminal, View,
  and File describe ownership without acting as separate palettes.
- Disabled rows stay searchable and show a localized reason such as “No
  terminal selection” or “Preview is not available.”
- The search field is initially focused and includes a leading Fluent Search
  icon, localized placeholder **Search RedSalamander commands**, and a clear button
  only after text is entered.
- Each result row shows a 20-DIP Fluent command icon, 14-DIP command title,
  12-DIP one-line description, and every active shortcut as trailing keycap
  chips. Use the command visual metadata and a deterministic text fallback; do
  not use emoji or hard-coded private-use glyphs without font-availability
  fallback.
- Search/ranking uses the command-centric `ShortcutCommandCatalog` projection,
  independent of the Helper's binding rows. Highlight matching substrings in
  title/description and matching keycap chips without changing row height.
  For a nonempty query, rank exact/prefix title matches, word-prefix title
  matches, shortcut matches, description/keyword matches, then localized title
  and stable command ID for deterministic ties.
- The selected row uses the resolved accent selection fill and focus stroke.
  Disabled rows use the shared disabled palette. High Contrast suppresses
  decorative elevation and uses system-resolved borders/text; Rainbow uses the
  existing selected-row policy. No mockup color becomes a hard-coded product
  color.
- Target size is 760 x 560 DIP, clamped to the current work area and main-window
  client size. Rows are 58 DIP, the search field is 44 DIP, and at least five
  complete rows remain visible at the minimum supported palette height.

### Interaction contract

- `Ctrl+Shift+P` opens the palette and selects the first enabled result; pressing
  it again while open focuses/selects the current search text rather than
  opening another palette.
- Typing edits search. `Up`/`Down`, `PageUp`/`PageDown`, `Home`/`End`, and mouse
  hover/click navigate. `Enter` revalidates action state, closes the palette,
  then invokes the selected command once. A disabled result never dispatches.
- `Esc` closes without action and restores the exact originating app window and
  focus target. For a Terminal origin, restore the Terminal child only if its
  HWND, instance ID, and session generation are still current; otherwise
  restore the selected content's normal focus target.
- While the palette owns focus, no typed/search/navigation key reaches the
  originating app control, including ConPTY for a Terminal origin. Closing the
  palette does not replay buffered keys. App deactivation closes it without
  invocation.
- The palette exposes one UIA dialog/window root, an editable search
  `ValuePattern`, a results `SelectionPattern`, named command `DataItem`s, and
  shortcut/disabled-reason text in each accessible row name. Announce result
  count changes without speaking on every paint.
- Command state is snapshotted for list construction and revalidated through
  the owning host dispatcher or, for terminal-local actions,
  `ITerminalActions` immediately before execution. Process output may continue
  behind a palette opened from Terminal; no Terminal mutex or render lock is
  held while filtering, painting, or dispatching a host command.

### Illustrative mockup

The mockup shows the intended hierarchy, command icons, primary/secondary text,
multiple shortcut chips, selected state, and keyboard footer. It is illustrative
only; implementation colors come from the active DxUi palette.

![Operation Chordforge global RedSalamander Command Palette mockup](../../Assets/OperationChordforge_CommandPaletteMockup.svg)

## P1 implementation sequence

### 1. Freeze defaults and settings behavior

- Add a reviewed, table-driven default inventory for the Terminal scope and a
  coverage test that accounts for every P1 chord in this plan.
- Register the complete table of 39 Terminal-scope-eligible command
  definitions and make the Preferences Terminal command picker derive from
  that registry.
  Assert that every one of the 46 default bindings resolves to an eligible,
  dispatchable definition; no default may reference a router-only string.
- Extend `Common/SettingsStore.h`, `Common/Common/SettingsStore.cpp`,
  `Specs/SettingsStore.schema.json` where applicable, `ShortcutManager`,
  `ShortcutDefaults`, Preferences Keyboard, and shortcut import/export.
- Add the pass-through sentinel and per-scope conflict behavior. Preserve all
  existing user bindings and current folder default restoration.
- Add the non-conflicting shared Folder bindings: `Ctrl+Shift+T`, `Ctrl+Tab`,
  `Ctrl+Shift+Tab`, `Ctrl+Alt+1..3`, and `Ctrl+Alt+9`.

Verification: settings round-trip, malformed-section recovery, missing-default
restoration, explicit pass-through/no-action persistence, import/export,
Restore defaults, sort/search/collapse, same-scope conflict, cross-scope
non-conflict, and localized/UIA coverage.

### 2. Install the narrow terminal action ABI and router

- Add `ITerminalActions` and its records to
  `Common/PlugInterfaces/Terminal.h`; implement `QueryInterface` and UI-thread
  dispatch in Terminal.dll.
- Refactor the plugin's existing hard-coded clipboard key cases through one
  terminal-action dispatcher without moving selection/paste/security ownership
  into the EXE.
- Insert the host lookup/router call only at the current terminal-target branch
  in `RedSalamander.cpp`. Preserve the existing non-terminal message ordering.
- Add generation-aware down/repeat/up suppression with teardown clearing and
  no owning raw COM pointers.
- Extend `ITerminal` with the raw cookie callback registration and add the typed
  session-exit record. Marshal the reader's final exit through the child
  UI thread, then through `PostMessagePayload(...)` to the owning host window;
  never call or destroy host UI from the reader thread or inline child WndProc.

Verification: ABI layout/QI, old-interface fallback, unknown-command
pass-through, failed-call pass-through, exact modifier snapshots, key repeat,
key-up symmetry, focus loss, HWND reuse, session replacement, callback clear
barrier, duplicate/stale exit events, child teardown, and no dispatch through
stale pointers.

### 3. Deliver plugin-local P1 actions

- Clipboard/selection: copy aliases, the merged selection-aware
  `Enter`/`Ctrl+C` copy-or-pass-through command, post-copy selection clearing,
  bounded safe paste, select all, and accessible context menu.
- Scrollback: line/page/top/bottom actions through the existing viewport model,
  including repeat behavior and follow-bottom restoration.
- Visual: session-local font increment/decrement/reset with bounds, DPI/reflow,
  scrollbar, selection, cursor, IME caret, and device-recreation correctness.
- Find: generation-tagged buffer snapshots, cancellable filter/highlight
  publication, navigation, focus, UIA, primary/alternate-screen scope, and no
  ConPTY leakage.
- Suggestions: trusted history-provider abstraction, bounded PSReadLine and
  current-session sources, the reference-inspired popup, ranking, safe
  insertion without execution, privacy, cancellation, and unavailable-shell
  state.

Verification: exact bytes for pass-through input, zero bytes for consumed copy,
one ETX for unselected Ctrl+C, one Enter for unselected Enter, selection cleared
only after successful copy, confirmation precedence, clipboard limits,
bracketed paste, exact viewport rows, primary versus alternate screen, zoom
bounds/reset, Find matches/navigation/cancellation, history-source
trust/ranking/privacy/insertion, and no settings mutation.

### 4. Deliver host/pane P1 actions

- Reuse `cmd/app/exit`, `cmd/app/fullScreen`, `cmd/app/preferences`,
  `cmd/pane/connect`, and `cmd/pane/openCommandShell`; add
  `cmd/app/openSettingsFile`. Keep F11 on Connect, put `Ctrl+,` and
  `Ctrl+Shift+,` in Application scope, and add no Terminal Settings command.
- Add `cmd/terminal/tab/*` context-aware commands. Embedded execution uses
  stable Folder/Preview/Terminal identities and passes unavailable indices 4…8
  through; floating execution uses ordered tab IDs/indices, cycles only live
  tabs, wraps, and never confuses a reordered tab with its former index. Folder
  scope retains separate `cmd/pane/contentTab/*` bindings.
- Add deterministic left/right physical-pane focus and DPI-scaled main-splitter
  movement with existing minimum-width rules.
- Add Terminal Session menu and close-session command without inventing hidden
  embedded sessions or multiple sessions per physical pane.
- Add one app-owned `FloatingTerminalWindow` with an unbounded-by-policy tab
  model, per-tab Terminal ownership, tab strip/reorder/selection, one remembered
  DPI/monitor placement, ordered tab/profile/path persistence, and quiet-point
  teardown.
- Consume validated session-exit events through the existing close-session path
  so shell `exit` removes an embedded tab or only its matching floating tab;
  close the singleton root after its last tab exits.

Verification: both pane directions, unavailable Preview, both alternate tabs,
terminal reuse, focus restoration, splitter min/max/repeat/DPI, close teardown,
F11 Connect versus Alt+Enter fullscreen, global Settings/Open Settings File,
singleton enforcement, unlimited tab creation, tab order/active/profile/path
restore, duplicate-path tabs, embedded-versus-floating next/previous/direct/last
tab routing, cwd relinking, shell-exit tab/last-window close, manual-close/exit
races, menus, and TUI pass-through.

### 5. Deliver the complete helper and global Command Palette

- Create an app-layer `ShortcutCommandCatalog` shared by
  `ShortcutsWindow` and a new `RedSalamanderCommandPalette` surface. Extend
  `CommandRegistry` with typed command visual/context/search metadata.
- Keep the F1 Helper binding-centric, add its new Application and Terminal
  groups, and render one keycap per binding row while preserving its
  independent-window, search, sort, collapse, persistent layout, activation,
  and UIA contracts.
- Make only `RedSalamanderCommandPalette` command-centric: group active aliases
  by canonical command ID and show their shortcuts as multiple trailing chips.
- Implement the transient palette from the mockup using existing DxUi
  `WindowHost`, `TextField`, grid/list, typography, theme, tooltip, and Fluent
  icon paths. Do not add a legacy common-control fallback.
- Add `cmd/app/commandPalette` and its one `Ctrl+Shift+P` Application binding.
  Route open/close/focus/dispatch through the originating app context; use
  generation-aware host state and `ITerminalActions` revalidation only for a
  Terminal origin/action.

Verification: exact 46 binding-centric Terminal Helper rows backed by all 39
Terminal-scope-eligible command definitions, three relevant global Application
rows, and retained Function Bar F11 Connect; separate Helper alias rows, Palette-only
alias aggregation, pass-through rows, per-binding conflicts, icon fallback,
global/context ranking and ties, empty/no-match states, disabled reasons,
keyboard/mouse invocation from Folder/Preview/Terminal,
repeated open/close, terminal replacement while open, app deactivation, focus
restoration, theme/Rainbow/High Contrast/DPI, UIA, and bounded long-run
filtering/scrolling.

### 6. Documentation, localization, and closeout

Update the durable contracts together:

- `Specs/Terminal/Terminal_EmbeddedPlugin.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/UI/UI_TopLevelToolWindows.md`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/Core/Core_SharedHelpers.md`
- `Specs/UI/UI_KeyboardManagement.md` and
  `Specs/UI/UI_PreferencesDialog.md` where their contracts are affected
- `docs/KeyboardShortcuts.md`, `docs/UserGuide.md`, `docs/SettingsFile.md`, and
  `docs/Preferences.md`

Document context-specific conflicts explicitly, including:

| Chord | Folder/Navigation | Terminal |
|---|---|---|
| `Alt+Enter` | Properties | Full screen |
| `F11` | Connect | Connect; a user Terminal pass-through override may yield to a TUI |
| `Ctrl+,` | Global Settings | Global Settings |
| `Ctrl+Shift+,` | Open global settings file | Open global settings file |
| `Ctrl+Shift+Space` | Insert Current Dir | Terminal Session menu |
| `Ctrl+Shift+1..9` | Set Hot Path | Pass through in P1 |
| `Ctrl+0` | Hot Path 10 | Reset terminal font size |
| `Ctrl++`, `Ctrl+-` | Select/Unselect dialogs | Terminal font zoom |
| `Alt+Left`, `Alt+Right` | Folder history | Focus physical pane |

## Test and evidence plan

### Deterministic behavior coverage

- Extend `Tests/PluginContractTests` terminal selftests for router ABI,
  selection/clipboard, confirmation, exact child bytes, scrollback, font zoom,
  Find snapshots/results, trusted history providers, suggestion insertion,
  physical plus/minus scan positions, AltGr, alternate screen, key repeat/up,
  session-exit callbacks, and teardown.
- Extend Commands selftests under a `terminal_shortcut_` prefix for host message
  routing, settings/defaults, Preferences, content-tab navigation, pane focus,
  splitter movement, global settings commands, singleton floating-window
  placement, unlimited tab creation/reorder/restore and per-tab independence,
  shell-exit tab/last-window close, binding-centric Helper rows,
  complete Preferences command enumeration, command-centric Palette
  aggregation/search/activation/focus, and UIA.
- Extend `Tools/Tests/TerminalPluginSourceContracts.Tests.ps1` only for genuine
  ownership/ABI boundaries. Runtime behavior tests are mandatory; do not encode
  the implementation as a broad source-shape contract.
- Cover English plus every shipped Terminal/RedSalamander satellite in existing
  localization and resource-contract suites.
- Add one table-driven test whose expected partition totals are exactly
  `48 + 12 + 8 = 68`, so a documentation edit cannot silently omit an
  upstream-reviewed binding.

### Performance scenarios and budgets

Install instrumentation before behavior changes and capture same-machine
Release x64 baseline/candidate archives under `Specs/TestRuns/`. Do not emit one
JSONL row per key. Aggregate route count, handled/pass-through/blocked count,
total routing time, maximum routing time, and allocation count per bounded
batch/focus episode as `terminal.shortcut.route_batch_us` plus counters. Reuse
`terminal.render.frame_us` for scroll/zoom repaint. History instrumentation may
record entry counts, byte counts, and latency only; it must never record query
or command text.

Required scenarios:

1. **Unmapped/pass-through input:** 50,000 deterministic key messages after
   warm-up. Zero allocations per message; candidate median and p95 routing cost
   no more than 15% plus 2 us above the same-machine bypass baseline, and p95 no
   greater than 25 us.
2. **Held scrollback chord:** 1,000 repeat messages. Exactly one bounded
   viewport mutation per eligible repeat, no input/worker queue growth, no
   unbounded JSONL growth, and p95 route decision no greater than 50 us before
   paint.
3. **Zoom/reflow:** 100 increase/decrease/reset cycles at the standard terminal
   geometry. No retained-resource growth after settling; candidate
   `terminal.render.frame_us` p95 no more than 10% slower than its same-machine
   baseline at the same geometry/adapter.
4. **Command catalogue/palette:** 100 warm opens and 1,000 incremental filters
   across the complete catalogue. Emit aggregated
   `app.command_palette.open_to_visible_us`,
   `app.command_palette.filter_batch_us`, and
   `shortcut.catalog.rebuild_us`; Release x64 p95 open-to-visible is at most
   100 ms, filter-to-visible at most 16.7 ms, with no full-catalogue rebuild in
   paint and no retained-resource growth after close.
5. **Terminal Find:** search a deterministic 100,000-line primary buffer with
   1,000 incremental queries and repeated next/previous navigation. Emit
   `terminal.find.snapshot_batch_us`, `terminal.find.query_to_visible_us`,
   scanned-line/result counts, and cancellation counts. Release x64 p95 warm
   query-to-visible is at most 50 ms, UI-thread snapshot work stays below
   4 ms per batch, stale generations publish zero results, and closing Find
   releases all search snapshots/workers.
6. **History Suggestions:** open/filter against deterministic 10,000-entry,
   4 MiB history plus current-session entries. Emit
   `terminal.suggestions.load_batch_us`,
   `terminal.suggestions.open_to_visible_us`, and
   `terminal.suggestions.filter_to_visible_us`. Perform no file I/O on the UI
   thread; Release x64 p95 warm open is at most 100 ms and filter-to-visible at
   most 16.7 ms, with bounded cancellation and no command text in evidence.
7. **Floating lifecycle:** create the singleton, add/reorder/switch/restore 100
   deterministic tabs with 25 concurrently live, move/resize the root, then
   shell-exit and close every tab. Emit
   `terminal.floating.window_open_to_visible_us`,
   `terminal.floating.tab_open_to_visible_us`,
   `terminal.session.root_exit_to_host_close_us`,
   `terminal.session.exit_to_host_close_us`, placement/state-write coalesced
   counts, and retained root/tab/Terminal/callback/payload counts. Exactly one
   floating root may exist. Release x64 p95 warm window open is at most 250 ms,
   tab open at most 150 ms, and exit-to-tab-close at most 100 ms; settled
   move/resize/reorder produces one persisted write episode, restore preserves
   tab order/active/profile/path, and final retained live root/tab/Terminal/
   callback/payload counts are zero.

If a budget is infeasible because measurement overhead dominates, stop and
correct the instrumentation or obtain explicit review of a revised budget;
do not relabel a regression as diagnostic-only closeout evidence.

### Verification commands

Focused development checks:

```powershell
.\build.ps1 -ProjectName PluginContractTests -Platform x64 -Configuration Debug
.\.build\x64\Debug\PluginContractTests.exe --terminal-selftests
.\build.ps1 -ProjectName RedSalamander -Platform x64 -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter terminal_shortcut_ -TimeoutMultiplier 2
Invoke-Pester -Script .\Tools\Tests -PassThru
git diff --check
```

Closeout gate:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

Record exact run IDs, repository commit, machine hash, configuration, adapter,
sample counts, summaries, and archive paths. Raw local console success is useful
during development but is not the required performance or final closeout
evidence.

## Follow-up candidates, not P1 work

Promote each accepted item into a separate scoped plan after P1 evidence:

- profile-slot creation (`Ctrl+Shift+1..9`) and safe duplicate-to-other-pane
  (`Ctrl+Shift+D`);
- additional trusted history-provider adapters beyond the P1 PowerShell and
  explicit command-boundary providers; and
- mark mode (`Ctrl+Shift+M`) and real buffer clearing (`Ctrl+Shift+K`) if
  supported without shell injection.

The listed unimplemented chords pass through in P1 and must not be reserved
preemptively. Additional history adapters add capability only and do not claim
new default chords.

## Done criteria

- Every P1 decision above is implemented with configurable Terminal defaults,
  explicit pass-through, and no regression to current Folder/Navigation
  shortcut semantics.
- The F1 Helper exposes 46 separate Terminal binding rows backed by all 39
  Terminal-scope-eligible command definitions, while its Application group owns
  the three relevant global bindings and Function Bar retains F11 Connect; one
  keycap and conflict state per binding, search, activation, and accessibility
  are complete.
- Preferences enumerates all 39 Terminal-scope-eligible definitions from
  `CommandRegistry`, allows every factory binding and user alias to be changed,
  and offers explicit Pass through and No action choices without hard-coded
  command lists.
- The global Command Palette is command-centric and aggregates active aliases
  into shortcut chips on a single canonical command row.
- Global `Ctrl+Shift+P` opens the illustrated RedSalamander Command Palette from
  Folder, Navigation, Preview, or Terminal with shared catalogue/contextual
  ranking, correct enabled state, safe invocation, and exact focus return.
- F11 remains Salamander Connect, `Ctrl+,` opens global Preferences,
  `Ctrl+Shift+,` opens the active settings file in its default editor, and no
  `cmd/terminal/settings` definition exists.
- Terminal Find searches/navigates current buffer content without ConPTY input;
  history Suggestions use only trusted bounded providers, match the specified
  compact visual/interaction contract, and insert without executing.
- Exactly one floating Terminal window may exist, with as many independent tabs
  as the user opens. It restores normalized size/position, tab order, active
  tab, profile, and trusted path for every tab without changing the
  one-embedded-session-per-pane invariant.
- A real shell exit closes its embedded Terminal tab or corresponding floating
  tab exactly once after final output; the floating window closes when its last
  tab exits. Callback, payload, generation, manual-close, restore, and shutdown
  races are proven safe.
- Main-keyboard `Ctrl++` and `Ctrl+-` match physical scan positions `0x0D` and
  `0x0C` across layouts and never match by character value.
- The terminal input router proves security precedence, TUI/alternate-screen
  yield, AltGr/IME correctness, exact pass-through bytes, repeat/key-up symmetry,
  and teardown safety.
- Settings, Preferences, command metadata, import/export, localization, UIA,
  docs, and authoritative specs agree.
- Focused tests, every Full gate (the immutable aggregate plus exact post-fix
  revalidation of its three late failures), spec inventory, and same-machine
  archived performance gates are green with no unexplained skip.
- The upstream 68-binding review is refreshed at closeout and still accounts
  for every current default.
- This plan moves to `Specs/Plans/Done/` only after the durable contracts above
  contain the final behavior.

## STOP conditions

Stop and reconcile before implementation or closeout if:

- the I0 terminal remediation plan changes the same routing/ABI/lifecycle
  boundary;
- the `ITerminal` callback extension cannot be updated release-lockstep across
  every in-tree host/plugin consumer, or implementation would move
  selection/paste/security ownership into the host;
- the raw Terminal event callback cannot provide a clear barrier and plugin
  unload quiet point, or shell exit cannot be marshaled without destroying the
  child inline from its WndProc;
- confirmation state cannot be made authoritative before host dispatch;
- AltGr, IME, key-up symmetry, HWND/session reuse, or exact pass-through cannot
  be proven deterministically;
- the implementation would create more than one floating Terminal top-level
  window, cap normal user tab creation by product policy, or cannot restore the
  singleton placement and ordered tab/profile/trusted-path state safely;
- an action requires multiple independent terminals inside one physical pane, a
  global hotkey, raw Windows Terminal settings, rendered-output history
  scraping, hidden shell commands, command-text telemetry, automatic execution
  of a suggestion, or untrusted path-based placement;
- current upstream defaults have drifted without updating this review; or
- the focused/full/performance gates fail or their required evidence is
  unavailable.

## Maintenance

After closeout, any change to a canonical terminal shortcut, routing precedence,
pass-through semantics, command eligibility, or Terminal ABI must update the
authoritative Terminal, keyboard, and settings specs plus the consolidated
shortcut documentation and table-driven tests in the same change. Review the
upstream defaults when intentionally refreshing compatibility; do not track
Microsoft `main` automatically.
