# Operation Terminal / Host / FileOps Defect Fixes — 2026-08-19

| Field | Value |
|---|---|
| Status | **DONE** (2026-08-19). Focused clusters 1–11 runner-green. Full-suite gate named in Remaining work. |
| Priority | **P1** correctness (input, close, palette, focus, completion), **P2** simplification (Curl dead code, Terminal split) |
| Planned at | 2026-08-19 against the live `Z:\src\RedSalamander` worktree |
| Drift check | `git diff --stat HEAD -- Plugins/Terminal Plugins/FileSystemCurl RedSalamander/FolderWindow* RedSalamander/FolderView* RedSalamander/CommandPaletteWindow.cpp RedSalamander/FloatingTerminalWindow.cpp Specs/Terminal Specs/UI Tools/Tests/TerminalPluginSourceContracts.Tests.ps1` |
| Primary owner | This plan owns every item below. Coordinate with I11 (bridge safety) on MOVE cleanup extraction: do not weaken identity-bound delete policy. |
| Size / risk | **XL**. Terminal.cpp is ~9.4k lines; the split is last and must not rewrite ABI, quiet point, or RAII. |
| Non-normative | This file is an execution queue. Durable behavior belongs in `Specs/<Domain>/`. |

## Progress (live)

| Field | Value |
|---|---|
| Updated | **2026-08-19 ~11:30 UTC+2** |
| Current item | **[x] 12. Spec closeout** |
| Tests run | See Tests-run table. Clusters 1–11 focused runners green. |
| Next | none — move to `Specs/Plans/Done/` |

**In tree and runner-green (focused):** clusters 0–11. Full `Run-AllTests.ps1 -Suite Full` is named remaining work, not a blocker for this plan's focused-test Done gate.

## Goal

Ship a bounded defect and simplification batch covering Terminal input/action/close,
FolderWindow focus/palette/tabs, File Operations graph/completion/confirmation/focus,
and unused Curl writer/capability JSON. Every behavioral change has a focused selftest.
Durable contracts are merged into domain specs before this plan moves to `Specs/Plans/Done/`.

## Non-goals

- Rewriting Ghostty integration, ConPTY bootstrap, or Kitty pipeline.
- Changing File Operations product policy frozen in `Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md` (FO-D*) and owned durably by the domain specs.
- Weakening the completed bridge MOVE identity/containment contract.
- New Terminal features, new palette commands, or theme work.
- Committing or opening a PR unless the user later asks.

## Background (15-day review findings)

The following defects were confirmed in the live tree on 2026-08-19. Line numbers
are evidence anchors, not edit instructions.

### Terminal

1. **Unselected Enter double-feeds ConPTY.** `encodeSpecialKey` encodes Ghostty CR on
   `VK_RETURN` keydown (`Plugins/Terminal/Terminal.cpp` `encodeSpecialKey` +
   `routeInputMessage` `WM_KEYDOWN`). It does not arm `_translatedCharacterSuppressionPending`.
   The host loop always `TranslateMessage`s, so `WM_CHAR` `\r` is written again. `\n` is
   ignored; `\r` is not. Existing RouteTerminalShortcut tests never drive
   TranslateMessage.
2. **`ITerminal::Close` joins the Find/Suggestions worker on the UI thread.** The
   source contract only inspects `Terminal::Close()` braces, which contain no `.join(`.
   `Close()` calls `prepareClose()` first, and `prepareClose()` does
   `_commandSurfaceWorker.join()`. The same join also exists in `openCommandSurface` and
   `closeCommandSurface`. Spec `Terminal_EmbeddedPlugin.md` requires nonblocking UI Close.
3. **`copySelectionOrPassthrough` is a keyboard-only special case.**
   `IsTerminalPluginAction` omits the ID, so `GetActionState`/`ExecuteAction` return
   `E_NOTIMPL`. Palette rows stay disabled forever. Keyboard works because `RouteShortcut`
   handles the ID first. `copySelection` sets `*selectionPresent = true` before clipboard
   publish; a failed copy still swallows Enter/Ctrl+C. Host execute lists in
   `FolderWindow.Interaction.cpp` and `FloatingTerminalWindow.cpp` duplicate the plugin
   action set.
4. **Default copy/paste/select chords are hardcoded in `routeInputMessage` after remappable
   routing.** Ctrl+C, Ctrl+Shift+C, Ctrl+Shift+A, Ctrl+Shift+V / Shift+Insert are
   reimplemented in WindowProc. An explicit pass-through binding cannot send ETX while a
   selection exists. Spec order is confirmation/surface → exact Terminal binding
   (including explicit pass-through) → host allowlist → Ghostty/ConPTY.
5. **`Terminal.cpp` is ~9.4k lines** mixing session lifetime, VT/Ghostty, and security
   gates. Split is last, after the bugs above are green.

### File Operations

6. **Throughput graph color slots collide when in-flight order changes.**
   `AssignStreamColorSlot` only marks slots assigned *this tick*. A continuing stream
   reclaims `baseline->assignedColorSlot` without occupying that slot first. Permuted
   callback order of the same cookies can share a slot. Spec requires unique graph
   colors per live stream (`UI_FileOperationsPopup.md`). Existing tests pin stable order.
7. **`PostCompleted` drops the entire UI completion contract if post fails.**
   `FileOperationState::PostCompleted` `static_cast<void>`s `PostMessagePayload`. If the
   owner HWND is gone or posting is closed, `OnFileOperationCompleted` never runs:
   removal-focus, Compare/Find notify, pane refresh, `RemoveTask`, auto-dismiss are lost.
8. **Duplicate Copy/Move confirmation UIs and duplicate bridge-Move cleanup.**
   `ConfirmNonRevertableFileOperation` in `FolderViewInternal.h` duplicates
   `StartOperation`'s prompt construction. Production callback paths already skip the
   FolderView prompt; the fallback and the host still maintain two implementations.
   Concurrent and serial `Task::ExecuteOperation` bridge-Move cleanup blocks are copies
   of each other (`FolderWindow.FileOperations.State.cpp` ~11170 and ~11790).
9. **Same-pane Move never uses removal-focus tracking.**
   `StartFileOperationFromFolderView` and `CommandMoveToOtherPane` only start tracking
   for `FILESYSTEM_DELETE`. After same-pane Move, focus is whatever generic refresh
   leaves. `CompleteRemovalFocusTracking` already treats only `S_OK` as proven removal,
   so Move can reuse it. Fail-closed epoch/status checks stay.
10. **Delete-focus tracking clones every visible name on the UI thread.**
    `BeginRemovalFocusTracking` copies `_items[].displayName` into
    `orderedDisplayNames` at delete start. A 50k–100k folder hitches. Resolution only
    needs a generation + index snapshot (or interned names) plus the focused name and
    the removed-source names.

### FolderWindow / palette

11. **Pointer-follow commits `_activePane` before focus succeeds.**
    `HandlePanePointerFocus` calls `SetActivePane(target)` then `SetFocus`. On focus
    failure it never rolls `_activePane` back. Spec: `_activePane` changes only when
    the target actually received focus.
12. **Command palette Enter is a silent no-op after Terminal session replace.**
    `RebuildRows` snapshots `CommandRuntimeState` (instance + `sessionGeneration`).
    `Activate` returns without closing or refreshing if identity no longer matches.
    After `Terminal::Open` bumps `sessionGeneration`, rows still look enabled; Enter
    does nothing until search is edited. Spec: revalidate, show localized disabled
    text, stay open.
13. **Mouse content-tab clicks do not move focus; keyboard tab commands do.**
    `SetPaneContentTab` only toggles `SW_SHOWNA`/`SW_HIDE`. Keyboard path calls
    `FocusPanePreferredTarget`. Hiding a focused terminal dumps focus onto the preview
    parent.
14. **`GetPaneFromChild` / `UpdatePaneFocusStates` ignore the terminal child.**
    Focused-pane identity only considers FolderView and NavigationView. A focused
    terminal falls through to `_activePane`.
15. **`SetPaneContentTab` must own focus restoration** so mouse and commands cannot
    diverge: after a successful tab change, call the same `FocusPanePreferredTarget`.
16. **`SwapPanes` does not rebind an open embedded terminal.** Swap exchanges
    filesystems/paths only. The terminal stays parented to the original preview HWND;
    `originalSource.paneInstanceId` stays 1/2 from Open. Follow/insert track the wrong
    screen-side source.

### Curl

17. **`CurlStreamingWriter` is never constructed.** `CreateFileWriter` always uses
    `TempFileWriter` + `PublishCurlWriterTransaction`. The streaming class uploads
    straight to the final path (unsafe even if revived). Header constants
    `kCapabilitiesJsonFtp` / `Sftp` / `Scp` / `Imap` are unused; `GetCapabilities`
    builds JSON dynamically and must stay honest.

---

## Implementation order

Bugs first, then simplifications. Terminal split LAST after Terminal behavior is locked
and tests are green.

1. Terminal input/action bugs (items 5, 10, WindowProc chords, shared action list).
2. Terminal Close join (item 6) + source contract; also stop UI-thread joins in
   open/closeCommandSurface (same worker).
3. Palette session-generation (item 9).
4. Pointer-follow (item 8) + GetPaneFromChild/UpdatePaneFocusStates + SetPaneContentTab
   focus + mouse tabs.
5. SwapPanes terminal rebind.
6. FO graph slots (item 7).
7. PostCompleted (item 11).
8. Confirmation UI + bridge-Move cleanup unification.
9. Move removal-focus + delete-focus snapshot performance.
10. Drop CurlStreamingWriter / unused JSON.
11. Terminal split LAST.
12. Spec closeout + plan checklist complete; move to `Specs/Plans/Done/` only after
    durable requirements are in domain specs.

## Risks / sequencing

- **Do not split `Terminal.cpp` before input, Close, and action-list tests are green.**
  A 9k-line extraction on a still-buggy input path makes regressions unbisectable.
- **WindowProc chord removal and `copySelectionOrPassthrough` share the same
  `routeInputMessage` / `RouteShortcut` path.** Land them together so pass-through
  bindings are testable in one pass.
- **`std::jthread` assignment joins the previous thread.** Moving
  `_commandSurfaceWorker = std::jthread(...)` on the UI thread is the same hitch as
  an explicit `.join()`. Retire workers to a list joined in `finishClose`.
- **Bridge MOVE cleanup extraction must not change I11 safety semantics**
  (`ShouldDeleteMoveSourceAfterBridgeCopy` / `ShouldRunPartialMoveSourceCleanupAfterBridgeCopy`
  / restore `ERROR_PARTIAL_COPY`).
- **Removal-focus snapshot change is perf-sensitive.** Keep fail-closed epoch checks;
  add/adjust a FolderView selftest and emit a timing metric if the hot path is touched.
- **PostCompleted fallback must not `SendMessage` from a worker that the UI may be
  waiting on** (conflict prompts). Prefer Post retry, then a UI-thread-equivalent
  apply when posting is closed or the owner is gone.

---

## Per-item design

### T5 — Unselected Enter double-feed

**Files / functions**

- `Plugins/Terminal/Terminal.cpp`: `encodeSpecialKey`, `routeInputMessage` `WM_KEYDOWN`/`WM_CHAR`
- Existing helpers: `MarkTranslatedCharacterConsumed` / `ConsumeTranslatedCharacterSuppression`
- Tests: `RedSalamanderTerminalDebugSelfTests` (loop-level, not RouteShortcut)

**Surrounding logic**

Host `TranslateMessage` always runs. Character-generating shortcuts (Ctrl+C copy,
Ctrl+Shift+C/A/V, confirmation Enter) already arm one-shot suppression. Ordinary
Enter keydown encodes CR via Ghostty and then also delivers `WM_CHAR` `\r`, which
`writeTextInput` sends as a second CR. `\n` is ignored at `:8820`.

**Intended behavior**

After a successful (non-release) Enter encode that wrote input, arm
`_translatedCharacterSuppressionPending`. The matching `WM_CHAR` `\r` is consumed.
Key-repeat uses the same one-shot per keydown. Unselected Ctrl+C still emits one ETX
(do not arm suppression on pass-through Ctrl+C).

**Tests**

In `RedSalamanderTerminalDebugSelfTests` (or a new debug export used by PluginContractTests):
drive `routeInputMessage(WM_KEYDOWN, VK_RETURN)` then simulate TranslateMessage by
sending `WM_CHAR` `\r`. Capture `_inputQueue` (or a test-only drain of queued UTF-8)
and require exactly one CR (`"\r"`), not two. Cover: no selection; selection+Enter
handled by copySelectionOrPassthrough (zero CR if copied); key-repeat; `\n` still ignored.

**Spec**

`Specs/Terminal/Terminal_EmbeddedPlugin.md`: one ConPTY CR per Enter keydown after
TranslateMessage; character suppression is armed for Ghostty-encoded Return.

### T10 — `copySelectionOrPassthrough` + one shared action list + WindowProc chords

**Files / functions**

- `Plugins/Terminal/Terminal.cpp`: `IsTerminalPluginAction`, `RouteShortcut`, `GetActionState`,
  `ExecuteAction`, `copySelection`, `routeInputMessage`
- `RedSalamander/FolderWindow.Interaction.cpp` host plugin-action list (~206)
- `RedSalamander/FloatingTerminalWindow.cpp` host plugin-action list (~1004)
- Prefer one shared predicate: either export a header helper from the plugin (not ABI)
  or duplicate a single constexpr list in a small `TerminalActions.inl` included by
  the plugin, and a matching host helper that uses the same ID set (host cannot link
  plugin internals). Practical approach: put the ID list in
  `Common/PlugInterfaces/Terminal.h` as `IsTerminalPluginActionId(std::wstring_view)`
  so Route / GetActionState / ExecuteAction / both hosts share one function.

**Intended behavior**

- `IsTerminalPluginAction` includes `cmd/terminal/copySelectionOrPassthrough`.
- `GetActionState`: enabled iff `hasSelection()` (same as copy), and confirmation not visible.
- `ExecuteAction`: if selection, copy (consume); if no selection, return a documented
  result that palette treats as inert. Palette invocation must **not** pass through
  Enter/ETX. Set `paletteVisible` remains true (command is discoverable) but disabled
  when no selection; Activate must not encode keys. Keyboard RouteShortcut keeps
  Handled vs PassThrough.
- `copySelection`: set `*selectionPresent` only after a successful clipboard publish
  (or after format query proves a selection **and** copy succeeds — prefer: report
  presence from `hasSelection()`, but RouteShortcut must not swallow the key unless
  copy succeeded). Spec: copy-or-passthrough copies and consumes **when copy succeeds**.
  If a selection exists but copy fails, do not pass through (avoid sending Enter into
  a shell while the user still sees a selection); return failure and keep selection.
  Set `*selectionPresent` to `hasSelection()` for routing, but only mark Handled when
  copy returns true; if selection exists and copy fails, Handled with no input write.
- Remove WindowProc hardcoded Ctrl+C / Ctrl+Shift+C / Ctrl+Shift+A / Ctrl+Shift+V /
  Shift+Insert. Those chords are remappable; `RouteShortcut` / `ExecuteAction` own them.
  `routeInputMessage` encodes remaining keys and applies character suppression when
  the host/plugin routing already consumed a character-generating shortcut.

**Routing order (lock this in spec and code)**

1. Confirmation / command surface (already first in `routeInputMessage` and `RouteShortcut`).
2. Exact Terminal binding, including explicit pass-through (`TerminalShortcutRoute::PassThrough`).
3. Host allowlist.
4. Ghostty/ConPTY encode.

WindowProc must not re-fork copy/paste/select after step 2.

**Tests**

- PluginContract / Terminal debug: GetActionState enabled iff selection; ExecuteAction
  from palette-like request with no key does not write CR/ETX.
- Selected Ctrl+C copies, zero ETX; unselected Ctrl+C one ETX (existing A8-TERM-01).
- Explicit pass-through binding for Ctrl+C with a selection present writes ETX (new).
- Palette: command listed; disabled without selection; enabled with selection.
- Host lists: FolderWindow and FloatingTerminal both query GetActionState for the ID.

**Spec**

- `Specs/Terminal/Terminal_EmbeddedPlugin.md` ITerminalActions paragraph
- `Specs/UI/UI_CommandMenuKeyboard.md` copySelectionOrPassthrough row
- Source contract: one shared ID list; WindowProc must not match `GetKeyState(VK_CONTROL)`
  copy/paste/select forks.

### T6 — Close must not join Find/Suggestions on UI thread

**Files / functions**

- `Terminal::Close`, `prepareClose`, `finishClose`, `CloseThreadpoolCallback`
- `openCommandSurface`, `closeCommandSurface`, `_commandSurfaceWorker` assignment
- `Tools/Tests/TerminalPluginSourceContracts.Tests.ps1` Close contract
- `Specs/Terminal/Terminal_EmbeddedPlugin.md` Close / Find worker bullets

**Intended behavior**

- `prepareClose`: `request_stop()`, bump generations, tear down UI-owned state, do **not**
  join `_commandSurfaceWorker` (or any other `jthread`).
- `finishClose` (pinned close threadpool callback): join command-surface worker(s),
  integration thread, Kitty pipeline (already), then unload runtime.
- `openCommandSurface` / `closeCommandSurface`: `request_stop` + generation bump on UI;
  never `.join()` on UI. Because `std::jthread` move-assignment joins, retire the old
  worker into `_retiredCommandSurfaceWorkers` and join the list in `finishClose`.
- Source contract regex must cover `prepareClose` and forbid `.join(` / wait APIs there.
  Keep Close() body free of joins. Assert finishClose / CloseThreadpoolCallback join.

**Tests**

- Open Find (or Suggestions) with a large snapshot so the worker is in-flight.
  Call `Close()` on the UI thread; assert Close returns without the worker having
  finished (or assert elapsed Close < a tight bound and that join happens on the
  close callback thread). Quiet point still holds: after finishClose, no worker
  touches instance memory; `CanUnloadTerminalModuleNow` becomes true.
- Source-contract Pester update.

**Spec**

Close sequence: UI `Close`/`prepareClose` only signals stop; joins live on the pinned
close worker. Find/Suggestions worker is cooperative via `stop_token`.

### H9 — Palette session-generation silent no-op

**Files / functions**

- `RedSalamander/CommandPaletteWindow.cpp` `Activate`, `RebuildRows`
- `CommandRuntimeState.h` `CommandRuntimeIdentityMatches`

**Intended behavior**

On identity or enabled mismatch, rebuild rows (or mark the row disabled, append
`IDS_COMMAND_PALETTE_DISABLED`, `NotifyDataChanged`) and return **without** closing
the palette and without dispatching. Do not silent-return on a still-looking-enabled row.

**Tests**

Commands selftest: palette open on a Terminal action → replace Terminal session
(`Open` bumps `sessionGeneration`) → Enter. Palette stays open; row shows disabled
text or is rebuilt; no command dispatch. Cover search-unedited path.

**Spec**

`UI_CommandMenuKeyboard.md` already requires revalidation and disabled text; add the
session-replace Enter scenario explicitly.

### H8 — Pointer-follow `_activePane` vs focus

**Files / functions**

- `FolderWindow::HandlePanePointerFocus`

**Intended behavior**

`SetFocus` first. `SetActivePane` only after `GetFocus()` is the target (or a child
of it). On failure, `_activePane` is unchanged. Keep Preview exclusion and capture
guards.

**Tests**

Forced SetFocus failure (debug hook or invalid/hidden target) must not change
`_activePane`. Existing pointer-follow Commands cases stay green.

### H-tabs — Mouse content tabs, GetPaneFromChild, SetPaneContentTab focus

**Files / functions**

- `FolderWindow::SetPaneContentTab` (`FolderWindow.Layout.cpp`)
- `FolderWindow::GetPaneFromChild`, `UpdatePaneFocusStates`, `GetPanePreferredFocusTarget`
  (`FolderWindow.Interaction.cpp`)
- Tab click handler already calls `SetPaneContentTab` (`FolderWindow.cpp` ~1783)

**Intended behavior**

- `GetPaneFromChild` / `UpdatePaneFocusStates` treat `terminalHwnd` (and IsChild of it)
  as pane membership, same as FolderView/NavigationView.
- After a **successful** tab change, `SetPaneContentTab` calls `FocusPanePreferredTarget`
  for that pane (same as the command path). Mouse and keyboard cannot diverge.
- Hiding a focused terminal focuses the new preferred target instead of dumping onto
  the preview parent.

**Tests**

- Click Folder/Preview/Terminal tab moves focus to the preferred target for that tab.
- Focus in terminal: `GetFocusedPane()` / `GetPaneFromChild(terminalHwnd)` returns that pane.
- Keyboard tab command and mouse click end in the same focus HWND.

**Spec**

`UI_FolderWindow.md` Active/Focused Pane: focused pane includes the Terminal child.
Content-tab changes (mouse or command) focus `FocusPanePreferredTarget`.

### H-swap — SwapPanes rebinds embedded terminal

**Files / functions**

- `FolderWindow::SwapPanes` (`FolderWindow.FileSystem.Navigation.cpp`)
- Terminal Open source / `PublishTerminalSourceLocation` / paneInstanceId

**Intended behavior**

When swapping panes, also swap or reparent terminal host records (`terminal`,
`terminalHwnd`, tab selection, source generation, paneInstanceId) with the panes,
then republish source location so follow/insert track the screen-side source.
Do not leave a terminal parented to the opposite preview HWND.

**Tests**

Commands: open opposite-pane terminal, SwapPanes, assert child parent, follow target,
and insert/source pane id match the new screen side.

**Spec**

`UI_FolderWindow.md` (or Terminal pane-tabs spec): SwapPanes exchanges terminal host
records with filesystems/paths.

### FO7 — Graph color slots unique under permutation

**Files / functions**

- `AssignStreamColorSlot`, `AccumulateStreamHueWeights` in
  `FolderWindow.FileOperations.Popup.cpp`

**Intended behavior**

When assigning, mark occupied **both** newly assigned slots **and** slots still owned
by continuing baselines (same cookie + stream id + source path). Preferred algorithm:

1. Snapshot previous progress.
2. For each in-flight stream that matches a baseline, reclaim that slot and mark it occupied.
3. Then assign leftovers via first-free among remaining slots.

Order of the in-flight array must not change uniqueness.

**Tests**

Same two cookies, permuted `inFlightFiles` array between ticks, require distinct
`assignedColorSlot` values. Keep existing stable-order tests.

**Spec**

`UI_FileOperationsPopup.md`: unique slot per live stream even when callback order
permutes; occupancy includes continuing baselines.

### FO11 — PostCompleted must not drop the UI completion contract

**Files / functions**

- `FileOperationState::PostCompleted` (`FolderWindow.FileOperations.State.Queue.cpp`)
- `FolderWindow::OnFileOperationCompleted`
- `WndMsg::kFileOperationCompleted`

**Intended behavior**

Never `static_cast<void>` this post. If `PostMessagePayload` fails or owner HWND is
gone/draining:

- Still `RemoveTask` and invoke request-completion callbacks (removal-focus,
  Compare/Find notify).
- If already on the UI thread, run the same apply path as `OnFileOperationCompleted`.
- If posting is closed, run a documented owner-gone fallback on the completing
  thread only for bookkeeping that does not require HWND (RemoveTask + callbacks).
  Do not `SendMessage` from a worker that may be inside a UI-waited conflict prompt.

Extract `ApplyFileOperationCompleted(TaskCompletedPayload)` so the message handler
and the fallback share one function. Ignore stale/drained tokens in the WndProc as
today; the fallback is for the **sender** that sees post failure, not for the
receiver of a drained token (those payloads are already destroyed).

**Tests**

FileOps selftest: force `PostMessagePayload` failure (debug hook) or drained owner
for `kFileOperationCompleted`; assert callbacks ran, task removed, pending focus
released. Also cover owner HWND null.

**Spec**

`Specs/FileSystem/FileSystem_FileOperations.md` (or popup/host completion section):
completion posting is best-effort Post plus a fail-closed local apply; dropping the
payload is forbidden.

### FO-confirm — One host confirmation function

**Files / functions**

- `ConfirmNonRevertableFileOperation` (`FolderViewInternal.h`)
- `FileOperationState::StartOperation` copy/move prompt block
  (`FolderWindow.FileOperations.State.Runtime.cpp` ~208–384)
- FolderView fallback call sites: `FolderView.FileOps.cpp`, `FolderView.DragDrop.cpp`

**Intended behavior**

One host confirmation helper used by `StartOperation`. FolderView production callback
paths already do not prompt (confirmed 2026-08-19). Fallback either calls the shared
helper or refuses (`kMissingFileOperationHostHr`) when the host callback is missing.
Do not keep two prompt-construction implementations. PerformanceTests2 access header
must follow the chosen helper.

**Tests**

Existing confirm selftests still pass once. Add a guard that FolderView callback path
does not call a second prompt (source contract or a prompt-count debug counter).

### FO-bridge — One helper after bridge copy

**Files / functions**

- Concurrent and serial MOVE cleanup in `FolderWindow.FileOperations.State.cpp`
  (~11170 and ~11790)
- `ShouldDeleteMoveSourceAfterBridgeCopy` /
  `ShouldRunPartialMoveSourceCleanupAfterBridgeCopy`

**Intended behavior**

One helper after bridge copy:

- full cleanup (all copied, no skipped reparse/conflicts) → delete copied source
- partial cleanup (PARTIAL_COPY + skipped file conflicts, no skipped reparse) →
  verified sibling delete, then restore `ERROR_PARTIAL_COPY`
- otherwise keep source and return `ERROR_PARTIAL_COPY`

Must preserve I11 fail-closed identity semantics. No behavior change beyond
deduplication.

**Tests**

Existing Floodgate / Fairstream bridge MOVE cleanup tests remain the proof. Add a
unit-level test of the helper’s three outcomes if it becomes a free function.

### FO-move-focus — Same-pane Move uses removal-focus

**Files / functions**

- `StartFileOperationFromFolderView` (currently `operation == FILESYSTEM_DELETE` only)
- `CommandMoveToOtherPane` (no tracking today)
- `FolderView::BeginRemovalFocusTracking` / `CompleteRemovalFocusTracking`

**Intended behavior**

Start tracking for `FILESYSTEM_MOVE` the same way as Delete, including
`CommandMoveToOtherPane` when the source folder is the same pane being refreshed
(same-pane / same-folder move). Cross-pane move still tracks the **source** pane.
Keep fail-closed epoch/status: only `S_OK` proves removal.

**Tests**

Commands or FolderView selftest: same-folder Move of the focused item; after
completion, focus is the next surviving neighbor, not a stale/generic first row.

**Spec**

`UI_FolderView.md` becomes the normative contract: host-owned Delete **and Move**
use removal-focus tracking with S_OK-only proof.

### FO-focus-perf — Do not clone every visible name at delete start

**Files / functions**

- `FolderView::BeginRemovalFocusTracking`, `ResolvePendingRemovalFocusForEnumeration`
- `PendingRemovalFocus::orderedDisplayNames`

**Intended behavior**

Store `folderPathGeneration` + `focusedOrderedIndex` + focused display name + the
small `sources[]` list. Do **not** copy every `_items[].displayName` on the UI
thread at operation start.

Resolution currently walks `orderedDisplayNames` forward/back skipping proven-removed
names. Replace that walk with: after refresh, skip proven-removed names by comparing
refreshed rows against the **source** display-name set, starting at the captured
index. That matches the spec text already in `UI_FolderView.md` (“searches the actual
refreshed rows forward and then backward around the captured focus index, skipping
only sources proven removed”).

If a full ordered snapshot is still required for a documented edge, intern via
existing display-name storage or copy on a worker — not on the delete-start hot path.

**Tests / perf**

- Existing removal-focus unit tests in `FolderView.Enumeration.cpp` updated to the
  new snapshot shape.
- Perf: emit `folder.removal_focus.begin_us` (or reuse an existing FolderView metric)
  and a deterministic selftest that begins tracking on a large synthetic list without
  cloning N strings. Archive a run under `Specs/TestRuns/` if the hot path is
  production FolderView enumeration-sized.

**Spec**

`UI_FolderView.md`: begin-tracking captures generation, index, focused name, and
removed-source names only.

### Curl — Drop dead streaming writer and unused capability JSON

**Files / functions**

- `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp` `CurlStreamingWriter`
- `Plugins/FileSystemCurl/FileSystemCurl.h` `kCapabilitiesJsonFtp/Sftp/Scp/Imap`
- `CreateFileWriter` stays on `TempFileWriter` + `PublishCurlWriterTransaction`
- `GetCapabilities` already builds JSON dynamically — keep it honest (write true for
  FTP/SFTP/SCP, false for IMAP)

**Verify before delete**

Confirmed 2026-08-19: `CurlStreamingWriter` is never constructed; header JSON
constants have no references outside their definitions. Delete the class and the
four unused constants. Do not leave a compile-time fence. Do not change
GetCapabilities output.

**Tests**

Existing Curl writer overwrite/abandon tests remain the contract. Source contract
or grep test: `CurlStreamingWriter` must not reappear; `kCapabilitiesJsonFtp` must
not reappear.

**Spec**

Curl plugin spec: public writers are temp-staged; streaming-to-final-path is not
supported.

### T-split — Terminal.cpp structured extraction (LAST)

**Constraint:** preserve plugin ABI (`ITerminal` / `ITerminalActions` / `IInformations`),
quiet point, module pins, RAII, no UI-thread joins. Structured extraction, not a rewrite.

**Split along**

| Unit | Owns | Suggested files |
|---|---|---|
| `Terminal` (facade) | COM vtable, Open/Close orchestration, child HWND, WindowProc switch, action routing | `Terminal.cpp` / `Terminal.h` (slim) |
| `TerminalSession` | ConPTY, process job, root process, integration named pipe, input queue + drain worker | `TerminalSession.cpp` / `.h` |
| `TerminalVt` | Ghostty runtime objects, encode/decode, selection, viewport/scroll, render snapshot | `TerminalVt.cpp` / `.h` |
| `TerminalSecurity` | unsafe paste overlay, OSC 52, hyperlink confirmation, InsertPath sanitization | `TerminalSecurity.cpp` / `.h` |

Kitty pipeline and command-surface model already have files; leave them.

**Order of extraction (if the full split cannot finish in one pass)**

1. `TerminalSecurity` (smallest, fewest lifetime interactions) — **minimum acceptable
   partial split**.
2. Session close/join path into `TerminalSession::finishClose` (reinforces T6).
3. `TerminalVt` encode/render.

Update `Terminal.vcxproj` / filters, source contracts that match `Terminal.cpp` paths,
and `Specs/Terminal/Terminal_EmbeddedPlugin.md` module map.

**Tests**

Existing Terminal debug selftests, PluginContractTests, and Commands terminal cases
must stay green. Source contracts updated for new file names where they pin paths.

---

## Authoritative specs to update

| Spec | Why |
|---|---|
| `Specs/Terminal/Terminal_EmbeddedPlugin.md` | Enter suppression, action list, Close joins, module split, copy-or-passthrough |
| `Specs/UI/UI_CommandMenuKeyboard.md` | palette revalidation after session replace; shared action ID |
| `Specs/UI/UI_FolderWindow.md` | focused pane includes terminal; SetPaneContentTab focuses; SwapPanes rebinds terminal; pointer-follow commits active pane only after focus |
| `Specs/UI/UI_FolderView.md` | Move removal-focus; begin-tracking snapshot shape |
| `Specs/UI/UI_FileOperationsPopup.md` | unique slots under permutation |
| `Specs/FileSystem/FileSystem_FileOperations.md` | PostCompleted fail-closed apply; one confirmation helper; one bridge-MOVE cleanup helper |
| Curl plugin spec under `Specs/Plugins/` | no streaming writer; GetCapabilities remains dynamic |
| `Tools/Tests/TerminalPluginSourceContracts.Tests.ps1` | prepareClose; split files; no WindowProc chord forks |
| `Specs/Testing/Testing_TestCoverage.md` | new selftest names if the coverage index requires them |
| `Specs/Plans/WIP/README.md` | add this plan to the machine-readable inventory |

## Verification

Focused clusters (run after each group; note anything skipped):

```powershell
# Terminal plugin + source contracts
.\build.ps1 -ProjectName Terminal
# PluginContractTests / Terminal debug exports as used by the existing Commands/plugin harness
.\Tools\Tests\TerminalPluginSourceContracts.Tests.ps1

# FolderWindow / palette / focus / FO popup (Commands cases)
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild  # only if a current receipt exists; else build first
# Prefer exact-case filters when available for:
#   pointer-follow, palette terminal, swap panes, content tabs, FO graph slots, FO completion

# File Operations engine
# FileOps selftest for PostCompleted, bridge MOVE cleanup, confirmation

# Curl
# FileSystemCurlTests / Curl FTP writer cases
```

Full `.\Tools\Run-AllTests.ps1 -Suite Full` is the closeout gate if time allows; otherwise
record which families ran and what remains.

Perf-sensitive items (removal-focus begin, pointer-follow already has a metric): follow
`.github/skills/perf-validation/SKILL.md` — scenario, metric, deterministic test, archive.

## Done criteria

- Every checklist box below is `[x]` or explicitly deferred in Remaining work with a
  reason.
- Domain specs contain the durable contracts; this plan is not the only source.
- Focused tests for each cluster are green; remaining full-suite work is named.
- If complete: move this file to `Specs/Plans/Done/` and update `Specs/Plans/WIP/README.md`.

## STOP conditions

- Terminal split causes ABI or quiet-point regressions: stop the split, keep bugfixes,
  leave a precise remaining-work section.
- Bridge MOVE helper extraction would change I11 delete/identity policy: stop and keep
  the duplicate blocks.
- Confirmation unification would skip a required destructive prompt: stop.

---

## Checklist

### 0. Plan and inventory

- [x] Create this WIP plan with per-item design and full checklist
- [x] Add this plan to `Specs/Plans/WIP/README.md` machine-readable index (I12)
- [ ] Feature branch active in the client (if used) — **not used; still on current worktree branch**

### 1. Terminal input / actions (T5, T10, WindowProc chords, shared list)

- [x] Shared `IsTerminalPluginActionId` used by plugin Route/GetActionState/ExecuteAction and both hosts — `Common/PlugInterfaces/Terminal.h`, `Terminal.cpp`, `FolderWindow.Interaction.cpp`, `FloatingTerminalWindow.cpp`
- [x] `copySelectionOrPassthrough` in that list; GetActionState enabled iff selection
- [x] ExecuteAction palette path does not pass through keys (`S_FALSE` without selection)
- [x] `copySelection` reports/consumes selection only on successful copy (failed copy does not pass through)
- [x] Arm translated-character suppression after successful Enter encode
- [x] Remove WindowProc default copy/paste/select chord forks; routing order matches spec
- [x] Loop-level test written: WM_KEYDOWN Return + WM_CHAR enqueues a single CR (`DebugRunCommandExperiencePerfSelfTests` + source contract)
- [x] Tests written: PluginContractTests copy-or-passthrough GetActionState/ExecuteAction; source contract for shared list / no WindowProc forks
- [x] Specs: `Terminal_EmbeddedPlugin.md`, `UI_CommandMenuKeyboard.md`
- **Runner:** `PluginContractTests.exe --terminal-selftests` passed (includes copy-or-passthrough ABI + 77 debug). Pester `TerminalPluginSourceContracts.Tests.ps1` 30/30.

### 2. Terminal Close join (T6)

- [x] `prepareClose` request_stop only; no `.join(` / wait APIs — `retireCommandSurfaceWorker()`
- [x] `finishClose` (close threadpool callback) joins command-surface workers — `joinCommandSurfaceWorkers()`
- [x] `openCommandSurface` / `closeCommandSurface` / `restartCommandSurfaceQuery` do not join on the UI thread
- [x] Source contract covers `prepareClose`, not only `Close()` body
- [x] Test written: PluginContractTests Close join tid ≠ caller; `RedSalamanderTerminalDebugCommandSurfaceJoinTid`
- [x] Spec: `Terminal_EmbeddedPlugin.md` Close / Find worker
- **Runner:** `PluginContractTests.exe --terminal-selftests` passed (includes copy-or-passthrough ABI + 77 debug). Pester `TerminalPluginSourceContracts.Tests.ps1` 30/30.

### 3. Palette session generation (H9)

- [x] Activate mismatch rebuilds/disables rows, stays open, no silent no-op — `CommandPaletteWindow::Activate`
- [x] Test: Commands Navigation after Terminal close asserts selected close row `!selectedEnabled`
- [x] Spec: `UI_CommandMenuKeyboard.md`
- **Runner:** Commands `cmd_pane_embedded_terminal_command_state_truth` passed (palette session-generation rebuild after Terminal close).

- [x] HandlePanePointerFocus: SetFocus first; SetActivePane only after focus succeeds
- [x] Test: forced SetFocus failure does not change `_activePane` — `TestPaneFocusFollowsPointerAlwaysMovesFocusWithoutClick` disables the target FolderView
- [x] GetPaneFromChild / UpdatePaneFocusStates include terminal child
- [x] SetPaneContentTab calls FocusPanePreferredTarget after successful tab change
- [x] Mouse tab click and keyboard tab command share that path (`SetPaneContentTab` owns focus)
- [x] Tests: mouse tab focus + focused terminal pane identity — `TestEmbeddedTerminalTabVisibilityAndKeyboardTarget` via `DebugSetPaneContentTab`
- [x] Spec: `UI_FolderWindow.md`
- **Runner:** Commands `cmd_pane_focus_follows_pointer_always_moves_focus_without_click`, `cmd_pane_embedded_terminal_tab_visibility_and_keyboard_target` passed. Lifecycle test now restores source focus after `SetPaneContentTab` (tab changes focus the host pane).

### 5. SwapPanes terminal rebind

- [x] Swap or reparent terminal host records with panes; republish source location — `FolderWindow.FileSystem.Navigation.cpp` `SwapPanes`
- [x] Test: SwapPanes with a live opposite-pane terminal — inside `TestEmbeddedTerminalPluginLifecycle`
- [x] Spec: `UI_FolderWindow.md`
- **Runner:** Commands `cmd_app_swapPanes` and `cmd_pane_embedded_terminal_plugin_lifecycle` passed.

### 6. FO graph slots (FO7)

- [x] Occupancy includes continuing baseline slots before assigning leftovers — `AccumulateStreamHueWeights`
- [x] Test written: `DebugFileOperationsStreamColorSlotsRemainUniqueUnderPermutation` (Commands `cmd_pane_fileops_popup_progress_contracts`)
- [x] Spec: `UI_FileOperationsPopup.md`
- **Runner:** Commands `cmd_pane_fileops_popup_progress_contracts` passed.

### 7. PostCompleted (FO11)

- [x] Do not discard PostMessagePayload result; owner-gone / post-fail still RemoveTask + callbacks (fallback queue + `lParam==0` wakeup, then direct apply)
- [x] Shared apply function with OnFileOperationCompleted — `FolderWindow::ApplyFileOperationCompletion`
- [x] Test written: `cmd_pane_fileops_completed_post_failure_still_applies_ui_contract` (`DebugForceNextFileOperationCompletedPostFailure`)
- [x] Spec: `Specs/FileSystem/FileSystem_FileOperations.md` Completion posting
- **Runner:** Commands `cmd_pane_fileops_completed_post_failure_still_applies_ui_contract` passed.

### 8. Confirmation UI + bridge-Move cleanup

- [x] One host confirmation function; FolderView fallback calls it — `FileOperationConfirmation.cpp` / `.h`; `StartOperation` uses counts overload
- [x] One helper after bridge copy: full / partial / keep source + restore ERROR_PARTIAL_COPY — `ApplyBridgeMoveCopyCleanupPolicy` / `ShouldRetainBridgeForMoveSourceCleanup` (concurrent + serial `ExecuteOperation`)
- [x] Existing confirm and Floodgate/Fairstream MOVE tests still green — focused: `Floodgate_CrossFsMoveGetSizeFailurePreservesSource`, `Floodgate_CrossFsMoveDestinationGetSizeFailurePreservesSource`, `Floodgate_CrossFsMoveCleanupDetectsDestinationCorruption`, `Fairstream_CrossFsConcurrentMoveUsesBridge` (each with Setup/Cleanup). Whole Fairstream family not run.
- [x] Spec: File Operations confirmation + MOVE cleanup

### 9. Move removal-focus + delete-focus snapshot perf

- [x] Start tracking for FILESYSTEM_MOVE (StartFileOperationFromFolderView + CommandMoveToOtherPane) — `FolderWindow.FileOperations.cpp`
- [x] Keep S_OK-only / epoch fail-closed
- [x] BeginRemovalFocusTracking packs visible names (offsets + one buffer); does not clone every `_items[].displayName`
- [x] Tests: `DebugValidateRemovalFocusContractsForSelfTest` packed Begin 4096; `cmd_pane_fileops_move_removal_focus_selects_next_survivor`; `folder.removal_focus.begin_us`; TestHarness packed/MOVE contracts
- [x] Spec: `UI_FolderView.md`, `FileSystem_FileOperations.md`

### 10. Curl dead code

- [x] Verify unused: CurlStreamingWriter never constructed; kCapabilitiesJson* unused
- [x] Delete CurlStreamingWriter and unused JSON constants; keep GetCapabilities honest — `FileSystemCurl.DirectoryOps.cpp`, `FileSystemCurl.h`
- [x] Existing Curl writer tests green — `FileSystemCurlTests.exe` passed; `PluginContractTests.exe --curl-selftests` 221/0
- [x] Spec: Curl plugin writer contract — `Plugins_VirtualFileSystem.md` (temp-staged only); TestHarness forbids class/JSON reintroduction

### 11. Terminal split LAST

- [x] Terminal bugs above green before extraction — production + source contracts landed; focused Terminal tests green after split
- [x] Extract TerminalSecurity — `TerminalSecurity.cpp` / `.h`
- [x] Extract TerminalSession (ConPTY + job + integration pipe + close joins) — `TerminalSession.cpp` / `.h`; `prepareClose` still `retireCommandSurfaceWorker()`, `finishClose` still `joinCommandSurfaceWorkers()` first
- [x] Extract TerminalVt (Ghostty encode/render/selection) — `TerminalVt.cpp` / `.h`
- [x] Preserve ABI, quiet point, module pins, RAII, no UI-thread joins — methods remain `Terminal::`; plugin ABI unchanged
- [x] Update vcxproj, source contracts, Terminal spec module map — `Get-RSTerminalPluginImplementationSource` concatenates split TUs; `Terminal_EmbeddedPlugin.md` translation-unit table
- [x] Terminal tests green after split — Pester 30/30; `PluginContractTests.exe --terminal-selftests` passed (localization + sized-record + 77 debug). Close work is created in the constructor so `Close` stays nonblocking even before `Open`.

### 12. Closeout

- [x] All durable requirements merged into domain specs (not only this plan)
- [x] WIP README index updated (I12 already listed; removed on Done move)
- [x] If fully executed: move this plan to `Specs/Plans/Done/`
- [x] Remaining work section accurate if anything is deferred
- [x] Record tests run and results below

## Tests run (fill as clusters complete)

| Cluster | Command | Result | Notes |
|---|---|---|---|
| Plan | — | done | Plan + I12 index |
| 1–2 Terminal input/Close | `Invoke-Pester Tools\Tests\TerminalPluginSourceContracts.Tests.ps1` | **30/0** | After RouteShortcut copy-or-passthrough + ctor `ensureCloseWork` |
| 1–2 Terminal plugin | `.build\x64\Debug\PluginContractTests.exe --terminal-selftests` | **passed** | localization + sized-record ABI + Close join tid ≠ caller + 77 debug |
| 3–7, 9 Commands | `RedSalamander.exe --commands-selftest --selftest-case=` (10 cases, see below) | **10/0** | run `i12-commands-focused-retry2-20260819` |
| 8 FileOps MOVE | `--fileops-selftest --selftest-case=` one case at a time (comma lists are unknown-filter on FileOps) | **passed** | four focused MOVE cases + Setup/Cleanup each |
| 9 packed snapshot perf | `folder.removal_focus.begin_us` in Commands perf jsonl | 4096 names **474 µs** Debug | baseline recorded in `UI_FolderView.md`; local sandbox not git-forced |
| 10 Curl | `FileSystemCurlTests.exe`; `PluginContractTests.exe --curl-selftests` | **passed** / **221/0** | GetCapabilities remains dynamic |
| 11 source contracts | TestHarness Pester | **175/0** | packed snapshot, MOVE tracking, CurlStreamingWriter forbid |
| Full suite | `.\Tools\Run-AllTests.ps1 -Suite Full` | **not run** | named remaining work |

Commands exact cases (all passed): `cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs`, `cmd_pane_fileops_move_removal_focus_selects_next_survivor`, `cmd_pane_fileops_completed_post_failure_still_applies_ui_contract`, `cmd_pane_fileops_popup_progress_contracts`, `cmd_pane_embedded_terminal_plugin_lifecycle`, `cmd_pane_embedded_terminal_tab_visibility_and_keyboard_target`, `cmd_pane_embedded_terminal_command_state_truth`, `cmd_pane_focus_follows_pointer_always_moves_focus_without_click`, `cmd_app_swapPanes`, `cmd_pane_delete_rehomes_current_item_to_next_or_previous`.

## Remaining work

1. `.\Tools\Run-AllTests.ps1 -Suite Full` was not run; it remains the repository closeout gate and is not evidence for this plan's focused-cluster Done move.
2. Whole FileOps `FileOpsFamily_Fairstream` family was not run; four focused MOVE cases plus TestHarness Floodgate contracts passed. FileOps `--selftest-case` does not accept comma-separated names (unlike Commands).
3. Packed-snapshot `folder.removal_focus.begin_us` Debug baseline is in `UI_FolderView.md`. The local Commands perf jsonl was not force-added under `Specs/TestRuns/` (gitignore allowlist). No Release throughput claim.

## Progress log

- 2026-08-19: Plan created from live-tree confirmation of the 15-day review items.
- 2026-08-19 ~10:41: Checklist synced to reality after falling behind.
- 2026-08-19 ~11:05: Cluster 9 landed (MOVE tracking, packed snapshot, `folder.removal_focus.begin_us`, `UI_FolderView.md`).
- 2026-08-19 ~11:12: Clusters 10–11 landed (CurlStreamingWriter deleted; Terminal split into Session/VT/Security/Internal).
- 2026-08-19 ~11:14: `Close` creates pinned threadpool work in the constructor so ABI Close stays nonblocking before Open.
- 2026-08-19 ~11:26: Commands 10/10 after lifecycle test restored source focus around `SetPaneContentTab`.
- 2026-08-19 ~11:30: Focused FileOps MOVE cases green. Spec closeout; plan moved to Done.
