# Operation Iron Ledger — FolderView Data Integrity (Drag/Drop, Clipboard, Selection)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the FolderView **data-loss / wrong-target / crash / security** defects from `Specs/Reviews/FolderView-Audit.md`, re-verified against current HEAD, by fixing them at their shared root causes — not the performance hazards (those are owned by the sibling WarpDrive plan).

**Architecture:** Group fixes by the audit's cross-cutting root causes so one change closes several findings: one drop-point/self-drop chokepoint, one set of GlobalSize-bounded clipboard/OLE read helpers, one "establish the gesture's target before any destructive op reads it" discipline, one "report MOVE only after it happened" rule, plus two one-line hardening fixes. Single-gesture data-loss bugs ship first. Every task gates on **reproduction → minimal fix → deterministic regression test**, not on perf metrics.

**Tech Stack:** C++23, Win32, WIL RAII, OLE drag-drop (`IDropTarget`/`IDataObject`), CF_HDROP / `CFSTR_PREFERREDDROPEFFECT`, the FolderWindow file-operations bridge, Commands selftests, `Specs/TestRuns/` archived evidence, vcpkg/MSBuild via `build.ps1`.

---

## Status And Supersession

**Status:** WIP implementation plan.
**Created:** 2026-06-28.
**Audited HEAD:** `045a1f773`.
**Codename:** Iron Ledger.

> **Anchor drift (2026-07-02 folder review):** the `file:line` citations in this plan predate the
> 2026-06-28 auto-format (`45ae2a9a9`) and subsequent WarpDrive commits. Known drift:
> `FolderView.Interaction.cpp` ~+10 lines; `FolderView.Enumeration.cpp` `:1908` → `:1948`;
> `FolderView.Rendering.cpp` `:2176` → `:2295`; `FolderView.FileOps.cpp` `:267` → `:268`.
> Re-locate every anchor by its quoted text before editing.

This plan is the remediation track for the **Tier-1 / correctness** findings in
`Specs/Reviews/FolderView-Audit.md` and `Specs/Reviews/FolderView-Findings.md`. It is a
**sibling** of `Specs/Plans/WIP/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`
(performance), not a successor. The two do not overlap:

- **WarpDrive owns** the perf-relevant Tier-3 items: B1 (thumbnail block), B2/B3
  (virtual-FS thumbnail + WIC cap), B7 (device-loss), B9 (draw-loop brushes), B10
  (paste-shortcut off the UI thread), and the **fix** for R1 (worker torn-read, its G3).
- **Iron Ledger owns** the data-loss / wrong-target / crash / security cluster: D2, D3,
  D4, D5, D6, D7, D8, B4, B5, B6, B8, S1, S2, S3, plus six newly-found defects (N1–N3 and
  a write-side overflow). D1 is already mitigated on the production path (see Out Of Scope).

Created because the WarpDrive plan explicitly noted (its gap **G1**) that superseding the
*performance* backlogs does **not** address the audit's data-safety defects, several of
which are reachable by an ordinary single mouse gesture and lose data silently.

---

## Master Implementation Checklist

Single source of truth for *what to do and how to know it is done*. Each task has its own
per-step boxes; this is the roll-up. Work top to bottom by severity.

### Definition of Done (applies to every box)

A box may be checked only when **all** hold:

1. The change is in its **own task-scoped commit** (do not batch tasks).
2. Every PowerShell verification command in that task **exits `0`**.
3. The new/extended selftest **reproduces the defect before the fix** (red) and **passes
   after** (green) — record both in the Findings Log.
4. The evidence path (selftest trace / `Specs/TestRuns/<MachineHash>/.../`) is pasted into
   the Implementation Findings Log.
5. The box is checked with the commit short-hash, e.g. `- [x] Task 9 — DrawItem bounds (`a1b2c3d`).`
6. No previously-green selftest regressed (`.\Tools\Run-AllTests.ps1 -Suite Full` before closeout).

### Pre-flight

- [ ] **P0 — Clean build.** `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exits `0` on current HEAD (record it).
- [ ] **P1 — Read the audit + the verified status table below** so you fix the *current* sites, not the audit's stale line numbers.

### Phase 1 — Single-gesture data-loss / wrong-target (ship first)

- [ ] **Task 1 — Drop point + self/descendant chokepoint (RC-A; closes D2, D3, D1 defense-in-depth, B5 trigger).** Done when a drop onto a hovered subfolder lands *in that subfolder*, a background drop lands in the displayed folder, and a drop into a dragged directory's own descendant is rejected — all proven by new selftests.
- [ ] **Task 2 — Click/right-click collapse selection; disarm empty-background drag (RC-C; closes D5, D6, N3).** Done when a plain or right click on an *unselected* item collapses selection to it before any drag/menu reads it, and an empty-background drag starts nothing.

### Phase 2 — Untrusted-buffer hardening (adds shared test harness)

- [ ] **Task 3 — GlobalSize-bounded, count-clamped clipboard/OLE reads (RC-B; closes S1, S2, S3, N1).** Done when malformed CF_HDROP, a hostile internal-drop `pathCount`, an undersized preferred-drop-effect, and a huge Ctrl+V count all fail safe (no OOB, no `std::terminate`), proven by adversarial-buffer selftests. *(2026-07-02 folder review: Step 3's S3 guard already landed externally via WarpDrive B10 `b9af84d72` — see the Step 3 box and findings log; the S3 selftest is still owed here.)* **Adds `DebugPerformDropFromDataObjectForSelfTest` + `FolderViewDataObject(suppressInternalFormat)` — Tasks 4 and 7 depend on these, so land Task 3 before them.**

### Phase 3 — Stale-target / timing races

- [ ] **Task 4 — Route external imports through the local FS, not the destination pane FS (RC-A; closes B5).** Done when dropping a local file onto a virtual-FS pane reads the source with the builtin local provider (or rejects with a compatible-FS error), never using the pane FS as the source reader.
- [ ] **Task 5 — Gate every destructive destination on `IsCurrentFolderEnumerated()` (RC-C; closes D7, N2).** Done when a drop/paste/cross-pane-copy/move into an in-flight (not-yet-displayed) folder is rejected, not silently re-targeted.
- [ ] **Task 6 — Snapshot the menu/deferred-command target; kill positional focus migration (RC-C; closes D8 + deferred-external-command).** Done when an external change during an open context menu / between post and dispatch can no longer migrate a destructive command onto a neighbor.

### Phase 4 — Effect semantics + robustness one-liners

- [ ] **Task 7 — Report COPY to external OLE until complete; same-volume no-modifier = MOVE (RC-D; closes D4, B6).** Done when an external async-queued move echoes COPY to the source (our verified worker still moves and deletes on success), and a no-modifier same-volume drag selects MOVE.
- [ ] **Task 8 — Clamp plugin `GetCount` before `reserve()`; keep reserve in the worker `try` (RC-F; closes B4).** Done when a hostile plugin count degrades to a failed enumeration, not `std::terminate`.
- [ ] **Task 9 — Add `< _items.size()` to DrawItem's hovered subscript (RC-E; closes B8).** Done when a stale out-of-range `_hoveredIndex` cannot subscript OOB during paint.
- [ ] **Task 10 — Overflow-guard `BuildFileDropHGlobal` allocation size (RC-B; closes N-build).** Done when an overflowing CF_HDROP size request returns `nullptr` instead of an undersized buffer.
- [ ] **Task 11 — Clear/invalidate the clipboard after a completed move-paste (RC-B/RC-D; closes NS-1, harvested from Notes_Scratch 2026-07-02).** Done when a second Ctrl+V after a completed MOVE-effect paste cannot retry moving already-moved sources (`CFSTR_PERFORMEDDROPEFFECT` written / clipboard cleared per the OLE contract), proven by a red-before/green-after selftest.

> The detailed `- [ ]` boxes inside each Task remain authoritative for sub-steps. Keep the
> roll-up, the per-task boxes, and the final Acceptance Checklist in sync.

---

## Why This Plan Exists

FolderView trusts inputs and timing it does not control — the OLE data object, the
clipboard, the drop coordinate, the OS move-completion handshake, the async-enumeration
generation, and the live selection/focus — and acts **destructively** on them. The audit
found a tight cluster of confirmed data-loss, wrong-target, crash, and OOB defects. A
multi-agent verification pass (2026-06-28) re-confirmed **15 of 16** against current HEAD
and found **6 new** related defects. The fixes share six root causes; this plan attacks
the root causes.

---

## Cross-Cutting Root Causes

- **RC-A — Drop ignores gesture geometry and source/destination relationship.** `PerformDrop`
  never receives or hit-tests the drop point and never rejects self/same-parent/descendant/cross-FS
  targets. The one chokepoint that should resolve and validate the destination does not exist.
  *(D2, D3, D1 defense-in-depth, B5.)*
- **RC-B — Untrusted OLE/clipboard buffers parsed without `GlobalSize` bounds; attacker counts
  drive `reserve()` across `noexcept` boundaries.** The CF_HDROP walk, the internal-drop
  `pathCount` reserve, the preferred-drop-effect DWORD deref, and the Ctrl+V CF_HDROP reserve
  each lack the bounded pattern the codebase already uses elsewhere. *(S1, S2, S3, N1, N-build.)*
- **RC-C — Destructive ops re-resolve their target from live mutable state at execution time**
  (selection / `_focusedIndex` / in-flight `_currentFolder`) instead of a snapshot captured at
  gesture/menu time, and never gate on `IsCurrentFolderEnumerated()`. *(D5, D6, D7, D8, N2, N3.)*
- **RC-D — External OLE MOVE effect reported before the async worker verifies/completes**, and
  the no-modifier default effect is key-only (volume-blind). *(D4, B6.)*
- **RC-E — Stale/out-of-range indices dereferenced without the bounds clause sibling sites use.** *(B8.)*
- **RC-F — Plugin-reported enumeration count treated as a trusted allocation size.** *(B4.)*

---

## Verified Finding Status (re-checked against HEAD `045a1f773`, 2026-06-28)

| ID | Audit severity | Current status | Owner |
| --- | --- | --- | --- |
| D1 | data-loss | **mitigated on production path** by host overlap guard; FolderView + ENABLE_TESTS fallback still unguarded | Task 1 (defense-in-depth) |
| D2 | wrong-target | **present** — `PerformDrop` has no point/hit-test | Task 1 |
| D3 | data-loss | **present** (host guard catches the `_currentFolder` case only) | Task 1 (+ FS-layer noted out of scope) |
| D4 | data-loss | **present** — MOVE echoed at queue-time `S_OK` | Task 7 |
| D5 | wrong-target | **present** — click on unselected only `FocusItem`s | Task 2 |
| D6 | wrong-target | **present** — empty-bg drag carries stale focus | Task 2 |
| D7 | wrong-target | **present** — no `IsCurrentFolderEnumerated()` gate on drop/paste | Task 5 |
| D8 | wrong-target→data-loss | **present** — menu focus migration + positional fallback | Task 6 |
| B4 | robustness/DoS | **present** — unbounded `reserve` outside worker `try` | Task 8 |
| B5 | incorrect-behavior | **present** — external drop uses pane FS as source reader | Task 4 (+ Task 1 trigger) |
| B6 | incorrect-behavior | **present** — effect is key-only, volume-blind | Task 7 |
| B8 | robustness | **present** — `DrawItem` `_hoveredIndex` lacks bounds clause | Task 9 |
| S1 | security/crash | **present** — CF_HDROP walk, no `GlobalSize` bound | Task 3 |
| S2 | crash/DoS | **present** — `reserve(pathCount)` in `noexcept` | Task 3 |
| S3 | security | **fixed externally (2026-07-02 folder review)** — `GlobalSize` guard landed incidentally via WarpDrive B10 commit `b9af84d72` (2026-06-28), verified at `FolderView.FileOps.cpp:268-271` | Task 3 (Step 3 done; selftest still owed) |
| R1 | race | **present** | **WarpDrive G3** (verify-only here) |
| N1 | crash/DoS | **new** — `ReadFileDropClipboard` reserve from attacker count (Ctrl+V) | Task 3 |
| N2 | wrong-target | **new** — cross-pane Copy/Move (F5/F6) no enumerated gate | Task 5 |
| N3 | wrong-target→data-loss | **new** — right-click unselected item acts on stale selection | Task 2 |
| N-build | robustness | **new** — `BuildFileDropHGlobal` size overflow (write-side) | Task 10 |

---

## Out Of Scope (verified, not designed here)

- **D1 production path is already fixed.** `FolderWindow::FileOperationState::StartOperation`
  runs an unconditional per-source overlap guard (`FolderWindow.FileOperations.State.Runtime.Part.cpp:523-666`):
  `destinationItemText=JoinFolderAndLeaf(destFolder,leaf)` (:612), `IsSameOrChildPath(source,dest)`
  (:630), on match shows an error overlay and returns `S_FALSE` (:652-665) **without** calling
  `MoveItems`. `IsSameOrChildPath` (`FolderWindow.FileOperations.State.cpp:2029-2056`) returns
  true for identical paths. Every drop MOVE funnels here via `_fileOperationRequestCallback`
  (`FolderView.DragDrop.cpp:575`) → `StartFileOperationFromFolderView` → `StartOperation`. The
  FS chain (`MovePathInternal`, `MoveDirectoryMergeByRename`, `DeleteCopiedSourceEntryForMove`)
  is unsafe in isolation but **unreachable in production**. Residual gap (ENABLE_TESTS direct
  fallback at `DragDrop.cpp:607`) is covered by Task 1's UI-layer guard.
- **R1 fix** (worker torn-read of `_itemsFolder`) is owned by WarpDrive **G3**. Verified present
  here; no fix/test designed.
- **B1/B2/B3/B7/B9/B10** are owned by WarpDrive Tasks 1–4. Note only: Task 7's D4 fix must **not**
  block the UI thread on the worker (that hang is perf-owned).
- **FS-layer defense-in-depth for D3** (destination-subtree exclusion in `CopyDirectoryInternal`,
  same-path early-out in `MovePathInternal`) is worthwhile but not designed here; the UI-layer
  guard (Task 1) plus the existing host overlap guard are the primary mitigations. Reopen as a
  FileSystem-plugin task if desired.
- **Long paths (>MAX_PATH), junction/symlink loops, deep trees** — out of scope per the perf
  plan; not data-loss in single-folder view.

---

## Task 1: Plumb The Drop Point + Reject Self/Same-Parent/Descendant At One Chokepoint

**Root cause:** RC-A. **Severity:** wrong-target / data-loss (single gesture). **Closes:** D2, D3, D1 (FolderView defense-in-depth + ENABLE_TESTS fallback), B5 (routing trigger).

**Files:**

- Modify: `RedSalamander/FolderView.DragDrop.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`

**Repro:** Drag an external file over a pane showing `root/` (which contains `root/Archive`),
hover the `Archive` row, drop → today it lands in `root` (`DragDrop.cpp:572/603/607` always use
`_currentFolder`), not `root/Archive`. Separately, drag `root/Work` onto a pane whose
`_currentFolder` is `root/Work/Sub` → the FolderView layer issues a MOVE into its own descendant
(host guard catches the `_currentFolder` case; the ENABLE_TESTS direct fallback at :607 is unguarded).

**Fix summary:** Thread the client drop point into `PerformDrop`, hit-test the hovered directory
to pick the real destination (fall back to the live `_currentFolder`, never `_itemsFolder`), and
add one shared rejection of self/same-parent/descendant targets before dispatch so copy/move/link
all honor the gesture and never recurse into a dragged tree.

- [ ] **Step 1: Add a client `POINT` param to `PerformDrop`, forward the converted point from `DropTarget::Drop`.**
  Add `POINT clientPoint` to `FolderView::PerformDrop` (decl `FolderView.h:1361`, def `DragDrop.cpp:353`).
  In `DropTarget::Drop` (`DragDrop.cpp:126`) the screen-space `POINTL` is passed only to `_helper->Drop`
  at :142; add `POINT pt{point.x, point.y}; ScreenToClient(_owner.GetHWND(), &pt);` and pass `pt` to
  `_owner.PerformDrop` at :147 (leave the helper using the screen point — correct for `IDropTargetHelper`).
- [ ] **Step 2: Resolve the destination once from a hit-tested directory.**
  After the `!_currentFolder` early-out (`DragDrop.cpp:361`): `std::filesystem::path dest = *_currentFolder;
  if (auto h = HitTest(clientPoint); h && *h < _items.size() && _items[*h].isDirectory) dest = GetItemFullPath(_items[*h]);`
  Replace every destination use with `dest`: callback :572, `ConfirmNonRevertableFileOperation` :586,
  `CopyItems` :603, `MoveItems` :607, and the LINK loop `GenerateShortcutPath(*_currentFolder,...)` :651.
  The no-hit fallback MUST be `*_currentFolder` (live), **not** `_itemsFolder` (which `GetItemFullPath`
  composes from, `FolderView.cpp:131`, and which can diverge mid-enumeration).
- [ ] **Step 3: Add the shared self/same-parent/descendant rejection chokepoint.**
  After `dest` is resolved and after the `paths.empty()` check (`DragDrop.cpp:544`), before the MOVE/COPY
  switch: normalize `dest`; for each source reject (`*performedEffect=DROPEFFECT_NONE`, return
  `DRAGDROP_S_CANCEL`) when `dest` equals a source's `parent_path` (same-parent no-op), equals a dragged
  path, or — for directory sources — is a descendant of the source (case-insensitive prefix test with a
  separator boundary, copied from `FolderWindow.FileOperations.cpp:527-544`). This single guard sits
  before the switch so it covers the callback (:564-577) and the ENABLE_TESTS direct fallback (:586-608).
- [ ] **Step 4: Extend the debug entry point to carry a `POINT`.**
  Extend `DebugPerformFileDropForSelfTest` (`DragDrop.cpp:261`, decl `FolderView.h:235-237`) with a
  `POINT clientPoint` defaulted to off-screen `{-1,-1}` (= background = `_currentFolder`), forwarded to
  `PerformDrop` at :278. Existing call sites compile unchanged via the default.

**Regression tests:**

- `cmd_pane_dragDrop_onto_subfolder_targets_subfolder` — CALLBACK path (destination observable via
  `RequireQueuedShellFileOperationTask`). Seed `root` + `root/Archive` + external `src/drop-alpha.txt`;
  resolve the `Archive` row center via `DebugGetColumnLayoutSnapshot` + `DebugVisibleItemEntry`, invert
  `HitTest`'s transform to a client `POINT`, drop `DROPEFFECT_COPY`; assert destination == `root/Archive`.
- `cmd_pane_dragDrop_onto_background_targets_current_folder` — same setup, off-screen `{-1,-1}`; assert
  destination == `root` (pins the `_currentFolder` fallback).
- `cmd_pane_dragDrop_into_own_descendant_is_rejected` — pre-create `.../Work/Sub/Work` + `Work/keep.txt`;
  `SetFolderPath(.../Work/Sub)`; drop `{.../Work}` MOVE; assert `performed==DROPEFFECT_NONE`, no new task,
  source unchanged, Error overlay shown.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_dragDrop_onto_subfolder_targets_subfolder --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_dragDrop_onto_background_targets_current_folder --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_dragDrop_into_own_descendant_is_rejected --selftest-timeout-multiplier=4
```

---

## Task 2: Click/Right-Click Collapse Selection; Disarm Empty-Background Drags

**Root cause:** RC-C. **Severity:** wrong-target / data-loss (single gesture). **Closes:** D5, D6, N3.

**Files:**

- Modify: `RedSalamander/FolderView.Interaction.cpp`
- Modify: `RedSalamander/FolderView.Menus.cpp`
- Modify: `RedSalamander/FolderView.DragDrop.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`

**Repro:** **D5** — select A,B,C; plain-click unselected Z and drag → FolderView drags A,B,C
(`FocusItem` at `Interaction.cpp:479` never clears `.selected`). **D6** — focus an item, click empty
background (`Interaction.cpp:485` clears selection but not `_focusedIndex`/`dragging`), drag past
`SM_CXDRAG` → an OLE drag of the previously-focused item starts. **N3** — select A,B,C; right-click
unselected Z, pick Delete/Move → A,B,C are recycled/relocated (`OnContextMenu` only `FocusItem`s at
`Menus.cpp:354`; handlers act on `.selected`).

**Fix summary:** Make every pointer gesture establish the selection it visually implies before any
destructive command or drag reads it: collapse to the clicked item when it is not already selected,
and fully disarm the drag and stale focus on empty background.

- [ ] **Step 1: Plain click on unselected item collapses selection (D5).**
  In `OnLButtonDown`, replace the no-modifier else-branch at `Interaction.cpp:478-480` with
  `if (!_items[*hit].selected) { SelectSingle(*hit); } else { FocusItem(*hit,false); } _anchorIndex = *hit;`.
  `SelectSingle` (`Selection.cpp:3`) clears all other `.selected`. Do NOT touch the shift (:464-471) or
  ctrl (:472-476) branches. (No marquee exists in FolderView, so zero marquee-regression risk.)
- [ ] **Step 2: Right-click on unselected item collapses selection before building the menu (N3).**
  In `OnContextMenu` (`Menus.cpp:352-356`), when the hit item is not already selected, call
  `SelectSingle(*hit)` instead of `FocusItem(*hit,false)`; keep `_anchorIndex=*hit`. Preserve the
  multi-selection only when `*hit` is already selected (Explorer semantics). `UpdateContextMenuState`
  and the Cmd handlers already read `.selected`, so no handler change is needed.
- [ ] **Step 3: Disarm empty-background drags (D6).**
  In the no-hit else branch (`Interaction.cpp:483-487`) after `ClearSelection()`, add
  `_drag.dragging = false; _focusedIndex = static_cast<size_t>(-1); _anchorIndex = static_cast<size_t>(-1);`.
  Belt-and-suspenders: at the top of `BeginDragDrop` (`DragDrop.cpp:201`), before `GetSelectedOrFocusedPaths`
  at :203, add `if (_drag.anchorIndex == static_cast<size_t>(-1)) { return; }`. Do NOT alter
  `GetSelectedOrFocusedPaths`/`ClearSelection` (focus fallback is intended for keyboard/command paths).

**Regression tests:**

- `cmd_pane_fileops_plainClick_collapses_selection_for_drag` — ctrl-click A,B,C; assert 3 selected;
  no-modifier click Z; assert 1 selected and a new `DebugGetDragSourcePaths(Left)` accessor (returning
  `GetSelectedOrFocusedPaths()`) == `{folder/Z.txt}`. Pre-fix returns `{A,B,C}`.
- `cmd_pane_fileops_empty_background_drag_does_not_drag_focused_item` — focus an item, `WM_LBUTTONDOWN`
  on empty background, `WM_MOUSEMOVE` past threshold; a DEBUG `_debugBeginDragDropEntryCount` +
  `DebugGetBeginDragDropEntryCount()` must be unchanged and `GetSelectedOrFocusedPaths()` empty. A
  positive case asserts a real-item drag still increments the counter.
- `cmd_pane_fileops_contextMenu_unselected_item_collapses_selection` — select a/b/c; invoke
  `OnContextMenu` at z's hit point; trigger CmdMove with `DebugSetNextMoveSelectedItemsDestinationForSelfTest`;
  assert only z is the operation target and a/b/c remain.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_fileops_plainClick_collapses_selection_for_drag --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_fileops_empty_background_drag_does_not_drag_focused_item --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_fileops_contextMenu_unselected_item_collapses_selection --selftest-timeout-multiplier=4
```

---

## Task 3: GlobalSize-Bounded, Count-Clamped Clipboard/OLE Read Helpers

**Root cause:** RC-B. **Severity:** security / crash (OOB read + `std::terminate` DoS). **Closes:** S1, S2, S3, N1.

**Files:**

- Modify: `RedSalamander/FolderView.DragDrop.cpp`
- Modify: `RedSalamander/FolderView.FileOps.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`

**Repro:** **S1** — external Explorer drag with a `DROPFILES` whose `pFiles` points past the
allocation faults at `DragDrop.cpp:532-536` (`GlobalSize` never called in this branch). **S2** — a data
object exposing `RedSalamander.InternalFileDrop.V1` with a 20-byte buffer (`version=1`,
`pathCount=0xFFFFFFFF`) reaches `reserve(0xFFFFFFFF)` at :460 → throw across `noexcept` → `std::terminate`.
**S3** — clipboard owner sets `CFSTR_PREFERREDDROPEFFECT` to a sub-DWORD HGLOBAL; Ctrl+V reads OOB at
`FileOps.cpp:267`. **N1** — a peer process places an oversized CF_HDROP of many tiny names; Ctrl+V hits
`reserve(huge)` at `FileOps.cpp:221` → `std::terminate`.

**Fix summary:** Apply the GlobalSize-bounded / count-clamped pattern the codebase already uses
(`DragDrop.cpp:405-409`, `ReadFileDropClipboard`'s `DragQueryFileW` loop) to every untrusted
OLE/clipboard read path.

- [ ] **Step 1 (S1): replace the hand-rolled `DROPFILES` walk with `DragQueryFileW`.**
  In `PerformDrop`'s CF_HDROP fallback, delete the `GlobalLock`/`fWide`/`pFiles`/`wcslen` walk at
  `DragDrop.cpp:520-537`; inside the existing tymed/null guard (:515-518) use the proven
  `ReadFileDropClipboard` pattern (`FileOps.cpp:220-235`): `DragQueryFileW(hdrop,0xFFFFFFFFu,...)` count,
  then per-index length-safe copies. Clamp the reserve as in Step 4.
- [ ] **Step 2 (S2): clamp/remove the internal-drop `pathCount` reserve inside the `noexcept` fn.**
  In `tryReadInternalDrop` (`DragDrop.cpp:377` noexcept), replace `paths.reserve(header->pathCount)` (:460)
  with `const size_t remaining = (offset <= bytesAvailable) ? (bytesAvailable - offset) : 0u;
  if (header->pathCount > remaining / sizeof(uint32_t)) return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
  paths.clear(); paths.reserve(header->pathCount);` (or delete the reserve — the loop at :461-482 already
  bounds-checks and grows).
- [x] **Step 3 (S3): add a `GlobalSize` guard before the preferred-drop-effect deref** (externally completed — WarpDrive B10 commit `b9af84d72`).
  *2026-07-02 folder review: landed incidentally by the WarpDrive B10 commit `b9af84d72` (2026-06-28). Verified
  at `FolderView.FileOps.cpp:268-271`: `if (GlobalSize(handle) < sizeof(DWORD)) return std::nullopt;`. No
  Iron Ledger code change remains; the `TestPaneClipboardPasteIgnoresUndersizedPreferredDropEffect` regression
  selftest below is still owed under this task (green-only pin — red-before is no longer reproducible).*
  Original design (superseded by the landed fix): in `ReadPreferredDropEffectClipboard` (`FileOps.cpp:241`, noexcept), after `GlobalLock` (:261) and before
  `const DWORD result = *effect;` (:267), insert `if (GlobalSize(handle) < sizeof(DWORD)) { GlobalUnlock(handle); return std::nullopt; }`.
  The sole caller (:738) already `.value_or(DROPEFFECT_COPY)`, so an undersized buffer degrades to COPY.
- [ ] **Step 4 (N1): clamp the Ctrl+V CF_HDROP reserve.**
  In `ReadFileDropClipboard` (`FileOps.cpp:205`, noexcept free fn), replace `result.reserve(fileCount);` (:221)
  with `result.reserve(std::min<UINT>(fileCount, 65536u));` (or drop it). Mirrors the `std::min` reserve caps
  in `FolderView.Icons.cpp:283/326`.
- [ ] **Step 5: add debug entry points that inject raw/malformed buffers.**
  Add `FolderView::DebugPerformDropFromDataObjectForSelfTest(IDataObject*, DWORD effect, DWORD* performed)`
  (decl `FolderView.h` near :235) wrapping an arbitrary `IDataObject` into `PerformDrop`, so tests can feed
  a malformed CF_HDROP and an internal-format-only object with a hostile `pathCount`. **This and the
  `FolderViewDataObject(suppressInternalFormat)` ctor flag are shared by Tasks 4 and 7.**

**Regression tests:**

- `TestFolderViewDropRejectsMalformedHdrop` — `BuildMalformedHDropForShellCommandTest` producing (a)
  `pFiles` past the allocation, (b) alloc `< sizeof(DROPFILES)`, (c) no in-buffer double-NUL; each via
  `DebugPerformDropFromDataObjectForSelfTest`; assert `performed==DROPEFFECT_NONE` / no task and tree
  untouched (faults pre-fix under App Verifier/ASan).
- `TestFolderViewInternalDropRejectsHugePathCount` — 20-byte internal buffer `pathCount=0xFFFFFFFF` via a
  CF-only `IDataObject`; require `ERROR_INVALID_DATA`, no crash. Pre-fix the reserve terminates the harness
  (process death = the signal). Add `TestFolderViewInternalDropAcceptsValidPaths` to prove no regression.
- `TestPaneClipboardPasteIgnoresUndersizedPreferredDropEffect` — seed valid CF_HDROP, then
  `SetClipboardData(CFSTR_PREFERREDDROPEFFECT, 1-byte 0x02)`; paste via `IDM_PANE_CLIPBOARD_PASTE`; assert
  COPY semantics (sources still exist AND copies appear). Do not read via the test helper (same unchecked deref).
- `TestPaneClipboardPasteSurvivesHugeHdropCount` — CF_HDROP packed with many single-char names; paste;
  assert the process survives and the paste no-ops/proceeds bounded.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=TestFolderViewDropRejectsMalformedHdrop --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=TestFolderViewInternalDropRejectsHugePathCount --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=TestPaneClipboardPasteIgnoresUndersizedPreferredDropEffect --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=TestPaneClipboardPasteSurvivesHugeHdropCount --selftest-timeout-multiplier=4
```

---

## Task 4: Route External CF_HDROP Imports Through The Local FS, Not The Destination Pane FS

**Root cause:** RC-A. **Severity:** incorrect-behavior (wrong source reader). **Closes:** B5. **Depends on:** Task 3 harness primitives.

**Files:**

- Modify: `RedSalamander/FolderWindow.FileOperations.cpp`
- Modify: `RedSalamander/FolderViewInternal.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`

**Repro:** Activate a virtual/archive/cloud plugin on the destination pane, drag a local `C:\` file from
Explorer onto it. `tryReadInternalDrop` returns `S_FALSE`, so `request.sourceContextSpecified=false`
(`DragDrop.cpp:569`) skips the entire source-context block (`FolderWindow.FileOperations.cpp:386-471`);
`fileSystem` stays `= destinationState.fileSystem` (:383) and the virtual FS is used as the SOURCE reader
for the local path. The gate at :474 passes for any provider advertising copy/move, so the misroute
proceeds silently.

**Fix summary:** Detect external (non-context-specified) copy/move imports and read the local source with
the builtin local provider while writing through the destination pane FS, rejecting with a compatible-FS
error when the provider cannot import.

- [ ] **Step 1: Add an explicit external-import branch in `StartFileOperationFromFolderView`.**
  After `isCopyMove` (`FolderWindow.FileOperations.cpp:379`) and before the same-FS gate (:474), branch on
  `isCopyMove && !request.sourceContextSpecified`. Resolve the builtin local provider via
  `FileSystemPluginManager::GetInstance()...FindPluginById(L"builtin/file-system")` as `localSource`. If the
  destination pane IS `builtin/file-system`, leave behavior unchanged. Otherwise treat as cross-FS import:
  `fileSystem=localSource`, `destinationFileSystem=destinationState.fileSystem`, gate with
  `CanCrossFileSystemCopyMove(localSource, L"builtin/file-system", destinationState.fileSystem, destinationState.pluginId, request.operation)`
  (:288-320). On false, reject with `ShowAlertOverlay(...IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS)` (mirror
  :479-482) and return `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)`. Mirrors the internal cross-FS branch at
  :464-469; no `FolderView.DragDrop.cpp` change is needed.

**Regression tests:**

- `Commands_FileOpsExternalDropOntoVirtualFsImportsLocally` — prerequisite: add a `bool suppressInternalFormat`
  ctor flag to `FolderViewDataObject` so `QueryGetData`/`GetData` (`FolderViewInternal.h:1303-1368`) return
  `DV_E_FORMATETC` for the internal format, plus a `DebugPerformDropFromDataObjectForSelfTest` overload using
  it with `includeHDrop=true` (so `internalDrop=false`). Activate a virtual test plugin on the destination
  pane; drop a local `C:\` source; assert EITHER (a) provider supports importCopy → task whose source FS is
  the builtin local provider and dest FS is the pane FS, OR (b) rejected with `ERROR_NOT_SUPPORTED` + Error
  overlay — never the pane FS as the source reader for local paths.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=Commands_FileOpsExternalDropOntoVirtualFsImportsLocally --selftest-timeout-multiplier=4
```

---

## Task 5: Gate Every Destructive Destination On A Settled (Enumerated) Folder

**Root cause:** RC-C. **Severity:** wrong-target / data-loss (timing race). **Closes:** D7, N2.

**Files:**

- Modify: `RedSalamander/FolderView.DragDrop.cpp`
- Modify: `RedSalamander/FolderView.FileOps.cpp`
- Modify: `RedSalamander/FolderWindow.FileOperations.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`

**Repro:** **D7** — `SetFolderPath` assigns `_currentFolder` synchronously (`FolderView.cpp:254`) then
starts async enumeration; `_displayedFolder` updates only on completion (`Enumeration.cpp:1672`) and is not
reset on A→B navigation, so during a slow nav the pane shows A while `_currentFolder=B` and a drop/paste
lands in B. **N2** — with the inactive (destination) pane mid-navigation to a slow folder, F5/F6 resolves
the destination as the in-flight `_currentFolder` (`FolderWindow.FileOperations.cpp:1238/1279`); a MOVE
silently relocates into the not-yet-shown folder.

**Fix summary:** Reject (not re-target) any drop, paste, or cross-pane copy/move whose destination is not
the settled, displayed folder, reusing the existing `IsCurrentFolderEnumerated()` guard.

- [ ] **Step 1: Gate `PerformDrop` (D7 drop).** After the `!_currentFolder` block (`DragDrop.cpp:361-364`),
  add `if (!IsCurrentFolderEnumerated()) { *performedEffect=DROPEFFECT_NONE; return DRAGDROP_S_CANCEL; }`
  before destination/data resolution. `IsCurrentFolderEnumerated()` exists (`FolderView.cpp:291-303`).
  **If both Task 1 and Task 5 land, place this gate before Task 1's dest resolution.**
- [ ] **Step 2: Gate `PasteItemsFromClipboard` (D7 paste).** After the `!_currentFolder || !_fileSystem`
  block (`FileOps.cpp:733-736`), add `if (!IsCurrentFolderEnumerated()) { return; }`.
- [ ] **Step 3: Gate cross-pane Copy/Move (N2).** In `SanityCheckBothPanes` (called by `CommandCopyToOtherPane`
  :1238 and `CommandMoveToOtherPane` :1279 at :1214/:1255), add
  `if (!dest.folderView.IsCurrentFolderEnumerated()) { /* reject via the existing IDS_MSG_PANE_OP_* overlay */ return ...; }`
  so its same-folder comparison (:1161-1164) and the destination both use a settled folder.

**Regression tests:**

- Add a `DebugSetCurrentFolderWithoutEnumerationForSelfTest(path)` hook (assigns `_currentFolder=B`, leaves
  `_displayedFolder=A`, so `IsCurrentFolderEnumerated()` is false).
- `cmd_pane_drop_into_unenumerated_folder_rejected` — set A; hook `_currentFolder=B`; drop; assert
  `DRAGDROP_S_CANCEL`, `DROPEFFECT_NONE`, no task, no file under B.
- `cmd_pane_paste_into_unenumerated_folder_rejected` — same state; paste; assert no task, no file under B.
- Cross-pane case (extend `FolderWindow.FileOperations.SelfTest.Phases05_06.cpp`) — Right pane `_currentFolder=B`
  but `_displayedFolder=A` (withhold `kFolderViewEnumerateComplete`); `CommandCopyToOtherPane(Left)`; assert the
  task's destination == A (displayed) OR the command is rejected.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_drop_into_unenumerated_folder_rejected --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_paste_into_unenumerated_folder_rejected --selftest-timeout-multiplier=4
```

---

## Task 6: Snapshot The Context-Menu / Deferred-Command Target; Stop Positional Focus Migration

**Root cause:** RC-C. **Severity:** wrong-target → data-loss (timing race). **Closes:** D8, deferred-external-command (`Enumeration.cpp:1908`).

**Files:**

- Modify: `RedSalamander/FolderView.Menus.cpp`
- Modify: `RedSalamander/FolderView.Enumeration.cpp`
- Modify: `RedSalamander/FolderView.FileOps.cpp`
- Modify: `RedSalamander/FolderWindow.FileOperations.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

**Repro:** Right-click an item (`OnContextMenu` captures no snapshot, `Menus.cpp:352-356`; command posted by
id at :374); while the menu is open, an external change posts `kFolderViewEnumerateComplete` which the nested
DxUi menu loop pumps → `ProcessEnumerationResult` refresh → `ApplyCurrentSort` positional fallback
(`Enumeration.cpp:1324-1326`) migrates `_focusedIndex` to a different file at the old slot; the posted
`CmdDelete` re-resolves the migrated focus (`FolderWindow.FileOperations.cpp:1057`) and recycles the wrong
item with **no confirmation**. Deferred path: `QueueCommandAfterNextEnumeration` validates by name then
`PostMessageW` (`Enumeration.cpp:1908`); a second debounced refresh migrates focus before dispatch.

**Fix summary:** Capture the menu/automation target identity at the gesture and act only on a still-matching
snapshot, and stop the refresh-time positional focus fallback that silently re-points focus to a neighbor.

- [ ] **Step 1: Re-validate the menu target by name before posting destructive commands.**
  In `OnContextMenu`, after `FocusItem` (:354) snapshot intended targets via `GetSelectedOrFocusedDisplayNames()`
  (`Selection.cpp:993`). After `ContextMenu::Show` (:367) and before `PostMessageW` (:374), re-resolve the
  snapshot names against current `_items`; if the set differs (an item migrated during the nested loop),
  suppress the post for target-bound destructive commands (CmdDelete, CmdMove, CmdRename, CmdCopy). Navigation
  commands still post.
- [ ] **Step 2: Stop `ApplyCurrentSort`'s positional focus fallback on a refresh.**
  In `ApplyCurrentSort` (`Enumeration.cpp:1318-1331`), when the named focus target disappeared during a refresh
  (`newFocusedIndex==invalidIndex`, nothing selected), do NOT fall back to the raw positional index (:1324-1326)
  — leave focus cleared so a later destructive command resolves to an empty path set and no-ops. Keep the
  positional fallback only for the navigation case (distinguish via the existing `fallbackFocusIndex` signal).
- [ ] **Step 3: Eliminate the post-then-execute window for deferred external commands.**
  In `ProcessEnumerationResult`'s pending-external-command block (`Enumeration.cpp:1884-1909`), after name
  validation (:1892-1901), execute the command inline against the just-validated `_focusedIndex` instead of
  `PostMessageW` (:1908) — removing the window where a second debounced refresh migrates focus.

**Regression tests:**

- `ContextMenuDeleteFocusMigrationGuard` — seed A (target) and B (slot A vacates); `DebugFocusItemByDisplayName(A)`;
  externally rename/delete A on disk; post `kFolderViewEnumerateComplete` to force a refresh so `ApplyCurrentSort`
  runs with the fallback at A's old slot; invoke `cmd/pane/moveToRecycleBin`; assert `exists(B)==true` and the
  recycle task targeted nothing or A-by-name only — never B. Drive refresh + command via posted messages in fixed
  order for determinism.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=ContextMenuDeleteFocusMigrationGuard --selftest-timeout-multiplier=4
```

---

## Task 7: Report COPY To External OLE Until Complete; No-Modifier Same-Volume = MOVE

**Root cause:** RC-D. **Severity:** data-loss (premature MOVE echo) + incorrect-behavior (volume-blind effect). **Closes:** D4, B6. **Depends on:** Task 1 (B6 uses the resolved dest), Task 3 harness.

**Files:**

- Modify: `RedSalamander/FolderView.DragDrop.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`

**Repro:** **D4** — a Shift+drag from Explorer (external, `internalDrop=false`) is queued via the callback
(`DragDrop.cpp:575`) which spawns a jthread returning `S_OK` at
`FolderWindow.FileOperations.State.Runtime.Part.cpp:764` before the move runs; `PerformDrop` then writes raw
`DROPEFFECT_MOVE` at :693, `DropTarget::Drop` echoes it to OLE at :151, and Explorer deletes the source before
our async copy succeeds. No drop-effect handshake exists (grep for `PERFORMEDDROPEFFECT`/`PASTESUCCEEDED` = 0
hits). **B6** — a no-modifier same-volume drag yields `DROPEFFECT_COPY` because `ResolveDropEffect`
(`DragDrop.cpp:311-314`) prefers COPY with no volume comparison, violating the Windows same-volume=MOVE
convention.

**Fix summary:** Tell external OLE sources COPY for asynchronously-queued moves (our worker still performs the
real move and deletes only on success), and choose MOVE for no-modifier same-volume drags.

- [ ] **Step 1 (D4): do not echo `DROPEFFECT_MOVE` for an async-queued external move.**
  Add `DWORD reportedEffect = effect;`. Inside the callback block (`DragDrop.cpp:564-577`), before the break
  (:576): `if (effect == DROPEFFECT_MOVE && !internalDrop) { reportedEffect = DROPEFFECT_COPY; }`. Change the
  success write at :693 from `*performedEffect = effect;` to `*performedEffect = reportedEffect;`. The verified
  worker still performs the real `FILESYSTEM_MOVE` and deletes the source on success; `internalDrop==true` and
  the synchronous direct-FS/LINK fallbacks keep `reportedEffect==effect`. Do **not** block the UI thread on the
  worker (that hang is perf-owned).
- [ ] **Step 2 (B6): make the no-modifier default volume-aware in `PerformDrop`.**
  Add a file-local `static bool AreSameVolume(const std::wstring& srcRoot, const std::wstring& destRoot) noexcept`
  using `GetVolumePathNameW` + `CompareStringOrdinal(...TRUE)` (any failure → false → safe COPY). Make `effect`
  a non-const local; when keyState has no Ctrl/Shift/Alt AND `effect==DROPEFFECT_COPY` AND
  `(allowedEffects & DROPEFFECT_MOVE)`, compute the destination root (from Task 1's resolved `dest`); if EVERY
  source shares that volume, set `effect=DROPEFFECT_MOVE` (mixed volumes keep COPY). Leave `ResolveDropEffect`
  and DragEnter/DragOver key-only so visual feedback never blocks on volume enumeration. (D4 still gates the echo.)

**Regression tests:**

- Prereq harness: extend `DebugPerformFileDropForSelfTest` (`DragDrop.cpp:261`) with `bool externalSource`,
  `DWORD allowedEffects` (default COPY|MOVE), `DWORD keyState` (default 0); when `externalSource`, build a
  CF_HDROP-only data object (`suppressInternalFormat`).
- `Commands_FileOpsExternalMoveDropReportsCopyUntilComplete` — external MOVE drop; assert `SUCCEEDED(hr)`,
  `performed==DROPEFFECT_COPY` (evaluated synchronously before the worker finishes), and a host task WAS queued.
  Companion: an INTERNAL move still reports `DROPEFFECT_MOVE`.
- `TestPaneDragDropNoModifierSameVolumeDefaultsToMove` — single-volume source+dest, keyState=0; assert
  `performed==DROPEFFECT_MOVE` and the file moved. Pre-fix returns COPY.
- Unit-test `AreSameVolume` (`C:\` vs `D:\` → false; same root → true) for machine-independent coverage.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=Commands_FileOpsExternalMoveDropReportsCopyUntilComplete --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=TestPaneDragDropNoModifierSameVolumeDefaultsToMove --selftest-timeout-multiplier=4
```

---

## Task 8: Clamp Plugin-Reported Enumeration Count Before `reserve()`

**Root cause:** RC-F. **Severity:** robustness / crash (`std::terminate` DoS). **Closes:** B4.

**Files:**

- Modify: `RedSalamander/FolderView.Enumeration.cpp`
- Modify: `Plugins/FileSystemDummy/FileSystemDummy.cpp` / `.h` (test mode)
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`

**Repro:** A plugin returns a hostile `GetCount` (`Common/PlugInterfaces/FileSystem.h:342`, `unsigned long`)
decoupled from its real buffer; `estimatedFiles=entryCount` (`Enumeration.cpp:391`) drives `files.reserve` (:393)
**outside** the `try` (which opens at :396). `reserve(0xFFFFFFFF*sizeof)` exceeds `vector::max_size` →
`std::length_error` escapes `ExecuteEnumeration` → `EnumerationWorker` (:272) → the non-noexcept jthread lambda
(:223) → `std::terminate`.

**Fix summary:** Treat `GetCount` strictly as a hint: clamp against the actual buffer size and a fixed ceiling
before reserving, and keep the reserve inside the worker `try`.

- [ ] **Step 1: Clamp by ceiling + buffer size before computing reserve estimates.**
  Add `static constexpr size_t kEnumerationReserveCeiling = 1u << 20;`. Hoist `GetBufferSize` above the reserves:
  `unsigned long bufferSizeHint=0; if (FAILED(filesInformation->GetBufferSize(&bufferSizeHint))) bufferSizeHint=0;`
  then `const size_t maxByBuffer = bufferSizeHint ? (static_cast<size_t>(bufferSizeHint)/sizeof(FileInfo)) : static_cast<size_t>(entryCount);
  const size_t clampedCount = std::min<size_t>(static_cast<size_t>(entryCount), std::min(maxByBuffer, kEnumerationReserveCeiling));`
  and base `estimatedDirs/estimatedFiles` (:390-391) on `clampedCount`.
- [ ] **Step 2: Defense-in-depth: move the `try {` above the reserves.**
  Move the `try {` from `Enumeration.cpp:396` up to immediately before :392 so any residual throw degrades to
  `E_FAIL` via the existing `catch(exception)` (:896-900) instead of escaping the jthread lambda. Do not rely on
  the `catch(bad_alloc)->terminate` handler at :892-894 (unreachable for this path and itself a latent DoS).

**Regression tests:**

- Add `SetMismatchedCount(unsigned long)` to `DummyFilesInformation` (override `GetCount` to `0xFFFFFFFF` while
  `GetBufferSize`/`GetAllocatedSize` keep the tiny real buffer). Add `folderView_enumeration_hostile_count_does_not_terminate`
  in `Commands.SelfTest.Navigation.cpp`: navigate to a dummy-FS path enumerating that stub, pump the loop, assert
  the process survives and `ProcessEnumerationResult` completes (payload trace at `Enumeration.cpp:1475/1495`). Pre-fix
  the unclamped reserve terminates the harness.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_enumeration_hostile_count_does_not_terminate --selftest-timeout-multiplier=4
```

---

## Task 9: Add The Missing `< _items.size()` Bounds Clause To DrawItem's Hovered Subscript

**Root cause:** RC-E. **Severity:** robustness / defense-in-depth (latent OOB read). **Closes:** B8.

**Files:**

- Modify: `RedSalamander/FolderView.Rendering.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

**Repro:** Latent today (every `_items`-shrinking path resets `_hoveredIndex` on the UI thread before the next
paint). Inject a stale out-of-range `_hoveredIndex` and force a paint: the render loop (`Rendering.cpp:1245-1255`)
calls `DrawItem`, which subscripts `_items[_hoveredIndex]` (:2176) with only the sentinel guard → OOB read / UB
under ASan or `_ITERATOR_DEBUG_LEVEL=2`.

**Fix summary:** Add the same `< _items.size()` bounds clause the three sibling `_hoveredIndex` sites already use.

- [ ] **Step 1: Add the bounds clause before the subscript.**
  In `DrawItem` (`Rendering.cpp:2176`), change `const bool isHovered = (_hoveredIndex != static_cast<size_t>(-1) && std::addressof(item) == std::addressof(_items[_hoveredIndex]));`
  to `const bool isHovered = (_hoveredIndex != static_cast<size_t>(-1) && _hoveredIndex < _items.size() && std::addressof(item) == std::addressof(_items[_hoveredIndex]));`.
  The `< _items.size()` clause must precede the subscript so `&&` short-circuit prevents the OOB index from
  forming. Matches `Interaction.cpp:55/:556/:566` exactly.

**Regression tests:**

- Add an ENABLE_TESTS `DebugSetHoveredIndexForTest(size_t)` setter (`_hoveredIndex` is private, `FolderView.h:922`).
  Add `folderView_drawitem_stale_hover_index_is_bounded`: populate N items, `DebugSetHoveredIndexForTest(_items.size()+5)`,
  force a paint, assert the process survives and no item reports `isHovered`. Pre-fix the subscript is OOB UB under ASan.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_drawitem_stale_hover_index_is_bounded --selftest-timeout-multiplier=4
```

---

## Task 10: Overflow-Guard `BuildFileDropHGlobal` Allocation Size

**Root cause:** RC-B. **Severity:** low / defense-in-depth (heap OOB write under pathological input). **Closes:** N-build.

**Files:**

- Modify: `RedSalamander/FolderView.FileOps.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`

**Repro:** Inputs are trusted (`GetSelectedOrFocusedPaths`), so reaching 2^64 total bytes needs synthetic path
counts; with a crafted vector of very long paths, `totalChars*sizeof(wchar_t)` or `sizeof(DROPFILES)+` wraps
`SIZE_T`, `GlobalAlloc` gets a too-small `bytes`, `GlobalLock` succeeds, and the copy loop (:125-133) writes past
the allocation (heap OOB write).

**Fix summary:** Reject integer overflow when sizing the CF_HDROP allocation, returning `nullptr` instead of an
undersized buffer, matching the overflow-checked pattern already used by the internal-format writer.

- [ ] **Step 1: Compute the allocation size with overflow-checked add/multiply.**
  In `BuildFileDropHGlobal` (`FileOps.cpp:99-107`), accumulate `totalChars` with an overflow-checked add (reject
  if `totalChars > SIZE_MAX - (path.native().size()+1)`) and compute `bytes` with a checked multiply/add (reject
  if `totalChars > (SIZE_MAX - sizeof(DROPFILES)) / sizeof(wchar_t)`), returning `nullptr` (the existing failure
  contract) before `GlobalAlloc` at :108. Mirrors `CreateInternalFileDrop` (`FolderViewInternal.h:1466-1502`).

**Regression tests:**

- Add a debug hook invoking `BuildFileDropHGlobal` with a crafted vector of very long paths; assert it returns
  `nullptr` instead of allocating; also assert a normal selection still succeeds via the pane Copy command path.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_buildfiledrop_rejects_size_overflow --selftest-timeout-multiplier=4
```

---

## Task 11: Clear/Invalidate The Clipboard After A Completed Move-Paste

**Root cause:** RC-B/RC-D (clipboard/OLE contract — the missing write-side handshake). **Severity:** data-integrity (repeat-paste retries a completed MOVE). **Closes:** NS-1 (harvested from Notes_Scratch NS-1, 2026-07-02; code-verified same day).

**Files:**

- Modify: `RedSalamander/FolderView.FileOps.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`

**Repro (verified 2026-07-02 folder review, post-drift line numbers):** `FolderView::PasteItemsFromClipboard`
(`FolderView.FileOps.cpp:731-820`) reads `CFSTR_PREFERREDDROPEFFECT` (:744) and dispatches the move, but never
empties/invalidates the clipboard nor writes `CFSTR_PERFORMEDDROPEFFECT` (grep = 0 hits repo-wide), so a second
Ctrl+V retries moving already-moved sources — a repeat-paste data-integrity hazard. (Consistent with Task 7's
D4 repro note that no drop-effect handshake exists anywhere in the codebase.)

**Fix summary:** After a MOVE-effect paste completes, write `CFSTR_PERFORMEDDROPEFFECT` / clear the clipboard
per the OLE clipboard contract, so a stale cut cannot be replayed.

- [ ] **Step 1: Complete the OLE move-paste handshake.**
  In `PasteItemsFromClipboard` (`FolderView.FileOps.cpp:731-820`), after a MOVE-effect paste completes
  successfully, write `CFSTR_PERFORMEDDROPEFFECT` (DWORD `DROPEFFECT_MOVE`) back to the clipboard and/or clear
  the clipboard so the cut source list is invalidated, per the OLE clipboard contract. COPY-effect pastes remain
  repeatable (no clearing). If the paste is dispatched asynchronously, invalidate only on verified completion —
  do not clear on queue (mirrors Task 7's "report MOVE only after it happened" rule).

**Regression tests:**

- `TestPaneClipboardSecondPasteAfterMoveDoesNotRetryMove` — cut a file (preferred effect MOVE), paste into
  folder B, wait for completion; paste again; assert no second move/task is dispatched for the already-moved
  source and no error-on-missing-source occurs. Red before the fix (second Ctrl+V retries the move), green
  after — gate on this red-before/green-after selftest like the other tasks.

**Verify:**

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=TestPaneClipboardSecondPasteAfterMoveDoesNotRetryMove --selftest-timeout-multiplier=4
```

---

## Sequencing

Order is by severity/blast-radius; the two single-gesture data-loss bugs ship first.

1. **Tasks 1 & 2** — highest-impact wrong-target/data-loss; self-contained.
2. **Task 3** — security/crash cluster; **adds the shared harness** (`DebugPerformDropFromDataObjectForSelfTest`
   + `FolderViewDataObject(suppressInternalFormat)`) that Tasks 4 and 7 depend on, so land it before them.
3. **Task 1 before Task 7's B6 step** (B6 computes the destination root from Task 1's resolved `dest`).
4. **Task 5's drop gate** is independent of Task 1, but if both land, place the enumerated-folder gate **before**
   Task 1's dest resolution so a pending folder is rejected before hit-testing.
5. **Tasks 6, 8, 9, 10, 11** are independent and can land in any order after the cluster above; 9 and 10 are
   one-liners suitable for single small commits. (Task 11 added 2026-07-02 from Notes_Scratch NS-1; it touches
   `PasteItemsFromClipboard` like Task 5 Step 2 — review together if both are in flight.)

Each task is its own commit (Definition of Done). Do not batch. Tasks 1 and 4 both touch the external-drop
path — review them together to avoid conflicts in `PerformDrop`/`StartFileOperationFromFolderView`.

---

## Implementation Findings Log

Append discoveries as work proceeds (newest at the bottom). A step is not Done until its
red-before/green-after evidence path is recorded here.

### Pre-implementation verification (multi-agent pass, 2026-06-28, HEAD `045a1f773`)

- 16 audit findings re-checked: **15 present**, **1 (D1) mitigated on the production path**. 6 new findings
  (N1, N2, N3, N-build) confirmed distinct from the audit set.
- [Task 1 · D1] D1 is mitigated by the host overlap guard `StartOperation` (`FolderWindow.FileOperations.State.Runtime.Part.cpp:523-666`,
  `IsSameOrChildPath` `FolderWindow.FileOperations.State.cpp:2029`), reached by every drop MOVE. The FS chain is
  unsafe in isolation but unreachable in production; Task 1 covers the ENABLE_TESTS direct fallback. Evidence: as cited.
- [Task 3 · N1] `ReadFileDropClipboard` (`FolderView.FileOps.cpp:220-221`) reserves from an attacker CF_HDROP count
  inside a `noexcept` free fn — the ordinary Ctrl+V path, distinct from S2/B4. Evidence: FolderView.FileOps.cpp:205-235.
- [Task 5 · N2] Cross-pane Copy/Move (F5/F6) resolves the destination from the destination pane's in-flight
  `_currentFolder` with no `IsCurrentFolderEnumerated()` gate. Evidence: FolderWindow.FileOperations.cpp:1238,1279.
- [Task 2 · N3] Context-menu Delete/Move/Copy on a right-clicked UNSELECTED item acts on the stale prior
  multi-selection (`OnContextMenu` only `FocusItem`s). Evidence: FolderView.Menus.cpp:354.
- [Task 10 · N-build] `BuildFileDropHGlobal` (`FolderView.FileOps.cpp:99-107`) sizes its `GlobalAlloc` with
  unchecked integer math (write-side overflow). Evidence: FolderView.FileOps.cpp:99-133.
- [Out of scope · R1] Worker torn-read of `_itemsFolder` (`FolderView.Icons.cpp:539,911`) confirmed present; fix
  owned by WarpDrive G3.

### During implementation (append below)

<!-- - [DATE · Task N] <red-before/green-after, archive path> (commit <hash>) -->

- [2026-07-02 folder review · Task 3 · S3] Step 3's `GlobalSize` guard was landed incidentally by the WarpDrive
  B10 commit `b9af84d72` (2026-06-28), not by this plan. Verified by an independent code-verification pass
  (2026-07-02) at `FolderView.FileOps.cpp:268-271`: `if (GlobalSize(handle) < sizeof(DWORD)) return std::nullopt;`.
  Step 3 box checked as externally completed; no red-before evidence is reproducible now, so the
  `TestPaneClipboardPasteIgnoresUndersizedPreferredDropEffect` selftest remains owed as a green-only pin under Task 3.
- [2026-07-02 folder review · Task 11 · NS-1] New task harvested from Notes_Scratch NS-1 and code-verified same
  day: `PasteItemsFromClipboard` (`FolderView.FileOps.cpp:731-820`) reads `CFSTR_PREFERREDDROPEFFECT` (:744) but
  never writes `CFSTR_PERFORMEDDROPEFFECT` / clears the clipboard after a completed MOVE paste (grep = 0 hits
  repo-wide), so a second Ctrl+V retries moving already-moved sources. Tracked as Task 11.

---

## Acceptance Checklist

- [ ] Drop honors the cursor (lands in the hovered subfolder) and rejects self/same-parent/descendant.
- [ ] Click / right-click / empty-background gestures establish the selection they imply before any drag or command.
- [ ] Malformed CF_HDROP, hostile internal `pathCount`, undersized preferred-drop-effect, and huge Ctrl+V counts all fail safe (no OOB, no `std::terminate`).
- [ ] External local-file imports onto a virtual-FS pane read via the local provider or reject with a compatible-FS error.
- [ ] Drop / paste / cross-pane copy/move into an un-enumerated folder are rejected, not re-targeted.
- [ ] A nested-loop or debounced refresh can no longer migrate a destructive command onto a neighbor.
- [ ] External async-queued moves report COPY to the source; no-modifier same-volume drags select MOVE.
- [ ] A hostile plugin `GetCount` degrades to a failed enumeration, not process termination.
- [ ] `DrawItem` cannot subscript `_items` out of bounds on a stale hover index.
- [ ] `BuildFileDropHGlobal` rejects size overflow.
- [ ] A second Ctrl+V after a completed move-paste cannot retry the move (`CFSTR_PERFORMEDDROPEFFECT` written / clipboard cleared).
- [ ] Every fix has a red-before / green-after selftest; `.\Tools\Run-AllTests.ps1 -Suite Full` passes.
- [ ] Durable behavior contracts merged into `Specs/UI/UI_FolderView.md`; plan moved to `Specs/Plans/Done/`.
