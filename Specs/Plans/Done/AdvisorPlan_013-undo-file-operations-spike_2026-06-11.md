# Advisor Plan 013 - Design spike for file-operations undo

> **NON-NORMATIVE RETIRED RECORD.** Archived from root `plans/013-undo-file-operations-spike.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** RETIRED
- **Archived:** 2026-08-25
- **Original path:** `plans/013-undo-file-operations-spike.md`
- **Live owner:** Specs/Plans/WIP/Product_WhimFilesRemainingDecisions_2026-08-25.md (D1 G2 routing); mutation identity remains in Specs/FileSystem/FileSystem_FileOperations.md. I11 owns the File Operations contract. Not an implementation owner.
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 013: Design spike — undo for file operations (building on the batch-rename undo plan)

> **Frozen historical instructions (do not execute)**: This is a DESIGN SPIKE. The deliverable is a design
> document in the repo's own planning location (`Specs\Plans\WIP\`), grounded in
> the existing code — NOT an implementation. Do not write production code. Honor
> STOP conditions. Historical: status-row updates no longer apply in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat d6bfccc42..HEAD -- RedSalamander/BatchRenameWindow.cpp RedSalamander/BatchRenameEngine.h RedSalamander/FolderWindow.FileOperations.State.cpp Specs/Plans`
> On drift, re-verify anchors before proceeding.

## Status

- **Priority**: P3
- **Effort**: M (spike only; the eventual feature is L-XL)
- **Risk**: LOW (doc-only)
- **Depends on**: none
- **Category**: direction
- **Planned at**: commit `d6bfccc42`, 2026-06-11

## Why this matters

RedSalamander mutates user files — and offers no undo. The only artifact in this direction is the batch-rename "Copy Undo Plan" (an exported tab-separated report of actually-renamed rows; see commit `84bdeb5b1` "Add batch rename undo plan report"), which is a manual escape hatch, not an undo. Competing dual-pane managers don't generally solve this either, so even a session-scoped undo for rename/move (the cheap-to-reverse operations) would be a differentiator; for delete, Recycle-Bin-mediated restore is plausible for local files. The just-built batch-rename machinery (per-row outcome tracking, collision-safe ordering, execution reports) is the natural seed. What's needed first is a scoped design with eyes-open trade-offs — full transactional undo across remote filesystems is a tar pit, and the design must explicitly fence it off.

## Current state (anchors to investigate from)

- Batch rename undo artifacts (recent work, commits `84bdeb5b1`, `a48d34909` "Add batch rename execution report copy", `5682ba807` "Cover case-only batch rename execution"): find the undo-plan generation/report in `RedSalamander\BatchRenameWindow.cpp` (`Grep -in "undo" RedSalamander\BatchRename*.cpp RedSalamander\FileSystemRenameBatch.cpp`) — document its exact shape: what rows it records (old path → new path), when it's offered, what guarantees it has (post-execution actual outcomes vs. planned).
- File-operations engine state machine: `RedSalamander\FolderWindow.FileOperations.State.cpp` (+ `.Queue.Part.cpp`, `.Runtime.Part.cpp`, `.Diagnostics.Part.cpp`) — operations, per-item outcome tracking, error/conflict handling, cancellation. The design must answer: does the engine already record, per item, what actually happened (needed for any undo log)?
- Plugin capability surface: `Common\PlugInterfaces\FileSystem.h` — operations are capability-gated per filesystem (`GetCapabilities`: "operations", "crossFileSystem"); undo on a remote FS means issuing reverse operations through the same plugin. Delete has `deleteRecycleBin*` capability hints (see the doc comment around :478-490) — relevant for restore-from-recycle-bin feasibility.
- Selftest vehicle for the eventual feature: `RedSalamander\SelfTest\FileOperations\` (50+ phases) — the design should name which phases get siblings.
- UI command surface: `RedSalamander\CommandRegistry.cpp`, `RedSalamander\ShortcutDefaults.cpp` (where an Undo command/shortcut would register), `Specs\UI\UI_CommandMenuKeyboard.md`.
- Repo planning convention (AGENTS.md): design docs live in `Specs\Plans\WIP\`, move to `Done\` when implemented, durable contracts merge into `Specs\<Domain>\`.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build (only to validate nothing changed) | `.\build.ps1` | exit 0 |
| Searches | Grep/Glob as above | anchors found |

## Scope

**In scope**:
- Reading; the design doc `Specs\Plans\WIP\Core_FileOperationsUndo.md`; a row in `plans/README.md`.

**Out of scope**:
- ANY production or test code.
- Redesigning the file-operations engine — the design must work WITH the existing state machine or explicitly call out the minimal extension it needs.
- Archive (7z) and cloud-drive undo in v1 scope — the doc may mention them only as explicit non-goals.

## Git workflow

- Branch: `advisor/013-undo-spike`
- Single commit: the design doc.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Inventory what exists

Document with file:line: (a) batch rename's undo-plan content and lifecycle; (b) what the fileops engine records per completed item (paths, outcome, timestamps?); (c) which operations are theoretically reversible per type — rename↔rename, move↔move (same-volume vs cross-volume!), copy↔delete-the-copy, delete↔recycle-bin-restore (local only; check what the engine currently does for delete: recycle bin or permanent? find it in the FileOperations code, the capability hints suggest both exist); (d) where operation completion is signaled (the hook point for writing an undo journal).

**Verify**: every item has a file:line citation.

### Step 2: Write the design doc

`Specs\Plans\WIP\Core_FileOperationsUndo.md` containing:

1. **Goal and non-goals.** v1 candidate: session-scoped (in-memory, lost on exit), single-level-or-small-stack undo for LOCAL filesystem rename, same-volume move, and batch rename; copy-undo (delete the copies) optional; delete-undo only via recycle-bin restore if and only if Step 1 shows deletes go to the recycle bin. Non-goals: persistence across sessions, remote/cloud/archive filesystems, redo, mid-operation undo.
2. **Undo journal design.** What gets recorded (per-item actual outcomes — reuse/extend the batch-rename undo-plan row shape), where it hooks into the state machine (the Step 1(d) point), staleness checks before undo (target unchanged since the operation? mtime/size compare; on mismatch → per-item skip+report, mirroring batch rename's fallback-reporting style from commit `d6bfccc42`).
3. **UX.** Where Undo surfaces (menu/shortcut via CommandRegistry; the post-operation report popup), what the confirmation shows, how partial undo failures report (reuse the execution-summary pattern from `e599ee80a` "Report batch rename execution summaries").
4. **Risk register.** Cross-volume moves (copy+delete — not reversible as "move back" without a copy), case-only renames, collisions at undo time, watcher-triggered refresh races, multi-window/same-folder conflicts, undo of an operation whose source folder is now gone.
5. **Test strategy.** Named new selftest phases under `RedSalamander\SelfTest\FileOperations\` (e.g. UndoRenameHappyPath, UndoMoveStaleTargetSkips, UndoAfterPartialFailure) and which existing phases they pattern on.
6. **Phasing + coarse effort.** Phase 1 rename/move undo; Phase 2 batch-rename integration (turn the export-only plan into an executable undo); Phase 3 delete/copy. Per-phase effort (S/M/L) and the open questions only the maintainer can answer (e.g. is persistence wanted? is a 1-deep undo acceptable for v1?).

**Verify**: doc ≤ ~4 pages; every design claim that references current behavior carries a file:line.

### Step 3: Clean finish

`git status` shows only the doc; `.\build.ps1` still exits 0 (nothing touched).

## Test plan

None (spike). Section 5 of the doc IS the future test plan.

## Done criteria

- (archived, not live)  `Specs\Plans\WIP\Core_FileOperationsUndo.md` exists with all six sections, citations included
- (archived, not live)  No code changes (`git diff --stat -- RedSalamander/ Common/ Plugins/` empty)
- (archived, not live)  `plans/README.md` status row updated with a one-line summary of the recommended v1 scope

## STOP conditions

- The fileops engine does NOT record per-item actual outcomes anywhere (Step 1(b) comes up empty) — the design's foundation is missing; write that up as the headline finding (the doc becomes "prerequisite: outcome journal") and finish.
- You find an existing undo design/spec in `Specs\` (search first: `Grep -rin "undo" Specs\ --glob "*.md"`) — reconcile with it instead of writing a competing one.

## Maintenance notes

- If the maintainer green-lights a phase, implementation becomes a new plan (or the repo's own WIP-plan workflow takes over — follow AGENTS.md spec-closeout from there).
- The other three direction findings from the audit (archive write support, in-app update check, plugin developer guide) were recorded in plans/README.md as unplanned candidates — this spike was selected first.
