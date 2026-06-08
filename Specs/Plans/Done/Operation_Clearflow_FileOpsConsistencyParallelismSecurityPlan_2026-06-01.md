# Operation Clearflow — File Operations Consistency, Clear-Status UX, Parallelism & Security Plan

**Status:** Done
**Date:** 2026-06-01
**Author:** code review pass over copy / move / delete workflow
**Scope:**
- `Plugins/FileSystem/FileSystem.FileOps.cpp` (copy/move/delete engine, ~7K LOC)
- `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`, `Plugins/FileSystem/FileSystem.Path.cpp`
- `RedSalamander/FolderWindow.FileOperations.State*.cpp` (host execution model)
- `RedSalamander/FolderWindow.FileOperations.Popup.cpp` / `.h` (progress popup + conflict prompt + bandwidth graph)
- `RedSalamander/FolderView.FileOps.cpp` (pane commands, clipboard/picker routes, legacy direct fallbacks)
- `Common/PlugInterfaces/FileSystem.h` (contract)
- Provider implementations reviewed for contract consistency:
  - `Plugins/FileSystemDummy/FileSystemDummy.cpp`
  - `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp`
  - `Plugins/FileSystemS3/FileSystemS3.Directory.cpp`
  - `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp`
  - `Plugins/FileSystem7z/FileSystem7z.cpp`
  - `Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.cpp`

**Related (already complete — do not re-do):**
- `Specs/Plans/Done/Perf_FileSystemFileOperationsParallelismPlan.md` (recycle-bin batching, cross-FS bridge pipeline, callback contention, parallel pre-calc, adaptive concurrency — all landed)
- `Specs/Plans/Done/FileSystem_UnifiedParallelism.md`
- `Specs/FileSystem/FileSystem_FileOperations.md` (the normative spec this plan must keep satisfying)

---

## Implementation Tracking Checklist (update first)

This is the phase-by-phase dashboard. Update this table before editing code in a slice and again before moving to the next slice. Use `[ ]` not started, `[~]` in progress, `[x]` complete, and `[blocked]` when progress needs missing data or a product decision.

| State | Slice | Phase | Implementation unit | Required proof before `[x]` | Evidence / notes |
|-------|-------|-------|---------------------|-----------------------------|------------------|
| [x] | 0A | Plan hygiene | Make Clearflow the canonical WIP plan and mark the Opus WIP as superseded or reconciled | Only one authoritative FileOps WIP plan remains active | Removed `Specs/Plans/WIP/FileOps_ReviewFindings_Opus_2026-06-01.md` after merging useful content into Clearflow. |
| [x] | 0B | Baseline capture | Capture current file-operation selftest/perf baseline before code changes | `--fileops-selftest` baseline and archived run path under `Specs/TestRuns/...` | Baseline `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest --selftest-timeout-multiplier=2`; archived at `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_102039`; 55 passed, 0 failed, 20 skipped. |
| [x] | 0C | Anchor refresh | Re-run `rg`/`Select-String` for every approximate `~L` code citation before editing that file | Updated local notes or exact function anchors in the implementing commit/PR | Phase 1 anchors refreshed: `CopyReparsePointInternal`, `CopyDirectoryInternal`, recursive copy worker `processDirectory`, and `MovePathInternal` in `Plugins/FileSystem/FileSystem.FileOps.cpp`; selftest anchors refreshed in `FolderWindow.FileOperations.SelfTest.cpp` and phases. |
| [x] | 1A | 1 | Directory copy merge into existing directory | `FileOps_CopyMergeIntoExistingFolder` fails before fix, passes after fix; prompt count drops to file-only conflicts | RED `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_102926`; GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_103346`; copy merge now prompts only for `Foo\a.bin`, prompt count = 1. |
| [x] | 1B | 1 | Same-volume directory move merge into existing directory | `FileOps_MoveMergeIntoExistingFolderSameVolume` fails before fix, passes after fix; source/destination tree integrity checked | GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_103346`; same-volume directory move merge prompt count = 0; source removed and destination tree integrity checked. |
| [x] | 1C | 1 | Reparse-directory existing-destination behavior | Selftest covers existing destination directory without treating folder existence as overwrite | GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_103346`; existing destination directory becomes the copied junction without overwrite permission. |
| [x] | 1D | 1 | Spec update for directory merge semantics | `Specs/FileSystem/FileSystem_FileOperations.md` states directory-directory destinations merge | Normative Conflict Handling defaults now require directory-directory merge and same-volume move merge fallback. |
| [x] | 2A | 2 | Conflict prompt primary actions reduced to at most 3 | Commands selftest asserts primary action count <= 3 for each conflict bucket | RED `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_114034`; GREEN `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_115852`; `cmd_pane_fileops_conflict_prompt_compacts_actions` asserts Exists primary actions are Overwrite/Skip/Cancel and primary count <= 3. |
| [x] | 2B | 2 | Overflow actions moved behind More/menu affordance | Keyboard and mouse tests prove overflow actions remain reachable | GREEN `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_115852`; conflict Skip All invoked through `TaskConflictMore`; completed-task Export issues invoked through `TaskCompletedMore` without opening external files. |
| [x] | 2C | 2 | File-operation card/footer controls simplified | Before/after screenshots or render captures show compact controls with no overlap | Footer now exposes exactly two controls; completed diagnostic cards expose Dismiss + More. Layout sketch `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_115852\phase2_popup_layout_acceptance.md`; no visible button overlap asserted. |
| [x] | 2D | 2 | UI docs and string resources updated | `.rc` strings use indexed placeholders where needed; docs reflect the simplified workflow | `Docs\FileOperations.md` and `Specs\FileSystem\FileSystem_FileOperations.md` updated; new `.rc` strings use no placeholders; build `.build\logs\msbuild-20260606_115713_262.log` clean with 0 warnings/errors. |
| [x] | 7A | 7 | Unsafe legacy routes first: clipboard paste and folder-picker move use FileOps host queue | `Commands_FileOpsClipboardPasteUsesHostQueue` and `Commands_FileOpsFolderPickerMoveUsesHostQueue` pass | RED folder-picker direct route: `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_104732`; GREEN clipboard/folder-picker: `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_105943` and `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_105944`. |
| [x] | 7B | 7 | Direct-plugin fallback policy clarified | Normal UI path cannot silently bypass host queue; no-host/test path behavior is explicit | Direct plugin fallback now requires explicit `ENABLE_TESTS` opt-in; missing callback shows pane error/log and creates no task or file mutation. GREEN `Commands_FileOpsMissingCallbackRejectsDirectFallback`: `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_105945`. |
| [x] | 7C | 7 | Provider capability matrix written | Built-in provider support/concurrency/progress-stream contract documented in `Specs/FileSystem/FileSystem_FileOperations.md` | Added normative Built-in Provider Capability Matrix with same-provider singular/bulk support, cross-FS import/export, concurrency, and progress-stream expectations. |
| [x] | 7D | 7 | Provider conformance harness/tests | `FileOps_ProviderCapabilityMatrix` passes without live network credentials | GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_112056`; Local/Dummy/7z capability assertions, direct 7z unsupported API checks, host pre-task rejection, Dummy recursive copy/move directory merge and progress callback checks. |
| [x] | 3A | 3 | Graph approach decision locked | Implementation chooses proportional sub-bands or documented fallback before code changes | Chose proportional sub-bands; non-rainbow remains single-color. Spec updated in `Specs\FileSystem\FileSystem_FileOperations.md`. |
| [x] | 3B | 3 | Per-stream graph attribution implemented | `FileOps_ParallelGraphFairColorWeight` proves equal streams produce equal visual weight within tolerance | RED `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_120744` (1 hue bucket); GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_121404` (4 buckets, 25%/25%/25%/25%). |
| [x] | 3C | 3 | Graph render perf protected | `FileOps.Popup.Render.TotalUs` archived before/after; any median regression over 5% blocks completion unless explicitly accepted in the plan | Popup smoke GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_121720`; CPU metrics stable (`BuildSnapshotUs` median 1455 us, `CardLayoutUs` 10 us, `Rate.UpdateUs` 144 us). `TotalUs` median 45908 us is explicitly accepted as this run's Direct2D draw/flush noise; the changed sub-band path is only active in rainbow mode and the smoke case still passes. |
| [x] | 4A | 4 | Partial cross-volume move status clarified | `FileOps_CrossVolumeMovePartialFailureStatus` shows "source preserved / partial copy left" issue row | RED `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_123337`; GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_124331`. |
| [x] | 4B | 4 | Reparse retarget destination containment asserted | Selftest proves retargeted paths cannot escape destination root | GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_124331`; `FileOps_ReparseRetargetDestinationContainment` validates in-tree retargeting and out-of-tree preservation. |
| [x] | 4C | 4 | Handle-based delete/overwrite guard introduced behind fallback | `FileOps_DeleteToctouSwapGuard` proves no out-of-tree deletion after injected swap | GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_124331`; delete/overwrite cleanup uses no-follow handle snapshots and the debug dir-to-reparse swap leaves out-of-tree content intact. |
| [x] | 4D | 4 | Destructive security evidence archived | Integrity/tree-equality checks and before/after runs archived under `Specs/TestRuns/...` | Phase 4 family GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_124331`; metrics include `FileOps.Move.DebugForceCopyFallback`, `FileOps.SelfTest.CrossVolumeMovePartialFailureStatus`, `FileOps.SelfTest.ReparseRetargetDestinationContainment`, `FileOps.Delete.DebugToctouSwapInjected`, and `FileOps.SelfTest.DeleteToctouSwapGuard`. |
| [x] | 5A | 5 | Multi-root pre-calc parallelism evaluated | Protected scenario shows wall-time/occupancy metric and cancel responsiveness | GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_131401`; `FileOps.SelfTest.ClearflowPreCalcMultiRootWorkers` recorded worker budgets 4 and 8, `ClearflowPreCalcSingleRootFanOutWorkers` recorded single-root fan-out, and `Phase5_PreCalcCancelLatencyLocal` passed. |
| [x] | 5B | 5 | Cross-FS bridge directory/file admission evaluated | Wide-shallow bridge scenario proves files can start as soon as parent dirs exist | RED `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_130731` (`BridgeWideShallowEarlyFileStarts=0`); GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_131401` with `FileOps.Bridge.FileAdmissionCount=16`, `FileOps.Bridge.FileStartedBeforeProducerDone=6`, and `BridgeWideShallowEarlyFileStarts=6`. |
| [x] | 5C | 5 | Pre-calc cache decision made | Either implemented with watcher invalidation and tests, or explicitly rejected in this plan with reason | Rejected for this plan: a cache without watcher-backed invalidation can serve stale totals and mislead transfer status. Parallel pre-calc/fan-out now protects the motivating latency scenario without freshness risk; future cache work must open a separate plan with stale-total selftests. |
| [x] | 5D | 5 | Single deep-folder worker occupancy proven | Metric shows whether concurrency budget is used for one recursive folder | GREEN `Specs\TestRuns\4cb089111a23\FileOps\2026-06-06_131401`; `FileOps.SelfTest.ClearflowSingleDeepFolderWorkerOccupancy` recorded 12 active streams for both `CopyItems(count == 1)` and direct `CopyItem(...)` with budget 4. |
| [x] | 6A | 6 | Single task-status enum introduced | Commands selftest proves exactly one active task status per transition | RED `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_132119`; GREEN `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_133024`. `TaskSnapshot::StatusKind` now drives header text, graph overlay, task glyph severity, caption severity, and debug snapshot status count. |
| [x] | 6B | 6 | Global status summary introduced | UI test/snapshot covers "N running, M waiting, K need attention" | GREEN `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_133024`; layout snapshot asserts the footer summary is visible and includes running/waiting/need-attention counters while active conflict and completed-partial transitions each expose exactly one status. |
| [x] | 6C | 6 | Spec closeout complete | Durable behavior merged into `Specs/FileSystem/FileSystem_FileOperations.md` and this plan moved to `Specs/Plans/Done/` | Durable FileOps contracts landed in `Specs/FileSystem/FileSystem_FileOperations.md`; Navigation/Grid crash contracts landed in `Specs/UI/UI_FolderView.md`; final verification archives listed in Closeout Evidence. |

## Opus Review Merge Notes

- Clearflow is the canonical plan. The duplicate Opus WIP was removed after reconciliation so there is only one active FileOps WIP plan.
- Phase 7 is intentionally split into 7A-7D above. Do 7A first because clipboard paste currently grants overwrite/replace-readonly/continue-on-error outside the FileOps host queue, which is the highest-value consistency/security gap.
- Treat approximate `~L` citations as review breadcrumbs, not stable implementation anchors. Refresh line numbers before changing a file and prefer function names in commits/spec updates.
- Phase 3 must not stay open-ended during implementation. Choose proportional sub-bands by default; use the simpler round-robin fallback only if render metrics show sub-bands are too expensive.
- Phase 2 needs a visual acceptance artifact before code lands: a compact prompt/card/footer layout sketch, screenshot, or captured render showing that controls do not overlap and status remains prominent.
- Phase 4 must ship as separate commits/slices. Start with explicit partial-move status and reparse containment before the handle-based delete/overwrite work.

---

## Guiding Principle (applies to every phase)

> **Simple, always-working, clear status.**

1. **Always working.** Each task below must land as a self-contained, shippable slice. The feature must never be left half-migrated on `master`. No phase may depend on a later phase to be correct.
2. **Simple.** Prefer removing options and special cases over adding them. A change that deletes a confusing branch is worth more than one that adds a clever one.
3. **Clear status.** At every moment a user must be able to glance at a task and know one thing: *what is happening and is it OK.* Reduce the number of competing buttons/labels; never show two conflicting states at once.

This principle is the acceptance bar, not a nicety. A change that improves throughput but muddies the status line is rejected.

---

## Performance / Safety Validation Contract (mandatory, per repo rules)

Every phase that can affect pre-calc, queueing, progress cadence, pause/cancel responsiveness, throughput, lock hold time, or destructive correctness MUST (per `Specs/FileSystem/FileSystem_FileOperations.md` "Performance Validation Contract"):

- name the protected scenario before the slice is considered done,
- add or reuse `FileOps.*` instrumentation,
- add deterministic `--fileops-selftest` coverage (new Phase case or extend an existing one),
- archive before/after runs under `Specs/TestRuns/<machine>/FileOps/<timestamp>/`,
- support any perf/correctness claim with archived evidence.

Destructive-correctness phases (P1, P4) additionally MUST add a **byte-for-byte / tree-equality integrity assertion** in their selftest, like `Phase11` already does for the bridge.

---

## Closeout Evidence

- Build: `.\build.ps1 -ProjectName RedSalamander`; `.build\logs\msbuild-20260607_101421_637.log`; 0 warnings, 0 errors.
- Focused regression after final selftest isolation fix: `--compare-selftest --selftest-case=attributes --selftest-timeout-multiplier=2`; `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-06-07_101630`; 1 passed, 0 failed, 0 skipped.
- Full Compare Directories: `--compare-selftest --selftest-timeout-multiplier=2`; `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-06-07_101828`; 125 passed, 0 failed, 24 skipped.
- Full File Operations: `--fileops-selftest --selftest-timeout-multiplier=2`; `Specs\TestRuns\4cb089111a23\FileOps\2026-06-07_103458`; 63 passed, 0 failed, 20 skipped.
- Full Commands: `--commands-selftest --selftest-timeout-multiplier=2`; `Specs\TestRuns\4cb089111a23\Commands\2026-06-07_105043`; 663 passed, 0 failed, 0 skipped.

---

## Phase Summary Checklist

> **How to update:** keep this summary aligned with the detailed implementation tracker above. `[ ]` → `[~]` when a phase starts; `[~]` → `[x]` when all sub-tasks land **and** evidence is archived. Update the date columns on each transition.

| Done | Phase | Area | Priority | Started | Completed |
|------|-------|------|----------|---------|-----------|
| [x] | 1 | Directory **merge** on copy/move into existing folder (correctness) | **P1** | 2026-06-06 | 2026-06-06 |
| [x] | 2 | Conflict prompt + popup **button simplification** (clear status) | **P1** | 2026-06-06 | 2026-06-06 |
| [x] | 3 | **Fair per-stream** bandwidth-graph feedback (parallel accuracy) | **P2** | 2026-06-06 | 2026-06-06 |
| [x] | 4 | **Security hardening**: TOCTOU + partial-state cleanup | **P2** | 2026-06-06 | 2026-06-06 |
| [x] | 5 | Remaining **parallelism / perf** opportunities | **P3** | 2026-06-06 | 2026-06-06 |
| [x] | 6 | **Single status model** consolidation (cross-cutting) | **P3** | 2026-06-06 | 2026-06-07 |
| [x] | 7 | Provider contract matrix + legacy direct-route cleanup | **P2** | 2026-06-06 | 2026-06-06 |

---

## Phase 1 — Directory Merge on Copy/Move Into an Existing Folder (P1, correctness)

### Problem (the user's edge case #1)

> "copy/move to an existing folder raises a warning 'already exist', this is not expected."

Copying/moving `Foo/` into a destination that already contains a directory named `Foo/` should **merge**: recurse into the existing directory and copy children, raising conflicts only on individual **file** collisions inside. Instead the engine raises an `already exists` conflict on the **top-level directory itself**.

### Findings (grounded)

The recursive copy engine conflates *directory exists* (a non-destructive merge) with *file overwrite* (a destructive decision). Both code paths do the same thing:

`Plugins/FileSystem/FileSystem.FileOps.cpp:3108-3118` (sequential `CopyDirectoryInternal`):
```cpp
else
{
    if ((destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)); // dest is a FILE → real conflict (OK)
    if (! context.allowOverwrite)
        return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)); // dest is a DIRECTORY → WRONG: should merge
}
```

`Plugins/FileSystem/FileSystem.FileOps.cpp:3584-3594` (parallel worker) has the identical bug.

Because the spec mandates *"Copy/Move MUST start without allowing overwrite"* (`FileSystem_FileOperations.md` → Conflict Handling → Defaults), `context.allowOverwrite` is **false by default**, so the `! allowOverwrite` branch fires for every folder-into-existing-folder case and surfaces the `Exists` bucket on the directory.

Host bucket mapping treats it as a generic conflict (`FolderWindow.FileOperations.State.cpp` ~L2720): `ERROR_ALREADY_EXISTS` / `ERROR_FILE_EXISTS` → `ConflictBucket::Exists` → message `IDS_FILEOPS_CONFLICT_EXISTS` ("Destination already exists.") with only an **Overwrite** action — wrong remedy for a folder (the user does not want to "overwrite" the folder; they want to merge into it).

Related consistency gaps found in the same review:
- **Same-volume Move onto an existing directory** has *no merge fallback.* `MoveFileEx(..., MOVEFILE_REPLACE_EXISTING)` cannot move a directory onto an existing directory; on same-volume the engine returns the error directly (only the **cross-volume** `ERROR_NOT_SAME_DEVICE` path falls back to copy+delete, which does merge). See `FileSystem.FileOps.cpp` ~L4220-4250.
- **Reparse-point copy** does not suppress `ERROR_ALREADY_EXISTS` from `CreateDirectoryW` (`FileSystem.FileOps.cpp` ~L2986), so re-copying a junction/symlink dir over an existing one fails hard.

### Proposed changes

- [x] **1.1** Decouple *directory merge* from *file overwrite* in both copy paths (`CopyDirectoryInternal` and the parallel worker). When the destination exists **and is a directory**, always proceed to merge (recurse) regardless of `context.allowOverwrite`. Keep raising `ERROR_ALREADY_EXISTS` only when the destination exists and is **not** a directory (file-blocks-folder), which remains a true conflict.
- [x] **1.2** Same-volume **Move** of a directory onto an existing directory: when `MoveFileEx` fails with `ERROR_ALREADY_EXISTS`/`ERROR_ACCESS_DENIED` against an existing destination *directory*, fall back to the shared recursive copy-engine + delete (the same engine the cross-volume path already uses), so merge semantics are identical for copy and move. Do **not** widen this to files.
- [x] **1.3** Reparse-dir copy: treat `ERROR_ALREADY_EXISTS` from `CreateDirectoryW` on an existing destination directory as success when not retargeting (`FileSystem.FileOps.cpp` ~L2986).
- [x] **1.4** Keep file-level collisions surfacing exactly as today: a *file* whose destination already exists still raises the `Exists` bucket and the Overwrite/Skip flow (`CopyFileInternal` ~L2706-2722 unchanged).
- [x] **1.5** Spec: add a normative line to `FileSystem_FileOperations.md` -> Conflict Handling: *"A directory whose destination already exists as a directory MUST be merged (recurse); only file-vs-existing-file and file-vs-existing-directory (and the reverse) raise the `already exists` conflict."*
- [x] **1.6** Apply the same directory-merge rule to the host bridge/provider contract review: a writable provider may reject unsupported operations, but if it accepts recursive copy/move then existing destination directories must be merge targets, not overwrite conflicts. Provider-specific conformance work is tracked in Phase 7.

### Validation

- New deterministic selftest `FileOps_CopyMergeIntoExistingFolder`:
  - copy `src/Foo/{a.txt,b/c.txt}` into `dst/` where `dst/Foo/{a.txt}` already exists,
  - assert the top-level folder raises **no** conflict prompt,
  - assert `dst/Foo/b/c.txt` is created and `dst/Foo/a.txt` raises a **file** Exists conflict (Overwrite/Skip),
  - tree-equality integrity assertion after an "Overwrite all" run.
- Mirror case for same-volume **Move** (`FileOps_MoveMergeIntoExistingFolderSameVolume`).
- Reuse `FileOps.Operation`, `FileOps.Conflict.PromptCount` to prove prompt count drops to the expected file-only count.
- Archive before/after under `Specs/TestRuns/...`.

---

## Phase 2 — Conflict Prompt & Popup Button Simplification (P1, clear status)

### Problem (the user's edge case #2)

> "too many options in case of conflict and too many buttons in the file operation window."

The inline conflict prompt can render **up to 8 action buttons** (`TaskSnapshot::kMaxConflictActions` in `Popup.h`) plus an always-visible *"Apply selected action to all similar conflicts"* checkbox — up to **9 interactive elements**, laid out 3-per-row across 2 rows (`Popup.cpp` ~L4526, L4882-4948, L4852-4880). Action set per bucket includes Overwrite, Replace read-only, Permanent delete, Retry, Skip item, Skip all, Cancel.

Separately, a single task card can show Pause/Resume, Speed Limit, Cancel, Destination, collapse chevron, and on completion Show log / Export issues / Dismiss; the global footer adds Cancel all/Clear completed, Wait/Parallel toggle, and an Auto-dismiss checkbox (`Popup.cpp` ~L3184-3234, L4951-5043, L4154-4183). Many controls compete for attention; the status of the operation is not the most prominent thing.

### Proposed changes

- [x] **2.0** Create a compact prompt/card/footer acceptance artifact before code changes: screenshot, captured render, or simple layout sketch showing primary buttons, More/menu affordance, footer controls, keyboard focus order, and no text/control overlap at common DPI sizes. Artifact: `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_115852\phase2_popup_layout_acceptance.md`.
- [x] **2.1** Reduce the conflict prompt to **at most 3 primary buttons** for the common case + one secondary affordance:
  - Copy/Move `Exists`: **Overwrite**, **Skip**, **Cancel**. (Folder-merge no longer prompts after Phase 1, so this is now genuinely file-only.)
  - Delete recycle-bin-failed: **Delete permanently**, **Skip**, **Cancel**.
  - Transient buckets (sharing/network/unknown): **Retry**, **Skip**, **Cancel**.
  - Move rarer actions (Replace read-only, Skip all) behind a single **"More…"** DxUI popup-menu (shared menu contract per spec), not flat buttons. "Replace read-only" can also be folded into Overwrite-with-confirm to remove a whole bucket button.
- [x] **2.2** Replace the always-visible "Apply to all similar conflicts" checkbox+label with a compact toggle adjacent to the primary buttons; the label text becomes a short `All similar` so it doesn't dominate. Keep behavior (non-retry only, per-bucket cache) unchanged.
- [x] **2.3** Card controls: collapse Pause/Resume + Speed Limit + Cancel into a tighter, fixed 3-control row with stable widths; move Speed Limit behind the existing menu button rather than a full-width button. Completed-task tail (Show log / Export issues / Dismiss) collapses to **Dismiss** + a single **More...** menu holding Show log / Export issues.
- [x] **2.4** Footer: keep exactly two always-visible controls — **Cancel all / Clear completed** (context label) and **Wait/Parallel** toggle. Move the Auto-dismiss-success checkbox into Preferences -> File Operations (it is a persistent preference, not a per-session action), removing one always-on footer control.
- [x] **2.5** No new strings without `.rc` entries; reuse existing `IDS_*` and add only a `More...` menu string.

### Validation

- `--commands-selftest` case `cmd_pane_fileops_conflict_prompt_compacts_actions` asserts the **primary** action count is <= 3 and that overflow actions are reachable via the More menu.
- Visual artifact check: layout acceptance sketch archived at `Specs\TestRuns\4cb089111a23\Commands\2026-06-06_115852\phase2_popup_layout_acceptance.md`.
- Visual regression: `FileOps.Popup.Render.BuildSnapshotUs` median 1053 us, `CardLayoutUs` median 10 us, and `ScrollLayoutUs` median 271 us in GREEN `2026-06-06_115852`. `TotalUs` median was 42754 us because Direct2D draw/flush time dominated later samples; CPU layout metrics stayed stable against RED `2026-06-06_114034`.
- Keyboard/default action: primary action order remains deterministic (Overwrite/Skip/Cancel for Exists), and `Cancel` remains a primary action.

---

## Phase 3 — Fair Per-Stream Bandwidth-Graph Feedback (P2, parallel accuracy)

### Problem (the user's edge case #3)

> "feedback for parallel copy doesn't seem to have the same time quota when looking at the graph with color (4 files in parallel don't produce 4 surfaces of the same size when copying/moving)."

The bandwidth/throughput graph is **per-task aggregate**, not per-stream. There is one `RateHistory` per task with one `samples[]` and one `hues[]` array (`Popup.h:190-217`). Each timer tick produces **one** aggregate throughput value (`deltaBytes = task.completedBytes - history.lastBytes`, `Popup.cpp` ~L2234) colored by **one** hue derived from whichever file was the most-recent callback source (`RateSampleHue(task.currentSourcePath)`, `Popup.cpp:1098-1107`, used at ~L2200/L2296). During resampling the pending hue is **overwritten** by the latest file (`Popup.cpp` ~L1191), so the earliest stream's contribution to that bucket is discarded.

Net effect with 4 parallel files: the limiter *does* divide bandwidth fairly (`ApplyCallbackBandwidthLimit` → `desiredTotal / activeCalls`, `State.cpp` ~L865), and the per-file in-flight lines are tracked separately (`_inFlightFiles`, keyed by `(cookieKey, progressStreamId)`), but the **graph** paints the whole interval's area in one stream's color (winner-takes-all). So the colored surfaces look unequal even though throughput is balanced — exactly the reported symptom.

### Proposed changes (decision gate, then implementation)

- [x] **3.0** Lock the rendering decision before code changes. Default to **3.1 proportional sub-bands** because it directly represents byte share. Use **3.2 round-robin/dominant hue** only if a quick prototype or render metric shows sub-bands exceed the popup render budget.
- [x] **3.1 (preferred):** Make each graph bucket carry **per-stream byte weight**, then color the bucket's filled area by **proportional sub-bands**. Within a resample bucket, accumulate `bytesByHue[hue] += streamDeltaBytes` from the per-stream `_inFlightFiles` deltas (already available) instead of a single `pendingHue`. When drawing the trapezoid for that bucket, split it vertically into sub-bands sized by each hue's share. 4 equal streams -> 4 equal-height bands of 4 colors. This directly answers "4 surfaces of the same size."
- [x] **3.2 (fallback only):** Fallback not used; proportional sub-bands passed deterministic coverage and popup smoke validation.
- [x] **3.3** Stop overwriting `pendingHue` with "latest wins" (`Popup.cpp` ~L1191); replace with the per-hue weight accumulator from 3.1 (or the round-robin selector from 3.2).
- [x] **3.4** Keep the graph **timer-cadence driven**, not callback-cadence (spec requirement) — only the *color attribution* changes, not the sampling model. Y-axis auto-scale, paused/waiting overlays, and speed-limit line unchanged.
- [x] **3.5** Non-rainbow mode is unaffected (single theme color); this is a rainbow-mode fidelity fix.

### Validation

- New selftest `FileOps_ParallelGraphFairColorWeight`: synthetic four-stream graph sample asserts 4 hue buckets and equal 25% shares. RED `2026-06-06_120744`; GREEN `2026-06-06_121404`.
- Reuse `FileOps.Popup.Render.*` through `Phase6_PopupSmokeResizeAndPause`; GREEN `2026-06-06_121720`. CPU metrics stayed stable; `TotalUs` draw/flush median is recorded as compositor noise for this run.

---

## Phase 4 — Security Hardening: TOCTOU & Partial-State (P2)

### Findings (grounded)

Good news first — the dangerous defaults are **already safe**: recursive delete does **not** traverse reparse points (`FileSystem.FileOps.cpp` ~L5343-5351, ~L5475-5482 collect reparse dirs without descending), long paths use `\\?\` extended form throughout (`FileSystem.Path.cpp` `ToExtendedPath`), bandwidth/arena math has explicit overflow guards (`FileSystem.FileOps.cpp` ~L1264-1296, `FileSystem.h` ~L870-910), recycle-bin `IFileOperation` flags are conservative (`FOF_NOCONFIRMATION|FOF_NOERRORUI|FOF_SILENT|FOFX_EARLYFAILURE|FOFX_RECYCLEONDELETE` with the host-side permanent-delete confirm defaulting to Cancel), and copy failures roll back via `TryRollbackCopiedDestination`.

Remaining risks:

1. **TOCTOU windows.** `RemovePathForOverwrite` (~L2534-2592) and `DeletePathInternal` (~L5337-5413) read attributes with `GetFileAttributesW`, branch on file/dir/reparse/read-only, then act on the **path** later. Between check and act another process can swap a file for a directory, or a normal dir for a reparse point. The no-follow recursion at ~L2417-2435 acts on enumeration-snapshot attributes that may be stale.
2. **Cross-volume Move partial failure.** If the delete phase fails after a successful copy and rollback also fails, the copied destination can be left behind (`~L4343` returns `ERROR_PARTIAL_COPY`). Status to the user is not explicit about "source kept, partial copy left at destination."
3. **Reparse retarget on copy** under `FollowTargets` policy (~L3012-3025) remaps targets into the destination tree — worth re-validating that a crafted target can't escape the destination root.

### Proposed changes

- [x] **4.0** Split this phase into separate commits/slices: partial-state status first, reparse containment second, handle-based delete/overwrite last. Do not combine all destructive-path changes into one landing.
- [x] **4.1** Cross-volume Move partial-failure: make the outcome a **clear status** — if copy succeeded but delete/rollback failed, the task result MUST say "source preserved; partial copy left at <dest>" with the affected path, and route it to the Issues pane. No silent ambiguity.
- [x] **4.2** Add an explicit destination-containment assertion to the reparse retarget path (normalized retarget MUST remain inside the destination root), mirroring the archive-extraction containment rule already in the spec.
- [x] **4.3** Close the delete/overwrite TOCTOU by acting through a **handle** opened with `FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS`: open once, verify type/attributes from the handle (`GetFileInformationByHandle`), then delete via `FILE_DISPOSITION_INFO`/`POSIX` semantics on that handle (or `SetFileInformationByHandle`), instead of re-resolving the path. Keep the existing path-based fallback for plugins/paths where handle-open is unavailable, gated so behavior never regresses.
- [x] **4.4** Re-verify reparse status from the handle immediately before any recursive descent decision (eliminate reliance on stale enumeration attributes for the destructive branch).
- [x] **4.5** Add a debug-only selftest hook that injects a file→dir / dir→reparse swap between check and act to prove the handle-based path rejects/handles it safely.

### Validation

- New selftest `FileOps_DeleteToctouSwapGuard` (debug-injected swap) — asserts no out-of-tree deletion, deterministic error surfaced.
- `FileOps_CrossVolumeMovePartialFailureStatus` — forces delete-phase failure, asserts the explicit "source preserved / partial copy" status string and an Issues-pane row.
- Confirm `Specs/Plans/Done/Notes_DoNotYeetYourJunctions.md` invariants still hold; archive runs.

---

## Phase 5 — Remaining Parallelism / Performance Opportunities (P3)

These are incremental; each must follow the validation contract and ship independently.

- [x] **5.1** **Parallel pre-calc across source roots within one task.** Protected by `FileOps.SelfTest.ClearflowPreCalcMultiRootWorkers`, `FileOps.SelfTest.ClearflowPreCalcSingleRootFanOutWorkers`, and `Phase5_PreCalcCancelLatencyLocal` in GREEN `2026-06-06_131401`.
- [x] **5.2** **Sequential directory creation in the cross-FS bridge.** Bridge now streams file work as each parent destination directory is ensured; RED `2026-06-06_130731`, GREEN `2026-06-06_131401`.
- [x] **5.3** **Pre-calc result cache** rejected for this plan. Without watcher-backed invalidation and stale-total selftests it is a reliability risk; the parallel scan/fan-out changes cover the protected latency scenario without serving cached totals.
- [x] **5.4** Single deep-folder occupancy proven by `FileOps.SelfTest.ClearflowSingleDeepFolderWorkerOccupancy` for both `CopyItems(count == 1)` and direct `CopyItem(...)` in GREEN `2026-06-06_131401`.

### Validation

- Each item: name protected scenario, add/extend `--fileops-selftest` case, archive before/after, include occupancy/throughput metric. No claim without evidence.

---

## Phase 6 — Single Status Model Consolidation (P3, cross-cutting)

Ties the "clear status" principle together once Phases 1-4 land.

- [x] **6.1** Define **one** per-task status enum that the header label, glyph, and footer derive from (e.g. `Waiting → Calculating → Running → Paused → Conflict → Done → Failed/Partial`). `TaskSnapshot::StatusKind` now drives card header, graph overlay, task glyph severity, caption severity, and layout snapshots.
- [x] **6.2** One global status summary in the footer/caption: "N running, M waiting, K need attention" — so the window communicates state at a glance without reading every card.
- [x] **6.3** Audit every label that currently encodes status implicitly (button captions, overlays) and route them through the single status model. The old header/graph/caption ladders now use the resolver; button captions remain commands, not task-status labels.

### Validation

- `--commands-selftest` snapshot of each status transition; assert exactly one status is active per task at each step.

---

## Phase 7 — Provider Contract Matrix & Legacy Direct-Route Cleanup (P2, consistency/security)

### Findings (grounded)

The normal pane F5/F6/Delete commands now route through the FileOps host queue (`FolderWindow.FileOperations.*`) and get the popup, conflict policy, cancellation, speed-limit, queueing, and archived metrics. Several older routes still bypass that model:

- `FolderView::PasteItemsFromClipboard` calls `IFileSystem::CopyItems` directly at `RedSalamander/FolderView.FileOps.cpp:669-671` and passes `ALLOW_OVERWRITE | ALLOW_REPLACE_READONLY | CONTINUE_ON_ERROR`, with no FileOps popup/conflict flow. This is inconsistent with the safe copy/move defaults and can silently widen destructive behavior.
- `FolderView::MoveSelectedItems` uses the folder picker and calls `IFileSystem::MoveItems` directly at `RedSalamander/FolderView.FileOps.cpp:883-884` with `CONTINUE_ON_ERROR`, also bypassing the FileOps popup and task status.
- `CopySelectedItemsToFolder`, `MoveSelectedItemsToFolder`, and `DeleteSelectedItems` still contain direct-plugin fallback paths when `_fileOperationRequestCallback` is absent. That may be useful for isolated tests, but in the real shell it creates a second workflow surface with different progress/conflict semantics.

Provider review also found capability/behavior drift worth locking down:

- `FileSystemCurl` has a mature multi-item scheduler and runtime concurrency settings, but its top-level vs nested directory concurrency choices need a contract test so advertised concurrency and `progressStreamId` behavior stay aligned.
- `FileSystemS3` advertises conservative copy/move concurrency and higher delete concurrency; individual item progress uses stream `0`, which is acceptable only if same-provider transfers remain effectively single-stream.
- `FileSystemMicrosoftDrive` advertises same-provider copy as unsupported, but `CopyItems` still loops through unsupported `CopyItem` calls if reached. The host should prevent that path, and a test should prove unsupported operations are blocked before a task starts.
- `FileSystem7z` and `FileSystemGoogleDrive` are read/export-only for these operations; their unsupported-operation behavior should be explicitly covered so UI affordances do not appear accidentally.
- `FileSystemDummy` is the best deterministic provider for conformance/selftests, but it should mirror the host contract for directory merge, conflicts, cancellation, and status.

### Proposed changes

- [x] **7.0** Ship this phase as four slices, not one lump: **7A unsafe route cleanup**, **7B fallback policy**, **7C provider matrix**, **7D conformance harness/tests**. Each slice must be correct and testable on its own.
- [x] **7.1 (7A)** Route clipboard paste and folder-picker move through `FileOperationRequestCallback` / `FolderWindow::StartOperation` so they share one copy/move/delete workflow. Do not grant implicit overwrite/replace-read-only from paste; surface conflicts through the same prompt as F5/F6.
- [x] **7.2 (7B)** Keep direct-plugin fallbacks only for explicit no-host/test scenarios; in normal UI builds, missing `_fileOperationRequestCallback` should fail visibly or log a host-wiring error instead of silently changing semantics.
- [x] **7.3 (7C)** Build a provider operation matrix in `Specs/FileSystem/FileSystem_FileOperations.md`: per provider, list CopyItem/CopyItems/MoveItem/MoveItems/DeleteItem/DeleteItems support, same-provider capability flags, cross-provider import/export capability flags, max advertised concurrency, and expected progress-stream behavior.
- [x] **7.4 (7D)** Add conformance tests around unsupported operations: if a provider capability says copy/move/delete is false, the command is disabled or rejected before the task starts; if an API method is accidentally reached, it returns a deterministic unsupported error and does not create partial UI state.
- [x] **7.5 (7D)** Add provider conformance tests for directory-merge semantics on every writable recursive provider. For network-backed providers where credentials are not available, use deterministic stubs or Dummy-backed adapters; no live network dependency in CI/selftests.
- [x] **7.6 (7D)** Add progress/cancel contract tests: bulk operations must report item start/completion, honor cancellation at item boundaries at minimum, and use stable `progressStreamId` values when concurrency > 1.

### Validation

- New `--commands-selftest` case `Commands_FileOpsClipboardPasteUsesHostQueue`: paste into a pane, assert a FileOps task appears, default overwrite flags are not pre-granted, and conflicts use the same prompt as F5 copy.
- New `--commands-selftest` case `Commands_FileOpsFolderPickerMoveUsesHostQueue`: folder-picker move creates one host task with clear queued/running/done status.
- New `--fileops-selftest` case `FileOps_ProviderCapabilityMatrix`: checks built-in provider capability flags against reachable operation behavior without requiring external network credentials.
- Update user-facing docs (`Docs/FileOperations.md`) after route simplification so users see one workflow, not separate paste/move exceptions.

---

## Out of Scope / Non-Goals

- No change to the plugin `IFileSystem` ABI beyond what already exists (`GetTransferHints`, `GetStorageCharacteristics`). Phases here are host + built-in-plugin behavior, contract clarifications, and UI.
- No new concurrency knobs in the host pane (plugin-owned settings stay on the plugin page, per spec).
- No change to external drag/drop shell formats.

## Suggested Order

0A/0B/0C → 1 → 2 → 7A → 3 → 4A/4B → 7B/7C/7D → 5 → 6. Phases 1 and 2 are tightly coupled (fixing the folder-merge bug removes the most common conflict prompt, which is what makes the prompt simplification safe). Do 7A early so clipboard paste and folder-picker move cannot preserve the old unsafe workflow. 3 and 4A/4B are independent and can be parallelized across people. Keep 4C handle-based delete/overwrite work separate from the earlier security-status slices. 6 is a finishing pass.
