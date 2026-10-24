# Operation Crosscut — CompareDirectories Remediation: Sync Data-Safety, Cache Coherence & Options Simplification

**Status:** Done - implementation slices complete, durable Compare Directories contracts merged into the authoritative spec, and the unrelated Commands order/input closeout blocker split into its own WIP plan.
**Closeout gate:** closed on 2026-06-24 by explicit split policy. The remaining broad Commands instability is not CompareDirectories correctness work and is now tracked by `Specs/Plans/WIP/Operation_CommandsSelfTestInputIsolation_2026-06-24.md`. Latest test-enabled Debug build is green, focused menu/preference/credential regressions that failed in the broad run are green when filtered, and prior FileOperations/Monitor closeout blockers were cleared in focused reruns. The broad Commands run remains red (`751 passed / 11 failed / 2 skipped`), but that risk is owned by the new Commands plan.
**Date:** 2026-06-15
**Author:** Deep multi-agent code review of the CompareDirectories subsystem (branch `claude/elated-sinoussi-a39160`). 14 specialized review lenses (engine concurrency / cache-memory / content-compare / decision-correctness; window sync-copy / lifecycle / navigation / progress / options / menu+prefs; architecture; spec-conformance; test-coverage) → per-finding adversarial refutation pass (2 independent skeptics for critical/high, 1 for medium/low) → synthesis. **57 candidate findings → 43 confirmed, 2 disputed, 12 rejected.** The two highest-stakes data-safety claims (CX0-1 sync manifest, CX0-2 invalidate version bump) were re-traced by hand in source.
**Builds on / does NOT re-open:** `Specs/Plans/Done/CompareDirectories_MemoryAndHangs.md` and `Specs/Plans/Done/Core_CompareDirectoriesReview.md`. Those addressed memory/hang issues and the prior review; the version/coherence skeleton, the LRU+budget eviction structure, the `(version, cancelToken)` in-flight stamp model, the throttled-notification design, and the size-known content read loop are sound and are NOT reworked here except where a specific defect is listed. Crosscut fixes the defects those passes' green self-tests did not catch.
**Normative spec:** `Specs/Core/Core_CompareDirectories.md` (every "MUST" referenced below is from this document).
**Follow-up plan:** `Specs/Plans/WIP/Operation_CommandsSelfTestInputIsolation_2026-06-24.md` owns the remaining broad Commands order/input failures.

**Scope (anchors are worktree `elated-sinoussi-a39160` @ current HEAD — line numbers are HEAD-relative; re-grep every `~line` before editing):**
- `RedSalamander/CompareDirectoriesEngine.cpp` / `.h` — session/engine: scan + content workers, decision cache, version/coherence, compare-scoped `IFileSystem` wrapper.
- `RedSalamander/CompareDirectoriesWindow.cpp` — lifecycle, sync navigation, selection, directory sync-copy/move.
- `RedSalamander/CompareDirectoriesWindow.Internal.h` — private window class declaration.
- `RedSalamander/CompareDirectoriesWindow.Options.cpp` — options panel (hidden-HWND + DxUi dual model, settings reload).
- `RedSalamander/CompareDirectoriesWindow.Progress.cpp` — progress UI, ETA, watermark, File Operations task card.
- `RedSalamander/CompareDirectoriesWindow.Menu.cpp` — themed menu bar.
- `RedSalamander/Preferences.CompareDirectories.cpp` / `.h` — global Preferences page.
- `RedSalamander/FolderWindow.FileOperations.cpp` + `.State.cpp` — copy/move submission path that compare-mode sync must extend with a resolved-item manifest executor.
- `RedSalamander/SelfTest/CompareDirectories/*` and `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp` / `Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp` — new self-tests.

---

## Closeout Snapshot - 2026-06-24

User approved splitting the unrelated Commands failures into a dedicated WIP plan. This CompareDirectories plan is therefore moved to Done; do not reopen it for the broad Commands failures unless new evidence ties them back to CompareDirectories behavior.

Current state:
- Latest clean test-enabled Debug build: `RSBuildEnableTests=true; .\build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander`, logs `.build/logs/msbuild-20260624_142759_244.log` and `.build/logs/msbuild-20260624_142801_196.log` (`0 warning(s), 0 error(s)`).
- The previous menu mouse-open blocker is fixed and focused green: `.build/logs/run-commands-menu-mouse-open-20260624_143208_163.log` (`cmd_app_menuBar_mouse_open_keeps_popup_selection_clear`, `1 passed / 0 failed`). The test still uses real cursor movement only for the path that requires it and displays the large "don't touch the mouse" warning there.
- Latest broad Commands closeout run: wrapper `.build/logs/run-commands-closeout-20260624_143239.out.log`, archived results `Specs/TestRuns/7d3a1247382a/Commands/2026-06-24_150328/commands_results.json` (`751 passed / 11 failed / 2 skipped`, duration about 30m45s).
- The exact 11 broad-run failures passed when rerun in focused clusters:
  - Credential prompt cluster: `.build/logs/run-commands-failing-credential-20260624_150453_924.log` (`2 passed / 0 failed`).
  - Preferences cluster: `.build/logs/run-commands-failing-preferences-20260624_150512_968.log` (`7 passed / 0 failed`).
  - Menu cluster: `.build/logs/run-commands-failing-menu-20260624_150555_203.log` (`2 passed / 0 failed`).
- A wider plugin/connection-manager/credential replay still reproduced order/input-sensitive UIA instability: `.build/logs/run-commands-plugin-connection-credential-20260624_150743_323.log` (`41 passed / 5 failed`). Notable signal: `cmd_connection_manager_window_uses_localized_strings_for_dynamic_labels` expected the localized default name `New connection` but observed `rce`, indicating text reached the wrong focused control or stale input leaked into the run.
- FileOperations prior full-gate failure was not reproducible in focused reruns: `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-21_205655` (`19 passed / 0 failed / 0 skipped`).
- Monitor ETW latency passed after rebuilding Monitor with tests enabled: `Specs/TestRuns/7d3a1247382a/Monitor/2026-06-21_210002`.

Current Commands failures from the broad suite:
- `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible`
- `cmd_connection_credential_prompt_live_dx_interaction`
- `cmd_preferences_dialog_plugins_theme_cycle_keeps_surface_legible`
- `cmd_preferences_dialog_viewers_theme_cycle_keeps_surface_legible`
- `cmd_preferences_dialog_viewers_live_search_dx_interaction`
- `cmd_preferences_dialog_viewers_add_update_live_dx_interaction`
- `cmd_preferences_dialog_viewers_selection_survives_legacy_list_clear`
- `cmd_preferences_dialog_plugin_tree_selection_keeps_single_visible_pane`
- `cmd_preferences_dialog_plugins_main_selection_survives_legacy_list_clear`
- `cmd_app_menuBar_arrow_switches_top_level_popup`
- `cmd_app_menuBar_submenu_placement_matches_spec`

Transferred follow-up focus, owned by `Specs/Plans/WIP/Operation_CommandsSelfTestInputIsolation_2026-06-24.md`:
- Treat the broad Commands blocker as an input/focus isolation problem until disproven. The `rce` stray-text symptom in the wider replay is stronger evidence than the isolated pass/fail split alone.
- Minimize direct mouse/keyboard paths; keep them only for coverage that cannot be driven through UIA/messages. Any path that moves the real cursor must create the on-screen warning before movement and dismiss it immediately after.
- Harden the UIA value-entry/cancel-reopen helpers used by credential, plugin configuration, connection manager, and Preferences tests so they verify active window/focus ownership before mutating controls and drain stale input/messages after modal teardown.
- Rebuild `RedSalamander` with `RSBuildEnableTests=true`, rerun the wider plugin/connection/credential replay, rerun full Commands, then run `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` for closeout.

---

## Why this plan exists

The CompareDirectories subsystem is **not currently safe to represent as a "sync" feature.** The review found three independent silent-data-loss paths, one confirmed heap-corruption race, and a complete failure of the headline cross-plugin scenario (compare a local folder against S3/FTP). The most damning is **CX0-1**: the spec's entire "Directory copy/move behavior (sync)" section — build a diff manifest, copy only the different items — is **unimplemented**. Copy/Move-to-other-pane delegates straight to the generic recursive file-op, so a "sync" copies the *whole* subtree including identical files, and **Move silently deletes identical files from the source.**

The architecture amplifies the risk: the Options panel keeps a *full hidden Win32 control set as its load-bearing state model* behind a DxUi mirror (**CX3-3** — the user-requested removal), and `CompareDirectoriesWindow` is a ~13.7k-line god-class. Every option must be wired in three places, which is precisely how the UI defects below accumulated.

**Standing rule for every Crosscut slice:** a slice is not done because a self-test is green. The current suite is green *because the risky paths are unexercised* (no test for diff-manifest construction, equal-size+equal-mtime+different-content, unknown-size streaming, `SetRoots` reset, or content-worker-vs-`Invalidate` races). For each P0/P1 slice the **RED-on-current-HEAD demonstration is part of the required proof** — the test must fail for the dangerous reason before the fix and pass after.

---

## Challenge Review Amendments — apply before implementation

This plan is not implementation-ready unless the following architecture corrections are treated as part of the contract, not optional cleanup:

1. **Do not force `(source,destination)` sync pairs through the current `FolderWindow` file-op API.** `FileOperationState::StartOperation(...)` currently accepts `sourcePaths + destinationFolder`; `CopyItems`/`MoveItems` append each source leaf to that one folder. That cannot preserve nested relative paths from a sync manifest. CX0-1 therefore MUST add a host-side resolved-item path in the file-operation layer, for example:

   ```cpp
   struct FileOperationResolvedItem
   {
       std::filesystem::path sourcePath;
       std::filesystem::path destinationPath;
       FileSystemFlags flags = FILESYSTEM_FLAG_NONE; // per item; recursive only for explicit WholeSubtree entries
   };
   ```

   Add a `StartResolvedOperation(...)` overload or equivalent internal task mode. It must reuse the existing per-item conflict/progress/circuit-breaker/cross-filesystem bridge code, but it must take `destinationPath` from the resolved item instead of `JoinFolderAndLeaf(destinationFolder, leaf)`. Do **not** change the plugin ABI and do **not** add fake path encodings to `sourcePaths`.

2. **The diff manifest belongs to the compare session, not the window.** The window should collect selected relative paths and ask the session for a cache-only manifest. The session owns decision keys, version checks, pending states, ignore behavior, and relative-to-absolute translation. The window owns only command routing, user messaging, and submitting the returned manifest to file operations.

3. **Directory sync must fail closed while the subtree is incomplete.** A command handler must not “force-drain” compare work by doing synchronous traversal or waiting on workers from the UI thread. If any selected subtree has missing, failed, stale, `SubdirPending`, or `ContentPending` decisions, the command requests/reschedules high-priority scans and shows a localized “comparison still in progress / cannot sync yet” status. It does not build a partial manifest.

4. **One-sided directories are the only recursive exception.** The earlier blanket “never pass a directory with `FILESYSTEM_FLAG_RECURSIVE`” is too broad, while passing mixed directories recursively is unsafe. The manifest may include a recursive `WholeSubtree` item only when the selected item or descendant is known to exist only on the source side. Mixed existing-on-both-sides directories are expanded into directory-shell creation plus explicit file/subtree items; identical descendants are never represented.

5. **Options simplification should use an explicit draft model.** Do not first polish the hidden-HWND model and then remove it. Introduce a `CompareOptionsDraft`/`CompareDirectoriesSettings` draft as the single editable state, bind DxUi controls directly to it, and delete the hidden Win32 controls plus dead split-host layout branches in one Options slice. Preferences.CompareDirectories already demonstrates the DxUi-only pattern.

6. **Self-test registration proof must match the actual suite.** Compare engine cases must be added to `kCompareCaseNames`, appear in `CompareDirectoriesSelfTest::ListCases(...)`, and be invoked by the included case dispatchers. Commands/Preferences cases must be reachable from their `Run*Cases(...)` dispatcher. Do not describe this as a generic “case registry” unless one is introduced.

These amendments intentionally reduce the implementation surface: a session-owned manifest planner, one file-op resolved-item mode, one options draft model. Any implementation that instead spreads manifest logic through menu/window code or overloads `CopyItems` semantics is not following this plan.

---

## Implementation Tracking Checklist (update first, before editing code in a slice)

Use `[ ]` not started, `[~]` in progress, `[x]` complete, `[blocked]` needs a product decision.

| State | Slice | Phase | Implementation unit | Required proof before `[x]` |
|-------|-------|-------|---------------------|-----------------------------|
| [x] | CX0-1 | P0 | Session-owned cache-only **sync manifest** + file-op resolved-item mode; Move deletes manifest items only after destination success; recursive only for explicit one-sided `WholeSubtree`; sync confirmation/accurate count | Verified 2026-06-20: manifest operation-set, Move manifest safety, missing-cache and pending-content/subdir fail-closed tests, route-level resolved file-op side effects, localized sync confirmation/status copy, and `compare.sync.manifest.*` archived perf counters are green. |
| [x] | CX0-2 | P0 | Scoped invalidation preserves unrelated cache while pruning affected scan/content queues and in-flight stamps; scan writeback requires the live in-flight stamp | Verified 2026-06-20: added `invalidate_drops_stale_inflight_scan_writeback`; focused stale-writeback and existing `invalidateForPath` cases green; `.\build.ps1 -ProjectName RedSalamander` green; full Compare suite green (151 total / 127 passed / 0 failed / 24 skipped). |
| [x] | CX0-3 | P0 | Gate all 3 decision-cache inserts on `SUCCEEDED(decision->hr)`; failed decisions returned but not stored | Verified 2026-06-20: added `failed_enumeration_retries_without_version_bump`; focused case green; `.\build.ps1 -ProjectName RedSalamander` green; full Compare suite green (150 total / 126 passed / 0 failed / 24 skipped). |
| [x] | CX0-4 | P0 | Snapshot keys in `FlushPendingContentCompareUpdatesBudgeted` before apply (kill dangling-iterator UB) | Verified 2026-06-20: source review confirms no live `_pendingContentCompareUpdates` iterator crosses `ApplyPending…`; `.\build.ps1 -ProjectName RedSalamander` green; `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2` green (125 passed / 0 failed / 24 skipped). ASAN not run locally. |
| [x] | CX0-5 | P0 | `GetCurrentPath`→`GetCurrentPluginPath` at the TryMakeRelative-fed compare-window sites | Verified 2026-06-20: added non-`file` dummy-plugin command selftest covering Restore/Invert selection + post-completion empty-state; `.\build.ps1 -ProjectName RedSalamander` green; focused Commands case green; full Compare suite green (150 total / 126 passed / 0 failed / 24 skipped). |
| [x] | CX0-6 | P0 | Options hot-reload: register participant in `ShowOptionsPanel`; clear `_optionsStaleFromExternalReload` on reopen | Verified 2026-06-20: focused hot-reload Commands case green; `cmd_compare_directories_` family green (13 passed / 0 failed); Debug build green. |
| [x] | CX0-7 | P0 | Preferences Compare page renders its 4 section headers (`layoutHeader` lambda + fixed index→section mapping) | Verified 2026-06-20: `PreferencesDebugSnapshot.compareDirectoriesVisibleSectionHeaderCount == 4`; `.\build.ps1 -ProjectName RedSalamander` green; focused Preferences statics case green; Compare Directories Preferences family green (7 passed / 0 failed). |
| [x] | CX1-1 | P1 | Decision-cache budget enforcement: evict non-pinned/non-ancestor `anyPending`; call eviction on aggregation paths; single-pass eviction scan | Verified 2026-06-20: pending-subdir aggregate snapshots allow non-pinned pending child decisions to evict while preserving ancestor repair; focused wide pending budget case asserts current and high-water bytes stay within 4x budget and archived perf counters are green. |
| [x] | CX1-2 | P1 | Leave-scope prompt deferred out of the navigation callback (revert + PostMessage, or `_leaveScopePromptActive` guard) | Verified 2026-06-20: pane reverts before the callback returns, prompt is queued via posted compare-window message, stale run/path guards are enforced, and focused Commands regression is green. |
| [x] | CX1-3 | P1 | Unambiguous window ownership (no unique_ptr + self-delete double-free on create failure) | Verified 2026-06-20: create-in-progress guard prevents self-delete while `CreateWindowExW` is on the stack; injected `OnCreate` failure leaves no window and normal open still works. |
| [x] | CX1-4 | P1 | `_dxMenuBar` null re-check after modal `ContextMenu::Show` | Verified 2026-06-20: post-modal `_hWnd`/`_dxMenuBar` guard present; normal DxUI popup open/dismiss regression remains green. |
| [x] | CX2-1 | P2 | Trailing space/dot name-collision no longer silently drops a same-side entry | Selftest: two same-side names differing only by trailing dot/space both represented. |
| [x] | CX2-2 | P2 | Ignore patterns applied to the folder's own leaf/ancestors on direct navigation | Selftest: navigate directly into an ignored-dir relative path → empty/identical, not compared. |
| [x] | CX2-3 | P2 | Leave-scope Cancel with empty `previousPath` falls back to root (or proceeds + cancels) | Verified 2026-06-20: pending leave-scope navigation falls back to the session root when no previous path exists; existing Commands regression proves the pane is not stranded out-of-scope while compare remains active. |
| [x] | CX2-4 | P2 | Watermark pulse driven by its own state, not the spinner control | Verified 2026-06-20: progress pulse timer is shared by spinner and pane watermark; animated watermark invalidation no longer depends on `_scanProgressBar`. |
| [x] | CX2-5 | P2 | `ScheduleDecisionRefresh` has a fallback when `SetTimer` fails | Verified 2026-06-20: timer failure logs `ErrorWithLastError` and posts a one-shot refresh fallback so pending refreshes are not stranded. |
| [x] | CX2-6 | P2 | ETA clamped / "estimating…" above sane threshold | Verified 2026-06-20: ETA is shown only for finite estimates with rate >= 1 KiB/s and duration <= 7 days. |
| [x] | CX2-7 | P2 | Task-card trailing flush after throttle window | Verified 2026-06-20: throttled task-card updates schedule a one-shot trailing flush and cancel it on finish/dismiss/teardown. |
| [x] | CX2-8 | P2 | Reset `_optionsWheelRemainder` when not scrollable | Verified 2026-06-20: non-scrollable options layout and wheel routes clear accumulated remainder before later scrollability changes. |
| [x] | CX2-9 | P2 | `compareContent` persisted value AND-masked with support (or unchecked on disable) | Verified 2026-06-20: unsupported content compare clears the disabled toggle and save masks persisted `compareContent=false`. |
| [x] | CX3-1 | P0-foundation | Extract one `ApplyCriteriaDiffAndSelection(...)`; call from all 3 engine sites | Verified 2026-06-20: `rg` shows one definition + 3 call sites; `.\build.ps1 -ProjectName RedSalamander` green; `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2` green (125 passed / 0 failed / 24 skipped). **Done before CX0-1.** |
| [x] | CX3-2 | P3 | Fold dead split-host DxUi layout branch deletion into the Options draft-model migration unless it lands as a separate no-behavior cleanup | Verified 2026-06-20: separate no-behavior cleanup removed never-attached per-control Dx placeholders and split-host fallback branches. |
| [x] | CX3-3 | P3 | **Remove hidden-Win32 shadow control model — DxUi + `CompareOptionsDraft` are the single source of truth in Options** | Verified 2026-06-20: Options-panel selftests read/write through DxUi + draft state; grep confirms hidden toggle/edit/static owners and legacy sync round-trips are deleted. |
| [x] | CX3-4 | P3 | Decompose god-class state into ChromeController / OptionsPanelController / ProgressController | Verified 2026-06-20: flat chrome/options/progress state moved into owned controllers; compare, options, chrome, and progress batteries green. |
| [x] | CX3-5 | P3 | Drop redundant button-host attach (B-1); reuse shared `DxUiNativeMenuInterop` (MD-2); fence `GetOrComputeDecision` behind `ENABLE_TESTS` (API-1) | Verified 2026-06-20: duplicate attach removed, menu conversion uses shared helper, blocking decision API is test-only. |
| [x] | CX4 | P4 | All pinning + standing self-tests (T-1..T-8, L-1b, PT-1) listed and reachable from the actual compare/commands dispatchers | Verified 2026-06-20/21: required cases are listed/filterable/dispatched; six missing standing cases added; full Compare suite and focused command dispatchers are green and evidence archived. |
| [x] | Closeout | Gate | Move the plan to `Specs/Plans/Done/` only after the full-suite closeout gate passes or unrelated blockers are split into their own tracked plan | Closed 2026-06-24: CompareDirectories work is complete and durable contracts are in `Specs/Core/Core_CompareDirectories.md`. The remaining broad Commands failures are unrelated order/input instability and are transferred to `Specs/Plans/WIP/Operation_CommandsSelfTestInputIsolation_2026-06-24.md`. FileOperations `Phase7_AutoConcurrencyHints` and Monitor ETW latency were cleared in focused reruns. |

---

## Guiding Principle

> **Simple, always working, clear status.** A "sync" must never copy an identical file or delete an unchanged source file. Never serve a stale decision after a mutation. One source of truth per control. Prove every destructive path on the user's real route.

## Performance / Safety Validation Contract (mandatory, per repo rules)

Compare Directories is a performance-sensitive subsystem (`Specs/Core/Core_CompareDirectories.md` "Performance Validation Contract"). Every slice that touches destructive correctness (**CX0-1, CX0-2, CX0-3, CX2-1**) MUST include a **tree-equality / byte-integrity assertion** and an **out-of-tree sentinel assertion**, and MUST archive before/after runs under `Specs/TestRuns/<machine>/Compare/<timestamp>/`.

- New compare self-test cases MUST be added to `kCompareCaseNames`, appear in `CompareDirectoriesSelfTest::ListCases(...)`, and be invoked by the included case dispatcher. Commands/Preferences tests must be reachable from their `Run*Cases(...)` dispatcher. Verify listability, filterability, and actual execution as part of CX4.
- Run via `--compare-selftest` (or `.\Tools\Run-AllTests.ps1 -Suite Full`, ≈18-20 min). A **session-global mutex serializes runs across worktrees — expect exit 3 on contention** (per the `fileops-selftest-run-timing` memory); do not interpret exit 3 as a failure.
- Plugin/engine perf counters land in **ETW**, not the readable log. Wire any `compare.ui.*` / cache-budget high-water assertions to ETW counters, not stdout. CX0-1 must add or reuse `compare.sync.manifest.*` counters for manifest build items, blockers, elapsed time, and submitted resolved-item count.

---

## Phase P0 — Stop the data loss / correctness / crashes (ship-blocking; slices independent unless noted)

### CX0-1. `[CRITICAL]` Directory sync Copy/Move has no diff manifest — copies the whole subtree; Move deletes identical source files
*(covers findings: `no-diff-manifest-whole-subtree-copy`, `sync-copy-diff-manifest-missing`, `cross-vs-same-context-copy-inconsistency`, `sync-copy-no-overwrite-confirmation-spec-gap`)*

**Symptom (reachable, normal use):** In compare mode, Copy-to-other-pane on a directory that mostly matches re-copies every identical descendant (slow, overwrites identical destination files). Move-to-other-pane is worse: it removes identical files from the *source* even though the user intended to transfer only differences — **silent data loss / unexpected source mutation.**

**Root cause:** `CompareDirectoriesWindow::OnCommand` routes `IDM_PANE_COPY_TO_OTHER` / `IDM_PANE_MOVE_TO_OTHER` straight to `_folderWindow.CommandCopyToOtherPane(pane)` / `CommandMoveToOtherPane(pane)` (`CompareDirectoriesWindow.cpp` ~868-869). Those pass the raw selected directory paths plus `FILESYSTEM_FLAG_RECURSIVE` (`FolderWindow.FileOperations.cpp` ~1058-1138). In the common same-context case the engine calls `_fileSystem->CopyItem(dir, RECURSIVE)` where `_fileSystem` is the compare wrapper, which delegates **unconditionally** to `_baseFs` with the recursive flag (`CompareDirectoriesEngine.cpp` ~3656-3759). The base plugin then runs its own on-disk recursion, blind to compare decisions → the entire real subtree is transferred. There is **no diff-manifest code anywhere** (repo-wide grep for `manifest`/`SyncCopy`/`sourceAbsolutePath`/`BuildSyncList` returns nothing). The spec's "Directory copy/move behavior (sync)" section (`Core_CompareDirectories.md` ~573-587) is entirely unimplemented.

*Divergence trap:* when the two panes happen to use *different* plugins (cross-FS bridge), `CopyDirectorySequential` enumerates via the filtered compare wrapper (`State.cpp` ~6877) and accidentally skips identical files — so the **same gesture produces different result sets depending on plugin identity**, and cross-FS testing masks the same-FS data loss.

**Fix architecture (do this exact shape; do not invent a window-local traversal):**
1. Add a small compare-owned manifest model in `CompareDirectoriesEngine.h`, for example `CompareSyncManifest`, `CompareSyncManifestItem`, `CompareSyncManifestBlocker`, and `CompareSyncManifestStatus`.
   - Each item carries `sourceAbsolutePath`, `destinationAbsolutePath`, `relativePath`, and a kind/flag: `File`, `DirectoryShell`, or `WholeSubtree`.
   - `WholeSubtree` is the only kind that may carry `FILESYSTEM_FLAG_RECURSIVE`, and only when the item exists exclusively on the source side.
   - `DirectoryShell` means create/ensure the destination directory before descendant file items; it does not copy/delete a subtree.
2. Expose a cache-only session method, e.g. `TryBuildSyncManifest(ComparePane fromPane, std::span<const std::filesystem::path> selectedRelativePaths, CompareSyncManifest& out, CompareSyncManifestBlocker& blocker) noexcept`.
   - It may call `TryGetCachedDecision` or inspect cache state under `_mutex`; it must not call `GetOrComputeDecision`, `ReadDirectoryInfo`, filesystem I/O, or wait on workers.
   - It snapshots the current `_version` and rejects any stale/missing decision instead of mixing versions.
   - It treats `FAILED(decision->hr)`, missing decisions, `anyPending`, `SubdirPending`, and `ContentPending` as `NotReady`/`Failed` blockers. Absence of identical entries when `keepIdenticalItems=false` is acceptable; absence is not a reason to copy.
3. For selected directories that exist on both sides, build the manifest by walking cached descendant decisions and collecting only source-selected/different items. If `compareSubdirectories=false`, block directory sync with a localized status until the product/spec explicitly defines a shallow-only sync mode; a destructive directory sync without subtree decisions is not safe.
4. For one-sided source directories, emit a single `WholeSubtree` item. This is the safe recursive fast path because every descendant is new to the destination by definition.
5. Add a host-side file-operation resolved-item mode (no plugin ABI change). The task executes each manifest item with the resolved destination path from the manifest, not by recomputing `JoinFolderAndLeaf(destinationFolder, leaf)`.
   - Reuse existing per-item conflict handling, circuit breaker, progress, pause/cancel, and cross-filesystem bridge code.
   - Ensure destination parent directories before file copies/moves using the destination side's `IFileSystemDirectoryOperations` where available; fail with a clear diagnostic if the destination provider cannot create required parents.
   - Same-context copy may use `CopyItem` for `File` and `WholeSubtree` items. Same-context move may use `MoveItem` only as the implementation of a single manifest item when it cannot touch unlisted descendants (`File`, or explicit one-sided `WholeSubtree` with recursive flag). Cross-FS move remains bridge copy + source delete. In every move path, source deletion is restricted to the manifest item only and only after the destination operation succeeds.
6. In `CompareDirectoriesWindow::OnCommand`, intercept Copy/Move-to-other-pane while compare mode is active. Gather selected/focused plugin paths, convert to selected relative paths, ask the session for a manifest, and submit the returned manifest. If the manifest is not ready, request high-priority scans for the blocker paths and show status; do not fall back to generic recursive copy/move.
7. Surface a sync confirmation and accurate task-card item count reflecting manifest items before a destructive Move. The confirmation must say when a one-sided directory will be copied/moved as a whole subtree.
8. After completion, invalidate both source and destination manifest paths (and ancestors) via the corrected CX0-2 path so the visual state converges without a full rescan.

**Required proof:** new `Crosscut_SyncManifestCopiesOnlyDifferences` (T-8) — multi-depth left/right tree mixing identical + different + only-on-one-side files; drain scan to idle + flush subdir/content updates; run the session manifest builder and assert the **exact** operation set: source path, destination path, relative path, kind, and flags. Different files present, identical files absent, one-sided subtree represented as exactly one recursive `WholeSubtree`, mixed directories represented without recursive flags. Add `Crosscut_SyncMoveKeepsIdenticalSource` asserting Move does not delete identical source files. Add a not-ready case proving pending content/subdir state blocks sync and schedules a high-priority scan instead of producing a partial manifest. All RED on current HEAD. Byte-equality + out-of-tree sentinel assertions per the Validation Contract. **Largest slice — see deeper refactors.**

**Status 2026-06-20:** complete. Implemented:
- `CompareSyncManifest*` model + session-owned `TryBuildSyncManifest(...)`, cache-only, version-stamped, with `File` / `DirectoryShell` / source-only recursive `WholeSubtree` items.
- Resolved-item file-operation mode carrying exact `sourcePath`/`destinationPath`, per-item flags, and `DirectoryShell` placeholders; resolved operations run single-item ordered execution to preserve manifest order.
- Compare-window Copy/Move command interception; no fallback to generic recursive copy when a manifest is empty, not-ready, failed, or unsupported.
- Completion invalidation uses exact resolved destination paths when available.
- Destructive resolved Move now uses sync-specific localized confirmation text with manifest item and whole-subtree counts.
- Not-ready sync blockers surface localized "comparison still in progress" status and schedule high-priority scan work.
- `compare.sync.manifest.*` counters cover manifest build count, elapsed time, ready item count, blocker reason/count, and submitted resolved-item count.

Validation:
- During validation, `sync_manifest_nested_differences_only` exposed a self-deadlock: `TryBuildSyncManifest(...)` held `_mutex` while emitting through `ResolveAbsolute(...)` → `GetRoot(...)` → `_mutex`. Fixed by adding a root-snapshot resolver (`ResolveAbsoluteFromRoot`) and by correcting whole-subtree coverage key ordering.
- `.\build.ps1 -ProjectName RedSalamander` green: `.build/logs/msbuild-20260620_191715_299.log` (`0 warning(s), 0 error(s)`).
- `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter sync_manifest_ -TimeoutMultiplier 2` green (`4 passed / 0 failed`): `sync_manifest_nested_differences_only`, `sync_manifest_move_preserves_identical_children`, `sync_manifest_not_ready_without_cached_decision`, `sync_manifest_pending_blocks_and_schedules`; archive `Specs/TestRuns/LT-PF5VDAGE/Compare/2026-06-20_192106_sync_manifest_cx0_1_focus_final/`.
- `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2` green (`155 total / 131 passed / 0 failed / 24 skipped`); archive `Specs/TestRuns/LT-PF5VDAGE/Compare/2026-06-20_192323_compare_full_cx0_1_final/`.
- `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOpsFamily_ClearflowPhase04_Security -TimeoutMultiplier 2` green (`6 passed / 0 failed`); archive `Specs/TestRuns/LT-PF5VDAGE/FileOps/2026-06-20_192518_resolved_items_route_cx0_1_final/`.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -EnableExit` green (`5 passed / 0 failed`).
- `Invoke-Pester -Script .\Tools\Tests\ResourceLocalizationContracts.Tests.ps1 -EnableExit` green (`4 passed / 0 failed`).
- A broad `cmd_compare_directories_` rerun was attempted after this slice but exceeded a 180s tool timeout in the existing `cmd_compare_directories_options_live_dx_body_interaction` path, so it is not counted as CX0-1 route proof.

### CX0-2. `[HIGH]` `InvalidateForAbsolutePath` lets in-flight workers re-insert stale pre-mutation decisions
*(finding: `invalidate-for-path-no-version-bump`; hand-verified)*

**Symptom:** After a sync copy/move/delete in the compare window, the affected folder can show stale results — items just made identical still shown as different, or freshly created/deleted files mis-reported. The user makes data-movement decisions on this view, so a stale "different" causes redundant copies and a stale "identical" skips a needed copy.

**Root cause:** `InvalidateForRelativePathLocked` erases the affected cache entries + pending content updates under `_mutex` and ends with `++_uiVersion;` (`CompareDirectoriesEngine.cpp` ~621), but it did not drop affected keys from `_scanScheduledKeys` / `_scanInFlightKeys` nor did scan writeback verify that its in-flight stamp was still live. The directory enumeration in `ComputeDecisionForFolder` runs **without `_mutex`**; the lock is only re-acquired for writeback, formerly guarded solely by `if (_version.load() == job.version) { _cache[job.key] = storedDecision; ... }` (~3228). A worker that enumerated the folder *before* the file op and finishes *after* the invalidation keeps `job.version == _version`, passes the old guard, and re-inserts a pre-mutation decision into the just-invalidated key. `OnFolderWindowFileOperationCompleted` calls this on every file-op completion while workers run (`CompareDirectoriesWindow.cpp` ~2586-2603).

**Fix:** Use the scoped design, not a global `_version` bump. Existing `invalidateForPath` coverage requires sibling subtrees to remain cached after a targeted invalidation, so a full generation bump is too broad for this codebase.

- Keep `_version` stable for scoped `InvalidateForAbsolutePath(...)` so unrelated cached decisions remain reusable.
- Erase affected decision-cache entries, pending content updates, pending subdir updates, scan queues, high-priority duplicate markers, scheduled keys, scan in-flight stamps, content queues, and content in-flight stamps using the same case-insensitive exact/subtree key semantics as the decision cache.
- Correct scan/content active counters when invalidation removes scheduled or in-flight work, and notify scan/content waiters so blocked workers do not strand behind erased queues.
- Require the original `_scanInFlightKeys` `(version, cancelToken)` stamp to still exist before a scan writes a decision, enqueues child scans, or queues pending-subdir propagation.

Land **before/with CX0-1**: CX0-1's Move issues deletes that hit this exact path; without CX0-2 the manifest's own mutations get stale-overwritten.

**Required proof:** extend T-3 with an in-flight scan/writeback race: during a delayed background scan, mutate and call `InvalidateForAbsolutePath`, then assert the affected scan/content queues and in-flight maps are drained, the stale worker does not repopulate the cache, and a fresh post-mutation decision does not show the removed item. Also keep the existing `invalidateForPath` sibling-cache test green to prove scoped invalidation did not become a full cache reset.

**Status 2026-06-20:** complete via `invalidate_drops_stale_inflight_scan_writeback`; scoped invalidation now prunes affected queued/in-flight scan and content work, scan writeback is gated by the live in-flight stamp, and unrelated cached subtrees remain reusable (`invalidateForPath` green). Debug build, focused stale-writeback case, focused `invalidateForPath`, and full CompareDirectories suite are green.

### CX0-3. `[HIGH]` Failed-enumeration decisions are cached forever, violating the spec retry contract
*(covers findings: `failed-decision-cached-no-retry`, `failed-enumeration-cached-not-retried`)*

**Symptom:** A folder that hits a transient/fixable enumeration failure (NTFS access-denied later corrected, S3/FTP throttle/timeout, network blip, locked folder) stays stuck showing the failure on every navigation back — no retry — until a full Rescan.

**Root cause:** Both decision-producing paths cache the decision whenever its version matches, with **no `SUCCEEDED(decision->hr)` guard**. `ComputeDecisionForFolder` returns a decision carrying the failed HRESULT from `TryReadDirectoryEntries` (~2602/2617); `ScanWorker` stores it into `_cache[job.key]` (~3228-3235) and `GetOrComputeDecision` into `_cache[rootKey]` (~3071-3077) unconditionally. `TryGetCachedDecision` (~745-759) then serves the cached failure on any version match. The spec MUST (`Core_CompareDirectories.md` ~344-349): "The failed decision is **not cached** — the engine retries enumeration on the next access."

**Fix:** Gate all three inserts on `SUCCEEDED(decision->hr)` (the missing-folder case is already mapped to `S_OK` and stays cacheable). Return failed decisions to the immediate caller (so the wrapper's `ReadDirectoryInfo` can surface the HRESULT in the details line) but do not store them; guard the cache-hit fast paths in `GetOrComputeDecision`/`TryGetCachedDecision` to not serve a stale failed entry.

**Required proof:** L-1b — `CreateReadDirectoryBehaviorFileSystem` with `forcedHr=E_ACCESSDENIED` then `S_OK` (toggle): first access asserts `FAILED(hr)`; after flipping to success, next access asserts `SUCCEEDED(hr)` and correct items **without a version bump** (proves not cached). Assert decision-cache byte estimate did not retain the failed entry. RED on current HEAD. **Status 2026-06-20:** complete via `failed_enumeration_retries_without_version_bump`; failed decisions are not cached or counted, retry succeeds at the same version, and full CompareDirectories suite is green.

### CX0-4. `[HIGH]` (disputed 1/2, mechanically verified) Dangling iterator in `FlushPendingContentCompareUpdatesBudgeted` → heap corruption
*(finding: `flush-pending-content-iterator-invalidation`)*

**Symptom:** Intermittent heap corruption / crash during the UI decision-refresh flush when the decision cache is over its 300MB budget while content-compare updates are pending — exactly the large S3↔FS subtree-with-content scenario.

**Root cause:** The flush loop copies `key = it->first`, does `++it`, then calls `ApplyPendingContentCompareUpdatesLocked(key)` (`CompareDirectoriesEngine.cpp` ~698-700). That call invokes `MaybeEvictDecisionCacheLocked()` when it applies an update (~1892), which does `_pendingContentCompareUpdates.erase(candidateKey)` (~1223). If the evicted candidate is the folder `it` currently points at, `it` is invalidated; the next loop test/deref is UB. A pending-update folder is eligible for eviction (eligibility only requires not-pinned, not in `_pendingSubdirUpdates`, not `anyPending` — none guaranteed). One verifier disputed reachability, but the erase mechanism was confirmed and the snapshot fix is correct regardless. Low effort, high blast radius — do early.

**Fix:** Do not hold a live `_pendingContentCompareUpdates` iterator across `ApplyPendingContentCompareUpdatesLocked`. Snapshot up to `maxFoldersToApply` keys into a local `std::vector<std::wstring>` under the lock, then iterate the vector re-checking membership; or re-fetch `begin()` each iteration (as `FlushPendingSubdirUpdatesBudgeted` already does). Sequence **before CX1-1** (CX1-1 changes how eviction touches the pending-updates map).

**Required proof:** T-3 with a low cache budget (`SetDecisionCacheBudgetBytesForSelfTest`) + many pending content updates + concurrent `Invalidate`: no crash under ASAN/repeated runs; code review confirms no live iterator across the apply call. **Status 2026-06-20:** implementation now snapshots keys before applying, so no `_pendingContentCompareUpdates` iterator survives across `ApplyPendingContentCompareUpdatesLocked`; Debug build and full CompareDirectories suite are green. ASAN was not run locally.

### CX0-5. `[HIGH]` Non-`file` plugin compares are broken — display-form path fed to `TryMakeRelative`
*(finding: `getcurrentpath-vs-pluginpath-trymakerelative`)*

**Symptom:** For any cross-plugin / non-`file` compare (S3↔FS, FTP, 7z — the headline scenario): **Restore/Invert Differences Selection do nothing**, the post-completion "No differences" / "doesn't exist in this hierarchy" empty-state is cleared instead of applied, and panes can appear stuck on "Comparing…".

**Root cause:** Several sites feed `FolderWindow::GetCurrentPath(pane)` (the display/history URL form `shortId:ctx|/path`) into `CompareDirectoriesSession::TryMakeRelative()`, which expects the **plugin** path form (`/path`). For non-`file` roots, `TryMakeRelative` normalizes the URL to `/s3:ctx|/sub` which never prefix-matches root `/sub` → returns `nullopt` → the feature silently no-ops. Affected: `CompareDirectoriesWindow.cpp` ~1054, 1058, 1073 (Restore/Invert selection) and `CompareDirectoriesWindow.Progress.cpp` ~203, 208, 235, 1118, 1122 (empty-state + decision-refresh). The working root path uses `GetCurrentPluginPath` (~2004-2010). For `file` plugins display==plugin form, so file-based selftests never see it. Note `CompareDirectoriesWindow.cpp:2247` (`hasContextChanged`) is a **correct** use of `GetCurrentPath` (it parses the display URL via `TryParseLocation`) — leave it.

**Fix:** Replace `GetCurrentPath` with `GetCurrentPluginPath` at the 8 listed sites. Leave ~2247 as-is.

**Required proof:** new dummy-plugin (non-`file`) compare selftest exercising Restore/Invert Differences Selection and post-completion empty-state; assert they take effect (they silently no-op today on URL-form roots). RED on current HEAD.

**Status 2026-06-20:** complete via `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state`; compare-window selection/empty-state paths now use `GetCurrentPluginPath(...)` before `TryMakeRelative(...)`, and the debug snapshot exposes internal pane plugin paths, selection counts, item counts, and empty-state text for regression proof. Debug build, focused Commands case, and full CompareDirectories suite are green.

### CX0-6. `[HIGH]` Options hot-reload: surprise modal while hidden + silent loss of "kept" edits
*(covers findings: `external-reload-prompt-while-panel-hidden`, `stale-flag-survives-reopen-data-loss`)*

**Symptom:** (a) While the user is just browsing the compare result (Options panel hidden), an external/background settings edit can pop a modal "Reload from disk / Keep editing" prompt. (b) After "Keep editing", reopening the panel silently discards the kept edits and a *phantom* OK-time conflict prompt fires on an actually-clean panel.

**Root cause:** The Options dialog is created once at window construction and lives hidden for the window lifetime; `OnOptionsInitDialog` calls `SettingsHotReload::RegisterParticipant(dlg)` (~557) and removes it only at `WM_NCDESTROY` (~542). So `kSettingsReloadedFromDisk` → `OnOptionsSettingsReloadedFromDisk` runs even when hidden, with no `IsWindowVisible` guard (~1333-1354). Separately, "Keep editing" sets `_optionsStaleFromExternalReload = true` (~1347) but `ShowOptionsPanel(true)` unconditionally calls `LoadOptionsControlsFromSettings()` (clobbering kept edits) and never clears the stale flag, so a clean reopened panel still triggers `ResolveOptionsStaleSaveConflict` → `PromptStaleSaveConflict` on OK (~3101-3124, 3652-3726). Spec scopes this behavior to "External settings reload **while Options is open**" (~268-276).

**Fix:** (a) Register/unregister the reload participant in `ShowOptionsPanel(true/false)` (preferred), or gate `OnOptionsSettingsReloadedFromDisk` on `_optionsDlg` visibility — when hidden, silently adopt disk state and never prompt. (b) Clear `_optionsStaleFromExternalReload` at the top of `ShowOptionsPanel(true)` (a freshly loaded panel is by definition not stale). If kept edits must truly survive hide/show, snapshot and re-apply them instead of calling `LoadOptionsControlsFromSettings`.

**Required proof:** selftest fires `kSettingsReloadedFromDisk` with the panel hidden → asserts **no** modal; second selftest "keeps editing", reopens, asserts the panel is clean and OK raises **no** phantom conflict prompt.

**Status 2026-06-20:** complete via `cmd_compare_directories_options_hot_reload_visible_only_and_reopen_clean`. The Options dialog no longer registers for settings reloads during `WM_INITDIALOG`; `ShowOptionsPanel(true)` now reloads controls, clears `_optionsStaleFromExternalReload`, shows the panel, and registers the reload participant, while `ShowOptionsPanel(false)` unregisters and clears stale state before hiding. `OnOptionsSettingsReloadedFromDisk` also has a defensive hidden-panel gate that silently reloads without prompting for already-posted messages. The debug snapshot exposes `optionsReloadParticipantRegistered`, `optionsStaleFromExternalReload`, and `optionsDialogDirty`. Validation: `.\build.ps1 -ProjectName RedSalamander` green (`.build/logs/msbuild-20260620_172000_263.log`, 0 warnings / 0 errors); focused hot-reload case green; `cmd_compare_directories_options_` prefix initially had a transient tab-traversal miss that passed in isolation and passed on rerun (10 passed / 0 failed); broader `cmd_compare_directories_` family green (13 passed / 0 failed). Inventory Pester is now aligned after the 2026-06-20 Compare count and FileOperations active-phase count updates.

### CX0-7. `[HIGH]` Preferences "Compare Directories" page never renders its section headers
*(finding: `prefs-compare-headers-never-shown`)*

**Symptom:** The Compare Directories page in the global Preferences dialog shows a flat, undifferentiated list of cards with no section headings (Subdirectories / Compare / Additional / Ignore) — inconsistent with every other Preferences page. (The separate in-window Options panel is a different code path and shows its 4 headers correctly.)

**Root cause:** `LayoutDxPage` hides all headers up front (`Preferences.CompareDirectories.cpp` ~712-718: `header->SetVisible(false)`) and **never** calls `SetVisible(true)`/`SetBounds` on any header — whole-file grep confirms no `header->SetBounds`, no re-show, no `layoutHeader` equivalent. The sibling `Preferences.FileOperations.cpp` (~968-977) defines a `layoutHeader` lambda doing exactly that before each section. The header index→section mapping is also muddled (~245-249, 519-529).

**Fix:** Add a `layoutHeader` lambda mirroring `Preferences.FileOperations.cpp` (~968-977) that makes the header visible, sets its bounds, and advances `y`; invoke it before each section's cards. Fix the header index→section mapping so headers 1..4 (SUBDIRS/COMPARE/ADVANCED/IGNORE) precede their cards; decide whether the page-level header[0] is shown. Reserve `headerHeight` vertical space.

**Required proof:** PT-1 — extend the Preferences debug snapshot to report the visible Compare section-header count; assert `== 4` in `validateCompareDirectoriesPageChrome` (mirrors the Options-panel `visibleDxBodyHeaderCount == 4` check). RED on current HEAD.

**Status 2026-06-20:** complete. `LayoutDxPage` now lays out headers 1..4 for Subdirectories / Compare / Additional / Ignore, leaves the page-level header hidden, and the retained DxUi child order uses the same 1..4 mapping. `PreferencesDebugSnapshot` exposes `compareDirectoriesVisibleSectionHeaderCount`, and `cmd_preferences_dialog_compare_directories_page_uses_dxui_statics` asserts it is `4`. Debug build and the `cmd_preferences_dialog_compare_directories_` family are green.

---

## Phase P1 — Concurrency / memory / crash-class (each independent)

### CX1-1. `[MEDIUM]` Decision-cache 300MB budget is effectively unenforced under a wide pending subtree
*(covers findings: `decision-cache-budget-unbounded-under-pending-subtree`, `eviction-inner-loop-rescans-skipped-prefix`)*

**Symptom:** On large/wide remote subtree compares (the exact S3↔FS scenario the budget was added for), the decision cache grows well past its 300MB target while a scan is in progress — partially re-introducing the GiB-scale memory pressure eviction was designed to prevent.

**Root cause:** `MaybeEvictDecisionCacheLocked` treats `anyPending` as an **absolute eviction veto** (`CompareDirectoriesEngine.cpp` ~1208). During a `compareSubdirectories` scan, every directory with still-scanning descendants has `SubdirPending` → `anyPending==true`, so tens of thousands of decisions are simultaneously unevictable; the budget check runs but evicts nothing and `_decisionCacheEstimatedBytes` grows unbounded. Secondary: `FlushPendingSubdirUpdatesBudgeted` (~786-818) and `PropagateChildAggregateToAncestorsLocked` (~1643) grow decisions via `TrackDecisionCacheInsertOrUpdateLocked` but **never** call eviction, so the budget is not enforced on the aggregation path at all. Tertiary (E-1): the eviction inner loop restarts from `rbegin()` each eviction (~1186), wasting the inspection budget when the LRU tail is a long run of pinned/pending keys.

**Fix:** Pending directory decisions are cheap to recompute (`TryGetCachedDecision` miss → `RequestScanForFolder` re-queues), so do **not** treat `anyPending` as an absolute veto — allow eviction of `anyPending` entries that are neither pinned nor an ancestor of a pinned (visible) folder once the budget is materially exceeded, or cap retained pending decisions. Call `MaybeEvictDecisionCacheLocked()` after the `FlushPendingSubdirUpdatesBudgeted` batch (and consider it in the propagate path). Collapse the `rbegin()` re-scan into a **single reverse pass** gathering up to `kMaxEvictionsPerCall` candidates, then erase. Sequence **after CX0-4**.

**Required proof:** budget test with `SetDecisionCacheBudgetBytesForSelfTest` low + a wide pending subtree: assert `_decisionCacheEstimatedBytes` (ETW high-water counter) stays within a bounded multiple of budget. RED on current HEAD (grows unbounded today).

**Status 2026-06-20:** complete. `MaybeEvictDecisionCacheLocked()` now performs one reverse LRU candidate pass and no longer treats `anyPending` as an absolute veto. Scan-worker keys remain protected only until their pending-subdir aggregate (`version`, `HRESULT`, `anyPending`, `anyDifferent`) is recorded; `PropagateChildAggregateToAncestorsLocked()` can consume that aggregate after the full child decision is evicted. Aggregation paths call eviction after pending-subdir/content propagation, and invalidation/reset paths clear aggregate state with the related pending keys.

Validation: `.\build.ps1 -ProjectName RedSalamander` green with `.build/logs/msbuild-20260620_194901_078.log` (`0 warning(s), 0 error(s)`). Focused `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter decision_cache_eviction_budget_pending_wide_tree -TimeoutMultiplier 2` green (`1 passed`), archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_195237/`. Archived `compare.selftest.decision_cache.pending_budget_bytes` counter: `budget=32768`, allowed `131072`, current bytes `31350`, high-water bytes `34596`. Full Compare green (`156 total / 132 passed / 0 failed / 24 skipped`), archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_195612/`.

### CX1-2. `[MEDIUM]` Leave-scope `MessageBox` runs a nested modal pump inside the navigation callback

**Symptom:** A pending deferred Rescan (or in-flight nav) processed while the leave-scope prompt is open can change roots, re-enter compare mode, or refresh panes; when the modal returns, the code may re-navigate to a stale `previousPath` or `CancelCompareMode` on a just-restarted session, leaving panes/roots inconsistent.

**Root cause:** `SyncOtherPanePath` is invoked synchronously from `OnPanePathChanged`, the `_panePathChangedCallback` fired at the tail of `FolderWindow::SetFolderPath`. When the user navigates outside the roots while a run is pending it shows `MessageBoxCentered` (`CompareDirectoriesWindow.cpp` ~2276), spinning a nested pump while still inside the originating `SetFolderPath`. During that pump, `kCompareDirectoriesDeferredStart` and progress/decision handlers run; the recursion guard `_syncingPaths` is a single non-reentrant bool, so nested ops leave it in an arbitrary state and the window state machine (`_compareActive`/`_compareRunPending`/`_compareRunId`/deferred phase) advances under a mid-navigation stack acting on now-stale `previousPath`/`changedPane`. Single UI thread, so not memory-unsafe — a state-coherence hazard around the most safety-relevant prompt.

**Fix:** Do not show a modal from the navigation callback. Either (a) revert the pane to `previousPath` immediately and `PostMessage` a request to show the prompt once the stack unwinds (decide cancel/proceed there — this also removes the visual flash of the out-of-scope folder behind the box), or (b) add a dedicated `_leaveScopePromptActive` flag that suppresses deferred-start processing and further sync while the prompt is open.

**Required proof:** selftest/inspection that a deferred-start cannot run reentrantly while the prompt is open; pane state coherent after either choice.

**Status 2026-06-20:** complete. `SyncOtherPanePath` no longer shows `MessageBoxCentered` from the pane path-change callback. Pending-run leave-scope navigation resolves an in-scope revert target (`previousPath` or the compare root), silently reverts the pane immediately, stores the latest leave-scope request, and posts `kCompareDirectoriesLeaveScopePrompt` for after the navigation stack unwinds. The posted handler revalidates `_compareRunId`, compare state, and the reverted pane path before showing the modal; OK cancels the still-current run and then navigates to the attempted path, while Cancel leaves the reverted compare pane coherent. `CancelCompareMode`/destroy clear stale queued prompt state.

Validation: `.\build.ps1 -ProjectName RedSalamander` green with `.build/logs/msbuild-20260620_201727_700.log` (`0 warning(s), 0 error(s)`). Focused `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_compare_directories_leave_scope_prompt_defers_out_of_navigation_callback -TimeoutMultiplier 2` green (`1 passed`), archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_202105/`. The selftest drives the embedded compare pane outside scope under a debug-armed pending run and asserts the run remains active/pending, `leaveScopePromptPending` is set only after navigation returns, and the pane path is already back at the compare root.

A broad `cmd_compare_directories_` rerun was attempted after the focused proof but exceeded a 300s tool timeout in the existing `cmd_compare_directories_options_live_dx_body_interaction` path; the partial trace only showed the first three family cases passing, so that run is not counted as CX1-2 green evidence.

### CX1-3. `[MEDIUM]` Latent double-free if the window self-destructs during `CreateWindowExW`

**Symptom:** Heap corruption / double-free if the window ever fails to fully create (or self-closes during creation). Currently dormant only because `OnCreate` unconditionally returns true.

**Root cause:** Mixed ownership. `ShowCompareDirectoriesWindow` holds `std::unique_ptr<CompareDirectoriesWindow>` across `Create()` and `release()`s only after it returns true. But `CreateWindowExW` (inside `Create`) synchronously dispatches `WM_NCCREATE..WM_NCDESTROY`; `OnNcDestroy` sets `_deletePending=true` and the `WndProcThunk` `scope_exit` does `delete self` when `_dispatchDepth==0` (`CompareDirectoriesWindow.cpp` ~595-619, 2629-2636). If the window is destroyed during creation, the object self-deletes, `Create()` returns false, and the owning `unique_ptr` deletes the **same object again** (~803-808).

**Fix:** Make ownership unambiguous — either (a) `Create()` relinquishes C++ ownership up front (`new`, hand to `CreateWindowExW`, let only `OnNcDestroy` own lifetime; return false without deleting because the object already self-deleted), or (b) keep the `unique_ptr` but guarantee the object can never self-destruct before `Create()` returns (assert `_deletePending==false` after `CreateWindowExW`; disable the deferred-delete path during creation).

**Required proof:** inspection + an `OnCreate`-returns-false injection: no double-free; ASAN clean.

**Status 2026-06-20:** complete. `Create()` now marks `_createInProgress` around `CreateWindowExW`, and both `WndProcThunk` deferred deletion and `OnNcDestroy` self-delete skip deletion while that flag is set. If `WM_CREATE` fails and `WM_NCDESTROY` runs during `CreateWindowExW`, `_deletePending` is recorded but the owning `std::unique_ptr` in `ShowCompareDirectoriesWindow` remains the sole owner until `Create()` returns false. Successful creation still releases the `unique_ptr` and the existing `WM_NCDESTROY` self-owned lifetime remains unchanged.

Validation: `.\build.ps1 -ProjectName RedSalamander` green with `.build/logs/msbuild-20260620_202843_064.log` (`0 warning(s), 0 error(s)`). Focused `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_compare_directories_create_failure_does_not_double_delete -TimeoutMultiplier 2` green (`1 passed`), archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_203227/`. The selftest injects a one-shot `OnCreate` failure, asserts no live compare window remains and the main window survives, then opens and closes Compare Directories normally. ASAN was not run locally.

### CX1-4. `[LOW]` `_dxMenuBar` dereferenced after modal `ContextMenu::Show` without null re-check

**Symptom:** If the compare window / menu-bar host is destroyed while a top-level menu popup is open (window close routed through the modal loop), post-`Show` code dereferences a null `_dxMenuBar` and crashes.

**Root cause:** `OpenDxMenuBarPopup` null-checks `_dxMenuBar` before `ContextMenu::Show` (~783) but `Show` runs a nested modal loop during which `HandleDxChromeHostMessage`'s `WM_NCDESTROY` sets `_dxMenuBar = nullptr` (~882-886). After `Show` returns, `_dxMenuBar->SetSelectedIndex(std::nullopt)` (~856) has no re-check (`CompareDirectoriesWindow.Menu.cpp` ~852-859).

**Fix:** After `ContextMenu::Show` returns, re-check `if (! _hWnd || ! _dxMenuBar) return;` before touching `_dxMenuBar` (mirror the guard at the top of the function).

**Required proof:** inspection; close-during-popup does not deref null.

**Status 2026-06-20:** complete. `OpenDxMenuBarPopup` now re-checks `_hWnd` and `_dxMenuBar` immediately after `ContextMenu::Show` returns, before clearing selected index, syncing the menu bar, restoring focus, or posting the chosen command. If the modal popup loop destroyed the compare window or menu-bar host, the function returns without dereferencing `_dxMenuBar`.

Validation: `.\build.ps1 -ProjectName RedSalamander` green with `.build/logs/msbuild-20260620_204144_572.log` (`0 warning(s), 0 error(s)`). Existing focused `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons -TimeoutMultiplier 2` green (`1 passed`), archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_204528/`. A direct close-during-popup automation attempt was abandoned because the test scaffolding could strand the UI thread; it was removed and is not counted as evidence.

---

## Phase P2 — Correctness edges / UX defects (batch as convenient; each independent)

- **CX2-1 `[MEDIUM data-safety]` Trailing space/dot name-collision silently drops a same-side entry.** `NormalizeEntryNameForCompare` strips trailing spaces/dots and the trimmed name is the map key via `outEntries.emplace(...)` (`CompareDirectoriesEngine.cpp` ~2039); `std::map::emplace` does not overwrite, so two same-side names normalizing to the same key (`report` vs `report.`, `data` vs `data `) drop the second — invisible in compare and excluded from the sync manifest. **Fix:** detect `inserted==false`; fall back to the untrimmed key for the colliding entry (or flag ambiguity), and retain the original on-disk name. **Proof:** selftest with two same-side names differing only by trailing dot/space — both represented. **Status 2026-06-20:** complete via `InsertSideEntryPreservingNormalizedCollisions(...)`. Singleton entries still use the normalized key for cross-backend/Win32 parity; when a same-side normalized collision is detected, the existing trimmed-key entry is relocated to its original enumerated name and the incoming entry is inserted by its original name. `normalized_name_collision_preserves_same_side_entries` uses the dummy filesystem to create `report.` before `report` on the left and `report` on the right; it asserts the untrimmed `report` pair remains size-identical, the trailing-dot item is `OnlyInLeft`, and an out-of-scope sentinel is not represented. Debug build `.build/logs/msbuild-20260620_210758_239.log` is green; focused case archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_211136/`; full Compare archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_211403/` (`157 total / 133 passed / 0 failed / 24 skipped`).
- **CX2-2 `[LOW]` Ignored directory's own leaf not re-checked on direct navigation.** `ShouldIgnoreEntry` runs only on child entries; `ComputeDecisionForFolder`/`RequestScanForFolder` never test `relativeFolder`'s leaf/ancestors against `ignoreDirectoryPatterns` (`CompareDirectoriesEngine.cpp` ~298-320, 2530-2544), so an ignored dir reached by direct/synchronized navigation is still compared (spec ~258 says subtrees are excluded). **Fix:** return empty/identical without enumerating when any ancestor/leaf matches an enabled ignore-directory pattern. **Proof:** selftest navigating directly into an ignored-dir relative path. **Status 2026-06-20:** complete via `ShouldIgnoreRelativeFolder(...)`, called before side enumeration in `ComputeDecisionForFolder`. The helper checks every normalized relative-path component against the parsed ignore-directory patterns so an ignored leaf or ancestor short-circuits to a successful empty decision. `ignore_direct_navigation_subtree` creates differences below an ignored folder, asserts the root decision excludes the folder, and asserts direct decisions for the ignored folder and its descendant are successful and empty. Debug build `.build/logs/msbuild-20260620_211808_480.log` is green; focused case archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_212156/`; full Compare archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_212418/` (`158 total / 134 passed / 0 failed / 24 skipped`).
- **CX2-3 `[LOW]` Leave-scope Cancel with empty `previousPath` strands the pane out-of-scope.** The original review target was the pending-run leave-scope path: if navigation moved outside the compare root and `previousPath` was empty, the pane could remain out-of-scope while compare stayed active. **Fix:** fall back to `ResolveAbsolute(changedPane, {})` (session root) before queueing the deferred leave-scope prompt, or cancel compare mode if no in-scope fallback exists. **Status 2026-06-20:** complete as part of the CX1-2 leave-scope prompt reentrancy architecture. `CompareDirectoriesWindow::SyncOtherPanePath(...)` now resolves `rootPath = _session->ResolveAbsolute(changedPane, {})`, uses `previousPath.value_or(rootPath)`, revalidates the fallback with `TryMakeRelative(...)`, silently reverts the pane, and posts the prompt via `QueueLeaveScopePrompt(...)`. If the user cancels the prompt, the pane is already back in scope; if the user confirms OK, compare mode is canceled and navigation proceeds to the attempted path. Focused Commands regression `cmd_compare_directories_leave_scope_prompt_defers_out_of_navigation_callback` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_212725/`.
- **CX2-4 `[LOW]` InProgress watermark only pulses if the spinner control exists.** `FolderView::SetBackgroundWatermark` invalidates once; the old repeated-invalidation driver was `OnProgressSpinnerTimer`, gated on `_scanProgressBar` being non-null (`CompareDirectoriesWindow.Progress.cpp` ~740-805, 899-942). If the spinner control failed to create, the spec-required animated watermark rendered frozen. **Fix:** drive the watermark pulse from progress state, not from the spinner HWND. **Status 2026-06-20:** complete via the shared progress pulse timer (`kCompareProgressPulseTimerId`, `_progressPulseTimerActive`). `EnsureProgressPulseTimer()` starts the timer when either the spinner is visible or `_watermarkState==InProgress`; `OnProgressPulseTimer()` invalidates the spinner only when it exists/visible and independently calls `InvalidateCompareWatermarkPanesIfDue(...)` for the animated watermark. Debug build `.build/logs/msbuild-20260620_213339_896.log` is green; focused progress regression `cmd_compare_directories_progress_perf` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_213739/`.
- **CX2-5 `[LOW]` `ScheduleDecisionRefresh` drops a pending refresh if `SetTimer` fails.** On failure, `_decisionRefreshTimerActive` stayed false and `_decisionRefreshPending` stayed true with nothing re-arming it (`Progress.cpp` ~157-167) -> fresh decisions might never reach the panes. **Fix:** on failure, `PostMessage` a fallback or flush synchronously; log `Debug::ErrorWithLastError`. **Status 2026-06-20:** complete via `WndMsg::kCompareDirectoriesDecisionRefreshNow` and `_decisionRefreshFallbackPending`. `ScheduleDecisionRefresh()` now logs `Debug::ErrorWithLastError(...)` when `SetTimer` fails, posts a run-id-scoped refresh fallback, suppresses duplicate fallback posts while one is pending, and falls back to an immediate flush only if posting the fallback also fails. Debug build `.build/logs/msbuild-20260620_213922_473.log` is green; focused normal-path progress regression `cmd_compare_directories_progress_perf` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_214325/`.
- **CX2-6 `[LOW]` ETA shows absurd values at tiny rates.** `secondsD = remaining / _contentEtaSmoothedBytesPerSec` was gated only by `> 1.0` (`Progress.cpp` ~486-491) -> e.g. `277777:46:40` on a stalled remote. **Fix:** clamp / show "estimating..." above a sane threshold; raise the rate floor. **Status 2026-06-20:** complete by suppressing ETA display unless the smoothed rate is finite and at least 1 KiB/s and the computed ETA is finite and no more than 7 days. Values outside that window leave `_contentEtaSeconds` empty, so the banner/task card omit ETA instead of showing absurd durations. Debug build `.build/logs/msbuild-20260620_214522_317.log` is green; focused progress regression `cmd_compare_directories_progress_perf` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_214910/`.
- **CX2-7 `[LOW]` Task-card throttle leaves stale intermediate state.** `UpdateCompareTaskCard` throttled to 50ms and returned early with no trailing flush (`Progress.cpp` ~968-979), so a bursty-then-quiet compare could leave stale per-file bytes until the next message. **Fix:** set a dirty flag + one-shot trailing-flush timer. **Status 2026-06-20:** complete via `kCompareTaskCardTrailingFlushTimerId`, `_compareTaskCardTrailingFlushPending`, and a forced `UpdateCompareTaskCard(false, true)` path. Throttled intermediate updates now schedule a one-shot flush for the remaining throttle window; actual updates cancel any stale trailing timer; finish, dismiss, and teardown cancel the timer so a stale in-progress card cannot overwrite Done/dismissed state. Debug build `.build/logs/msbuild-20260620_215041_527.log` is green; focused progress/task-card regression `cmd_compare_directories_progress_perf` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_215432/`.
- **CX2-8 `[LOW]` Wheel remainder not reset when unscrollable.** The `_optionsScrollMax<=0` early-out skipped resetting `_optionsWheelRemainder` (`Options.cpp` ~1641-1660), so accumulated sub-notch deltas could leak across a non-scrollable->scrollable transition. **Fix:** reset `_optionsWheelRemainder = 0` when `_optionsScrollMax <= 0`. **Status 2026-06-20:** complete in the Win32 options host wheel handler, DxUi wheel route, and layout recomputation when `_optionsScrollMax <= 0`. Debug build `.build/logs/msbuild-20260620_215541_640.log` is green; focused options scroll regression `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_215932/`.
- **CX2-9 `[LOW]` `compareContent` persists checked-while-disabled.** When `IFileSystemIO` is unsupported the toggle is disabled but its checked state is not cleared, and `ReadOptionsControlsToSettings` reads it verbatim (`Options.cpp` ~2059-2067, 3168-3171) -> persists `compareContent=true`. Cosmetic (engine guards every use with `&& IsContentCompareSupported()`). **Fix:** AND-mask `compareContent` with support in `ReadOptionsControlsToSettings`, or uncheck on disable. **Status 2026-06-20:** complete. `LayoutOptionsControls()` now clears the content toggle before disabling it when either side lacks `IFileSystemIO`; `LoadOptionsControlsFromSettings()` mirrors persisted settings as unchecked for unsupported scopes; `ReadOptionsControlsToSettings()` AND-masks the saved `compareContent` value with `_session->IsContentCompareSupported()`. Debug build `.build/logs/msbuild-20260620_220311_534.log` is green; focused no-IO engine guard `content_no_io_disables_compareContent` passed, archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_220654/`; focused options/plugin dialog guard `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_220703/`. *(CX3-3 still removes the hidden-control source-of-truth design entirely.)*

---

## Phase P3 — Architecture & simplification (de-risks future work; largest payoff first)

CX3-1 is an architecture item but executes as **P0-foundation** because CX0-1 depends on one authoritative selection/difference contract. The remaining P3 items can wait until the P0/P1 data-safety work is stable.

### CX3-1. `[MEDIUM / P0-foundation]` Extract one `ApplyCriteriaDiffAndSelection(...)` — diff/selection logic is triplicated

The logic turning `(sizeDifferent, timeDifferent, attrsDifferent, contentDifferent)` into a `differenceMask` + `selectLeft`/`selectRight` (Size→bigger, DateTime→newer, Attributes→both, Content→both) is written verbatim **three times**: `ComputeDecisionForFolder` file branch (`CompareDirectoriesEngine.cpp` ~2843-2910), `ApplyPendingContentCompareUpdatesLocked` new-item branch (~1717-1782) and existing-item branch (~1814-1877). The copies must stay byte-identical or the post-content-compare render disagrees with the initial scan about which side to select — a subtle data-safety risk when the user then copies the "selected" side. **Fix:** extract a single free function `void ApplyCriteriaDiffAndSelection(CompareDirectoriesItemDecision&, const Settings&, bool canCompareContent, bool contentDifferent)`; call from all three sites. **Do this first** — CX0-1's manifest and content-compare both depend on a single-sourced selection contract. **Proof:** existing compare battery green; single definition asserted by grep. **Status 2026-06-20:** complete; build and CompareDirectories suite are green as recorded in the checklist.

### CX3-2. `[MEDIUM]` Delete the dead split-host DxUi layout branch + never-attached per-control members
*(covers findings: `dead-legacy-split-host-layout-path`, `options-dead-per-control-dx-hosts-and-fallback-layout`)*

The stabilized panel always uses the single `body` host: `LayoutOptionsControls` returns via the body branch whenever `_optionsDxUi && usesDxUiStatics && body.hostHwnd` (`Options.cpp` ~2493/2738), and `EnsureOptionsDx*Hosts` only ever create `body` (per-control `hostHwnd`s stay null). The ~360-line fallback (~2741-3099) runs only when `_optionsDxUi==nullptr`, and even then every `dxHeader/dxCard/dxToggle/dxEdit` argument is `nullptr`, so those sub-branches are permanently dead. The backing members (`OptionsSectionDxLabel`, `OptionsCardDxText`, `OptionsToggleDx`, `OptionsEditDx` in `Internal.h` ~462-595, flagged "Compatibility-only placeholders") are never attached.

**Fix:** Prefer folding this deletion into CX3-3 so the Options code is touched once: remove the dead per-control Dx structs/members and fallback layout while migrating to the draft model. If CX3-2 lands separately, it must be a no-behavior cleanup that keeps the existing hidden-HWND data path untouched and proves by grep that only unreachable compatibility code was removed. **Proof:** compare battery + Options-panel test green; grep confirms removal.

**Status 2026-06-20:** complete as a separate no-behavior cleanup. `OptionsSectionDxLabel`, `OptionsCardDxText`, `OptionsToggleDx`, and `OptionsEditDx` plus their never-attached `OptionsDxUiState` members were removed. The live body-host branch remains unchanged; the fallback after that branch is now legacy-HWND-only instead of accepting always-null `dxHeader`/`dxCard`/`dxToggle`/`dxEdit` pointers. The current hidden-HWND state model is intentionally left intact for CX3-3. Debug build `.build/logs/msbuild-20260620_222034_595.log` is green. Focused options guards passed: `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics` (`Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_222945/`), `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable` (`Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_222955/`), and `cmd_compare_directories_options_enter_and_escape_route_default_cancel` (`Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_223010/`). Full Compare suite passed (`158 total / 134 passed / 0 failed / 24 skipped`), archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_223229/`. A broad `cmd_compare_directories_options_` prefix attempt exceeded a 300s tool timeout before producing a summary and is not counted as evidence.

### CX3-3. `[HIGH — user-requested]` Remove the hidden-Win32 shadow control model — make DxUi + draft state the single source of truth in Options
*(finding: `options-legacy-hwnd-shadow-state-model`)*

**This is the explicit ask: "remove the Options panel that keeps a hidden Win32 control."**

**Current (broken) design:** The Options panel maintains **two complete control trees**. `EnsureOptionsControlsCreated()` creates real but **hidden** Win32 controls for every option — `BS_AUTOCHECKBOX` buttons (`makeToggle`), `Edit`s (`makeFramedEdit`), DxHost statics (`_optionsUi.*`) — and a parallel DxUi tree (`_optionsDxUi->body.*`). The **hidden controls are the canonical data model**: `ReadOptionsControlsToSettings()` (the function that persists user choices) reads exclusively from `_optionsUi.*.toggle` via `GetTwoStateToggleState` (`Options.cpp` ~3160-3178), never from the DxUi toggles. A click on a DxUi toggle is round-tripped to the hidden legacy control via a synthetic `BN_CLICKED` (`syntheticHiddenToggle`, `OnOptionsCommand` ~1384), then re-mirrored by `SyncOptionsDxToggles()`. State lives in hidden HWNDs; rendering happens in DxUi; hand-written `Sync*` mirrors keep them aligned. Every option is wired in **three** places and can desync silently (enable-state, content-compare-unavailable — see CX2-9). This is the single largest source of risk and bloat in the 3745-line Options file.

**Proof it is viable:** `Preferences.CompareDirectories.cpp` already implements a DxUi-only compare-settings panel with **no** hidden-control model. Mirror its state flow: working settings/draft first, controls second.

**Migration runbook (ordered; do after CX0-6 so hot-reload behavior is known):**
1. **Inventory the contract.** Enumerate every option the panel owns (`compareSize`, `compareDateTime`, `compareAttributes`, `compareContent`, `compareSubdirectories`, `compareSubdirectoryAttributes`, `selectSubdirsOnlyInOnePane`, `keepIdenticalItems`, `ignoreFiles`+pattern, `ignoreDirectories`+pattern). For each, note: persisted field, enable/disable rule, dirty/reload behavior, and interlock (e.g. `compareContent` disabled when `!IsContentCompareSupported()`).
2. **Introduce explicit draft state.** Add a small `CompareOptionsDraft` member or reuse `Common::Settings::CompareDirectoriesSettings` as `_optionsDraft`. `ShowOptionsPanel(true)` loads the draft from effective settings; user edits update the draft; `SaveOptionsControlsToSettings()` persists the draft. `IsOptionsDialogDirty()` compares draft to effective settings.
3. **Bind DxUi directly to the draft.** Toggle/edit callbacks update `_optionsDraft` and then call one `SyncOptionsDxFromDraft()`/`ApplyOptionsDraftInterlocks()` function. `LoadOptionsControlsFromSettings()` becomes `LoadOptionsDraftFromSettings()` + `SyncOptionsDxFromDraft()`. `ReadOptionsControlsToSettings()` either disappears or returns the draft.
4. **Delete the hidden control round-trip.** Remove hidden toggle/edit/static HWND creation, `syntheticHiddenToggle`, `BN_CLICKED` re-dispatch, `GetTwoStateToggleState(_optionsUi...)` reads, `SetTwoStateToggleState(_optionsUi...)` writes, and legacy-to-Dx `Sync*` functions. Keep only the option dialog/host HWNDs DxUi needs.
5. **Fold CX3-2 cleanup here.** Remove never-attached per-control Dx host structs/members and unreachable split-host layout branches while deleting the hidden model. Do not keep compatibility placeholders unless a live call site remains.
6. **Re-home enable/disable + interlocks** (CX2-9) onto the draft + DxUi controls: `compareContent` disabled+unchecked when unsupported; ignore-pattern edits enabled by their toggle; `keepIdenticalItems=false` forces `showIdenticalItems=false` if that field is surfaced later.
7. **Dirty-tracking & reload.** Re-point CX0-6 hot-reload/stale logic at the draft. Hidden panel reloads silently into the draft; visible dirty panel prompts; "Keep editing" keeps the draft and stale flag only until save/reload/hide semantics say otherwise.
8. **Focus traversal & accessibility.** Verify forward/reverse keyboard traversal at small window sizes flows through DxUi controls only. Update `Commands.SelfTest.CompareOptions.cpp` to read debug state from DxUi/draft rather than `_optionsUi`, and keep UIA TogglePattern/ValuePattern coverage.

**Required proof:** Options-panel selftest (`Commands.SelfTest.CompareOptions.cpp`) green reading state from DxUi/draft; round-trip test (set each toggle/edit via DxUi → OK → reopen → values preserved); hot-reload tests from CX0-6 still green; grep confirms `_optionsUi.*` toggle/edit/static owners, `Sync*FromLegacy`/legacy-to-Dx mirrors, `syntheticHiddenToggle`, and hidden `GetTwoStateToggleState` persistence reads are gone. Run the full compare + commands batteries.

**Status 2026-06-20:** complete. The Options body no longer creates hidden native statics, checkboxes, or edits as a shadow state model; `_optionsUi` owns only the scroll/container host, and live DxUi toggles/edits update `_optionsDraft` (`Common::Settings::CompareDirectoriesSettings`) directly. `LoadOptionsControlsFromSettings()` initializes the draft from effective settings, `ApplyOptionsDraftInterlocks()` enforces content-compare support and `keepIdenticalItems`/`showIdenticalItems`, `ReadOptionsControlsToSettings()` returns the draft copy for save, OK persists it, and Cancel discards it. The legacy `OptionsToggleCard` / `OptionsIgnoreCard` owners, `SetTwoStateToggleState`, `GetTwoStateToggleState`, `syntheticHiddenToggle`, hidden Button/Edit creation, legacy-to-Dx state reads, `_optionsFrameStyle`, and `ThemedInputFrames` are gone. Same-process DxUi UIA ValuePattern tests now use message-pumped helpers so the UI thread can service retained-control providers while tests read/write values.

Validation: Debug `RedSalamander` build green at `.build/logs/msbuild-20260620_230155_024.log` (`0 warning(s), 0 error(s)`). Full `cmd_compare_directories_options_` Commands family passed (`10 passed / 0 failed`), archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_231049/`. Full Compare suite passed (`158 total / 134 passed / 0 failed / 24 skipped`), archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_230833/`. Focused earlier runs also covered pointer toggles, scroll stability, access keys, tab traversal, hot reload, and live DxUi ValuePattern round-trips after the message-pumped helper update.

### CX3-4. `[MEDIUM]` Decompose the `CompareDirectoriesWindow` god-class

`CompareDirectoriesWindow` declares the entire window — menu (Dx+legacy), banner (Dx+legacy), options (after CX3-3: DxUi only), progress/ETA/spinner/watermark/task-card, sync navigation, selection, details/metadata providers, ~25 `ENABLE_TESTS` accessors — ~150 members / ~120 methods across 6 `friend`-coupled `.cpp` files sharing a flat mutable flag set (`_compareStarted`/`_compareActive`/`_compareRunPending`/`_compareRunSawScanProgress`/`_bannerRescanIsCancel`/`_syncingPaths`/`_deferredCompareStart*`). **Fix:** extract cohesive sub-controller state owned by the window — `ChromeController` (menu+banner), `OptionsPanelController` (Options host/body/draft/theme state), `ProgressController` (`BannerProgressState`+ETA+spinner/watermark/task-card state). Challenge the original "move all methods and friend shims in one pass" framing: the Win32 callback shims should stay routed through the owning window until each message boundary can be split and tested separately. This slice removes the flat shared-state surface without rewriting message dispatch. Depends on CX3-3 (Options) and CX3-2 landing. **Proof:** compare battery green after extraction; focused chrome/options/progress command tests green; grep shows the old flat state names gone and grouped state references routed through `_chrome`, `_optionsPanel`, and `_progress`. `B-1`, `MD-2`, `API-1` (CX3-5) fold in naturally.

**Status 2026-06-20:** complete as a scoped state-controller extraction. `CompareDirectoriesWindow.Internal.h` now owns explicit `ChromeController`, `OptionsPanelController`, and `ProgressController` buckets with deleted copy/move operations. Chrome/menu/banner state is routed through `_chrome`; Options dialog/body/draft/theme/scroll state is routed through `_optionsPanel`; progress banner, ETA, spinner, watermark, task-card id/result, and progress-host state is routed through `_progress`. The previously flat `_compareRunSawScanProgress`, `_bannerRescanIsCancel`, `_compareTaskId`, `_compareRunResultHr`, `_watermarkState`, `_options*`, `_dxMenuBar*`, `_dxBanner*`, and `_scanProgress*` families are gone. The existing Win32 callback friend shims remain intentionally, because they are message-dispatch adapters and not the state ownership defect closed here.

Validation: Debug `RedSalamander` build green at `.build/logs/msbuild-20260620_231739_525.log` (`0 warning(s), 0 error(s)`). Focused chrome/menu/banner guard `cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_232056/`. Full `cmd_compare_directories_options_` Commands family passed (`10 passed / 0 failed`), archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_232125/`. Focused progress guard `cmd_compare_directories_progress_perf` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_232133/`. Full Compare suite passed (`158 total / 134 passed / 0 failed / 24 skipped`), archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_232352/`. An initial parallel run of these suites hit the documented self-test mutex with exit 3 for the overlapping jobs; the serial reruns above are the green evidence.

### CX3-5. `[LOW]` Dedup & API hygiene (fold into CX3-4)

- **B-1** `EnsureOptionsDxButtonHosts` runs twice at creation (WM_INITDIALOG path + `CreateChildWindows` ~1678), tearing down and rebuilding just-created hosts (`attachButtonHost` begins with `slot.Detach()`). **Fix:** drop the second call, or no-op when already attached.
- **MD-2** `CompareDirectoriesWindow.Menu.cpp` (~44-244) reimplements `SplitMenuText`/`StripMenuMnemonicMarkers`/`FindMenuMnemonic`/`TryGetMenuItemPresentationText`/`ConvertHMenuToDxFlyoutItems`/`BuildDxMenuBarItems` byte-for-byte vs the shared `Common/DxUi/DxUiNativeMenuInterop.h` (`ConvertNativeHMenuToFlyoutItems`/`BuildNativeMenuBarItems`, already used by the viewers). **Fix:** delete the local copies; call the shared functions with a default `NativeMenuFlyoutOptions{}`.
- **API-1** `GetOrComputeDecision` is public but has zero production callers (UI uses `TryGetCachedDecision` only) and can block on `_contentCompareQueueNotFullCv` backpressure — it invites a future UI-thread stall. **Fix:** move behind the `ENABLE_TESTS` hook (like `SetDecisionCacheBudgetBytesForSelfTest`) or rename to expose the blocking nature.

**Status 2026-06-20:** complete as a separate low-risk cleanup before CX3-4. `CreateChildWindows()` no longer calls `EnsureOptionsDxButtonHosts()` after `WM_INITDIALOG` already attached footer hosts; `CompareDirectoriesWindow.Menu.cpp` deletes the local native-menu text/conversion helpers and calls `RedSalamander::DxUi::BuildNativeMenuBarItems(...)` / `ConvertNativeHMenuToFlyoutItems(...)`; `CompareDirectoriesSession::GetOrComputeDecision(...)` is declared/defined only under `ENABLE_TESTS`, leaving production UI on the non-blocking `TryGetCachedDecision(...)` path. Debug build `.build/logs/msbuild-20260620_221106_949.log` is green. Focused menu/banner guard `cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_221500/`; focused options DxUi host guard `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics` passed, archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-20_221517/`; full Compare suite passed (`158 total / 134 passed / 0 failed / 24 skipped`), archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_221733/`.

---

## Phase P4 — Tests (each pins a fix above; write alongside its phase, run full suite at closeout)

For compare engine tests, add every new case to `kCompareCaseNames`, verify it appears in `CompareDirectoriesSelfTest::ListCases(...)`, and invoke it from the relevant included case file. For commands/preferences tests, add the case to the appropriate `Run*Cases(...)` dispatcher and verify a case-filtered run executes it. A case name that lists but never executes, or executes but cannot be selected by filter, does not close a finding.

**P0 pinning tests:**
- **T-8** `Crosscut_SyncManifestCopiesOnlyDifferences` + `Crosscut_SyncMoveKeepsIdenticalSource` — pins CX0-1 (disputed-as-missing today; mandatory once the manifest lands). Multi-depth tree; assert exact manifest operation set `(source,dest,relative,kind,flags)`, nested destination paths, one-sided `WholeSubtree`, no recursive flags on mixed directories, and Move keeps identical source. Include a pending-subtree/content blocker case. Byte-equality + out-of-tree sentinel.
- **L-1b** `Crosscut_FailedEnumerationNotCachedRetries` — pins CX0-3. `forcedHr=E_ACCESSDENIED`→`S_OK`; recompute without version bump.
- **CX0-5 test** non-`file` dummy-plugin compare exercising Restore/Invert selection + post-completion empty-state — pins CX0-5.
- **CX0-6 tests** reload-while-hidden (no modal) + keep-editing-reopen (clean, no phantom prompt) — pins CX0-6.
- **PT-1** Preferences visible Compare header count `== 4` — pins CX0-7. Add a real debug-snapshot field for visible Compare section headers first; do not leave the current `true /* Phase 8: removed field */` placeholders as the proof.

**P1 pinning tests:**
- **T-3** `Crosscut_ContentWorkerVsInvalidateStress` — `compareContent=true` over many equal-size byte-different files + readers + repeated `Invalidate`/`SetBackgroundWorkEnabled` toggles + low cache budget. Asserts: no crash (pins CX0-4 dangling iterator), no null decision, no stale-equal, and `_decisionCacheEstimatedBytes` bounded (pins CX1-1). Also drives the CX0-2 stale-writeback assertion. The single most important new concurrency test.
- **T-4** `Crosscut_SetRootsResetsAndBumpsVersion` — warm rootsA + content compares, `SetRoots(rootsB)`; assert versions bumped, old keys stale/none, queues/in-flight drained, new decisions correct. (`SetRoots` is currently called by **no** test.)

**Standing data-safety tests (write with their nearest phase; they guard against future regressions even though current code is correct on these paths):**
- **T-1** `Crosscut_ContentEqualSizeEqualMtimeDiffers` — force identical sizes + identical `SetFileLastWriteTime` both sides, `compareContent=true` only; assert `isDifferent` + Content bit + `selectLeft&&selectRight`. Mirror with identical bytes+mtime asserting NOT different (proves the byte compare ran). **The regression net for the most likely future data-loss optimization** (a size+mtime dedup fast path).
- **T-2** `Crosscut_UnknownSizeStreamingCompare` — `IFileReader` wrapper whose `GetSize` returns `E_FAIL` (forces `sizeKnown=false`); cases: both empty, equal nonzero, left-prefix-of-right, right-prefix-of-left, mid-stream mismatch, all with short reads. Guards the unknown-size EOF logic (`CompareDirectoriesEngine.cpp` ~2243-2262) for FTP/S3/archive backends.
- **T-5** `Crosscut_ContentCacheHitSkipsIo` — counting `IFileSystemIO`; second compare of an unchanged pair opens no files; assert `ContentCompareKeyEq` distinguishes mtime-only and size-only keys (hash/eq contract, ~408-422).
- **T-6** `Crosscut_SelectSubdirsOnlyInOnePane` — assert one-sided-directory selection differs between `selectSubdirsOnlyInOnePane=false` vs `true`; reconcile with the existing `unique` case so both match the spec.
- **T-7** `Crosscut_MissingSideEmptyEnumeration` — build a relative subfolder only on the left, call the right wrapper's `ReadDirectoryInfo` on the missing-side relative folder; assert `SUCCEEDED` + zero entries (not a failure). Pins the empty-state contract.
- **CX2-1 test** two same-side names differing only by trailing dot/space — both represented. **Status 2026-06-20:** complete via `normalized_name_collision_preserves_same_side_entries`.

**Status 2026-06-20:** complete. CX4 closeout audited listability/filterability/dispatch and added the missing standing cases:

- **T-1** `content_equal_size_equal_mtime_differs` — equal size + equal mtime still byte-compares; different bytes select both with `Content`, equal bytes remain identical.
- **T-2** `content_unknown_size_streaming_compare` — forced `GetSize` failure with short reads covers empty, equal nonzero, left-prefix, right-prefix, and mid-stream mismatch.
- **T-3** `invalidate_drops_stale_inflight_scan_writeback` + `decision_cache_eviction_budget_pending_wide_tree` — stale in-flight scan writeback, pending update safety, and low-budget pending eviction guards.
- **T-4** `setRoots_resets_and_bumps_version` — `SetRoots` bumps version, clears old cached decisions, and computes the new roots.
- **T-5** `content_cache_hit_skips_io` — counting reader proves settings-only recompute of an unchanged content pair reuses `_contentCompareCache` and performs no extra reads. This exposed and fixed an overbroad settings reset that cleared completed content cache entries.
- **T-6** `select_subdirs_only_in_one_pane` — one-sided directory selection follows `selectSubdirsOnlyInOnePane=false/true`.
- **T-7** `missing_side_empty_enumeration` — compare filesystem wrapper returns `SUCCEEDED` + zero entries for the missing side of a one-sided folder.
- **T-8** `sync_manifest_nested_differences_only`, `sync_manifest_move_preserves_identical_children`, `sync_manifest_not_ready_without_cached_decision`, `sync_manifest_pending_blocks_and_schedules`.
- **L-1b** `failed_enumeration_retries_without_version_bump`.
- **PT-1** `cmd_preferences_dialog_compare_directories_page_uses_dxui_statics` with `compareDirectoriesVisibleSectionHeaderCount == 4`.

Validation: Debug `RedSalamander` build green at `.build/logs/msbuild-20260620_233426_995.log` (`0 warning(s), 0 error(s)`). New focused cases passed in archives `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_233323/` (`content_equal_size_equal_mtime_differs`), `2026-06-20_233335/` (`content_unknown_size_streaming_compare`), `2026-06-20_233803/` (`content_cache_hit_skips_io` after the cache-preservation fix; pre-fix failure archived at `2026-06-20_233340/`), `2026-06-20_233813/` (`setRoots_resets_and_bumps_version`), `2026-06-20_233821/` (`select_subdirs_only_in_one_pane`), and `2026-06-20_233826/` (`missing_side_empty_enumeration`). `RedSalamander.exe --selftest-list-cases` lists all six new case names. Full Compare suite passed (`164 total / 140 passed / 0 failed / 24 skipped`), archive `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-20_234056/`; the full-suite closeout rerun also archived a green CompareDirectories leg at `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-21_002337/`. Focused command dispatcher proof passed: `cmd_compare_directories_` (`15 passed / 0 failed`), archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-21_001914/`, and `cmd_preferences_dialog_compare_directories_page_uses_dxui_statics` (`1 passed / 0 failed`), archive `Specs/TestRuns/7d3a1247382a/Commands/2026-06-21_001936/`.

Closeout validation updates on 2026-06-24:

- Latest clean test-enabled Debug build is green at `.build/logs/msbuild-20260624_142759_244.log` / `.build/logs/msbuild-20260624_142801_196.log` (`0 warning(s), 0 error(s)`).
- CompareDirectories remains green from the last full compare evidence recorded above; no June 24 change reopened Compare correctness. `Specs/Core/Core_CompareDirectories.md` has been updated with the durable contracts for sync manifest behavior, targeted invalidation/cache coherence, bounded ignore-pattern matching, Options draft-state ownership, progress/watermark behavior, and coverage expectations.
- The previous menu mouse-open blocker is fixed and focused green: `.build/logs/run-commands-menu-mouse-open-20260624_143208_163.log` (`1 passed / 0 failed`).
- Latest broad Commands run completed and failed overall: `.build/logs/run-commands-closeout-20260624_143239.out.log`, archived results `Specs/TestRuns/7d3a1247382a/Commands/2026-06-24_150328/commands_results.json` (`751 passed / 11 failed / 2 skipped`).
- The exact 11 failing Commands cases passed when filtered by their failure clusters:
  - Credential prompt cluster: `.build/logs/run-commands-failing-credential-20260624_150453_924.log` (`2 passed / 0 failed`).
  - Preferences cluster: `.build/logs/run-commands-failing-preferences-20260624_150512_968.log` (`7 passed / 0 failed`).
  - Menu cluster: `.build/logs/run-commands-failing-menu-20260624_150555_203.log` (`2 passed / 0 failed`).
- A wider plugin/connection-manager/credential replay still failed (`41 passed / 5 failed`) at `.build/logs/run-commands-plugin-connection-credential-20260624_150743_323.log`. The most useful diagnostic signal is the localized-default-name case observing `rce` instead of `New connection`, which points to stale keyboard/focus contamination or an equivalent value-entry ordering defect.
- FileOperations `Phase7_AutoConcurrencyHints` was cleared in focused rerun `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-21_205655` (`19 passed / 0 failed / 0 skipped`), and Monitor ETW latency was cleared in `Specs/TestRuns/7d3a1247382a/Monitor/2026-06-21_210002`.

This is Done evidence for the CompareDirectories remediation under the explicit split policy. The Commands order/input instability must be fixed under `Specs/Plans/WIP/Operation_CommandsSelfTestInputIsolation_2026-06-24.md` before treating the broader Commands gate as green.

---

## Verified correct — do NOT re-open without new evidence

The review's refutation pass rejected 12 candidate findings, and the following were examined and found sound (do not rework as part of Crosscut):
- **Version/coherence skeleton** — `_version` increments on `SetRoots`/`SetSettings`/`Invalidate`; cached decisions tagged with version; in-flight `(version, cancelToken)` stamps discard stale global-reset scan/content results. Scoped `InvalidateForAbsolutePath` preserves unrelated cache entries and relies on targeted queue/in-flight pruning plus live-stamp writeback checks (CX0-2).
- **Size-known and unknown-size content read loops** (`CompareFileContent`) — short-read handling, 256KB streaming, min-available compare, the post-loop extra-byte EOF check (reading 1 trailing byte each side to catch a lying `GetSize`), `sizeKnown` size/zero shortcuts, and forced-unknown-size EOF behavior are covered by `content short reads` and `content_unknown_size_streaming_compare`.
- **Content equality is always byte-verified** — there is no metadata-only "assume equal" shortcut; `_contentCompareCache` only ever stores results produced by an actual byte comparison. The size+mtime cache key is a narrow staleness window (in-place edit preserving size *and* mtime), not a metadata-only equality decision.
- **Name matching** — the ordered `std::map` with `WStringViewNoCaseLess` (`CompareStringOrdinal` case-insensitive) correctly avoids the hash/equality contract violation an `unordered_map` would create. The **only** matching defect is the trailing dot/space collision drop (CX2-1).
- **Posted scan/content progress coalescing** and the `WM_NCDESTROY` payload draining were not found defective. This does not cover the separate task-card trailing-flush issue in CX2-7.

If a future change appears to contradict one of these, attach new source evidence before reopening.

---

## Concrete execution order

1. **CX3-1** (extract `ApplyCriteriaDiffAndSelection`) first — single-sources the selection contract the manifest and content-compare depend on. Compare battery stays green.
2. Add RED tests for **CX0-2, CX0-3, CX0-5** on current HEAD; confirm they fail for the dangerous reason. Fix each (cheapest correctness/crash fixes first to stabilize the ground CX0-1 stands on): **CX0-4** (snapshot), **CX0-3** (SUCCEEDED gate), **CX0-2** (coherent invalidation/queue policy), **CX0-5** (plugin path).
3. Add the **file-operation resolved-item mode** with focused unit/selftest coverage before wiring Compare Directories to it. Prove nested destination paths are preserved and generic `CopyItems`/`MoveItems(sourcePaths, destinationFolder)` semantics are unchanged.
4. **CX0-1** (session sync manifest + command interception) with **T-8** RED→GREEN. This is the largest slice; build `TryBuildSyncManifest(...)` in the session so T-8 drives it directly. Include manifest-only Move deletion after destination success, one-sided `WholeSubtree`, not-ready blockers, confirmation, and manifest perf counters. Depends on CX0-2, the resolved-item file-op mode, and CX3-1.
5. **CX0-6** (Options hot-reload) and **CX0-7** (Preferences headers + real debug snapshot field) with their tests — self-contained UI fixes.
6. **CX1-1** (eviction budget, after CX0-4), then **CX1-2/CX1-3/CX1-4** (reentrancy + crash-class guards).
7. **CX2-1** (name collision — data-safety) first in P2, then batch the remaining P2 UX fixes.
8. **CX3-3** (draft-model Options migration, folding CX3-2 unless already done) → **CX3-4** (god-class decomposition) → **CX3-5** (dedup). Each keeps the compare + Options + DxUi-menu batteries green.
9. Write standing data-safety tests (T-1, T-2, T-5, T-6, T-7) with their nearest phase; T-3/T-4 in P1.
10. **Closeout:** verify all new cases are listable/filterable/executed; run `.\Tools\Run-AllTests.ps1 -Suite Full` (expect exit 3 only on cross-worktree mutex contention); archive before/after perf+safety evidence under `Specs/TestRuns/<machine>/Compare/<timestamp>/`; update `Specs/Core/Core_CompareDirectories.md` for clarified contracts (resolved-item file-op mode, one-sided recursive exception, not-ready sync blockers, Options draft model); remove the TODO block at the top of that spec for items now done; move this plan to `Specs/Plans/Done/`.

## Out of scope / non-issues

- The size+mtime content-cache key is a deliberate design (attributes excluded to avoid spurious misses); only an *in-place edit preserving both size and mtime* could serve a stale "equal", and content is always byte-verified at least once per `(key)` per run. Not reopened; T-1 stands as the regression net should anyone add a true metadata-only fast path.
- The cross-FS bridge's accidental diff-filtering is **not** treated as "correct" — CX0-1's explicit manifest replaces both transport paths so the divergence disappears.

---

*Provenance: 14-lens adversarial multi-agent review of the CompareDirectories subsystem, branch `claude/elated-sinoussi-a39160`, 2026-06-15 (57 candidates → 43 confirmed / 2 disputed / 12 rejected; CX0-1 and CX0-2 hand-traced). Anchors are HEAD-relative — re-grep before editing.*
