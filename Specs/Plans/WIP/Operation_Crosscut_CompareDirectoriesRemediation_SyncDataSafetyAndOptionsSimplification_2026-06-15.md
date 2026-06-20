# Operation Crosscut — CompareDirectories Remediation: Sync Data-Safety, Cache Coherence & Options Simplification

**Status:** WIP
**Date:** 2026-06-15
**Author:** Deep multi-agent code review of the CompareDirectories subsystem (branch `claude/elated-sinoussi-a39160`). 14 specialized review lenses (engine concurrency / cache-memory / content-compare / decision-correctness; window sync-copy / lifecycle / navigation / progress / options / menu+prefs; architecture; spec-conformance; test-coverage) → per-finding adversarial refutation pass (2 independent skeptics for critical/high, 1 for medium/low) → synthesis. **57 candidate findings → 43 confirmed, 2 disputed, 12 rejected.** The two highest-stakes data-safety claims (CX0-1 sync manifest, CX0-2 invalidate version bump) were re-traced by hand in source.
**Builds on / does NOT re-open:** `Specs/Plans/Done/CompareDirectories_MemoryAndHangs.md` and `Specs/Plans/Done/Core_CompareDirectoriesReview.md`. Those addressed memory/hang issues and the prior review; the version/coherence skeleton, the LRU+budget eviction structure, the `(version, cancelToken)` in-flight stamp model, the throttled-notification design, and the size-known content read loop are sound and are NOT reworked here except where a specific defect is listed. Crosscut fixes the defects those passes' green self-tests did not catch.
**Normative spec:** `Specs/Core/Core_CompareDirectories.md` (every "MUST" referenced below is from this document).

**Scope (anchors are worktree `elated-sinoussi-a39160` @ current HEAD — line numbers are HEAD-relative; re-grep every `~line` before editing):**
- `RedSalamander/CompareDirectoriesEngine.cpp` / `.h` — session/engine: scan + content workers, decision cache, version/coherence, compare-scoped `IFileSystem` wrapper.
- `RedSalamander/CompareDirectoriesWindow.cpp` — lifecycle, sync navigation, selection, directory sync-copy/move.
- `RedSalamander/CompareDirectoriesWindow.Internal.h` — private window class declaration.
- `RedSalamander/CompareDirectoriesWindow.Options.cpp` — options panel (hidden-HWND + DxUi dual model, settings reload).
- `RedSalamander/CompareDirectoriesWindow.Progress.cpp` — progress UI, ETA, watermark, File Operations task card.
- `RedSalamander/CompareDirectoriesWindow.Menu.cpp` — themed menu bar.
- `RedSalamander/Preferences.CompareDirectories.cpp` / `.h` — global Preferences page.
- `RedSalamander/FolderWindow.FileOperations.cpp` + `.State.cpp` — copy/move submission path that compare-mode sync must bypass with a flat list.
- `RedSalamander/SelfTest/CompareDirectories/*` and `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp` / `Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp` — new self-tests.

---

## Why this plan exists

The CompareDirectories subsystem is **not currently safe to represent as a "sync" feature.** The review found three independent silent-data-loss paths, one confirmed heap-corruption race, and a complete failure of the headline cross-plugin scenario (compare a local folder against S3/FTP). The most damning is **CX0-1**: the spec's entire "Directory copy/move behavior (sync)" section — build a diff manifest, copy only the different items — is **unimplemented**. Copy/Move-to-other-pane delegates straight to the generic recursive file-op, so a "sync" copies the *whole* subtree including identical files, and **Move silently deletes identical files from the source.**

The architecture amplifies the risk: the Options panel keeps a *full hidden Win32 control set as its load-bearing state model* behind a DxUi mirror (**CX3-3** — the user-requested removal), and `CompareDirectoriesWindow` is a ~13.7k-line god-class. Every option must be wired in three places, which is precisely how the UI defects below accumulated.

**Standing rule for every Crosscut slice:** a slice is not done because a self-test is green. The current suite is green *because the risky paths are unexercised* (no test for diff-manifest construction, equal-size+equal-mtime+different-content, unknown-size streaming, `SetRoots` reset, or content-worker-vs-`Invalidate` races). For each P0/P1 slice the **RED-on-current-HEAD demonstration is part of the required proof** — the test must fail for the dangerous reason before the fix and pass after.

---

## Implementation Tracking Checklist (update first, before editing code in a slice)

Use `[ ]` not started, `[~]` in progress, `[x]` complete, `[blocked]` needs a product decision.

| State | Slice | Phase | Implementation unit | Required proof before `[x]` |
|-------|-------|-------|---------------------|-----------------------------|
| [ ] | CX0-1 | P0 | Window-level **diff manifest** for dir Copy/Move sync; flat non-recursive submission; Move = copy-verify-delete manifest-only; sync confirmation/accurate count | T-8 manifest pair-set test RED→GREEN; identical files NOT copied; Move does NOT delete identical source files; out-of-tree sentinel intact. |
| [ ] | CX0-2 | P0 | `InvalidateForRelativePathLocked` bumps `_version` + drops affected `_scanScheduledKeys`/`_scanInFlightKeys` (or per-key epoch at writeback) | T-3 contention test: stale in-flight writeback dropped after invalidate; no stale decision served post-op. RED on current HEAD. |
| [ ] | CX0-3 | P0 | Gate all 3 decision-cache inserts on `SUCCEEDED(decision->hr)`; failed decisions returned but not stored | L-1b retry test: access-denied→S_OK recomputes without version bump. RED on current HEAD. |
| [ ] | CX0-4 | P0 | Snapshot keys in `FlushPendingContentCompareUpdatesBudgeted` before apply (kill dangling-iterator UB) | T-3 stress (budget low + pending content + Invalidate) no crash/ASAN; code review confirms no live iterator across `ApplyPending…`. |
| [ ] | CX0-5 | P0 | `GetCurrentPath`→`GetCurrentPluginPath` at the 8 TryMakeRelative-fed sites | New dummy-plugin (non-`file`) compare selftest: Restore/Invert selection + post-completion empty-state take effect. RED on current HEAD. |
| [ ] | CX0-6 | P0 | Options hot-reload: register participant in `ShowOptionsPanel`; clear `_optionsStaleFromExternalReload` on reopen | Selftests: reload while panel hidden → no modal; "keep editing"+reopen → clean panel, OK raises no phantom conflict. |
| [ ] | CX0-7 | P0 | Preferences Compare page renders its 4 section headers (`layoutHeader` lambda + fixed index→section mapping) | PT-1: Preferences debug snapshot reports visible Compare header count == 4. RED on current HEAD. |
| [ ] | CX1-1 | P1 | Decision-cache budget enforcement: evict non-pinned/non-ancestor `anyPending`; call eviction on aggregation paths; single-pass eviction scan | Budget test (`SetDecisionCacheBudgetBytesForSelfTest` low + wide pending subtree): `_decisionCacheEstimatedBytes` stays within bounded multiple of budget. RED on current HEAD. |
| [ ] | CX1-2 | P1 | Leave-scope prompt deferred out of the navigation callback (revert + PostMessage, or `_leaveScopePromptActive` guard) | Selftest/inspection: deferred-start cannot run reentrantly while the prompt is open; pane state coherent after either choice. |
| [ ] | CX1-3 | P1 | Unambiguous window ownership (no unique_ptr + self-delete double-free on create failure) | Inspection + `OnCreate`-returns-false injection test: no double-free; ASAN clean. |
| [ ] | CX1-4 | P1 | `_dxMenuBar` null re-check after modal `ContextMenu::Show` | Inspection: guard present; close-during-popup does not deref null. |
| [ ] | CX2-1 | P2 | Trailing space/dot name-collision no longer silently drops a same-side entry | Selftest: two same-side names differing only by trailing dot/space both represented. |
| [ ] | CX2-2 | P2 | Ignore patterns applied to the folder's own leaf/ancestors on direct navigation | Selftest: navigate directly into an ignored-dir relative path → empty/identical, not compared. |
| [ ] | CX2-3 | P2 | Leave-scope Cancel with empty `previousPath` falls back to root (or proceeds + cancels) | Selftest/inspection: pane never stranded out-of-scope with compare active. |
| [ ] | CX2-4 | P2 | Watermark pulse driven by its own state, not the spinner control | Inspection: animated watermark survives `_scanProgressBar==null`. |
| [ ] | CX2-5 | P2 | `ScheduleDecisionRefresh` has a fallback when `SetTimer` fails | Inspection: pending refresh never stranded; logs `ErrorWithLastError`. |
| [ ] | CX2-6 | P2 | ETA clamped / "estimating…" above sane threshold | Inspection/visual: near-stalled rate shows no absurd ETA. |
| [ ] | CX2-7 | P2 | Task-card trailing flush after throttle window | Inspection: last intermediate state applied during quiet period. |
| [ ] | CX2-8 | P2 | Reset `_optionsWheelRemainder` when not scrollable | Inspection: no jump on non-scrollable→scrollable transition. |
| [ ] | CX2-9 | P2 | `compareContent` persisted value AND-masked with support (or unchecked on disable) | Inspection: disabled toggle does not persist `compareContent=true`. |
| [ ] | CX3-1 | P3 | Extract one `ApplyCriteriaDiffAndSelection(...)`; call from all 3 engine sites | Existing compare battery green; single definition asserted by grep. **Do before CX0-1 / CX3-3.** |
| [ ] | CX3-2 | P3 | Delete dead split-host DxUi layout branch + never-attached per-control members | Compare battery + Options-panel test green; grep confirms removal. |
| [ ] | CX3-3 | P3 | **Remove hidden-Win32 shadow control model — DxUi is the single source of truth in Options** | Options-panel selftest green reading from DxUi; grep confirms `_optionsUi.*` toggle/edit/static + Sync*FromLegacy deleted. |
| [ ] | CX3-4 | P3 | Decompose god-class into ChromeController / OptionsPanelController / ProgressController | Compare battery green after extraction; reduced friend coupling. |
| [ ] | CX3-5 | P3 | Drop redundant button-host attach (B-1); reuse shared `DxUiNativeMenuInterop` (MD-2); fence `GetOrComputeDecision` behind `ENABLE_TESTS` (API-1) | Compare + DxUi menu batteries green; grep confirms dedup. |
| [ ] | CX4 | P4 | All pinning + standing self-tests (T-1..T-8, L-1b, PT-1) registered in the compare selftest family | Each test FAILS on a deliberately reverted fix; full `--compare-selftest` green; evidence archived. |

---

## Guiding Principle

> **Simple, always working, clear status.** A "sync" must never copy an identical file or delete an unchanged source file. Never serve a stale decision after a mutation. One source of truth per control. Prove every destructive path on the user's real route.

## Performance / Safety Validation Contract (mandatory, per repo rules)

Compare Directories is a performance-sensitive subsystem (`Specs/Core/Core_CompareDirectories.md` "Performance Validation Contract"). Every slice that touches destructive correctness (**CX0-1, CX0-2, CX0-3, CX2-1**) MUST include a **tree-equality / byte-integrity assertion** and an **out-of-tree sentinel assertion**, and MUST archive before/after runs under `Specs/TestRuns/<machine>/Compare/<timestamp>/`.

- New self-test cases MUST be registered in the compare selftest family/case registry — order-array membership alone reports-but-never-runs (same class of pitfall as `kFileOpsFamilyDefinitions`, per the `fileops-selftest-family-registration` memory). Verify the registration list vs defined cases as part of CX4.
- Run via `--compare-selftest` (or `.\Tools\Run-AllTests.ps1 -Suite Full`, ≈18-20 min). A **session-global mutex serializes runs across worktrees — expect exit 3 on contention** (per the `fileops-selftest-run-timing` memory); do not interpret exit 3 as a failure.
- Plugin/engine perf counters land in **ETW**, not the readable log. Wire any `compare.ui.*` / cache-budget high-water assertions to ETW counters, not stdout.

---

## Phase P0 — Stop the data loss / correctness / crashes (ship-blocking; slices independent unless noted)

### CX0-1. `[CRITICAL]` Directory sync Copy/Move has no diff manifest — copies the whole subtree; Move deletes identical source files
*(covers findings: `no-diff-manifest-whole-subtree-copy`, `sync-copy-diff-manifest-missing`, `cross-vs-same-context-copy-inconsistency`, `sync-copy-no-overwrite-confirmation-spec-gap`)*

**Symptom (reachable, normal use):** In compare mode, Copy-to-other-pane on a directory that mostly matches re-copies every identical descendant (slow, overwrites identical destination files). Move-to-other-pane is worse: it removes identical files from the *source* even though the user intended to transfer only differences — **silent data loss / unexpected source mutation.**

**Root cause:** `CompareDirectoriesWindow::OnCommand` routes `IDM_PANE_COPY_TO_OTHER` / `IDM_PANE_MOVE_TO_OTHER` straight to `_folderWindow.CommandCopyToOtherPane(pane)` / `CommandMoveToOtherPane(pane)` (`CompareDirectoriesWindow.cpp` ~868-869). Those pass the raw selected directory paths plus `FILESYSTEM_FLAG_RECURSIVE` (`FolderWindow.FileOperations.cpp` ~1058-1138). In the common same-context case the engine calls `_fileSystem->CopyItem(dir, RECURSIVE)` where `_fileSystem` is the compare wrapper, which delegates **unconditionally** to `_baseFs` with the recursive flag (`CompareDirectoriesEngine.cpp` ~3656-3759). The base plugin then runs its own on-disk recursion, blind to compare decisions → the entire real subtree is transferred. There is **no diff-manifest code anywhere** (repo-wide grep for `manifest`/`SyncCopy`/`sourceAbsolutePath`/`BuildSyncList` returns nothing). The spec's "Directory copy/move behavior (sync)" section (`Core_CompareDirectories.md` ~573-587) is entirely unimplemented.

*Divergence trap:* when the two panes happen to use *different* plugins (cross-FS bridge), `CopyDirectorySequential` enumerates via the filtered compare wrapper (`State.cpp` ~6877) and accidentally skips identical files — so the **same gesture produces different result sets depending on plugin identity**, and cross-FS testing masks the same-FS data loss.

**Fix:**
1. Build the spec manifest in the window from **cache-only** decisions (never synchronous traversal on the UI thread). Add a testable, extractable helper — e.g. `CollectCompareSyncManifest(ComparePane fromPane, const std::vector<path>& selectedRelative, std::vector<std::pair<path,path>>& outPairs)` — that iteratively walks `TryGetCachedDecision` for the selected directory and all descendant folders, collecting only `isDifferent` items (respecting enabled criteria, ignore patterns, `compareSubdirectories`), and translates each to `(sourceAbsolutePath, destinationAbsolutePath)`.
2. A directory that exists **only on one side** → add as a single recursive top-level item (whole subtree is new; no per-file filtering).
3. Submit the flat pair list through a **non-recursive** entry point (e.g. `CopyItems`/`MoveItems` without `FILESYSTEM_FLAG_RECURSIVE`, pre-creating intermediate destination directories) so the base plugin never re-enumerates. **Never pass a directory with `FILESYSTEM_FLAG_RECURSIVE` in compare mode.**
4. For **Move**, use copy-verify-delete restricted to the manifest: identical source files must never be deleted. (Closes `sync-copy-no-overwrite-confirmation-spec-gap`.)
5. Surface a sync confirmation / accurate task-card item count reflecting the manifest size before a destructive Move.
6. Gate the whole path on a fully-scanned subtree: if any descendant decision is still `SubdirPending`/`ContentPending`, either block with "comparison still in progress" or force-drain before building the manifest — a manifest built on pending state can omit genuinely-different (not-yet-compared) files. (This is why CX0-1 depends on a correct pending model; see CX0-2.)

**Required proof:** new `Crosscut_SyncManifestCopiesOnlyDifferences` (T-8) — multi-depth left/right tree mixing identical + different + only-on-one-side files; drain scan to idle + flush subdir/content updates; run the manifest collector and assert the **exact** `(source,destination)` pair set: different present, identical **absent**, one-sided subtree whole. Add `Crosscut_SyncMoveKeepsIdenticalSource` asserting Move does not delete identical source files. Both RED on current HEAD. Byte-equality + out-of-tree sentinel assertions per the Validation Contract. **Largest slice — see deeper refactors.**

### CX0-2. `[HIGH]` `InvalidateForAbsolutePath` never bumps `_version` → in-flight workers re-insert stale pre-mutation decisions
*(finding: `invalidate-for-path-no-version-bump`; hand-verified)*

**Symptom:** After a sync copy/move/delete in the compare window, the affected folder can show stale results — items just made identical still shown as different, or freshly created/deleted files mis-reported. The user makes data-movement decisions on this view, so a stale "different" causes redundant copies and a stale "identical" skips a needed copy.

**Root cause:** `InvalidateForRelativePathLocked` erases the affected cache entries + pending content updates under `_mutex` and ends with `++_uiVersion;` (`CompareDirectoriesEngine.cpp` ~621) — it **never increments the atomic `_version`** and never drops the affected keys from `_scanScheduledKeys` / `_scanInFlightKeys` nor cancels content jobs. The directory enumeration in `ComputeDecisionForFolder` runs **without `_mutex`**; the lock is only re-acquired for writeback, guarded solely by `if (_version.load() == job.version) { _cache[job.key] = storedDecision; ... }` (~3228). A worker that enumerated the folder *before* the file op and finishes *after* the invalidation keeps `job.version == _version`, passes the guard, and re-inserts a pre-mutation decision into the just-invalidated key. `OnFolderWindowFileOperationCompleted` calls this on every file-op completion while workers run (`CompareDirectoriesWindow.cpp` ~2586-2603). The spec (`Core_CompareDirectories.md` ~509) explicitly requires "Bumps `_version` so stale decisions are not served"; `SetRoots`/`SetSettings`/`Invalidate` all do — `InvalidateForAbsolutePath` does not.

**Fix:** In `InvalidateForRelativePathLocked`, after erasing keys: increment `_version` (`fetch_add`) **and** bump `_backgroundWorkCancelToken` / clear content in-flight for the affected scope, **and** erase the affected keys from `_scanScheduledKeys`/`_scanInFlightKeys`, then re-enqueue a fresh high-priority scan. A global `_version` bump is correct but coarse (invalidates the whole cache) — if that proves too aggressive for large trees, store a **per-key epoch** in the decision and check it at writeback instead. Land **before/with CX0-1**: CX0-1's Move issues deletes that hit this exact path; without CX0-2 the manifest's own mutations get stale-overwritten.

**Required proof:** extend T-3 (content-compare-active contention): during an in-flight scan/content compare, call `InvalidateForAbsolutePath` and assert via perf stats that old-version writebacks are dropped (`_version` incremented, old `_scanInFlightKeys` cleared) and no stale decision is served after invalidation. RED on current HEAD.

### CX0-3. `[HIGH]` Failed-enumeration decisions are cached forever, violating the spec retry contract
*(covers findings: `failed-decision-cached-no-retry`, `failed-enumeration-cached-not-retried`)*

**Symptom:** A folder that hits a transient/fixable enumeration failure (NTFS access-denied later corrected, S3/FTP throttle/timeout, network blip, locked folder) stays stuck showing the failure on every navigation back — no retry — until a full Rescan.

**Root cause:** Both decision-producing paths cache the decision whenever its version matches, with **no `SUCCEEDED(decision->hr)` guard**. `ComputeDecisionForFolder` returns a decision carrying the failed HRESULT from `TryReadDirectoryEntries` (~2602/2617); `ScanWorker` stores it into `_cache[job.key]` (~3228-3235) and `GetOrComputeDecision` into `_cache[rootKey]` (~3071-3077) unconditionally. `TryGetCachedDecision` (~745-759) then serves the cached failure on any version match. The spec MUST (`Core_CompareDirectories.md` ~344-349): "The failed decision is **not cached** — the engine retries enumeration on the next access."

**Fix:** Gate all three inserts on `SUCCEEDED(decision->hr)` (the missing-folder case is already mapped to `S_OK` and stays cacheable). Return failed decisions to the immediate caller (so the wrapper's `ReadDirectoryInfo` can surface the HRESULT in the details line) but do not store them; guard the cache-hit fast paths in `GetOrComputeDecision`/`TryGetCachedDecision` to not serve a stale failed entry.

**Required proof:** L-1b — `CreateReadDirectoryBehaviorFileSystem` with `forcedHr=E_ACCESSDENIED` then `S_OK` (toggle): first access asserts `FAILED(hr)`; after flipping to success, next access asserts `SUCCEEDED(hr)` and correct items **without a version bump** (proves not cached). Assert decision-cache byte estimate did not retain the failed entry. RED on current HEAD.

### CX0-4. `[HIGH]` (disputed 1/2, mechanically verified) Dangling iterator in `FlushPendingContentCompareUpdatesBudgeted` → heap corruption
*(finding: `flush-pending-content-iterator-invalidation`)*

**Symptom:** Intermittent heap corruption / crash during the UI decision-refresh flush when the decision cache is over its 300MB budget while content-compare updates are pending — exactly the large S3↔FS subtree-with-content scenario.

**Root cause:** The flush loop copies `key = it->first`, does `++it`, then calls `ApplyPendingContentCompareUpdatesLocked(key)` (`CompareDirectoriesEngine.cpp` ~698-700). That call invokes `MaybeEvictDecisionCacheLocked()` when it applies an update (~1892), which does `_pendingContentCompareUpdates.erase(candidateKey)` (~1223). If the evicted candidate is the folder `it` currently points at, `it` is invalidated; the next loop test/deref is UB. A pending-update folder is eligible for eviction (eligibility only requires not-pinned, not in `_pendingSubdirUpdates`, not `anyPending` — none guaranteed). One verifier disputed reachability, but the erase mechanism was confirmed and the snapshot fix is correct regardless. Low effort, high blast radius — do early.

**Fix:** Do not hold a live `_pendingContentCompareUpdates` iterator across `ApplyPendingContentCompareUpdatesLocked`. Snapshot up to `maxFoldersToApply` keys into a local `std::vector<std::wstring>` under the lock, then iterate the vector re-checking membership; or re-fetch `begin()` each iteration (as `FlushPendingSubdirUpdatesBudgeted` already does). Sequence **before CX1-1** (CX1-1 changes how eviction touches the pending-updates map).

**Required proof:** T-3 with a low cache budget (`SetDecisionCacheBudgetBytesForSelfTest`) + many pending content updates + concurrent `Invalidate`: no crash under ASAN/repeated runs; code review confirms no live iterator across the apply call.

### CX0-5. `[HIGH]` Non-`file` plugin compares are broken — display-form path fed to `TryMakeRelative`
*(finding: `getcurrentpath-vs-pluginpath-trymakerelative`)*

**Symptom:** For any cross-plugin / non-`file` compare (S3↔FS, FTP, 7z — the headline scenario): **Restore/Invert Differences Selection do nothing**, the post-completion "No differences" / "doesn't exist in this hierarchy" empty-state is cleared instead of applied, and panes can appear stuck on "Comparing…".

**Root cause:** Several sites feed `FolderWindow::GetCurrentPath(pane)` (the display/history URL form `shortId:ctx|/path`) into `CompareDirectoriesSession::TryMakeRelative()`, which expects the **plugin** path form (`/path`). For non-`file` roots, `TryMakeRelative` normalizes the URL to `/s3:ctx|/sub` which never prefix-matches root `/sub` → returns `nullopt` → the feature silently no-ops. Affected: `CompareDirectoriesWindow.cpp` ~1054, 1058, 1073 (Restore/Invert selection) and `CompareDirectoriesWindow.Progress.cpp` ~203, 208, 235, 1118, 1122 (empty-state + decision-refresh). The working root path uses `GetCurrentPluginPath` (~2004-2010). For `file` plugins display==plugin form, so file-based selftests never see it. Note `CompareDirectoriesWindow.cpp:2247` (`hasContextChanged`) is a **correct** use of `GetCurrentPath` (it parses the display URL via `TryParseLocation`) — leave it.

**Fix:** Replace `GetCurrentPath` with `GetCurrentPluginPath` at the 8 listed sites. Leave ~2247 as-is.

**Required proof:** new dummy-plugin (non-`file`) compare selftest exercising Restore/Invert Differences Selection and post-completion empty-state; assert they take effect (they silently no-op today on URL-form roots). RED on current HEAD.

### CX0-6. `[HIGH]` Options hot-reload: surprise modal while hidden + silent loss of "kept" edits
*(covers findings: `external-reload-prompt-while-panel-hidden`, `stale-flag-survives-reopen-data-loss`)*

**Symptom:** (a) While the user is just browsing the compare result (Options panel hidden), an external/background settings edit can pop a modal "Reload from disk / Keep editing" prompt. (b) After "Keep editing", reopening the panel silently discards the kept edits and a *phantom* OK-time conflict prompt fires on an actually-clean panel.

**Root cause:** The Options dialog is created once at window construction and lives hidden for the window lifetime; `OnOptionsInitDialog` calls `SettingsHotReload::RegisterParticipant(dlg)` (~557) and removes it only at `WM_NCDESTROY` (~542). So `kSettingsReloadedFromDisk` → `OnOptionsSettingsReloadedFromDisk` runs even when hidden, with no `IsWindowVisible` guard (~1333-1354). Separately, "Keep editing" sets `_optionsStaleFromExternalReload = true` (~1347) but `ShowOptionsPanel(true)` unconditionally calls `LoadOptionsControlsFromSettings()` (clobbering kept edits) and never clears the stale flag, so a clean reopened panel still triggers `ResolveOptionsStaleSaveConflict` → `PromptStaleSaveConflict` on OK (~3101-3124, 3652-3726). Spec scopes this behavior to "External settings reload **while Options is open**" (~268-276).

**Fix:** (a) Register/unregister the reload participant in `ShowOptionsPanel(true/false)` (preferred), or gate `OnOptionsSettingsReloadedFromDisk` on `_optionsDlg` visibility — when hidden, silently adopt disk state and never prompt. (b) Clear `_optionsStaleFromExternalReload` at the top of `ShowOptionsPanel(true)` (a freshly loaded panel is by definition not stale). If kept edits must truly survive hide/show, snapshot and re-apply them instead of calling `LoadOptionsControlsFromSettings`.

**Required proof:** selftest fires `kSettingsReloadedFromDisk` with the panel hidden → asserts **no** modal; second selftest "keeps editing", reopens, asserts the panel is clean and OK raises **no** phantom conflict prompt.

### CX0-7. `[HIGH]` Preferences "Compare Directories" page never renders its section headers
*(finding: `prefs-compare-headers-never-shown`)*

**Symptom:** The Compare Directories page in the global Preferences dialog shows a flat, undifferentiated list of cards with no section headings (Subdirectories / Compare / Additional / Ignore) — inconsistent with every other Preferences page. (The separate in-window Options panel is a different code path and shows its 4 headers correctly.)

**Root cause:** `LayoutDxPage` hides all headers up front (`Preferences.CompareDirectories.cpp` ~712-718: `header->SetVisible(false)`) and **never** calls `SetVisible(true)`/`SetBounds` on any header — whole-file grep confirms no `header->SetBounds`, no re-show, no `layoutHeader` equivalent. The sibling `Preferences.FileOperations.cpp` (~968-977) defines a `layoutHeader` lambda doing exactly that before each section. The header index→section mapping is also muddled (~245-249, 519-529).

**Fix:** Add a `layoutHeader` lambda mirroring `Preferences.FileOperations.cpp` (~968-977) that makes the header visible, sets its bounds, and advances `y`; invoke it before each section's cards. Fix the header index→section mapping so headers 1..4 (SUBDIRS/COMPARE/ADVANCED/IGNORE) precede their cards; decide whether the page-level header[0] is shown. Reserve `headerHeight` vertical space.

**Required proof:** PT-1 — extend the Preferences debug snapshot to report the visible Compare section-header count; assert `== 4` in `validateCompareDirectoriesPageChrome` (mirrors the Options-panel `visibleDxBodyHeaderCount == 4` check). RED on current HEAD.

---

## Phase P1 — Concurrency / memory / crash-class (each independent)

### CX1-1. `[MEDIUM]` Decision-cache 300MB budget is effectively unenforced under a wide pending subtree
*(covers findings: `decision-cache-budget-unbounded-under-pending-subtree`, `eviction-inner-loop-rescans-skipped-prefix`)*

**Symptom:** On large/wide remote subtree compares (the exact S3↔FS scenario the budget was added for), the decision cache grows well past its 300MB target while a scan is in progress — partially re-introducing the GiB-scale memory pressure eviction was designed to prevent.

**Root cause:** `MaybeEvictDecisionCacheLocked` treats `anyPending` as an **absolute eviction veto** (`CompareDirectoriesEngine.cpp` ~1208). During a `compareSubdirectories` scan, every directory with still-scanning descendants has `SubdirPending` → `anyPending==true`, so tens of thousands of decisions are simultaneously unevictable; the budget check runs but evicts nothing and `_decisionCacheEstimatedBytes` grows unbounded. Secondary: `FlushPendingSubdirUpdatesBudgeted` (~786-818) and `PropagateChildAggregateToAncestorsLocked` (~1643) grow decisions via `TrackDecisionCacheInsertOrUpdateLocked` but **never** call eviction, so the budget is not enforced on the aggregation path at all. Tertiary (E-1): the eviction inner loop restarts from `rbegin()` each eviction (~1186), wasting the inspection budget when the LRU tail is a long run of pinned/pending keys.

**Fix:** Pending directory decisions are cheap to recompute (`TryGetCachedDecision` miss → `RequestScanForFolder` re-queues), so do **not** treat `anyPending` as an absolute veto — allow eviction of `anyPending` entries that are neither pinned nor an ancestor of a pinned (visible) folder once the budget is materially exceeded, or cap retained pending decisions. Call `MaybeEvictDecisionCacheLocked()` after the `FlushPendingSubdirUpdatesBudgeted` batch (and consider it in the propagate path). Collapse the `rbegin()` re-scan into a **single reverse pass** gathering up to `kMaxEvictionsPerCall` candidates, then erase. Sequence **after CX0-4**.

**Required proof:** budget test with `SetDecisionCacheBudgetBytesForSelfTest` low + a wide pending subtree: assert `_decisionCacheEstimatedBytes` (ETW high-water counter) stays within a bounded multiple of budget. RED on current HEAD (grows unbounded today).

### CX1-2. `[MEDIUM]` Leave-scope `MessageBox` runs a nested modal pump inside the navigation callback

**Symptom:** A pending deferred Rescan (or in-flight nav) processed while the leave-scope prompt is open can change roots, re-enter compare mode, or refresh panes; when the modal returns, the code may re-navigate to a stale `previousPath` or `CancelCompareMode` on a just-restarted session, leaving panes/roots inconsistent.

**Root cause:** `SyncOtherPanePath` is invoked synchronously from `OnPanePathChanged`, the `_panePathChangedCallback` fired at the tail of `FolderWindow::SetFolderPath`. When the user navigates outside the roots while a run is pending it shows `MessageBoxCentered` (`CompareDirectoriesWindow.cpp` ~2276), spinning a nested pump while still inside the originating `SetFolderPath`. During that pump, `kCompareDirectoriesDeferredStart` and progress/decision handlers run; the recursion guard `_syncingPaths` is a single non-reentrant bool, so nested ops leave it in an arbitrary state and the window state machine (`_compareActive`/`_compareRunPending`/`_compareRunId`/deferred phase) advances under a mid-navigation stack acting on now-stale `previousPath`/`changedPane`. Single UI thread, so not memory-unsafe — a state-coherence hazard around the most safety-relevant prompt.

**Fix:** Do not show a modal from the navigation callback. Either (a) revert the pane to `previousPath` immediately and `PostMessage` a request to show the prompt once the stack unwinds (decide cancel/proceed there — this also removes the visual flash of the out-of-scope folder behind the box), or (b) add a dedicated `_leaveScopePromptActive` flag that suppresses deferred-start processing and further sync while the prompt is open.

**Required proof:** selftest/inspection that a deferred-start cannot run reentrantly while the prompt is open; pane state coherent after either choice.

### CX1-3. `[MEDIUM]` Latent double-free if the window self-destructs during `CreateWindowExW`

**Symptom:** Heap corruption / double-free if the window ever fails to fully create (or self-closes during creation). Currently dormant only because `OnCreate` unconditionally returns true.

**Root cause:** Mixed ownership. `ShowCompareDirectoriesWindow` holds `std::unique_ptr<CompareDirectoriesWindow>` across `Create()` and `release()`s only after it returns true. But `CreateWindowExW` (inside `Create`) synchronously dispatches `WM_NCCREATE..WM_NCDESTROY`; `OnNcDestroy` sets `_deletePending=true` and the `WndProcThunk` `scope_exit` does `delete self` when `_dispatchDepth==0` (`CompareDirectoriesWindow.cpp` ~595-619, 2629-2636). If the window is destroyed during creation, the object self-deletes, `Create()` returns false, and the owning `unique_ptr` deletes the **same object again** (~803-808).

**Fix:** Make ownership unambiguous — either (a) `Create()` relinquishes C++ ownership up front (`new`, hand to `CreateWindowExW`, let only `OnNcDestroy` own lifetime; return false without deleting because the object already self-deleted), or (b) keep the `unique_ptr` but guarantee the object can never self-destruct before `Create()` returns (assert `_deletePending==false` after `CreateWindowExW`; disable the deferred-delete path during creation).

**Required proof:** inspection + an `OnCreate`-returns-false injection: no double-free; ASAN clean.

### CX1-4. `[LOW]` `_dxMenuBar` dereferenced after modal `ContextMenu::Show` without null re-check

**Symptom:** If the compare window / menu-bar host is destroyed while a top-level menu popup is open (window close routed through the modal loop), post-`Show` code dereferences a null `_dxMenuBar` and crashes.

**Root cause:** `OpenDxMenuBarPopup` null-checks `_dxMenuBar` before `ContextMenu::Show` (~783) but `Show` runs a nested modal loop during which `HandleDxChromeHostMessage`'s `WM_NCDESTROY` sets `_dxMenuBar = nullptr` (~882-886). After `Show` returns, `_dxMenuBar->SetSelectedIndex(std::nullopt)` (~856) has no re-check (`CompareDirectoriesWindow.Menu.cpp` ~852-859).

**Fix:** After `ContextMenu::Show` returns, re-check `if (! _hWnd || ! _dxMenuBar) return;` before touching `_dxMenuBar` (mirror the guard at the top of the function).

**Required proof:** inspection; close-during-popup does not deref null.

---

## Phase P2 — Correctness edges / UX defects (batch as convenient; each independent)

- **CX2-1 `[MEDIUM data-safety]` Trailing space/dot name-collision silently drops a same-side entry.** `NormalizeEntryNameForCompare` strips trailing spaces/dots and the trimmed name is the map key via `outEntries.emplace(...)` (`CompareDirectoriesEngine.cpp` ~2039); `std::map::emplace` does not overwrite, so two same-side names normalizing to the same key (`report` vs `report.`, `data` vs `data `) drop the second — invisible in compare and excluded from the sync manifest. **Fix:** detect `inserted==false`; fall back to the untrimmed key for the colliding entry (or flag ambiguity), and retain the original on-disk name. **Proof:** selftest with two same-side names differing only by trailing dot/space — both represented.
- **CX2-2 `[LOW]` Ignored directory's own leaf not re-checked on direct navigation.** `ShouldIgnoreEntry` runs only on child entries; `ComputeDecisionForFolder`/`RequestScanForFolder` never test `relativeFolder`'s leaf/ancestors against `ignoreDirectoryPatterns` (`CompareDirectoriesEngine.cpp` ~298-320, 2530-2544), so an ignored dir reached by direct/synchronized navigation is still compared (spec ~258 says subtrees are excluded). **Fix:** return empty/identical without enumerating when any ancestor/leaf matches an enabled ignore-directory pattern. **Proof:** selftest navigating directly into an ignored-dir relative path.
- **CX2-3 `[LOW]` Leave-scope Cancel with empty `previousPath` strands the pane out-of-scope.** The IDCANCEL branch reverts only when `previousPath.has_value()`, else returns without reverting or canceling compare mode (`CompareDirectoriesWindow.cpp` ~2278-2295), leaving the pane out-of-scope with `_compareActive` true. **Fix:** fall back to `ResolveAbsolute(changedPane, {})` (session root) or treat as "proceed" + `CancelCompareMode()`.
- **CX2-4 `[LOW]` InProgress watermark only pulses if the spinner control exists.** `FolderView::SetBackgroundWatermark` invalidates once; the only repeated-invalidation driver is `OnProgressSpinnerTimer`, gated on `_scanProgressBar` being non-null (`CompareDirectoriesWindow.Progress.cpp` ~740-805, 899-942). If the spinner control fails to create, the spec-required animated watermark renders frozen. **Fix:** drive watermark pulse from its own timer keyed on `_watermarkState==InProgress`, or invalidate panes even when `_scanProgressBar` is null.
- **CX2-5 `[LOW]` `ScheduleDecisionRefresh` drops a pending refresh if `SetTimer` fails.** On failure, `_decisionRefreshTimerActive` stays false and `_decisionRefreshPending` stays true with nothing re-arming it (`Progress.cpp` ~157-167) → fresh decisions may never reach the panes. **Fix:** on failure, `PostMessage` a fallback or flush synchronously; log `Debug::ErrorWithLastError`.
- **CX2-6 `[LOW]` ETA shows absurd values at tiny rates.** `secondsD = remaining / _contentEtaSmoothedBytesPerSec` gated only by `> 1.0` (`Progress.cpp` ~486-491) → e.g. `277777:46:40` on a stalled remote. **Fix:** clamp / show "estimating…" above a sane threshold; raise the rate floor.
- **CX2-7 `[LOW]` Task-card throttle leaves stale intermediate state.** `UpdateCompareTaskCard` throttles to 50ms and returns early with no trailing flush (`Progress.cpp` ~968-979), so a bursty-then-quiet compare leaves stale per-file bytes until the next message. **Fix:** set a dirty flag + one-shot trailing-flush timer.
- **CX2-8 `[LOW]` Wheel remainder not reset when unscrollable.** The `_optionsScrollMax<=0` early-out skips resetting `_optionsWheelRemainder` (`Options.cpp` ~1641-1660), so accumulated sub-notch deltas leak across a non-scrollable→scrollable transition. **Fix:** reset `_optionsWheelRemainder = 0` when `_optionsScrollMax <= 0`.
- **CX2-9 `[LOW]` `compareContent` persists checked-while-disabled.** When `IFileSystemIO` is unsupported the toggle is disabled but its checked state is not cleared, and `ReadOptionsControlsToSettings` reads it verbatim (`Options.cpp` ~2059-2067, 3168-3171) → persists `compareContent=true`. Cosmetic (engine guards every use with `&& IsContentCompareSupported()`). **Fix:** AND-mask `compareContent` with support in `ReadOptionsControlsToSettings`, or uncheck on disable. *(Largely obviated by CX3-3, which makes DxUi the single source of truth.)*

---

## Phase P3 — Architecture & simplification (de-risks future work; largest payoff first)

### CX3-1. `[MEDIUM]` Extract one `ApplyCriteriaDiffAndSelection(...)` — diff/selection logic is triplicated

The logic turning `(sizeDifferent, timeDifferent, attrsDifferent, contentDifferent)` into a `differenceMask` + `selectLeft`/`selectRight` (Size→bigger, DateTime→newer, Attributes→both, Content→both) is written verbatim **three times**: `ComputeDecisionForFolder` file branch (`CompareDirectoriesEngine.cpp` ~2843-2910), `ApplyPendingContentCompareUpdatesLocked` new-item branch (~1717-1782) and existing-item branch (~1814-1877). The copies must stay byte-identical or the post-content-compare render disagrees with the initial scan about which side to select — a subtle data-safety risk when the user then copies the "selected" side. **Fix:** extract a single free function `void ApplyCriteriaDiffAndSelection(CompareDirectoriesItemDecision&, const Settings&, bool canCompareContent, bool contentDifferent)`; call from all three sites. **Do this first** — CX0-1's manifest and content-compare both depend on a single-sourced selection contract. **Proof:** existing compare battery green; single definition asserted by grep.

### CX3-2. `[MEDIUM]` Delete the dead split-host DxUi layout branch + never-attached per-control members
*(covers findings: `dead-legacy-split-host-layout-path`, `options-dead-per-control-dx-hosts-and-fallback-layout`)*

The stabilized panel always uses the single `body` host: `LayoutOptionsControls` returns via the body branch whenever `_optionsDxUi && usesDxUiStatics && body.hostHwnd` (`Options.cpp` ~2493/2738), and `EnsureOptionsDx*Hosts` only ever create `body` (per-control `hostHwnd`s stay null). The ~360-line fallback (~2741-3099) runs only when `_optionsDxUi==nullptr`, and even then every `dxHeader/dxCard/dxToggle/dxEdit` argument is `nullptr`, so those sub-branches are permanently dead. The backing members (`OptionsSectionDxLabel`, `OptionsCardDxText`, `OptionsToggleDx`, `OptionsEditDx` in `Internal.h` ~462-595, flagged "Compatibility-only placeholders") are never attached. **Fix:** delete the dead dxCard/dxToggle/dxHeader/dxEdit branches (collapse `layoutSectionHeader`/`layoutToggleCard`/`layoutIgnoreCard` to pure native-HWND positioning) and remove the unused structs/members; keep `body`/`okButton`/`cancelButton`. **Proof:** compare battery + Options-panel test green; grep confirms removal. *(Prerequisite for CX3-3.)*

### CX3-3. `[HIGH — user-requested]` Remove the hidden-Win32 shadow control model — make DxUi the single source of truth in Options
*(finding: `options-legacy-hwnd-shadow-state-model`)*

**This is the explicit ask: "remove the Options panel that keeps a hidden Win32 control."**

**Current (broken) design:** The Options panel maintains **two complete control trees**. `EnsureOptionsControlsCreated()` creates real but **hidden** Win32 controls for every option — `BS_AUTOCHECKBOX` buttons (`makeToggle`), `Edit`s (`makeFramedEdit`), DxHost statics (`_optionsUi.*`) — and a parallel DxUi tree (`_optionsDxUi->body.*`). The **hidden controls are the canonical data model**: `ReadOptionsControlsToSettings()` (the function that persists user choices) reads exclusively from `_optionsUi.*.toggle` via `GetTwoStateToggleState` (`Options.cpp` ~3160-3178), never from the DxUi toggles. A click on a DxUi toggle is round-tripped to the hidden legacy control via a synthetic `BN_CLICKED` (`syntheticHiddenToggle`, `OnOptionsCommand` ~1384), then re-mirrored by `SyncOptionsDxToggles()`. State lives in hidden HWNDs; rendering happens in DxUi; hand-written `Sync*` mirrors keep them aligned. Every option is wired in **three** places and can desync silently (enable-state, content-compare-unavailable — see CX2-9). This is the single largest source of risk and bloat in the 3745-line Options file.

**Proof it is viable:** `Preferences.CompareDirectories.cpp` already implements a DxUi-only compare-settings panel with **no** hidden-control model. Mirror its structure.

**Migration runbook (ordered; do after CX3-1 and CX3-2 so dead paths are already gone):**
1. **Inventory the contract.** Enumerate every option the panel owns (`compareSize`, `compareDateTime`, `compareAttributes`, `compareContent`, `compareSubdirectories`, `compareSubdirectoryAttributes`, `selectSubdirsOnlyInOnePane`, `keepIdenticalItems`, `ignoreFiles`+pattern, `ignoreDirectories`+pattern). For each, note: persisted field, enable/disable rule, and any interlock (e.g. `compareContent` disabled when `!IsContentCompareSupported()`).
2. **Make DxUi the source of truth for reads.** Rewrite `ReadOptionsControlsToSettings()` to read each value from the DxUi `Toggle`/`TextField` in `_optionsDxUi->body.*` (use the same accessors `Preferences.CompareDirectories.cpp` uses), not from `_optionsUi.*`.
3. **Make DxUi the source of truth for writes.** Rewrite `LoadOptionsControlsFromSettings()` to set DxUi control state directly; delete the legacy `SetCheck`/`SetWindowText` calls into `_optionsUi.*`.
4. **Delete the synthetic round-trip.** Remove `syntheticHiddenToggle` / `BN_CLICKED` re-dispatch in `OnOptionsCommand` (~1356-1417) and the `Sync*FromLegacy` mirrors (`SyncOptionsDxToggles`/`Edits`/`Statics`/`Buttons` that exist solely to copy legacy→Dx).
5. **Delete the hidden control set.** Remove `EnsureOptionsControlsCreated`'s creation of hidden toggle/edit/static HWNDs and the `_optionsUi` `OptionsToggleCard`/`OptionsIgnoreCard` members from `Internal.h` (~342-381). Keep only the host window(s) DxUi needs.
6. **Re-home enable/disable + interlocks** (CX2-9) onto the DxUi controls: `compareContent` disabled+unchecked when unsupported; ignore-pattern edits enabled by their toggle.
7. **Dirty-tracking & reload.** Re-point `IsOptionsDialogDirty()` and the CX0-6 hot-reload/stale logic at the DxUi control values (single comparison source).
8. **Focus traversal & accessibility.** Verify forward/reverse keyboard traversal at small window sizes now flows through the DxUi controls only (the hidden HWNDs were also tab stops); update `Commands.SelfTest.CompareOptions.cpp` expectations (`visibleDxBodyHeaderCount`, toggle/edit descendants) accordingly.

**Why it must follow CX3-1/CX3-2:** CX3-2 removes the dead per-control Dx members this work would otherwise have to reason about; CX3-1 single-sources the selection contract the options drive. Doing CX3-3 first would mean editing code slated for deletion.

**Required proof:** Options-panel selftest (`Commands.SelfTest.CompareOptions.cpp`) green reading state from DxUi; round-trip test (set each toggle/edit via DxUi → OK → reopen → values preserved); grep confirms `_optionsUi.*` toggle/edit/static + `Sync*FromLegacy` + `syntheticHiddenToggle` are gone. Run the full compare + commands batteries.

### CX3-4. `[MEDIUM]` Decompose the `CompareDirectoriesWindow` god-class

`CompareDirectoriesWindow` declares the entire window — menu (Dx+legacy), banner (Dx+legacy), options (after CX3-3: DxUi only), progress/ETA/spinner/watermark/task-card, sync navigation, selection, details/metadata providers, ~25 `ENABLE_TESTS` accessors — ~150 members / ~120 methods across 6 `friend`-coupled `.cpp` files sharing a flat mutable flag set (`_compareStarted`/`_compareActive`/`_compareRunPending`/`_compareRunSawScanProgress`/`_bannerRescanIsCancel`/`_syncingPaths`/`_deferredCompareStart*`). **Fix:** extract cohesive sub-controllers owned by the window — `ChromeController` (menu+banner), `OptionsPanelController` (owns the Options state + methods), `ProgressController` (`BannerProgressState`+ETA+spinner+task card) — each with a narrow interface instead of friend access to all state. Depends on CX3-3 (Options) and CX3-2 landing. **Proof:** compare battery green after extraction; reduced friend coupling. `B-1`, `MD-2`, `API-1` (CX3-5) fold in naturally.

### CX3-5. `[LOW]` Dedup & API hygiene (fold into CX3-4)

- **B-1** `EnsureOptionsDxButtonHosts` runs twice at creation (WM_INITDIALOG path + `CreateChildWindows` ~1678), tearing down and rebuilding just-created hosts (`attachButtonHost` begins with `slot.Detach()`). **Fix:** drop the second call, or no-op when already attached.
- **MD-2** `CompareDirectoriesWindow.Menu.cpp` (~44-244) reimplements `SplitMenuText`/`StripMenuMnemonicMarkers`/`FindMenuMnemonic`/`TryGetMenuItemPresentationText`/`ConvertHMenuToDxFlyoutItems`/`BuildDxMenuBarItems` byte-for-byte vs the shared `Common/DxUi/DxUiNativeMenuInterop.h` (`ConvertNativeHMenuToFlyoutItems`/`BuildNativeMenuBarItems`, already used by the viewers). **Fix:** delete the local copies; call the shared functions with a default `NativeMenuFlyoutOptions{}`.
- **API-1** `GetOrComputeDecision` is public but has zero production callers (UI uses `TryGetCachedDecision` only) and can block on `_contentCompareQueueNotFullCv` backpressure — it invites a future UI-thread stall. **Fix:** move behind the `ENABLE_TESTS` hook (like `SetDecisionCacheBudgetBytesForSelfTest`) or rename to expose the blocking nature.

---

## Phase P4 — Tests (each pins a fix above; write alongside its phase, run full suite at closeout)

Register every new case in the compare selftest family/case registry (membership in the order-array alone reports-but-never-runs — `fileops-selftest-family-registration` memory).

**P0 pinning tests:**
- **T-8** `Crosscut_SyncManifestCopiesOnlyDifferences` + `Crosscut_SyncMoveKeepsIdenticalSource` — pins CX0-1 (disputed-as-missing today; mandatory once the manifest lands). Multi-depth tree; assert exact `(source,dest)` pair set + Move keeps identical source. Byte-equality + out-of-tree sentinel.
- **L-1b** `Crosscut_FailedEnumerationNotCachedRetries` — pins CX0-3. `forcedHr=E_ACCESSDENIED`→`S_OK`; recompute without version bump.
- **CX0-5 test** non-`file` dummy-plugin compare exercising Restore/Invert selection + post-completion empty-state — pins CX0-5.
- **CX0-6 tests** reload-while-hidden (no modal) + keep-editing-reopen (clean, no phantom prompt) — pins CX0-6.
- **PT-1** Preferences visible Compare header count `== 4` — pins CX0-7.

**P1 pinning tests:**
- **T-3** `Crosscut_ContentWorkerVsInvalidateStress` — `compareContent=true` over many equal-size byte-different files + readers + repeated `Invalidate`/`SetBackgroundWorkEnabled` toggles + low cache budget. Asserts: no crash (pins CX0-4 dangling iterator), no null decision, no stale-equal, and `_decisionCacheEstimatedBytes` bounded (pins CX1-1). Also drives the CX0-2 stale-writeback assertion. The single most important new concurrency test.
- **T-4** `Crosscut_SetRootsResetsAndBumpsVersion` — warm rootsA + content compares, `SetRoots(rootsB)`; assert versions bumped, old keys stale/none, queues/in-flight drained, new decisions correct. (`SetRoots` is currently called by **no** test.)

**Standing data-safety tests (write with their nearest phase; they guard against future regressions even though current code is correct on these paths):**
- **T-1** `Crosscut_ContentEqualSizeEqualMtimeDiffers` — force identical sizes + identical `SetFileLastWriteTime` both sides, `compareContent=true` only; assert `isDifferent` + Content bit + `selectLeft&&selectRight`. Mirror with identical bytes+mtime asserting NOT different (proves the byte compare ran). **The regression net for the most likely future data-loss optimization** (a size+mtime dedup fast path).
- **T-2** `Crosscut_UnknownSizeStreamingCompare` — `IFileReader` wrapper whose `GetSize` returns `E_FAIL` (forces `sizeKnown=false`); cases: both empty, equal nonzero, left-prefix-of-right, right-prefix-of-left, mid-stream mismatch, all with short reads. Guards the unknown-size EOF logic (`CompareDirectoriesEngine.cpp` ~2243-2262) for FTP/S3/archive backends.
- **T-5** `Crosscut_ContentCacheHitSkipsIo` — counting `IFileSystemIO`; second compare of an unchanged pair opens no files; assert `ContentCompareKeyEq` distinguishes mtime-only and size-only keys (hash/eq contract, ~408-422).
- **T-6** `Crosscut_SelectSubdirsOnlyInOnePane` — assert one-sided-directory selection differs between `selectSubdirsOnlyInOnePane=false` vs `true`; reconcile with the existing `unique` case so both match the spec.
- **T-7** `Crosscut_MissingSideEmptyEnumeration` — build a relative subfolder only on the left, call the right wrapper's `ReadDirectoryInfo` on the missing-side relative folder; assert `SUCCEEDED` + zero entries (not a failure). Pins the empty-state contract.
- **CX2-1 test** two same-side names differing only by trailing dot/space — both represented.

**Regression gate:** none of T-1, T-2, T-3(content), T-4, T-8 exist today — the suite is green **because** these paths are unexercised. A fix without its pinning test does not close the finding; each pinning test must FAIL on a deliberately reverted fix.

---

## Verified correct — do NOT re-open without new evidence

The review's refutation pass rejected 12 candidate findings, and the following were examined and found sound (do not rework as part of Crosscut):
- **Version/coherence skeleton** — `_version` increments on `SetRoots`/`SetSettings`/`Invalidate`; cached decisions tagged with version; in-flight `(version, cancelToken)` stamp model correctly discards stale scan/content results. The **only** coherence defect is `InvalidateForAbsolutePath` omitting the bump (CX0-2).
- **Size-known content read loop** (`CompareFileContent` ~2160-2327) — short-read handling, 256KB streaming, min-available compare, the post-loop extra-byte EOF check (reading 1 trailing byte each side to catch a lying `GetSize`), and `sizeKnown` size/zero shortcuts are defensive and correct. The **unknown-size** branch is correct-but-untested (T-2).
- **Content equality is always byte-verified** — there is no metadata-only "assume equal" shortcut; `_contentCompareCache` only ever stores results produced by an actual byte comparison. The size+mtime cache key is a narrow staleness window (in-place edit preserving size *and* mtime), not a metadata-only equality decision.
- **Name matching** — the ordered `std::map` with `WStringViewNoCaseLess` (`CompareStringOrdinal` case-insensitive) correctly avoids the hash/equality contract violation an `unordered_map` would create. The **only** matching defect is the trailing dot/space collision drop (CX2-1).
- **Throttled progress notifications** and the `WM_NCDESTROY` payload draining were not found defective.

If a future change appears to contradict one of these, attach new source evidence before reopening.

---

## Concrete execution order

1. **CX3-1** (extract `ApplyCriteriaDiffAndSelection`) first — single-sources the selection contract the manifest and content-compare depend on. Compare battery stays green.
2. Add RED tests for **CX0-2, CX0-3, CX0-5** on current HEAD; confirm they fail for the dangerous reason. Fix each (cheapest correctness/crash fixes first to stabilize the ground CX0-1 stands on): **CX0-4** (snapshot), **CX0-3** (SUCCEEDED gate), **CX0-2** (version bump), **CX0-5** (plugin path).
3. **CX0-1** (diff manifest) with **T-8** RED→GREEN. This is the largest slice; build the extractable `CollectCompareSyncManifest` helper so T-8 drives it directly. Include the Move copy-verify-delete + confirmation. Depends on CX0-2 (its mutations must invalidate correctly) and CX3-1.
4. **CX0-6** (Options hot-reload) and **CX0-7** (Preferences headers) with their tests — self-contained UI fixes.
5. **CX1-1** (eviction budget, after CX0-4), then **CX1-2/CX1-3/CX1-4** (reentrancy + crash-class guards).
6. **CX2-1** (name collision — data-safety) first in P2, then batch the remaining P2 UX fixes.
7. **CX3-2** (delete dead layout) → **CX3-3** (remove hidden-Win32 shadow model — the user's explicit ask) → **CX3-4** (god-class decomposition) → **CX3-5** (dedup). Each keeps the compare + Options + DxUi-menu batteries green.
8. Write standing data-safety tests (T-1, T-2, T-5, T-6, T-7) with their nearest phase; T-3/T-4 in P1.
9. **Closeout:** verify all new cases are registered in the compare selftest family; run `.\Tools\Run-AllTests.ps1 -Suite Full` (expect exit 3 only on cross-worktree mutex contention); archive before/after perf+safety evidence under `Specs/TestRuns/<machine>/Compare/<timestamp>/`; update `Specs/Core/Core_CompareDirectories.md` for any clarified contract; remove the TODO block at the top of that spec for items now done; move this plan to `Specs/Plans/Done/`.

## Out of scope / non-issues

- The size+mtime content-cache key is a deliberate design (attributes excluded to avoid spurious misses); only an *in-place edit preserving both size and mtime* could serve a stale "equal", and content is always byte-verified at least once per `(key)` per run. Not reopened; T-1 stands as the regression net should anyone add a true metadata-only fast path.
- The cross-FS bridge's accidental diff-filtering is **not** treated as "correct" — CX0-1's explicit manifest replaces both transport paths so the divergence disappears.

---

*Provenance: 14-lens adversarial multi-agent review of the CompareDirectories subsystem, branch `claude/elated-sinoussi-a39160`, 2026-06-15 (57 candidates → 43 confirmed / 2 disputed / 12 rejected; CX0-1 and CX0-2 hand-traced). Anchors are HEAD-relative — re-grep before editing.*
