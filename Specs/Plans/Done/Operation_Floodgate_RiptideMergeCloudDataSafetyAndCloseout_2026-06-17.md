# Operation Floodgate — Riptide-Merge Cloud Data-Safety & Closeout

**Status:** Done - final uninterrupted Full gate green on 2026-06-19
**Date:** 2026-06-17
**Author:** Independent adversarial code review of the **last two commits on master** — merge `0b6b8080a` ("Merge claude/zealous-poitras-ce383a into master") + checkpoint `92ed8ba06` ("Checkpoint Riptide remediation state"), i.e. the landing of Operation Riptide. 70-agent workflow (`wf_fc21c0a7-a2c`): 18 area×lens reviewers → per-finding adversarial verification (two perspective-diverse skeptics per critical/high) → synthesis. 45 findings, 39 survived verification, 6 refuted. The two headline cloud data-loss vectors were re-traced by hand in the merged blobs.
**Original reviewed tree:** net diff `e99d076de..0b6b8080a` (80 files, ~14.8k insertions). All finding-body line anchors are relative to the **`0b6b8080a` blobs** unless an explicit 2026-06-18 HEAD anchor is shown. **Current reconciliation:** master `8231f79a11` on 2026-06-18. Re-grep every anchor before code edits (repo standing rule).
**Builds on:** `Operation_Riptide_FairstreamRemediation_DataSafetyConflictParity_2026-06-15.md` (closed with this plan) and `Operation_Riptide_CrashQuickSearchStabilization_Findings_2026-06-17.md`. Riptide's structural work and its R0-1 fix are sound and are NOT re-opened here. Floodgate fixes what the Riptide merge shipped anyway, against its own hold.

---

## Final Closeout - 2026-06-19

**Verdict:** close this plan and move it to `Specs/Plans/Done/`. The final uninterrupted Full-suite gate is green after the closeout-order fixes for synthetic DxUI menu hover/input retry and FileOps Phase 11 local-plugin concurrency reset.

**Final green evidence:**

- `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 3.0`: passed `1010/963/0/47`, wall time `98m 44.2s`, run UTC `2026-06-19T10:36:22.1678835Z` to `2026-06-19T12:15:06.3925326Z`.
- Build inside the Full run passed with `0 warning(s), 0 error(s)`, log `.build/logs/msbuild-20260619_123255_015.log`.
- Final aggregate and per-suite artifacts archived at `Specs/TestRuns/LT-PF5VDAGE/FileOps/2026-06-19_123255/Floodgate_Final_Closeout/`.
- Focused post-fix checks before the Full gate: Commands exact closeout failures passed `3/0/0`; FileOps `Phase11_` prefix passed `9/0/0`.

**Closeout fixes after the first 2026-06-19 red Full gate:**

- Hardened `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup` so the synthetic persistent menu hover repeats until the popup session observes the root switch, avoiding a single stale/suppressed move as the whole oracle.
- Hardened Find result context-menu opening so the pointer path positions the cursor, sends a hover, and retries while waiting for the DxUI popup.
- Hardened Preferences Panes live toggle validation so it drives the left Status Bar toggle toward an expected checked state instead of assuming one Space key always lands after reopen.
- Reset local plugin copy/move concurrency before `Phase11_ConnectionOverridePrecedence`, so the test owns the baseline instead of inheriting the prior bridge phase's manual `copyMoveMaxConcurrency=4`.

Residual parity/perf/proof-depth items remain intentionally deferred to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`; they do not block this Done move.

---

## Closeout Update - 2026-06-18

**Current verdict:** almost ready for Done move, but do not move this plan yet. The Floodgate safety code is implemented, broad Commands ordering/state contamination is now fixed and green, residual non-blocking parity/perf/proof-depth work has been split to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`, and the only remaining closeout gate is one uninterrupted `.\Tools\Run-AllTests.ps1 -Suite Full` archived on a single machine.

**Pause/resume status - 2026-06-18:** closeout was paused by user request while the final Full gate was running. The run was interrupted during Commands after roughly 6,409 seconds, did not produce a final aggregate `run-all-tests-results.json`, and is not green evidence. Its diagnostic `last_run` snapshot was archived to `Specs/TestRuns/LT-PF5VDAGE/CloseoutInterrupted/2026-06-18_223549/`. All spawned Debug selftest/build processes from the interrupted run were stopped; the only remaining `RedSalamander.exe` observed was the user's Release app at `.build\x64\Release\RedSalamander.exe`.

**Latest resume evidence before the pause:**

- `.\build.ps1 -ProjectName RedSalamander`: passed, `0 warning(s), 0 error(s)`, log `.build/logs/msbuild-20260618_204230_399.log`.
- Compare exact runtime: `search_service_status_and_query_roundtrip` passed `1/0/0`; `Specs/Core/Core_Search.md` now records store-readiness versus request-specific fallback semantics.
- `.\build.ps1 -ProjectName DxUiTests`: passed, `0 warning(s), 0 error(s)`, log `.build/logs/msbuild-20260618_192437_226.log`; `.\.build\x64\Debug\DxUiTests.exe` passed.
- Commands exact runtime: `cmd_preferences_dialog_viewers_live_search_dx_interaction` passed `1/0/0` after bounded UIA/message-pump hardening.
- FileOps exact runtime: `Phase11_ConnectionOverridePrecedence` passed `3/0/0`.
- Broad Commands closeout runtime remains green: `cmd_pane_` passed `272/0/0`, archived at `Specs/TestRuns/LT-PF5VDAGE/Commands/2026-06-18_181213/`.

**Resume command:** rerun `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 3.0` with a long wrapper timeout. If green, archive the aggregate/per-suite artifacts, add final evidence here and in the Riptide plan, then move both plans to `Specs/Plans/Done/`. If red, triage the first failure before moving either plan.

**Implemented in this pass:**

- S3 planned-destination ancestor collision guard: an object-vs-prefix shape planned by the same transfer is rejected before the transfer deletes a destination object written earlier in that same plan.
- Cross-filesystem bridge MOVE integrity guard: source `GetSize` failure is no longer treated as "verification optional" for MOVE; source is preserved with `ERROR_PARTIAL_COPY`, and destination size is rechecked after promotion when needed.
- Curl staged-upload guard: staged remote uploads are re-statted before promotion/source-delete, with staged cleanup and `ERROR_PARTIAL_COPY` on size mismatch.
- Microsoft Drive recursive merge guard: invalid empty-name children mark enumeration incomplete so source folder deletion is not authorized; malformed `FileSystemOptions::sizeBytes` on Move is rejected.
- Quick Search refresh/focus guard: same-folder refresh preserves active query, and `WM_KILLFOCUS` exits search only when focus actually leaves the FolderView subtree.
- Test/source contracts and authoritative specs updated: `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`, and `Tools/Tests/TestInventory.Tests.ps1`.

**Green focused evidence:**

- `.\build.ps1 -ProjectName RedSalamander`: passed, `0 warning(s), 0 error(s)`, log `.build/logs/msbuild-20260618_113228_788.log`.
- `.\.build\x64\Debug\PluginContractTests.exe`: passed; provider debug selftests included FileSystem `25/0`, MicrosoftDrive `44/0`, S3 `109/0`, Curl `68/0`.
- `.\.build\x64\Debug\FileSystemCurlTests.exe`: passed.
- `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`: passed `30/0/0`.
- `Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru`: passed `5/0/0`.
- Quick Search focused runtime: `cmd_pane_quickSearch_integrated_navigation` passed `1/0/0`, archived at `Specs/TestRuns/7d3a1247382a/Commands/2026-06-18_113956/`.
- Floodgate bridge MOVE runtime: `Floodgate_CrossFsMoveGetSizeFailurePreservesSource` passed with setup/cleanup `3/0/0`, archived at `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-18_114202/`.
- Batch Rename focused rerun after the broad failure passed `46/0/0`, archived at `Specs/TestRuns/7d3a1247382a/Commands/2026-06-18_151252/`; representative isolated failures outside Batch Rename also passed (`Find`, `Shortcuts`, Compare Options).
- `.\build.ps1 -ProjectName RedSalamander`: passed, `0 warning(s), 0 error(s)`, log `.build/logs/msbuild-20260618_173823_556.log`.
- Exact Find editable-combo runtime: `cmd_pane_find_dialog_editable_combo_keyboard_editing_keys` passed `1/0/0`.
- Narrow Find ordering runtime: `cmd_pane_find_dialog_tab_traversal_matches_expected_order,cmd_pane_find_dialog_editable_combo_keyboard_editing_keys` passed `2/0/0`.
- Broad Commands closeout runtime: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_ -FailFast -TimeoutMultiplier 2.0` passed `272/0/0`, archived at `Specs/TestRuns/LT-PF5VDAGE/Commands/2026-06-18_181213/`.

**Broad Commands blocker resolution:**

- Earlier in closeout, `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2.0` built successfully (`.build/logs/msbuild-20260618_114511_570.log`, `1 warning(s), 0 error(s)`; the warning was `DxUiTests.Menu.cpp` C5245), but the wrapper timed out after 2 hours while Commands was still running. The orphaned Commands process later exited red and archived to `Specs/TestRuns/7d3a1247382a/Commands/2026-06-18_150527/`: `719 passed, 14 failed, 2 skipped`.
- The state-contamination failures from that run were narrowed through exact comma-separated sequence filtering, clipboard fallback hardening, Find destination-navigation hit-test/open-count instrumentation, Find tab traversal keyboard-state restoration, and editable-combo stale edit-host cleanup.
- The formal broad `cmd_pane_` rerun now passes `272/0/0`, so the prior broad Commands blocker is closed.

**Next closeout step:** run one uninterrupted `.\Tools\Run-AllTests.ps1 -Suite Full` with a timeout long enough for Commands to complete, archive the aggregate/per-suite artifacts under `Specs/TestRuns/<machine>/FileOps/<timestamp>/Floodgate_Final_Closeout/`, add the final evidence here and in the Riptide plan, then move both plans to `Specs/Plans/Done/` if green.

---

## Post-closeout adversarial review - 2026-06-19

Independent deep review of the closeout commit `68306b9ec` ("harden Floodgate safety closeout"). The four P0 data-safety fixes are sound and were re-verified in code (S3 guard pre-empts the self-delete on both COPY and MOVE before any write; cross-FS MOVE `GetSize`-fail hard-fail is the only safe option because `FileSystemBasicInformation` carries **no** file-size field, so the plan's old "use `sourceBasicInfo.sizeBytes` fallback" idea in FG-P0-2 was a misread and was correctly not implemented; MSDrive incomplete-enumeration propagates through every recursion level via the by-ref `anySkipped` + `subtreeFullyMovedOut`; Curl staged re-stat is correctly ordered before promote/source-delete). The review found one **critical test-credibility defect** and several reliability/perf/proof-depth gaps.

**Fixed in this pass (2026-06-19):**

- **[CRITICAL · test-credibility] FG-TEST-1 — the FG-P0-2 regression test never ran in the Full suite.** `Floodgate_CrossFsMoveGetSizeFailurePreservesSource` was registered in the `Step` enum, `StepToString`, and `kFileOpsPhaseOrder`, but **not** in any `kFileOpsFamily*` membership array. `BuildRunFiltersImpl` builds a full run exclusively from `kFileOpsFamilyDefinitions`, so the case was reported-but-silently-skipped in every `-Suite Full` / per-family run — it only "passed" via exact `--selftest-case=`. The plan's mandated final Full gate would therefore have given **zero** coverage of the single most important data-safety fix. This is the project's recurring "green self-test exercised a benign path" failure mode. **Fix:** added the step to `kFileOpsFamilyFairstream` (size 30→31); confirmed via `--selftest-case=FileOpsFamily_Fairstream --selftest-list-cases` that the case is now selected. The FG-P0-2 `[x]` runtime evidence above predates this fix and did **not** come from a family/Full run.
- **[MEDIUM · data-loss] FG-MSDRIVE-1 — MSDrive `ParseChildren` dropped non-object array elements without flagging enumeration incomplete** (`FileSystemMicrosoftDrive.cpp` ~2984), the same source-folder-delete loss class as the empty-name case the original fix targeted. **Fix:** the non-object skip path now sets `*incompleteDueToInvalidChildNameOut = true`, so a malformed child blocks the recursive source-delete and surfaces `ERROR_PARTIAL_COPY`. Provider selftests still `44/0`.
- **[MEDIUM · test-credibility] FG-TEST-2 — the source contract gave false confidence.** `TestHarnessSourceContracts` asserted only that the Floodgate case name appears *somewhere* in the file (true via enum/order array), never that it is a family member — which is exactly why FG-TEST-1 went unnoticed. **Fix:** added a tempered-dot assertion that the case appears **inside** the `kFileOpsFamilyFairstream` literal. Pester `30/0`.

**Evidence correction:** the `[x]` for FG-P0-2 below should read: the **source-`GetSize`-fail** branch is runtime-proven, but the **post-promote destination size-mismatch** branch (`State.cpp` ~7373-7402) has **zero** runtime coverage and no fault-injection hook (FG-A1). Treat FG-P0-2 as partially-proven until a destination-side injection test exists (tracked in the follow-ups plan).

**Validation after the FG-TEST-1 fix (2026-06-19):** with the case now registered in `kFileOpsFamilyFairstream`, the full Fairstream family was rerun through the harness (`.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOpsFamily_Fairstream -TimeoutMultiplier 2.0`) and reported **33 passed / 0 failed / 0 skipped — Overall PASSED** (exit 0). `Floodgate_CrossFsMoveGetSizeFailurePreservesSource` is now in the 33-case family inventory and ran-and-passed in family-sequence context (0 skipped) — so the source-`GetSize`-fail branch of FG-P0-2 is, for the first time, exercised by a family/suite run rather than only an exact `--selftest-case=`. The remaining formality is one uninterrupted whole-`-Suite Full` run; the other suites (Commands, etc.) are unaffected by these changes.

**Open items routed to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` (not blockers, but real):**

- **[HIGH · reliability] FG-P0-3-FALLBACK-SIZE-ZERO** — the Curl staged re-stat hard-requires `stagedInfo.sizeBytes == fileSize`, but FTP servers whose `LIST` output the parser cannot size (MLSD-only/EPLF/VMS/AS400/some IIS modes → `CurlListParseContext` fallback-names branch, `sizeBytes` always 0) make every **non-empty** copy/MOVE trip the guard, delete a correct staged upload, and fail with `ERROR_PARTIAL_COPY`. Fail-**safe** (source preserved, no data loss) but it **breaks copy/MOVE to an entire class of FTP servers**. The data-safety primitive must distinguish "size unknown" from "size 0" (add a `sizeKnown` flag to `Entry`, set only when the listing parsed a numeric size; on unknown, verify via a single-object size query — FTP `SIZE` / SFTP `fstat` — or gate the source-delete instead of failing).
- **[MEDIUM · reliability] FG-P0-2-RESTAT-FALSEPOS** — the post-promote re-stat (`State.cpp` ~7386) does a fresh by-path `GET`; on an eventually-consistent destination (e.g. MS Drive path lookup) a transient miss leaves a **duplicate** in both source and destination plus a spurious `ERROR_PARTIAL_COPY`. Distinguish "could not verify" (transient → short retry/backoff, don't strand a duplicate) from "verified wrong size" (true corruption → consider deleting the partial destination).
- **[MEDIUM · performance] FG-P0-3-EXTRA-LIST-ROUNDTRIP / FG-P0-2-PERF-RESTAT** — the Curl guard issues a **full destination-directory `LIST` per file** (O(N²) on a flat many-file move), and the bridge MOVE re-stat adds one metadata round-trip per file on **every** cross-FS MOVE. Use a single-object size probe; skip the re-stat for files the copy phase already byte-proved on strong-consistency destinations (aligns with FG-P2-7).
- **[MEDIUM · proof-depth] FG-A1 / FG-MSDRIVE-2 / FG-P0-2-TEST-SERIAL-ONLY** — add: (a) a destination-side `GetSize` fault-injection hook + Fairstream step covering the post-promote mismatch branch; (b) an MSDrive selftest where the empty-name child lives under a **nested** existing-on-both-sides subfolder (locks the recursion-propagation contract); (c) a concurrency>1 variant of the bridge GetSize-fail test (current case is serial-only).
- **[LOW · reliability/UX] FG-P0-1-A** — the S3 ancestor-collision guard correctly aborts the whole transfer (fail-closed, plan-sanctioned; do **not** switch to per-item-continue — that re-opens the self-delete loss) but surfaces a misleading `ERROR_ALREADY_EXISTS` and logs nothing. Use a distinct error and log the offending key for diagnosability.
- **[LOW · hygiene] FG-A3 / FG-A4 / FG-A5 / FG-QS-1 / FG-QS-2** — `Plugins_VirtualFileSystem.md` was tightened to a normative MUST (sizeBytes guards on Delete/Rename/batch) the code does **not** implement (only MSDrive Move); the commit bundles 6+ unrelated concerns under a "safety closeout" subject while this plan says the gate is **not** green; same-folder refresh during active Quick Search does not re-resolve focus/highlight (stale state if the matched item was deleted); and the new `focusStayedInView` `IsChild`/self-HWND branch is effectively dead (FolderView has no child HWNDs) so the OnKillFocus change is behaviorally a no-op vs the old null check.

---

## Historical merge-readiness verdict at reviewed merge point: NOT ship-ready

At the reviewed merge point, both Riptide documents explicitly said **"do not merge to master yet" / "do not move to Done"** and required one uninterrupted `Tools\Run-AllTests.ps1 -Suite Full` captured green on a single machine. **That gate was never recorded green** (last full `-SkipBuild` aggregate failed `960/3/46`; the only later attempt was user-interrupted and wrote no aggregate). The branch was merged regardless, with bare commit messages that neither acknowledged nor overrode the hold.

**Genuinely fixed — verified in code, do NOT re-litigate (see "Verified correct" below):** R0-1 (the critical size+mtime move data-loss), R2-2 staging-key entropy (S3 + Curl), R3-2 MOVE delete-gate, S3 backup ordering, R3-4 `AddHueWeight`, R3-1/R3-3 dead-code deletions, ConnectionManager UAF + `WM_NCDESTROY` teardown, R1-4/R1-5 concurrency clamp, R2-5 storage probe.

**Still shipped broken at that reviewed point:** four distinct silent-data-loss vectors in the cloud providers, one unfixed documented Quick Search blocker, and an ABI-guard hole — none covered by the "green" notes because their self-tests exercised benign paths, not the dangerous condition (the project's own standing lesson, recurring). Current code status is superseded by the reconciliation section below.

**Historical recommendation:** treat `0b6b8080a` (and any release built from it) as **not shippable for the file-ops area** until at least all P0s and FG-0 are closed. For current HEAD, use the closeout checklist and final-gate requirements below.

---

## Current HEAD Reconciliation - 2026-06-18

**Verdict:** the code is no longer in the original Floodgate "start fixing P0s" state. The main safety fixes have landed, focused validation is green, broad Commands is green, the final Full-suite aggregate is green, and lower-priority parity/perf/proof-depth work has a named follow-up plan.

| Slice | Current code status | Finish action |
|-------|---------------------|---------------|
| FG-P0-1 | Fixed. S3 now rejects a planned destination ancestor/self-collision before estimate or execution can delete a key written by the same transfer. | Keep the source-contract guard and S3 debug selftest; no further code work required for this closeout. |
| FG-P0-2 | Fixed and focused-green. Bridge MOVE treats source `GetSize` failure as `ERROR_PARTIAL_COPY`, preserves the source, and rechecks destination size when a size proof exists. | Runtime evidence: `Floodgate_CrossFsMoveGetSizeFailurePreservesSource` passed at `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-18_114202/`. |
| FG-P0-3 | Fixed for current closeout. Curl staged uploads are re-statted before promotion, and mismatched staged objects are deleted with `ERROR_PARTIAL_COPY`. | Source-contract + `FileSystemCurlTests.exe` coverage is accepted closeout proof for this release scope; true short-2xx fake-server fault injection is split to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| FG-P0-4 | Fixed. Microsoft Drive marks enumeration incomplete when malformed empty-name children are dropped, so recursive source deletion is not authorized from an incomplete child set. | Covered by MicrosoftDrive debug selftests and source contracts. |
| FG-P1-1 | Fixed and focused-green. Same-folder refresh preserves active Quick Search state; focus loss exits only when focus leaves the FolderView subtree. | Runtime evidence: `cmd_pane_quickSearch_integrated_navigation` passed at `Specs/TestRuns/7d3a1247382a/Commands/2026-06-18_113956/`. |
| FG-P1-3 | Deferred with explicit release decision. The source-delete helper remains guarded, and source contracts protect the helper shape, but the skipped-child behavioral selftest is still desirable proof-depth work. | Tracked in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`; not a blocker for the Floodgate/Riptide Done move after Full is green. |
| FG-P1-4 | Fixed. Temporary Quick Search diagnostics were removed from production source. | Source contract verifies the temporary traces stay out. |
| FG-P2-1 / FG-P2-2 | Deferred with explicit release decision. Microsoft Drive ambiguous transport reconciliation and delete-404-as-success semantics were not part of this pass. | Tracked in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| FG-P2-3 | Fixed. Microsoft Drive Move validates malformed `FileSystemOptions::sizeBytes` before copying the options tail. | Covered by MicrosoftDrive debug selftest and source contract. |
| FG-P2-4 | Partially closed and deferred. The Microsoft Drive Move hole is covered; broader malformed-options coverage across every Move/batch/provider path is not fully expanded here. | Tracked in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| FG-P2-5 / FG-P2-6 / FG-P2-7 | Deferred performance debt. No new perf implementation is in this pass. | Tracked in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`; any later performance claim needs archived before/after metrics. |
| FG-P2-9 | Deferred parity simplification. Serial and concurrent bridge-MOVE conflict classification have not been unified in this pass. | Tracked in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |

**Done evidence:** final Full gate passed on 2026-06-19 and is archived at `Specs/TestRuns/LT-PF5VDAGE/FileOps/2026-06-19_123255/Floodgate_Final_Closeout/`. Move this plan plus its dependent Riptide plan to `Specs/Plans/Done/`.

---

## Guiding Principle (inherited from Fairstream/Riptide)

> **Simple, always working, clear status.** Prefer removing options and special cases. Never a silent destructive act; never two conflicting states at once. **Move MUST NOT delete a source unless the destination is byte-proven written.** Prove it on the user's route.

## Performance / Safety Validation Contract (mandatory, per repo rules)

Every slice that touches destructive correctness (all P0s, P1-3) MUST include a **byte-integrity / tree-equality assertion** and an **out-of-tree sentinel assertion**, and MUST archive before/after runs under `Specs/TestRuns/<machine>/FileOps/<timestamp>/`. New selftest cases MUST be registered in `kFileOpsFamilyDefinitions` (order-array membership alone reports-but-never-runs — silently skipped). The host pre-flight only guards TOP-LEVEL conflicts; design P0/P1 selftests around **nested** children / cloud providers so they exercise the engine, not the host gate. **A slice is not done because a self-test is green — the proof must make the dangerous condition observable and FAIL on the unfixed code.**

---

## Implementation Tracking Checklist (update first, before editing code in a slice)

Use `[ ]` not started, `[~]` in progress, `[x]` complete, `[blocked]` needs a product decision.

| State | Slice | Phase | Implementation unit | Required proof before `[x]` |
|-------|-------|-------|---------------------|-----------------------------|
| [blocked] | FG-0 | Gate | Capture one uninterrupted `Run-AllTests.ps1 -Suite Full` after focused safety closure and broad Commands stabilization | Broad `cmd_pane_` is now green (`272/0/0` at `Specs/TestRuns/LT-PF5VDAGE/Commands/2026-06-18_181213/`); only the final Full aggregate remains. |
| [x] | FG-P0-1 | P0 | S3: never delete an "ancestor blocker" that is itself a planned write of this transfer | Implemented via planned-destination ancestor collision guard; covered by S3 debug selftest, provider contract run, and source contract. |
| [x] | FG-P0-2 | P0 | Cross-FS bridge: MOVE must byte-prove the destination before deleting the source even when `GetSize` fails | Implemented and runtime-focused green: `Floodgate_CrossFsMoveGetSizeFailurePreservesSource` at `Specs/TestRuns/7d3a1247382a/FileOps/2026-06-18_114202/`. |
| [x] | FG-P0-3 | P0 | Curl/WebDAV/SFTP: re-stat staged remote object before promote+source-delete | Implemented; current closeout accepts source-contract + `FileSystemCurlTests.exe` proof. True short-2xx fake-server proof depth is tracked in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| [x] | FG-P0-4 | P0 | MicrosoftDrive: do not drop empty-name children at the shared parser before a recursive source delete | Implemented via incomplete-enumeration propagation; covered by MicrosoftDrive debug selftest and source contract. |
| [x] | FG-P1-1 | P1 | Quick Search: gate `ExitIncrementalSearch()` on an actual location change | Implemented and runtime-focused green: `cmd_pane_quickSearch_integrated_navigation` at `Specs/TestRuns/7d3a1247382a/Commands/2026-06-18_113956/`. |
| [blocked] | FG-P1-2 | P1 | Release-hygiene: resolve the merge-against-hold (move plan to Done or revert; decision record) | The residual decisions are resolved by follow-up split; final Full gate must be green before Riptide/Floodgate plans move to `Done/`. |
| [~] | FG-P1-3 | P1 | Add bridge-MOVE skipped-child data-safety selftest | Deferred with release decision to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`; not a blocker for this plan after Full is green. |
| [x] | FG-P1-4 | P1 | Remove or document the temporary `ENABLE_TESTS` Quick Search diagnostics | Temporary diagnostics deleted; source contract verifies they stay out. |
| [~] | FG-P2-1 | P2 | MSDrive PATCH/DELETE: re-query-by-id reconciliation on ambiguous transport failure | Deferred to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| [~] | FG-P2-2 | P2 | MSDrive: retried DELETE seeing 404 reports success | Deferred to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| [x] | FG-P2-3 | P2 | MSDrive Move: add the `sizeBytes` ABI guard every other provider has | Implemented in `MoveSingleItemWithConflicts`; covered by MicrosoftDrive debug selftest and source contract. |
| [~] | FG-P2-4 | P2 | Extend the ABI `sizeBytes`-rejection selftest across Move/batch on every provider | Deferred to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`; MSDrive Move is covered in this pass. |
| [~] | FG-P2-5 | P2 | S3: stop double-probing destination objects; skip the wasted upfront pass on the interactive path | Deferred to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| [~] | FG-P2-6 | P2 | MSDrive merge: thread fetched child metadata into `MoveOrRenameItem`; add `$batch`/pacing for large merges | Deferred to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| [~] | FG-P2-7 | P2 | Cross-volume move: skip the deep destination re-read for files the copy phase just produced | Deferred to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| [~] | FG-P2-9 | P2 | Route both bridge-MOVE conflict classifications through one shared `ClassifyBridgeMoveConflict(...)` | Deferred to `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. |
| [ ] | FG-P3-* | P3 | Lower-severity correctness/simplification (see P3) | Per-item smoke / source-contract guard where applicable. |

---

## Phase P0 — Silent data loss (blocks ship; slices independent)

### FG-P0-1. `[HIGH · data-loss · NEW regression]` S3 transfer deletes an ancestor object it wrote itself

**Anchor:** `Plugins/FileSystemS3/FileSystemS3.Directory.cpp` ~1527-1568 (ancestor-conflict block in the transfer loop; the backup+`DeleteS3Object` at ~1561). Upfront probe: ~1399-1465.

**Symptom (silent permanent loss on MOVE):** S3 legally allows an object `src/a` and a prefix `src/a/` to coexist. The transfer writes `dest/a` first (lexicographic order), then when processing `dest/a/b`, `FindAncestorObjectConflict` discovers `dest/a` — **the object this same run just wrote** — as an "ancestor blocker," backs it up to a hidden `.rs-bak`, **deletes it**, and `CleanupBackupObjects` removes the backup on success. There is **no guard that the discovered ancestor key is itself a planned destination write of this transfer.** On MOVE the source `src/a` was already deleted, so the content is **permanently and silently lost; the op reports success.** Pre-Riptide this returned `ERROR_ALREADY_EXISTS` and aborted (safe); the Riptide diff converted a safe refusal into silent destruction. The added self-test only seeds *pre-existing* destination blockers, never the same-transfer-wrote-it shape — a synthetic green.

**Fix:** before treating an ancestor as a blocker, skip it if its key is itself a planned `destinationKey` of this transfer (check `destinationStates`/`plan.objects`, or maintain a `std::set` of keys written this run). On an intrinsic source-shape collision (object-vs-prefix within the same source), fail the item/transfer with a distinct surfaced error — never resolve by deleting a sibling we are responsible for writing.

**Required proof:** new `Floodgate_S3PrefixObjectAncestorOfSelfPreservesData` — source contains object `src/a` and `src/a/b`; MOVE to `dest/`. RED on current engine (after `dest/a` written, processing `dest/a/b` deletes it, op succeeds, `src/a` content lost). GREEN after fix (collision surfaced, no self-delete, source preserved on the failing item). Byte-equality + out-of-tree sentinel.

### FG-P0-2. `[P0 · data-loss · medium likelihood / catastrophic impact]` Cross-FS MOVE deletes source after an unverified cloud upload when `GetSize` fails

**Anchor:** `RedSalamander/FolderWindow.FileOperations.State.cpp` — `CopyFileWithBuffer` ~6900-6905 (`static_cast<void>(reader->GetSize(&fileTotalBytes));`), the only completeness check ~7270 (`if (fileTotalBytes > 0 && fileCompletedBytes != fileTotalBytes)`), promote/finalize ~7283-7347, MOVE delete decision via `ShouldDeleteMoveSourceAfterBridgeCopy` ~8104-8130 / 8638-8693. Confirmed by hand.

**Symptom:** The source `GetSize` HRESULT is discarded; the copy loop is EOF-driven; the byte-completeness check is gated on `fileTotalBytes > 0`. If a cloud reader's `GetSize` fails/returns transient (S3 `S3RangedFileReader::GetSize` can fail) **and** a transient short-but-2xx range read is surfaced as EOF, `CopyFileWithBuffer` returns S_OK with a truncated/empty destination, `fileTotalBytes` stays 0 so the size check is skipped, `ShouldDeleteMoveSourceAfterBridgeCopy` returns true, and the cloud source is deleted. Both faults live in the same network dependency, so co-occurrence is correlated. Violates the Guiding Principle.

**Fix:** for a MOVE, never treat `GetSize` failure as "skip verification." When `GetSize` succeeds, **always** require `fileCompletedBytes == fileTotalBytes`. When unsupported/failed, after Promote re-stat the destination (`GetFileBasicInformation`/`GetAttributes`) and compare against a re-read source size before allowing the source delete; otherwise return `ERROR_PARTIAL_COPY`. Cheaper fallback: use `sourceBasicInfo.sizeBytes` (already fetched ~6883) as the integrity target when `reader->GetSize` yields 0.

**Required proof:** new `Floodgate_CrossFsMoveGetSizeFailurePreservesSource` — bridge MOVE with a fault-injected reader whose `GetSize` fails and a short read that EOFs early; assert source preserved + op `ERROR_PARTIAL_COPY`. RED before fix (source deleted), GREEN after.

### FG-P0-3. `[P0 · data-loss · pre-existing, in reviewed path]` Curl/WebDAV/SFTP MOVE deletes remote source after an upload that is never byte-verified

**Anchor:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp` — `CurlUploadFromFile` ~1260-1271 (returns S_OK on `curl_easy_perform()==CURLE_OK`, `CURLOPT_FAILONERROR` only catches HTTP ≥400), `CopyFileViaTemp` promote ~2348-2375 (no `GetEntryInfo` re-stat), MoveItem/MoveItems delete + `FinalizeOverwriteTarget` ~3134-3161.

**Symptom:** A server/gateway/proxy that ACKs a short PUT, or an SFTP backend that accepts a short write, yields a truncated destination and a deleted source. `CURLOPT_INFILESIZE_LARGE` makes client-side truncation/connection-drop fail with `CURLE_PARTIAL_FILE`, so the residual hole is the **server-side-short-but-2xx** case — narrower than FG-P0-2 but real and the dominant remaining data-safety hole in this provider.

**Fix:** after `CurlUploadFromFile` returns S_OK and before `PromoteStagedFileToDestination`, `GetEntryInfo(destinationConn, stagedRemotePath)` and require `info.sizeBytes == fileSize`; on mismatch delete the staged file and return `ERROR_PARTIAL_COPY`. Only promote, and only delete the source, after the size proof. Where remote size is unreliable, gate the source-delete on verification or skip the staging optimization.

**Required proof:** new `Floodgate_CurlShortUploadPreservesSource` — stub remote that returns 2xx for a short PUT; MOVE → assert source preserved + `ERROR_PARTIAL_COPY`. RED before fix.

### FG-P0-4. `[P0 · data-loss]` MicrosoftDrive merge-move recursively deletes the source folder, destroying an unmoved empty-name child

**Anchor:** `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp` — `ParseChildren` empty-name `continue` ~2984-2996, `MergeMoveFolderIntoExisting` enumeration + `allChildrenMoved` ~3866-3871, recursive `DeleteItemById` of the source folder ~4029-4051.

**Symptom:** `ParseChildren` now `continue`s on any empty-`name` child (added to dodge a self-recursion stack overflow already independently neutralized by the depth cap ~3852). Because `ParseChildren` backs `ListDirectory`, which the merge uses to enumerate the source, an empty-name child is invisible to the merge loop, so `allChildrenMoved` stays true and the function issues an **unconditional recursive Graph `DELETE` of the whole source folder** (`DeleteItemById`, which Graph deletes recursively without requiring emptiness). The unmoved child is destroyed with no error. Trigger requires Graph to report a malformed/empty `name`, but the consequence is irreversible.

**Fix:** do not drop empty-name children at the shared parser layer. Either set `allChildrenMoved=false` whenever enumeration could be incomplete, or re-enumerate before the recursive `DeleteItemById` and bail to `ERROR_PARTIAL_COPY` if any child (including empty-name) remains. Keep the empty-name skip ONLY where the recursion hazard is, not in the shared parser.

**Required proof:** extend the Debug Graph selftests — `Floodgate_MsDriveEmptyNameChildBlocksSourceDelete`: fake child with empty `name`; MOVE-merge → assert source folder + child survive, op `ERROR_PARTIAL_COPY`. RED before fix.

---

## Phase P1 — Reliability & release-hygiene

### FG-P1-1. `[HIGH · correctness · pre-existing bug, misdirected fix]` Quick Search fix targets the wrong cause; documented case still fails the full run

**Anchor:** shipped change `RedSalamander/FolderView.Interaction.cpp` ~160-203 (`OnKillFocus` null-target guard + `kPaneRestoreFolderFocus` post). Root cause: `RedSalamander/FolderView.cpp` `SetFolderPath` ~220 and `FolderView.Enumeration.cpp` `ProcessEnumerationResult` ~1513 call `ExitIncrementalSearch()` **unconditionally**, before the same-location check.

**Symptom:** Printable typing clears `_incrementalSearch.active` while focus never leaves the same HWND. A same-folder refresh (live FS change → `DirectoryInfoCache` → debounced `RequestRefreshFromCache` → `ProcessEnumerationResult`) lands during typing and silently wipes the query and overlay. The null-focus guard provably cannot help (focus stays on the same window in the captured failing trace). `cmd_pane_quickSearch_integrated_navigation` is documented as STILL failing in the full ordered run.

**Fix:** gate the `ExitIncrementalSearch()` calls in `SetFolderPath` and `ProcessEnumerationResult` so a same-folder refresh **preserves** `_incrementalSearch` (clear only on an actual location change), mirroring the same-location logic already added to `NavigationView::SetPath`.

**Required proof:** the **full-ordering** `cmd_pane_quickSearch_integrated_navigation` passes (not just the isolated case), and a focused selftest where a same-folder refresh during active Quick Search keeps the query.

### FG-P1-2. `[HIGH · release-hygiene]` Branch merged against two explicit holds; no green Full gate exists

**Anchor:** `Specs/Plans/WIP/Operation_Riptide_*` (both still under `WIP/`); 12 added `Specs/TestRuns/SINON/.../README.md` are per-slice RED/GREEN evidence only — none is a Full-gate aggregate.

**Fix:** after the mandatory Floodgate safety/ABI slices are closed, run one uninterrupted `Tools\Run-AllTests.ps1 -Suite Full` on a single machine and archive it. If green: move the dependent Riptide plan and this Floodgate plan to `Done/`, resolve/re-spec any explicitly deferred blockers, and add a closeout note. If not green: hot-fix before any release or record a release-blocking decision. Replace the bare merge/checkpoint commit intent with an explicit decision record. (This is FG-0 split out as a tracked release item.)

### FG-P1-3. `[HIGH→MEDIUM · test-credibility on a data-loss gate]` No selftest covers the bridge-MOVE source-delete gate's data-saving branch

**Anchor:** `RedSalamander/SelfTest/FileOperations/...Fairstream.cpp` ~1449-1591, 2425-2629, 3916-4038. Gate: `ShouldDeleteMoveSourceAfterBridgeCopy` (`State.cpp` ~7983-7998).

**Symptom:** the skip/failure-injecting bridge tests are all `FILESYSTEM_COPY` (which never deletes a source); the only bridge MOVE test is happy-path. `Phase12_ReparsePointPolicy` covers the reparse-skip terms, but the `skippedFileConflictCount == 0` term and the copy-FAILURE arm have **no MOVE coverage**. The gate is correct today — this is one regression-step from loss.

**Fix:** add a bridge MOVE selftest: cross-FS MOVE of a directory whose child collides with a pre-existing destination file; Skip the child; assert (a) op ends `ERROR_PARTIAL_COPY`, (b) the skipped child's source survives byte-for-byte, (c) the transferred sibling's source was deleted, (d) prompt count == 1. Confirm RED if the `skippedFileConflictCount == 0` term is removed.

### FG-P1-4. `[HIGH→LOW · merge-hygiene · Debug-only]` Temporary Quick Search `ENABLE_TESTS` diagnostics merged against the "remove before closeout" gate

**Anchor:** `FolderView.Interaction.cpp` ~25 (`DescribeQuickSearchFocusTarget`), ~163, ~1191; `FolderWindow.FileSystem.Navigation.Part.cpp` ~260, ~272; `RedSalamander.cpp` ~10092. All `#ifdef ENABLE_TESTS`.

**Symptom:** four investigation-only trace blocks + a helper, listed in the crash doc under "Temporary Diagnostics … remove before final merge," merged still present and undocumented as intentional. They compile out of Release (zero runtime/binary impact) — pure cleanup debt.

**Fix:** delete the blocks now the investigation is over, or add a one-line "intentional permanent selftest aid" comment at each site and record the decision before moving the plan to Done.

---

## Phase P2 — Correctness / performance / ABI

- **FG-P2-1 `[MEDIUM · correctness]`** MSDrive PATCH/DELETE never do the documented re-query-by-id reconciliation after an **ambiguous transport** failure (`FileSystemMicrosoftDrive.cpp` ~3696-3742, 1937-2148). Only POST/create got a reconcile; PATCH/DELETE just pass `allowRetry=true` (status-loop only). A network drop after Graph commits a child PATCH but before the response triggers the overwrite rollback ~4185, colliding with the already-moved source → `ERROR_PARTIAL_COPY` + orphaned `.redsalamander-rollback-*`. Visible failure, not silent loss. **Fix:** wrap MoveItemById/DeleteItemById so a transport-level failure re-fetches by id; MoveItemById returns S_OK only if `parentReference.id` + `name` match; DeleteItemById treats 404/itemNotFound as S_OK. Add a transport-failure-injecting selftest.
- **FG-P2-2 `[MEDIUM · error-handling]`** MSDrive retried DELETE that sees 404 on retry is reported as failure (`FileSystemMicrosoftDrive.cpp` ~3696-3706, 4043-4055). If the first DELETE removed the item but returned a throttle status, the retry sees 404 → `ERROR_FILE_NOT_FOUND` for a materially-successful delete; in `MergeMoveFolderIntoExisting` the source-folder delete result is not NotFound-guarded (only the GET is), so it aborts after children already moved. Data ends in intended state; HRESULT wrong. **Fix:** return S_OK when `statusCode == 204 || statusCode == 404`. Add a selftest where the item is gone on the retry.
- **FG-P2-3 `[MEDIUM · ABI]`** MSDrive Move copies `FileSystemOptions` without the `sizeBytes` guard every other provider applies (`FileSystemMicrosoftDrive.cpp` ~5172-5177 `MoveSingleItemWithConflicts`, reached from MoveItem ~5256 / MoveItems ~5491). Local FS (8 sites), S3, Curl, Dummy all guard this copy; the header makes it normative. A version-skewed deployment reads 8 bytes past the caller's object and propagates garbage for `copyMoveMaxConcurrency`; within a matched build it silently accepts a malformed `sizeBytes`. The Riptide README's claim that size checks protect the new tail is **false for this path**. **Fix:** add `if (options && options->sizeBytes != sizeof(FileSystemOptions)) return E_INVALIDARG;` at the top of `MoveSingleItemWithConflicts` (propagates safely via `ReportItemResult`); consider symmetric validation in CopyItems.
- **FG-P2-4 `[LOW · test-credibility]`** The ABI `sizeBytes`-rejection selftest covers only `Dummy::CopyItem` (`...SelfTest.Phases07_09.cpp` ~4641-4720). MoveItem/MoveItems/CopyItems and all remote providers are untested, so the tripwire has no test that would fail on the unguarded MSDrive Move (FG-P2-3). **Fix:** extend Phase8 (or PluginContractTests) to drive malformed `FileSystemOptions` through Move/batch/Delete on every provider that accepts it; assert `E_INVALIDARG`; confirm RED on the unguarded MSDrive Move.
- **FG-P2-5 `[MEDIUM · performance]`** S3 directory transfer probes every destination object **twice** per object (`FileSystemS3.Directory.cpp` ~1399-1413 upfront `RefreshDestinationState`, then ~1453-1465 re-runs the identical probe and overwrites it). When `reportIssue` is set (the normal interactive path) the upfront fail-closed branch is skipped, so all N·(D+1) upfront HeadObjects are pure waste (~2× billed S3 metadata round-trips, serial). **Fix:** on the interactive path skip the upfront loop; in the transfer loop reuse the upfront `destinationStates[i]` on the first iteration; re-probe only after an ancestor-blocker removal `continue` or a Retry answer.
- **FG-P2-6 `[MEDIUM · performance]`** MSDrive merge GETs each child's destination metadata, then `MoveOrRenameItem` GETs the identical path again (the known `existingChild` is never threaded down) — each leaf pays ~5 serial Graph round-trips (≥1 a pure duplicate), no `$batch`, only reactive throttle backoff (`FileSystemMicrosoftDrive.cpp` ~3866-3911, re-probe ~4116). New-in-diff. **Fix:** thread the already-fetched destination (and source-child) metadata into `MoveOrRenameItem`; longer term, use Graph `$batch` + light proactive pacing for large merges.
- **FG-P2-7 `[MEDIUM · performance]`** Cross-volume move delete phase re-reads the **full source AND full destination** of every file before deleting (`FileSystem.FileOps.cpp` `FileContentsEqual` ~3171-3202, `DestinationMatchesSourceFile` ~3241-3392, call site ~5467) ≈ 3× copy byte I/O, destination read on the slower target volume — observable as the move "hanging" after 100% on multi-GB moves. (This is the cost of the correct R0-1 guarantee.) **Fix:** track files the copy phase just produced (per-file id set or verified byte count) and skip the deep re-read for them; restrict deep re-verify to the resume/uncertain case; or stream a source hash during copy to avoid the second full destination read.
- **FG-P2-9 `[MEDIUM · correctness/architecture]`** The concurrent (`_perItemMaxConcurrency>1`) bridge-MOVE path classifies conflicts at `State.cpp` ~8276 with hard-coded `(_operation, fileSystemIo, …, false)`, omitting the bulk path's (~8516-8850) selection: destination IO for a copy-phase failure, a `failedDuringMoveDelete` → `FILESYSTEM_DELETE` reclassification, and the `bridge.unsupportedDirectoryReparseEncountered` hint. Concrete bite: a nested-child unsupported reparse under concurrency>1 falls through to `Unknown` (one wasted Retry + wrong bucket); an `ACCESS_DENIED` source-delete loses the `ReplaceReadOnly` recovery action. **No data loss** (a FAILED copy never deletes the source; retry capped to one). **Fix:** capture the reparse hint + a `failedDuringMoveDelete` flag in the concurrent path and route both classification calls through a single shared `ClassifyBridgeMoveConflict(...)` helper.

---

## Phase P3 — Lower-severity correctness / simplification

- **FG-P3-8 `[MEDIUM · architecture · highest structural payoff]`** Four near-duplicate recursive copy walkers (serial vs parallel × plugin vs bridge: `FileSystem.FileOps.cpp` ~3946-4188 vs ~4246-4842; `State.cpp` ~7393-7541 vs ~7543-7903) hand-maintain identical destination/reparse/child-classification/one-shot-grant logic — this very diff had to apply the `oneShotAllowOverwrite`/`ClearOneShotGrants` change to **both** plugin walkers by hand. Also a real concurrency-dependent UX divergence: at concurrency 1 the serial walker re-prompts the parent on a nested `ERROR_PARTIAL_COPY`; at concurrency>1 it's silently folded into `hadSkipped`. (No current grant leak.) **Fix:** factor the per-child decision + conflict/grant retry into one shared routine parameterized by an enqueue/execute callback; at minimum normalize nested-`ERROR_PARTIAL_COPY` handling so both engines behave identically. **This is the structural source of future data-safety drift — prioritize it.**
- **FG-P3-1 `[LOW · data-loss · TOCTOU-only]`** Move source reparse-point deleted on bare destination existence, not link-equivalence (`FileSystem.FileOps.cpp` ~3299-3316). Delete runs only after a successful copy that just recreated the link, so it requires a concurrent external destination swap in the verify→delete window. **Fix:** verify same **tag** (allow the retargeted target — copies legitimately retarget symlinks), not an identical target buffer.
- **FG-P3-2 `[LOW · data-loss · inherent race]`** Verify-then-delete microgap in the move delete phase (`FileSystem.FileOps.cpp` ~3241-3392): byte-compare opens by-path sharing WRITE; delete uses a separate snapshot handle; a writer in the gap loses bytes. **Fix:** open the delete snapshot with `FILE_SHARE_READ` only (a writer then yields a sharing violation → skip → source preserved), or compare through a DELETE+`FILE_READ_DATA` snapshot handle.
- **FG-P3-3 `[LOW · correctness · TOCTOU + content-identical]`** Move delete-phase directory marker accepts a junction destination (`FileSystem.FileOps.cpp` ~3318-3325) — lacks the `!IsReparsePoint` guard the copy phase uses. Per-child decisions are byte-verified, so the destructive direction needs both a concurrent junction swap and byte-identical content. **Fix:** add `!IsReparsePoint(destinationAttributes)` at ~3321.
- **FG-P3-5 `[LOW · performance]`** Dynamic-job scheduler over-dispatches workers against a transiently single-item queue (`FileSystem.FileOps.cpp` ~419-476, 4755-4782): `queue.ready` decremented at pop not dispatch, so a deep-but-narrow tree at high concurrency wakes extra workers that find the deque empty and `notify_all` re-wakes everyone — bounded, self-terminating, no correctness impact. **Fix:** decrement the readiness signal at dispatch and re-increment on the transiently-empty Idle path.
- **FG-P3-6 `[LOW · error-handling · pre-existing]`** Main operation-thread creation in `StartOperation` can throw and orphan a registered task (`...State.Runtime.Part.cpp` ~700-715): the task is moved into `_tasks` before the `std::jthread` ctor, which can throw `std::system_error`; the queue then wedges and the popup never dismisses. R1-6 hardened the precalc thread but not this sibling. **Fix:** mirror `TryStartPreCalculationThread` — catch `std::system_error` around the ctor, remove the just-added task + tear down the popup, return failure HRESULT; let `std::bad_alloc` stay fatal.
- **FG-P3-7 `[LOW · performance]`** Non-local file-vs-directory child collision can still re-transfer a whole bridge tree on Overwrite (`State.cpp` ~6796-6815, 8276-8334, 8849-8908): `IsExistsOverwriteDeadEnd` probes with **local** `GetFileAttributesW` (→ `INVALID` for S3/OneDrive), so Overwrite is offered though it can't replace a folder with a file; the outer loop re-runs `bridge.CopyPath` from the tree root with no skip-committed semantics, re-uploading the whole committed subtree on each click. **Fix:** suppress Overwrite for the bridge file-vs-directory child (Skip/Cancel only); give the outer-loop bridge re-run resume/skip-committed semantics.
- **FG-P3-9 `[LOW · correctness]`** `OnKillFocus` ships a Release behavior change (`FolderView.Interaction.cpp` ~168-199): Quick Search now persists across a null-target `WM_KILLFOCUS` (Alt+Tab) and, when our process is still foreground, posts `kPaneRestoreFolderFocus` → `SetActiveWindow`/`SetForegroundWindow`. The cross-process branch correctly does not steal activation (verified). Residual: an in-process modal/popup activating during the window could be briefly yanked back. **Fix:** confirm via a deactivation test that the in-process re-assert can't fight a legitimately-opening modal; otherwise restrict the post to when the folder view's own window is still active, or document the intended persistence.
- **FG-P3-10 `[LOW · test-credibility]`** No re-run-COPY same-size/different-content resume selftest (`...Fairstream.cpp` ~1592-1765): `Fairstream_RerunCopyIsResumeAware` only changes a child's size. COPY and MOVE share the single `CopyFileInternal` chokepoint and the MOVE test does RED on a size-only regression, so the gap is narrow. **Fix:** add a same-byte-length/different-fill stage to the COPY resume test; verify RED when `PlainFilesEqualByContentNoFollow` is swapped for a size/mtime check.

---

## Verified correct — do NOT re-open without new evidence

- **R0-1 (core):** Move source delete is content-proven; **no size+mtime path remains**; a non-identical existing destination hard-fails `ERROR_ALREADY_EXISTS`. (`FileSystem.FileOps.cpp` ~3204-3268, 3379-3392, 3517-3544, 5463-5487). Re-confirmed by hand. **Keep the byte-compare; never reintroduce a metadata-only shortcut.**
- **R3-2:** `ShouldDeleteMoveSourceAfterBridgeCopy` is correct, used at both bridge MOVE paths, and re-checks `skippedFileConflictCount==0` at runtime (not just a debug assert). (`State.cpp` ~7983-7998, 8104, 8638).
- **R2-2 entropy:** S3 (`FileSystemS3.Directory.cpp` ~790-816) and Curl (`FileSystemCurl.CopyMove.cpp` ~813-892) both use 128-bit `BCryptGenRandom` tokens, no weak fallback, failure propagated, PID removed.
- **S3 move/overwrite ordering:** backup-before-overwrite, cleanup only after success, `RollbackTransfer` returns `ERROR_PARTIAL_COPY` before deleting destinations when a deleted source can't be restored; move-delete iterates only successfully-transferred objects. (`FileSystemS3.Directory.cpp` ~1585-1662).
- **R3-4 `AddHueWeight`:** `++weightCount` now inside the capacity guard; consumers were already clamped (latent counter-overflow, never OOB). (`Popup.cpp` ~1515-1519).
- **R3-1/R3-3 cleanup:** `clearCachedDecision` and the dead scheduler backpressure are gone; `REDSALAMANDER_SELFTEST_START_AT_CASE` is absent; all fault-injection hooks are `#ifdef ENABLE_TESTS`.
- **ConnectionManager ESCAPE UAF fix** is correct, and the diff adds a genuine **production** `WM_NCDESTROY` teardown fix in `WindowProc` (~3575-3580). The cited blocker is a functional Quick Search assertion (FG-P1-1), not an open AV.
- **R1-4/R1-5 concurrency clamp, R2-5 storage probe, scheduler CV/finalize/`inFlight` balance** independently verified correct.

## Refuted (do not re-flag without new evidence)

- "Move deletes a source junction on bare destination existence" as routine high-severity loss — reachability is narrow TOCTOU only (kept as FG-P3-1).
- `_transferStartedBeforePreCalcComplete`, `CollectTaskDiagnosticSnapshot`, `_taskFinished` as "dead" — all live.
- MSDrive cross-FS move guard / same-path no-op → `ERROR_INVALID_PARAMETER` "behavior flip" — host already screens it.
- "Live finished branch flashes a clean 100% Done frame (R2-6)" — transient ≤100ms self-correcting single-frame cosmetic; the skipped-child scenario is protected by `ERROR_PARTIAL_COPY`. Not worth a fix unless user-reported.
- Scheduler-shutdown data-race "cannot be confirmed RED-able" — refuted as a finding.

---

## Concrete closeout order from current code

1. **Do not move either plan to `Done/` yet.** Current focused safety code and broad Commands are green, but the uninterrupted Full-suite aggregate is still missing.
2. **Keep the follow-up split intact.** Deferred proof-depth, provider parity, and performance work is now owned by `Specs/Plans/WIP/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`; do not re-expand those items into this closeout unless a Full-gate failure proves they are blocking.
3. **Run the final gate.** Execute one uninterrupted `.\Tools\Run-AllTests.ps1 -Suite Full` with a timeout long enough for Commands to finish, archive the same-machine aggregate under `Specs/TestRuns/<machine>/FileOps/<timestamp>/Floodgate_Final_Closeout/`, and record the build log plus per-suite artifacts.
4. **Close only after green evidence.** Add the final closeout note here and in the Riptide plan, then move Floodgate plus the dependent Riptide plan to `Specs/Plans/Done/`.

## Out of scope / non-issues

- The 2A/2B live-backend merge-execution conformance gap is Fairstream's documented conscious limitation, not a Floodgate defect.
- The size+mtime *resume* idea is legitimate; only an ungated content-blind implementation would be a bug (already removed by R0-1).

---

*Provenance: 70-agent adversarial review of master `0b6b8080a` (last two commits), workflow `wf_fc21c0a7-a2c`, 2026-06-17. 39 findings survived two-skeptic verification; the two headline cloud data-loss vectors were re-traced by hand in the merged blobs. Anchors are `0b6b8080a`-relative — re-grep before editing.*
