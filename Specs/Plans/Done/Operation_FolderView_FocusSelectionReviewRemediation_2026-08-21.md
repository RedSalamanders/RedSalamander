# Operation FolderView Focus/Selection Review Remediation

> **NON-NORMATIVE OPERATION RECORD.** Authoritative behavior remains in the owning specifications under `Specs/`. This follow-up records implementation, test, documentation, and evidence work discovered after the completed focus/selection state-model operation.

## Status

- **State:** COMPLETE
- **Priority:** P1 correctness and destructive-interaction safety
- **Planned at:** `8313c1950e69f7ef1e5e057c5b0850381f410d33` plus the live uncommitted FolderView state-model worktree
- **Ownership:** FolderView context-menu targeting, pointer/drag cancellation, current-item repair integration, focused regression coverage, user documentation, and closeout evidence for the findings below.
- **Coordination:** I12 (`Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md`) retains broader FileOps removal-focus ownership. This operation may replace the newly introduced pre-sort index walk only by preserving I12's exact-per-source-`S_OK`, epoch, folder/provider/sort, immutable-target, and newer-enumeration proof gates and feeding successful proof into the canonical FolderView successor/predecessor resolver.

## Drift check

```powershell
git diff --name-status 8313c1950e69f7ef1e5e057c5b0850381f410d33..HEAD -- `
  RedSalamander/FolderView.h `
  RedSalamander/FolderView.Enumeration.cpp `
  RedSalamander/FolderView.Interaction.cpp `
  RedSalamander/FolderView.Menus.cpp `
  RedSalamander/FolderView.DragDrop.cpp `
  RedSalamander/SelfTest/Commands `
  Specs/UI/UI_FolderView.md `
  Specs/Testing/Testing_SelfTests.md `
  docs/UserGuide.md `
  docs/MainWindow.md `
  docs/KeyboardShortcuts.md
git diff --name-status
```

## Outcomes

1. A pointer background context menu cannot execute any item-only command against retained current.
2. Escape cancels every armed pre-drag gesture while retaining current and clearing selection.
3. FolderView releases mouse capture only when it owns capture.
4. Successful host-owned Delete/Move removal proof reuses the canonical resulting-UI-order successor/predecessor resolver rather than indexing a pre-sort enumeration payload.
5. No removal fallback can nominate an identity already proven removed.
6. Named deterministic tests prove the cross-component contracts promised by `Testing_SelfTests.md`.
7. User documentation accurately describes current versus selection, Space/Insert, empty-parent behavior, and navigation versus explicit selection restore.
8. Focused, performance, inventory, and Fresh Full gates pass on one stable source snapshot.

## Checklist

### Phase 0 — Plan and contract reconciliation

- [x] Preserve the completed normative focus/selection model as source of truth.
- [x] Confirm the findings against production code and named tests.
- [x] Add the explicit normative rule that Escape clears selection, retains current, resets anchor, and disarms an armed pointer/drag source.
- [x] Remove the legacy host-removal adjusted-index clauses and require canonical resulting-order repair with proven-removed exclusions.
- [x] Re-run retired-clause and specification-inventory checks after spec edits.

### Phase 1 — Production fixes

- [x] Require an item target for every item-only pointer context-menu command, including Open, in both menu enablement and dispatch validation.
- [x] Disarm potential drag on normal and Quick Search Escape handling.
- [x] Remove the unconditional `ReleaseCapture()` from `OnLButtonUp`; retain ownership-aware cleanup.
- [x] Refactor host removal tracking to return proof/eligibility rather than a replacement chosen from unsorted enumeration positions.
- [x] Feed eligible host removal into the same sorted successor/predecessor probe as generic disappearance while preserving the specialized proof gates.
- [x] Remove the fallback that can nominate a proven-removed identity.
- [x] Explicitly retire dead current-resolution reasons without changing their archived telemetry values.
- [x] Keep drag cleanup ownership clear; do not treat the already caller-disarmed `DoDragDrop` failure path as a production bug.

### Phase 2 — Deterministic coverage

- [x] Extend `folder_view_empty_background_keeps_current` to invoke disabled/background Open behavior and prove no retained-current activation.
- [x] Extend the same named case through pending removal tracking, exact `S_OK`, a newer accepted enumeration, and correct rehome without an ownership-epoch bump.
- [x] Extend `folder_view_drag_source_target_contract` with Escape while LMB is held and system-metric threshold boundary points.
- [x] Add a host-removal fixture whose provider enumeration order differs from the resulting keyed display order.
- [x] Add proven-removed fallback coverage.
- [x] Assign proven selected-rename-chain transfer to the existing `cmd_pane_navigation_directory_impact_preserves_selection_across_chained_renames` case in the normative matrix instead of duplicating that E2E scenario.
- [x] Run the exact focused cases repeatedly and in their normal Commands-family order.

### Phase 3 — User documentation

- [x] Document current item and selection as independent; arrows move current; background click and Escape clear selection while retaining current.
- [x] Document Space as toggle, post-toggle size-work update, advance without wrap; document Insert as toggle and advance without size work.
- [x] Document the empty-folder Go to parent action as non-item state and exclude filter-empty.
- [x] Document Back/Forward current restoration without selection restoration and explicit Save/Restore Selection behavior.
- [x] Keep root `README.md` unchanged because no concrete contradiction was found.

### Phase 4 — Validation and closeout

- [x] Build the test-enabled x64 Debug solution with zero new warnings/errors.
- [x] Pass focused Commands and FileOps cases on the changed behaviors.
- [x] Run the existing `folderView_perf_focus_selection_state` scenario and archive/validate candidate evidence; compare with the accepted same-machine candidate without inventing a budget.
- [x] Pass `Get-SpecInventory.ps1 -FailOnFindings`, tooling/source-contract tests affected by the spec move, and `git diff --check`.
- [x] Pass `Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` on the final stable snapshot.
- [x] Record exact run IDs, receipts, archives, results, and caveats here.
- [x] Move this plan to `Specs/Plans/Done/` and remove it from the WIP index only after every required gate is green.

## Verification commands

```powershell
.\build.ps1 -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Affected -ImpactBase 8313c1950e69f7ef1e5e057c5b0850381f410d33
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh
git diff --check
```

Exact Commands case invocations and the governed performance command must be copied from the current selftest/performance specifications rather than guessed.

## Done criteria

- Production behavior conforms to the normative FolderView contract for every finding.
- Each repaired path has a deterministic regression that fails on the reviewed implementation and passes after the fix.
- User documentation no longer describes Space or Insert as unconditional selection.
- Existing performance direction is preserved with auditable candidate evidence.
- The final Fresh Full repository verdict is PASSED on the same final source snapshot recorded here.

## STOP conditions

- Exact removal proof would need to be weakened or host intent would be allowed to retarget an immutable destructive command.
- Canonical resulting-order repair cannot be reused without duplicating the sort comparator.
- A required interactive test cannot observe the actual menu/drag/removal boundary and no equivalent deterministic seam exists.
- A dirty-worktree overlap cannot be preserved safely.

## Closeout evidence

### Build and focused correctness

- Test-enabled x64 Debug build passed with zero warnings and zero errors. Final repair/rebuild receipt before Fresh Full: `259a037aa9a248e893135d2428a17b8ee037193456a26dbc67340533578f12ba`.
- The exact FolderView context-menu, empty-background/removal-focus, drag-source/threshold/Escape, keyed-sort host-removal, selection/filter, LRU/navigation, and Space/Insert cases passed in focused runs and in normal Commands-family order. The keyed-sort/removal fixture passed 10/10 after its case-sensitive ordinal fixture order was corrected.
- The diagnostic Affected campaign `20260821T070844Z-78932-7274caaa92374eb9b40eee269c4792e6` exposed two stale source contracts and that fixture-order dependency; both were corrected. Its unrelated Preferences settle failure passed 3/3 in isolation. The diagnostic campaign is not used as the final repository verdict.
- Affected Pester source contracts passed 175/175 after correction.

### Performance evidence

- Accepted baseline: `Specs/TestRuns/4cb089111a23/Commands/2026-08-20_154800_folder_view_focus_selection_baseline_release/`.
- Accepted prior candidate: `Specs/TestRuns/4cb089111a23/Commands/2026-08-20_183441_folder_view_focus_selection_final_candidate_release/`.
- Repeated remediation candidate: `Specs/TestRuns/4cb089111a23/Commands/2026-08-21_103616_folder_view_focus_selection_review_remediation_release_repeat3/` from governed run `20260821T083548Z-80844-ad2e6191887b4374ac9b22cde8951cd0`, Release receipt `2e4baa0ad87984873b5116e8d7483df2cf0852f6fd6412a5668af62ee30e81b3`, source snapshot `9b2ec2b86516ececf9b3ddbca8ea1aa3011cad161dfe011d487ab3ca47da3a8e`.
- Repeated scenario result: 3 passed, 0 failed, 0 skipped. Focus resolver had 634 samples and p95 `4328 us`, versus `4406 us` for the prior accepted same-machine candidate (`-1.8%`, noise). The sparse FolderView preset reported 0 regressions, 3 improvements, and 2 noise classifications with its explicit one-sample sparse-row rule. No hard budget was added.
- Checked compact JSONL: 641,749 bytes, 1,754 rows, SHA-256 `B7E7545455DFF0843F50197C15D19934FB783166792A73DDA2E829274B403356`. Retained local full raw JSONL: 11,691,107 bytes, 31,864 rows, SHA-256 `C7C00130AC88DBBC75098D1A8E05D95DC7A30CE31CA4F50131A9124D35A46478` under `.build/PerfRawArchive/2026-08-21_103616_folder_view_focus_selection_review_remediation_release_repeat3/`.

### Final repository validation

- Authoritative Fresh Full campaign: `20260821T091337Z-84588-b241f42170254e8c8161389aae1207a6`.
- Result: PASSED; all 18 governed entries promoted on stable snapshot `3c8b91d71b3fed9b526fa7c5ece4cd09a986692a93756547bca290fee0c34557` with Debug build receipt `aea58ef340b32743b34a1ec394e3fff95663f22cf71545d449309418c9a563d6` and plan digest `b996b39bea73fde9e70405fa819091a3288b53bbd323c49602eb673261094e7c`.
- Major results: tooling preflight 2 passed / 0 failed; tooling Pester 634 passed / 0 failed; Compare Directories 227 passed / 0 failed / 31 declared-capability skipped; Commands 865 passed / 0 failed / 2 skipped; File Operations 118 passed / 0 failed / 20 skipped; DxUi completed all 399 reported groups with 9 declared interactive-desktop skips and no failures; PerformanceTests2 14/14 passed. FileSystemCurl, ViewerPE, ViewerSqlite, Monitor, Localization, RedConfigure, PluginContract, SettingsSchema, CrashHandling, Monitor ETW latency, and the vcpkg synthetic entry all passed.
- The earlier diagnostic Fresh campaign `20260821T084259Z-48204-d3ecc265629c44709cf11ee2d64370ae` was intentionally interrupted after a one-off idle PluginContract process. The identical PluginContract binary then passed directly in 19.9 seconds, the required contamination repair rebuild passed, and the authoritative Fresh campaign above passed PluginContract in 19.3 seconds. The interrupted campaign is not green evidence.
- `Get-SpecInventory.ps1 -FailOnFindings` passed with 0 blocking findings; `Test-TestRunArchive.ps1 -Inventory` passed for 541 files; archive validation passed for all 12 files in the new repeated candidate; `git diff --check` passed (line-ending notices only).
