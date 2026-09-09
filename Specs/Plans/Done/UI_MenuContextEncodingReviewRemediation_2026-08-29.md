# Menu, context-action, and encoding review remediation

Status: COMPLETE
Priority: I18
Planned at: `4ad31825d3de59a4c4ec14a30772de58d7f751ee`
Drift check: `git diff 4ad31825d3de59a4c4ec14a30772de58d7f751ee..HEAD -- RedSalamander Plugins/ViewerSpace Plugins/ViewerText Tests Specs/UI Specs/Terminal Specs/Plugins Specs/SettingsStore.schema.json`

## Ownership boundary

This plan owned the bounded follow-up defects found after the completed I16 menu architecture work: FolderView context-menu truth and ordering, ViewerSpace host-menu synchronization, terminal session dispatch, the one-shot `Ctrl+Alt+T` migration, Find result-menu truth, and ViewerText encoding-catalog efficiency and dialog behavior. It did not reopen the accepted top-level menu tree, duplicate **Path from Other Pane** entries, satellite Batch Rename work, or File Operations product semantics owned by I3/I12/I14/I17.

## Outcomes and checklist

- [x] FolderView item/background choice is correct for pointer and keyboard invocation, including an empty pane.
- [x] Cut and Paste state reflects executable clipboard/provider behavior; target validation preserves duplicate multiplicity.
- [x] Artifact actions are inserted before Properties, with transactional separator/popup insertion.
- [x] ViewerSpace consumes the shared FolderView menu ID contract and gates Cut with Copy.
- [x] Terminal session-menu floating-window action reaches the same command as Commands > Terminal.
- [x] The legacy `Ctrl+Alt+T` migration is persisted and runs once, so later user rebinding is retained.
- [x] Find result context actions use one target model and expose truthful destination/provider enablement.
- [x] ViewerText uses an encoding-only resource catalog, avoids reparsing it on F8, shares encoding conversion helpers, and localizes picker rows/access keys.
- [x] Focused runtime/source-contract tests cover the corrected behavior, including picker current-row, Escape, rapid filtering, and teardown.
- [x] Authoritative UI, Terminal, settings, plugin, and performance contracts describe the durable behavior without overstating coverage.
- [x] Focused builds/tests, localization/resource validation, encoding performance evidence, and Fresh Full validation are green.

## Performance evidence

The accepted x64 Release evidence is archived at `Specs/TestRuns/4cb089111a23/Viewers/20260829_115300_menu_encoding_catalog_cache_release/`. Five fresh processes covered the full 140-row catalog. The cache was built exactly once per process; picker load p95 was 3,827 us, open-to-ready p95 was 19,474 us, filter p95 was 4,383 us, cache-build p95 was 3,841 us, and repeated F8 command p95 was 4,624 us. All established 50 ms open and 16.67 ms interaction budgets passed.

## Verification evidence

- x64 Debug and Release affected builds completed with zero warnings and zero errors.
- Focused FolderView/menu, settings/shortcut, Find, ViewerSpace, ViewerText, localization, settings-schema, and command-surface tests passed.
- The eight transient Commands UI cases from the superseded Fresh attempt each passed in isolated fresh-process reruns; the accepted Fresh run then passed all Commands cases without a flaky classification.
- `Tools/Get-SpecInventory.ps1 -FailOnFindings` reported zero findings before closeout and was rerun after archival/index reconciliation.
- Fresh Full run `20260829T110910Z-101328-4e2c00acfd5c463aa2284141a17d2a1d` passed with 2,012 total cases: 1,959 passed, 0 failed, and 53 permitted skips. It reported `flaky=0`, `regression=0`, `isolation_suspect=0`, and `unclassified_failure=0`. Validation evidence is under `.build/ValidationEvidence/runs/20260829T110910Z-101328-4e2c00acfd5c463aa2284141a17d2a1d/`.

## Done criteria

All checklist rows are complete; normative behavior is merged into authoritative specs; focused and Fresh Full validation pass; candidate performance evidence is archived; this file is in `Specs/Plans/Done/`; and the WIP index no longer lists I18.

## STOP conditions honored

- The accepted I16 top-level menu layout and deliberate duplicate **Path from Other Pane** entries were unchanged.
- No blocking provider capability query was added to the UI thread; unsupported capability expansion remains outside this bounded remediation.
- Localization edits were limited to structural catalog reduction, unambiguous access keys, and existing localized format resources.
- Unrelated I3, I12, I14, I15, I17, and D2 ownership/content was preserved.
