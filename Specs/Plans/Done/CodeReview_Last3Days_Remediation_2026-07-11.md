# Last Three Days Code Review Remediation - 2026-07-11

## Objective

Remediate every actionable finding from the review of commits
`188f40bbfa019b8ec538e1c1a77e618edbe40bd5..eaf640798eb0fa21f2b5b6e3883891ac2442a255`, from High through Low severity, with production fixes, deterministic regression coverage, authoritative specification updates, and performance evidence where the affected path is responsiveness-sensitive.

## Worktree preservation

The implementation started with unrelated uncommitted work in:

- `Common/DxUi/DxUi.Menu.cpp`
- `Common/DxUi/DxUi.h`
- `RedSalamander/RedSalamander.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_FolderWindow.md`
- `Tests/DxUiTests/DxUiTests.Menu.cpp`

Those changes are user-owned and must be preserved. Overlapping documentation updates must be merged without discarding them.

## Performance gate

- Scenario: File Operations popup pause/resume, pre-calculation aggregate state, compact progress rendering, settings toggles, and conflict metadata presentation.
- Scenario: foreground Search service readiness, query cutover state, request lifetime, and prompt process exit.
- Risk protected: indefinite worker/test waits, UI-thread persistence stalls, false determinate progress, false Ready/Watching status, and popup render/input regressions.
- Metrics: reuse `FileOps.Conflict.MetadataPromptUs`, `FileOps.Conflict.MetadataUs`, popup render/input metrics, Search service request/query/status metrics, and focused selftest case durations; add a bounded save-duration metric only if no existing metric covers popup setting persistence.
- Deterministic validation: focused Commands, FileOps, Compare, Pester source-contract, inventory, localization, and blocking-provider/lifetime cases introduced or strengthened by this plan.
- Evidence: create same-machine baseline/candidate archives under `Specs/TestRuns/<MachineHash>/...`, retain only small curated evidence in Git, and record exact analyzer commands and deltas before closeout.

## High

- [x] **CR-H1 — Eliminate File Operations pause/resume lost wakeups.** Mutate pause predicates under `_pauseMutex`, notify after unlocking, and prove real worker progress resumes for manual and queue pause.
- [x] **CR-H2 — Give foreground Search service tests explicit lifetime ownership.** Remove request-count coupling from readiness/termination, add an explicit graceful shutdown path or equivalent deterministic owner-controlled termination, and prove prompt exit after the final request.
- [x] **CR-H3 — Make UIA helper deadlines genuinely bounded.** A blocked UIA provider must return a failed assertion within the declared budget without an unbounded join; extend the fix to the related action/value helpers and add deterministic blocking coverage.
- [x] **CR-H4 — Repair the FileOps popup performance closeout.** Reopen the historical completion claim until same-machine baseline/candidate evidence, metric values, sample quality, and analyzer output are present and reviewable.

## Medium

- [x] **CR-M1 — Keep query cutover readiness honest for `CURRENTNESS_UNPROVEN`.** Separate legacy migration state from query-ready state and cover pre-query service store/sync status.
- [x] **CR-M2 — Preserve indeterminate aggregate progress during active pre-calculation.** Do not promote planned roots into known aggregate totals before calculation completes.
- [x] **CR-M3 — Derive Pause all / Resume all from command-eligible started tasks.** Cover the mixed paused-started plus pre-calculating-unstarted state.
- [x] **CR-M4 — Share one known-progress predicate between compact rendering and diagnostics.** Unknown rows must not render a determinate `0%` meter.
- [x] **CR-M5 — Remove synchronous settings/schema persistence from popup input handling.** Use serialized/debounced or shutdown persistence and retain bounded, crash-safe behavior.
- [x] **CR-M6 — Make FileOps pause-point tests report current blocked state.** Clear entered/blocked state on bailout and keep observer/bailout deadlines scaled consistently.
- [x] **CR-M7 — Isolate and split the monolithic popup Commands case.** Require clean setup/cleanup and ensure only case-owned tasks can satisfy assertions.
- [x] **CR-M8 — Repair TestRuns archive hygiene.** Remove raw oversized telemetry from the reviewed history/worktree where possible, correct machine-profile placement, and add size/profile lint.
- [x] **CR-M9 — Strengthen source-contract tests around decisive behavior.** Require expected-pipe matching and queue-entry/order predicates, with behavioral coverage where practical.
- [x] **CR-M10 — Restore real Preferences category-tree keyboard routing coverage.** Exercise HWND key-down/key-up behavior for Home/End/Up/Down while retaining direct control helpers only for model-level tests.

## Low and simplification

- [x] **CR-L1 — Make the File Operations non-default predicate exhaustive.** Include `issuesPaneSortDescending` and add one-field-at-a-time persistence coverage.
- [x] **CR-L2 — Reconcile README inventory counts and lint every duplicated count.** Generate or validate Commands and FileOps values as well as current covered families.
- [x] **CR-L3 — Localize conflict metadata timestamps.** Use the user locale and correct UTC-to-local conversion.
- [x] **CR-L4 — Avoid duplicate provider metadata round trips.** Use basic information first and attributes only as fallback, with call-count coverage.
- [x] **CR-L5 — Remove English-only popup assertions.** Assert structured state or localized resource output.
- [x] **CR-L6 — Document `popupFooterOnly` and `popupCompactDensity` in the authoritative SettingsStore specification.**
- [x] **CR-L7 — Restore weakened end-to-end routing/performance guards.** Attribute Preferences render counts and retain repeated Quick Search `WM_COMMAND` coverage.

## Validation and closeout

- [x] Build Debug `RedSalamander` with zero errors and no new warnings.
- [x] Run focused Commands, Compare, FileOps, Pester inventory/source-contract, and localization validations.
- [x] Run repeat/shuffle stress for pause/resume, popup isolation, Search service lifetime, and bounded UIA cases.
- [x] Run the full suite or document a concrete external blocker with preserved partial evidence.
- [x] Produce and analyze same-machine baseline/candidate performance evidence.
- [x] Update authoritative FileOps UI, Search, SettingsStore, selftest, performance, and TestRuns guidance.
- [x] Move this plan to `Specs/Plans/Done/` only when every item and the performance gate are complete.

## Closeout evidence

- Same-machine baseline/candidate metrics, sample-quality decisions, raw-artifact locations, and the keep decision are curated in `../../TestRuns/4cb089111a23/Commands/2026-07-11_130136/README.md`.
- The split popup successors passed 300/300 repetitions; final Search passed 35/35 runnable cases with four capability skips, and the exact readiness/cutover subset passed 5/5.
- The user-reported delete regression exposed a settings self-reload race: a popup save could be consumed as an external change, plugin refresh cleared both pane models, and retained folders were not re-enumerated. The atomic writer now publishes its exact finalized file stamp, reload epoch changes retry/repost, shutdown is bounded, and provider rediscovery restores raw provider paths and re-enumerates both panes.
- Final focused runs passed 4/4 once (`20260711T150508Z-71928-619bc6c6235f47ceb40d9d8528d23c01`) and 20/20 across five shuffled repetitions (`20260711T150534Z-93132-88f048b874144584b34962cddf32c806`). They cover exact stamp equality, external-write and failed-save races, bounded final process save, local panes, and a 7z pane with nonempty instance context.
- Final Debug x64 `Common` build `.build/logs/msbuild-20260711_170110_735.log` and isolated application build `.build/logs/msbuild-redsalamander-isolated-20260711_170148_410.log` completed with zero warnings / zero errors.
- The Full attempt stopped during the solution build at `.build/logs/msbuild-20260711_161736_710.log` on unrelated concurrent ViewerSqliteTests link gaps. A later dependency build also exposed an unrelated concurrent ViewerWeb lambda-capture error at `.build/logs/msbuild-20260711_164723_672.log`; both are classified explicitly in the curated evidence rather than calling Full green.
- Focused tooling gates passed: inventory 5/5, source contracts 132/132, documentation drift 9/9, archive policy 5/5, build-process guard 4/4, and localization 4/4.
