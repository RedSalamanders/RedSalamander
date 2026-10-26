# Lighthouse Track C — Windows session-end settings durability

Focused candidate evidence for Lighthouse `LH-7`, captured on 2026-07-15 from Debug x64 at the
`codex/lighthouse-session-end` working tree based on `6f61c38b3` plus the uncommitted Track C change.

## Scenario and result

- Scenario: `app/windows-session-end-settings-save`
- Commands filter: `cmd_app_session_end_`
- Run id: `20260715T084139Z-24412-aff83ed8f2e444dcbad9d511d7247008`
- Result: 1 passed, 0 failed, 0 skipped; no flaky, regression, isolation-suspect, or unclassified-failure classification.
- Build: `.build/logs/msbuild-20260715_103907_803.log`, Debug x64, 0 warnings and 0 errors.

The deterministic writer seam measured a 1 us `WM_QUERYENDSESSION` response and a 1,533 us runtime capture
plus injected writer call. It observed one writer call, zero normal teardown entries, and `S_OK`. These values
are characterization evidence for the new reliability path, not a claimed before/after speedup; the prior
implementation had no Windows session-end save scenario.

`focused_last_run/` preserves the focused runner manifest, Commands result/trace, and
`perf/session_end_settings_metrics.json`. The repository-level disk audit reported 236 pre-existing workspace
items during the focused run; the case and its perf artifact were green, and Track C does not claim a clean
workspace disk audit from this focused slice.

## Full-suite result

Full run `20260715T084345Z-73644-5868df5e04cd4cb5b00536ef772b6cb6` completed in 38m50s with 1,147 passed,
5 failed, and 54 skipped. The session-end case passed in the complete Commands run. FileOperations,
PluginContractTests, SettingsSchemaTests, CrashHandlingTests, PerformanceTests2 (14/14), and the inventory
contract also passed.

The five failures are outside Track C: one Preferences reverse-keyboard navigation case; the independently
known LocalizationTests and RedConfigureTests access violations; the existing Monitor ETW latency baseline;
and a Tools Pester source-contract `ArgumentNullException`. `full/` preserves the runner manifest, combined
results, complete streamed output, the session-end metric, and the three standalone-suite output files before
focused reruns can clean the raw TestSandbox run.

## Failure classification

- The exact Commands case `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation` passed
  1/1 in run `20260715T092433Z-65140-f5dd821388d24fa9828440da758eea18` (13.8 seconds).
- The complete `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` file passed 118/118 in 2.1 seconds; the
  previously failing payload-capture assertion passed in 91 ms.
- LocalizationTests and RedConfigureTests remain independently reproducible access violations, and Monitor ETW
  remains the established latency baseline. Their owner and prior reproduction evidence are recorded under
  `Specs/TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_b_full/`.

The two fresh Full failures are therefore classified as non-reproducing unrelated flakes, while the three
remaining failures are established external blockers. `classification/` preserves the Commands rerun result
and a concise classification ledger.
