# Operation Chordforge closeout continuation — 2026-08-12

## Purpose

This checkpoint preserves the exact state needed to resume Operation Chordforge
closeout. The implementation, tests, normative specifications, user
documentation, mockups, and performance archive are present in the current
unstaged, uncommitted worktree. No branch, stash, commit, stage, or destructive
Git operation was created.

The WIP plan remains active at
`Specs/Plans/WIP/Operation_Chordforge_TerminalCommandExperience_2026-08-11.md`.
Do not move it to `Specs/Plans/Done/` until the two full-suite failures below
are reconciled and every remaining closeout gate is green.

## Green evidence captured in this session

- All eight archived `Show-PerfRuns.ps1 -FailOnQuality` analyzers passed against
  `Specs/TestRuns/4cb089111a23/Commands/20260811_192050_operation_chordforge_terminal_command_experience_release/`.
- Full Debug x64 solution rebuild passed with 0 warnings and 0 errors. Log:
  `.build/logs/msbuild-20260812_084412_823.log`.
- Terminal source contracts passed 27/27.
- `PluginContractTests.exe` passed; Terminal debug selftests passed 56/56 and
  the Terminal unload quiet point passed.
- `SettingsSchemaTests.exe` passed, including singleton floating-window
  placement and ordered tab/profile/path persistence.
- `DxUiTests.exe` passed.
- Focused `terminal_shortcut_` Commands passed 10/10. Runner:
  `20260812T065107Z-10160-8919ba1965f54dc3b09a51705d655eaa`.
- Ordinary production Release x64 build, with `RSBuildEnableTests` absent,
  passed with 0 errors. Log: `.build/logs/msbuild-20260812_085128_320.log`.
  It reported three unrelated existing C5245 unused-function warnings.
- The build contamination marker is clear.

## Full-suite result that still blocks closeout

Command:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full
```

Runner `20260812T065351Z-74552-6c75e0bca16346eba46902ea7be2e3af`
returned exit code 1 after 2,088,591 ms. Aggregate: 490 passed, 712 failed,
53 skipped. Fifteen suites passed; two were classified
`UNCLASSIFIED_FAILURE`:

1. `Commands`: 132 passed, 711 failed, 2 skipped. The first real failure was
   `pane_view_options_preview_uses_configured_embedded_viewer_and_preserves_focus`
   after 75,582 ms: `Failed to focus image preview item after media preview.`
   The main window then became invalid, causing most of the remaining 710
   failures to cascade at zero duration. A few theme-overlay model checks also
   reported substantive failures after the window loss. This first failing case
   is unrelated to the Terminal shortcut implementation until a focused repro
   proves otherwise.
2. `ToolsPesterTests`: the aggregate Pester process returned exit code 1 after
   577,181 ms. The full runner did not retain a per-test failure list in its
   structured result, so this suite must be rerun directly to expose the exact
   failed test(s).

The long File Operations phase was investigated under the crash-forensics
workflow because the GUI process appeared non-responsive. It was CPU-bound,
not crashed: there was no new dump, it advanced through
`Phase11_BridgePipelineDummyToDummyPerf`, completed cleanup, and ended with
`FileOpsSelfTest: PASS`. File Operations finished 118 passed, 0 failed,
20 skipped. Do not treat it as a blocker.

The volatile full-run structured result is currently at:

`Z:/src/RedSalamander/.build/TestSandbox/runs/20260812T065351Z-74552-6c75e0bca16346eba46902ea7be2e3af/artifacts/selftest/last_run/run-all-tests-results.json`

A future runner may remove that sandbox run, so the decision-relevant evidence
is reproduced above.

## Exact resume sequence

1. Confirm `git status --short` still shows the saved dirty Operation Chordforge
   worktree and `.build/artifact-operation-contaminated.json` is absent. Do not
   reset, checkout, clean, or overwrite the worktree.
2. Reproduce the first Commands failure in isolation:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Debug -SkipBuild `
     -CaseFilter pane_view_options_preview_uses_configured_embedded_viewer_and_preserves_focus `
     -TimeoutMultiplier 4
   ```

   Preserve the focused trace before another selftest. If it passes, rerun the
   complete Commands family with `-ClassifyFailures`; if it fails, diagnose the
   preview-focus/window-lifetime root cause before interpreting downstream
   cases.
3. Expose the exact tooling failures:

   ```powershell
   Import-Module Pester
   Invoke-Pester -Script .\Tools\Tests -PassThru
   ```

   Fix or reconcile every failure, then rerun the affected Pester file and the
   complete tooling suite.
4. After Commands and tooling are independently green, rerun:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Full -ClassifyFailures
   ```

5. Run the remaining closeout gates:

   ```powershell
   .\Tools\Get-SpecInventory.ps1 -FailOnFindings
   .\Tools\Test-TestRunArchive.ps1 -Inventory
   git diff --check
   ```

6. Perform the requirement-by-requirement audit listed in the WIP plan. Record
   final run IDs and results, change the plan state to complete, move it to
   `Specs/Plans/Done/`, and rerun the spec/archive inventories after the move.

## Worktree scope at pause

`git diff --stat` reported 73 tracked files changed, 7,747 insertions, and 818
deletions, plus the expected new Terminal/host implementation files, durable
mockups under `Specs/Assets/`, the Release performance archive, and this
continuation checkpoint. The deletions under `Specs/Plans/WIP/Assets/` are the
old mockup locations; their durable replacements are under `Specs/Assets/`.
