# Terminal Gate 0 full-suite continuation and classified baseline exceptions

- Frozen review commit: `3e84e87b717841e28e11e1fe7e5bc04b4e3a3e0d`
- Canonical command: `.\Tools\Run-AllTests.ps1 -Suite Full`
- Wrapper start: `2026-07-31T15:38:29Z`
- Commands harness start: `2026-07-31T15:46:08Z`
- Last Commands progress: `2026-07-31T15:59:23Z`
- Live capture: `2026-07-31T16:17:45Z`
- Hung process: `RedSalamander.exe`, PID `35952`
- Last-started case:
  `cmd_preferences_dialog_plugins_live_search_dx_interaction`
- Wrapper completion: `2026-07-31T16:56:01Z`
- Wrapper exit: `1`
- Aggregate: 776 total, 722 passed, 3 failed suite/case entries, 51 skipped

The Commands harness recorded 371 passes and no failures before starting the
case above. It then made no trace or result progress for more than 15 minutes.
The process remained responsive, accumulated CPU time, retained about 11,075
handles, and did not produce a crash artifact. `Tools/Run-AllTests.ps1` waits
for the native self-test process with an unbounded `WaitForExit()`, so there was
no runner watchdog that could complete this invocation. After preserving the
live trace, partial results, process command lines, process/thread snapshots,
and wrapper prefix, the operator stopped only PID `35952`; the canonical
wrapper then continued to its remaining lanes. `commands-native-trace.txt`
pinpoints the last completed UIA work: the case found the visible `Search`
field, found the shell `Cancel` action, and never returned from that synchronous
interaction. A full dump attempt through `comsvcs.dll,MiniDump` was denied by
the local process policy.

This is classified as a pre-existing Preferences/UIA/D3D self-test boundary,
not a Terminal Gate-0 regression:

- `base-pre-freeze-commands-results.json` and
  `base-pre-freeze-commands-trace.txt` preserve an earlier invocation that
  began while HEAD was still the Gate-0 base
  `1a5d838d0c18298d9e560e243e67b821a40fcf6a`: it also recorded 371 passes,
  started this exact case, and produced no outer `Case returned` record;
- the same exact case and hang boundary were previously reproduced and archived
  at
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-21_1832_commands_plugins_uia_hang/`;
- the already-completed
  `Specs/Plans/Done/UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md`
  records that prior owner classification and keeps Preferences remediation
  outside unrelated feature scope;
- an earlier green full-suite archive at
  `Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`
  records the exact case passing in 640 ms. That proves the interaction has
  completed historically, but does not by itself prove flakiness on the current
  base;
- the frozen Gate-0 commit changes governance, specifications, schemas, and
  Gate-0 PowerShell tooling only. It changes no product, native self-test, or
  full-suite harness source, and the target test function is byte-identical
  across the archived July-12 revision, the Gate-0 base, and the frozen review
  commit.

The wrapper completed every remaining lane after the exact Commands child was
stopped. Compare Directories passed 227 and skipped 31; File Operations passed
112 and skipped 20; all 18 File Operations families completed. DxUiTests and
the other ordinary native/managed/script suite entries passed. The aggregate
has 17 suite entries: 14 passed and three are classified below. Therefore the
nonzero wrapper result is retained honestly; it is not described as a green
full-suite pass.

The unchanged aggregate also records
`test_sandbox_audit.is_clean=false` with nine
`unexpected-test-root-child` issues. Those issues are not suite/case entries,
but they are part of the full-suite result and required explicit closeout. All
nine named TerminalEngine fixture directories had creation/last-write times
between `2026-07-30T22:18:12Z` and `2026-07-30T23:01:46Z`, before both final
focused runs and this full suite. Neither later run modified them, and an exact
prefix search of frozen commit
`3e84e87b717841e28e11e1fe7e5bc04b4e3a3e0d` found no source match. They were
therefore earlier Round-3/ad-hoc residue, not a leak reproduced by the frozen
tests.

The operator moved exactly those nine disposable children recoverably into an
ignored quarantine outside `TestSandbox`, then repeated the exact 20-file
focused Gate-0 set under short isolated run `r4clean-284e0752`. It started
`2026-07-31T17:30:43Z`, completed `2026-07-31T17:41:25Z`, and passed all 239
tests with zero failed, skipped, pending, or inconclusive tests in 639.63
seconds. The disk audit executed in the same process after Pester and returned
`SandboxClean=true`, `SandboxIssueCount=0`; the root contained only `runs`, and
only the canonical full-suite run plus the new explicitly allowed run were
present. `sandbox-remediation-focused.stdout.txt` has SHA-256
`b253a566b3bcfee8a3f2878c1e6d06e01a077ae57fb19a2fc87e6cb5c4d6873c`;
`sandbox-remediation-focused.stderr.txt` contains only Git line-ending warnings
and has SHA-256
`d2d2470773f41a8075b387bd73d02148cbf72f0c2fce554e85321cee8aa10250`.
`test-sandbox-remediation.md` preserves the bounded remediation record.

The three failed aggregate entries are:

1. **Commands / `selftest_result_coverage`.** This is the expected aggregate
   consequence of externally stopping the unbounded child at the known
   Preferences Plugins UIA hang. The child had 371 passes and zero ordinary
   case failures; every later case is reported missing because the process
   could not return to the coordinator. Owner classification:
   Preferences/DxUi self-test infrastructure.
2. **ViewerPETests.** The exact underlying failure is
   `ViewerText shows one localized warning after truncating an over-cap hex
   clipboard selection`, which also makes the enclosing fresh-harness assertion
   fail. `base-pre-freeze-viewerpe.output.log` preserves the same exact failure
   from the earlier base-HEAD continuation. The frozen Gate-0 change touches no
   Viewer product or test source. Owner classification: ViewerText.
3. **ToolsPesterTests.** Pester itself completed 584 passes and three unrelated
   baseline failures: a documentation-drift test still reads the already-moved
   File Operations popup plan from WIP; a historical-review banner assertion
   expects an ASCII hyphen while the unchanged reviewed file uses an em dash;
   and the unchanged test-inventory contract expects 725 Commands registrations
   while the unchanged source inventory derives 714. The two Pester test files
   and all named inputs have no diff from Gate-0 base
   `1a5d838d0c18298d9e560e243e67b821a40fcf6a` to frozen review
   `3e84e87b717841e28e11e1fe7e5bc04b4e3a3e0d`. Owner classification:
   Testing/documentation baseline maintenance.

The exact aggregate, case history, native logs, process/trace capture, and final
wrapper streams are retained here. `SHA256SUMS.txt` binds each archived file,
including the accepted clean remediation rerun. The Terminal engine decision
records this classified result together with the separate 239/239 focused
Gate-0 pass, zero-issue post-remediation sandbox audit, and 12/12
dependency-lifecycle pass.
