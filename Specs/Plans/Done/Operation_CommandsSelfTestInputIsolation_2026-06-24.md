# Operation Commands SelfTest Input Isolation

**Status:** Complete — closed 2026-07-16 at `f4e0c8c3b`. CI5-R is complete across every current Commands real-cursor family: source contracts, the test-enabled build, and all five affected GUI cases are green. The broad Commands run finished 798 passed / 6 failed / 6 skipped; none of its six failures executes a modified warning/cursor path, so they are classified as unrelated to this plan and transferred to Observatory Track 0 rather than holding completed input-isolation work open.
**Date:** 2026-06-24
**Source plan:** `Specs/Plans/Done/Operation_Crosscut_CompareDirectoriesRemediation_SyncDataSafetyAndOptionsSimplification_2026-06-15.md`
**Primary scope:** `RedSalamander/SelfTest/Commands/*`, with implementation fixes only where a failing selftest proves a product-side UIA/focus/input defect.

## Problem Statement

The CompareDirectories Crosscut implementation is complete, but the broad Commands selftest gate is still red. The failures are not deterministic in focused runs:

- Latest broad Commands run: `.build/logs/run-commands-closeout-20260624_143239.out.log`
- Archived results: `Specs/TestRuns/7d3a1247382a/Commands/2026-06-24_150328/commands_results.json`
- Result: `751 passed / 11 failed / 2 skipped`, duration about 30m45s.
- The exact 11 failed cases passed when rerun in focused clusters.
- A wider plugin/connection-manager/credential replay still reproduced value-entry/focus instability (`41 passed / 5 failed`), including a default-name field reading `rce` instead of `New connection`.

Treat this as a Commands harness/input isolation problem until source evidence proves otherwise. Product fixes are allowed only when the harness demonstrates the application violates its UIA, focus, modal teardown, or input-routing contract.

## Current Failing Broad Cases (section removed 2026-07-02 folder review)

Superseded: all cases in the former failing-case table now pass in repeated broad-green Commands runs — `774 passed / 0 failed / 2 skipped` (`Specs/TestRuns/7d3a1247382a/Commands/2026-07-01_093435/`) and `776 passed / 0 failed / 2 skipped` (`Specs/TestRuns/7d3a1247382a/Commands/2026-07-02_121623/`). The failing-case table, focused-cluster green logs, and wider-replay evidence lists live in this file's git history.

## Non-negotiable Input-Safety Rule

Minimize direct mouse/keyboard paths. Prefer UIA providers, debug hooks, message-pumped helpers, and direct window messages where they exercise the same contract.

When a test genuinely requires real cursor movement, display a large on-screen warning before the movement and dismiss it immediately after:

```text
don't touch the mouse
```

Do not add new direct `SetCursorPos`, `SendInput`, or global keyboard paths without either:

- proving no UIA/message/debug route covers the behavior, and
- wrapping the real input window with the warning above.

## Working Hypotheses (section removed 2026-07-02 folder review)

Superseded: the hypotheses were resolved by the CI0-CI2 diagnostics/hardening/drain work and confirmed closed by the repeated broad-green runs cited above. The original five-hypothesis list lives in this file's git history.

## Implementation Tracking Checklist

Use `[ ]` not started, `[~]` in progress, `[x]` complete, `[blocked]` blocked.

| Status | ID | Priority | Work item | Required proof |
|---|---:|---|---|---|
| [x] | CI0 | P0 | Add a shared input/focus audit trace around Commands UIA value mutation, invoke, modal close, and menu popup transitions. | Done (2026-07-02 folder review): diagnostics landed — failure messages carry focus HWND/class/pane state, and the popup driver emits `AppendSuiteTrace` (`Commands.SelfTest.ViewCommands.cpp:9969-9993`); commit 3cb869a0d. |
| [x] | CI1 | P0 | Harden `RunUiaActionWithMessagePump`, `SetVisibleDescendantValueByNameWithMessagePump`, `InvokeVisibleDescendantByNameWithMessagePump`, and related helpers so they verify target window liveness, foreground/active ownership when required, and post-action state convergence. | Done (2026-07-02 folder review): helpers hardened with liveness guards + pumped waits (`Commands.SelfTest.Settings.cpp:5027-5082`, 56 call sites); proven by repeated broad-green runs 774/0/2 (`2026-07-01_093435`) and 776/0/2 (`2026-07-02_121623`). |
| [x] | CI2 | P0 | Drain stale input/messages after modal teardown and before reopen for credential, plugin configuration, connection manager, and Preferences cancel/reopen paths. | Done (2026-07-02 folder review): modal-teardown drain landed in commit 3cb869a0d; the `rce` symptom no longer reproduces in the broad-green runs cited in CI1. |
| [x] | CI3 | P1 | Strengthen Preferences page settle predicates for Plugins/Viewers so they wait for one visible pane, expected selection, and stable visible DxUi ValuePattern state. | Seven-case Preferences focused cluster remains green; full Commands no longer reports Preferences page settle failures. |
| [x] | CI4 | P1 | Harden menu popup tests around owner lookup, capture/focus cleanup, arrow-switch state, and cascading submenu return behavior. | Two-case menu cluster remains green; broad Commands no longer reports the arrow/submenu failures. |
| [x] | CI5-R | P0 | Reopened 2026-07-02 after two real cursor warps lost their warning/restore contract. Reconciled and completed 2026-07-16: the required all-site audit found two additional warning gaps in the Preferences category-tree churn probe and Find destination-history hover probe, plus duplicate warning-window implementations in Search and ViewCommands. The helper is now canonical in the Commands translation unit, all current real-cursor families use it, and restoration is armed before the first cursor warp. | Proven by `TestHarnessSourceContracts.Tests.ps1` 129/129, the test-enabled RedSalamander build with 0 warnings/errors, and the five affected Commands cases 5/5. |
| [x] | CI6 | P2 | Update `Specs/Testing/Testing_SelfTests.md` if this work establishes a durable Commands input-isolation contract. | Durable rule is documented outside this WIP plan before closeout. |
| [x] | Closeout | Gate | CI5-R focused proof is green. The 2026-07-16 broad Commands run produced six failures in unrelated Preferences page settling, Compare options, FolderView selection focus, and NavigationView popup focus paths; none uses a changed warning/cursor function. Per the maintainer convergence rule, Observatory Track 0 now owns their classification. | Complete: move this plan to `Specs/Plans/Done/`. |

## Progress Update - 2026-07-16

- Reconciled the plan against current `master` at `f4e0c8c3b`.
- Confirmed the durable warning/restore rule remains authoritative in `Specs/Testing/Testing_SelfTests.md`.
- Confirmed both previously unguarded hover functions now construct `DirectedSelfTestInputWarning` and restore the original cursor via `wil::scope_exit`; no new source edit is required.
- Focused two-case proof passed 2/2 on current `master`.
- The all-site source audit found and fixed two additional warning gaps: Preferences category-tree churn and Find destination-history hover. It also consolidated the duplicate Search/ViewCommands warning implementations into one Commands translation-unit helper and added source-contract coverage.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` passed 129/129.
- Test-enabled Debug x64 RedSalamander build passed with 0 warnings and 0 errors; log: `.build/logs/msbuild-20260716_174825_121.log`.
- The two original hover cases plus Preferences category churn, Find result-action/navigation, and Find command-enablement passed 5/5; run id `20260716T155106Z-74420-9c93904786cd472b9b7636cf0d295612`.
- The runner reported 563 pre-existing TestSandbox disk-audit issues. They are not input-isolation failures and are routed to Observatory Track 0 for ownership classification.
- Broad Commands completed with 798 passed / 6 failed / 6 skipped; run id `20260716T155228Z-41252-4f4ecd8fb7ab4123af0ee60f6a607100`.
- The six failures are outside CI5-R's modified source paths: three Preferences page/focus-settle cases, one Compare options body-lifetime case, one FolderView Select All focus case, and one NavigationView full-path-popup focus case. All are transferred to Observatory Track 0 for focused classification.
- The same run reported 563 TestSandbox disk-audit issues, also transferred to Observatory Track 0.
- CI5-R is fully proven and the Commands input-isolation plan is complete.

## Progress Update - 2026-06-28

Meaningful Commands issues from this pass are fixed, but this plan remains WIP because the explicit `Full` suite closeout gate has not been rerun after the latest FileOps/Compare fixes.

Completed or proven:

- Full Commands suite is green: `762 passed / 0 failed / 2 skipped`. Evidence: `Specs/TestRuns/7d3a1247382a/Commands/2026-06-28_140401/commands_results.json`.
- The remaining direct cursor paths are limited to two focused popup/hover probes, both guarded by an on-screen `DirectedSelfTestInputWarning`-style warning and cursor restoration. `rg "SetCursorPos|SendInput" RedSalamander/SelfTest/Commands` shows only those two `SetCursorPos` call sites and no `SendInput` call sites. (2026-07-02 folder review: this claim is no longer true — d46f506d0 and the MTP merge b824c118e brought the count to 8 sites with two unguarded warps; see CI5-R.)
- DxUi popup tests now scope lookup/dismissal to the owning window where possible and pump the UI thread while waiting for popup-driver threads to finish.
- Menu popup debug access is now routed through popup-window messages when called cross-thread, avoiding raw `GWLP_USERDATA` reads from the wrong thread.
- NavigationView edit-mode tests re-resolve and refocus the live edit host/input after refresh/display changes instead of reusing stale child HWNDs.
- Durable Commands input-isolation and owner-scoped popup rules were added to `Specs/Testing/Testing_SelfTests.md`.

Still required before moving this plan to `Specs/Plans/Done/`:

- Run `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2`.
- If Full is green, move this plan to Done.
- If Full fails for an unrelated subsystem, create or reference that subsystem's WIP plan and then close this plan only if the failure is not caused by Commands input isolation.

2026-07-02 folder review: the third bullet's rule was applied — the Suite Full gate's current blocker (`cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit`) is not a Commands input-isolation failure and is owned by `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md`, so Full-gate ownership transferred there. This plan now closes when CI5-R is fixed.

## Validation Commands

Build first:

```powershell
$env:RSBuildEnableTests='true'
.\build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander
```

Rerun the two CI5-R hover cases after adding the guards:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup,cmd_app_menuBar_persistent_view_to_files_hover_highlight_follows_pointer" -TimeoutMultiplier 2
```

(The superseded 47-case wider-replay CaseFilter and the three focused failure-cluster reruns were removed 2026-07-02 folder review; those clusters are covered by the broad-green runs cited in CI1, and the old command blocks live in this file's git history.)

Closeout:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
```

(The former `-Suite Full` closeout command is no longer this plan's gate — Full-gate ownership transferred to `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md`; see the Closeout row.)

## Exit Criteria (closed 2026-07-16)

- Met: every current Commands real-input family uses the canonical translation-unit `DirectedSelfTestInputWarning` and restores cursor state on all exits; the prior duplicate Search/ViewCommands warning implementations are removed.
- Met by focused proof and explicit transfer: all five affected GUI cases pass, while the six unrelated broad Commands failures are owned by Observatory Track 0. The former literal “Full Commands remains green” criterion is not claimed; it is superseded by the maintainer's convergence rule not to hold a completed change open on unrelated failures.
- Met already: wider replay, focused clusters, and full Commands green — broad-green runs 774/0/2 (`2026-07-01_093435`) and 776/0/2 (`2026-07-02_121623`).
- Met by transfer: the Full-suite gate is owned by `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md` (blocker `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit` is not input-isolation class), satisfying the original "own explicit WIP plan" clause.
- Met already: durable rules merged into `Specs/Testing/Testing_SelfTests.md` (CI6).
