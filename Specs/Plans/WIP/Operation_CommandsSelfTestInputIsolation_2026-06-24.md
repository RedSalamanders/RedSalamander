# Operation Commands SelfTest Input Isolation

**Status:** WIP - owns the broad Commands order/input failures split out of the CompareDirectories Crosscut closeout. 2026-07-02 folder review: the broad input-isolation work is done (CI0-CI4, CI6 complete, broad Commands repeatedly green); the only remaining engineering item is CI5-R (unguarded SetCursorPos regression, see checklist). Full-gate ownership transferred to `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md`.
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
| [ ] | CI5-R | P0 | Reopened 2026-07-02 (was CI5, previously [x]): the 2026-06-28 "only two SetCursorPos sites, both guarded" claim is no longer true. There are now 8 `SetCursorPos` sites in `RedSalamander/SelfTest/Commands`, and two REAL cursor warps are unguarded, violating the Non-negotiable Input-Safety Rule and `Specs/Testing/Testing_SelfTests.md:129-134`: `Commands.SelfTest.ViewCommands.cpp:9979` (`cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup` — no `DirectedSelfTestInputWarning` anywhere in the function AND no cursor restore) and `:10337` (`cmd_app_menuBar_persistent_view_to_files_hover_highlight_follows_pointer` — no warning, restore only). Introduced by d46f506d0 (2026-07-01) and the MTP merge b824c118e (partially reverting this plan's own CI5 cleanup). This is the plan's next action and its only remaining engineering item. | Fix: add the `DirectedSelfTestInputWarning` guard + cursor restore to both sites (pattern at `Commands.SelfTest.ViewCommands.cpp:10561-10563` and `Search.cpp:1048-1050`); then `rg "SetCursorPos\|SendInput"` in `RedSalamander/SelfTest/Commands` shows every real input path warning-guarded with cursor restore, and the two hover cases stay green. |
| [x] | CI6 | P2 | Update `Specs/Testing/Testing_SelfTests.md` if this work establishes a durable Commands input-isolation contract. | Durable rule is documented outside this WIP plan before closeout. |
| [~] | Closeout | Gate | Full-gate ownership transferred (2026-07-02 folder review): the Suite Full gate's current blocker (`cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit`) is owned by `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md`, per this plan's own close rule that an unrelated-subsystem failure is tracked by its own plan and does not hold this plan open. | This plan closes (moves to `Specs/Plans/Done/`) when CI5-R is fixed. |

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

## Exit Criteria (updated 2026-07-02 folder review)

- CI5-R fixed: both unguarded cursor warps (`Commands.SelfTest.ViewCommands.cpp:9979` and `:10337`) carry the `DirectedSelfTestInputWarning` guard + cursor restore, and `rg "SetCursorPos|SendInput" RedSalamander/SelfTest/Commands` shows no unguarded real input path.
- Full Commands remains green after the CI5-R fix.
- Met already: wider replay, focused clusters, and full Commands green — broad-green runs 774/0/2 (`2026-07-01_093435`) and 776/0/2 (`2026-07-02_121623`).
- Met by transfer: the Full-suite gate is owned by `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md` (blocker `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit` is not input-isolation class), satisfying the original "own explicit WIP plan" clause.
- Met already: durable rules merged into `Specs/Testing/Testing_SelfTests.md` (CI6).
