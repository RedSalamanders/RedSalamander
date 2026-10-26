# I30 delivered review and implementation history

Status: DONE — historical review and committed implementation only, archived 2026-10-10 at `ab5f40b57`. This is not production qualification. Thirty commits remain outside origin/master. Outstanding implementation, verification and owner actions are retained in [the remaining plan](../WIP/CodeReview_MasterMerges_2026-10-05.md).

The following text is frozen historical evidence; its old verdicts and remaining-work statements describe their original dates, not current status.
## How the review was done

- Reviewed a clean sparse checkout of `d93dea1c6`, not the working tree (the working tree carried 234 uncommitted
  changes of the scoped-testing work at review time).
- 27 independent reviewers: 21 slices covering every changed production file plus tests and tooling, and 6
  cross-cutting lenses (end-to-end data safety, lifetime/threading, performance, architecture/simplification,
  user experience/accessibility, spec conformance).
- Findings were deduplicated, then adversarially verified. High findings had three skeptics (trace, context,
  impact), Medium two, Low/Info one.
- Result: 180 distinct findings, 6 refuted. Of the 174 survivors: 0 Critical, 7 High, 40 Medium,
  80 Low, 47 Info (simplification/architecture). There were also 148 owner questions and 12 coverage gaps.
- The seven High findings and `fs-fileops#2` were then re-checked by hand against the snapshot (next section).
  All other entries are agent-verified only.
- Static review only: nothing was built or run.

## Revalidation against master `076d72cd3` (2026-10-06)

Master moved after the review: scoped testing (PR #114 `bcf1a8348`, PR #116 `3847e5bdd`) and a mockup npm bump
(PR #115). Every surviving finding was remapped from `d93dea1c6` to `076d72cd3` with a line-level diff of each
cited file. **All 174 surviving findings remain valid; none was fixed or invalidated.**

- **Production code:** the only changes are renamed test-header includes and clang-format reflow, in files
  such as `State.cpp`, `Popup.cpp/.h`, `FolderView.*.cpp`, `RedSalamander.cpp`, `ViewerPluginManager.cpp`,
  `IconCache.cpp` and `SettingsStore.cpp`. Behavior is unchanged.
- **Unchanged files:** `Plugins/FileSystem/`, `FileOperationItemResultReducer.h`, `RedSalamander/Ui/`,
  `vcpkg.json`, `vcpkg-tool.json`, `.github/dependabot.yml`, `Tools/Modules/Testing/TestSandbox.psm1`, and the
  File Operations and popup specs.
- **Anchors:** 158 findings keep the same file and line, and 15 moved only because their test file was renamed
  (`*.SelfTest.*` to `*.Tests.*`, same line numbers). The ledger and this plan now use the new paths. The one
  edited anchor (`tests-other#9`, line 1 header comment) still describes the same behavior: the runtime
  source-text check had to be re-pointed at the renamed files, which is the fragility it reports.
- **`tests-fileops#1` exposure increased.** The new `Test-Changes.ps1` runs `Run-AllTests.ps1` for every scoped
  local iteration, and `Run-AllTests.ps1:1277` runs `Remove-RSTestSandboxStaleRunDirectories` unconditionally.
  Any iteration after a failed or killed F7 run now sweeps the retained mount tree. Wave 0 item 2 stays first.
- **`x-spec#1` and question 18 got worse.** The duplicate I28 (Release dependency recovery and Scoped testing)
  is now committed in the WIP index, in addition to the I27 collision.
- **Not reviewed:** the scoped-testing changes themselves (`Test-Changes.ps1`, `ScopedTesting.psm1`, the
  `Run-AllTests`/`TestRunPlan` changes and the test-file renames). They fall after the original window.
- **Workflow note:** under the new AGENTS.md scoped-testing policy, iterate on each wave with
  `Test-Changes.ps1` and use `-Mode PrePush` before pushing. The Waves 0-2 release gate still needs the I29
  Fresh Full outcomes.

## Verdict

**Not production-ready as merged.**

- **Master cannot install its dependencies on a clean machine.** A Dependabot PR merged without CI rolled the
  vcpkg baseline back.
- **One data-loss path.** It is pre-existing, but it is reachable through the newly hardened Recycle flow:
  Recycle can permanently destroy items and report them as recycled.
- **Wrong outcomes and wrong decision targets in the new File Operations UI:**
  - Cancel labels untouched items "Outcome unknown".
  - Escape can answer another task's conflict.
  - The popup drops clicks while copying.
  - One graphics failure permanently degrades the popup.
- **A self-test can leave a whole-volume mount** where the next run's sweep deletes recursively.
- **Qualification is still open.** I29 records that Fresh Full Debug/Release on the API-3 DxUi pin has not
  passed.

What holds up after tracing:
- The core conflict-consent design: generation-bound prompts, one-shot acceptance under the arbiter lock, and
  `SetEvent` inside the lock.
- Cancellation wake-ups.
- The ENABLE_TESTS gating of the FileSystem plugin's fault-injection hooks (off in shipping Release).
- The SettingsStore merge resolution (formatting only).
- The DxUi rename adoption.

## Hand verification of the High findings

| ID | Finding | My check | Severity | Introduced |
|---|---|---|---|---|
| `build-tooling#1` | `builtin-baseline` rolled back to `9e593bb` by Dependabot #111; master cannot restore | Confirmed. [vcpkg.json:5](../../../vcpkg.json) is `9e593bb`; [vcpkg-tool.json](../../../vcpkg-tool.json) is `d53f25d`. Commit `c0597174c` had documented 11 missing version-database entries at `9e593bb`, and `fbe113a0f` had aligned both pins at `d53f25d`. PRs #111-#113 merged while CI could not start (billing). | High | `441e1de91` |
| `fs-fileops#1` | Recycle can silently permanently delete and report "recycled" | Confirmed. Both recycle routes use `FOF_NOCONFIRMATION` without `FOF_WANTNUKEWARNING` ([FileSystem.FileOps.cpp:7328](../../../Plugins/FileSystem/FileSystem.FileOps.cpp), 7489). Neither `PreDeleteItem` checks `TSF_DELETE_RECYCLE_IF_POSSIBLE`, and both `PostDeleteItem` sinks treat any `SUCCEEDED(hrDelete)` as "committed, original removed". The flags are unchanged since before the range. | High | pre-existing |
| `x-datasafety#1` | Cancel (or any per-item failure, which cancels) reports every never-dispatched root as Canceled with Unknown publication/source, so the task summary says "Outcome unknown" | Confirmed. `MarkSourceItemsMutationPossible()` ([State.cpp:10399](../../../RedSalamander/FolderWindow.FileOperations.State.cpp), body 11239) marks every builder before execution. Finalization (11400) passes that phase to the reducer, whose cancel branch ([FileOperationItemResultReducer.h:181](../../../RedSalamander/FileOperationItemResultReducer.h)) sets Unknown axes. Per-root `executionPhase` exists only for roots that reached `Process()`. | High | `b883d5f8e` |
| `popup-b#2` | Escape on task A's focused path region submits task B's conflict Escape action (e.g. Cancel) | Confirmed. In `SyncHostedControls` ([Popup.cpp:7150-7178](../../../RedSalamander/FolderWindow.FileOperations.Popup.cpp)), only a focused button sets `focusedTaskPresent`. A focused path region falls through to "the sole actionable decision", which is spec-forbidden. | High | `b859cbc37` |
| `popup-b#1` | The hosted control tree is rebuilt whenever the in-flight stream count changes, so clicks on Pause/Cancel/decision buttons are lost | Confirmed. One `Source.{i}` region per in-flight row (Popup.cpp:9496). The rebuild condition compares tooltip-descriptor counts (6703), and a rebuild calls `ResetInteractionState()` (6783), dropping capture/pressed/hover/focus. | High | `b859cbc37`, `7a47d8b77` |
| `popup-a#1` | One transient `PrimeForShow` failure (e.g. `ResizeBuffers` after a device reset) switches the popup to the native failure surface for the life of the HWND; conflict decisions become unreachable | Confirmed. `_failureSurface` is only ever set true (Popup.cpp:11860). DxUi `49a9988` treats a `ResizeBuffers` failure as "deferring recreation to a later visible render". The spec's "Close and reopen re-attempts the host" ([UI_FileOperationsPopup.md:638](../../UI/UI_FileOperationsPopup.md)) is not implemented. Low likelihood but severe consequence, hence High. | High | `34b37cccf` |
| `tests-fileops#1` | The F7 admission baseline mounts the entire alternate (D:) volume inside a run tree and deliberately retains the mount on failure; the next run's stale sweep runs `Remove-Item -Recurse -Force` on that tree | Confirmed. The mount is retained by `retainMount` ([AdmissionBaseline.cpp:187](../../../RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.AdmissionBaseline.cpp)); crash or kill also leaves it in place. The sweep ([TestSandbox.psm1:495](../../../Tools/Modules/Testing/TestSandbox.psm1)) never checks descendant reparse points. Whether PowerShell's recursion descends into a *volume* mount point (rather than unlinking it like a junction) must be proven on a VM before relying on it. At minimum the retained evidence is destroyed and the mount orphaned. Test-only, opt-in (alternate-volume flag), but worst case is developer data loss. | High | `02841dd4b` |

The Medium `fs-fileops#2` was also confirmed by hand. Native Move computes `allowOverwriteEffective =
HasOverwriteGrant(context)` ([FileSystem.FileOps.cpp:6451](../../../Plugins/FileSystem/FileSystem.FileOps.cpp))
without the `destinationExisted &&` guard that Copy has (4248). An absent-at-probe destination therefore gets an
unbound `MOVEFILE_REPLACE_EXISTING`. Pre-existing; race-only; one-line fix. I would raise it to Medium-High
because it is a data-loss vector.

## Remediation progress (2026-10-08)

Work happens on branch `claude/i30-review-remediation`, one commit per item, each reviewed by an independent
reviewer before it is committed (a fix-up runs when the review asks for changes). The orchestration paused when the
agents' weekly usage limit was reached; it resumes from the next item.

| Commit | Item | Ledger IDs / decision |
|---|---|---|
| `21b2e26d4` | Review, plan and findings ledger | - |
| `6c4a4db42` | Restore the vcpkg pins Dependabot rolled back | build-tooling#1-#3, Q11 |
| `b531e1dcc` | Initialize build.ps1's overlay flag for StrictMode callers | regression of `6c4a4db42` |
| `0113de0aa` | Sandbox cleanup never follows links or mount points | tests-fileops#1 (cleanup) |
| `fc222f854` | F7 uses a test-owned virtual disk, never a real volume | tests-fileops#1, Q12 |
| `d85927d0d` | Confine test links, shares and SUBST drives to their sandbox | Q12 (rule) |
| `32202ee19` | Format only a proven-blank test VHDX; keep it after FileOps fails | Q12 follow-ups |
| `9a7abe5a1` | Report NOT_EVALUATED, not a crash, when -NonInteractive skips all tests | runner defect found in this run |
| `07f3167a9` | Refuse to package or release a test-enabled build | Q16 |
| `f67bf98c5` | Recycle never deletes permanently without asking | fs-fileops#1 (High) |
| `e55a68fb5` | Move never replaces a destination that appeared after the probe | fs-fileops#2 |
| `30cd32412` | Recycle escalation deletes a non-empty folder when accepted | state-c#2, Q3 |
| `29495bbaf` | Structural loopback-share contract; skip logs and formatting | regressions found in this run |
| `5ae046033` | Self-tests never leave recycled items in the user's Recycle Bin | Q12 (pre-existing) |
| `251c042d2` | Never delete a rejected MSIX through a link or outside the lease | review Lows of `07f3167a9` |
| `5cccbaa90` | Clear the Release-only warning C5245 | pre-existing |
| `a83359175` | Clear completed keeps interrupted-Move notices and unseen results | popup-c#2, Q8 |
| `260c9189a` | Never-dispatched per-item roots are retained, not "Outcome unknown" | x-datasafety#1 (High), Q1 |
| `f290374e9` | Empty-folder Move merges complete; vanished sources fail | x-datasafety#3, fs-fileops#3, Q19 |
| `e8fa0be57` | Batch Rename rows a stopped run never reached are not attempted | state-c#1 |
| `c91fb6837` | Name a canceled or error-stopped Copy/Move instead of reporting success | x-ux#3 |
| `1ed166a1f` | Format FileSystem.Internal.h for the whole-file clang-format gate | CI formatting gate |

**Remaining, in order:** Wave 2 items 2-4 (`popup-b#2` Escape focus, `runtime-queue#1`/`l10n#2` retarget,
`host#2`), Q14 (refuse new work during a pending exit), Q18 (plan IDs), the Wave 2 Low leftovers, the Q2 task-scoped
escalation with the pane-Delete continue-on-error decision, Q17 (regex warning), Waves 3-10, then a build and test
checkpoint that gates the simplification track, then a final integration check. The detailed follow-ups and the
owner actions found while checking each item are in
[remediation-followups.md](CodeReview_MasterMerges_2026-10-05/remediation-followups.md).


## Additional delivered commits

ab5f40b57 Stop the Issues pane misordering rows and showing facts they lack
f520fd9e1 Stop a held Recycle ignoring Cancel after another item fails
45779b0df Stop the Issues pane swapping a producer's sentence for generic text
ff888db98 Stop Apply to all flooding the Issues pane and evicting real errors
bc7728fd5 Stop the Issues pane showing English engine text as the problem
6442a5b91 Stop a conflict inspection failing on refresh or hanging on cancel
ff2773d2b Stop a forged window message running code in the alert overlay
5e581d8c0 Stop alert content jumping on option clicks and announce focus moves
69caa9dbb End a prompt when its owner dies and stop alerts unhooking each other
cb62fc40e Drop a stale popup Escape decision and stop re-fitting the notice
56ced1a1c Stop pane Delete at a failed escalation and let held Recycles finish
6c45c6e0c Stop the popup leaving scroll panels pointing at removed controls
0ec8da3c4 Slow the popup and owner-window timers while copies wait
584b08204 Keep popup clicks and focus intact while a copy's streams change
a0615dc35 Retry failed popup frames and recover from the failure surface
77b01b0ee Abbreviate a long refused name in the Destination unchanged notice
614980c39 Ask before closing a Compare window and count its tasks at app exit
d395e4fbe Fix refusal notice race, double alert, pending exit and queue place
5442c95f5 Return focus to the owning task when a path detail's row vanishes
c2d4c4a47 Bound a pane Delete's Recycle resume to one extra provider call
cd623ed10 Warn when a search regex gives up instead of dropping items silently
3d902c06b Offer Delete all remaining permanently when Recycle fails
d57a2c1bf Explain queued destination change refusals and follow the new overlap
67f7076c8 Report Skip, Keep Both and partial Move results as they happened
c085f76c8 Reject reused plan IDs in the spec inventory and renumber the clashes
12854e334 Refuse new file operations while an accepted exit is pending
54f3b7046 Report a refused pane Copy/Move/Delete instead of dropping it silently
9fa079c3c Report a refused queued destination change as keeping the destination
e7fd11571 Never bind Escape to another task's decision from a focused path region
ff5d6b0d1 Stop Dependabot from proposing any vcpkg port or baseline update
