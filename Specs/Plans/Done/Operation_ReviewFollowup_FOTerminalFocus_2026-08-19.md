# Operation Review Follow-up — FO completion, Terminal focus/input, SwapPanes

> **NON-NORMATIVE COMPLETION RECORD.** Durable behavior belongs in `Specs/<Domain>/`.
> This file records the 2026-08-19 review follow-up. Do **not** move
> `Specs/Plans/Done/Operation_TerminalHostFileOpsDefectFixes_2026-08-19.md`
> back to WIP.

## Status

- **State:** COMPLETED
- **Priority:** P1 correctness (FO completion, preview focus, VT suppression, copy routing); P2 (SwapPanes tabs, command-surface reaper, removal-focus snapshot)
- **Planned at:** `150864345147c167808cc309d9c1ae955e666304`
- **Ownership boundary:** review findings agreed after the 15-day Terminal/host/FO defect batch. Preserve the delivered MOVE identity contract and I13's removal-focus resolver; do not weaken cleanup policy. The September 8 owner request adds the reproducible blank embedded terminal and disabled first floating-terminal opener to this bounded follow-up, followed by a visible warning for tests that require foreground input.
- **Drift check:** `git diff 150864345147c167808cc309d9c1ae955e666304 -- RedSalamander/FolderWindow.FileOperations.State.Queue.cpp RedSalamander/FolderWindow.Layout.cpp RedSalamander/FolderWindow.FileSystem.Navigation.cpp RedSalamander/FolderView.Enumeration.cpp Plugins/Terminal Specs/FileSystem/FileSystem_FileOperations.md Specs/UI/UI_FolderWindow.md Specs/UI/UI_FolderView.md Specs/Terminal/Terminal_EmbeddedPlugin.md`

## Completion record

| Field | Value |
|---|---|
| Updated | **2026-09-08** |
| Current item | **Complete — Fresh Full passed; all 35 entries promoted** |
| Tests run | Corrected test-enabled Release: 9 passed / 0 failed / 0 skipped, run `20260908T163756Z-28496-52f8387c7b26486cb0c3fd2d3b69e1cb`; build 0 warnings/errors. Live Alt+7, Ctrl+Alt+T and menu prompt checks passed. The predecessor-inclusive 46-case replay repeated three times passed 138/0/0 (`20260908T182930Z-97660-586ba5c9e75145f5b8839707c4bb9b6c`). Earlier Debug host subset: 6/0/0; linked Pester: 298/0/0. Failed experiments are retained in the closeout archive. |
| Next | None within this bounded plan. New work follows the authoritative domain specs and its own active owner. |

## Qualified outcomes

1. FO `PostCompleted` dual-post failure never applies/removes from `Task::ThreadMain` or a threadpool worker; completion stays on the UI thread, and shutdown has a race-free lifetime barrier even if callback submission fails.
2. Preview content tab restores source FolderView focus; Folder/Terminal tabs focus the host preferred target; Terminal tab commits `SetActivePane`.
3. Ghostty encode arms translated-character suppression only for keys that produce WM_CHAR/WM_SYSCHAR; Up Arrow then `a` enqueues `a`; Enter stays single CR.
4. Keyboard `copySelectionOrPassthrough` is Handled when a selection existed even if clipboard publish fails; selection is retained.
5. `SwapPanes` normalizes exactly one of Folder/Preview/Terminal per pane, round-trips the live terminal HWND/ConPTY, and insertion after an odd swap uses the immutable Open-time source-pane identity.
6. Command-surface workers are joined off the UI thread without retaining one joinable `jthread` per keystroke until Close, and the Close quiet point uses race-free synchronization.
7. `BeginRemovalFocusTracking` stops copying every visible display name while retaining only the small source/outcome proof required by FolderView. The later I13 FolderView remediation owns replacement ordering and feeds successful proof into the canonical resulting-UI-order successor/predecessor resolver; this supersedes the original adjusted-index walk without weakening any I12 proof or epoch gate.
8. Focused test-enabled Release evidence is archived for the removal-focus performance scenario. The authorized September 8 Fresh Full closeout requirement is satisfied by the final qualification below.
9. A failed FO reaper submission cannot expose a payload-less UI completion until the running `jthread` is again owned by `Task`; successful admission transfers the handle to exactly one reaper that joins before posting.
10. Escape consumed by a visible confirmation or command surface arms translated-character suppression, so the following `WM_CHAR 0x1B` never reaches ConPTY.
11. Alt+7/Command Shell opens an embedded terminal with visible shell output. Ctrl+Alt+T/Command Shell Window opens the first floating terminal from a supported file pane without requiring an existing terminal. Preserve prompt customization and prevent shell output from escaping to the host's inherited standard handles.

12. Focus-independent tests retain their activation blocker. Focus-taking entries warn visibly before and during execution without the notice activating or intercepting input. Preserve the lease and real keyboard/focus assertions.

## Authoritative specs

- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/UI/UI_FolderWindow.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/Terminal/Terminal_EmbeddedPlugin.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/Testing/Testing_SelfTests.md`
- `Specs/Testing/Testing_ToolingGovernance.md`
- `AGENTS.md`

## Implementation / tests

See checklist. Every behavioral change has a focused selftest. Run FO completion, Commands preview/tabs/swap, Terminal input/copy/Close, and FolderView removal-focus clusters after each lands.

## Verification

```powershell
.\build.ps1 -ProjectName RedSalamander
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_fileops_completed_post_failure_still_applies_ui_contract
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=pane_view_options_toggle_preview_pane_tabs_and_selection,cmd_pane_embedded_terminal_tab_visibility_and_keyboard_target,cmd_pane_embedded_terminal_plugin_lifecycle
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs,cmd_pane_fileops_move_removal_focus_selects_next_survivor
Invoke-Pester Tools\Tests\TerminalPluginSourceContracts.Tests.ps1,Tools\Tests\TestHarnessSourceContracts.Tests.ps1
.\.build\x64\Debug\PluginContractTests.exe --terminal-selftests
```

## Done criteria

All checklist items marked done with files/tests. Domain specs updated. The authorized September 8 Fresh Full passed on the repaired inputs. Only closeout Markdown and curated evidence change after that gate; the archived source comparison records those exact differences.

## STOP conditions

- Do not move the original 2026-08-19 Done plan back to WIP.
- Do not join File Operations or command-surface workers on the UI thread.
- Do not `SendMessage` FO completion from a worker that the UI may be waiting on.
- Do not weaken bridge MOVE identity/containment.

## Checklist

- [x] **1. FO completion UI affinity and fallback lifetime barrier**
  - Files: `FolderWindow.FileOperations.State.Queue.cpp`, `FolderWindow.FileOperations.State.Runtime.cpp`, `FolderWindow.FileOperationsInternal.h`, `FolderWindow.FileOperations.cpp`, `FileSystem_FileOperations.md`, `Commands.SelfTest.FileOps.cpp`
  - Required proof: dual-post failure applies on the known UI thread; drain submission failure still wakes the UI; shutdown cannot return while a drain callback retains raw state.
- [x] **2. Preview tab must not steal source focus**
  - Files: `FolderWindow.Layout.cpp` `SetPaneContentTab`, `UI_FolderWindow.md`
  - Tests: `pane_view_options_toggle_preview_pane_tabs_and_selection`; `cmd_pane_embedded_terminal_tab_visibility_and_keyboard_target`
- [x] **3. Nav keys must not swallow the next character**
  - Files: `TerminalVt.cpp`, `TerminalInternal.h`, `Terminal_EmbeddedPlugin.md`
  - Tests: Terminal debug selftest including Up Arrow then `a` and Enter single CR
- [x] **4. Failed clipboard copy must remain Handled**
  - Files: `Terminal.cpp` `RouteShortcut`, `UI_CommandMenuKeyboard.md`
  - Tests: debug selftest `selectAll` + force-fail copy → Handled + `hasSelection()`
- [x] **5. SwapPanes exclusive content tab and immutable insertion source identity**
  - Files: `FolderWindow.FileSystem.Navigation.cpp`, `FolderWindow.Layout.cpp`, `TerminalSession.cpp`, `TerminalVt.cpp`, `TerminalInternal.h`, `UI_FolderWindow.md`, `Terminal_EmbeddedPlugin.md`, `Commands.SelfTest.Navigation.cpp`
  - Required proof: current-directory and focused-path insertion both succeed after one `SwapPanes` without replacing the live Terminal.
- [x] **6. Command-surface worker quiet point**
  - Files: `Terminal.h`, `TerminalSession.cpp`, `Terminal.cpp`, `Terminal_EmbeddedPlugin.md`, `TerminalPluginSourceContracts.Tests.ps1`
  - Required proof: race-free predicate wait; retired-worker stress drains to zero; Close join thread remains distinct from caller.
- [x] **7. Removal-focus source-index adjustment**
  - Files: `FolderView.h`, `FolderView.Enumeration.cpp`, `UI_FolderView.md`, `TestHarnessSourceContracts.Tests.ps1`
  - Required proof: deleting `A` and focused `C` from `[A,B,C,D,E]` selects `D`, while Begin remains proportional to requested sources rather than visible rows.
- [x] **8. FO reaper admission must preserve Task lifetime** *(newly identified)*
  - Files: `FolderWindow.FileOperations.State.Queue.cpp`, `FolderWindow.FileOperationsInternal.h`, `Commands.SelfTest.FileOps.cpp`, `FileSystem_FileOperations.md`.
  - Required proof: reaper admission owns the moved `jthread`; admission failure restores it to `Task` before any UI wakeup; the forced-failure test proves UI completion cannot erase `Task` before `ThreadMain` exits.
- [x] **9. Consumed Escape must suppress translated character delivery** *(newly identified)*
  - Files: `Terminal.cpp`, `Terminal_EmbeddedPlugin.md`, Terminal deterministic debug selftests.
  - Required proof: confirmation Escape and command-surface Escape both close their UI state and a simulated following `WM_CHAR 0x1B` writes zero ConPTY bytes.
- [x] **10. Focused validation and evidence**
  - Run the linked Debug builds/exact cases/Pester contracts/Terminal plugin tests, test-enabled Release evidence, archive validation, and the September 8 authorized Fresh Full.
  - [x] Debug `RedSalamander` build: 0 warnings, 0 errors (`msbuild-20260819_183147_318-pid72476-0d8eaa65.log`).
  - [x] Exact Commands cases: 11 passed, 0 failed (`i12-review-fixes-final-debug-20260819`).
  - [x] Terminal plugin contracts: localization/ABI/Close contracts passed; deterministic debug selftests 96 passed, 0 failed.
  - [x] Linked Pester: `TerminalPluginSourceContracts`, `TestHarnessSourceContracts`, `DocumentationDriftContracts`, and `SpecInformationArchitecture` — 244 passed, 0 failed.
  - [x] `Get-SpecInventory.ps1 -FailOnFindings`, `Get-ToolInventory.ps1 -FailOnFindings`, and `git diff --check` passed (line-ending notices only).
  - [x] Test-enabled Release `RedSalamander` build: 0 warnings, 0 errors (`msbuild-20260819_183523_760-pid62032-9ada158c.log`).
  - [x] Release removal-focus exact case passed; `folder.removal_focus.begin_us` measured 490 us for 4096 visible items / 1 requested source. Archive `Specs/TestRuns/4cb089111a23/Commands/2026-08-19_183856` passed `Test-TestRunArchive.ps1`.
  - [x] Post-High-fix Debug `RedSalamander` build: 0 warnings, 0 errors (`msbuild-20260819_192746_099-pid90724-f752f481.log`).
  - [x] Exact FO forced drain-admission failure case passed: 1 passed, 0 failed (`i12-highs-fo-debug-20260819`).
  - [x] Terminal plugin contracts passed; deterministic debug selftests 99 passed, 0 failed.
  - [x] Post-High-fix linked Pester: `TerminalPluginSourceContracts` and `TestHarnessSourceContracts` — 205 passed, 0 failed.
  - [x] Post-High-fix `Get-SpecInventory.ps1 -FailOnFindings` passed with 0 blocking findings.
  - [x] September 8 Fresh Full `20260908T200723Z-69288-1db8a2860de7474ebbc2f9164aa99216`: **2,116 passed / 0 failed / 52 declared skips**, all 35 entries executed/promoted, no reused/provisional/incomplete entries. Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-09-08_200340_i12_final_full`. The earlier linked-tests-only restriction was superseded by the owner continuation request.
- [x] **11. Terminal entry points and visible output**
  - Regressions: `cmd_terminal_session_menu_floating_dispatch` (first opener via shortcut dispatch and menu), `cmd_pane_embedded_terminal_plugin_lifecycle` (visible initial prompt in addition to Running/trusted-idle state).
  - [x] Reproduce both failures before the production repair. The owner independently confirms a blank Terminal tab.
  - [x] Fix and validate the application command gate and ConPTY output attachment in focused Release and actual UI; Full remains item 10.
  - [x] Reconcile the Ghostty update decision with runtime evidence; retain the admitted pin because the host repair works with it.
  - [x] Restore File Operations completion/lifetime and Terminal close/launch contracts in authoritative specs; admit the exact Done path with positive/negative tests.

- [x] **12. Warn before tests that need foreground input** *(September 8 owner request)*
  - [x] Reuse the existing `DirectedSelfTestInputWarning`, moved from the Commands translation unit to `Tests/TestSupport/DirectedSelfTestInputWarning.h`. Remove the duplicate PowerShell/WinForms implementation after owner feedback. Commands, DxUi Menu/NativeTextInput, and interactive ViewerPE/ViewerSqlite cases share the native helper; nested input probes borrow one surface.
  - [x] Enlarge the warning, center it on the owned app/dialog with DPI/work-area handling, name both keyboard and mouse, and explain that input resumes when it closes. Keep no-activation/click-through styles and a three-second lead. Commands starts it before startup activation; the runner's existing desktop lease remains authoritative. No-activation lanes and case enumeration remain silent.
  - [x] Preserve the first failed Full (`20260908T164307Z-104260-ba28d93d5bb041e5900189f2db0b2445`: 2101/14/52), including all failure and skip reasons. The owner confirmed interacting with its test windows; this alone does not establish every failure's cause.
  - [x] First 46-case replay repeated three times: 136/2/0, run `20260908T181539Z-8780-44014c9a25e444f7b11082f75ca7cf9d`. Only Find result shortcuts failed on repetitions two and three: the on-disk recreated files existed, but the pane reused the prior deletion/move snapshot. Reuse `ForceRefreshPaneForCommandSelfTest` in shared Find fixture setup before waiting for the new contents. No timeout increase or product-cache change.
  - [x] Repaired predecessor-inclusive replay: **138 passed / 0 failed / 0 skipped**, run `20260908T182930Z-97660-586ba5c9e75145f5b8839707c4bb9b6c`. All 14 original failures and their predecessor groups passed three times. Debug rebuild: 2m44s, zero warnings/errors. The subsequent startup placement adjustment is qualified by the final Fresh Full.
  - [x] Final Fresh Full (item 10), 265 focused contract checks, and owner confirmation that the reused warning is larger, centered and readable. Exact admission/archive checks are recorded in the linked closeout evidence.

Evidence: `Specs/TestRuns/4cb089111a23/Commands/2026-09-08_180534_i12_terminal_closeout/README.md`. Failed experiments and the superseded warning probe remain explicitly historical; their results do not qualify the final native helper.

### September 8 validation scope corrections

- Actual Release UI verification passed Alt+7, Ctrl+Alt+T from a file pane,
  and Commands > Terminal > Command Shell Window; each displayed the owner's
  customized PowerShell prompt. No terminal command was entered through UI
  automation. The app and companion monitor were closed normally afterward.
- The first UIA-based prompt probe proved unreliable: the nine-case Release
  attempt `20260908T162713Z-48948-a9cf39e8a96e410192229af7769cfe81`
  hit its bounded-worker fail-fast while discovering the terminal provider.
  Preserve that incomplete/failed run. The replacement test-only hook reads
  the existing bounded VT formatter directly; it introduces no UIA worker or
  public Terminal ABI change. The replacement passed the nine-case Release run; Fresh Full remains item 10.
- The lifecycle fixture must cancel the text inserted by its earlier SwapPanes
  checks before testing follow-when-idle. Real input correctly defers following
  while PSReadLine contains unsubmitted text; the former disconnected shell
  concealed that fixture dependency. The repaired launch passes prompt visibility.
- The preview-tabs case reads ViewerText document contents through a snapshot
  available only under `_DEBUG` (`FolderWindow.FileSystem.Commands.cpp` and
  `ViewerText.cpp`). Its Release failure does not establish a preview product
  regression and is retained as an unsupported measurement attempt. Keep this
  required preview witness in Debug/Fresh Full; focused test-enabled Release
  covers the nine terminal/removal/completion cases that expose the required
  instrumentation. Do not label the original ten-case Release run green.

## September 8 final-gate follow-up

The next Fresh Full (`20260908T183914Z-71172-fa6f9f19305b4ba7b23d571d207daf93`) remained failed: **2,113 passed / 2 failed / 53 skips**, with 33 of 35 entries promoted. Commands passed **896/0/2**, including all 14 original failures and every I12 terminal, preview-focus, removal-focus, and failed-completion-post witness. The terminal plugin passed 162 assertions. This is retained failure evidence, not a successful closeout.

- FileOps `Phase6_LocalBandwidthThrottle` measured 265 ms from cancel request to UI completion against its existing 250 ms limit. The worker cancellation metric was 15 ms. The path is unchanged by the terminal/warning repair, and the warning is disabled in FileOps. Five exact repetitions then passed (15 total cases including setup/cleanup), with cancellation samples 32, 31, 16, 31, and 32 ms. No budget, product code, or test assertion was relaxed; the isolated replay does not erase the failed Full.
- ViewerPE interactive failed the ViewerSpace tooltip-paint witness, including its reuse by the combo-host long-run case. The exact case initially passed twice with the warning visible. Queueing `WM_MOUSELEAVE` reproduced the failure deterministically: the debug hook painted synchronously, but the subsequent message pump dismissed the tooltip before its state was checked. The test now captures the real paint-count increase and nonempty state immediately after the hook; the queued dismissal remains as an adversarial fixture. Five corrected repetitions passed. Production rendering and the approved warning presentation are unchanged.
- [Failure and replay evidence](../../TestRuns/4cb089111a23/Commands/2026-09-08_183818_i12_fresh_full/README.md) retains the aggregate, detailed results, both relevant metrics, and the before/after viewer reproduction. The final qualification record determines closeout; this interim attempt cannot do so.

## Final September 8 qualification

- Fresh Full: `20260908T200723Z-69288-1db8a2860de7474ebbc2f9164aa99216` — **2,116 passed / 0 failed / 52 declared skips**; all 35 entries promoted. Build receipt: `f94d01029dc0bbf39e4e1786920704a08df4a05614c8d872d471cb0fd7312956`; tested snapshot: `018041afc4cea6efb7d4b0575f7857b4f2c52ff335058c2095bc85ffb01915a1`.
- [Final gate and closeout evidence](../../TestRuns/4cb089111a23/Commands/2026-09-08_200340_i12_final_full/README.md) retains per-case results, skip reasons, plan/decision/provenance, plugin contracts, and bounded terminal/removal-focus metrics. Prior failed attempts remain in the linked repair archive.
- Terminal qualification: nine test-enabled Release cases passed; actual Alt+7, Ctrl+Alt+T from a file pane, and the menu route displayed customized prompt output. The final Debug Full supplies the preview-tabs witness and the full terminal/plugin/FO completion coverage.
- Ghostty remains at admitted commit `9c3ec931d64561a8407dde7ac984ce156ae91539`; the defect was in host command state and ConPTY attachment. I15's separate runtime/ARM64 qualification is not claimed here.
- The 14 initial Commands failures were followed by a 136/2/0 replay, a targeted fixture cache refresh, and a 138/0/0 replay. An intermediate Full then passed all Commands but failed a 265/250 ms FileOps UI-completion sample and a ViewerSpace tooltip witness. Five unchanged-limit cancellation repetitions passed at 16–32 ms. The tooltip race was reproduced with queued mouse-leave, corrected by capturing the synchronous paint before pumping dismissal, and passed five exact repetitions plus the six-cycle viewer churn test. This final Fresh Full passed on the resulting inputs. Prior failed runs remain failed; not every original failure is attributed to desktop interference.
- WIP priority routing now starts with I10's independent closeout qualification, then I18 characterization and metadata-outcome foundations. This completion does not activate HOLD plans or close I10/I15.
