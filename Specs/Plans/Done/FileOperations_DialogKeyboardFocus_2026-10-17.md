# I38 Decision dialog keyboard focus

Status: DONE — implementation and focused x64 Debug/Release acceptance complete.
Owner: this change, requested 2026-10-17.

The pane decision mirror must take focus when it appears in the foreground application and accept
Tab/Shift+Tab, arrows, Enter, Escape and Space immediately. Keep the safe initial action, F6 opt-out,
focus restoration, background-window isolation, and held-key repeat suppression.
The owning contract is [File Operations popup](../../UI/UI_FileOperationsPopup.md#decision-mirror).
This changes consumer policy only; the pinned DxUi library and optional shared overlay guard stay intact.

- [x] Trace the screenshot to the mirror's opt-in keyboard guard and pane-only focus condition.
- [x] Remove the guard opt-in on show/refresh; take foreground focus from either pane.
- [x] Extend deterministic mirror checks and the real-input Commands case for immediate navigation,
  Escape, safe Enter, other-pane focus return, and repeat/F6 behavior.
- [x] Update the UI contract and user guide.
- [x] Build x64 Debug and run the affected mirror and real-input cases through the canonical runner.
- [x] Review formatting, spec consistency and evidence; record precise remaining qualification limits.
- [x] Build x64 Release and run the same affected mirror group and real-input case after the user closed
  the existing Release instance and authorized testing. Keep other platform gates separate.

The person at the desktop agreed to the existing focus-taking regression harness on 2026-10-17.
Use its warning and desktop lease, with foreground/focus restoration. No manual desktop capture.
This is an input-policy fix, with no new polling, timer, rendering path or throughput work.
Existing FileOps.DecisionMirror.Show telemetry records focus acquisition.

## Verified evidence

Both Debug builds completed with zero warnings/errors. The final source and dependency bindings remain
in each run's original workspace/result artifacts; the two runtime scopes have distinct build receipts.

| Scope | Result | Run / build receipt |
| --- | --- | --- |
| FileOpsFamily_DecisionMirror (seven mirror cases plus two setup cases) | 9 passed, 0 failed/skipped | `20261017T154024Z-105624-bb0e0a2e8f1940d098c199bda68fe705`; `b112c39fb842d0d15ce9df21f94153a488b07b3ce0deed563c3c5fd1249ac613` |
| cmd_pane_fileops_decision_mirror_real_focus | 1 passed, 0 failed/skipped | `20261017T154635Z-50796-ad9f20c6ef10459095214c3f92511999`; `0898f939c17f88f61321aa370de6f5760da45ac502934d3a2546ba92534fbc5b` |
| DocumentationDriftContracts and TestHarnessSourceContracts | 222 passed, 0 failed/skipped | Repeated after final test changes |
| Changed C++ ranges / whitespace | Passed | Visual Studio-discovered clang-format dry-run/Werror and git diff --check |
| Specification inventory | No blocking findings | Get-SpecInventory -FailOnFindings |

Original result JSON, trace and workspace artifacts were copied with matching SHA256 hashes under
`Z:\RedSalamander.Perf\evidence\dialog-keyboard-focus-2026-10-17`, separately for FileOps and Commands.
That folder also records tested code hashes and the initial failure observation. The first native
attempt failed an overlap test's stale guarded-focus expectation (4 passed, 1 failed, 4 skipped);
the corrected complete group passes above. Its raw run was removed by the canonical runner's next-run
cleanup before copying, so the retained observation is explicitly not an original raw result artifact.
The first compile's receipt was rejected because source changed during that build; neither attempt
is promoted to a passing result.

Initial scoped explanation correctly left FileOperations pending without Interactive. PrePush explanation
could not verify PR equivalence and retains broader local obligations; no Full, native ARM64,
ASan, assistive-technology or performance non-regression claim is made. Existing unrelated dirty
governance work is preserved on branch `ej/dialog-keyboard-focus`.

## Release closeout — 2026-10-17

The user closed RedSalamander and explicitly authorized Release build and both fixes' runtime checks.
Both test-enabled x64 Release builds completed with zero warnings/errors. The mirror group passed
9/9 (seven mirror cases and two setup/cleanup cases), run
`20261017T162800Z-80208-c502e7dfb5484d5d89ea79eb3fcd6d7b`, receipt
`4022a40b1a3b7a4443c1dd1ea29afd9bf3df56449260acdadd3dd78e10a40da4`.
The exact real-input case passed with the natural-sorting case in one Commands run (2/2),
`20261017T163444Z-107504-ef07845e7c224789a110c4039566a6fa`, receipt
`479a733e2a272305e7b9ef5a166e51a37c90f9b9b14599f8eaebd9d1ed14c2a1`.
There were no failures, skips, flaky classifications or repair attempts. Original results, raw traces,
build logs, receipts, workspace identities and tested C++ hashes are in the
[Release evidence packet](../../TestRuns/Local-x64/Navigation/2026-10-17_ReleaseDialogAndNaturalSorting/README.md).
The six changed C++ files match their passing Debug identities; subsequent edits only close plans,
admit their exact Done paths and preserve evidence. These focused runs do not close Full, ASan,
native ARM64 or the separate File Operations post-merge qualification program.
