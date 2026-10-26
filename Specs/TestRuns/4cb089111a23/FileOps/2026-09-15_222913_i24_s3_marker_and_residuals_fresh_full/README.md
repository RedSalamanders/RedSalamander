# I24 S3 marker transfer and two I22 residuals: Fresh Full closeout

Source: `359e4419`, a clean working tree.
Governed run: `20260915T222913Z-82160-20b6d19dd10f404ca88a3a5e4f395a83`.
Command: `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -AllowAlternateVolumeTestRoot`.
Profile: test-enabled x64 Debug; the full-solution build reported zero warnings and zero errors.
Accepted build receipt: `775d70d6cb58cab762d038a6ed30d4cc721d247d916d420d4e9fc82c69d5a9ba`.
Workspace snapshot: `1aa17ccf0f528a52d317d28b5537cbc149b3f23e7f3a4223798a56f688a57c22`.
Validation plan digest: `58336f513976e44ce58d3aabd7df254fc5020e40c27108305ad67ae98bf6d344`.

The three fixes this run closes each landed with a focused witness (see the plan's Execution record);
this run is the proof that the three together leave the whole suite green on the machine that closed
I22 and I23 the same day.

## Verdict

Run status is **failed** in `Fresh` mode for suite `Full`: all 20 plan entries executed and were
promoted, and the runner recorded **2,143 passes, one failure, 52 declared skips**, 2,196 total, in
65 minutes 39.2 seconds. The one failure is classified below and is not a product defect; it is kept
in this archive as it happened rather than relabelled. This is one Fresh Full invocation, not a
composite.

The single failure, `cmd_pane_navigationView_history_dropdown_keyboard_navigation`, reported
`keyboard=-1, hovered=18, cursorBefore=(1476, 1176)` against a popup rect of `(1012, 488, 1726, 1253)`:
the machine's physical mouse cursor was parked inside the popup, so hover selection overrode the
keyboard selection the case drives. It passed in the three previous gates on this machine (I23, I22 and
this plan's first attempt, 170 to 192 ms each) and 3/3 immediately afterwards once the cursor was
elsewhere, and nothing in this plan touches navigation history. It is the live-cursor interference
class the test suite already records.

This was the second attempt. The first, run `20260915T211116Z-48752-87739ba450f44ef18a76815c9764e343`,
had every suite green except FileOperations, where `Phase7_RecycleBinBatchDeleteMultiBatch` hung in the
shell's recycle operation and timed out at 240 s (it had passed in about 10 s in both earlier gates on
this machine and passed 9/9 immediately afterwards), the family aborted, and the next family's `Setup`
could not recreate the temp root the hung operation still held: 2 failures and 36 "not reached" skips
from one shell hang. That cascade is recorded as a harness residual in the successor plan. Both attempts
built with zero warnings, both passed `PluginContractTests` with `FileSystem.dll` 163/0 and
`FileSystemS3.dll` 238/0, and between them every suite has been green on this tree.

| Suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| Commands | 897 | 1 | 2 |
| CompareDirectories | 242 | 0 | 30 |
| FileOperations | 193 | 0 | 20 |
| ToolsPesterTests | 784 | 0 | 0 |
| ToolsPesterBuildToolchain | 12 | 0 | 0 |

## Witnesses this run carries

| Case | Result | What it proves |
|---|---|---|
| `terminal_shortcut_command_palette_closes_on_deactivation` (Commands) | Passed, 35 ms | Task 2: after the palette closes on deactivation both debug hooks refuse; before the fix this exact case crashed with an access violation. |
| `cmd_pane_embedded_terminal_command_state_truth` (Commands) | Passed, 509 ms | Task 2: the case that crashed the laptop's Commands suite, now with a named precondition on the palette being open. |
| `Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker` | Passed, 969 ms | Task 3: runs `FileSystem.dll`'s debug self-tests in-process; every check now reports through the shared `Check`. |
| `PluginContractTests` (standalone) | Passed | Tasks 1 and 3: `FileSystem.dll` 163/0 and `FileSystemS3.dll` 238/0, the latter including the seven marker-transfer checks. |
| `DiscoveryScope_S3DirectApi` | Passed, 859 ms | Task 1 control: the I22 S3 witness on the same relay route, unchanged. |
| `R0fS3_FakeS3ReadWriteCreateDelete` | Passed, 2,016 ms | Task 1 control: cross-provider Copy both ways, CreateDirectory, Rename and Delete on the same fixture. |

## Suite-wide invariant

`perf/discovery-scope-metrics.jsonl` retains this run's `FileOps.Discovery.OpenUs`,
`FirstMutationBeforeClose` and `GrowthAfterClose` rows and the `FileOps.SelfTest.DiscoveryScope.*`
rows, 1,161 rows. `FileOps.Discovery.GrowthAfterClose` is 0 across all **382** recorded
task completions; nothing in this plan touched discovery, and this shows it.

## Declared skips

All 52 skips are environment-gated and none were waived or deleted by this work.

- **CompareDirectories, 30.** Live journal cursor unavailable (nine), an NTFS-backed root (one), SMB (one),
  and absent remote connection profiles (nineteen).
- **FileOperations, 20.** Eighteen `Phase16_Remote*` cases for the same absent profiles, plus the two
  `Phase15_FileSystem7z*` cases owned by Compare.
- **Commands, 2.** The two explicit ViewerSpace opt-ins.

The same skip shape as the I22 and I23 gates on this machine.

The runner also reports pre-existing test-root layout audit residuals. They are preserved in the
aggregate and were not deleted or waived here.

## Retained artifacts

| File | Contents |
|---|---|
| `results.json` | Full aggregate case statuses and skip reasons. |
| `run-all-tests-results.json` | Runner summary with per-entry durations. |
| `run-state.json` | Validation run identity, mode, build receipt, snapshot and entry promotion. |
| `validation-plan.json` | The executed Fresh plan. |
| `workspace-final.json` | Final workspace identity binding. |
| `trace.txt` | Whole-run trace. |
| `fileops/results.json`, `fileops/trace.txt` | FileOperations suite detail. |
| `commands/results.json` | Commands suite detail, including the two palette witnesses. |
| `perf/discovery-scope-metrics.jsonl` | Discovery metric rows for this run, machine-bound to profile `4cb089111a23`. |
