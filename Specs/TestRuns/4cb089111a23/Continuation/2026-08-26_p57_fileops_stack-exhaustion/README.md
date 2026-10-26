# P5.7 File Operations self-test stack-exhaustion continuation

## Trigger

- Command: `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild`
- Run: `20260826T162158Z-34212-fe29ecddad7d4fd7a406d3de9217cf61`
- Build attestation: `349fce5cc3c6a920970b24d537becabd52bf0fdfc6203c0b3e2d507325ef15dd`
- Result: 53 File Operations cases passed; `RedSalamander.exe` crashed while starting `Phase11_BridgeMultiFolderParallelCopyInFlightLines`, so the runner reported an incomplete suite.

## Crash evidence

- Exception: `0xC0000005`
- Faulting process: PID 67316, Debug `RedSalamander.exe`
- WER dump: `%LOCALAPPDATA%\CrashDumps\RedSalamander.exe.67316.dmp`
- Dump size: approximately 129 MiB; the dump is intentionally not copied into the repository.
- Dump SHA-256: `461E5127C854111C37EC17A7D301606E884EA1731BA9C95AF0EFA11042B7CA91`
- WER report, FileOps trace/results, and runner result are retained beside this note.

## Diagnosis

The exception write was 16 bytes below the committed UI-thread stack. PDB frame data showed the anonymous `TickPhases10To13` helper reserving 331,968 bytes, while `FileOperationsSelfTest::Tick` reserved another 13,248 bytes. Nested Phase 11 popup/rendering and crash-reporting calls therefore exhausted the default 1 MiB UI stack. The crash-handler frames visible in the dump were secondary reporting work, not evidence of a File Operations executor access violation.

## Continuation rule

Split the Phase 10–13 include dispatcher into separately compiled phase helpers, rebuild, confirm the helper-frame reduction, rerun the failing Phase 11 case, and then rerun the complete FileOps suite. Do not count this interrupted run as P5.7 closeout evidence.

## Harness correction

The dispatcher now compiles Phase 10, 11, 12, and 13 into separate helpers. PDB `S_FRAMEPROC` measurements from the corrected Debug executable are:

| Helper | Stack frame |
|---|---:|
| `TickPhase10` | 194,336 bytes |
| `TickPhase11` | 83,232 bytes |
| `TickPhase12` | 39,296 bytes |
| `TickPhase13` | 19,952 bytes |

The crashing Phase 11 path therefore no longer inherits the former 331,968-byte combined frame. Final behavioral evidence is recorded only after the governed focused and complete FileOps reruns terminate cleanly.

## Sequence rerun

The corrected Phase 11 case passed 3/3 in isolation (`20260826T170937Z-99416-ef7f90f74ea14583bd14928bb3cba545`). The next complete FileOps run (`20260826T171158Z-46352-fb37097fcec0403aa05a028ed3b74b27`) passed 53 cases, completed the former Phase 11 crash case, advanced to `Phase7_CopyItemsSingleFolderRecursiveParallelism`, and then terminated with another `0xC0000005`:

- WER dump: `%LOCALAPPDATA%\CrashDumps\RedSalamander.exe.45692.dmp`
- Dump SHA-256: `7D37668FF0DB1AA0D0B42330CE028CA9148E873BD12D957126AF69C0C848B159`
- Fault RVA: `0x01CF439E`

PDB inspection exposed the remaining oversized dispatchers: `TickPhases05To06` reserved 762,656 bytes, `TickFairstreamPhases` 360,000 bytes, and `TickPhases07To09` 232,104 bytes. A nested timer/UI callback can enter a default dispatcher while an active phase frame is still live, so an unrelated 762 KiB default switch can exhaust the UI stack. The continuation therefore splits Fairstream, merge/provider-matrix, Phase 5, Phase 6, Phase 7, Phase 8, and Phase 9 into independently compiled helpers as well; successful final PDB measurements and suite evidence supersede the intermediate table above.

## Final correction and result

The final Debug PDB reports these bounded reentrant dispatcher frames:

| Helper | Stack frame |
|---|---:|
| `TickCopyMerge` | 9,776 bytes |
| `TickMoveMerge` | 34,592 bytes |
| `TickReparseMerge` | 8,048 bytes |
| `TickFairstreamA` | 113,632 bytes |
| `TickFairstreamB` | 110,288 bytes |
| `TickFairstreamC` | 139,360 bytes |
| `TickPhase5` | 72,800 bytes |
| `TickPhase6` | 41,392 bytes |
| `TickPhase7` | 125,480 bytes |
| `TickPhase8` | 27,424 bytes |
| `TickPhase9` | 82,432 bytes |
| `TickPhase10` | 194,336 bytes |
| `TickPhase11` | 83,232 bytes |
| `TickPhase12` | 39,296 bytes |
| `TickPhase13` | 19,952 bytes |

The complete governed File Operations rerun
`20260826T172428Z-24428-b09b4e4a5d2b48dbb9af0cfb5e2b7b0c` then terminated cleanly with
**123 passed / 0 failed / 20 environment/delegation skips**. Its compact validated archive is
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-26_194120/`. This establishes that both access violations
were self-test UI-stack exhaustion, not executor memory corruption, and that the split dispatchers close the
continuation without suppressing either former crash path.
