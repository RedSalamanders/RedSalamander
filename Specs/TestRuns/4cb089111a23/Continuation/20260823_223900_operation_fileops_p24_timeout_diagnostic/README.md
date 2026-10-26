# File Operations P2.4 timeout diagnostic

- Date: 2026-08-23
- Run: `20260823T202549Z-62976-a15ee289e4da4c9b8578be00ee156bb1`
- Command: `./Tools/Run-AllTests.ps1 -Suite FileOps -SkipBuild -ValidationMode Fresh`
- Build: Debug x64, attestation `bd06e00e24b0439e9b62a6dfe4d92ec722a1beb6b661eecb7e714bf470554c12`
- Outcome: diagnostic run stopped after the first failure and after it entered the duplicate dedicated bridge family.

## Failure

`Phase11_BridgeMultiFolderParallelCopyInFlightLines` reached its 240-second step timeout. The retained trace contains:

```text
[114828500] NextStep: Phase11_BridgeMultiFolderParallelCopyInFlightLines
[115068578] FAIL: Phase11_BridgeMultiFolderParallelCopyInFlightLines timed out.
```

This was not a bridge deadlock. The task completed successfully at the timeout boundary. The run's performance stream recorded 48 admitted and completed files (100,663,296 bytes), two ensured directories, admission queue high-water 16, retained-entry high-water 21, retained-path high-water 11,634 bytes, retained-metadata high-water 9,674 bytes, and zero traversal-limit hits. The nested workers ran for about 99 seconds. The remaining approximately 141 seconds were spent creating the 96 MiB fixture before task admission, but the phase timeout started before fixture creation.

## Resolution

Reset the step timeout clock immediately after the bridge task is admitted. This keeps the 240-second operation bound while excluding variable fixture-write latency caused by filesystem filters. The diagnostic run is retained so the timeout classification and its successful bridge counters are auditable.
