# I26 four-primary-action policy — baseline

Expected failing x64 Debug witness at HEAD `a212edec433d658748ae1273cba68633c86de8b2`, with dirty
test/measurement inputs and the unchanged three-primary production policy. The previous ordinary
action-height repair is present in both members of this new comparison.

- Runner: `20260919T104753Z-20280-c712b06da27a4f9ca3081373271d9008`.
- Receipt: `278e6746cd3448257259723b4373efc697915bf60beb401570fac949d900908f`.
- Source snapshot: `28521752445aaa8d02529db120acab7dd34fd83d53b42af7db18441833c6924a`.
- Retained exact policy fixture SHA-256: `90688FBD9B4343BBF323622013B3AF17AA3A45CB87D8A75551093123BC8A612E`.

Build: zero warnings/errors. `Phase10_DeferredConsentAndRecycleEscalation` failed its ordered
policy expectation; Setup/Cleanup passed. The expected policy now requires four eligible direct
collision actions and preserves typed default/Escape, membership, scope and metadata-loading tests.
The matrix covers all 24 conflict classes and 64 permission combinations. It validates unique,
complete primary/overflow partitioning and prohibits actions whose corresponding authority is absent.

`FileOps.ConflictPolicy.MatrixBatchUs` measures five 16-pass batches after one full warm-up:
**32802, 33219, 35004, 33337, 33084** microseconds; median **33219**. Each batch constructs 24576
policies. `value0` is the round; `value1` is the observed placement/default checksum plus warm-up
checksum (3490816 here). This is construction cost, not rendering, input latency or filesystem I/O.
The exact fixture source and complete raw metric stream are retained; compare the unchanged fixture.

Command: `Run-AllTests.ps1 -Suite FileOps -CaseFilter Phase10_DeferredConsentAndRecycleEscalation
-Platform x64 -Configuration Debug -TestRoot C:/RedSalamander.Perf -TimeoutMultiplier 3`, with
`RSBuildEnableTests=true`. The separate sandbox audit still reports 225 preserved pre-existing run
directories. This expected failure is not Full closeout evidence.
