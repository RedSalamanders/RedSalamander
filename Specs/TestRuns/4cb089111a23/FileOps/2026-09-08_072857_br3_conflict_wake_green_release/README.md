# Conflict cancellation wake GREEN control

Source: c36d3497a0d5214168c062bf3f67a16e6baa909a plus the P1/P2/BR-3 worktree.
Test-enabled x64 Release, zero build warnings/errors.
Build receipt: 6cddbbae87c79d403f089219dae346f77578a6d518be87290223290e02b22e7c.
Runner: 20260908T072857Z-81648-142f935e646a4c0284d1ee2cae71ba80.

Command: Tools/Run-AllTests.ps1 -Configuration Release -Suite FileOps -CaseFilter Causeway_BridgeFailureStatusAndPausedReader -SelfTestRepeat 5 -TestRoot Z:\RedSalamander.Perf
with RSBuildEnableTests=true.

All 15 setup/case/cleanup results passed. All five ConflictCancelWake rows have
value0=1 (passed), value1=0 (no premature notification), hr=0. Each held peer
returned within two seconds after the gate opened, without failure recovery.
The existing external-stop paused-reader check also passed in every repetition.
The complete bounded results, trace and metric stream are retained verbatim.

Compared with the five-repeat RED control in the sibling 071617 archive, the
production change synchronizes cancel/stop notifications with both predicate
mutexes. This establishes the fixed ordering, not a throughput improvement.
Fresh Full is separately required for whole-process Phase 7 and shutdown proof.