# Forced conflict cancellation lost-wake RED control

Source: c36d3497a0d5214168c062bf3f67a16e6baa909a plus the P1/P2/BR-3 worktree,
including the forced conflict-predicate gate, before the notifier mutex fix.
Test-enabled x64 Release, zero build warnings/errors.
Build receipt: 1f68765c6c4ef38f7f074c8235edd69de2b78476cd640eec2b808b9df336e026.
Runner: 20260908T071617Z-66932-d222086f33354384814a9bc45acf7f62.

Command: Tools/Run-AllTests.ps1 -Configuration Release -Suite FileOps -CaseFilter Causeway_BridgeFailureStatusAndPausedReader -SelfTestRepeat 5 -TestRoot Z:\RedSalamander.Perf
with RSBuildEnableTests=true.

All five repetitions failed; their ten setup/cleanup cases passed. Each peer was
held under its conflict mutex after a false predicate and before condition-variable
sleep. RequestCancel returned before the gate opened, losing its notification.
All five FileOps.SelfTest.ConflictCancelWake rows have value0=0 (failed), value1=1
(premature notification), hr=2147500037. The two-second peer-return assertion
failed; only failure recovery released the conflict so teardown remained finite.

The complete bounded metric stream, results and trace are retained verbatim.
This is an ordering/correctness control, not a throughput baseline or proof of the
exact worker stack in the historical Full-gate failure.