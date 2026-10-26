# Paused-reader stop RED observation

Source: c36d3497a0d5214168c062bf3f67a16e6baa909a plus the P1/P2/BR-3 worktree
before stop/cancel notifications synchronized with both wait mutexes.
Test-enabled x64 Release, zero build warnings/errors.
Build receipt: 53a943a93bd8bb83fad7c12873c2a8859b1b0c05b0829a51701947c4b4ed8d50.
Runner: 20260908T071033Z-78696-a24db26346694f068b45ab77d1633524.

Command: Tools/Run-AllTests.ps1 -Configuration Release -Suite FileOps -CaseFilter Causeway_BridgeFailureStatusAndPausedReader -TestRoot Z:\RedSalamander.Perf
with RSBuildEnableTests=true.

Two setup/cleanup cases passed; the selected case failed after 2,047 ms at the
existing paused-reader stop assertion. No ConflictCancelWake metric was reached,
so this run does not establish the forced peer-conflict ordering. The following
five-repeat RED archive supplies that proof. Results and bounded traces/metrics
are copied verbatim. This focused failure is not repository qualification.