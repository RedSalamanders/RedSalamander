# Passive-listener fixture — compile failure preserved

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-passive-lifecycle-debug-20260905`.

Based on `61815385`, with the fixture-only passive-listener state and two new raw
FTP lifecycle checks. The first non-elevated launch stopped before building:
`Unable to verify that no residual RedSalamander build tools are active: Access denied`.
The same governed command was rerun with process-inventory access; no ownership
guard was disabled and no unrelated process was stopped.

Build failed in 01:48: two C4625/C4626 warnings (implicitly deleted resource-owner
copy operations) and C2039 at the old ABOR branch (`PassiveListener.reset`). The
complete 20260905_182637_044-pid68312-df9f94c0 log is retained as `build-failure.txt`.
One whitespace-only line was trimmed for patch hygiene; all diagnostic text and
line positions remain. Original log SHA-256:
`612730a7cb325a5b7249102b6b15e33950552865754225d9fc88a36357c5044d`.
**No runtime tests ran and no successful receipt was issued.** Source/localization
checks completed separately: 197/197, 15.47 s; that is not a compiled-test pass.

Correction: explicitly declare the listener state's non-copy/non-move ownership,
and clear its WIL socket/pending advertisement on ABOR. The next run must compile
warning-free and execute the complete intended C0 filter, retaining all wide/deep
content and deadline assertions.
