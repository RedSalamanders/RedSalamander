# Observed pre-fix reproduction

This is a transcription of the governed runner output and trace inspected during this chat,
not a recovered raw archive. The next runner invocation automatically retired the completed
sandbox run before its raw files were curated. No timing or full qualification claim uses it.

Command: Test-Changes.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts
Run: 20261016T210953Z-92952-0d353ef4841f42aa80d29404a02d32e7
Debug x64 build: 0 warnings, 0 errors.
Build receipt: 2f8e0bc3a8774620d46d110384d3270e317772ac0c3f4d6687b066d5dd04fa7f
Runner: exit 1; setup and cleanup passed; FileOps_ReviewRemediationContracts failed.

The production guard was unchanged in this run; the new resources and fixture were present.
Cases 0, 1 and 2 (existing destination with different bytes, identical bytes and identity-less
parent respectively) all reported HRESULT 0x80070032, preserved=true, passed=false and:

    The request is not supported.

Cases 3 and 4 (missing destination and unsupported source respectively) both reported the
same HRESULT with preserved=true and passed=true. The original text was also wrong after
French/English re-rendering. The candidate fixes the fixture's running-note observation order:
the running diagnostic is captured before completion moves it into the completed summary.
Candidate measurements are observational only; there is no retained raw baseline timing.

The first fixture build failed because ExecutionMode was initially qualified through Task
instead of FileOperationState (C3083/C2039/C2065). That was corrected before this reproduction.
