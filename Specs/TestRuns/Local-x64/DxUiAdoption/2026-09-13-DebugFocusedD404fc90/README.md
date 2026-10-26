# Debug focused qualification at d404fc90

The x64 Debug root build passes in 3m24s with zero warnings/errors. Pending-content,
removal-focus and bandwidth each pass ten repeats; each post-run check verifies all
14 actual consumer module identities and exact DxUi pin 13788e95.

Queue-after repeat 1 passes in 1,610 ms. Repeat 2 reuses the configured Dummy provider's
destination and waits for an unrelated overwrite decision. A noninvasive stack snapshot
shows WaitForConflictDecision beneath CrossFileSystemBridge::PromptDestinationCollision.
The exact owned diagnostic process is then interrupted; the summary and partial native
coverage are failures, not a complete passing suite. Debugger inspection also disqualifies
this attempt as timing acceptance. The driver stops before don't-start or Fresh Full.

The next fixture clears only its own destination before each repeat, rejects unexpected
cleanup errors and retains the existing publication gate, workload and assertions.
The original logs, partial results, process identity and interruption reason are retained.
No production code or DxUi pin changes are made by that correction.
