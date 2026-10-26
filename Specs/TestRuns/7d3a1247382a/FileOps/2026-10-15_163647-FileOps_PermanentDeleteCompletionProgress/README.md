# Permanent Delete progress: failed focused attempt

Focused x64 Debug run: **failed**, exit code 1; 2 records passed, 1 failed, 0 skipped.
Setup and Cleanup passed; `FileOps_PermanentDeleteCompletionProgress` failed.

- Run: `20261015T143618Z-436-040ba2be9fd04b4d9ad806282b2d34ea`.
- [Build receipt](build-receipt.json): `6b42b3d9d25bb8f351668a577f0ef2fcd4ee2279c5aff47505ef71e196ff27e7`.
- The [trace](fileops_trace.txt) records successful internal deletion checks: partial: 1 file/17 B, confirmed: 2 files/1 folder/46 B, canceled: 1 file/7 B, and refused: 0 completed objects with read-only attributes retained.
- The actual host-driven normal-parent/read-only-child flow passed its separate permanent-delete and read-only confirmations, exact completed counts, and foreground preservation.
- The later sequential bridge check failed. Follow-up source review identified that its shared guard required the parallel-only `bridgeDirectoryEnsureCount` value on the sequential route. The selected-read-only-roots helper was not reached in this attempt.

The earlier [callback assertion](../../Continuation/2026-10-15_162805-delete-fixture-callback-assertion/README.md)
did not recur after the fixture adopted callback-scope RAII. That correction does not turn this
different fixture-guard failure into a pass. The complete Delete case still requires a clean rerun
on the corrected test build; this archive retains its failed classification.

The [raw aggregate](run-all-tests-results.json) retains all 3 legacy sandbox run-directory findings
and `is_clean:false`. No raw evidence or audit statuses were changed. This focused attempt gives no
repository-wide green verdict; Release remains unqualified.
