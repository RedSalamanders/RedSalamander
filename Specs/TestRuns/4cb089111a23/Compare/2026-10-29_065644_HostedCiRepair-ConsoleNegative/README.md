# Original captured-console launcher control

Original helper at 7a3a7057669c5cd9d2ae5c629133f6cdebd23875 with only the
retained test/diagnostic patch, x64 Debug Embedded. Run 20261029T065644Z-83484-edc90692a8724fcdb79c716693275808:
zero passed, one failed, zero skipped; wrapper exit 1. The independent child
returns exit zero and console-state Y, failing exactly the new no-console oracle.
All original stream, quoting, exit-37, inheritance, timeout-tree and cancellation
assertions remain. Native duration: 2230 ms.

The full build receipt binds the uncommitted input snapshot. This expected red
control proves the console-isolation defect, not the precise cause of the earlier
broad duplicate-writer exit-zero failure. That earlier failure remains retained.
