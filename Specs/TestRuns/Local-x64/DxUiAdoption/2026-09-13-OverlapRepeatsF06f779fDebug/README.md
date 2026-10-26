# Overlap regression repeats: x64 Debug

RedSalamander `f06f779fd45105d251c2761a82138e7c8215326e`, DxUi `13788e95`, passes ten queue-after and ten don't-start
repeats. Each native File Operations run reports 30 passed, zero failed and zero
skipped including setup/cleanup. All 14 consumer module identities are verified
against actual bytes after each run. No hosted CI is used.

The first destination publication is held until Delete's queued wait or explicit
cancel is observed. The fixture now clears only its own Dummy destination before
each repeat and rejects unexpected cleanup errors. All 200 files, original overlap
prompt/default, queue-order, cancellation and source-retention assertions remain.
Only the File Operations test source changes between d404fc90 and f06f779f; no
production code or dependency pin changes.

The [original interrupted repeat](../2026-09-13-DebugFocusedD404fc90/README.md) remains
evidence: repeat 1 passed, but repeat 2 waited for an unrelated overwrite conflict.
These focused passes do not replace Fresh Full qualification. Exact run IDs and
original receipt/driver hashes are retained alongside the native per-repeat results.
