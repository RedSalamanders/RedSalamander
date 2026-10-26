# Qualified source snapshot and focused migration handoff

The source at 582d644be5e6147acb0b8a8aae05511b75f64d17 passed all six x64/ARM64 build profiles and
all eight original ASan regressions. Ten explicit focused launches recorded
51 native case passes, zero failures and zero skips, including mandatory
FileOps setup/cleanup. The no-frame A/B UIA checks and original inspection-return
assertions passed. A seeded heap overflow produced the required real ASan
heap-buffer-overflow rejection. All three actual application profiles passed
real ETW Unicode, severity, PID, TID and timestamp roundtrips. The unchanged
500 ms nonblocking search gates passed nine C-root repetitions.

Full runtime reruns were explicitly omitted at the user's request after the
earlier green qualification. This packet claims focused qualification only.
Earlier failed Full and focused receipts remain failed and are retained beside
this packet. Search seed/resource and export code are unchanged from C4; the
C4 resource/24-export results are historical component evidence, not relabelled
as fresh results on this source. The new popup binding path was tested in all
three x64 profiles, including stale-generation, input, teardown and resource
lifetime coverage. Build warnings are retained in predecessor/final receipts;
this is not a warning-free ARM64 claim.

The first public snapshot 397e15adf847aef4f9d5afe190b007d13ed5778e has the exact qualified source tree.
A fresh public-only clone verified private commit-object absence, complete
reachable commit-object inventory, immutable protected-plan bytes/ancestry and
the exact public DxUi dependency restore. This delivery commit adds only this
evidence, WIP/changelog bookkeeping and archive attributes. Public PR/main CI
and the primary checkout switch are subsequent delivery operations; their final
results are reported in the migration handoff. No pending CI is claimed passed.
Installed-helper/package/release/manual IME/AT qualification remains separate,
and the RedMonitor consumer plan therefore remains WIP.

The manifest binds complete native verdicts/traces, the six source-bound build
receipts, diagnostics, explicit scope waiver and the fresh public verification.
PrePush is planning-only and returned NOT_EVALUATED; it does not establish
repository qualification or equivalence to a future CI result.
