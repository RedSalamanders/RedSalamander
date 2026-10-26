# Native entry lookup — clean red baseline

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlEntryLookupTruth -FailFast -RunId c0-curl-entry-lookup-red-debug-20260905`.
Production remains `ec92c9fe`; only fixture, registration and test spec changed.
Build 03:17, zero warnings/errors. Receipt:
`7dd674440b5edbde44f5e4373e1facda0fa7c5d74f9cd9f802aad3efb662f458`.
Workspace: `9e15007ee02e6e566f977029d6326f0c4a365ef27ecf79e1502bed2cafc65b63`.

Result: 2 harness passes (Setup/Cleanup), one failure, zero skips; native checks
**117 passed / 80 failed**. All 64 route scenarios completed. Every single/bulk
Copy/Move/Delete/Rename route accepts each of five unsafe parent shapes: malformed
before/after the match, duplicate match, out-of-parent name and oversized row.
Those 40 scenarios mutate the fixture and return success instead of refusing.
Each fails both exact-state/status and completion assertions. Literal names,
missing source, cancellation and both wide controls pass. Cancellation is about
one second and is cooperative, with zero mutation commands.

Sixteen 4,097-row lookups take 460,448 us (Unix) and 381,552 us (DOS), with 16
LIST requests each and nonzero requested timestamps. These are a same-machine
baseline, not a throughput improvement. Full perf, all results/reasons and both
traces are retained without filtering. Raw perf is 145,500 bytes.

The [earlier oracle run](../2026-09-05_c0_curl_entry_lookup_oracle_red/README.md)
and [initial compile failure](../../Continuation/2026-09-05_c0_curl_entry_lookup_build_red/README.md)
are retained. The corrected oracle requires Copy's status-only completion and
conservative unknown-source destructive receipts after failed admission. No
originalStillPresent or safe-Retry proof is fabricated from a failed lookup.

Source/localization policy: 197/197, 20.04 seconds; spec inventory: zero blockers.
Disk audit: 77 issues, none deleted. Next: bounded strict lookup, remove whole-parent
bulk cache, preserve requested metadata and repeat this gate plus native wide Copy.
C0 and Done remain open.
