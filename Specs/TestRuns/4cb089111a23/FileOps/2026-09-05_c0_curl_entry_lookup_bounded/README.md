# Bounded strict entry lookup — focused pass

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlEntryLookupTruth -FailFast -RunId c0-curl-entry-lookup-bounded-debug-20260905`.
Based on regression commit `d9e0093e`. Build 03:07, zero warnings/errors.
Receipt `bca6184473f97b7aaeaf22a6af772e50944d5532b35a9887d03296c93f72c1a3`;
workspace `0439d4b0f9cde5bac34ac0bfce9bdabd7fe82bddb64e7ec517d7dad10f887c66`.

**197/197 assertions; 3/3 harness cases, zero skips.** The 40 unsafe admissions
from the [clean baseline](../2026-09-05_c0_curl_entry_lookup_red/README.md) now
refuse without mutation. Literal/missing controls and all eight cancellation routes
pass. Cancellation takes 55,538–59,533 us versus roughly one second before.

GetEntryInfo now drains the strict cursor and publishes only a complete unique
match, retaining the current row and candidate. Bulk operations use that same
owner; the whole-parent cache and its credential-bearing key/hash are removed.
Directory-size root lookup reuses the helper with its own progress checkpoint.
Unrelated rows skip timestamp conversion; full requested metadata remains exact.

Same-machine 16-query/4,097-row probes: Unix 460,448 → 386,833 us; DOS 381,552 →
350,564 us. Candidate additionally asserts exact ordinary-parser timestamp and
4,097 observed rows. Owned reservations peak at 208 path bytes and 51,400 metadata
bytes, excluding opaque transport state. This modest timing comparison does not
prove wide-copy throughput or whole-operation memory.

Full perf retained: 155,763 bytes, 424 rows, SHA-256
`cfd5a5f85f1c315c1576cea5c29e00ec765719266a3e023c7699baebc2c6eceb`.
Native case 5,813 ms; suite 7,406 ms. Disk audit 78 issues, none deleted.
Source/localization 197/197 (19.05 s); spec inventory zero blockers.
The [same-source broad run](../2026-09-05_c0_curl_entry_lookup_connection_red/README.md)
exposes connection churn. This focused pass does not qualify the candidate or Done.
