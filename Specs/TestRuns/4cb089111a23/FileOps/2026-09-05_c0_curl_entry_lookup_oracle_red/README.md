# Entry lookup — diagnostic run with an incorrect receipt oracle

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlEntryLookupTruth -FailFast -RunId c0-curl-entry-lookup-baseline-fixed-debug-20260905`.
Production is unchanged from `ec92c9fe`. Build 03:30, zero warnings/errors;
receipt `ed402767bf2e985ddaf40a4be973a7a27305344d963c6c011f1ade8233c9b073`;
workspace `7168f04bc79690387a2012840f03da9b45dd0f21d1c4924463c131f85a77cbf6`.

Result: 2 harness passes, 1 failure, zero skips; 99/197 native assertions pass.
The oracle incorrectly demanded a destructive receipt for Copy and retained-source
proof when admission never observed a trustworthy source. Those assertions were
corrected without changing production; this run is diagnostic, not the clean baseline.
The independent state/status assertions already witness the same 40 unsafe mutations.
Full results, reasons, traces and perf remain intact. Wide repeated lookups measured
456,011 us Unix and 349,911 us DOS. Disk audit: 76 issues, none deleted.
