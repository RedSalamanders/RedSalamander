# IMAP UID enumeration: Debug failing baseline

Result: **260/348 assertions pass, 88 fail**, with clean native completion.
The runner reports two passed cases, one failed case and no skips; it did not
perform a retry/classification run. This is the failing baseline, not green evidence.

The production parser is unchanged from `5ed9032c4dd8b7f46e7be9dc53bf0e25ecaa06b3`.
`tested.patch` adds 35 four-check SEARCH response scenarios to the existing fixture;
`tested-source-hashes.json` retains the four exact working-file hashes.
Full test-enabled Debug receipt:
`9b34bd2e9d50209ff8e871b3f3155bee2d50b1a1ff7c6c6d273a5dcef1d6687c`.
Snapshot: `1fc681ecc47292496042a015097a330c1b303cd125c84b0cd5150081d7d479a0`.
Build duration: 214,465 ms, eleven C5246 warnings from the new fixture aggregate
braces and zero errors. The candidate must correct these warnings.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-CaseFilter C0_CurlImapListingTruth -TestRoot C:/RedSalamander.Perf
-RunId c0-imap-uids-baseline-debug-20260906`.

The scenarios protect exact UID tokens, case/empty/missing response distinctions,
unique descending message identities, opaque literal framing and zero metadata
requests after identity-enumeration failure. ESEARCH support is not implemented;
it must be reported as unsupported rather than an empty mailbox. Existing
1/201/4,097-message, metadata-repair and cancellation controls are retained.

Raw results/traces, full receipt, tested snapshot and ordered selected metrics
are retained with their provenance. The runner reports 211 disk-audit findings;
existing runtime data was not cleaned. This fixture is not live IMAP/TLS,
whole-response memory, mutation or final Fresh Full qualification.
