# IMAP listing candidate — Debug

The focused candidate passes 70/70 assertions (runner: 3 passed, zero failures/skips).
Broad C0 passes 18/0/0 in the [companion archive](../2026-09-06_c0_imap_listing_broad_debug/README.md).
The 12 existing standalone IMAP helper functions pass; all 177 attested artifacts
were checked before their launch. Source/resource policy passes 197/197, no skips.

Source base: `9b56081f6fc67924ee62c1831f1de6a656f13d99` plus the exact ten inputs in
`tested-source.json`. Snapshot:
`fa008de406f84ba9fd25319d492403a4e60f8cc1059ca8e34daf520fcd2675b9`.
Full test-enabled Debug receipt:
`bf72858a59745777812fcc204aae08b44df65d5ea88035d9e9bf84e967a68275`.
Build: 186,403 ms, zero warnings/errors. No source/test/tool changes occurred
between this build and the focused, helper and broad runs.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -CaseFilter C0_CurlImapListingTruth -TestRoot C:\RedSalamander.Perf -RunId c0-imap-listing-candidate-debug-20260906
```

The [baseline](../2026-09-06_c0_imap_listing_baseline_debug/README.md) passed 21/38
with 17 failures. The candidate retains those assertions and adds 32 assertions
covering unsolicited unrelated UIDs/FLAGS updates, one-shot control errors,
deadlines and overflowing/signed sizes. No failed assertion was removed.

`baseline-comparison.json` pairs all 12 original scenarios; the candidate adds ten.
For 1/201/4,097 complete messages, retained summary counts change from 1/201/4,097
to 1/200/200. Request counts remain 1/2/37. Debug durations change from
1,031/5,478/110,785 to 1,097/5,811/112,616 us. These single observations include
fixture work and the candidate's additional no-op control fixture; they are not
a throughput improvement claim. Cancellation-phase request counts change from
2/5/5/15 to 1/2/3/2, with cancellation HRESULT and no partial returned entries.

The existing parser/repair owner now builds entries per bounded summary chunk,
retains unknown-size truth, ignores unsolicited unrelated UIDs, preserves metadata
on FLAGS-only updates, and stops on cancellation/deadline/control failure.
Targeted lookup and Properties do not synthesize unknown sizes as zero.

This does not qualify all IMAP buffering, live libcurl/TLS behavior, real server
quotas or mutation routes. Quoted/literal FETCH attribute-key matching and partial
summary repair need further coverage. No full-plan gate is closed by this archive.
The focused runner reported 196 historical disk-audit issues; they were not removed.
All raw source copies, selected metrics and copied runtime files have digest provenance.
Later plan/consistency/README changes are evidence annotations, not a new native run.
The frozen `.gitattributes` input contains three pre-existing trailing blanks on
comment lines. Its narrow archive-only whitespace exception preserves exact input
bytes and hashes; production source and other evidence retain normal hygiene checks.
