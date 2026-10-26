# IMAP listing regression baseline

Source base: `9b56081f6fc67924ee62c1831f1de6a656f13d99` plus the six exact inputs
in `tested-source.json`. Full test-enabled Debug build receipt:
`fb9fed83c76573ce9e3903907c070c4f5fa86be0f9e3bee5a416799eca29ac53`;
snapshot `29cccf82755064683f944eeea1d17753f0483f209e34b6fb3fb86fa72fcd42d9`.
The build succeeded in 190,877 ms. This is deliberately failing baseline evidence.

Command:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -CaseFilter C0_CurlImapListingTruth -TestRoot C:\RedSalamander.Perf -RunId c0-imap-listing-baseline-debug-20260906
```

The selected native case reports 21/38 checks passed, 17 failed. Including harness
cases, the runner reports 2 passed, 1 failed, zero skipped. No retries or failure
reclassification were requested. The runner retained 194 disk-audit issues;
no historical files were removed to suppress them.

Four cancellation scenarios incorrectly return success and continue fetching.
Missing and malformed sizes become known values. For complete 1/201/4,097-message
listings the summary map peaks at 1/201/4,097 entries, instead of at most 200.
All three normal listings still return every requested message.

The fixture drives actual LIST/STATUS/SEARCH/FETCH parsing and repair orchestration
without credentials, sockets or saved-profile changes. It does not prove live TLS,
transport cancellation or total mailbox/transport memory bounds. Metric selections
preserve claim-bearing rows in original order with raw digest and line positions.
Build-log normalization and exact copied-file hashes are recorded separately.
