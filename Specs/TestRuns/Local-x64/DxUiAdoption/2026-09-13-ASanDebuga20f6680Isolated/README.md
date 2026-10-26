# Local x64 ASan Debug qualification at a20f6680

Frozen RedSalamander `a20f6680dd5f82835a549f4c231e5666e4047eda`, DxUi `13788e95`,
**fails** Fresh Full run `20260913T054610Z-95756-605d08f0a71e45ac91d7d81bac52cdab`. This is local execution; no hosted CI ran.

The aggregate reports **1,831 passed, 8 failed and 63 skipped**;
16/20 entries are promoted. The aggregate duration is
5,548,257 milliseconds. Skipped capabilities remain unqualified.
Unpromoted entries: selftest.compare-directories, selftest.commands, selftest.file-operations, standalone.red-configure.

| Product suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| CompareDirectories | 241 | 1 | 30 |
| Commands | 632 | 0 | 0 |
| FileOperations | 162 | 5 | 33 |

Native archives are absent for: selftest.commands. The table retains partial entry-result counts, not successful suite completion. Original terminal/evidence records preserve the incomplete coverage and failed entry status.


Original per-case results retain every passed, failed and skipped case and each skip reason.
The immutable validation plan and final decision accompany the aggregate. The final run state
retains entry fingerprints, promotion digests and the accepted build-receipt identity.

All 14 module identities were checked against actual output bytes after Full's build-tooling
re-attestation. They match the exact pin, configuration, public-header tree and x64 compiler host.

The native seeded heap-overflow probe exits nonzero with the expected AddressSanitizer diagnostic. Its executable hash and original log are retained.

`retained-files.json` records the original hashes and locations of 28 reviewed receipts.
The complete native archives and per-entry validation evidence remain at
`Z:/RedSalamander.Perf/evidence/I19-20260909/a20f6680-isolated-x64-ASan-Debug-Full`. Previous failed attempts stay retained; this run does not erase them.

Native ARM64 execution, measured resource acceptance and publication remain separate gates
in the [I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).

The additional File Operations evidence includes its original trace and exact JSONL
metric excerpts with the complete source-file hash and selection rule. R0d completes
all counts successfully but takes 2,404,647 us against 2,000,000 us. Curl's serial
4,097-file copy hits its 180-second deadline after 3,280 publications. These are
observed failures, not accepted performance baselines. Source comparisons record
the unchanged Curl implementation and the baseline fixtures; the capability matrix
matches after canonical clang-format 22 formatting. Candidate diagnosis remains open.
