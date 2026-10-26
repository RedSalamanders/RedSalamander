# Local x64 Release qualification at a20f6680

Frozen RedSalamander `a20f6680dd5f82835a549f4c231e5666e4047eda`, DxUi `13788e95`,
**fails** Fresh Full run `20260913T034047Z-88456-99c436c59afe45f698e35f3b19b2dd59`. This is local execution; no hosted CI ran.

The aggregate reports **2,105 passed, 8 failed and 56 skipped**;
19/20 entries are promoted. The aggregate duration is
3,287,810 milliseconds. Skipped capabilities remain unqualified.
Unpromoted entries: tooling.pester.

The eight failures are module-scope lookup errors in the general tooling entry
(762 passed, eight failed). The combined local driver loaded same-named modules
from two checkouts. The [isolated module probe](../2026-09-13-LocalDriverIsolation/README.md)
reproduces that error with both copies and succeeds with one. All 12 build-tooling
checks, both viewer entries and performance entries pass. All five ViewerText
regressions corrected at `836f4f97` pass in Commands. This remains a failed Full;
the clean-process retry uses the same source with no product or assertion changes.

| Product suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| CompareDirectories | 239 | 0 | 33 |
| Commands | 898 | 0 | 2 |
| FileOperations | 179 | 0 | 21 |

Original per-case results retain every passed, failed and skipped case and each skip reason.
The immutable validation plan and final decision accompany the aggregate. The final run state
retains entry fingerprints, promotion digests and the accepted build-receipt identity.

All 14 module identities were checked against actual output bytes after Full's build-tooling
re-attestation. They match the exact pin, configuration, public-header tree and x64 compiler host.

`retained-files.json` records the original hashes and locations of 11 reviewed receipts.
The complete native archives and per-entry validation evidence remain at
`Z:/RedSalamander.Perf/evidence/I19-20260909/a20f6680-final-x64-Release-Full`. Previous failed attempts stay retained; this run does not erase them.

Native ARM64 execution, measured resource acceptance and publication remain separate gates
in the [I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).
