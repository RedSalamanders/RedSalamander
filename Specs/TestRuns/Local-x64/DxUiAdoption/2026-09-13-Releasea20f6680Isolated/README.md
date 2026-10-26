# Local x64 Release qualification at a20f6680

Frozen RedSalamander `a20f6680dd5f82835a549f4c231e5666e4047eda`, DxUi `13788e95`,
passes Fresh Full run `20260913T044529Z-83108-0cc8f2230c0d4d2ab341a44b5e1cdb3b`. This is local execution; no hosted CI ran.

The aggregate reports **2,113 passed, 0 failed and 56 skipped**;
20/20 entries are promoted. The aggregate duration is
3,283,811 milliseconds. Skipped capabilities remain unqualified.

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
`Z:/RedSalamander.Perf/evidence/I19-20260909/a20f6680-isolated-x64-Release-Full`. Previous failed attempts stay retained; this run does not erase them.

Native ARM64 execution, measured resource acceptance and publication remain separate gates
in the [I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).
