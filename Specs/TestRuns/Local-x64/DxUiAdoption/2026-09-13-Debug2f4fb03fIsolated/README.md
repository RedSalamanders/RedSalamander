# Local x64 Debug qualification at 2f4fb03f

Frozen RedSalamander `2f4fb03fffb9484783ac0badb47f2f9b9f61c739`, DxUi `13788e95`,
**fails** Fresh Full run `20260913T115725Z-105808-1f4f447e83b040d3ba36c8f21ee75a1a`. This is local execution; no hosted CI ran.
Latest master `e8a9f857` is an ancestor of this candidate; the five release-tooling
files match its Git blobs exactly, as recorded in `master-integration-2f4fb03f.json`.

The aggregate reports **2,120 passed, 2 failed and 52 skipped**;
18/20 entries are promoted. The aggregate duration is
3,940,533 milliseconds. Skipped capabilities remain unqualified.
Unpromoted entries: selftest.commands, tooling.pester.

| Product suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| CompareDirectories | 242 | 0 | 30 |
| Commands | 897 | 1 | 2 |
| FileOperations | 180 | 0 | 20 |



Original per-case results retain every passed, failed and skipped case and each skip reason.
The immutable validation plan and final decision accompany the aggregate. The final run state
retains entry fingerprints, promotion digests and the accepted build-receipt identity.

All 14 module identities were checked against actual output bytes after Full's build-tooling
re-attestation. They match the exact pin, configuration, public-header tree and x64 compiler host.

`retained-files.json` records the original hashes and locations of 16 reviewed receipts.
The complete native archives and per-entry validation evidence remain at
`Z:/RedSalamander.Perf/evidence/I19-20260909/2f4fb03f-isolated-x64-Debug-Full`. Previous failed attempts stay retained; this run does not erase them.

ARM64 and ASan qualification are explicitly user-deferred. The measured resource trade-off
is accepted; publication remains a separate release handoff. See the [I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).
