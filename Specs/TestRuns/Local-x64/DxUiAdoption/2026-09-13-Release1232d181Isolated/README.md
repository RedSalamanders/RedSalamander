# Local x64 Release qualification at 1232d181

Frozen RedSalamander `1232d181e74ae7118d3fce7974196811b3cd4398`, DxUi `13788e95`,
passes Fresh Full run `20260913T143644Z-39692-35d3bb90eb2f4f908ced437115524c37`. This is local execution; no hosted CI ran.
Latest master `e8a9f857` is an ancestor of this candidate; the five release-tooling
files match its Git blobs exactly, as recorded in `master-integration-2f4fb03f.json`.

The aggregate reports **2,118 passed, 0 failed and 56 skipped**;
20/20 entries are promoted. The aggregate duration is
3,248,520 milliseconds. Skipped capabilities remain unqualified.

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

`retained-files.json` records the original hashes and locations of 12 reviewed receipts.
The complete native archives and per-entry validation evidence remain at
`Z:/RedSalamander.Perf/evidence/I19-20260909/1232d181-isolated-x64-Release-Full`. Previous failed attempts stay retained; this run does not erase them.

ARM64 and ASan qualification are explicitly user-deferred. The measured resource trade-off
is accepted; publication remains a separate release handoff. See the [I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).
