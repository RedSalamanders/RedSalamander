# Local x64 Debug qualification at a8350f4b

Frozen RedSalamander `a8350f4bf3e8d18cd666c720e2628167b4eb3af3`, DxUi `13788e95`,
**fails** Fresh Full run `20260913T103734Z-66012-6055b950b9de426fb4e3587e66330fd6`. This is local execution; no hosted CI ran.

The aggregate reports **2,119 passed, 1 failed and 52 skipped**;
19/20 entries are promoted. The aggregate duration is
3,955,407 milliseconds. Skipped capabilities remain unqualified.
Unpromoted entries: selftest.commands.

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

`retained-files.json` records the original hashes and locations of 13 reviewed receipts.
The complete native archives and per-entry validation evidence remain at
`Z:/RedSalamander.Perf/evidence/I19-20260909/a8350f4b-isolated-x64-Debug-Full`. Previous failed attempts stay retained; this run does not erase them.

ARM64 and ASan qualification are explicitly user-deferred. The measured resource trade-off
is accepted; publication remains a separate release handoff. See the [I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).
