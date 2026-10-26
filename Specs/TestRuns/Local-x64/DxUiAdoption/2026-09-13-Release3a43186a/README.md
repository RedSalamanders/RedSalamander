# Local x64 Release qualification at 3a43186a

Frozen RedSalamander `3a43186a3507ca48eb41e7ad8dec8585aed732f9`, DxUi `13788e95`,
**fails** Fresh Full run `20260913T020805Z-95632-7863a01c443845c68ff6b908af4aaf69`. This is local execution; no hosted CI ran.

The aggregate reports **2,108 passed, 5 failed and 56 skipped**;
19/20 entries are promoted. The aggregate duration is
3,364,005 milliseconds. Skipped capabilities remain unqualified.
Unpromoted entries: selftest.commands.

| Product suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| CompareDirectories | 239 | 0 | 33 |
| Commands | 893 | 5 | 2 |
| FileOperations | 179 | 0 | 21 |

Original per-case results retain every passed, failed and skipped case and each skip reason.
The immutable validation plan and final decision accompany the aggregate. The final run state
retains entry fingerprints, promotion digests and the accepted build-receipt identity.

All 14 module identities were checked against actual output bytes after Full's build-tooling
re-attestation. They match the exact pin, configuration, public-header tree and x64 compiler host.

`retained-files.json` records the original hashes and locations of 11 reviewed receipts.
The complete native archives and per-entry validation evidence remain at
`Z:/RedSalamander.Perf/evidence/I19-20260909/3a43186a-final-x64-Release-Full`. Previous failed attempts stay retained; this run does not erase them.

Native ARM64 execution, measured resource acceptance and publication remain separate gates
in the [I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).

All five failures concern ViewerText snapshots that were compiled out by `_DEBUG` in
test-enabled Release. The [focused correction record](../2026-09-13-ViewerTextReleaseDiagnostics/README.md)
reproduces 0/5/0 before and passes 5/0/0 after the `ENABLE_TESTS` correction. This does not
convert this failed Full into a pass; a fresh final run is required.
