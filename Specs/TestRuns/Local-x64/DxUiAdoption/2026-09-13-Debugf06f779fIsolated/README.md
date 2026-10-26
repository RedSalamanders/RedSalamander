# Local x64 Debug qualification at f06f779f

Frozen RedSalamander `f06f779fd45105d251c2761a82138e7c8215326e`, DxUi `13788e95`,
**fails** Fresh Full run `20260913T090138Z-94432-6d3ad46dee364289838dc5503e04dccd`. This is local execution; no hosted CI ran.

The aggregate reports **2,118 passed, 2 failed and 52 skipped**;
19/20 entries are promoted. The aggregate duration is
3,949,148 milliseconds. Skipped capabilities remain unqualified.
Unpromoted entries: selftest.commands.

| Product suite | Passed | Failed | Skipped |
|---|---:|---:|---:|
| CompareDirectories | 242 | 0 | 30 |
| Commands | 896 | 2 | 2 |
| FileOperations | 180 | 0 | 20 |



Original per-case results retain every passed, failed and skipped case and each skip reason.
The immutable validation plan and final decision accompany the aggregate. The final run state
retains entry fingerprints, promotion digests and the accepted build-receipt identity.

All 14 module identities were checked against actual output bytes after Full's build-tooling
re-attestation. They match the exact pin, configuration, public-header tree and x64 compiler host.

`retained-files.json` records the original hashes and locations of 14 reviewed receipts.
The complete native archives and per-entry validation evidence remain at
`Z:/RedSalamander.Perf/evidence/I19-20260909/f06f779f-isolated-x64-Debug-Full`. Previous failed attempts stay retained; this run does not erase them.

ARM64 and ASan qualification are explicitly user-deferred. The measured resource trade-off
is accepted; publication remains a separate release handoff. See the [I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).

## Retained failures and diagnosis

Commands reports 896 passed, two failed and two skipped. The failed cases are
`cmd_pane_navigation_refresh_keeps_navigation_shell_stable` and
`cmd_app_openDriveMenus_keeps_navigation_shell_stable`. Their original diagnostics
are evaluated in the same function call as the waits, so the snapshot/focus values
may precede the failed wait. They do not establish an empty-path or foreign-focus cause.
Diagnostic-only commit `e9099fb7` orders those reads after the waits and includes
baseline counters and foreground process identity. It changes no assertion or deadline;
its native qualification follows this retained run.

The two fixture functions are byte-identical between master `0dd0bd6d` and `f06f779f`.
`f06f779f-navigation-fixture-identities.json` retains the exact function hashes;
fixture identity does not prove product equivalence or explain the failures.
Build-tooling tests pass 12/0 and general tooling tests pass 773/0. Compare passes
242/0/30, File Operations passes 180/0/20, and all other suite entries pass.
