# Operation FileOperations Fault-Injection Ratchet - 2026-07-05

Owner for the long-term remainder routed from Granite GR-A5 after the first
interface-boundary bridge IO decorator slice landed.

## Status

Active follow-up. This plan is not a blocker for closing Granite; it owns the
opportunistic migration of remaining FileOperations self-test fault hooks.

## Baseline

Current baseline at routing time:

- `RedSalamander/FolderWindow.FileOperations.State.cpp`: 31 `ForSelfTest`
  occurrences.
- `RedSalamander/FolderWindow.FileOperations.State.cpp`: 30
  `#ifdef ENABLE_TESTS` blocks.
- `RedSalamander/`: 629 `ForSelfTest` occurrences.
- `RedSalamander/`: 647 `#ifdef ENABLE_TESTS` blocks.

The first Granite GR-A5 ratchet slice established
`SelfTestBridgeIoDecorator` / `SelfTestBridgeFileReader`, injected through
`DecorateBridgeIoForSelfTest(fileSystemIo, SelfTestBridgeIoRole::Source)`, and
migrated the cross-filesystem bridge source `GetSize` fail-next hook out of the
inline `CopyFileWithBuffer(...)` consumption path.

## Rules

- New FileOperations bridge data-safety fault-injection tests should use the
  interface-boundary decorator seam when the failure belongs to `IFileSystemIO`
  or `IFileReader`.
- Do not add new inline production-pipeline globals/call-site hooks in
  `FolderWindow.FileOperations.State.cpp` unless the hook cannot be expressed at
  an existing boundary; document any exception in this plan.
- When a touched pipeline site still has an inline self-test hook, migrate it to
  the decorator seam or a shared helper before adding adjacent coverage.
- Track the baseline counts above; a ratchet slice should reduce the relevant
  count or explain why a helper rename/centralization changed the count while
  still reducing pipeline coupling.

## Open Ratchet Slices

| ID | Status | Slice | Direction |
|----|--------|-------|-----------|
| FIR-1 | OPEN | Destination-side bridge `GetSize` fault injection, currently tracked as Floodgate FG-A1. | Coordinate with Floodgate: implement the destination `GetSize` failure through `SelfTestBridgeIoDecorator` with `SelfTestBridgeIoRole::Destination`, add the RED FileOps coverage required by FG-A1, then update both plans. |
| FIR-2 | OPEN | Bridge file-copy failure hook still consumed inline near the copy loop. | When next touching bridge copy error handling, move the failure into a boundary decorator/writer seam instead of checking a global at the operation site. |
| FIR-3 | OPEN | Destination mutation and create-directory race hooks still live as pipeline call-site helpers. | When touching move cleanup or bridge directory creation, convert the hook to a shared boundary helper or decorator-capable role and preserve the existing deterministic FileOps coverage. |
| FIR-4 | OPEN | Count ratchet and source-contract guard. | Add or extend a source-contract check that prevents new inline FileOperations bridge fault hooks from being added without updating this plan. |

## Closeout Gate

This plan can move to Done when the remaining FileOperations bridge fault
injection hooks either use boundary seams or have documented exceptions, the
ratchet counts have not regressed, and the relevant FileOps/Floodgate focused
cases plus `TestHarnessSourceContracts.Tests.ps1` pass.
