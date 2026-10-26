# File Operations card information hierarchy

Status: complete for the requested layout and scoped x64 Debug qualification. Owner request: 2026-10-16.

## Checklist

- [x] Capture a deterministic baseline using owned harness windows without activation.
- [x] Retain the original single-root name or top-level selection counts through completion and collapse.
- [x] Implement paired header/status and progress/percent/rate/ETA rows, top progress bar, paths, throughput, actions and Technical details.
- [x] Cover narrow widths, disclosure transitions, long names, folders, mixed/unknown selections, terminal states and localization.
- [x] Update authoritative UI contract and user documentation.
- [x] Build, run affected checks, inspect harness screenshots and archive measured rendering evidence.
- [x] Move this completed plan to Done.

## Contract and validation

The subject is admission identity, not the current child or recursive discovery count. Normal expanded cards use the horizontal space for paired short facts. Paths and throughput retain full-width rows. Technical details is a task-owned accessible disclosure; automatic concurrency is secondary text there. Conflict decisions and precise terminal facts retain priority. Throughput keeps its current expanded default and remembered preference.

Extend the existing reviewed no-activation `FileOps_ReviewRemediationContracts` case with owned-window synthetic layout captures. Use the existing screenshot and theme helpers, no global focus/input assertions. Capture identical 480/960 DIP scenes before and after layout changes, in English/French and dark/light. Record `FileOps.Popup.Render.*` and an explicit warmed layout batch metric; bound layout work and check retained controls. Evidence stays under the marked test root and is promoted before another runner invocation removes previous runs. No throughput improvement is claimed from UI timings.

The earlier queued missing-source correction remains part of the working tree and must continue passing its existing ten fixtures.

## Closeout evidence

The authoritative contract is merged into `Specs/UI/UI_FileOperationsPopup.md`,
with the missing-source facts in `Specs/FileSystem/FileSystem_FileOperations.md`.
The user guide and its native harness screenshot are refreshed in `docs/FileOperations.md`.

The compact archive is
`Specs/TestRuns/4cb089111a23/FileOps/2026-10-16_card-hierarchy-closeout/README.md`.
It retains 60 deterministic layout captures (English/French, 480/960 DIP, actual
144 DPI), 20 live-copy captures, the two build receipts, source diff, raw rendering
samples and summaries, and the 214/214 localization/source-contract result.

- `FileOps_ReviewRemediationContracts`: run `20261016T173738Z-83008-21037ccaa05048f38559cb8314746aa0`, passed, including the ten missing-source fixtures and geometry/disclosure checks.
- `DiscoveryScope_FolderLiveProgress`: run `20261016T174403Z-41384-034b5e0cebd6459b8554f6aa0b1263d5`, passed, with 64 behavioral assertions across six local/simulated-provider routes.
- Both x64 Debug application builds completed without compiler warnings/errors.
- The compact TestRuns archive passes the standard size and evidence contract without overrides.
- The 624 warmed frame samples have median 8.099 ms and p95 8.497 ms; synchronous render median is 1.929 ms and p95 6.069 ms. These are characterization, not a matched before/after claim.

These focused results do not qualify Full, Release, ARM64, ASan, foreground input,
other DPI profiles, real SMB storage performance, or long-run retention. The test
disk audit remains unclean from unrelated existing data; the live run also
recorded the completed concurrent Pester scratch directory, which was subsequently
removed with the governed exact-run cleanup helper. Historical reports are retained.
