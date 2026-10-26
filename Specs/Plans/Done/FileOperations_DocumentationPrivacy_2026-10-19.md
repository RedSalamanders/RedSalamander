# File Operations documentation screenshot privacy

- [x] Review every image embedded in `docs/FileOperations.md`: popup contains a media filename; Issues image is an obsolete completed card; Properties shows a name/path example.
- [x] Add dedicated no-activation native popup and current Issues-pane fixtures with hidden identity placeholders and deterministic privacy assertions. Use private operation state and an empty owned breadcrumb store.
- [x] Document the capture/privacy contract; remove the Properties illustration and historical stress-gallery link from this guide.
- [x] Build and run `FileOps_ReviewRemediationContracts`; inspect the generated native frames, preserving all previous card checks.
- [x] Replace both guide images, bind provenance, archive the run and review documentation consistency.
- [x] Move this completed plan to Done.

Scope: documentation and test fixtures only; no product UI behavior or renderer change. No manual
desktop capture or image editing. No foreground input is required. All runtime artifacts remain
under the existing marked `D:\RedSalamander.Perf` sandbox. Native capture validation does not
constitute Full, Release, ARM64, ASan or storage-performance qualification.

Closeout: test-enabled x64 Debug full-solution build passed with zero warnings/errors. Focused
`FileOps_ReviewRemediationContracts` run `20261019T105155Z-35676-105649eb7ccd4ba0a96e2eaa11b64964`
passed all three runner cases and retained all 132 previous card captures plus the two new documentation
frames. Both new frames were visually inspected, contain only hidden identities, and are published
byte-for-byte with independent provenance sidecars. TestHarnessSourceContracts passed 204/204.

Durable privacy/capture requirements are in `Specs/Testing/Testing_SelfTests.md`. Evidence and source
bindings are in `Specs/TestRuns/4cb089111a23/FileOps/2026-10-19_documentation-privacy/`. The complete
raw run remains under the marked sandbox. Build and final runner source snapshots match; only image,
provenance, evidence and plan closeout changed after the runner exited. Capture sources, normative
requirements and the DxUi pin stayed unchanged. These are focused capture results, not broader
runtime/platform qualification. The shared test-root audit still reports 274 existing issues.

Follow-up: the user requested realistic synthetic names and paths instead of hidden labels.
`ScreenshotSyntheticData_2026-10-19.md` records that replacement and the general screenshot policy.
This plan and its original captures describe the earlier iteration; the authoritative testing
specification and current guide use invented Workshop media data.
