# Synthetic data in application screenshots

- [x] Replace the redaction-only File Operations fixture with realistic invented Workshop media names and distinct source/destination paths; assert the actual subject, accessible paths and selected issue identities.
- [x] Make realistic synthetic data mandatory for every published screenshot in the normative testing specification and canonical agent guidance; align the guide text.
- [x] Build and run the no-activation native capture case, retaining the existing card checks; inspect both new frames.
- [x] Publish the two unedited frames with current provenance, archive validation and review the guide/spec consistency.
- [x] Move this plan to Done.

This follows the earlier privacy capture work. Its hidden-label frames remain historical evidence;
the current guide will use realistic invented data instead. No product UI or renderer changes are
needed. Example paths are display-only and never grant test I/O outside the marked sandbox.
No desktop input is required. Focused x64 Debug validation does not imply Full or other profiles.

Closeout: the final test-enabled x64 Debug full-solution build passed with no warnings or errors.
Run `20261019T111255Z-93580-75ec71ff7cd246e38b52f1c121b2c827` passed all three runner cases,
captured both current documentation surfaces and preserved all 132 existing card captures.
All 204 source-contract checks passed. The first fixture-expectation failure and original build
receipt remain archived; its incorrect filename/parent assumption was corrected to the plain
diagnostic's full-path fallback without changing the product UI.

The two native images were visually inspected and published byte-for-byte, with current hashes,
scenario data, DPI and build/source provenance. General requirements are in
`Specs/Testing/Testing_SelfTests.md` and `AGENTS.md`. Evidence is in
`Specs/TestRuns/4cb089111a23/FileOps/2026-10-19_documentation-synthetic/`.
Only image/provenance, evidence and plan closeout changed after the final runner exited.
No broader runtime/platform/performance qualification is claimed.
