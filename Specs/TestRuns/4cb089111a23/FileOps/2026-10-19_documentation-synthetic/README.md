# Realistic synthetic screenshot data

The current File Operations guide uses invented Workshop media files and distinct locations under
`C:\Demo\Media\Workshop` and `D:\Demo\Archive\Workshop`. These are display-only fixture values;
no operation is admitted and no displayed file is opened. The popup and current selected-error
Issues pane were captured by the native harness without activation, global input or image editing.
The full frames were visually reviewed: readable paths, no personal data and no redaction labels.

`Specs/Testing/Testing_SelfTests.md` now requires realistic invented identifying data in every published
screenshot across all features, with the same rule in `AGENTS.md`. Redaction labels are superseded
by synthetic examples. The earlier documentation-privacy archive is historical capture evidence,
not the current guide illustration or current screenshot policy.

## Runs and verification

- First attempt `20261019T110614Z-13344-b0e85834f75b4ecd9112d6dba0f3cf1c`: 2 runner cases passed, 1 failed. The fixture incorrectly expected a plain diagnostic to provide a leaf filename and parent paths. The production Issues projection correctly retains full paths when qualified display metadata is absent. `first-attempt/` preserves the failed result, trace, source diff, source hashes, warning-free build receipt and metrics provenance. No Issues image from this attempt was published.
- Corrected run `20261019T111255Z-93580-75ec71ff7cd246e38b52f1c121b2c827`: 3 passed, 0 failed/skipped. Exact command: `Test-Changes.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts -Configuration Debug -SkipBuild -TestRoot D:\RedSalamander.Perf`. No interactive entries are required for this exact selection.
- Build: full solution, test-enabled x64 Debug, zero warnings/errors; receipt `6cbe868c8a7aa850d1197d80628c84b1afef8327399be5315c316758299eaee6`, source snapshot `6a6c606bfbb05f705d7accc2d32f707fa03a13422d03cb736c10bbf5b9b085e5`.
- Final source contracts: 204 passed, 0 failed/skipped. The final native case captures 2 documentation surfaces and all 132 existing card frames at actual 144 DPI, including all 124 positive percent/concurrency correctness witnesses.

Privacy/layout checks assert the exact popup subject and accessible source/destination paths, count
text and non-overlapping controls; the selected Issues row and exact full source/destination details
must use the invented values. Personal-account path markers, UNC/account identifiers and old
redaction labels are rejected by this fixture. Plain-diagnostic fallback is exercised without
manufacturing conflict metadata. Capture failures now have their own failure reason and trace.

`captures/` contains the exact native bytes published under `docs/res/`; independent sidecars bind
image hashes, fixture values, display profile and build provenance. Native scenario manifests,
source identities, validation explanation, aggregate results and final source snapshots are retained.
`documentation-review.json` records the post-publication image/provenance consistency review.

Complete raw runs are preserved under `D:\RedSalamander.Perf\evidence\fileops-documentation-synthetic-20261019-first`
and `...-final`. The final metric stream exceeds the checked-in per-file limit; both streams'
hashes, sizes and row counts are retained. `card-frames.jsonl` retains all 1,488 final card-frame metric rows in
original order. No product renderer or operation code changed, so this is capture correctness
validation, not a new performance comparison or optimization claim.

## Scope and publication boundary

The build and final runner used the same complete source snapshot. No source/index edits occurred
while either build or runner was active. Capture code, normative policy, agent guidance and DxUi
pin remained unchanged after the final run. Images, provenance, plan closeout and evidence were
published only after the runner exited; those documentation-only updates change the final workspace
identity, so this receipt does not authorize exact reuse against it.

This focused result is not Full, PrePush, Release, ARM64, ASan, foreground-input, storage-performance
or CI qualification. The existing shared sandbox audit reports 274 issues and is not clean.

Archive closeout: the explicit pre-stage check passed for both documentation archives. The full
checked-in archive inventory passed for 15,507 files; its output is retained in `archive-inventory.txt`.
