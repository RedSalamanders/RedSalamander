# File Operations documentation screenshot privacy

Historical iteration: these hidden-label frames were superseded by the realistic synthetic captures
in `../2026-10-19_documentation-synthetic/`. They are retained as original evidence, not current guide
illustrations. The current normative policy requires invented names and paths for all screenshots.

The user guide previously embedded an identifying media filename/path, an obsolete completed-task card under Issues, and a Properties name/path example. The guide now embeds two current native frames whose item identities and locations are explicitly hidden. The Properties text remains without its illustration; the historical stress-gallery link is removed from this guide.

## Validation

- Build: test-enabled full solution, x64 Debug, zero warnings/errors. Receipt `879ae09d2d858c3599140587beff520d0f4ac6a9ecb7f4a037b1386aaf2d959b`; source snapshot `0737e5676b188a6c1e4bbd23d58a155f5b54e8bf2b7244bebf663f6dab6a7b11`.
- Candidate run: `20261019T105155Z-35676-105649eb7ccd4ba0a96e2eaa11b64964`. `Test-Changes.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts -Configuration Debug -SkipBuild -TestRoot D:\RedSalamander.Perf` passed all 3 runner cases, with 0 failed/skipped and 0 pending interactive entries for this exact selection.
- Captures: 2 documentation frames plus all 132 existing card frames at actual 144 DPI. Both new frames were visually inspected before publication. No activation, desktop input, manual screenshot or image editing was used.
- Privacy assertions: exact retained popup subject/source/destination identities, recursive count text and no button overlap; one selected current Issues row, details present with hidden source/destination, and no absolute drive/UNC paths or account identifiers in selected cells/details. Fixtures have private operation state and an empty owned breadcrumb store; no operation is admitted.
- Source contracts: 204 passed, 0 failed/skipped. Whole checked-in archive inventory passed for 15,456 files before adding this archive; the completed new archive is separately checked before staging.

`captures/` contains the two unedited native frames copied byte-for-byte into `docs/res/`. Their sidecars bind image hashes, scenario, build receipt, theme, window size, DPI and privacy policy. `documentation-scenarios.tsv` and `card-scenarios.tsv` retain the native manifests. `source-identity.json` binds both test sources, the normative capture spec and the unchanged DxUi pin.

The full raw run is preserved at `D:\RedSalamander.Perf\evidence\fileops-documentation-20261019-final`. The raw metric stream exceeds the checked-in per-file limit. `card-frames.jsonl` retains all 1,488 existing card-frame metric rows in original order; `metric-stream-provenance.json` binds the full stream and retained rows. No renderer or operation code changed, so there is no new optimization comparison or storage-performance claim; the earlier renderer baseline/candidate analysis remains in the card-counts-throughput archive.

## Boundaries

This is focused documentation-capture validation, not Full, PrePush, Release, ARM64, ASan, foreground-input or storage-performance qualification. The shared test-root audit reports 274 existing issues and is not clean. No CI pass is claimed.

Source, index and documentation inputs stayed frozen throughout the build and complete runner invocation; final source snapshot equals the build receipt snapshot. Only image/provenance publication, evidence and plan closeout changed afterward. The final working-tree snapshot therefore differs, and these receipts do not authorize exact reuse against it. Capture sources, normative capture requirements and the DxUi pin remained unchanged after validation.

Durable capture/privacy requirements are in `Specs/Testing/Testing_SelfTests.md`; the completed plan is in `Specs/Plans/Done/FileOperations_DocumentationPrivacy_2026-10-19.md`.
