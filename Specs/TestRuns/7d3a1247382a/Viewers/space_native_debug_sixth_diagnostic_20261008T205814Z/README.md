# ViewerSpace native Debug focused diagnostic — 2026-10-08

This archive preserves focused diagnostic outputs from run `20261008T205814Z-10476-05ebb0b8203840de8291ee1f2a2a0033`. It records x64 Debug, build receipt `b93b682663ee44fb758b85354772fe46b044e0eb519350be64b447b3a1fc90ca`, and build source snapshot `368330d0b4123882b5af27128fb8969337f55b85e63250defd40b4af34a11761`. The focused JSON reports `repository_verdict: NOT_EVALUATED`; this archive does not qualify the repository or claim an overall pass.

`manifest.txt` binds the SHA-256, byte length, archived filename and original source path for all 23 copied artifacts: eight exact self-test output logs, the focused diagnostic JSON, and seven lifecycle PNGs with their seven provenance JSON files. Copies were verified byte-for-byte. The archive-local `.gitattributes` keeps the preserved evidence bytes unchanged in Git. Failure logs and invalid capture provenance are retained unchanged.

## Focused case outcomes

The focused diagnostic reports 3 cases with exit code 0 and 5 with exit code 1:

| Case | Activation | Exit |
| --- | --- | ---: |
| `TestViewerSpaceScanPolicyHelpers` | noninteractive | 0 |
| `TestViewerSpaceAnalysisHostVirtualGridAndLifetime` | noninteractive | 1 |
| `TestViewerSpaceLifecycleCaptures` | noninteractive | 1 |
| `TestViewerSpaceListingValidation` | noninteractive | 0 |
| `TestViewerSpaceBoundedHostileProviderScanning` | noninteractive | 1 |
| `TestViewerSpaceBlockedProviderCloseAndPostUpdateRace` | noninteractive | 0 |
| `TestViewerSpaceDirectorySummaryProtocol` | noninteractive | 1 |
| `TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses` | interactive, desktop lease | 1 |

## Capture limits

The lifecycle whole-frame PNGs look visually clean. They are **diagnostic evidence pending corrected harness DPI/probe handling, not qualification captures**: provenance reports 96 DPI and a 1200×720 window rectangle, while the PNG frames are 1800×1080 pixels (144-DPI physical scale). The recorded probe coordinate basis therefore does not match the captured physical pixels. In addition, lifecycle provenance for scenarios 02 through 06 has `valid: false` and capture HRESULT `0x8007000D`; scenarios 01 and 07 report valid captures. Preserve these failures as recorded; do not promote this packet to a qualification pass.