# File Operations documentation refresh

Four unedited native frames illustrate the current running card, open counting,
the fully wrapped terminal note and the selected-error Issues pane. Invented
Workshop paths are display-only fixture data. The fixture admits no operation,
opens none of those paths and uses an empty owned breadcrumb store. Each complete
frame was visually reviewed, including title bars, paths, text and issue details.
PNG bytes match the published `docs/res` assets; adjacent provenance binds their
hashes, actual pixel sizes, DPI, language/theme, scenario and build identity.

## Native validation

Command:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts -NonInteractive -Configuration Debug -Platform x64 -TestRoot D:\RedSalamander.Perf
```

- Run: `20261019T152718Z-81388-86b8516089d74e17b705f9e4bdb07900`.
- Build: test-enabled x64 Debug; zero compiler warnings/errors.
- Build receipt: `ba972175710cbc82901766e11f512848e29ee8ae81b02f531558ad9e4bdf8515`.
- Source snapshot: `60b1a81973dc6018985c9132d3bc2a497595dd4efb5f5f1f2fe5394601903b87`.
- Result: Setup, FileOps_ReviewRemediationContracts and Cleanup_RestorePluginConfig
  passed; zero failures/skips. No interactive entries were selected.
- New native predicates require the exact provisional counts, at least two note
  lines, complete note fit above progress and no overlapping action buttons.
  Existing identity, selected-issue and card-gallery checks remain enabled.

The complete raw run was preserved in a private marked-root evidence archive,
verified file-by-file through independent ZIP readback. `raw-evidence-provenance.json`
records the archive and original metric-stream digests and byte/row counts.
The broad raw stream stays outside Git; there is no performance comparison claim.

## Documentation and publication checks

The accompanying source changes expand missing product guides, correct stale
File Operations/Preferences/developer descriptions, and document sandbox evidence
and capture workflows. The coverage review accounts for 25 top-level settings
sections and 16 plugin DLL families. A shared heading-fragment checker now retains
identifier underscores and punctuation-created adjacent hyphens; focused fixture
tests cover those links and duplicate-heading suffixes.

The final focused tooling run passed 316 tests with zero failures/skips. Link and
coverage checks cover 46 Markdown documents. The specification inventory has zero
blocking findings; the human-managed scratch note was excluded before content
reads. Historical informational findings retain their existing classification.

Whole-frame review plus OCR retired 43 original `docs/res` images and three old
normative UI illustrations with unsafe identifiers, desktop content, clipped or
obsolete UI. Original image bytes and metadata were hash-verified in a private
maintenance archive before removal. Retained older illustrations are not current
build qualification. Written product coverage and the remaining capture backlog
are recorded separately in `docs/DocumentationMap.md` and `docs/res/README.md`.

## Evidence boundary

No source/index edits occurred during the build or native runner. Publishing PNGs,
provenance, the final heading-link test and prose corrections happened afterward;
the receipt identifies the native capture source, not the final documentation
workspace. This is focused presentation/tooling evidence, not Full, PrePush, CI,
Release, ARM64, ASan, foreground-input or storage-performance qualification.
