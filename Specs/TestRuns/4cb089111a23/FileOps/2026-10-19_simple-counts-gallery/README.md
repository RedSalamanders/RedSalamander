# Simple File Operations counts and expanded documentation gallery

The expanded Copy/Move count row now uses plain Folders/Files labels. Open totals
retain Counting…; final totals use plain values, and published completion receipts
use compact completed/total fractions. The existing five resource sets (English,
French, Japanese, Czech and Slovak) preserve their positional tokens and stable
resource IDs. Transfer/counting algorithms and the DxUi pin did not change.

## Native validation and captures

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts -NonInteractive -Configuration Debug -Platform x64 -TestRoot D:\RedSalamander.Perf
```

- Final run: `20261019T160349Z-75836-17f67a74d0c94f78b4c6410376addba0`.
- Test-enabled x64 Debug build: zero warnings/errors.
- Receipt: `fd8fb5c5cc804e83c1833649ea5273cdc48062ba5b4dc9437470d99bc8d0a880`.
- Source snapshot: `da6e3a4590a465cb7e6a8d4f2bb69f87865ab2cd7e4fcfd751394801fb44e165`.
- Setup, the selected native case and Cleanup_RestorePluginConfig passed, with
  zero failures/skips. No interactive entry or desktop-global input was selected.
- Literal English assertions independently check Counting… with no discovered
  counts, provisional counts, full completion fractions, totals-only Move and
  file-only completion. The existing multilingual/narrow card-gallery assertions
  continue to check layout, accessible counts, action overlap and wrapped notes.

Ten full native frames cover running Copy, open counting, totals-only Move,
file-only completion, paused/queued/completed Copy, destination conflict, the
complete terminal note and the selected Issues pane. All use private invented
Workshop identities, no admitted operation, and an empty owned breadcrumb store.
Each complete frame was visually reviewed, copied unedited and hash-bound to the
adjacent scenario/build/display/privacy provenance. Actual pixel sizes and DPI
come from the artifacts. The published `docs/res` PNGs match these exact bytes.

The first attempt (`20261019T155731Z-90092-81f4e1bada864135a31bbb6300042770`)
passed Setup/Cleanup and failed the selected case. Its new conflict fixture
incorrectly expected a status chip. The owning spec and production renderer put
that status in the decision body and intentionally omit a duplicate chip. The
fixture now requires the visible, unclipped decision prompt and absence of a
painted chip; the count and other status assertions remain intact. No incomplete
first-attempt image was published. `first-attempt/` preserves the failure, trace,
build receipt and source identity; the complete original run and fixture source
remain in a verified private archive outside Git.

## Tooling, documentation and privacy

The initial resource/source run passed 243 checks; the final publication/source
run passed 252. Their overlap is 233 checks: this is 262 unique passing checks,
including 10 localization checks and 19 documentation/capture checks. The five
resource files remained byte-identical between the localization pass and final
build; the reuse proof is retained privately. No failures/skips remain.

Links across 46 guides, all 25 top-level settings sections and all 16 plugin
families reconcile. The specification inventory has zero blocking findings;
its 43 historical informational findings are not promoted to errors or rewritten.
The human-managed scratch note was excluded before content reads. All 41 current
public image assets completed Windows OCR; new frames also received full visual
review. Image hashes/provenance supplement, and do not replace, that review.

## Frame-cost observation

Both runs use the same machine (`4cb089111a23`), Debug profile and toolchain.
The unchanged CardHierarchy timed loop contributes 1,488 samples per run. Baseline
median/p95 were 7.314/10.017 ms; candidate median/p95 were 7.359/9.544 ms. All selected
raw rows are retained in their original order, with original-stream digests and
byte counts in `card-frame-comparison.json`. The additional documentation saves
are outside this timed scope. These sequential observations do not establish an
optimization or controlled A/B/A performance qualification.

## Evidence boundary

No source/index changes occurred during either build/native runner. Image
publication, final prose and archive metadata followed the successful runner.
The production resource files, popup source and native fixture still match that
build snapshot. Complete raw runs were ZIP-readback verified in the marked test
root before scratch cleanup. This is focused x64 Debug presentation/localization
evidence, not Full, PrePush, CI, Release, ARM64, ASan, foreground-input or storage
performance qualification.
