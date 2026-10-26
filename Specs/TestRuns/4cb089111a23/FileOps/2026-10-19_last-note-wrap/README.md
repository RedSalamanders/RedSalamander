# Last note wrapping: scoped native evidence

Protected scenario: expanded terminal File Operations cards with long localized Last note text. Captures are actual application screenshots from deterministic harness snapshots, never proposal mockups or desktop screenshots.

The baseline reproduces the owner's destination-already-exists diagnostic clipped to one line. The candidate wraps the complete paragraph and reserves its measured height before progress and completed actions. English and French at 480/960 DIP cover diagnostic text, explicit newlines, long unbroken tokens, Unicode, optional informational notes, diagnostic priority, a separate result summary, resize, collapse/reopen, changed/cleared notes, and a note taller than 4096 pixels. Scrolling reaches its final line and actions. Actual DPI: 144; light at 480 DIP, dark at 960 DIP.

## Runs and source identity

- Baseline: `20261019T083744Z-89496-4bbe7f3ad8fa41719bd0e161ee5df7d8`; expected failure (2 runner cases pass, 1 fails). Receipt `c978861df28da86479281244cea41246d2633d1d14e67a0dc0b7773431095ed2`, source snapshot `6614113d9e1f8a0a0c499e971d32d7c8e8ef104b6eb07166552cb3f5a9c7ccde`. New visible-note predicates fail in all 16 original note-bearing combinations; the four hidden-information-note combinations pass. Native captures independently show the clipping.
- Candidate: `20261019T084519Z-85416-dd3e0a6f0be34ac493880a6ad57c7b9f`; 3 runner cases pass, 0 fail/skip. Receipt `6dee06f4856bbbc38f34deb2b948eae5eb1c47b2e751c47af4c5b0ad2d931315`. All 24 note-scene combinations, four lifecycle witnesses and four scroll witnesses pass; all original card-hierarchy checks remain green. The complete raw capture set contains 128 PNGs; eight representative candidate images are retained here.
- Both x64 Debug builds report zero warnings/errors. Full MSBuild logs and exact build receipts are retained. Both use the workspace's existing DxUi pin `f65e4a4b9311370b3e3070088e2a8abe29adcb27`; this task does not stage or alter that pin change.
- `TestHarnessSourceContracts.Tests.ps1`: 204 pass, 0 fail/skip. Transcript and summary retained.
- Archive validation: explicit pre-stage passes all 36 files; whole-inventory validation passes all 15374 files.

Command for both native runs:

```powershell
.\Test-Changes.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts -Configuration Debug -TestRoot D:\RedSalamander.Perf
```

Raw runs were preserved under `D:\RedSalamander.Perf\evidence\fileops-last-note-20261019-{baseline,candidate}` before any later runner cleanup. `capture-provenance.json` binds retained image bytes, scenario manifests, DPI, no-activation method and run/dependency identities. The ordinary capture is `captures/candidate/scene24-en-US-480dip.png`; the end of the stress paragraph is visible in `captures/candidate/last-note-bottom-en-US-480dip.png`.

## Rendering observations

`performance-comparison.json` uses all matched warm samples (frames 2..11) in each language/width group, nearest-rank percentiles, with no timing exclusions. Same machine and dependency pin:

| Scene set | Frames per run | Baseline median / p95 | Candidate median / p95 |
| --- | ---: | ---: | ---: |
| Original scenes 0..23 | 960 | 8.152 / 8.542 ms | 8.158 / 8.540 ms |
| Note scenes 24..28 | 200 | 8.196 / 8.536 ms | 8.172 / 8.578 ms |

The baseline clips notes while the candidate draws their full text, so note-scene work differs. The added candidate stress scene and lifecycle steps are outside this matched comparison. These single-run observations do not prove performance equivalence, causality, or storage-throughput improvement.

Complete claim-bearing `FileOps.Popup.CardHierarchy.FrameUs` rows are retained (1392 baseline, 1440 candidate). The full raw metric streams exceed the per-file archive limit and remain beneath the marked sandbox; metric provenance records their SHA-256, bytes and row count. Other metric families are not used for a claim here.

## Qualification limits

This is focused x64 Debug, no-activation correctness and rendering evidence. It is not Full, Release, ARM64, ASan, real foreground interaction, other-DPI, or storage-performance qualification. The runner's sandbox audit remains unresolved: 272 issues at baseline and 273 at candidate. The affected explanation selects wider obligations and is not a passing repository gate. No CI result is claimed. Normative spec and user guide were present in the candidate snapshot; plan closeout and evidence promotion followed without changing product code/tests.
