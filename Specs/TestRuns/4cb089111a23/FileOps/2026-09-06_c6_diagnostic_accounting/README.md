# C6: trustworthy build diagnostic accounting

The final candidate passes 31/31 focused checks and 98/98 surrounding tooling
checks, with no skips. This closes only the warning-accounting repair, not C6,
native FileOps qualification, production Release staging, or final Fresh Full.

## Scope and identity

Starting commit: `0b2ef88e5d2a7e6b1392a98a0f6186c03eb1e24a`.
The six changed inputs at policy-test time are copied verbatim as `input-*`;
`copy-manifest.json` binds every promoted raw artifact's path, size and SHA-256.
Each diagnostic run also retains its exact module/test bytes, provenance, NUnit
results and all ten ordered timing rows. Other implementation dependencies remain
at the starting commit. Plan/archive/attribute edits followed testing and claim
no additional runtime result. `policy-inputs.json` records the selected policy
inputs, not a replacement for the repository's full build-receipt closure.

The historical warning witness is the unchanged
[560-line build log](../2026-09-06_c0_curl_parser_baseline_debug/build.txt), SHA-256
`597b110ea0e8e2c3299a32c5140ebbaced54d45806df8b921a0daff8f4ae1a0a`.
Its line 134 contains an unnumbered vcpkg warning about app-local dependency
gathering. The old summary incorrectly reported zero warnings. This archive
does not rewrite that old log or convert its run into zero-warning evidence.

## Correctness evidence

| Run suffix | Passed / failed / skipped | Interpretation |
|---|---|---|
| `baseline` | 16 / 14 / 0 | Old parser misses optional origins/codes/subcategories and miscounts diagnostic-looking message text. |
| `candidate` | 30 / 0 / 0 | Shared header parser fixes those cases; locale coverage not yet present. |
| `locale-baseline` | 31 / 0 / 0 | In-process culture test is insufficient: an already compiled regex cache masks the defect. Retained, not accepted as locale proof. |
| `cold-locale-baseline` | 30 / 1 / 0 | Contained fresh PowerShell process reproduces uppercase WARNING failure under Turkish culture. |
| `invariant-candidate` | 31 / 0 / 0 | Explicit invariant regex matching fixes the same cold-process failure. |
| `policy` | 98 / 0 / 0 | Diagnostic, streaming, build-output/containment, inventory and governance coverage. |

Raw run IDs are `c6-diagnostics-<suffix>-20260906` beneath the initialized
`C:\RedSalamander.Perf\runs` root. `probe-command.ps1.txt` and
`policy-command.ps1.txt` preserve the executed commands; choose new governed run
IDs for reproduction rather than overwriting these runs. The cold-process test
uses the existing contained streaming runner and does not change machine culture.

The shared classifier covers optional origin/subcategory/code, Windows/UNC paths,
node prefixes, case-insensitive locale-neutral categories and one anchored header.
It rejects ordinary filenames, incomplete headers and numeric summary lines.
Both counts and colors use it; stderr and project-completion presentation remain
intact. Exported APIs, launch selection and MSBuild exit status are unchanged.
Counts represent physical diagnostic lines, not deduplicated build events.

## Same-machine timing check

Protected path: scanning and classifying the same 560-line captured build log.
For each mode, exclude one warm sample and retain five samples of ten complete
passes. Summary mode includes file reading; color mode uses already loaded lines
and excludes terminal rendering. Values below are median/min/max milliseconds
per complete pass, derived from the preserved raw batch times divided by ten.

| Mode | Original baseline | Final invariant candidate |
|---|---|---|
| Summary | 16.84 / 15.67 / 19.06 | 20.67 / 18.28 / 24.20 |
| Replay color classification | 11.60 / 9.14 / 13.12 | 33.14 / 28.96 / 36.50 |

The correctness repair costs roughly 4 ms per log scan and 22 ms per complete
color-classification pass in this small same-machine experiment. This is not a
throughput improvement or a general performance bound. The build's live process,
terminal rendering and native FileOps execution are outside this measurement.
Preserve the intermediate candidate timing rows as well; do not cherry-pick their
faster summary measurements as final results.

## Continuation

Validation: 186/186 tooling files classified with zero blockers; spec inventory
has zero blocking findings. Archive pre-stage checks pass for 38 files, changed-set
checks for 2,196 files, and whole-inventory checks for 3,409 files. All 34 raw-copy
manifest entries, five ordered timing streams and 38 raw/index archive identities
were verified. Preserve the failed NUnit XML's original trailing stack-trace blank
through its narrow attribute exception rather than altering evidence bytes.

Commit this coherent code/test/spec/evidence step. Then validate a real full-solution
test-enabled-to-production Release transition, preserving shipping SQLite/runtime
dependencies while retiring only test-owned outputs. The preceding Clean fixtures
do not substitute for that transition. Keep the core plan in WIP until C0–C6 and
the inherited I17/final Fresh Full requirements have their own evidence.
