# I39 Natural sorting in breadcrumb sibling menus

Status: DONE — implementation and focused x64 Debug/Release qualification complete.
Owner: this change, requested 2026-10-17.

Use Windows natural name comparison for sibling folders in every NavigationView, including the main
panes, Find/Batch Rename embeddings, destination bars and full-path popup breadcrumbs. The shared
TryGetSiblingFolders query owns ordering and the bounded earliest-99 selection, retaining the current
folder and the existing overflow action. Canonical paths and provider identity comparisons are unchanged.
The owning contract is [NavigationView](../../UI/UI_NavigationView.md#separator-click-behavior-showsiblingsdropdown).

- [x] Trace both sibling popup builders to the same query; search shared helpers before adding utilities.
- [x] Add a regression through the real cached query: mixed case, numbers, multiple digit runs, Unicode,
  dot names, leading zeros, files excluded, empty/exact-limit/oversized sets and current-folder retention.
- [x] Retain a baseline native run and 32 cached-query timing samples for 12 and 1000 siblings.
- [x] Apply StrCmpLogicalW with deterministic ordinal ties to both selection and sorting; use a bounded
  heap instead of repeated largest-item scans, with current-folder priority only during retention.
- [x] Run the focused Debug case and compare timings on the same machine/fixture; archive original results.
- [x] Review specs, user docs, formatting and source contracts; record all remaining configuration gates.
- [x] Build x64 Release and run the natural-sorting regression after the user closed the existing instance.
  Keep Full/native platform gates separate.

The user initially deferred Release while the application was running, then closed it and explicitly
authorized the Release build and both fixes' runtime checks on 2026-10-17. The governed Commands runner
owns the existing warning and desktop lease; this sorting case never opens a popup or sends input.
Keep focused Debug/Release evidence separate from Full and native platform qualification.

Baseline: Debug build `a4b1a6740e6dffb90da43da559cafc405a127ef31debf4facad22b6ac95a4543`,
run `20261017T160547Z-85008-039b803c38b84950ae38676d58c8a639`: the new regression fails the old
lexical numeric order as intended (0 passed, 1 failed, 0 skipped). Original results, trace, workspace
and hashes are retained under `Z:\RedSalamander.Perf\evidence\breadcrumb-natural-sorting-2026-10-17\baseline`.
All 64 cached-query samples are retained; no performance verdict is inferred before pairing.
The initial test build failed on the Windows `small` macro; the fixture variable was renamed before
this valid baseline. Its original diagnostic remains in `.build/logs/msbuild-20261017_180134_660-pid101620-25538142.log`.

Candidate: Debug build `45ba644e093f9c38bb8a9fa2d721895e2f9e2b5b63606d2a62b28a9096ab7a10`,
run `20261017T161239Z-96908-1c8cf00273f545189e4cb8e5623dcdb9`: 1 passed, 0 failed/skipped;
build has zero warnings/errors. All 222 source/documentation contracts and changed-range formatting pass.
The warm-query p95 improves from 0.3983 to 0.2777 ms (12 siblings), and 10.3279 to 8.1884 ms (1000).
These are paired Debug query observations for the combined comparator/heap change, not a claim about
Release, menu rendering, individual optimization contributions, or whole-product performance.

Reviewed original results, raw timing traces, all samples, source/runtime bindings, failed baseline and
initial compile diagnostic are in the [bounded evidence packet](../../TestRuns/Local-x64/Navigation/2026-10-17_NaturalSiblingSorting/README.md).

Release closeout: the test-enabled x64 Release build completed with zero warnings/errors. The exact
`cmd_navigation_sibling_natural_sorting` case and dialog real-input case both passed (2 passed,
0 failed/skipped), run `20261017T163444Z-107504-ef07845e7c224789a110c4039566a6fa`, receipt
`479a733e2a272305e7b9ef5a166e51a37c90f9b9b14599f8eaebd9d1ed14c2a1`.
Original results, all 64 Release query samples, raw traces, build logs, receipts, source/runtime bindings
and tested C++ hashes are in the
[Release evidence packet](../../TestRuns/Local-x64/Navigation/2026-10-17_ReleaseDialogAndNaturalSorting/README.md).
The six changed C++ files match their passing Debug identities; subsequent changes only close plans,
admit their exact Done paths and archive evidence. The Release run establishes functional acceptance,
not a paired Release performance result or Full/ASan/native ARM64 qualification.
