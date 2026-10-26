# Natural file and folder name sorting

Focused x64 acceptance, 2026-10-17. This packet closes the bounded I40 request;
it does not establish Full, ASan, ARM64, broad provider/locale or I33 qualification.

`Common/NaturalSort.h` owns Windows natural display comparison. Main pane Name,
Extension and metadata-name ties, Find Name/Path, autocomplete before trimming,
Batch Rename collection/Original/New/Path/manual ordering, and ViewerSpace
breadcrumb/analysis names use it. Numeric metadata and provider identity,
duplicate validation, collision and execution ordering retain their contracts.
Stable grids keep source/arrival ties; bounded choices use ordinal spelling ties.
Length-delimited names are copied safely without assuming a trailing NUL; ordinary
names use bounded stack storage, long slices are copied in full, and identical keys
avoid copies and collation calls. Owned strings compare directly.

| Final witness | Result |
| --- | --- |
| Debug app and runtime graph build | 0 warnings, 0 errors |
| Debug Commands | 8 passed, 0 failed, 0 skipped |
| Debug viewer harness build | 0 warnings, 0 errors |
| Debug viewer diagnostics | 3 passed |
| Release app and runtime graph build | 0 warnings, 0 errors |
| Release Commands | 8 passed, 0 failed, 0 skipped |
| Release viewer harness build | 0 warnings, 0 errors |
| Release viewer diagnostics | 3 passed |
| Final source/documentation/tooling contracts | 309 passed, 0 failed/skipped/pending |
| Changed-range clang-format and whitespace | Passed |

The eight Commands cases cover pane comparator keys/directions, directory priority,
numeric metadata, unsorted order, Unicode/case/leading zeros, deliberately bounded
views and stack/heap boundaries; autocomplete order/overflow; real Find header cycles
and both Path directions with selection; Batch Rename source counters, displayed-Path
stable ties, Original/New names and hidden-row manual mappings; breadcrumb limits;
and the real dialog keyboard-focus regression. The existing 5000-item pane fixture
also verifies stable counts, retained focus and no inactive incremental-search work.

The three viewer cases are `TestViewerSpaceAnalysisHostVirtualGridAndLifetime`,
`TestViewerSpaceAnalysisAllowsAmbientDesktopChanges` and
`TestViewerSpaceListingValidation`. They check actual retained-model sorting in both
Name directions, equal-size name ties, mixed numeric/case/dot/Unicode names, no
provider I/O, stale generations, UIA/lifetime and provider identity. Numeric names
`file2`, `file02` and `file10` remain distinct under both provider case policies.
The two UI cases run with `--no-activate`; their desktop witnesses report no owned
foreground transitions. The existing 2002-row root fixture remains intact.

`summary.json` records exact run IDs and receipts. Final app receipts:
Debug `c503b5a3c90abf30be9047bf79fc6e8a01ab38c908e8cca0f86b1a57ad450dba`;
Release `5bc4d5b91beaa4c5834ee131c4c1a2b87bcd287aa028a87786d8d87c0caed637`.
Viewer harness receipts and exact test/plugin hashes are retained per phase.
`tested-code-final.json` binds the current C++/project inputs. The Debug app checks
preceded a test-only local-variable rename clearing C4456; the viewer test project
was rebuilt and tested afterwards. Application code and runtime dependencies did
not change in that warning repair. Release was built and tested on the final source.

## Baselines and corrections

The app lexical baseline has 1 passed and 4 failed cases. Its observed Find order is
`folder1, folder10, folder2`; pane/autocomplete and Batch Rename assertions reject
the old policy. The independent viewer baseline fails exactly three new ordering
checks: both Name directions and equal-size directory ties. Its plugin hash matches
the retained `debug-fast` app receipt. Its navigation, identity, no-I/O and UIA checks
pass. Original failures are retained, never replaced by passing retries.

The initial test build failed on an invalid include path (C1083), corrected before
the valid lexical baseline. A later test-local shadow-name warning (C4456) was
cleared before final viewer builds. Both diagnostic logs are retained. The manual
mapping fixture's proposed names changed from `renamedN` to `z-renamedN` after its
baseline, preserving the pre-existing unchanged-row position while still testing
numeric order. This fixture adjustment does not affect the pane timing corpus.

## Performance observations

Every sort row is retained in `timing-comparison-final.json`, the original fixture
JSON and all ten `folder.sort_toggle_us` events per app phase. These are individual
end-to-end observations, including comparison, focus/scroll repair and rendering;
they are not percentiles, isolated comparator costs or a non-regression claim.

| 5000 items, milliseconds | Lexical Debug | Final natural Debug | Natural Release |
| --- | ---: | ---: | ---: |
| Name ascending | 29.118 | 34.334 | 22.691 |
| Name descending | 75.244 | 113.817 | 66.190 |
| Extension ascending | 102.450 | 162.400 | 50.937 |
| Extension descending | 84.484 | 157.236 | 65.243 |

Natural ordering has a measured Debug cost. The initial natural extension samples
were 200.616/182.470 ms; the identical-key shortcut produced 174.585/155.214 ms in
the next run. Final samples above retain the remaining cost. These sequential
observations do not isolate the shortcut's contribution. The requested natural
policy is retained. A future measured optimization could prove terminators within
the provider arena and bypass temporary copies on that proven path; it would need
separate arena-bound and lifetime tests. Any further optimization must preserve
the requested natural order and provider identity. No budget was silently rebaselined.
Release's ten sort samples range from 22.691 to 67.053 ms; there is no paired Release
baseline, so no Release speedup or non-regression is inferred. Source-contract runs
finished before final native pane timing. Viewer traces preserve all query samples;
they do not establish scan or popup-presentation performance qualification.

## Provenance and limits

Original phase results, traces, workspace records, build logs and receipts are
copied with SHA256 bindings in `artifact-bindings.json`. `external-phase-bindings.json`
files describe the full external phases, including bulk metrics. Full raw metrics
remain under `Z:\RedSalamander.Perf\evidence\natural-name-sorting-2026-10-17`.
To keep this packet bounded, `perf-selected.jsonl` contains original complete lines
for the documented metric selection; its source/output hashes and selection are
recorded separately. All named sort samples, including first and maximum, remain.

Commands used the governed interactive warning/desktop lease and restored focus.
Viewer diagnostics used the contained-process launcher, artifact profile lock,
isolated profile/TEMP and no activation. No desktop screenshots were taken.
Disk audits retain the existing 41 unowned legacy sibling run directories; final
Commands audits also see the completed concurrent Pester run as one stale sibling.
These advisory audit findings remain in original results. No unowned data was
deleted, and these focused passes do not claim a clean Full qualification.

All final application/test/plugin inputs are checked against passing hashes.
Subsequent edits close documentation, the exact I40 Done path and this archive;
the broader I29/I33 and platform qualification records remain separate.

The final automatic affected explanation sees all 51 dirty worktree paths, recommends
broader work and exits 2 (NOT_EVALUATED), with five broad interactive entries pending.
It was explanation-only and ran after the passing native cases. The named focused
cases above were explicitly selected and executed; they do not satisfy those broader
repository obligations. `scope-explanation.log` preserves the recommendation.
`final-artifact-verification.json` confirms 270 runtime artifacts and 22 C++/project
inputs still match passing evidence. The specification inventory has zero blocking
findings, and the tool inventory succeeds. Final packet validation is explicit and
covers ignored, pre-stage files.