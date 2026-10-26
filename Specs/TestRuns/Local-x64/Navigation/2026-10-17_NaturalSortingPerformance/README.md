# Natural sorting performance: focused x64 acceptance

2026-10-17, machine `4cb089111a23`, branch `ej/dialog-keyboard-focus`, base
`a48e7d7880cb37ea6898d462ae188a6213a5f3f8`. This is I41's bounded Debug/Release
optimization evidence. It does not qualify Fresh Full, ARM64, ASan or I33's
remaining ViewerSpace program.

## Result

Windows `StrCmpLogicalW` remains the comparator. Bulk bounded copies remove
checked iterator traversal; listing-local pane ranks avoid repeated collation
during sort toggles. Ranks include the existing ordinal/stable ties and never
replace provider identity. Removed-item probes use the uncached comparator.

Same-machine five-repeat medians, in microseconds:

| Request | Natural baseline Debug | Final Debug | Final Release |
|---|---:|---:|---:|
| Name ascending | 33435 | 28470 | 6527 |
| Extension ascending | 181383 | 87551 | 28344 |
| Time descending | 115650 | 103203 | 64860 |
| Size descending | 122424 | 102808 | 65305 |
| None ascending, first | 98464 | 97763 | 67153 |
| Name descending | 119756 | 68322 | 65773 |
| Extension descending | 151374 | 70058 | 65641 |
| Time ascending | 89933 | 83627 | 67303 |
| Size ascending | 110054 | 84757 | 65249 |
| None ascending, last | 84450 | 85226 | 66136 |

All sample values, ranges, render counts and timestamps are in
[performance-comparison.json](performance-comparison.json). The significant
Extension/Name-descending regression has returned below the earlier lexical
observations (102450/84484/75244 us respectively) in
[I40's original comparison](../2026-10-17_NaturalNameSorting/timing-comparison-final.json).
Those lexical values are single observations, not a new five-repeat lexical baseline.
Independent DxUi and RedPrism compilation was observed and recorded in
[ambient-build-observation.json](ambient-build-observation.json). These are
observed host timings with ranges; no quiet-host or tail-latency claim.

Final Debug rank preparation: median 5064 us, range 4799–8261, nine observed
5,000-row preparations. Release: median 617 us, range 566–1035, nine preparations.
The fields retain 80000 bytes per 5000 rows; preparation uses a temporary 40000-byte
index. Observed 5,000-row enumeration SortMerge medians are 41383.5 us baseline
and 41991 us final Debug (ten scopes each). Those scopes include fixture restoration
and canceled/stale work, so they do not establish cold-folder-open latency or
process-wide memory usage. The first Name request can reuse the loaded sort;
all interactive timings include warm rendering, layout and current-item repair.

## Fixture and validation

`folderView_perf_sort_toggle_stress`, existing 5000 local files: mixed case,
common prefixes, repeated extensions, extensionless files and Unicode names.
Each `--selftest-repeat=5` run executes ten Name/Extension/Time/Size/None requests
per repetition. Every repetition passes count/current/layout assertions, performs
zero inactive quick-search drawing-effect updates, and the candidate additionally
requires one enumeration throughout its ten requests.

Final Debug and Release each pass these ten Commands cases:

- `cmd_natural_name_sorting_contracts`
- `cmd_pane_find_dialog_header_click_sorts_results`
- `cmd_pane_batchRename_path_sort_uses_displayed_path_and_stable_ties`
- `cmd_pane_batchRename_manual_sort_like_preview_with_hide_unchanged`
- `cmd_pane_batchRename_folder_scope_collects_local_children_metadata`
- `cmd_navigation_sibling_natural_sorting`
- `cmd_pane_fileops_decision_mirror_real_focus`
- `folder_view_selection_refresh_filter_contract`
- `cmd_pane_delete_rehomes_current_item_to_next_or_previous`
- `cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs`

The natural-contract case checks direct Windows results for bounded/long/empty
views, autocomplete limits, and the actual pane comparator. The new rank matrix
checks six keys in both directions, natural/ordinal/exact ties, directories first,
numeric primary metadata and mixed cached/uncached resolver probes.

Each profile also passes three fresh native `--no-activate` ViewerSpace cases:
`TestViewerSpaceAnalysisHostVirtualGridAndLifetime`,
`TestViewerSpaceAnalysisAllowsAmbientDesktopChanges`, and
`TestViewerSpaceListingValidation`. These retain natural Name/breadcrumb assertions,
metadata ties, lifetime/ambient-desktop behavior and distinct provider identities.
All six native executions return zero. Source/tooling contracts: 310 passed,
zero failed/skipped. Changed-range formatting passes on 22 C++ files.

## Provenance and retained attempts

Final full-solution build receipts (both zero warnings/errors):

- Debug: `a8e229566928911a801b9f32d3037b0b46517040b4f5989e1ab4f4b942148221`.
- Release: `820a9ff82b273393a8aa7d10b5f47dee423f03e0ec656e5cef682ffe821add37`.

`debug-qualified`, `functional-debug-qualified` and all Debug viewer cases use
that same Debug receipt; `candidate-release`, `functional-release` and the
Release viewer cases use the same Release receipt. No rebuild occurs between
these final performance and correctness checks. [runtime-verification.json](runtime-verification.json)
verifies all 384 current receipt-bound outputs and all 23 changed C++/project
inputs, plus each viewer executable/plugin identity.

`baseline-debug` is the original Windows-natural implementation with checked
iterator copies and no ranks. `candidate-debug`, `debug-rebuilt` and
`functional-debug` retain successful intermediate evidence. Rebuilds changed
binary hashes, so final acceptance was repeated on the fixed full-solution
receipts; no intermediate pass replaces the final records. The two original
pre-execution blocks are retained in `functional-debug-runner.log` (target-only
receipt refused by SkipBuild) and `functional-debug-retry-runner.log` (active
Release shared-dependency build). Neither executed native cases.

All 250 sort-toggle events from all five measurement phases are retained.
`perf-selected.jsonl` files contain every original complete line for each named
metric family, including every ApplyCurrentSort, BuildItems, SortMerge and rank
preparation observation. No outlier is dropped. The per-case ten-request JSON is
the final repetition because the harness overwrites that filename; complete JSONL
retains all repetitions. Full original bulk metrics remain under
`Z:\RedSalamander.Perf\evidence\natural-sort-optimization-2026-10-17`.
[artifact-bindings.json](artifact-bindings.json) binds originals, selected-line
derivations and their source/output hashes. `external-phase-bindings.json` files
describe the external originals, including bulk files intentionally absent here.

`probe-debug.csv` is a seven-round, 500000-comparison diagnostic of checked copy,
bulk copy and direct terminated Windows comparison. It explains the copy overhead;
it is not an end-to-end or WIL/allocation benchmark. Existing sandbox audit
advisories remain in the unmodified run results (41 legacy entries, plus one
owned tooling fixture in an intermediate run). No independently owned application
or workspace was terminated or cleaned. The affected-plan explanation retains
broader pending obligations; this packet is focused acceptance only.
