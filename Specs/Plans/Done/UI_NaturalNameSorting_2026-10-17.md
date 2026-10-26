# I40 Natural file and folder name sorting

Status: DONE. Focused x64 Debug/Release acceptance completed 2026-10-17.

Use Windows natural display ordering throughout main panes, Find Name/Path,
address autocomplete and Batch Rename source collection/Original/New/Path.
Keep numeric metadata sorting, explicit source order, stable grid ties and
provider identity/collision/dependency rules under their existing contracts.
Breadcrumb menus already use natural ordering (I39 and the I33 follow-up). The wider
audit also covers ViewerSpace's analysis Name column and equal-metadata display ties;
its ordinal provider duplicate-name validator remains separate.

- [x] Add shared bounded-view-safe comparison and production-path regressions.
- [x] Retain lexical baseline failures and pane sort-toggle timing evidence.
- [x] Apply natural ordering across the four surfaces and review remaining name sorts.
- [x] Update owning UI contracts, shared-helper catalog and usage documentation.
- [x] Build and run affected native cases in x64 Debug and Release; compare the
  same-machine pane fixture and retain raw receipts and timing samples.
- [x] Validate formatting, source/spec/tool contracts, archive evidence and close I40.

Protected performance scenario: the existing 5000-entry
`folderView_perf_sort_toggle_stress` fixture and its `folder.sort_toggle_us`
measurements, baseline and candidate on this host. Ordinary name comparisons
must avoid per-comparison heap allocation; length-delimited provider names and
relative path slices must never be read beyond their declared length.
Long slices use an exact temporary copy; no truncation or locale-dependent
fallback is permitted. Focused checks do not establish Full/ASan/ARM64 coverage.

The owning contracts are [FolderView](../../UI/UI_FolderView.md),
[Find](../../UI/UI_FindFilesWindow.md), [NavigationView](../../UI/UI_NavigationView.md),
[Batch Rename](../../UI/UI_BatchRenameWindow.md),
[ViewerSpace](../../Plugins/Plugins_ViewerSpace.md) and the
[shared-helper catalog](../../Core/Core_SharedHelpers.md).

Final acceptance: the app/runtime graph and viewer harness build without warnings
or errors in both x64 configurations. Each configuration passes all eight selected
Commands cases and all three viewer cases, with no native case failures/skips.
The real-input dialog focus regression is included. All 309 final source/docs/tool
contracts pass, as do changed-range formatting, whitespace and the explicit
explicit evidence archive contract. The exact I40 Done path has a positive and
adjacent-date negative admission test; broad plan owners remain unchanged.

The [evidence packet](../../TestRuns/Local-x64/Navigation/2026-10-17_NaturalNameSorting/README.md)
retains original lexical failures, the viewer's three Name/tie failures, initial
compile and shadow-warning diagnostics, passing results, raw named timing samples,
workspace/build identities and source/artifact SHA256 bindings. Full bulk metrics
remain in the marked test root, with exact selected-line derivations archived.

Natural comparison has a measured Debug cost on the 5000-item fixture. Name
ascending is 34.334 ms (lexical baseline 29.118); extension ascending is 162.400 ms
(baseline 102.450). An identical-key shortcut reduced observed extension cost from
its first natural candidate; all samples and caveats are retained. Release's ten
sort samples are 22.691-67.053 ms. There is no paired Release baseline or isolated
comparator/p95 claim. Further optimization can prove provider-arena terminators to
avoid copies, with separate bounds/lifetime tests, while preserving natural order.
No threshold was rebaselined. Full/ASan/ARM64 and the existing I29/I33 qualification
work remain outside this bounded closeout.
The final automatic explanation covers all 51 dirty worktree paths and recommends
broader affected work, including existing work owned by other plans. It exits 2
(NOT_EVALUATED), with five broad interactive entries pending; no repository-wide
pass is claimed. This request's explicit eight Commands and three viewer witnesses
were selected, built and executed in each configuration. `scope-explanation.log`
retains the broader recommendation and its unclosed obligations.