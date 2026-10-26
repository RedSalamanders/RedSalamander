# I43 Debug sorting row movement

Status: DONE, 2026-10-20. Scope: Debug interactive row ordering, its correctness
contracts and fresh paired x64 Debug/Release measurements. This is not a Full or
platform qualification or a claim of uniform Release rendered-latency improvement.
Baseline source: `d7dd952d629985b48d25b23b56d1726fbb2a1ec5`.
Implementation branch: `ej/debug-sort-row-moves`.

- [x] Separate row ordering from selection/current repair and rendered-settling latency.
- [x] Retain fresh five-repeat 5,000-file Debug and Release baselines with no observed background builds.
- [x] Reduce measured ordering overhead without changing Windows natural comparison, numeric metadata keys, directories first or uncached resolver probes.
- [x] Cover permutation cycles, payload/ownership continuity, cached/uncached ordering, both directions, numeric extremes and current/selection restoration.
- [x] Update the authoritative FolderView and performance contracts and usage documentation.
- [x] Build and test x64 Debug/Release; retain every benchmark sample, rejected attempt and original failure; compare paired core and total timings.
- [x] Verify formatting, source/tool/spec contracts, archive, raw evidence and compiled source/output identities; archive this completed record.

Whole `FolderItem` introsort repeatedly moved strings and owned rendering resources
in checked Debug code. Interactive ordering now sorts a bounded, reusable index
permutation with the unchanged comparator, then applies whole-row cycles. Fixed
rows incur no payload moves. Other rows need at most `n + floor(n / 2)` moves.
One staging row is created lazily and reused across cycles: installed MSVC checked
string move construction allocates proxies, whereas equal-allocator move assignment
reuses them. Checked iterator safety remains enabled globally. Initial worker
enumeration sorting, the Windows comparison helper/rank preparation, build flags,
sort threshold and existing benchmark fixture are unchanged.

The [evidence packet](../../TestRuns/Local-x64/Navigation/2026-10-20_DebugSortingRowMoves/README.md)
retains all 50 request timings per phase, every relevant sorting metric, all
measurement snapshots, exact receipts and rejected/superseded attempts. Final
baseline/candidate pairs use the same isolated checkout at
`Z:/src/RedSalamander-worktrees/debug-sort-row-moves-validation`. Concurrent
ViewerSpace edits invalidated the main V2 build; those unrelated edits were
preserved and excluded from qualification. Both qualified profiles have zero
observed build-overlap polls and maximum snapshot gaps below 1,000 ms. This is
snapshot observation, not a lossless process-start audit or proof of an idle host.
The user-authorized build pause was released after the final measurement.

For 5,000-row core ordering records, including setup/restoration, Debug median
OrderItems falls from 33.385 ms to 5.899 ms (82.33%); complete ApplyCurrentSort falls
from 33.758 ms to 6.305 ms. Every one of the ten Debug request medians decreases.
Release median OrderItems falls from 0.971 ms to 0.529 ms (45.52%); complete
ApplyCurrentSort falls from 1.008 ms to 0.559 ms. Candidate counters retain exactly
40,000 index bytes and use 0-7,500 payload moves at 5,000 rows.

Release total-request medians improve for six requests and increase for four.
Extension Ascending increases from 28.266 ms to 49.861 ms (+76.399%); five-sample
ranges overlap (25.134-69.504 ms versus 47.696-79.084 ms). Size Descending, Time
Ascending and the final None request also increase. No samples were removed and no
threshold was weakened. These observations do not identify a cause or qualify
Release rendered-latency non-regression. A separate counterbalanced investigation
of rendering/settling is required before making that broader claim; I2 retains
the broad FolderView performance closeout. This bounded Debug repair does not
silently close that work.

Both final isolated solution builds are test-enabled, with zero warnings/errors.
All six focused native cases pass in each profile: natural sorting contracts,
bounded current memory, navigation roundtrip, refresh/filter selection,
visibility/selection restoration and focus/selection/removal state. The real
reorder helper is exercised for all six keys, both directions, three cache modes,
four input orders, all 720 six-row permutations, owned DirectWrite layouts and
empty/inline/heap strings, empty/singleton models, scratch reuse/release, the actual
sequential/parallel threshold and full-width numeric extremes. All 289 focused
tooling/source-contract tests pass. Specification/tool inventories, changed-range
formatting, explicit archive validation and independent raw-data verification pass.

Final artifact-lock readback matches the compiled C++ inputs and current runtime
closure in both isolated profiles. The main checkout's three sorting C++ files
match those qualified bytes. Later plan/evidence closeout is documentation only.
The running main Release application is undisturbed; qualified binaries are under
the isolated checkout's `.build/x64/Debug` and `.build/x64/Release` directories.
All four selected benchmarks and both final native runs retain exactly the same
41 pre-existing sandbox audit category/path/reason records; those issues remain
unresolved. Passing native cases do not imply a clean sandbox audit.

Name/Extension and main-pane metadata name ties retain Windows natural ordering.
Time uses signed 64-bit modification timestamps; file Size uses unsigned 64-bit
byte counts; folder Size ordering uses natural names; Attributes uses numeric
DWORD flags. Find Size/Modified remain numeric with arrival-order ties. Displayed
date/size formatting never supplies a primary numeric sort key.

Authority: `Specs/UI/UI_FolderView.md`,
`Specs/Testing/Testing_PerformanceValidation.md`, `Specs/Core/Core_SharedHelpers.md`.
I41 and the October 17 measurements remain immutable historical evidence. I2's
broader performance and Full/platform gates remain separate. No DxUi dependency
or consumer-migration change is included.
