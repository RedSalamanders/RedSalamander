# I41 Natural sorting performance

Status: COMPLETED (bounded x64 Debug/Release optimization, 2026-10-17). Planned at `a48e7d7880cb37ea6898d462ae188a6213a5f3f8`.
Drift check: `git diff HEAD -- Common/NaturalSort.h RedSalamander/FolderView.Enumeration.cpp RedSalamander/FolderView.h`.

Owner: this bounded successor to completed I40. Preserve Windows `StrCmpLogicalW`
ordering, all tie policies, provider identity and the I38 dialog-focus fix. I33's
broader ViewerSpace and platform qualification remain separate.

- [x] Retain the current Debug five-repeat 5,000-item sort-toggle baseline.
- [x] Remove checked iterator scratch-copy overhead and reuse listing-local natural
  name/extension ranks for pane sorting, with uncached focus-probe fallback.
- [x] Add tests for cached and uncached comparisons, ties, directions, metadata, bounded
  slices and all existing natural-sorting surfaces.
- [x] Build/test Debug and Release; compare the identical repeated fixture and
  record initial enumeration/preparation cost as well as interactive toggles.
- [x] Update authoritative contracts, archive evidence and move this plan to Done.

Authority: `Specs/Core/Core_SharedHelpers.md`, `Specs/UI/UI_FolderView.md`,
`Specs/Testing/Testing_PerformanceValidation.md`; consumers remain governed by
their existing natural ordering contracts. Implementation: `Common/NaturalSort.h`,
`FolderView.h/.Enumeration.cpp`; tests: Commands natural contracts, sort-toggle,
Find, Batch Rename, breadcrumbs and focus plus ViewerSpace's native cases.

Performance fixture: existing 5,000 local files, ten Name/Extension/Time/Size/None
requests, five repetitions per profile on the same machine. Retain all original
samples and build receipts beneath a fresh TestRuns packet. The scratch-copy
diagnostic is explanatory only, not end-to-end performance qualification.

Stop on Windows ordering differences, stale ranks crossing listing generations,
provider identity changes, an independently launched output-locking application,
or a measured unresolved Debug regression. No Full/platform claim from focused runs.

## Delivered result

`NaturalSort::Compare` still calls Windows `StrCmpLogicalW`; bounded scratch copies
use `memcpy` without checked iterator traversal. Pane ranks reuse the existing
Windows name presort and group extension keys once on the enumeration worker.
Both directions, all keys, logical/ordinal/stable ties and mixed cached/uncached
resolver comparisons pass the native matrix. The stress case also requires one
enumeration throughout all ten interactive requests.

Same-machine five-repeat medians (microseconds):

| Request | Natural baseline Debug | Final Debug | Final Release |
|---|---:|---:|---:|
| Extension ascending | 181383 | 87551 | 28344 |
| Extension descending | 151374 | 70058 | 65641 |
| Name descending | 119756 | 68322 | 65773 |

The large Debug regression is removed relative to the earlier I40 lexical
observations (102450/84484/75244 us respectively); those older values are single
observations, not a fresh repeated lexical baseline. All 250 sort samples across
the baseline, intermediate and final profiles remain retained. Independent
DxUi/RedPrism compilation was observed; these are measured host latencies, not
quiet-host or tail-latency qualification.

Observed 5,000-row SortMerge median: 41383.5 us baseline, 41991 us final Debug.
These scopes include fixture restoration and canceled/stale enumeration work.
Final Debug rank preparation median is 5064 us (4799–8261); Release is 617 us
(566–1035). Retained rank fields cost 80000 bytes at 5000 rows, plus a temporary
40000-byte index. No process-wide memory or cold-open latency claim.

Final full-solution build receipts:

- Debug: `a8e229566928911a801b9f32d3037b0b46517040b4f5989e1ab4f4b942148221`.
- Release: `820a9ff82b273393a8aa7d10b5f47dee423f03e0ec656e5cef682ffe821add37`.

Both builds have zero warnings/errors. Each profile passes ten Commands correctness
cases, five stress repetitions and three ViewerSpace native cases. All six native
viewer executions are nonactivating. Final hash verification covers 384 runtime
outputs and 23 changed C++/project inputs. Source/tool contracts: 310 passed.
Two pre-execution attempts were blocked by receipt scope and shared-dependency
coordination; their original diagnostics are retained. The final tests use fixed
full-solution receipts without intervening rebuilds.

Evidence: [focused performance packet](../../TestRuns/Local-x64/Navigation/2026-10-17_NaturalSortingPerformance/README.md).
The automatic affected plan remains broader than these named cases; I29/I33,
Fresh Full, ARM64 and ASan qualification are not closed here. Existing sandbox
audit advisories remain unchanged; independently owned applications/workspaces
were not terminated or cleaned.
