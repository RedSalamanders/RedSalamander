# File Operations UI review fixes — 19 October 2026

The four review findings are fixed and the final focused checks pass. Foreground validation also
found a stale focused-source tooltip when another stream finished; that defect is fixed and its
original failing assertion now passes. [Final review](FinalReview.md) records the conclusions.

Final x64 Debug solution build: zero warnings/errors, receipt
`ceb646bc0fd1e47dc2c9cf61cf1ea4110c3e4281814923330281745ef84b866e`, source snapshot
`e1790f1fba69722dda44433a7abe34997137935a2808f1a8f9994851dc7c78b4`.
The existing unstaged DxUi pin `f65e4a4b9311370b3e3070088e2a8abe29adcb27` was preserved.

| Final run | Result |
|---|---|
| Commands throughput disclosure, external reload and restart | 3 passed; actual foreground input authorized by the user after the build, governed warning/lease |
| FileOps_ReviewRemediationContracts | 3 runner cases passed; 124 card-fact checks, including 48 exact percentage requirements; 132 native captures |
| DiscoveryScope_FolderLiveProgress | 3 runner cases passed; 69 live-operation assertions, including five independent count-oracle rejection checks |
| TestHarnessSourceContracts.Tests.ps1 | 204 passed |

Run IDs, filters and limits are in [qualification-scope.json](qualification-scope.json).
Original results/traces, restart-child results, final build receipt/log and eight source-byte
checks accompany this document. The product, tests, normative spec, guide and dependency bytes
stayed fixed through the final build and all final runtime runs. Only plan closeout, curated
evidence and staging followed. All screenshot pixels came from deterministic native test harnesses;
the card matrix uses English/French, light/dark, 480/960 DIP at actual 144 DPI. See
[capture provenance](capture-provenance.json) and the scenario manifests.

Matched warm card rendering: 1,160 samples each, baseline median/p95 **8.159/8.550 ms**, candidate
**8.109/8.575 ms**. Changed Delete scene 12 and new unknown-limit scene 30 are excluded from the
matched figures. Matched path-access batches (20 batches, eight redraws each): baseline p95
**66.852 ms**, candidate **66.746 ms**. These are single observations, not performance equivalence
or causal proof, and make no storage-throughput claim. Complete selected metric families and
raw-stream hashes/paths are retained; see [comparison](performance-comparison.json).

Failed evidence remains failed: the first build caught 64-bit completion counts narrowed by the
new test snapshot; the fields were corrected. Commands attempts then exposed obsolete French
copy wording, an Escape/refocus fixture error, idle-popup setup assumptions and the real stale
other-stream tooltip. Their original results, traces, build identities and correction diffs are
retained here. Full raw runs remain beneath
`D:\RedSalamander.Perf\evidence\fileops-ui-review-fixes-20261019-*`.
The initial combined no-activation filter was not executed because the runner requires exact
allowlisted filters; both cases were subsequently run separately. It contributes no passing checks.

Qualification is focused x64 Debug only. Full, Release, ARM64, ASan, other DPI scales and external
assistive technology were not evaluated. The shared sandbox audit still reports 274 findings;
these selected case passes do not establish clean whole-repository qualification. Existing archive
inventory validation passed for 15,374 files before adding this packet; this packet is separately
validated in explicit pre-stage mode.
