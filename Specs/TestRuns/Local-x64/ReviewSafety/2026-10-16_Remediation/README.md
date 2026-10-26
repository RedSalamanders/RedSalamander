# Supplied-review remediation validation

Status: focused and final Fresh noninteractive PR checks passed. The repository verdict remains NOT_EVALUATED (partial scope), with the broader qualification listed below pending.
This is bounded remediation evidence, not Fresh Full repository or release qualification.
Branch: `ej/review-safety-fixes`; base: `cca51ab0935a7d2b8f23eae7c4498fadca523904`.
The user requires nonactivating validation. No interactive acceptance is claimed.

## Delivered scope

Review items 1, 2, 6, 7, 8 and 10: one checked native-directory decoder and migrated consumers;
retained no-follow delete-parent authority; execution-closure PR digest; independently verified
benchmark output on both routes; typed XP-09 prerequisite status; and random, destination/message/
type-bound payload adoption. The payload type identity uses writable per-type data because identical
code folding can merge deleter functions. All receivers were migrated.

Real benchmark verification also exposed Dummy's DWORD-only record strides. Dummy now aligns to
`alignof(FileInfo)` and reuses the checked indexed locator; PluginContractTests covers four filename
lengths and streamed/indexed parity. The existing search source guard now requires delegation to the
shared decoder, with an executable adapter case whose name overlaps a valid following record.

Owning Core, FileSystem, VFS and Testing specs, helper guidance, tooling/native inventories and the
bounded I36 checklist were updated. The larger I36 plan remains active for its separate residual work.
Items 3, 5 and 9 were not substantiated as originally claimed against this base and are outside this patch.

## Exact current build and focused runs

Full test-enabled x64 Debug build: **0 warnings, 0 errors**. Compiler debug information is `Embedded`
(`/Z7`, normal final linker symbols retained), the documented mode for coexistence with disjoint builds.
Receipt: `5dbf4bdcd2fc394639357578e1572a910c5a6edefa1ff2ffb2532359b29e8543`.
Source snapshot: `dafd34c62c531f9661bdc3743f208f784b187bc51a378c2ac0c416c52fa1acc5`.
The receipt identity and original compiler-log digest are alongside this file.

| Archive | Result | Canonical run id |
|---|---|---|
| `search-current` | 1 passed, 0 failed/skipped; actual search-adapter malformed-record and source guards | `20261016T103105Z-77916-e1acef352609460ebc467692b0393402` |
| `parent-current-repeat3` | 9 phase cases passed, 0 failed/skipped; 24 race witnesses, 8 variants x 3 repeats | `20261016T103409Z-76152-59bdddf1fb4d445e961792cca161297e` |
| `smallfile-current` | 3 phase cases passed, 0 failed/skipped; 2,000 exact names/contents on each route, missing/corrupt controls on both | `20261016T103518Z-76152-b7c233bfdcc24dcf979dc08907907dc1` |
| `xp09-current` | 3 phase cases passed, 0 failed/skipped; real reachable loopback-share provider route | `20261016T103632Z-76152-08104267b20243efad21baafb6163e4f` |

The delete matrix covers bound/unbound, configured concurrency 1/4, success/cancellation and 32 nested
levels (observed traversal depth 34). Bound deletion reports its effective concurrency separately.
During the real child-delete pause, ancestor rename and `GENERIC_WRITE` are denied, the outside
sentinel remains intact, and surviving-ancestor rename proves handle release after cancellation.
The same case checks the actual XP-09 helper's injected Unreachable=skipped and Refused=failed results.

Commands use `Tools/Run-AllTests.ps1 -SkipBuild -NonInteractive -ValidationMode Fresh -TestRoot
Z:\RedSalamander.Perf -FailFast`, with the named FileOps/Compare `-CaseFilter`; the parent case adds
`-SelfTestRepeat 3`. The small-file environment selects count 2000, repeat 1 and tag
`review-safety-current`. The environment selects `RSBuildCompilerDebugInformation=Embedded`.
No wide/wildcard FileOps filter was admitted to the nonactivating exception.

Observed copy intervals: native 1,609,315 us (configured concurrency 8), bridge 1,465,211 us
(concurrency 4). Output verification and negative controls are outside those intervals.
These are candidate observations; there is no matched performance baseline or speedup claim.
Injected delete-witness timings are not ordinary delete-throughput measurements.

## Wider checks and preserved failures

The final PR run is `20261016T103814Z-85692-45b948072c224588a20bf36aecdf7202`, archived in `pr-current`: all 13 entries passed, with 950 aggregate checks passed, 0 failed and 0 skipped; Pester contributed 938 passed, 0 failed/skipped. All 13 entries were freshly executed and promoted. The change verdict is PASSED; the repository verdict is NOT_EVALUATED and exit 2 correctly denotes partial coverage.
It requests `-Suite PR -SkipBuild -NonInteractive -ValidationMode Fresh`, with the same compiler mode.

`header-probes-final` retains native ASan parser assertions and an optimized `/GL /LTCG /OPT:ICF`
payload probe (10,000 destination/message/trivial-type/replay/drain iterations). All three exact
header/test input hashes still match the current source. These are focused probes, not ASan/Release
application qualification. `tooling-final` retains 352 selected Pester contracts from development.
The current specification inventory has 0 blocking findings (71 nonblocking notes), 62 normative
files and 29 WIP plans. The recorded workflow execution digest matches its recomputed closure.
`scoped-plan-explain.json` is explanation only; its five focus-taking obligations remain pending.
An initial console-promotion attempt used separate scratch which the governed cleanup swept;
the canonical explanation decision itself was retained instead.

`pr-debug-first` is **FAILED**, not superseded into a green historical result: 948 passed, 2 failed,
1 skipped. Pester printed 937 passed, 0 failed, 1 skipped; normalization correctly rejected the
unallowed pinned-vcpkg-database skip. The exact pinned database has since been restored in this
isolated worktree. ViewerPE's 09-explorer-right-panel WGC foreground-preservation witness also
failed; this evidence cannot distinguish external focus change from activation by the viewer.
No focus assertion was weakened and no causal/flaky classification is claimed.

`failed-parent-initial` and `failed-parent-diagnostic` retain both fixture failures, traces, metrics,
all case statuses and original build identities. The diagnostic found PATH_NOT_FOUND after parallel
cancellation: the deep child could legitimately have been removed. The corrected witness probes a
surviving ancestor, with extended paths and bounded sharing-error handling; production cancellation
was not changed to satisfy a mistaken survival expectation.

`failed-smallfile-first` retains bridge success/counters rejected by the observed output verifier,
leading to the Dummy stride correction. `failed-search-old-guard` retains the stale source-text guard
failure before its migration and stronger runtime case. Earlier passing `parent-final-repeat3`,
`smallfile-final` and `xp09-final` belong to their recorded prior receipts; the current rows above
are fresh executions against the final source snapshot.

`failed-build-resource` and `failed-build-pdb` retain compiler-log digests/diagnostics for disk-space
and PDB failures. Only this operation's identified MSBuild children were stopped. The task's AWS build
cache was preserved on Z: to free C: space; other users' builds/processes were not terminated. A later
successful build does not establish a unique cause for the earlier environmental failures.
Development probe scratch was swept before curation; important header/tooling checks were rerun and
archived immediately. Original compiler logs remain in repository `.build/logs`.

Large metric streams are retained under the marked sandbox's `evidence/review-safety-raw` directory.
Their compact archives bind original SHA-256, byte/row counts, exact retained line numbers and
verbatim claim-bearing rows in original order. Nested `.gitattributes` disables text normalization.
Explicit pre-stage validation and complete-inventory validation passed; the latter checked 13,599 files. Final pre-stage validation also covers promotion of its own log and the final source-verification record. The final source fingerprint still matches the tested receipt; both unstaged and staged whitespace checks pass.

## Separate shutdown investigation and remaining qualification

`shutdown-probe` compiles the exact production scheduler prefix with a synthetic parked nested
child. Its closing-thread handshake is observed; shutdown does not return in 500 ms before child
release; after release/notification, parent and shutdown reach quiescence. A 15-second contained
process bound prevents stranding diagnostic workers. This demonstrates a scheduler condition,
not a reproduced valid app-exit or DLL-unload lifecycle. Production shutdown is unchanged.
Returning early from WaitJob alone could leave callbacks using stack contexts; an eventual fix
needs admission/producer closure, cancellation/abandonment and in-flight quiescence together.

Fresh Full, the five interactive obligations, full Release/ASan application builds, native ARM64
and real SMB/NAS qualification remain pending. No release or complete I36 closeout is claimed.
