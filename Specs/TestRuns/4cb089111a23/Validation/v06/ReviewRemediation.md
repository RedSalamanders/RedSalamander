# I27 File Operations safety remediation

Implementation branch: `codex/fileops-review-remediation`, based on `e860022a088f8b5fc3c604a61e091452cf83337e`.
Qualification remains open. All seven reviewed P1/P2 findings and the six architecture recommendations
are implemented. The pre-merge R21 x64 Debug and Release builds succeeded with zero warnings and errors;
source governance passed 239/239, including specs/tool inventories and formatting. R19 broad FileOps
and Compare passed in both configurations with explicit skips. R19 Release PR passed; Debug PR
failed one unchanged ViewerPE close-time limit. Final R21 held-key contracts passed three repetitions
in each configuration, and six isolated ViewerPE close checks passed. These outcomes do not replace
the failed PR or the pending final Full gates.

The user explicitly selected **Continue non-focus work for now** after Zoom took foreground
ownership during the isolated throughput test. Foreground regressions and fresh Debug/Release
Full qualification are deferred. The plan remains WIP; no whole-repository verdict or flawless
runtime claim is made. The second fresh Debug Full run failed 55 cases; the third recorded
2,204 passed, six failed and 51 skipped, with 18/20 entries promoted. Both remain failed historical
evidence. Final R21 performance retains the original baseline comparison, three unchanged-I26
controls under current conditions, and five current-condition interleaved I26/R21 pairs.
Performance non-regression remains unqualified because graph redraw observations worsened.

The historical I26 Done plan is preserved. This packet follows the seven actionable findings in
`C:/RedSalamander.Perf/evidence/i26-code-review-20261001-e860022a0/Review.md` and the user-authorized architecture work.

| Finding | Implemented behavior and evidence seam |
|---|---|
| P1 stale consent | Prompt generation belongs to snapshots, primary/overflow/scope/Escape/UIA bindings, safe focus and asynchronous inspection. The real same-task A-to-B fixture rejects retained A inputs and preserves B. An in-flight enumeration-settlement fixture rejects A's late selection/focus continuation after generation changes. |
| P1 duplicate decisions and late signals | `FileOperationsConflictBroker` owns generation-checked one-shot acceptance, consumption and manual-event state under the arbiter lock. Wakeups are hints; an empty/stale wake cannot manufacture Cancel. Deterministic broker tests exercise duplicate different/identical choices, policy/scope rejection and event ownership. |
| P1 false source retention | `FileOperationItemResultReducer` separates the host mutation gate, terminal cause and provider storage truth. Mutation-possible missing receipts retain unknown destructive-operation axes. Real Dummy Move/Delete fixtures mutate, then cancel with neither callbacks nor a receipt. Preparing items and stronger host proofs remain precise. |
| P2 path-detail focus | The extracted path-detail policy retains logical state and selection/caret without activating native text input while another window owns focus. The real pinned read-only TextField reconstruction and background focus matrix are exercised. |
| P2 hidden Issues work | Pure revisioned Issues projection replaces the window-owned correlation/sort path. Hidden/minimized windows stop the timer; unchanged visible revisions skip copying/projection. Lifecycle tests pump beyond two timer periods and prove refresh on show/restore. |
| P2 stale alert providers/actions | The existing single alert provider is fenced by HWND instance and presentation generation. Owner-thread action dispatch revalidates identity; retired property/pattern/action access is unavailable. The real alert journey covers stale and queued actions. |
| P2 incomplete native fallback | Hidden child creation plus both canonical button hooks publish transactionally. Every child/hook fault rolls back and retries. Reentrant initialization is rejected. Keyboard, theme/DPI, hide/reopen and Cancel all use the real fallback. |

The parent reviewed delegated findings against current callers and tests. A delegated claim that the
stale-primary replay was missing was rejected: the current test already explicitly replays it.

<!-- I27-ARCHITECTURE-MAPPING -->

| Architecture recommendation | Implemented ownership boundary |
|---|---|
| One conflict broker | `FileOperationsConflictBroker` owns generation validation, one-shot acceptance, consumption, and event state; the production adapter holds the arbiter lock. |
| Pure typed-result reducer | `FileOperationItemResultReducer` owns gate, terminal cause, receipt validity, publication, source, and stage classification; stronger bridge proof stays explicit. |
| Issues projection separate from HWND | `IssuesProjection` owns identity/correlation/sort and revision reuse; the pane owns visibility, selection, persistence, and rendering. |
| Logical and native focus distinct | The path-detail policy restores text/caret/logical state separately from permission to activate native text input. |
| One alert provider with presentation identity | The existing provider checks instance and presentation generation for properties, patterns, owner-thread dispatch, and queued actions. |
| Split by responsibility | Conflict adapters move to `State.Conflict.cpp`, path detail to `Popup.PathDetail.cpp`, and projection/reducer/broker tests have dedicated files. These extractions remove overlapping policy rather than only moving lines. |
Additional defects found while validating the fixes were corrected:

- Native Move synthesized Unknown publication without provider proof, then treated it as precise bridge evidence and fabricated Retained source truth. Only real bridge evidence now supplies that stronger classification.
- A qualified descendant Skip after possible mutation could imply Published without a receipt. Exact whole-root Skip and partial-tree Skip now remain separate; the latter preserves unknown storage axes without retrying the root. The pure production adapter matrix covers Copy/Move/Delete, subsequent cancellation and unknown/malformed receipts.
- Fallback hook installation failures were ignored. The existing checked Win32 hook helper now participates in the child transaction; both hook failure points are tested.
- D2D fallback initialization reentered during retained-host teardown. The old member-handle count masked duplicate child publication. The runtime witness records one rejected reentrant entry; the actual HWND count remains three and Tab reaches the intended Cancel all button. Three repeated journeys also pass.
- Bulk callback receipt validation accepted structurally valid but contradictory known no-commit receipts claiming an absent source or Published stage. The pure reducer now rejects them after possible mutation as Indeterminate, including after Skip/cancel. The new matrix and real callback/finalizer tests cover Copy/Move/Delete with valid controls. The first Full Debug attempt was deliberately interrupted to correct this reachable defect; it is preserved as incomplete.
- The Issues adapter lost programmatic grid selection when the next model revision arrived. It now captures the actual stable selection before reconciliation and after programmatic selection; real UIA clear/reselect and deleted-row detail assertions cover the integration.
- Recursive Local Copy workers kept mutation facts private. Successful merges with published children reported no-commit, correctly triggering the reducer's `E_UNEXPECTED` rejection. Worker contexts now borrow one item's atomic destination facts; top-level stage authority and uncertainty precedence remain intact. The one-shot/apply-to-all collision control also asserts the typed whole-item receipt.
- Local bulk Copy/Move callbacks omitted mutation receipts despite the provider retaining that evidence. Bulk Copy now resets item facts between sequential roots and reports indexed per-root receipts; bulk Move resets native rename truth and shares the single-item receipt mapping. Sequential and parallel two-root controls check exactly-once indexed callbacks, successful publication followed by failed non-publication, unknown cleanup, exact bytes and source state.
- The clipboard-ready fixture left its intentionally failed Move's completed popup for following families. It now joins before dispatching completion and dismisses only its own result through production cleanup, then asserts an idle state with no popup.
- Generic Find UI fixtures waited for a cold volume index within UI deadlines. They now use scan; dedicated backend controls remain indexed/service and idle failures report internal state. Production cold-index latency is unchanged. Unsupported Retry is tested as rejected consent, and live Skip uses the generation-bearing UIA action identity. The two-card B focus witness retains strict post-focus viewport assertions and adds lookup/focus diagnostics.
- The independent-volume control was blocked by missing runner authorization. Subsequent runs use the governed marked C: alternate root and explicit runner opt-in; no skip or assertion was weakened.
- The isolated two-card probe rejected the clipping hypothesis: lookup completed without finding B's action. The visual-layout helper rebuilt policy from a zero-generation prompt and erased both live prompt identities. It now preserves the input generation; the fixture explicitly asserts both presented generations before focusing the exact B action. The original failed probe is retained.

- Find header-click fixtures sent snapshot DIPs as Win32 pixels. At 144 DPI the intended headers were missed; three fixtures failed and five restored-descending fixtures could pass without sorting. The repaired fixtures use canonical coordinate conversion, verify the intended logical header, prove ascending order, then verify descending order and exact selected path.
- The connected acceptance fixture captured prompt A's generation in every provider lookup, including Cancel for the subsequent prompt B. Lookup now takes an explicit generation and matches the exact AutomationId; the next prompt must agree with its worker generation before Cancel. Original A diagnostics remain bound to A.
- Local Native Move reported a vanished source as known retained after admission. Source discovery/validation failures now record uncertainty in the existing monotonic native-mutation tracker, without adding a filesystem probe or another receipt-policy path. Missing-source native rename failures and conditional absent-source truth also remain unknown, while a proven commit survives later uncertainty. Real Local producer-to-host-callback/finalizer, single-provider and three-index bulk controls cover source disappearance, successful publication, and precise existing-destination refusal. The callback/finalizer seam is labelled separately from ExecuteOperation: its earlier exact-source guard catches removal before provider admission; a new production pause hook was not added.
- A final adversarial review found that an empty-directory Copy merge into an already-existing destination returned `S_OK` with a truthful no-commit receipt, which the host normalized to `E_UNEXPECTED`. The pure reducer now keeps completion independent of mutation: an exact clean Copy no-op is Completed/NotPublished/Retained. Move/Delete non-commit success, Copy `S_FALSE` without Skip, contradictory receipts and unresolved stage debt retain failure/indeterminate classifications. The completed-task summary derives Copied from actual completed Copy records as well as publication. The final targeted Debug and Release contract groups pass 15/0/0 each. Their safety case exercises 20 real controls: five empty/existing/sentinel/nested/new-destination scenarios, configured concurrency caps one and four, and two separately labelled seams (executor/host bridge and Local producer/callback/finalizer). The real asynchronous merge fixture also verifies the canonical Copied summary with zero published items and unchanged sentinel bytes.

The first R8-built targeted contract run (`debug-fileops-contracts-r7`) preserved 3 passed, 1 failed and 11 not-reached skips. Its new executor controls proved that normal Local Copy uses the host-owned bridge, with no provider completion callback. The bridge recorded `NotPublished`/`NotCreated` for every existing empty-tree merge, while the final result store overwrote that exact fact as `Published`. The success mapping now preserves the bridge's explicit publication proof. The runtime matrix separates actual executor/bridge results from actual Local producer/callback/finalizer receipts; the earlier `E_UNEXPECTED` defect belongs to the provider-receipt route. Original results, trace and bridge metrics remain diagnostic evidence, rather than being relabelled as a successful qualification.

The pre-no-op-fix targeted R6 qualification passed Debug contracts 15/0/0, parallel-copy matrix 3/0/0, Find 64/0/0 and conflict journeys 3/0/0; Release contracts 15/0/0 and one-shot overwrite 3/0/0 also passed. The owned orchestrator was then stopped before Full to add the no-op repair. Its interruption identity and all completed targeted outcomes are retained; these do not qualify the final rebuilt candidate.

The R6 rebuild caught a missing chrono-literal scope in the new Find helper. It was corrected before native qualification; both original failed build logs remain in diagnostic history.
Retained raw evidence is at `V:/RedSalamander.Perf/evidence/fileops-review-remediation-20261001/`;
the original large R3 raw metrics were removed by governed runner retention, as detailed below.
The curated packet preserves per-case statuses/skipped reasons, historical Full provenance, failed
witnesses and ordered metric claim rows. Twenty-three capture files retain all 6,325 integrated
claim rows through verified lossless gzip, with raw digests, sizes and original line indices.
Two historical failed-run traces are also preserved losslessly. Build success or source-shape tests
alone do not establish runtime correctness.

Performance uses the same machine (`4cb089111a23`), compiler/dependency identities, test-enabled
Embedded profile, unchanged DxUi API-2 pin, workload digests and settled 144-DPI geometry. The
baseline's dirty-checkout identity is preserved: its 687 reviewed native input blobs match I26;
the candidate's 971 native inputs match the primary checkout byte-for-byte. The fixture remains
blob `9b244a8830403e2dae192d3ce180f2613f34d0cc`. No baseline was silently replaced.

The retained original five-process I26 baseline versus initial final-R21 captures showed UI CPU
1,625,000 to 2,437,500 us (+50%), process CPU 1,953,125 to 2,796,875 us (+43.2%), and several redraw
series worsening by 25–32%. Those results remain `performance/comparison.json`. Three subsequent
unchanged-I26 processes consumed 2,750,000–2,984,375 us UI CPU and 3,218,750–3,406,250 us process CPU.
The unchanged binary's slowdown establishes changed measurement conditions, without proving that
the source change has no performance regression.

Five fresh current-condition pairs alternated process order: I26/R21, R21/I26, I26/R21, R21/I26,
I26/R21. All ten captures passed. Full build inventories were SHA-checked before and after every
capture, including all 134 source-bound candidate runtime modules. The comparison validates build
profiles, exact workload digest tokens, all eight geometry/workload controls, 29 resource phases,
275 integrated claim rows and all timing sample counts per capture. The capture driver omitted a
comparer binding file; it was derived from the existing SHA-checked records with subset validation,
without rerunning or inventing evidence. The original driver failure remains in the external log.

| Current-condition timing, eight synchronous redraw calls | I26 median us | R21 median us | Change |
|---|---:|---:|---:|
| Graph open | 102,804 | 133,009 | +29.38% |
| Graph collapsed | 75,606 | 97,663 | +29.17% |
| Churn update | 96,785.5 | 79,032.5 | -18.34% |
| Issues selected detail | 66,746 | 66,761 | +0.02% |
| Changing scan: Compare | 72,459 | 67,262 | -7.17% |
| Changing scan: case | 67,897 | 69,750 | +2.73% |
| Changing scan: attributes | 78,357 | 67,206 | -14.23% |
| Changing scan: file list | 67,120 | 76,156 | +13.46% |
| Clean scan: Compare | 66,587 | 66,022 | -0.85% |
| Clean scan: case | 71,951 | 66,744 | -7.24% |
| Clean scan: attributes | 68,156 | 68,748 | +0.87% |
| Clean scan: file list | 66,561 | 68,344 | +2.68% |

| Current-condition resource observation | I26 median | R21 median | Interpretation |
|---|---:|---:|---|
| Observed peak private bytes | 271,974,400 | 272,785,408 | +0.30%; labelled observations, not a continuous allocator peak. |
| Observed peak working set | 195,743,744 | 195,092,480 | -0.33%. |
| Scenario UI CPU us | 3,031,250 | 2,687,500 | -11.34%; source improvement is not established. |
| Scenario process CPU us | 3,640,625 | 3,062,500 | -15.88%; source improvement is not established. |
| Peak handles / GDI / USER | 2,685 / 59 / 78 | 2,685 / 59 / 78 | Same observed counts. |
| Closed minus cold handles / USER | +35 / +3 | +36 / +3 | Existing fixture residuals remain visible. |
| Closed minus cold GDI / DxUi hosts / attachments | 0 / 0 / 0 | 0 / 0 / 0 | Starting counts restored in every capture. |
| Closed minus cold private bytes | +118,784 | +3,854,336 | Candidate median +3,735,552 bytes; no memory-improvement claim. |

Graph-open capture medians span 82,203–208,008 us in I26 and 66,728–313,846 us in R21.
Graph-collapsed medians span 64,046–83,318 and 66,634–163,736 us. Graph-only UI CPU endpoint
differences also worsen: open 156,250 to 171,875 us (+10%), collapsed 125,000 to 171,875 us
(+37.5%). Whole-fixture component medians do not isolate that cost: retained-control update
1,646 to 1,460 us, Render 4,336.5 to 4,113.5 us, and WM_PAINT 8,190 to 8,173.5 us. All captures
record seven control reconstructions; candidate capture five records 543 popup paints versus
720 in the other nine captures. This is a further measurement/capture-variation limitation,
despite matching integrated claim controls. CPU counters have 15.625 ms granularity.

Parent-verified source attribution found no diff in `DrawBandwidthGraph` or the hosted graph
descriptor/`SetSamples` path against frozen I26. The generation-bearing hit adds one scalar
comparison per reconciled button. This small reachable common-path cost does not explain the
observed shift. The delegate first cited message pumps in the separate Commands throughput fixture;
that attribution was rejected. The actual identical integrated fixture has no pump inside its
eight-call redraw lambda or the measured open/collapsed loops.

**Performance status: NOT_QUALIFIED.** The graph and endpoint-memory worsening remains actionable
qualification debt. Whole-fixture CPU improvements and variable conditions cannot clear it.
A stable repeated graph attribution pass is needed before performance closeout; a proven added
cost must be optimized and measured against the retained fixture. Do not weaken safety behavior,
change the baseline or infer a leak solely from these one-shot endpoint-memory differences.
Synchronous redraw durations include presentation and scheduling waits and do not establish
presented FPS. GPU/driver allocation peaks remain unavailable.

Legacy raw `WorkloadFNV64.durationUs` fields contain opaque digests; `ClientPixels.durationUs`
contains DPI; `GuiAndAttachments.durationUs` contains an attachment count. Their writer-generated
`us` labels are not elapsed times. Typed controls preserve this distinction. The lossless claim
verifier validates all 6,325 JSON rows, exact decoded/stored hashes and original row ordering.

H4's user-deferred ARM64/native-ASan adoption, native 192-DPI qualification and external assistive-technology
execution are not claimed. HWND message journeys and physical key delivery remain distinct. The R18 physical Space check passed narrowly; the full 40-cycle foreground journey is unqualified. Native fallback
UIA pattern presence is exercised; an external screen reader is not inferred from that coverage.

The R4 staged-copy probe's receipt passed but its nonzero-stream worker assertion failed. The scheduler legitimately numbers a worker zero. The fixture now proves the exact child's progress callback executes off the calling thread; the failed original probe remains diagnostic history.

The two FileOps archive skips have explanatory text saying 7z coverage remains in Compare.
That is general provider coverage, not equivalent scenario qualification: the actual Compare
registry and results contain `directory_size_7z_callback_contract` and
`host_fallback_search_7z_name_only`, both passed in Debug. They do not test the skipped
large-entry read/seek or mounted-pane response to backing-archive changes. Those two legacy
scenarios remain unqualified; neither is counted as a pass or silently substituted by Compare.
This is an additional coverage limit, not a demonstrated production defect. The original skip
reasons are retained verbatim in the native results.

<!-- I27-R21-NONFOCUS-STATUS -->

## R21 non-focus qualification and remaining gates

At R21 qualification, tested native sources were frozen and byte-identical across 971 inputs in the
primary and isolated Release checkouts. The pre-merge build receipts are Debug
`56725964cd481d537d5a9b0bc949c388f83ab073026ed750393b4d00efe2d0bc` and Release
`284d47cf32a3872c22c23e0cce74a17b6d36bd81c20cc52773cae19bab9dc2fc`.
Both have zero warnings and errors; governance is 239/239. R20's intermediate thirty warnings
per configuration were corrected before R21; R20 did not receive native qualification.

| Gate | Outcome | Run identity |
|---|---|---|
| R19 Debug FileOps | 198 passed / 0 failed / 20 skipped; exit 0 | `20261002T115046Z-96672-cce73452147e4e6aaafc500be9408410` |
| R19 Release FileOps | 197 passed / 0 failed / 21 skipped; exit 0 | `20261002T122039Z-72320-f4c376879a994dfb90be9243c14ef78c` |
| R19 Debug Compare | 252 passed / 0 failed / 20 skipped; exit 0 | `20261002T123812Z-71864-4e3ab4019c95445a9f5e7a3b34a75787` |
| R19 Release Compare | 249 passed / 0 failed / 23 skipped; exit 0 | `20261002T131027Z-114164-4ffe556e480548b5be9f7d5ce24e320b` |
| R19 Debug PR | 827 passed / 1 failed / 0 skipped; exit 1. ViewerPE close threshold miss remains failed. | `20261002T131805Z-105988-c87a3abbe0c34a399b8df3e670a287b8` |
| R19 Release PR | 828 passed / 0 failed / 0 skipped; exit 0 | `20261002T133301Z-79928-b99b84b4c74a46fabf9350dd05e9b7a0` |
| R21 Debug safety contracts | 9 passed / 0 failed / 0 skipped, including three safety-case repetitions | `20261002T142704Z-114260-1f3c7b04e75e4eccb6ca3807d6538b43` |
| R21 Release safety contracts | 9 passed / 0 failed / 0 skipped, including three safety-case repetitions | `20261002T142817Z-110804-96d78ccee63d4c11a84f61643f4ddc4c` |
| R21 exact ViewerPE isolated checks | Debug close 54 / 51 / 57 ms; Release 50 / 61 / 50 ms. Six passed; unchanged 500 ms limit. | `nonfocus/viewer-pe-close-isolated-r21/results.json` |
| Final Release performance | Ten original/initial captures, three unchanged-I26 controls and ten interleaved captures passed. Performance NOT_QUALIFIED. | `performance/selection-manifest.json` |
| Foreground regressions and fresh Full | DEFERRED_BY_USER; required before closeout | - |

PR totals combine 816 Pester cases with twelve standalone-program entry outcomes; they are not
828 independent product scenarios. The safety totals include setup/sanity rows plus three actual
safety-case repetitions. Counts are preserved in their native reporting units.

All official non-focus runs record clean disk audits. FileOps skips include eighteen missing remote
credential profiles, two explicitly unimplemented legacy archive scenarios, and one Release-only
Debug-export limitation. Compare skips include missing remote/SMB/MTP prerequisites; Release has
three additional Debug-only search-hook cases, which passed in Debug. The native result files retain
each exact reason. No skip is counted as a pass.

R21 rejects held Escape/Return/Space/Delete before Path Detail and hosted-control dispatch, using
the [documented previous-key-state bit](https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-keydown)
and its [system-key equivalent](https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-syskeydown).
The actual non-activating owned-HWND fixture proves repeat rejection across an A-to-B binding change,
fresh-press activation controls, Path Detail Escape handling and repeated navigation without native
focus changes. A repeat Delete branch has no dedicated direct assertion. The separate Commands
case verifies the real generation-B engine remains unresolved after ordinary/system Escape repeats;
it is compiled and remains unexecuted under the user's foreground deferral.

The original Debug PR close miss printed no elapsed duration. Its functional assertions prove the
window closes while the provider read remains blocked and the worker later exits after release.
Source review found no worker join in Close/WM_NCDESTROY. The assertion now prints elapsed time;
the six narrow R21 passes do not establish the original miss's cause or make the failed PR green.

The cheaper independent source audit established no new reachable P1/P2 defect in the conflict broker,
result reducer, recursive per-root receipt aggregation, Issues projection, or alert lifetime fencing.
The parent verified its cited production and test paths. Existing-directory recursive merge cancellation
already asserts a published child, unchanged conflicting file, and Canceled/Published/Retained root truth.
A narrow combined case remains unqualified: cancellation cleanup of another active file stage after a
recursive sibling has published. Aggregation, canceled known commitment, direct-final abort, and merge
cancellation have separate controls; this gap does not establish a product defect.

Diagnostic foreground failures remain available: an exact Ghost-class witness identified the unpumped
long-paint fixture; canonical bounded pumping was added outside timed batches. R18 physical Space and
initial Unicode detail checks passed before Zoom took foreground at cycle ten. The current reentry fixture
also forces a fresh enumeration and waits for actual current-visit Selected status before its callback.
These repairs do not yet have fresh foreground qualification, and no assertion was weakened.

Large future raw metrics are digest-retained under the marked C: evidence root by the v2 copier before
normal runner retention. The original R3 whole perf JSONL was later removed by governed runner GC;
its preserved statuses, traces and selected claim rows remain, but that original raw file is unavailable.
The lossless-compression record is a historical operation record, not proof of current raw availability.

The user authorized committing and merging the implemented changes on 2026-10-02 with these
qualification limits recorded. I27 remains ACTIVE/WIP. Final Full, foreground qualification and
performance non-regression remain open; publication does not certify flawless runtime behavior.

## Merge integration, 2026-10-02

The remediation commit is `b883d5f8e506db40bc6b02e20bc7a9a1c4f4847e`. Integration with master
`7d61f4979c9d59a48060f4545696885b51756abb` retains its ViewerSpace/dependency updates and the
I26 history underlying the remediation. The sole content conflict was the WIP index; both the
updated D2 row and I27 row are retained. Auto-merged ViewerPE tests retain the elapsed-time
diagnostic alongside master's added tests, and Common message definitions retain both changes.

Merged governance passed 258/258 tests, including dependency and localization contracts. The
first attempt passed 254 and failed four dependency fixtures because the task wrapper's nested
TEMP/TMP prefix was too long. The rerun uses the official runner's short authorized root and
preserves every assertion. Its original failing log is retained as diagnostic history.
The initial format inventory also included archived C++ evidence copies; those are excluded
from active-source formatting and kept unchanged. Six inherited active files were formatted:
SettingsStore implementation/header, FolderView enumeration/header, RedSalamander entrypoint,
and LocalizationTests. The 258 checks were rerun after these format-only changes.
Historical I26 console/evidence whitespace is preserved byte-for-byte and excluded from the
source/active-spec whitespace check. Merge governance and archive receipts are retained under
`checks/`. Pre-merge native receipts remain historical: native builds and runtime gates were not
repeated for the merged dependency/source combination. The open qualification gates remain open.
