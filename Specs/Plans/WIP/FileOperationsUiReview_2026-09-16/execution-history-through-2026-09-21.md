# File Operations UI execution history through September 21

Historical checkpoints, preserved from the implementation plan. Their status and next-action
statements are superseded by the [current execution checkpoint](../../Done/UI_FileOperationsReview_2026-09-16.md#execution-checkpoint).
This record retains earlier failures, accepted slices, source identities and evidence links;
it is not a separate active plan or permission to waive an unfinished gate.

### Execution checkpoint

**Current, 2026-09-21 late qualification — Full failures retained; repairs underway:**
Fresh Full `20260921T190208Z-107796-357e522a509840379dc9046cc0ed6aa1` finished with
Commands 912 passed, 2 failed, 2 optional skips and FileOps 192 passed, 1 failed,
20 capability skips. All other entries promoted. Full is **not passed**.
The throughput reload assertion incorrectly required an offscreen card's painted summary;
the fixture now holds a short owned viewport, verifies remembered state separately, and
scrolls the new card into view. Disclosure/reload/restart pass 9/9 across three repetitions.
Selection visibility passes 3/3 alone and the surrounding `folder_view_` cluster passes
30/30, but its Full-only failure remains unclassified; additional state diagnostics are
being placed in the failing fixture. Independent C/D copying passes its unchanged
three-repeat control (9/9 including setup/cleanup), preserving the failed Full timing.
Its unequal pre-admission fixture I/O is being excluded consistently from timing without
changing the 1.35 ratio, data correctness, overlapping-progress or 30-second limits.
Test-enabled Debug receipt before these last diagnostic/timing edits:
`20153c27c17fe93b4c6892ce3f7b6efa0418058ff871eb8f55f7db2c1689a8f9`.
[Reviewed Full/follow-up evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-21-full-followup/README.md).
Raw run logs and retained cluster/volume artifacts:
`C:/RedSalamander.Perf/evidence/i26-ui/full-followup-20260921/`.

DxUi menu scrollbar reflow candidate `7ada987f3d211b2b6522ca18a5692d22512fbd98`
passes all six native CI profiles in run `35650904698`; reported desktop capability
skips must remain explicit. All six x64 Release gallery PNGs match the described
baseline byte-for-byte. Local paired resource qualification is running against retained
`356006a` with identical test instrumentation. The earlier roughly 1.15 MiB process
delta is not a 100-KiB menu-entry object: native objects are 248/760 bytes before
owned allocations, and released live-heap scaling is approximately 40 KiB per described
entry including native layout-related storage. This is not allocation-stack attribution.
**No menu or grid memory tradeoff is approved; no consumer pin has changed.**
Grid experiments have concluded and rejected candidates remain reverted at `0ae8362`.

**Next dependency-ready work (Astra High):** finish the local menu baseline/candidate
comparison, verify the timing/diagnostic repairs and broader Commands/FileOps order,
then qualify explicit library adoption and native product DPI/input/AT. Preserve the
independent Release app and use one local build/test lane. Final Full/Release,
paired resource acceptance, gallery review and normative closeout remain open.


**Previous, 2026-09-21 — actual throughput reload and process restart verified:**
The real main-window settings watcher applies external collapsed/expanded defaults
without resetting existing cards or queuing a save. Three owned application
processes exercise default-expanded→collapsed→expanded startup and shutdown,
preserving both popup placements, including position, dimensions, state and DPI.
No production behavior change was needed. Changed code: test-only owned-popup
teardown in `FolderWindow.FileOperations.State.Queue.cpp` and
`FolderWindow.FileOperationsInternal.h`; the two Commands fixtures and their
shared child-process include in `Commands.SelfTest.FileOps.cpp` / `.cpp`.
Full test-enabled x64 Debug build: zero warnings/errors, receipt
`9b380865cd8dfa0106bfc2a89b7c98e5868cbc0374c908391785422aa928dd27`.
Combined disclosure/reload/restart: **3/3, zero skipped**, including three child
processes with exit 0. Reload latency 302,794/293,271 µs; queued generation 25→25.
[Raw results, child receipts and limits](../../../TestRuns/4cb089111a23/Commands/2026-09-21-throughput-reload-restart/README.md).
This closes the functional throughput settings item; final configuration,
native DPI/AT and paired resource gates remain open. The runner's 260 disk-audit
findings are retained, not reported as clean.

**Follow-up, 2026-09-21 — DxUi native heap attribution completed:** the independent
grid worktree now has opt-in process-local heap diagnostics, paired unpaced and
50-fps-paced retention receipts, and four retained/reverted experimental patches.
DxUi diagnostic/evidence commit: `2ba0dfb` on `codex/grid-line-clamp`.
V11 adds about 1.17 MiB of median live heap data and 9.02 MiB of reported free heap
space in the unpaced pair. This supports substantial allocator retention but does
not identify native allocation call stacks or explain the separate menu candidate.
WPR setup was denied with 0x80070005; in-process sampling works without elevation.
The simpler one-layout prototype passes three Release suites and lowers the paced
private median from 36,831,232 to 34,004,992 bytes in one comparison, still above
original 29,743,104. It remains an unqualified saved patch; no production change
or consumer pin adoption is retained. DxUi's final diagnostic-only code passes
nine x64 suite/configuration combinations with zero skips and three ARM64
cross-builds. Reviewed evidence is in the grid worktree at
`Measurements/GridTextOverflow/2026-09-21/heap-attribution/README.md`.

**Completed investigation:** the targeted Extra High grid/resource investigation has
ended. All rejected experiments remain reverted; do not repeat them unchanged. The
current work is the menu resource comparison and remaining product gates listed above.

**Previous, 2026-09-21 — neutral Issues title and provider fallback focused-verified:**
The Issues native caption now uses neutral wording in embedded English and the
four existing translated overrides. The same-HWND English/French live-language
test checks that a pane containing pending, accepted and completed observations
does not title them all as failures. An unqualified provider collision now uses
its bucket sentence while keeping the exact destination path only in its
labelled comparison row; no provider separator or leaf is inferred. The
comparison UIA test confirms both the path-free question and exact role path.
Full test-enabled x64 Debug build: zero warnings/errors, receipt
`deb3e0b3a7efe78ea015df4f594fb2669ebd216dac31b3db70180d74d50015cb`.
Both focused Commands fixtures pass 1/1 with zero skips; unchanged French and
English [harness captures](implementation/2026-09-21-neutral-title-fallback/README.md)
and [run IDs](../../../TestRuns/4cb089111a23/Commands/2026-09-21-neutral-title-fallback/README.md)
are retained. The French Issues image still exposes the next tall-row sliver:
the separate DxUi grid/resource gate is open, not silently accepted.
Final Full/Release, native 96-DPI/AT and paired resource evidence remain open.

**Previous next action (completed above):** exercise actual external settings
hot reload and an owned process restart for the default-expanded throughput
preference. Use the production watcher/merge route and governed isolated
settings, not a direct `g_settings` replacement or HWND-only recreation. The
focused investigation at
`C:/RedSalamander.Perf/evidence/i26-ui/throughput-external-reload-next-slice.md`
names the existing fixture and cleanup hooks. Qualify baseline behavior before
changing production code; preserve unrelated settings and the independent
Release app. This slice involves settings lifetime/watcher state, so route it
through the supported Astra High continuation override. Do not infer parent P3
or final Full gates from the current focused passes.

**Previous, 2026-09-21 — Queue repair and visible conflict-order gates focused-verified:**
Queue fix `843eef39` now passes the 57-case Fairstream predecessor family, including
inline F2 after its original predecessors. Phase5 passes 10/10 and R4A19 5/5.
The newly exposed copy-merge assertion was a snapshot error: it compared unbounded
scroll-content coordinates with fixed decisions. The test snapshot now clips only
comparison regions to the actual viewport, retains fixed text bounds unchanged,
and reports raw/visible/action extents. Production rendering and decision policy
are unchanged. The copy-merge baseline fails at 203/156 ms; the correction passes
three repetitions (234/250/234 ms). The French comparison case passes and publishes
12 native 144-DPI harness images, including scrolling and keyboard reveals.
[Evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-21-conflict-visible-order/README.md),
[gallery](implementation/2026-09-21-visible-conflict-order/gallery.html).
Full test-enabled Debug build: zero warnings/errors; receipt
`adde4b7fef9085c460c8ab3762b3a60886cfaa92d20fd8dc10d58d6a627dda41`.
No complete Full/Release, 96-DPI/AT or resource pass is inferred. I26 remains ACTIVE.

**Previous, 2026-09-21 — inline F2 Queue exemption repaired; broad UI gate open:**
`ApplyQueueMode` and `UpdateQueuePausedTasks` now preserve the typed inline-rename
exception after admission. The strengthened real-interlock regression fails before
the change (188 ms), then passes three complete repetitions (703/625/734 ms), with
explicit user pause and cancellation checks intact. Full test-enabled Debug builds
pass with zero warnings/errors; candidate receipt `3ce4eb37ee840e24ad7dd1653902f1b0c2c9118e406e9e92d13d9f8afad3f092`.
The Phase5 discovery/queue family passes 10/10; the R4A19 family passes 5/5 after
correcting the missing alternate-volume runner flag. Both historical failing cases
now pass independently, but original-order qualification remains open.
[Paired results and limits](../../../TestRuns/4cb089111a23/FileOps/2026-09-21-inline-queue-exemption/README.md).

The broad FileOps run stopped immediately at `FileOps_CopyMergeIntoExistingFolder`:
2 passed, 1 failed, 210 not reached; `promptBeforeActions=false`. An exact classified
run reproduces it. Evidence and native archives are retained in
`C:/RedSalamander.Perf/evidence/i26-ui/inline-queue-broad-initial-20260921` and
`C:/RedSalamander.Perf/evidence/i26-ui/conflict-order-classification-20260921`.
**Next:** stay Astra High, reconcile the snapshot's unbounded scroll-content bounds
with the actual comparison viewport and fixed actions. Retain the failing baseline,
add meaningful clipped/visible geometry coverage and a harness capture; do not weaken
real overlap or clipping checks. Then run the normal FileOps order with
`-AllowAlternateVolumeTestRoot` and required final Full/Release qualification.
The Issues neutral-title resource slice, remaining persistence/interaction/native-DPI/
AT/resource acceptance and separate DxUi grid/menu gates remain open. I26 is ACTIVE.

**Previous, 2026-09-21 — retained Issues language changes implemented and focused-verified:**
the existing Issues HWND now refreshes its native title, primary/technical headings,
row labels, detail/actions, copy feedback, empty-state and accessible detail name
together. It keeps the same model, stable selection, logical focus, semantic sort,
column identities/order/widths and technical disclosure. Historical diagnostic text
and provider paths remain original facts. The Common localization manager exposes
a synchronized successful-application revision; unrelated theme notifications do
not reload translations or rebuild the row projection.

The unchanged three-repeat fixture first reproduced the defect **0/3**, then passed
**3/3** after each repair. The final File Operations Commands family passes
**50/50, no skips**, and native LocalizationTests passes **76 assertions**. All
four test-enabled x64 Debug builds have **0 warnings / 0 errors**; final focused
receipt `0135cb7402b71af67df1c3d713bd3b959d128412ef6a8127847d1bc0ad52e7b9`.
[Paired raw evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-21-issues-live-language/README.md)
and [five reviewed before/after harness images](implementation/2026-09-21-issues-live-language/README.md)
retain exact source/build provenance. Twelve samples per metric show language
notification median 4.450 → 8.982 ms (required new work), and five unchanged
notifications 2.220 → 2.011 ms after removing the first candidate's unnecessary
work. All 60 unchanged notifications publish zero projections. Variable tails and
memory/GPU resources are not qualified by these small timing samples.

Changed code: `Common/LocalizationManager.h`, `Common/Common/LocalizationManager.cpp`,
`FolderWindow.FileOperations.IssuesPane.cpp/.h`,
`SelfTest/Commands/Commands.SelfTest.FileOps.cpp`, and
`Tests/LocalizationTests/LocalizationTests.cpp`. Core localization/helper contracts
and the normative popup/Issues contract describe the implemented behavior.

**Full Debug qualification failed:** generated run
`20260921T144647Z-81208-b80aa09981ff4858af41b888e8ff80e5` exited 1.
The aggregate records 1997 passed, 2 failed and 37 skipped cases, with 19 promoted
entries and one incomplete entry (FileOperations). Commands passes 912/0/2,
including all 50 File Operations UI cases; both Commands skips are declared opt-ins.
The tooling-driven full rebuild has zero warnings/errors, receipt
`cdd541db0ebbb420296a0c4a3d34d048e26b089c3e18f39dc3eb5c2131fb4eb3`.
FileOperations failed `Floodgate_InlineF2WorkerQueueBypassAndInterlock` (180046 ms)
and `Phase5_DiscoveryCancelReleasesSlot` (703 ms), then stopped advancing at
`R4A19_DiscoveryProviderControls` for over six minutes. The owned no-activation
test process 90424 was terminated only after retaining its identity, traces,
metrics and a live minidump; this is interrupted coverage, not a successful suite
or a proven crash. The independent Release application was untouched.
Raw evidence remains in
`C:/RedSalamander.Perf/evidence/i26-ui/issues-live-locale-full-20260921`;
the complete runner log is
`C:/RedSalamander.Perf/evidence/i26-ui/issues-live-locale-full-generated-20260921.log`.
The failures remain open; focused localization passes do not waive Full.

**Next dependency-ready action:** remain on Astra High and repair the Queue exemption using the retained plan in `C:/RedSalamander.Perf/evidence/i26-ui/inline-queue-repair-plan.md`. Reproduce the added regression before the product change, then verify the discovery failure/stall independently. Only after these repairs,
correct the neutral Issues title identified in the reviewed captures as a bounded
Sol High resource slice (notes: `C:/RedSalamander.Perf/evidence/i26-ui/issues-neutral-title-next-slice.md`).
Then continue the remaining P3/P4 interaction and persistence acceptance, including a
real external-settings reload and process restart rather than substituting the
existing parse/reopen assertions. Keep the required Release/native mixed-DPI/AT,
paired resource and final gallery gates explicit. The independently launched
Release application must remain untouched. DxUi pin `78b3de3` is unchanged;
the separate grid and described-menu performance/resource gates are not accepted
or adopted. I26 remains ACTIVE.


**Previous, 2026-09-21 — live popup title localization implemented and focused-verified:**
`FolderWindow.FileOperations.Popup.cpp` now reloads its native title from the active
satellite on the same `WM_THEMECHANGED` notification used by settings reload.
`Commands.SelfTest.FileOps.cpp` switches an existing conflict popup en-US then
fr-FR and verifies both different localized titles without recreating the HWND;
`Specs/UI/UI_FileOperationsPopup.md` states that contract. The full test-enabled
x64 Debug solution build passed with **0 warnings / 0 errors**, receipt
`dd1423f609f589d9c0c266590ccd80498debe23c150dbd9940a5cc890516e410`.
The first build exposed a test-only missing `Win32Text::` qualifier; the corrected
build passed. Changed-hunk clang-format and `git diff --check` pass. The repeated
inspection predecessor cluster passes **9/9, no skips** and captures the French
native caption. [Reviewed before/after harness images and runner result](implementation/2026-09-21-live-caption/README.md)
retain build/scenario provenance; raw logs remain under
`C:/RedSalamander.Perf/evidence/i26-ui/`.

**Next dependency-ready action:** qualify the Issues pane's live language change
as its own bounded slice. Its native caption and cached grid headings/buttons do
not currently refresh together, so this popup fix must not be described as
complete live relocalization. Then continue the remaining P3/P4 behavior and
Fresh Full/Release/native mixed-DPI/AT/paired resource/gallery gates. The
independently launched Release application must not be terminated; menu memory
acceptance (~1.1 MiB on the measured fixture) still needs a scoped developer
decision or optimization evidence. No library pin changed, no broader I26 gate
is waived, and this plan remains ACTIVE. Do not repeat the unchanged 35-minute
canonical Commands run just for this caption change; the next broad run should
cover the next substantive UI slice.

**Previous checkpoint, 2026-09-21 — clipboard/inspection fixture slice qualified in canonical Commands:**
qualified product commits remain `27a1670ce`, `ecb391780`, `eed511720`; DxUi pin
`78b3de3` is unchanged. The new Commands test-support slice and Core/Testing
contracts are ready for the authorized commit. No product/library behavior was
changed by this slice. The RedPrism crash-handler comparison changed no files there.

- Native clipboard shuffled order: **9/9**, three repetitions; independent
  HGLOBAL/stream/bitmap storage, valid HTML/PNG/WIC seeds, three exact original
  restorations, five real read-contention recoveries and three reviewed French
  focus captures. Read/flush retries are bounded and only for CLIPBRD_E_CANT_OPEN;
  publication/product copy are never retried. Generation interference remains a
  failed gate that preserves newer clipboard contents. Injected failures are labelled.
- Conflict inspection's retained-name setup race is fixed by awaiting a new owner
  enumeration settlement and exact selection before testing reentry. The same
  File Operations family passes **50/50**, zero skips.
- Repeated predecessors first failed **8/9** with Selected/focusUnchanged=false.
  That guard refusal is retained without inventing its exact external cause.
  Instrumented confirmation passes **9/9** with all 12 handoffs preserving input
  ticks; canonical order also passes its four UIA/keyboard source/destination
  handoffs. The focus protection is unchanged. Other interference/DPI/AT gates
  remain in the parent acceptance matrix.
- **Canonical Commands: 912 passed / 0 failed / 2 declared opt-in skips**, 914
  cases, 34m 58.8s. The skips are ViewerSpace 20k-layout and large-performance
  opt-ins, not File Operations. Final full-solution x64 Debug build has zero
  warnings/errors, receipt
  `770e0bccf4ec5fa37a9be15c67aba3680b0d35245d1641742d8585e650a8c98c`.
  Changed-hunk clang-format reports zero replacements. All builds/tests are finished.

[Retained failure/pass sequence, results, exact provenance and reviewed focus image](../../../TestRuns/4cb089111a23/Commands/2026-09-21-clipboard-ownership-native/README.md).
Archive validation passes for 69 files. Full raw logs/captures remain in
`C:/RedSalamander.Perf/evidence/i26-ui/`, including
`canonical-commands-clipboard-inspection-20260921`. The earlier bitmap-preflight
archive is preserved. Disk-audit findings remain visible (canonical: 253); do not
claim Fresh Full, Release, whole-UI performance, or other-owner gates are complete.
The independently launched Release app PID 32648 was last observed at its output
path; recheck and never terminate it. The menu memory decision remains open.

**Next dependency-ready slice and model boundary:** use an actual supported
**Sol High** continuation in this same task for the bounded live popup-caption
localization fix. Queue at most one continuation after the authorized commit.
Reviewed actual inspection captures still have an English native caption after
switching retained content to French. The real app settings route posts
WM_THEMECHANGED; Popup::OnThemeChanged refreshes content but not its native title.
Apply the prepared, unapplied patch at
`C:/RedSalamander.Perf/evidence/i26-ui/live-popup-title.patch` after checking its
anchors; investigation notes are `live-locale-next-slice.md` beside it. It reuses
IDS_FILEOPS_POPUP_TITLE and GetWindowTextString, tests en-US then fr-FR on the same
window, and leaves decision/focus authority unchanged. Update the normative popup
contract, format changed hunks, freeze source/docs for one full test-enabled Debug
build, then run the existing three-case inspection predecessor cluster/repeated
French harness capture. Keep the before images. Issues has the same caption gap;
review its cached columns/buttons separately before claiming complete live locale
support. Do not merely update this plan; implement the prepared fix and tests.
If a new failure requires complex focus/lifetime work, checkpoint and route back
through an actual Astra High override. Afterwards continue P3/P4 and remaining
Fresh Full/Release/native mixed-DPI/AT/resource/gallery gates. No plan moves to
Done or other owner's deferral closes early. One writer/build/capture lane only.
**Superseded routing checkpoint:**
**Continuation routing checkpoint, 2026-09-21 — targeted review implemented:**
three UI refinements are committed (`27a1670ce`, `ecb391780`, `eed511720`).
The uncommitted consumer fixture and Core/Testing contracts now include these
corrections from the narrowly scoped clipboard ownership review:

- Every admitted HGLOBAL is copied independently as well as streams/bitmaps.
  An actual IStream release owner in the owned fixture changes its shared global
  buffer after snapshot and proves the original changed; no transferred raw
  storage is touched. Existing stream and bitmap source invalidation also run.
- CF_BITMAP comparison canonicalizes the header and RGB channels, excluding GDI's
  reserved high byte. The fixture includes nonzero reserved bytes.
- The fixture tracks its last admitted clipboard owner window and sequence,
  checks them after waits/readback and before every path copy/cleanup, and retains
  OLE materialization generations only while the publishing window remains owner.
  Before flush, reader-triggered delayed rendering may advance the sequence only
  with exact OleIsCurrentClipboard object identity and the original owner window;
  rejecting that legitimate materialization would prevent safe restoration.
  Path Ctrl+C must produce a new sequence, even when the copied value is unchanged.
  Process identity alone no longer permits restoration. Cleanup does not write
  when the fixture never published, rejects newer generations, and always restores
  focus/cursor through its own scope guard. Win32/OLE still has no atomic CAS.

This review changes **test support only**; no new product/library pin is adopted.
`git diff --check` passes. The earlier owned bitmap run was 1/1, but it does **not**
validate these subsequent changes. The latest previous native build was rejected
for an in-flight source edit, and no new build/test has run for this review yet.
[Retained preflight failure and earlier owned-data proof](../../../TestRuns/4cb089111a23/Commands/2026-09-21-clipboard-bitmap-preflight/README.md)
remain staged unchanged. No agent-owned build/test is running.

The targeted Extra High review is complete; return through an actual supported
**Astra High** continuation in this same task, with at most one queued continuation.
Next: finish any narrowly required formatting, then freeze code/docs during a full
Debug solution build (`RSBuildEnableTests=true`, `build.ps1 -Configuration Debug
-Platform x64 -MaxCpuCount 4`). Reuse that receipt with `Run-AllTests -SkipBuild`:
first `cmd_pane_fileops_clipboard_snapshot_medium_contract` (no clipboard mutation),
then only if it passes the original predecessor cluster
`cmd_pane_fileops_popup_progress_contracts,cmd_pane_fileops_popup_global_summary_ignores_finished_tasks,cmd_pane_fileops_popup_throughput_disclosure`.
Retain publish/flush/restore HRESULTs and exact restoration checks, then required
order/native confirmation. Do not change docs during builds or before planned
receipt reuse. Archive evidence before the next runner cleans stale runs, validate
archives, commit the reviewed fixture once qualified, then continue dependency-ready
P3/P4/Full/Release work. The menu memory decision and other owners' platform/resource
work remain open; do not waive failures or adopt unqualified pins.
**Earlier clipboard diagnostic continuation:** that build added stage-specific HRESULT
metrics for publish/flush and restore publish/flush, plus the round and any process
holding the open clipboard on failure. No clipboard contents are logged; the
existing full snapshot/preflight and restoration guard remain unchanged. Run the
same three-case predecessor cluster once to distinguish publication, OLE flush,
readback and external clipboard contention before choosing a repair. Do not waive
the retained 2/3 failure or use retry success alone as final qualification.

The diagnostic predecessor cluster also failed 2/3, but safely before mutation:
current clipboard format 2 (`CF_BITMAP`) returns `DV_E_TYMED` (0x80040069) in the
old byte-only snapshot. It never reached the new OLE diagnostics. Raw evidence:
`C:/RedSalamander.Perf/evidence/i26-ui/clipboard-flush-diagnostic-20260921`.
The candidate now deep-copies bounded GDI bitmap handles and canonicalizes their
pixels for exact comparison. An owned 2x2 fixture replaces the source bitmap through IDataObject
before source release, alongside the existing stream invalidation test. First run
only `cmd_pane_fileops_clipboard_snapshot_medium_contract` (no desktop clipboard
mutation). Do not resume native clipboard mutation until this passes and the
unresolved flush/restoration path has been reviewed. Spec inventory passed with
zero blocking findings before this bitmap addition.
The first bitmap build compiled cleanly but its receipt was invalidated by an
in-flight fixture ownership correction, so it supplies no test verdict. The stable
fixture replaces the source through IDataObject and proves the value changed;
it never touches bitmap storage after SetData(TRUE) transfers ownership. Run ID
`i26-clipboard-bitmap-owned-final-20260921` is the next non-mutating qualification.
That stable owned-object test passed 1/1, zero skips with a clean Debug build;
raw evidence is preserved under `C:/RedSalamander.Perf/evidence/i26-ui/clipboard-bitmap-owned-final-20260921`.
The next candidate reuses the existing clipboard quiet-point helper before OLE
publication/flush, verifies current-object identity, rejects snapshot sequence
changes before mutation, and preserves externally replaced clipboard contents
during cleanup. Focus/cursor restoration remains unconditional. Run the original
three-case predecessor cluster with these guards, retain all HRESULT metrics and
exact restoration evidence; earlier failed runs remain open until native checks pass.

**Current, 2026-09-21: Delete presentation passes; clipboard interaction failed.**
The unchanged gallery fixture first failed at 144/96 DPI and 480/760 DIP: Delete
had a graph/disclosure and no accessible item-rate text. Full baseline remains in
`C:/RedSalamander.Perf/evidence/i26-ui/deletion-graph-baseline-20260921/last_run`.
Fixture SHA-256 is `6611c6abfa3123f9a8fc50858e2ca6a2bb23ad7196a2f09630f65ff14eeddebe`.
The candidate restricts byte graphs to Copy/Move, removes Delete's empty plot
reserve and exposes its item rate through the existing retained text/UIA path.
Final full-solution Debug build passes with zero warnings/errors. The unchanged
native gallery passes 1/1, zero skips at 480/760 DIP on both screens. Reviewed
480-DIP frames show the 152-DIP card without path/bar/button overlap. Matched
redraw medians remain within 1.3%, without a broad resource claim.
[Captures](implementation/2026-09-21-deletion-progress/README.md)
and [receipts](../../../TestRuns/4cb089111a23/FileOps/2026-09-21-deletion-progress/README.md)
retain baseline failure and candidate pass. Commands remains red: progress and
finished-task checks pass, but throughput disclosure fails in its later clipboard
roundtrip (`Flush the owned rich-format snapshot before readback.`). Restoration
readback also fails (7 formats, generic E_FAIL). Next: retain exact publish/flush
HRESULT diagnostics and diagnose without repeating an unchanged mutation.
Preserve clipboard/focus restoration and the required gate. Open-discovery Delete
rate, P3/P4, native/Release/resource and Full remain open. No library pin changed;
menu memory acceptance is still pending.

**Previous checkpoint, 2026-09-21: verification aggregate correction validated.** The footer
and taskbar now weight the transfer/proof whole-task fraction separately from real
byte counters. New focused assertions cover mismatched proof totals, mixed task
sizes, open discovery, conflict exclusion, pause, completion and saturated counters.
Full-solution Debug build passed; focused progress/finished-task contracts passed
2/2 and the unchanged native gallery passed 1/1, all zero skips. The reviewed French
144/96-DPI captures now show 75% in both the task and footer. The unchanged fixture
retains a matched redraw pair with overlapping ranges and mixed median changes;
it does not close final performance/resource gates.
[Captures](implementation/2026-09-21-verification-progress/README.md)
and [receipts](../../../TestRuns/4cb089111a23/FileOps/2026-09-21-verification-progress/README.md)
preserve exact evidence. Next: remove Delete's empty byte graph while retaining its
item rate and compact measured layout; continue P3/P4 and final qualification.

**Previous checkpoint, 2026-09-21: French path-label fit repaired on both native displays.**
The unchanged gallery fixture first failed its actual paint-bound checks:
Suppression overflowed by 12.970 px at 144 DPI and 11.305 px at 96 DPI;
Vérification overflowed by 1.030 px at 96 DPI. The candidate measures the localized
label column using four bounded cache entries invalidated by text/format changes.
Delete now says Supprimés rather than Transférés, with all five satellites updated.
The x64 Debug candidate gallery passed 1/1, zero skips. Resource contracts passed
9/9; spec inventory reports zero blocking findings. An explicit full-solution
Debug build then passed; its verified receipt was reused successfully by the
focused progress/deletion contract (1/1, zero skips). Focused suite builds target
the application project; they do not create reusable full-solution receipts.
[Before/after captures](implementation/2026-09-21-path-label-fit/README.md)
and [paired evidence](../../../TestRuns/4cb089111a23/FileOps/2026-09-21-path-label-fit/README.md)
retain results and provenance. Five warm eight-redraw batches per size/DPI show
no consistent slowdown and zero cache misses during batches; these noisy Debug
samples do not close broad resource/performance or tail-latency gates. Scratch
audits remain non-clean (229/230 gallery, 231 focused), with raw evidence preserved
under `C:/RedSalamander.Perf/evidence/i26-ui/path-labels-*`.
**Next:** fix verification progress: the expanded card is 75% while the footer is
100%; the shared whole-task model bypasses verification. Preserve raw byte counts
and test mixed/unknown cohorts. The empty deletion byte graph also needs review.
Qualified DxUi pin remains unchanged. Menu resource acceptance, native input/
resource qualification, P3/P4 and final Full/Release remain open.

**Previous checkpoint, 2026-09-21: the independent French progress-row slice is implemented and
captured; P3 and final qualification remain open.** The consumer on qualified DxUi
pin `78b3de3` now shows separate discovery and transfer bytes/items in its two
existing card rows, including explicit discovery-on-demand after Skip. The prior
four-value French sentence fitted poorly and blurred the open discovered count
with completed transfer. The formatting remains satellite-localized in all five
supported languages. The focused French progress contract passed 1/1; the
deterministic `FileOps_VisualGallery` passed 1/1, both with no skips. Four actual
480-DIP French captures at native 144/96 DPI were reviewed and preserved in
[the implementation gallery](implementation/2026-09-21-progress-rows/README.md).
Both new rows fit in those captures. The application-project x64 Debug build reported
zero warnings/errors, resource contracts passed, and spec inventory found zero
blocking findings. Full logs are under
`C:/RedSalamander.Perf/evidence/i26-ui/progress-rows-*-20260921.log`; the two
successful harness runs reported 229 pre-existing scratch disk-audit issues,
which are retained without a clean-sandbox claim. The initial project-only
build compiled but its receipt was invalidated by an in-flight spec edit; a
stable project build subsequently passed. A `-SkipBuild` harness attempt lacked
a full-solution receipt, so the focused test and gallery each used their own
governed application-project build. This slice does not establish paired paint-cost
non-regression, all progress/verification states, screen-reader acceptance or
Full/Release qualification. **Next:** review transfer/verification and terminal
fixtures on both native displays, then repair the next measured P3 or P1
interaction defect. Do not adopt the unqualified described-menu or grid branches;
menu resource acceptance and native gates remain open. Route any shared layout,
focus/UIA or resource regression work to Astra High.
The next review already found a concrete P3/P1 defect in the same harness run:
`17-verifying__dark__480dip__fr-FR__display2__96dpi.png` clips the
`Vérification` path label in the former 56-DIP `De :` slot. Preserve that
actual capture as a baseline and fix the semantic label width/layout before
claiming verification presentation acceptance. The gallery run path is above.

**Previous checkpoint, 2026-09-21: described menus committed; independent P3 localization was ready.**
Library work is isolated at `Z:/src/DxUi-worktrees/i26-menu-description`, branch
`codex/menu-description-layout`, commits `b5b58b8` (implementation/evidence) and
`a474cf2` (byte-exact raw receipts), based on the qualified `78b3de3` pin. Its
`Specs/Plans/WIP/MenuDescriptions_2026-09-21.md` records wrapped primary/secondary
fields, full UIA identity, scrolling/DPI reflow, cancellation-safe invocation and
screen-reader focus preservation. All 16 nonactivating x64 suites passed in Debug,
Release and ASan Debug. Final root-menu focus/allocation guards pass nine focused suite
runs across those three configurations; all three final ARM64 cross-builds pass.
Skills/spec/dependency/format validation and regenerated gallery review pass.
These are library results, not application adoption or native monitor/AT acceptance.

Three alternating Release pairs show ~0.8–1.18 MiB more clean private memory despite
unchanged counted allocations, surfaces and hidden work. The common fixture does not
open menus, so retained menu text cannot explain this result. The user requested its
cause before deciding on the cost. A sixty-cycle matched driver now finds almost equal
warmed private medians (~32.8 MiB); the early difference is not a fixed retained increment.
Timing still fails qualification (late slowdown). The separate actual-menu probe at 96 DPI
finds ~1.15 MiB extra while twelve described entries are open, plus ~0.81 MiB for harness
capture. The user has been asked to accept only the measured open-menu cost after this
explanation; **no reply/waiver is recorded yet**. It would not accept other resource/native gates.
Full evidence and source/fixture identities are under the library's
`Measurements/MenuDescriptions/2026-09-21/README.md`.

The corrected physical-cursor fixture passed a direct Release Menu run with zero skips
and restored child focus/foreground/cursor. Release NativeTextInput then passed functionally
but failed cursor verification; foreground runs stopped rather than overwriting unexpected
movement. The latest root-menu focus guard now also runs the described-menu cases in
the foreground Menu suite; that final directed matrix remains open, alongside native ARM64
CI, paired acceptance and publication/handoff. All failed attempts remain explicit.
No build/test process from this task remains active. The library working tree is clean.
No consumer product source or pin has changed; the separate grid blocker remains unchanged.

**Previous dependency-ready routine slice (Sol High), now implemented:** P3 French progress presentation.
Use the unchanged qualified consumer pin and existing deterministic gallery baseline.
Find and fix remaining English or fragmented discovery/transfer/verification progress
sentences through the existing resource/formatting helpers, preserving operation state,
unknown-total semantics and policy. Retain any additional focused baseline needed before
code changes; extend meaningful French/long-text fixtures, run targeted tests/captures,
update normative/proposal/checkpoint evidence and commit. Do not use this slice to waive
menu/grid gates or adopt their branches. If work becomes generic rendering, native focus,
concurrency or authority policy, checkpoint and route back to Astra High. Continue other
dependency-ready P3/P1/P2/P5 work; one implementation agent and one native build lane.

**Continuation boundary, 2026-09-21: destination-history unchanged-product witness retained.**
The full x64 Debug solution build passed with zero warnings/errors. The deterministic
`FileOps_VisualGallery` run passed 1/1 with zero skips and disk-audit issues; receipt and
captures are under `S:/RedSalamander.Perf/evidence/i26-destination-history-baseline-v2-20260921`
(build log `C:/RedSalamander.Perf/evidence/i26-ui/destination-history-baseline-v2-build.log`).
The actual task-destination flyout was captured on display 1 at 144 DPI and display 2 at
96 DPI, in French, without foreground activation. Both menu PNGs show the historical
paths clipped before their distinguishing destination-folder text. The first attempted
run (`i26-destination-history-baseline-20260921`) stopped at `-SkipBuild` attestation
before app launch; the fresh full build and successful V2 receipt supersede it. No product
source, test fixture or DxUi pin changed in this slice; both working trees are clean.

**Next dependency-ready slice (Astra High):** current `MenuFlyoutItem` has one `text`
field, fixed one-line Radio height and a 456-DIP width cap. A folder-name/parent-path
destination entry with complete keyboard/UIA identity needs a generic DxUi menu-item
layout capability; shortening consumer labels alone would lose the distinguishing
parent and still clip long French names. Implement and qualify that capability in an
independent DxUi feature worktree with retained library baseline, focused menu/input
tests and resource measurements, then explicitly adopt a qualified pin in RedSalamander.
Preserve exact parsed destination authority, plugin/profile eligibility and the existing
real-button dispatch. Extend the gallery with repeated leaf names, roots and long French
paths; compare candidate screenshots against the archived 144/96-DPI baseline and verify
selection through the actual task button. Update the proposal, normative UI contract,
gallery and receipts. The separate grid-overflow worktree retains its unresolved resource
and ARM runtime gates; do not silently adopt it or waive those gates. Continue onward
through remaining P3/P1/P2/P5 items after this slice, with one implementation agent and
one native build lane. This is a model-routing boundary, not an overall Done claim.

**Continuation boundary, 2026-09-21:** consumer implementation/evidence committed as
`45f2c97ba`; the independent library's rejected cache experiment is committed as
`d6c8bb8` on `codex/grid-line-clamp`. Both working trees are clean. No native build/test
is running. The eight-slot experiment did not remove the original memory increase
and was reverted: restored multiline throughput is 121.13 FPS; retention median
33,613,824 bytes. All measurements/flags, the experimental patch, three passing
Release correctness suites and the restored control are retained under
`Z:/src/DxUi-worktrees/i26-grid-line-clamp/Measurements/GridTextOverflow/2026-09-21/cache8-rejected/`.
There is no new library pin or performance waiver; developer resource advice remains
unanswered. No need to repeat that question or retry this rejected cache reduction.

**Next bounded routine slice (Sol High):** P3 destination-history presentation.
Current `FileOperationsPopupState::ShowDestinationMenu` in `Popup.cpp` around
10132–10330 displays full historical paths. Present distinguishable folder + parent
location using existing provider-qualified display helpers and current menu capabilities;
preserve exact selected destination, plugin/profile eligibility, queue/mode semantics
and dispatch through the real task button. Extend the existing deterministic menu
fixture with repeated leaf names, roots and long French text, retain an unchanged-product
witness before implementation, then code/tests, harness captures, paired evidence and
normative/spec checkpoint updates. Do not infer provider identity from a display label.
If the next work requires generic controls, focus/lifetime, provider-policy or rendering
changes, checkpoint and route it to Astra High rather than extending this routine slice.
Continue through dependency-ready work; one agent, one native lane, at most one queued
continuation. This boundary is not a user-approval pause or an overall Done claim.


**Latest, 2026-09-21: exact inline-rename result/Issues projection passes 87 focused
Debug checks (9 FileOps, 75 Issues, 3 accepted-choice matrix repetitions), zero skips.**
The full build has zero warnings/errors. Real overwrite namespace/result identity,
completed-task TSV, JSONL, five typed completion states, Info-disabled first-truth,
duplicate/invalid references, late original/choice, contradictions and eviction are covered.
[Reviewed captures and scope](implementation/2026-09-21-typed-results/README.md).
Changed paths: FileOperationsInternal.h, State.cpp, State.Diagnostics.cpp, IssuesPane.cpp/.h,
Commands.SelfTest.FileOps.cpp and FileOperations.SelfTest.Fairstream.cpp; normative UI/FileOps
contracts now describe exact observation propagation. Copy/bridge/root outcomes without
an explicit carrier remain unknown. Existing conflict authority and the qualified pin are unchanged.
Raw final receipts: `S:/RedSalamander.Perf/evidence/i26-typed-result-{candidate,issues,decisions}-v2-20260921`;
retained failing baseline: `i26-typed-result-baseline-v2-20260921`; earlier passing V1 is retained too.
Native mixed-DPI/AT, Release/Full, broader transitions, library resource/native qualification and
final adoption remain open. Next: continue P3 destination history/persistence/progress while the library
resource direction remains open; the cache-size experiment was rejected and reverted. Do not close deferred owners.


**Current, 2026-09-21: throughput save/reopen and popup regressions pass 21/21 Debug;
DxUi V11 x64 functional/input qualification is complete, with resource acceptance open.**

The corrected same-task reopen witness exposed a real 72-pixel first-layout shrink,
failing all three unchanged-product repetitions. The candidate preserves the saved
initial expanded geometry and passes the broader 21-check popup matrix, zero skips/audit
findings/build warnings. It also covers failed save/recovery and both placement keys.
[Reviewed captures, exact receipts and timing caveats](implementation/2026-09-21-throughput-persistence/README.md).
All 16 current French harness PNGs were reviewed through their nine distinct images.
External hot reload/process restart and final Release/native-DPI/AT remain open.

Issues minimum-row fix `e925bf7f8` passes 72/72; filename-first detail/primary columns are
in `6a574bb8`/`790616ca`. [Row captures and remaining pin limitation](implementation/2026-09-21-issues-minimum-row/README.md).
The independent library worktree is `Z:/src/DxUi-worktrees/i26-grid-line-clamp`, branch
`codex/grid-line-clamp`, plan `Specs/Plans/WIP/GridTextOverflow_2026-09-21.md`.
Production `71d6446`, input checkpoint `725e83c`, paired ASan receipts `877db9c`.
All 48 non-activating x64 suite runs and all six Menu/NativeTextInput runs pass, zero
skips; input runs verify original child focus/foreground/cursor restoration. All three
ARM64 cross-builds pass; native runtime CI remains open. Six gallery images were reviewed.

Release multiline FPS improves 62 -> 122 but short-run private memory rises ~6.7 MiB.
Retention private usage oscillates 30.1–42.6 MiB versus original ~28.6 MiB; a settled
bound is unproven. Matched ASan common grid stays within bands; multiline retains
clean-frame memory/timing flags. Developer advice remains pending. No resource waiver,
publication or new pin is inferred; consumer pin stays `78b3de3`. Reviewed library
receipts are in `Measurements/GridTextOverflow/2026-09-21/`; raw logs in
`C:/RedSalamander.Perf/evidence/i26-ui/`. Preserve other owners' checkouts.

Next dependency-ready work: remaining P3 behavior/persistence, final French/native DPI/AT acceptance, paired
consumer resources and fresh Full. Resolve library resource/native gates before explicit
pin adoption and final Issues recapture. Six accepted decisions and ownership boundaries
remain unchanged; I26 stays ACTIVE and no other owner's deferred work is closed.

DxUi PR #18 merged as `78b3de389a189c7f86f611787e0489fb6d474218`, tree-identical to qualified
`b9c9390732c2b586d55b6089f190fda097cec4d2`. Main CI `35531706045` passes all six profiles after
one retained x64 ASan setup retry. Library closeout `4099dde` is Done. Consumer adoption and measured
actions were committed as `ca9f116`. The previous pin handoff needs no approval; the later grid resource advice remains pending. Live DxUi EditorConsumerControls changes
belong to another owner: preserve them; do not implicitly adopt that checkout.

Committed consumer slice `46c10f61b`: common name once; differing names each bold and once; stacked folders;
full measured scope/actions; independent comparison scroll; fixed inspection feedback; UIA/keyboard
reveal; retained-panel replacement on prompt-context changes. The typed guard preserves authority.
[Gallery, receipts and limitations](implementation/2026-09-21-bounded-conflict-comparison/README.md).

- Corrected baseline: 3 passed / 3 failed; each comparison hides decisions. The initial auto-expanding
  fixture was rejected as a short-window witness and retained separately.
- Final V11: 6/6 comparison/guard, 15/15 inspection/close/reopen/navigation; no skips or disk-audit
  issues. Full Debug build zero warnings/errors. Source/localization Pester 199/199.
- All 12 comparison PNGs and the inspection-feedback PNG directly reviewed. Keyboard captures have
  text-population/focus-stroke probes; a pixel comparison also finds identical fixed-explanation pixels
  in the ordinary, short before/after and keyboard before/after captures. Prior visual concern is not
  itself proof of clipping. Source/destination name/parent and actions remain complete.
- V11 verifies display device, monitor/work rectangles and DPI every measured frame. V7/V10 opaque
  HMONITOR mismatch was not a reliable physical-display identity check. V10's abandoned initial
  build scratch is preserved under evidence; independently launched applications were untouched.
- Redraw baseline median/p95 common 66550/68069 us, differing 66606/66906 us. V11: 66557/1817984 and
  66542/67087 us. **Performance remains unresolved**: multiple 250–520 ms hosted-paint stalls occur
  with unchanged display facts. First V6 crossed DPI and was invalid as a matched timing sample;
  V11 cannot be dismissed on that basis. All samples remain; short-window CPU metrics are unpaired.
- Curated final archive: `Specs/TestRuns/4cb089111a23/Commands/2026-09-21_020725-i26-bounded-body-candidate-v11`;
  corrected baseline, V6 and V6-confirm archives remain alongside it. Raw runs:
  `S:/RedSalamander.Perf/evidence/i26-bounded-body-*-20260921`; scripts/logs/live checkpoint:
  `C:/RedSalamander.Perf/evidence/i26-ui`. Next: connect the pinned DxUi frame-stage metric sink in
  the comparison harness (V12 now instrumented), separate presentation stalls from render/update, retain qualified evidence
  and fix a demonstrated product regression. No performance baseline is silently replaced.

- V12 frame telemetry: 5/6, one repeated keyboard capture reports text 304/304 but focus 0. Its PNG
  was overwritten by the next repetition; do not present the last PNG as that failure. V13 gives each
  repeat its own folder and records pre/post capture modality, focused ID, foreground, pointer and
  scope geometry. The new initializer warning is corrected. Product behavior is unchanged.
  V12 totals/render/update/presentation each have 880 samples: presentation max 5389 us, render max
  286784 us; the V11 repeated 250–520 ms stall is not reproduced. The isolated render stall and raw
  batches remain retained; no performance pass or OS-occlusion root cause is inferred.

- V13: 5/6, zero warnings/errors and disk-audit issues. Repeat 1 now preserves its missing-focus
  PNG independently. Text 304/304; focus 0; keyboard modality, focused ID, foreground, cursor and
  scope geometry stay unchanged across WGC. This is a retained focus failure, not established user
  interference. Next: inspect actual control HasFocus and host reduced-motion state; wait for the
  posted theme update before assuming a reduced-motion fixture. Continue the independent typed
  decision producer slice; baseline and candidates are now completed below.
  V12/V13 curated archives: `2026-09-21_021829-i26-bounded-body-candidate-v12` and
  `2026-09-21_022622-i26-bounded-body-candidate-v13`.

- Accepted-decision producer completed: V4 passes 6/6 repeated Commands checks; existing conflict
  family passes 33/33; source/localization Pester 199/199. All builds have zero warnings/errors;
  final runs have zero skips/disk-audit issues. Eight real scenarios cover explicit Replace/Skip/
  Keep Both/Cancel, invalid ingress, cancellation without choice, cached replacement and cached Skip All.
  Worker consumption snapshots exact task/prompt/root/scope/origin; closes the prompt before logging;
  retains the original conflict HRESULT and bounded records with Info disabled. Exports preserve all
  24 records (18 user, 6 cached); independent JSON parsing verifies unique IDs and exact grant links.
  No result correlation or mutation authority is inferred. Changed source: FileOperationsInternal.h,
  State.cpp, State.Diagnostics.cpp, Commands.SelfTest.FileOps.cpp; Popup.h/.cpp add observation only.
- Final receipts: [Commands and structured logs](../../../TestRuns/4cb089111a23/Commands/2026-09-21_030040-i26-decision-candidate-v4/README.md),
  [existing conflict family](../../../TestRuns/4cb089111a23/FileOps/2026-09-21_030118-i26-decision-conflict-regression-v4/README.md).
  Baseline/V1/V2/V3 remain alongside the final Commands archive. Original baseline fails three missing
  record checks; V1/V2 pass 6/6. Replace acceptance-to-finish baseline median/p95 52937/53214 us;
  V4 31853/59478 us, three samples each. This is paired end-to-end evidence, not a CPU/resource or
  statistically established non-regression claim; retain the other candidate repetitions/tails.
- Focus fixture correction: set canonical reduced-motion preference as well as direct theme override,
  restore it after the fixture. V3 has all three keyboard pixel probes pass (text 304, focus 1294),
  actual focus and reduced-motion true before/after; all three PNGs directly reviewed and retained in
  [V3 evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-21_025520-i26-decision-candidate-v3/README.md).
  V3 overall is **8/9**, because repeat 2 still fails display/DPI stability: DISPLAY5 work area
  2560x672, observed 96/144 DPI transition. Do not label this native work-area or performance acceptance.
- Accepted-decision slice committed as `510c1fb3`. Original-problem references now pass 6/6
  native repeated checks, existing conflict family 33/33 and source/localization Pester 199/199;
  Debug build zero warnings/errors and zero skips/audit findings. Baseline: 24 original problems
  lacked a typed prompt reference. Candidate: 24 original problems, 18 explicit problem/choice
  pairs, 6 cached applications and 6 problems without an accepted choice. Independent JSON parsing
  verifies identity/class/root/path/condition correspondence; keys are scoped to one producing
  FileOperationState lifetime. Merged historical logs require separate run provenance.
  [Producer and paired timing](../../../TestRuns/4cb089111a23/Commands/2026-09-21_031620-i26-problem-reference-v1/README.md);
  [policy regression](../../../TestRuns/4cb089111a23/FileOps/2026-09-21_031702-i26-problem-conflict-regression-v1/README.md).
  Original-problem/accepted records snapshot under the arbiter lock and enqueue after unlock.
  JSONL/TSV extend additively. No Issues UI or later mutation-result correlation is claimed.
- Issues typed projection is implemented and verified: V2 **9/9**, zero skips/audit issues,
  full Debug build zero warnings/errors. The baseline has 3 passed / 3 intended failures: four logical
  issues collapse to two rows. Explicit prompt identity now keeps same-path generations separate,
  retains cached applications, merges late original warnings, rejects mismatched/contradictory choices,
  and preserves selection. Eviction clears selection; opaque legacy strings use length-prefixed keys.
  Unchanged sorted data does not rebuild the grid. This is a data prerequisite, not final P4 layout.
  [Paired captures and timing](implementation/2026-09-21-issues-projection/README.md).
  V2 refresh median/p95 865/1096 us versus incomplete baseline 504/555 us; unchanged V2 426/452 us.
  The full-data projection adds work; broad paired resource/Release qualification remains open.
  Curated archives: `2026-09-21_034340-i26-projection-baseline-v1`, `2026-09-21_035007-i26-projection-candidate-v1`,
  `2026-09-21_035454-i26-projection-candidate-v2` under the machine Commands archive. All 42 archive files
  validate. Pester 199/199 and tool-inventory validation pass. The source-contract guard now protects
  typed association and collision-free legacy keys.
- Projection commit: `80cd474f3`. Selected detail implemented and qualified in focused Debug:
  **V8 63/63** across three repetitions; zero skips/audit issues, full solution zero warnings/errors,
  source/localization Pester **199/199**, tool inventory valid. Final archive `2026-09-21_050323-i26-detail-candidate-v8`.
  [Actual French captures and evidence](implementation/2026-09-21-issues-detail/README.md).
- Complete measured detail: problem, operation, accepted choice/origin/scope, explicit unknown result,
  exact Source above Destination, independent scroll, fixed labelled Copy/Show/Hide actions. Shared
  action resources and diagnostic serializer are reused. Normal strong decision foreground fixes the
  informational color being used against the wrong background. No later result is inferred.
- Corrected host-consumed size/DPI dispatch and host-coordinate child layout. WM_NCDESTROY saves view
  state before host Detach; factory owns adapter through failed CreateWindowEx. Repeated forced
  creation failures and normal closes return both adapter and attached-host counters to baseline.
- Interaction fixtures now target the process-owned HWND. Unchanged refresh need not rebuild. Routed
  Tab traverses grid/detail/technical/copy in both directions, preserving selection; Shift is sent to
  the host's modifier tracker. Wider columns can reduce visible column count; subsequent sort/reopen
  must preserve the resized count. Width/order/selection/copy assertions remain. Text-layout observers
  are opt-in so ordinary churn snapshots do not create DirectWrite layouts.
- Retained failures: V1 visuals/overload warning; V2 footer outside actual client; V3 42/63; V4 6/9
  (reverse fixture omitted Shift messages); V5 48/63 (viewport count assumptions); V6 diagnostic-name
  compilation failure; V7 63/63 with C5246 corrected in V8. Raw logs/scripts:
  `C:/RedSalamander.Perf/evidence/i26-ui`; raw runs `S:/RedSalamander.Perf/evidence/i26-detail-*-20260921`.
  Every applied edit script is non-idempotent; do not rerun it.
- V8 refresh median/p95 3321.0/10408 us; unchanged 423.0/524 us.
  Pre-detail V2 865/1096 us; added content/work means broad paired resource/performance qualification
  remains open. Native 144-DPI captures are not native mixed-DPI/192/AT/Release acceptance.
- Selected detail committed as `8c9655d0a`; all 15 final PNGs directly inspected and 186 curated
  detail receipt files validate. The working tree was clean before the next fixture.
- Provider-qualified display observations now pass **9/9** Commands checks plus **33/33** existing
  conflict-policy checks, zero skips/audit issues; full Debug build zero warnings/errors and Pester
  **199/199**. Original and accepted/cached records copy four qualified name/parent strings; they do
  not retain inspection capabilities or infer remote path syntax. JSONL/TSV use the shared additive
  prompt object. Independently parsed 48 original + 48 accepted records, with 12 cached applications
  carrying their own file's facts. This prepares filename-first presentation without changing policy.
- Baseline **6 passed / 3 expected failures**; V1 **8/9** exposed an invalid fixture assumption that
  payload.txt always reaches the prompt before second-payload.txt. Raw JSON proves either worker
  may arrive first. V2 compares the exact prompted local fixture item; cached application must name
  the other item. All assertions pass; no production ordering constraint was added.
- Display receipts: Commands `2026-09-21_051102-i26-display-baseline-v1`,
  `2026-09-21_051602-i26-display-candidate-v1`, `2026-09-21_052148-i26-display-candidate-v2`;
  FileOps `2026-09-21_052255-i26-display-conflict-regression-v2`. Raw logs/scripts under the same
  i26-ui/evidence roots. Do not rerun the applied edit scripts.
- Matched unchanged accepted-record case, 24 acceptance-to-finish samples: baseline median/p95
  49269.5/61160 us; candidate 51862/66182 us. Paired rows preserve action/cache tags and case order.
  This is bounded latency evidence (+5.3% median/+8.2% p95), not CPU/memory/renderer/Release approval.
- Filename-first selected detail now passes **69/69** complete Issues-family checks (three repeats),
  zero skips/audit findings, Debug zero warnings/errors and Pester **199/199**. Baseline has 3 passes /
  6 intended missing-subject failures. Common name appears once above the problem; differing names
  are separately labelled. Qualified parents remove duplicate filenames; copy retains full paths;
  incomplete pairs retain exact fallback paths. UIA child order follows the visual order.
  [24 directly reviewed French captures](implementation/2026-09-21-issues-subject/README.md).
  Archives: `2026-09-21_053336-i26-subject-baseline-v1`, `2026-09-21_053809-i26-subject-candidate-v1`.
  Aggregate refresh 3077/9970 us median/p95; unchanged 430/840 us. Additional fixtures change the mix;
  this is not paired overall performance approval. Native 144 DPI does not close mixed-DPI/192/AT.
- Primary Issues projection V1 passes **72/72**, zero skips/audit findings, Debug zero warnings/errors,
  Pester **199/199**. Default file/problem/operation/decision/result, narrow operation-in-detail,
  old technical layout migration, hidden-sort label and opt-in technical interactions pass.
  Archives: `2026-09-21_055127-i26-primary-baseline-v2` (69 pass / 3 intended failures),
  `2026-09-21_055754-i26-primary-candidate-v1`. V1 filter attempt matched no cases; raw logs retained.
  **Visual acceptance remains open:** short summary problem text ends in a clipped partial line.
  Inspection of pinned DxUi Grid shows SetLineClamp only enables wrapping and cell layout uses the
  clipped rectangle. Fix the generic primitive; do not shorten or mutate diagnostic model text.
- New library follow-up: `Z:/src/DxUi-worktrees/i26-grid-line-clamp`, branch `codex/grid-line-clamp`,
  indexed `Specs/Plans/WIP/GridTextOverflow_2026-09-21.md`. Based on qualified 78b3 plus documentation
  closeout c52a8f5. Other-owner DxUi editor work is untouched. Consumer pin remains 78b3.
  Next: unchanged-production library pixel/performance baseline, bounded complete-line ellipsis,
  full-cell clipping, all library gates and explicit qualified pin adoption, then consumer recapture.
- No native consumer supervisor remains active. Primary columns are implemented but not visually accepted;
  resume the library grid follow-up above before final consumer qualification.
  Continue exact later-result correlation, P3, native display/AT, paired resources/Release and fresh Full.
  I26 remains WIP; other-owner DxUi changes and independently launched apps remain untouched.


Mandatory Full `20260920T195803Z-8452-e970f56506f94f15982bebe45684b956` remains FAILED:
20 entries, 15 promoted, 5 incomplete; 1836 passed / 54 failed / 21 skipped. Governed results remain
under `S:/RedSalamander.Perf/evidence/runs/<id>`; raw cleanup removed some captures/traces. Bounded
Compare callback correction passes (8 ms case / 31 ms suite). Whitespace-tolerant typed guards pass
repeatedly. FileOps fake-Drive repeats retain 6/6; Full last-family result loss is not a universal
aggregation failure. Investigate final publication/full-size coverage without reading live result
JSON or claiming a root cause. Settings/preview, Find UI, terminal, Change Case timing and ViewerPE
interactive still need focused reproduction/correction.

Next after this slice: file-first Issues columns, typed result correlation, P3 remaining implementation,
multi-card/removed-action, native DPI/focus/AT, paired resources/Release and fresh Full qualification.
Keep I18/I25/H4/Editor/RedXe ownership separate. Preserve independently launched Release applications;
use isolated outputs for Release qualification. I26 cannot move to Done yet.

### Historical checkpoints (superseded where stated above)

**Main CI diagnostic, 2026-09-20:** initial main job `106133346270` failed Accessibility in
x64 ASan while `ElementFromHandle` was still pending at the unchanged 3000 ms deadline. No
notification assertion was reached. Retain the failure (DxUi `Measurements/LocalizedAdaptiveLayout/
2026-09-20/main-78b3-first-asan`). All other jobs passed, including all three ARM64 profiles.
One diagnostic retry is now running as attempt 2, job `106137812800`; ten unchanged local ASan
Accessibility runs passed (93–109 ms subscription). Do not claim a root-cause fix or adopt before
current-main CI succeeds. Consumer resource validation also
exposed 11 missing inspection-action/outcome IDs in each Czech, Slovak and Japanese satellite.
The 33 translations were added; the same ResourceLocalizationContracts suite now passes 9/9,
zero skips (previous 8/9 failure retained). Log: `C:/RedSalamander.Perf/evidence/i26-ui/
conflict-inspection-resource-tests.log`. The prepared UI implementation still awaits pin adoption.

**Approved handoff in progress, 2026-09-20 19:15 UTC:** the user explicitly approved step 1
("1. go"). PR #18 was merged normally, matching tested head `b9c9390732c2b586d55b6089f190fda097cec4d2`,
as main commit `78b3de389a189c7f86f611787e0489fb6d474218`. The merge tree is identical to the
qualified PR tree. The previous pending-approval statements below are historical and superseded;
no further merge approval is needed. Main push CI `35531706045` is running. After it passes, run the
standard RedSalamander updater and its Full validation, then the focused prepared UI matrix.
Preserve rollback pin `f5c7fc7403e352425b8e82b7a4da1c042d942ede`, all unrelated edits and independently
launched applications. No plan closeout is claimed by merging alone.


- **Qualification update, 2026-09-20 18:39 UTC:** final DxUi revision `b9c9390732c2b586d55b6089f190fda097cec4d2`
  passes native push CI `35526329502` and PR CI `35526332077` in all six x64/ARM64 configurations.
  Eighteen suite receipts per profile, architecture, exit codes and ASan detection were verified;
  x64 has zero capability skips, ARM64 Menu retains nine per profile. External-consumer and gallery
  jobs pass. Raw receipts are retained under DxUi `Measurements/LocalizedAdaptiveLayout/2026-09-20/native-ci-b9c9390`.
  No new implementation failure was found. The previously requested PR #18 shared-main merge approval
  remains pending; the consumer pin and prepared application changes are unchanged. Do not re-request
  that approval, infer it from a heartbeat, rebuild against the obsolete pin, or repeat unchanged tests.
  Next: approved merge, successful main CI, normal pin updater, then the prepared focused application
  matrix and actual French width/name captures. Keep the complete spec ACTIVE until all later gates pass.


- **Current prepared slice, 2026-09-20: measured conflict actions and complete repeat scope.**
  The baseline at `9d4c5bee` adds a real hosted-text fit assertion and reproduces the reported
  480-DIP French clipping in all three comparison repetitions (15/18 cases pass overall).
  Retained archive: `Specs/TestRuns/4cb089111a23/Commands/2026-09-20_190204-i26-conflict-actions-baseline`;
  raw runner/build/scenario evidence: `S:/RedSalamander.Perf/evidence/i26-conflict-actions-baseline-20260920`.
  Baseline common/differing-name redraw median/p95: 66607/66740 and 66578/66836 microseconds,
  45 samples each. The earlier viewport and inspection receipts remain valid for their own source.
  Prepared code replaces equal-width action slots with cached measured `ArrangeMeasuredActions`
  bounds and actual wrapped Body-font heights. It reserves the scrollbar gutter consistently,
  preserves authoritative action order/defaults and reuses those rectangles for hosted input/UIA.
  Repeat scope uses the complete localized sentence, explicitly limited to remaining similar
  conflicts in this task, measured with Checkbox's indicator/padding and retained separately from
  actions. French, English, Czech, Slovak and Japanese resources preserve that meaning.
  These consumer changes are **not yet built or accepted**: the pin remains `f5c7fc7403e352425b8e82b7a4da1c042d942ede`.
  DxUi `54e956e` passed all six native CI configurations, but full-caption review subsequently
  exposed Checkbox ignoring inherited multiline mode. The library correction, negative pixel
  witness, paired performance investigation and new gallery tile are being qualified before handoff.
  Pending shared-main merge approval was requested once for PR #18; the final change now includes
  this caption correction. Final prepared library revision: `b9c9390732c2b586d55b6089f190fda097cec4d2`,
  pushed to PR #18. All 18 x64 Debug/Release/ASan suites pass (nine Menu capability skips per profile,
  none elsewhere); the ASan probe passes. Five gallery themes are reviewed; skills/specs/format and
  dependency validation from a clean tracked export pass. Native push CI `35526329502` and PR CI
  `35526332077` qualify this exact revision next. The original native six-profile matrix at `54e956e`
  passed and remains archived independently. No merge, pin bypass or duplicate implementation agent
  is authorized by an elapsed wait. After a qualified approved main commit, use the standard updater, run the focused
  18-case matrix, inspect all six width/name captures, retain candidates and compare to this baseline.
  Continue the short-card body/action and typed decision/Issues work after the visible action slice.


- **Current checkpoint, 2026-09-20: body viewport and first-frame collapse implemented.**
  Retained task controls use an existing DxUi ScrollPanel with the same clip as the legacy body;
  footer controls remain outside it. Native scroll state stays authoritative. Tab reveals an
  offscreen action including 4 DIP of focus chrome. Interactive resize cancels pending automatic
  expansion and preserves the explicit height. Checked scope survives the guarded inspection return.
  Completed-card content and child controls retire on the first collapsed frame while height
  continues its bounded transition. The first-frame regression no longer waits for animation completion.
  [Accepted receipt](../../../TestRuns/4cb089111a23/Commands/2026-09-20_184156-i26-hosted-body-accepted/README.md): **18/18**,
  zero compiler warnings/errors, skips or audit issues. Initial candidate 15/18 exposed auto-resize;
  final 18/18 was visually rejected for ring clipping; ring 17/18 exposed stale collapse children.
  Every iteration is retained. [Reviewed gallery](implementation/2026-09-20-hosted-body/gallery.html)
  contains two rejected and one accepted harness frame. Common/differing renderer median/p95:
  baseline 66585/66788 and 66522/66782 us; accepted 66569/66783 and 66583/66708 (45 samples each).
  This is not final resource/native/configuration qualification; original historical tails remain.
  Source snapshot `06a35d5fd49640ec32ee67dee0ed52fb4c61264e48abe52eb5ebf5b8f0ec7248`.
  Changed code: Popup.{h,cpp} and Commands.SelfTest.FileOps.cpp. Raw evidence:
  `S:/RedSalamander.Perf/evidence/i26-hosted-body-{candidate,final,ring,accepted}-20260920`.
  **Next:** qualified measured whole-action layout and independent oversized-card action regions;
  typed decision/Issues work, final Release/Fresh Full, native DPI/AT and resources remain open.
  DxUi PR #18 now has a subscription diagnostic at `1ea876fc`; its previous duplicate push CI
  failed native UIA setup with no HRESULT. New local Release passes all 18 suites (Menu nine explicit
  foreground skips), subscription 94 ms / S_OK; updated remote x64 Debug/Release pass. Other native
  CI jobs and a new library invalidation/selected-detail fixture are being qualified. Do not claim
  the original failure's cause is known. Exact-main and explicit pin adoption remain required.
  Preserve unrelated DxUi Editor plan/index edits. Continue without waiting for another heartbeat.

- **Previous checkpoint, 2026-09-20: guarded inspection focus/return and keyboard implemented.**
  One-shot owner ticket retains post-selection visit, live adapter/context and input/foreground witnesses;
  final focus commit rechecks these and prompt/lifetime across synchronous Windows calls. Successful
  foreground inspection focuses the exact item and hides the popup; the normal View command restores
  actual popup foreground and the same control, scroll, scope and default/cancel actions. Background
  inspection stays non-activating. Space no longer invokes the general task-pause shortcut on decision/
  inspection buttons. Both UIA roles/panes and keyboard Space/Enter pass; stale focus/ticket replay,
  missing result, stale painted prompt and unchanged bytes remain covered.
  [Keyboard receipt](../../../TestRuns/4cb089111a23/Commands/2026-09-20_180849-i26-inspection-focus-keyboard/README.md): **18/18**, zero warnings/errors/skips/audit issues.
  Candidate and final each also pass 18/18. The first returned frame was visually rejected as inactive;
  actual OS-focus assertions and View-command handling repaired it. [Focused gallery](implementation/2026-09-20-inspection-focus/gallery.html)
  contains one before and three current harness frames; both return frames directly reviewed.
  Source snapshot `a9fa1560a4575edfbc869e059d25919257a0c56f08d6a17b3b883455097c8edd`. Changed code: Dialog.cpp, Popup.{h,cpp},
  State.Runtime.cpp, FileOperationsInternal.h and Commands.SelfTest.FileOps.cpp.
  Redundant native scrollbar redraw is skipped when actual range/page/position is unchanged. Renderer
  common/differing p95: candidate 66788/66824 us; final 66859/66859; keyboard 66788/66782 (45 samples
  each). Original 88361/81740 tails remain retained; component analysis is under
  `C:/RedSalamander.Perf/evidence/i26-ui/redraw-tail-components.json`. No cause, final resources or
  Release acceptance is inferred. Raw roots are `S:/RedSalamander.Perf/evidence/i26-inspection-focus-{candidate,final,keyboard}-20260920`.
  **Next:** fix retained-control body clipping and reveal offscreen keyboard focus; legacy content clips
  above the footer while root-hosted controls currently do not. Prepared scratch implementation/test
  scripts: `implement-hosted-body-viewport.py`, `test-hosted-body-viewport.py` in the same C: evidence folder.
  Qualified DxUi source is unchanged; [draft PR #18](https://github.com/RedSalamanders/DxUi/pull/18) is now published,
  superseding the unpublished-branch blocker. x64 Debug/Release/ASan and formatting CI pass; native ARM64
  CI remains running. Exact-main qualification and explicit pin update remain required. The library
  plan checklist is reconciled with existing bounded fixture/API/state evidence; unrelated new Editor
  plan/index edits belong to another owner and are preserved. Both plans remain ACTIVE. Continue without
  stopping at this checkpoint; final native mixed-DPI/AT, provider/refusal, Release/Fresh Full and resource
  gates remain open.

- **Previous checkpoint, 2026-09-20: R3 inspection controls and accessible outcomes implemented.**
  Source/destination have separate full-width French hosted buttons; the painted immutable prompt
  owns dispatch. One bounded presentation/request and weak popup lifetime guard asynchronous results.
  Current-prompt results are measured, painted and announced; replacements cannot inherit old text.
  Expanded decision heights now snap to measured content so a new result cannot overlap scope/actions
  while an old shorter card animates. Normal transfer/completed-card animations remain unchanged.
  [Focused gallery](implementation/2026-09-20-inspection-controls/gallery.html)
  retains 10 current frames, two before frames and one rejected overlap frame; four current frames
  and the rejected frame directly inspected. New labels/results are complete; old narrow action clipping
  and short-viewport footer/action overlap remain unaccepted under R2.
  [Correctness receipt](../../../TestRuns/4cb089111a23/Commands/2026-09-20_172529-i26-inspection-controls-accepted/README.md): **18/18**,
  zero warnings/errors/skips/audit issues. Real UIA invokes both roles in opposite panes; exact item,
  missing result, stale painted prompt refusal, replacement-result removal, non-activating dispatch
  and unchanged destination bytes are covered. UIA invocation itself establishes button/popup focus
  in the existing pin; asynchronous navigation does not change it afterwards. The native fixture
  reader now shares READ/WRITE/DELETE to verify bytes while the admitted object binding is retained.
  Earlier 15/18 candidate/final failures, a verified frame rejected for overlap and a layout-edit
  compiler failure are retained in the evidence; none is treated as a passing acceptance run.
  Changed code: `FolderWindow.FileOperations.Popup.{h,cpp}`, `Resource.h`, English/French resources,
  `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`. Source: `25db5fe194861741fedcbbd103fd84dadcec7c420add73775c4c623aba680550`.
  **Performance gate remains open:** baseline common/differing median/p95 66496/66865 and
  66547/66755 us; accepted 66583/66852 and 66588/88361. One justified unchanged-source
  [confirmation](../../../TestRuns/4cb089111a23/Commands/2026-09-20_173013-i26-inspection-controls-perf-confirmation/README.md) passes **3/3**, records
  66620/79436 and 66504/81740 (45 samples each). No tail-latency non-regression claim and no
  silent rebaseline. Next performance work must separate CPU/layout from presentation waits and
  retain matched baseline evidence; do not keep repeating unchanged runs without a diagnostic change.
  Raw roots: `S:/RedSalamander.Perf/evidence/i26-inspection-controls-{candidate,final,verified,accepted,perf-confirmation}-20260920`.
  **Next, Astra High:** implement the guarded foreground commit and popup hide/return. The current
  controls navigate/select and report results but leave the popup visible. Capture a post-selection
  visit/owner witness; recheck prompt, pane, foreground/input and invoking control before focus,
  and across synchronous Windows focus reentry. Keep a bounded return token for the same unresolved
  card, scope, scroll and invoking hit; do not resurrect resolved/replaced prompts. Existing popup
  `_inspectionLifetime` rejects callbacks after destruction. Cover real foreground transitions with
  the existing notification/desktop lease/restoration, both roles/panes, late input/prompt replacement,
  close/reopen and UIA; then broaden provider/refusal/rollback and long-result/narrow/native-DPI cases.
  Validation: specs have zero blocking findings; 52 new TestRuns files and the complete 5575-file
  inventory pass. DxUi skills/specs/format and clean tracked-export dependencies pass.
  Docs/proposal/gallery and helper catalog updated. No DxUi implementation/API/pin changed; library
  gallery regeneration is unnecessary. R2 publication/exact-main/pin adoption, final Release/Fresh Full,
  native mixed-DPI/keyboard/AT and paired resource gates remain open. Both plans remain ACTIVE.

- **Previous checkpoint, 2026-09-20: R3 non-activating owner navigation implemented.**
  `BeginPendingConflictInspection` captures the active pane, dispatches the qualified parent through
  canonical navigation, and consumes enumeration settlement. It rechecks original visit, live pane
  adapter/profile/root and pending prompt before exact selection. Provider/context strings and COM
  ownership are retained across provider calls, then checked again for reentrant replacement.
  One request is retained: a duplicate returns busy; explicit cancellation and newer pane navigation
  complete stale. Missing/filtered/ambiguous items, unavailable qualification, failed enumeration and
  failed rollback are distinct outcomes. A failed enumeration restores the previous qualified location
  and waits for that enumeration before reporting restoration. Caller callbacks may start a new request;
  the old request and listener are removed first. Existing external settlement listeners are preserved.
  A focus/input/foreground witness is returned, but no OS focus is taken. This witness is not permission
  for a later caller to skip prompt/foreground/control revalidation. Popup hide/return and controls remain.
  [Baseline](../../../TestRuns/4cb089111a23/Commands/2026-09-20_160358-i26-inspection-target-final/README.md),
  [candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-20_162335-i26-inspection-navigation-candidate/README.md), and
  [final](../../../TestRuns/4cb089111a23/Commands/2026-09-20_162858-i26-inspection-navigation-final/README.md) each pass **18/18**.
  Final has zero skips, compiler warnings/errors or sandbox audit issues; it adds retained live-adapter
  protection and completion-callback reentrancy coverage. Native cases invoke both pane/role combinations,
  verify exact selected identity and unchanged OS focus, duplicate/cancel, superseding navigation,
  missing item, failed-folder rollback and stale equal-path prompt refusal. No operation consent changes.
  Changed code: `FolderView.h`, `FolderWindow.cpp`, `FolderWindow.FileOperationsInternal.h`,
  `FolderWindow.FileOperations.State.Runtime.cpp`, `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`.
  Source snapshot: `f3a57044436227c2315d3b73b7b12bcfcb13799c4aeea583c7e763ed39550bd3`. New metric: `FileOps.Inspection.NavigationUs`.
  Same French renderer fixture, 45 samples each: common median/p95 **66561/66843 ->
  66496/66865 us**; differing **66564/66846 -> 66547/66755 us**.
  Navigation timings are candidate-only, not paired resource acceptance. Raw roots:
  `S:/RedSalamander.Perf/evidence/i26-inspection-navigation-{candidate,final}-20260920`.
  **Next, Astra High:** expose the French source/destination controls and accessible result text;
  connect completion to guarded popup hide/pane focus and return to the same unresolved card, scope,
  scroll and invoking control. Preserve focus when the witness changed and recheck before committing.
  Popup source map: `FolderWindow.FileOperations.Popup.h` owns hit kinds/hosted control state;
  `.Popup.cpp` owns `BuildConflictPromptContent`, `MeasureConflictPromptContentHeight`, hosted text/
  accessibility and `OnActivatedHit`. Update measured height, drawn rows, hosted hit bounds, keyboard
  and UIA together. The owner helper is not itself a UI action or a return-to-control implementation.
  Extend actual harness interaction/captures for left/right, late prompt/user-focus changes, provider
  switching/refusal, filters/ambiguity, rollback failure, close/reopen and UIA. Refused/provider-changed
  navigation and remote rollback are not fully qualified by the current local enumeration-failure case.
  R3 stays open; buttons are not exposed yet. Docs/helper catalog reviewed/updated; no rendered control
  or library change, so gallery regeneration is unnecessary. R2 publication/exact-main/pin adoption and
  final Release/Fresh Full/native mixed-DPI/UIA/AT/resource gates remain open; both plans remain ACTIVE.

- **Previous checkpoint, 2026-09-20: R3 owner target qualification implemented and verified.**
  `ResolvePendingConflictInspectionTarget` retains the requested role's task adapter, plugin identity,
  raw pane instance context and immutable prompt. It does not infer authority from the active pane.
  The canonical qualifier reuses `CanonicalInstanceId`, `SamePathIdentityContract` and provider path
  helpers; typed queries check both exact item and parent against admitted provider/profile/root/
  comparison rules. It refuses changed or missing qualification, inconsistent leaf/parent and stale
  prompts. Failures clear output. Prompt freshness is rechecked after provider calls; no borrowed
  task spans those calls; a weak UI-lifetime guard prevents owner access after reentrant shutdown.
  No navigation, focus, consent, mutation, timer or worker is added here.
  Retained [pre-change baseline](../../../TestRuns/4cb089111a23/Commands/2026-09-20_152349-i26-inspection-settlement-confirmation/README.md)
  passes 18/18. [Initial candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-20_155849-i26-inspection-target-candidate/README.md)
  is retained at **17/18**: one select-next navigation-shell baseline capture failed (cause unproven);
  all three new target/selection cases passed. The later owner-lifetime guard and aliasing assertion
  are included in the [final matrix](../../../TestRuns/4cb089111a23/Commands/2026-09-20_160358-i26-inspection-target-final/README.md), which passes **18/18**,
  zero skips, warnings, errors or sandbox audit issues. New native checks cover both role adapters,
  output/context aliasing, raw default vs canonical instance context, changed profile/root/path comparison/separators/instance,
  another provider, inconsistent leaf, unknown path rules, absent provider, and stale equal-path prompt
  clearing a previously valid target. Existing exact selection, callback reentry, unchanged prompt and
  destination bytes remain covered. Remote disconnection/provider reentry and full interaction remain open.
  Source snapshot: `087b461222f74b64ddf7dcd2b14f63fae28597d487b597db650fb69c1420e6e2`. Changed code: `FolderWindow.FileOperations.cpp`,
  `FolderWindow.FileOperations.State.Runtime.cpp`, `FolderWindow.FileOperationsInternal.h`,
  `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`. Metric: `FileOps.Inspection.QualifyEndpointUs`.
  Same renderer fixture, 45 samples: common median/p95 **66539/66835 -> 66561/66843 us**;
  differing **66548/66834 -> 66564/66846 us**. New qualifier cost is candidate-only,
  not a paired latency/resource claim. Local qualifier: 39 calls, median/p95/max **398/1046/1089 us**.
  This does not qualify remote-provider responsiveness. Raw: `S:/RedSalamander.Perf/evidence/i26-inspection-target-{candidate,final}-20260920`.
  **Next, Astra High:** connect the owner request/completion transaction to the qualified target and
  pane settlement hook, preserving the captured last-active pane, navigation rollback and intervening
  interaction/foreground guards. Requalify role and prompt at completion before committing selection.
  Do not use `ExecuteInPaneLocation` as a passive dispatch: its `ExecuteInPane` path focuses the
  view immediately and queues permissive display-name selection. Reuse qualified `SetFolderPath`
  navigation and retain exact adapter/rollback ownership; commit exact selection/focus only at settlement.
  Then expose French source/destination controls and restore unresolved card/scope/scroll/invoking
  control through the File Operations command. These interactions are not implemented by this helper;
  **R3 remains open and no inspection buttons are exposed yet**. Docs/helper catalog updated;
  docs/gallery reviewed, no visual change requiring new images. No DxUi source/API/pin changes.
  R2 publication approval/exact-main/adoption and final Release/Fresh Full/native mixed-DPI/UIA/AT/
  resource gates remain open. Both plans and the existing continuation stay ACTIVE.

- **Previous checkpoint, 2026-09-20: R3 enumeration settlement hook verified.**
  `FolderWindow::SetPaneEnumerationSettledCallback` forwards a UI-thread receipt with a fresh pane
  visit, attempted folder and HRESULT after accepted enumeration and queued work. Both successful
  and failed results notify; superseded/reentrant navigation suppresses the old result. Callbacks
  can unsubscribe themselves safely. Uninterested panes incur no token or callback allocation;
  no worker, timer or polling was added. The attempted folder is retained before early failure returns.
  The lazy inspection marker now uses an exact `uint8_t` initializer, removing the previous C4242
  narrowing warnings. The earlier checkpoint's zero-warning claim is corrected above its old receipt.
  Retained baseline: previous [18/18 selection matrix](../../../TestRuns/4cb089111a23/Commands/2026-09-20_144559-i26-inspection-selection-final/README.md).
  [First attempt](../../../TestRuns/4cb089111a23/Commands/2026-09-20_150354-i26-inspection-settlement-candidate/README.md) and
  [second attempt](../../../TestRuns/4cb089111a23/Commands/2026-09-20_150917-i26-inspection-settlement-final/README.md) each fail **15/18**:
  the first checked a moved-from success folder; the second exposed absent folders on early failures.
  Both are retained and fixed. [Verified matrix](../../../TestRuns/4cb089111a23/Commands/2026-09-20_151853-i26-inspection-settlement-verified/README.md)
  passes **18/18**, zero skips, warnings, errors or sandbox audit issues. Coverage includes exact item
  selection inside success notification, exact failed status, self-unsubscription, later refresh and
  obsolete-result suppression when legacy completion navigates. Legacy completion stays success-only.
  Source snapshot: `7696814991fa75f55c691a38bb1c7f51f7cb62504c3e3a1d7c9c2554680a2546`.
  Changed code: `FolderView.{h,Enumeration.cpp,Selection.cpp}`, `FolderWindow.{h,cpp}` and
  `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`. Callback metric: `folder.enumeration.settled_callback_us`.
  Paired unchanged renderer fixture, 45 samples each: common median/p95 **66584/66851 ->
  66601/719862 us**; differing **66590/66863 -> 66594/66799 us**.
  The common p95 spike consists of four consecutive second-repetition samples; no cause is proven.
  One [unchanged-C++ confirmation](../../../TestRuns/4cb089111a23/Commands/2026-09-20_152349-i26-inspection-settlement-confirmation/README.md) passes **18/18**,
  zero warnings/errors/skips/audit issues. All six code-file hashes match; only docs changed between
  workspace receipts. Confirmation common median/p95 **66539/66835 us**, differing **66548/66834 us**,
  back near baseline. Retain the spike; no silent rebaseline or claim about its cause. Confirmation
  snapshot: `da4af5cd36c38726a99fd8752c7b6d7986cbc2c6ee873b1cb462735711178efe`.
  These are Debug renderer witnesses only; new callback costs have no matched baseline/resource claim.
  Raw roots: `S:/RedSalamander.Perf/evidence/i26-inspection-settlement-{candidate,final,verified,confirmation}-20260920`.
  **Next, Astra High:** implement the FileOperationState owner transaction and visible French inspection
  controls using this hook. Capture last-active pane; revalidate task/prompt and live role provider,
  profile/root/path identity and raw instance context; navigate qualified parent and select exact item;
  guard intervening interaction/foreground changes, handle failure/rollback, and restore unresolved
  card/scope/scroll/invoking control. Canonical `host/default` / `host/context/...` identity is not a raw
  pane context; reuse the existing canonical adapter helper. No inspection buttons are exposed yet;
  **R3 remains incomplete**. Domain specs/helper catalog updated; docs/gallery reviewed, no new pixels.
  No library source/API/pin change. R2 publication approval, exact-main CI and explicit adoption plus
  Release/Fresh Full/native mixed-DPI/UIA/AT/resource acceptance remain open. Both plans stay ACTIVE.

- **Previous checkpoint, 2026-09-20: R3 pane selection component implemented and verified.**
  `FolderView` now captures an owning-pane lifetime/visit token and commits inspection selection only
  for that path, provider epoch, enumeration and current-item ownership. Pending and failed enumeration
  have distinct outcomes; the failed outcome retains its HRESULT rather than selecting the old displayed
  model. Names/parents use the admitted provider's canonical path helpers, with one unique matching
  displayed item; no unconditional casing fallback, neighbor selection, file launch or OS-focus change.
  The normal single-item selection path now stops after a reentrant owner callback replaces the model.
  A lifetime marker is allocated lazily once per view, not per enumeration/paint; existing generations
  supersede tokens, and Destroy invalidates the marker. No new timer, worker or provider query is added.
  [Baseline](../../../TestRuns/4cb089111a23/Commands/2026-09-20_142554-i26-inspection-selection-baseline/README.md) passes **12/12**;
  [initial candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-20_143448-i26-inspection-selection-candidate/README.md) is retained at **9/12**.
  The new pre-completion byte assertion used ifstream, whose sharing mode conflicts with the unresolved
  destination's DELETE-capable binding. Its bounded native replacement shares read/write/delete and
  verifies exact length/bytes while the same prompt is still pending. No engine authority was relaxed.
  [Final six-case matrix](../../../TestRuns/4cb089111a23/Commands/2026-09-20_144559-i26-inspection-selection-final/README.md) passes **18/18**, zero skips, errors or audit issues. Correction: its runner
  contains 27 warning/context lines from two C4242 template sites, fixed in the next settlement slice. It includes request-before-enumeration, exact selection/current,
  wrong parent/pane, unstable identity, case-sensitive refusal/local case-insensitive match, missing
  items, refresh and identical-path revisit invalidation, failed enumeration, reentrant navigation,
  unchanged pending prompt/bytes, ordinary select-next and hide/show navigation-shell stability.
  Source: `dbd6a3cab69b25e60d4605fcb9e1509ab61163ca7b830e7a47d1eafe5e18fc85`. Changed code: `FolderView.{h,cpp,Selection.cpp,Enumeration.cpp}`
  and `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`. Runtime instrumentation: `folder.inspection.select_us`.
  Paired unchanged renderer fixture, 45 samples per presentation: common median/p95 **66541/66934 ->
  66584/66851 us**, differing **66497/67102 -> 66590/66863 us**. Median deltas are below 0.15%; p95
  is lower. Added inspection work is candidate-only and does not establish matched latency/resource
  qualification. Raw roots: `S:/RedSalamander.Perf/evidence/i26-inspection-selection-{baseline,candidate,final}-20260920`.
  **Next, Astra High:** wire an event-driven owning `FolderWindow` completion/failure transaction and
  the source/destination controls. Use admitted role facts to revalidate the live provider/profile/root
  and canonical-vs-raw instance context; capture the last-active pane and guard intervening user
  interaction/foreground changes. Existing enumeration callback reports success only; add a governed
  terminal failure path rather than polling. `SelectItemForInspection` checks pane/model freshness,
  not operation consent, live remote-provider qualification or global focus ownership. Preserve
  unresolved scope/card/scroll/invoking control and restore it through the ordinary File Operations
  command. Native filtered/ambiguous/remote/lifetime-recreation and complete left/right interaction
  coverage remain open. **R3 is not complete and no inspection button is exposed yet.**
  No DxUi source/API/pin or visual change; docs/gallery were reviewed and need no new pixels for this
  internal component. R2 still awaits approved publication, exact-main CI and explicit pin adoption;
  all parent Release/Full/native-DPI/AT/resource gates remain open. Both plans/continuation stay ACTIVE.
  Validation: all three explicit archives and the full 5,443-file archive inventory pass. Specification
  inventory has zero blocking findings. DxUi skills/specs/format pass; dependency validation passes
  in the clean tracked export, preserving unrelated ignored developer settings.

- **Previous checkpoint, 2026-09-20: R3 immutable inspection facts and freshness guard implemented.**
  Six consumer source files now carry per-role admitted endpoint/provider paths in one shared immutable
  prompt context. Its retained identity distinguishes consecutive prompts even with equal paths;
  metadata decoration and popup hiding preserve it. UI-thread freshness checks re-resolve the task ID
  and reject null, unrelated, cancelled, stopped, finished or replaced contexts. No provider call,
  navigation, focus change, decision or mutation is introduced by this preparatory slice.
  [Baseline](../../../TestRuns/4cb089111a23/Commands/2026-09-20_135420-i26-inspection-context-baseline/README.md),
  [candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-20_140350-i26-inspection-context-candidate/README.md) and
  [unchanged-source confirmation](../../../TestRuns/4cb089111a23/Commands/2026-09-20_140852-i26-inspection-context-confirmation/README.md)
  each pass **12/12**, zero skips, build warnings/errors or audit issues. Tests cover flattened/nested
  paths, provider identity/separators/case policy, source-only/delete qualification, exact-path fallback,
  immutable generation, blocked metadata, hide/reopen and cancellation. Equal-content token rejection
  is covered; actual asynchronous pane navigation and producer-replacement journeys remain open.
  Candidate source: `fe2e4681f0372bf3e8532c5ebb74edf4f52f01f1f71349acbbfa173daf0a8fe4`.
  Common-name redraw median/p95 baseline 66580/66788 us, candidate 66581/66883 us; differing-name
  66569/66764 us versus 66594/84154 us. The latter includes large outliers; the justified unchanged-source
  repeat gives 66559/66845 and 66593/66933 us respectively. All 45 ordered samples per presentation/run
  are retained; no silent rebaseline or final Release/resource claim. SkipBuild was rejected before tests
  for an incompatible targeted receipt, then a normal attested build ran. No attestation bypass.
  Raw roots: `S:/RedSalamander.Perf/evidence/i26-inspection-context-{baseline,candidate,confirmation}-20260920`.
  Validation: three explicit pre-stage archive checks and the full 5,413-file inventory pass;
  specification inventory has zero blocking findings. DxUi skills/specs/format pass, and dependency
  validation passes in the clean tracked export while preserving ignored developer settings.
  Changed code: `FolderWindow.FileOperations{.cpp,Internal.h,.Popup.cpp,.Popup.h,.State.Runtime.cpp}`
  and `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`. No shared-library API/pin or visual change;
  harness captures exercise the existing French renderer, so the published gallery needs no new pixels.
  **Next dependency-ready slice, Astra High:** implement the consumer inspection completion transaction
  using these facts. Capture last-active pane, validate live endpoint/context (canonical instance ID is
  not necessarily the pane adapter's raw context), observe enumeration/nav generation and exact
  provider-correct selection, handle failure/rollback and intervening focus, then wire full French
  actions and return-to-control. `ExecuteInPaneLocation` acceptance is not completion;
  `PrepareForExternalCommand` case-insensitive fallback is not qualified exact selection. R3 remains
  unchecked. R2 still awaits approved publication, successful exact-main CI and explicit pin adoption.
  No publication answer has arrived. Both plans and recurring continuation remain ACTIVE; all parent
  Release/Full/native mixed-DPI/AT/resource gates and other owners' deferred work remain open.

- **Previous checkpoint, 2026-09-20: native clipboard persistence blocker repaired and verified.**
  The user copied text and authorized directed testing after Windows denied private window-station
  creation (error 5, before clipboard access). A fresh [baseline](../../../TestRuns/4cb089111a23/Commands/2026-09-20_131249-i26-clipboard-ole-baseline/README.md) is **0/3**;
  the [diagnostic](../../../TestRuns/4cb089111a23/Commands/2026-09-20_132105-i26-clipboard-ole-diagnostic/README.md) reproduces `CLIPBRD_E_BAD_DATA` at the first OLE
  flush/readback with a 12-byte stream at position zero. The owned-only passing witness had not
  covered that boundary. A private binary-format probe also fails, so the issue is not PNG-specific.
  The test-only snapshot now materializes stream bytes into native owned HGLOBAL, preserves the
  source cursor, rejects unsupported descriptors/media or oversized data before mutation, and
  verifies every retained format byte-for-byte after restoration. Boundaries: 256 entries, 64 MiB
  total, unindexed content, no target device, HGLOBAL/IStream input. No format is silently dropped;
  arbitrary provider behavior and unsupported media are not claimed. Clipboard contents are never logged.
  [Focused candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-20_133008-i26-clipboard-ole-candidate/README.md) passes **6/6** and the original twelve-case
  [broader matrix](../../../TestRuns/4cb089111a23/Commands/2026-09-20_133432-i26-clipboard-ole-final/README.md) now passes **36/36**, zero skips, build warnings/errors
  or disk-audit issues. Each directed run records all **7/7** saved clipboard formats restored exactly.
  Corrupted-byte rejection, stream cursor preservation, three owned/OLE cycles, French path-copy/UIA,
  conflict actions/close, completed navigation, failure surfaces and popup lifetime pass. This closes
  the reproduced stream-persistence blocker; the historical **34/36** and new **0/3, 0/1** failures remain.
  Source changed only in `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`; no product
  renderer or DxUi pin changes. The broader matrix uses source snapshot
  `b959bedb4fe24e477d7eb9030e3c5ff12588a51dd8fb14b53f4dd5e9327c6ce2`.
  A final [preflight guard review](../../../TestRuns/4cb089111a23/Commands/2026-09-20_134243-i26-clipboard-ole-preflight/README.md) passes **6/6**, zero skips/warnings/errors:
  verify every saved value in the Shell snapshot before replacement (including coalesced format keys),
  and report zero fully verified formats when restoration fails. The 36-case matrix predates this
  test-only guard; the six affected cases were rerun on final source `552a4d9939e4bd70987c8170ad173123b31a6eb193a799db36532f0a0f2f7cca`.
  Raw evidence: `S:/RedSalamander.Perf/evidence/i26-clipboard-png-ole-{baseline,diagnostic,candidate,final,preflight}-20260920`.
  Validation: all five curated archives pass explicit pre-stage checks; the full archive inventory
  passes **5,383 files**. Specification inventory has zero blocking findings. DxUi skill/spec/format
  checks pass; dependency validation passes in the existing clean tracked export with the updated
  coordination plan, preserving the unrelated ignored developer project settings.
  **Next dependency-ready slice, Astra High: R3 pending source/destination inspection and return.**
  `ExecuteInPaneLocation` returns request acceptance before enumeration/selection completes;
  `PrepareForExternalCommand` can fall back to case-insensitive matching. Add a qualified completion
  transaction with prompt/navigation generation and provider-correct exact-item checks; do not reuse
  those success/fallback semantics as proof. Cover refusal, missing/filtered items and stale focus,
  preserve decision scope and return to the invoking control. R1 one-item/differing-name polish,
  R2 measured actions and final Release/Full/mixed-DPI/AT/resource gates remain open. R2 still requires
  approved DxUi publication, successful exact-main CI and explicit adoption; no publication answer
  has arrived. Both WIP plans and the existing recurring continuation remain ACTIVE.

- **Previous checkpoint, 2026-09-20: R1 common-name/status rendering implemented; qualification remains open.**
  The popup now displays one full-width bold common filename, an unnamed problem explanation and
  stacked parent folders/metadata when admitted leaf values match exactly. Expanded headers identify
  the operation; the duplicate fixed-width conflict chip is removed and one measured waiting/loading
  line remains. Collapsed cards retain a status-bearing title. UIA uses one `.Conflict.Name` provider
  and no hidden filename duplicates; announcements retain the common name. Different leaves and
  unknown/slash-provider paths retain their precise spelling/fallback. No conflict authority/default,
  library code or dependency pin changes are included. Gallery seeds now supply explicit local facts.
  Changed source: `FolderWindow.FileOperations.Popup.cpp`, `Commands.SelfTest.FileOps.cpp`, and
  `FolderWindow.FileOperations.SelfTest.VisualGallery.cpp` under their existing source owners.
  [Before/after evidence and regenerated gallery](implementation/2026-09-20-common-name/README.md): baseline **0/3**, candidate **3/3**;
  the expanded twelve-case final run is **34/36**, zero skips. Common-name UIA count/order/metadata
  pixels, provider fallback, real conflict actions/close, contrast/failure surfaces, completed navigation
  and popup resources pass. **The final run is not a pass:** throughput repetitions 2/3 cannot retain
  advertised clipboard format **49343 (`PNG`)**, `CLIPBRD_E_BAD_DATA (0x800401D3)`, before mutation.
  First repetition passes; original failure is retained. Exact cause and complete clipboard restoration
  are unproven. Do not clear the user's clipboard or repeat the unchanged failing test.
  The independent static gallery passes **1/1** at native **144/96 DPI**, zero skips, warnings/errors or
  audit issues. All **300** regenerated catalog PNGs plus eleven focused before/after images are in WIP;
  **15/311** are directly inspected. Long common filenames/parent folders fit on both displays; narrow
  actions still clip. The other 296 catalog frames are captured/indexed, not visually accepted.
  Paired Debug median eight-redraw batches: common names **66,545 -> 66,573 us (+0.042%)**;
  different names **66,568 -> 66,590 us (+0.033%)**, 45 samples per variant before/after. These are
  bounded fixture timings, not final Release FPS or paired resource acceptance. Final native source
  snapshot `bcf4e812cf450181567c43f4c816e60844da8a0597fa0e90f8b58c17e98494a6`.
  Full raw runs are retained at `S:/RedSalamander.Perf/evidence/i26-common-name-{baseline,candidate,final,gallery}-20260920`;
  compact archives and per-image source/build/hash provenance are linked from the evidence README.
  Follow-up [owned PNG stream characterization](../../../TestRuns/4cb089111a23/Commands/2026-09-20_124438-i26-clipboard-png-characterization/README.md)
  passes **3/3** without accessing the user clipboard: exact opaque PNG bytes survive invalidation of
  the original stream backing and repeated snapshots left at EOF. The existing Shell retainer already
  satisfies that boundary. This does **not** reproduce or fix the actual OLE flush/readback failure.
  No generic duplication workaround is applied; the scratch draft is unqualified. Next reproduce the
  OLE persistence boundary with a private, fixture-owned clipboard (no desktop activation), including
  carrier/format diagnostics and exact byte roundtrips, before another directed user-clipboard run.
  **Next dependency-ready slice, Astra High:** diagnose PNG snapshot independence with owned synthetic
  clipboard data; characterize returned carrier types/lazy backing and flush/readback before further
  directed user-clipboard mutation. Require exact bytes/all-format restoration, without treating a
  successful flush as proof. Keep the original 34/36 receipt. Then resume R3 pending inspection/return
  and remaining R1 one-item/differing-name polish. R2 measured actions/collapsed headers/short viewport
  still follows approved publication, successful exact-main CI and explicit qualified DxUi adoption.
  No publication/merge answer has arrived. Release/Full, final mixed-DPI input/AT and paired-resource
  gates, both WIP plans, and the recurring continuation remain ACTIVE.

- **Previous checkpoint, 2026-09-20: user interaction correction and audit.** The user rejected clipped
  actions/badge, three waiting messages, a quoted/buried filename and its repeated appearances, and
  requested inspection of the conflicting destination in the focused pane. The normative contract
  and proposal now require one common filename heading, one conflict indicator, stacked parent
  folders with metadata, measured complete labels and pending source/destination inspection/return.
  This supersedes the per-role repeated-name presentation below; retain its provider qualification.
  [Review and M01–M32 interaction inventory](interaction-review-2026-09-20/README.md)
  accounts for all 26 non-None popup action kinds, menus, adjacent dialogs, input, failure and return.
  It distinguishes source-existing tests, observed display states and missing executed outcomes.
  Fresh static and notified input cases pass **1/1 each**, zero skips, warnings/errors and audit issues;
  **309 French harness PNGs** report native **144/96 DPI**, including same-HWND return to the first
  display. **48 frames are directly reviewed**; the remaining catalog is indexed, not visually accepted.
  Workspace snapshot `6174d1ab2428470c9d3992b649ce3c20aac43490d573d34a790733f071937e67` binds both
  [static](../../../TestRuns/4cb089111a23/FileOps/2026-09-20_113651-i26-interaction-review-static/README.md) and
  [input](../../../TestRuns/4cb089111a23/FileOps/2026-09-20_114123-i26-interaction-review-input/README.md) receipts.
  The broad conflict gallery uses synthetic unqualified snapshots and therefore renders exact-path
  fallback; it must gain admitted production-backed fixtures before qualifying the new heading.
  Clipped action/badge/Clear/destination controls, short-display footer overlap and Issues message
  clipping remain observed failures. The checked glyph and speed-dialog spacing are visibly improved.
  [HTML proposal](mockups/07-conflict-single-name-french.html) has six
  checked browser layouts and simulated inspect/return; it is explicitly not native UI acceptance.
  Product/library source, APIs, pin and library gallery pixels are unchanged in this audit.
  Review closeout: all 309 raw/copy PNG hashes and dimensions, seven source hashes and the 26-action
  inventory match; local links pass. Both curated archives pass pre-stage validation, specification
  inventory has zero blocking findings, and DxUi skills/specs/format plus clean-export dependency
  validation pass. Two proposal previews are visually reviewed after six browser layout checks.
  **Review handoff (superseded by the current checkpoint): R1 single-name/status hierarchy.** Extend the real
  conflict fixture first, retain failing visible/UIA count and reading-order witnesses, then reuse
  admitted facts for the prominent common name and folder-only roles. Preserve differing leaves,
  unqualified fallback, metadata, policy/defaults and exact-path detail. Revise the old `.Name`-per-role
  expectations rather than satisfying them with hidden duplicates. Add qualified catalog coverage.
  Then R2 measured actions/header/short viewport follows explicit qualified pin adoption, and R3
  pending inspect/return uses the existing pane transaction with generation/focus checks. The prior
  publication/merge request remains unanswered; continue R1/R3 and other independent work, without
  copying shared helpers or bypassing exact-main CI. Existing library qualification is retained.
  Final Full/Release/native redesigned-UI mixed-DPI/AT/paired-resource gates and both plans remain open.

- **Latest completed slice (2026-09-20): qualified conflict filename/location rows.** Transfer/delete
  item prompts compute display parts using the selected child plan's own endpoint profiles. The UI
  places each name above its full-width parent, Incoming above Existing, then metadata. Questions
  reuse qualified names; unavailable profiles/indexes/roots and rename schedules preserve exact paths.
  No provider I/O in rendering or change to action authority, defaults, engine policy or DxUi pin.
  Baseline **0/3** is retained; first candidate **6/6**, final eleven-case Debug set **33/33**, zero skips.
  [Evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-20_111505-i26-conflict-paths-final/README.md) binds workspace
  `0ee181f1960de7bfa151b8fea81cf1125fcc25c29e1d847cf95d86d1ee7860b4` and receipt
  `3ff178561c0bf9fea5500ff5435b176e7e8a77f82b601451d6da561474641b94`. The pure multi-provider case and real published task plus French
  full-width semantic rows cover the new adapter. [Eight captures](implementation/2026-09-20-conflict-paths/README.md)
  retain the narrow French action/badge clipping and exact-path fallback. Timing samples are paired
  but changed painted content/Debug cannot close final FPS/resource gates.
  Changed source: `FolderWindow.FileOperations.cpp`, `.State.cpp`, `.Popup.cpp`, `.Popup.h`,
  `FolderWindow.FileOperationsInternal.h`, `Commands.SelfTest.FileOps.cpp`; normative UI/performance/
  helper contracts, proposal/gallery and both plans are updated. Library source/gallery unchanged.
  Closeout: changed-line clang-format 22.1.3 and whitespace checks pass; specification inventory
  reports zero blocking findings; all **5,280** archived TestRuns files pass validation. Eight
  reviewed PNG hashes match their original harness outputs. No build or test is active.
  **Next, Astra High:** honor any actual publication/merge answer, adopt qualified exact-main CI via
  the standard updater, then use shared measured action flow. While permission remains pending,
  continue independent short-viewport/path-detail or typed Issues work. Rename/task-wide display
  qualification, exact-path selection/copy, action/badge fit, Full/Release/native mixed-DPI/AT and
  paired resources remain open. No unchanged clipboard or library qualification retry is needed.

- **Latest completed prerequisite (2026-09-20): clipboard snapshot medium contract.** The harness
  now stores the actual returned medium type instead of advertising an enumerator's multi-medium
  mask. An owned-object probe reproduced stream-only reads returning HGLOBAL; the normalized
  variant preserves bytes and rejects unsupported streams. SetData(FALSE) is unsupported and
  was rejected; borrowed-byte aliasing was not demonstrated by this probe.
  The new owned native case validates bytes/type after source release across three snapshots.
  The existing directed fixture also runs three rich-format OLE flush/readback cycles inside its
  complete clipboard-restoration scope, then exercises French/Unicode Ctrl+C/focus/navigation.
  Focused checks pass **6/6**; the ten-case Debug regression set passes **30/30**, zero skips,
  compiler warnings/errors and audit issues. [Final evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-20_105016-i26-clipboard-medium-final/README.md)
  binds workspace `002c2fd168c6fca57d78150b0eef8f131962d1a3f116948b9d071af8f009c005` and receipt
  `e3d70ac1228ababf27363870de4c072ff169efadad06a47ebd6bf64329acbd8d`. The original 25/27 failure is retained. Early native 0/3 and 0/6 attempts
  fail an incorrect single-enumeration-entry assertion; that test defect and the initial C5246
  brace warning are preserved, not claimed as target-regression acceptance.
  Changed source: `Commands.SelfTest.FileOps.cpp` only; the format snapshot is reused by the
  directed fixture and new test. Testing performance/shared-helper specs and both plans are updated.
  No product rendering, engine policy, DxUi source/API/pin or gallery pixels changed in this repair.
  Closeout: pinned formatting, specification inventory, DxUi guidance checks and the complete
  **5,249-file** TestRuns inventory pass. Source/docs whitespace checks pass; four raw diagnostic
  logs intentionally retain their terminal blank lines for byte/hash provenance. No test/build is active.
  **Next, Astra High:** honor any actual answer to the qualified DxUi publication/merge request;
  adoption still requires successful exact-main CI through the standard updater. If pending,
  continue provider-qualified filename/parent display facts or independent Issues detail work.
  The prior decision-only body now has a passing broader regression receipt; remaining French
  action wrapping, short scrolling, full Issues and final Full/Release/native DPI/AT/resource gates
  are still open. Do not repeat resolved clipboard or standalone-library qualification unchanged.

- **Implemented slice (2026-09-20), broad validation still open: decision-only conflict body.** Ordinary transfer paths,
  mini-progress and concurrency text no longer precede a pending conflict. Card height is measured
  from the decision body and published action/scope rows. No engine policy or DxUi pin/API changed.
  The same French fixture retains a three-repetition failing old-renderer baseline, a **3/3**
  focused candidate and **25/27** final nine-case Debug set. Compiler/audit are clean, zero skips. The broad failures are the existing clipboard snapshot gate in throughput repetitions 2/3, not passes.
  Actual hosted-tree absence assertions supplement existing metadata UIA/pixel checks; the former
  policy-only progress flag missed this visible defect. [Evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-20_095955-i26-conflict-clean-final/README.md)
  binds workspace `6116ab197a55520e24566dd0a1759dc146a03daf4211e2d23f40a7affa87c2dc` and receipt
  `0c4eff1319fd6f97e8aae725936a37a2c0e2758027514d5561b8114e9c97db0d`. [Six reviewed captures](implementation/2026-09-20-conflict-decision/README.md)
  preserve the remaining narrow action/badge clipping. All timing repetitions and input hashes
  remain in the comparison; less painted content is not equal-work or final performance proof.
  Source changes: `FolderWindow.FileOperations.Popup.cpp`, `Commands.SelfTest.FileOps.cpp`;
  normative UI/performance specs, proposal/gallery and both checkpoints are updated.
  **Next, Astra High:** diagnose and repair clipboard snapshot independence before rerunning broad
  qualification. The first throughput repetition passes; repeats 2/3 cannot read restored HTML
  Format, registered ID 49344, `0x800401D3`. The format name was queried without reading/logging
  its contents. Preserve all formats and the original failure; do not clear the user clipboard or
  weaken assertions. An unproven ownership hypothesis is that `GetData` returns a medium still
  dependent on the old clipboard, while `SetData(TRUE)` transfers it without an independent copy.
  Investigate with owned synthetic rich-format data before touching the user clipboard again.
  [Microsoft's SetData contract](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-idataobject-setdata)
  distinguishes transferring the medium from caller-retained ownership (`FALSE`, use only during
  the call); do not assume a fix without verifying independent retained bytes and restore/readback.
  Then honor any answer to the pending qualified-library publication/merge request,
  then adopt successful exact-main CI through the standard updater for measured action wrapping.
  While that remains pending, implement provider-qualified filename/parent display facts or the
  independent Issues detail work. Do not duplicate the shared action-flow implementation locally.
  Full conflict/Issues design, long detail/short scrolling, final Full/Release/native DPI/AT/resources
  remain required. Both plans and the existing continuation stay ACTIVE.
  Closeout checks pass: pinned formatting, zero blocking spec findings, three explicit archives
  (31 files), whole TestRuns inventory (5,203 files), and DxUi skills/specs/format. These checks
  validate the records; they do not turn the two native failures into passes. No build/test lane
  remains active. Raw logs and the saved identical comparison fixture remain in `i26-ui` evidence.

- **Resume here after `b522a2bf` (consumer), `a5712b0` (library coordination).** Both checkouts are
  committed; pinned formatting and whitespace pass, spec inventory has zero blocking findings,
  four new TestRuns archives pass (41 files), and the full archive inventory passes (5,172 files).
  DxUi skills/specs/format checks pass; this consumer slice changes no shared renderer pixels.
  Read-only GitHub checks on 2026-09-20 still return main `b125edbf4cdd639d1cabd686f9dc3fe8aab1f374`
  and no PR for `codex/localized-adaptive-layout`. The concrete PR description is prepared at
  `C:/RedSalamander.Perf/evidence/i26-ui/dxui-layout-pr-body.md`; publication/merge authorization
  was requested asynchronously. Keep that request pending; elapsed time is not approval.
  If approved, publish the qualified branch, retain required checks/reviews, merge only within
  the authorized scope, then wait for successful exact-main CI before the standard pin updater.
  While pending, the next independent Astra High slice removes stale ordinary transfer paths,
  mini-progress and concurrency text above an active conflict. Reuse the French conflict fixture
  and retain a pre-change baseline; assert actual visible/UIA output rather than the existing
  policy-only `conflictTransferProgressVisible` flag. The ordinary rendering block around
  `autoConcurrencyUsed` through the Source/Destination rows currently runs during conflicts;
  measurement around `expandedTransferCardH` also reserves ordinary transfer content. Keep the
  complete stacked conflict context, engine action policy and loading-state contract intact.
  No build/test process remains active. One next continuation may run; do not start competing writers.

- **Latest completed slice (2026-09-20): full-width conflict metadata.** Consumer renderer now
  measures and paints complete metadata below exact incoming/existing paths, with independent
  stable UIA roles and normal readable foreground. No library source/API/pin changed.
  The identical French fixture runs 480/640/760 DIP at native 144 DPI. Both old-renderer baselines
  fail the new row-semantic requirement in all three repetitions; the initial metric-conversion
  warning is fixed before the clean baseline. Focused candidate passes 3/3. The final nine-case
  set passes **27/27**, zero skips; compiler and runner audit are clean.
  [Evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-20_094101-i26-conflict-rows-final/README.md) binds workspace
  `33c39494ca779d1e12176663d1d6243e9c62e27cb0149887d11c36a0158c3d3a` and receipt
  `5b4d64023d65d3ede0e1a5d92c223f20389e7de0760c2c1dc924daac65a3129b`. Its comparison retains all four runs and raw timing samples.
  [Six before/after harness captures](implementation/2026-09-20-conflict-metadata/README.md)
  are visually reviewed. Clean-baseline/focused-candidate eight-redraw medians remain about 66.5 ms;
  this is bounded Debug evidence, not final Release/FPS/resources. Final production bytes restore
  the original UTF-8 BOM after the focused candidate; fixture SHA is unchanged.
  Changed code: `FolderWindow.FileOperations.Popup.cpp` and `Commands.SelfTest.FileOps.cpp`;
  normative UI/performance specs, proposal and gallery track the bounded result.
  **Next, Astra High:** resolve qualified DxUi publication/main-CI/pin adoption for measured action
  flow, then repair clipped French conflict actions. Preserve existing policy order/default/Escape.
  Separate provider-qualified filename/parent rows, redundant progress removal, short scrolling
  body/fixed decisions, long selectable detail, typed Issues presentation and final Full/Release/
  native DPI/AT/resource qualification remain open. Existing library full-scene/matrix/clean-checkout
  evidence remains valid for its recorded inputs; do not repeat it unchanged. Both plans stay ACTIVE.

- **Completed execution lane (2026-09-20): isolated Release qualification.** Keyboard-path code/evidence
  is committed as `b859cbc37199a566aec9a62ee7cb2df2121f1114`; DxUi guidance is `e7e697fe`.
  A clean detached checkout at `C:/RsI26Release` runs the same eight named Commands cases as the
  final Debug run, three repetitions, Release, timeout multiplier 3, run ID `i26-path-release-short`
  on `S:/RedSalamander.Perf`. Log: `C:/RedSalamander.Perf/evidence/i26-ui/path-release-short.log`.
  The run passes **24/24**, zero skips and compiler warnings/errors. Its original sandbox audit
  flags only the inactive failed setup directory; preserving that run under evidence produces a
  separate clean post-archive audit. Original results are retained unchanged.
  [Release evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-20_003057-i26-path-keyboard-release/README.md)
  binds workspace `1e0a528ba61afea93fd457a6793ee7612f32daba5311d54e17d289e25f4b6095` and receipt
  `2657022283160eb730f770f075b61770754307d0df54a871c442b2af12422a57`.
  [Four reviewed French captures](implementation/2026-09-20-release/README.md)
  cover focused Unicode paths, long graph labels and expanded/collapsed throughput at native 144 DPI.
  Full, paired Release performance, physical mixed-DPI and AT remain separate open gates.
  The initial long checkout at `C:/Users/eric/.codex/worktrees/fileops-ui-release/RedSalamander`
  failed pinned-source restoration because four historical DxUi measurement filenames exceeded
  Git's path limit. Its log is retained as `path-release-isolated.log`; no tests ran, and that failed
  checkout is preserved inactive. The short checkout restores and builds the exact pin successfully.
  No generated products were edited or independently launched app stopped.
- **Latest implemented slice: exact-path keyboard access.** Ordinary path rows now expose a
  contrasting focus outline, full labelled tooltip on focus/Enter/Space and exact Unicode Ctrl+C.
  Tab/Shift+Tab follow visible order; Escape first dismisses detail. Value changes refresh the
  current row; removal restores a live focus target without invoking it. Consumer-only adapter
  changes are in `FolderWindow.FileOperations.Popup.{cpp,h}`; native fixtures are in
  `Commands.SelfTest.FileOps.cpp` and its `Commands.SelfTest.cpp` include/WIL guard. No DxUi code,
  public API or pin changed. The [reviewed gallery](implementation/2026-09-19-path-keyboard/README.md)
  retains both the rejected missing-outline capture and final native-144-DPI French/Unicode image.
- [Final eight-case Debug run](../../../TestRuns/4cb089111a23/Commands/2026-09-19_235421-i26-path-keyboard-final/README.md)
  passes **24/24**, zero skips, zero warnings/errors and zero runner audit issues. Raw UIA focus,
  exact copy, navigation/removal and all three 391-pixel focus-outline witnesses pass. Workspace
  `3c9c38e836f4a3fcc0a0198d190f734d94c5dbc0c0e4aa64d7373bde05551a8e`; receipt
  `7cd5d85650fdbfed0d799c8bd81f8e653587355a6ded8a3357d0671266b1c885`. Initial 2/3 clipboard-restore failure, compile-only rejection and intermediate
  23/24 animation failure remain archived. The fixture now snapshots available clipboard formats
  before mutation. Calling the base focus callback repairs unacknowledged UIA/paint focus. The
  presentation fixture establishes notified owned foreground, keeps its original deadline, and
  settles in 520–551 ms with eight paints; a retained 6.920-second paint gap explains the previous
  missing prerequisite without asserting an unobserved OS visibility cause.
- [Same-fixture resource comparison](../../../TestRuns/4cb089111a23/Commands/2026-09-19_235910-i26-path-keyboard-resources/README.md)
  passes **3/3**, zero skips/warnings/errors/audit issues: 75 lifetimes, 4,800 redraws, attachment
  counts return on each close and hidden frames stay zero. Baseline 32-redraw medians are
  246,061 / 245,517.5 / 244,920.5 us; candidate medians are
  257711.0 / 256039.5 / 254342.5 us. Candidate warm-to-settled private deltas:
  315392 / 557056 / -593920 bytes. All rounds, p95, handles and quantized CPU
  remain in comparison/summary records; bounded Debug evidence does not close Release/FPS/GPU/
  allocation/long-soak gates. The resource function hash is unchanged from the retained baseline.
- **Previous next-work checkpoint (metadata rows now completed above), Astra High:** complete
  scrollable/selectable long-path detail and adaptive conflict/Issues presentation. Primary Release
  PID 76260 was independently owned at the start of this slice and was not stopped. Library full-scene/matrix/
  interleaved evidence and clean-checkout dependency validation already exist further down its
  checkpoint; reconcile that evidence before repeating work. Publication/current-main CI and
  explicit qualified pin adoption remain pending. Full, native 96/144/192 DPI and AT gates remain
  open. Both plans/continuation stay ACTIVE; no required gate is waived.

- Resource timing investigation retains a fresh baseline/candidate pair, all three repetitions
  passing in each run. Baseline medians: 253847.0 / 253838.0 / 254097.0 us;
  candidate medians: 252884.0 / 253808.0 / 253131.5 us.
  The earlier baseline-to-first-candidate increase is retained alongside this repeat; source attribution
  and final performance acceptance are not inferred from a single pair. See the
  [complete four-run comparison](../../../TestRuns/4cb089111a23/Commands/2026-09-20_001041-i26-path-keyboard-candidate-repeat/comparison.json).
  Candidate source hashes were verified after restoring the four files from their saved snapshot.
- Slice closeout: pinned formatting, whitespace and spec inventory pass with zero blocking findings.
  Six explicit TestRuns archives pass for 69 files; the whole inventory passes for 5,118 files.
  Nine metric selections match the raw rows and
  hashes, all archive receipt/workspace identities agree, and the four tested source/two PNG hashes
  match. DxUi documentation-only skills/specs/format checks pass; library gallery pixels are unchanged.

- **Previous implemented slice: measured graph labels and resource-lifetime witness.** Speed and
  qualified ETA now paint in cached, measured rows above the history plot; long French text stacks
  and wraps. The shared graph retains its plot/current-rate marker. Two bounded text layouts per
  visible graph and stable tooltip/UIA Rate/Eta regions replace overlapping marker labels. Changes
  are in `FolderWindow.FileOperations.Popup.{cpp,h}` and `Commands.SelfTest.FileOps.cpp`; no canonical
  DxUi code, API or pin changed. Normal French 480/640/760-DIP captures and synthetic long French
  480-DIP capture are retained in the
  [reviewed gallery](implementation/2026-09-19-graph-labels/README.md).
- Native Debug evidence: [eight cases repeated three times](../../../TestRuns/4cb089111a23/Commands/2026-09-19_222308-i26-graph-labels-final-ui/README.md)
  pass **24/24**, then [long-French UIA/pixel/bounds stress](../../../TestRuns/4cb089111a23/Commands/2026-09-19_223447-i26-graph-labels-french-verified/README.md)
  passes **3/3**, zero skips. The latter had one C5246 initializer warning, subsequently fixed with
  explicit array braces; its warning-repair record distinguishes tested source from the final
  zero-warning resource build. Production popup code is unchanged between these final runs.
  A 6/6 intermediate was visually rejected because centered paragraph alignment hid the labels;
  near alignment plus exact label-pixel probes repairs that gap. A later 0/3 bounds witness counted
  invisible trailing wrap spaces; the corrected test reports zero visible overflow and separately
  retains 4.5197754 px trailing-whitespace overhang. Failed builds/attempts remain archived.
- A new exact `cmd_pane_fileops_popup_resource_lifetime` case runs five warm-up plus twenty measured
  lifetimes per repetition: three repetitions cover 75 lifetimes and 4,800 explicit redraws per run.
  Host/device attachments return to baseline after every close; all hidden intervals add no hosted
  frames. Actual geometry is 760 x 384 DIP at native 144 DPI after production auto-sizing, despite
  requesting 760 x 480. Private bytes, working set, GUI/process handles and quantized CPU totals are
  retained. This bounded witness does not establish allocations, GPU residency, wakeups or long soak.
- Both A/B pairs use the same digest-bound resource fixture and exact filter/order/configuration.
  Original baseline 32-redraw medians: **243,928 / 244,233.5 / 243,354.5 us**; first candidate:
  **255,824.5 / 253,749 / 253,442 us**. Repeat baseline: **242,986.5 / 244,851.5 / 243,820 us**;
  unchanged candidate repeat: **246,061 / 245,517.5 / 244,920.5 us**. The larger slowdown did not
  reproduce; retain all samples and keep final timing/Release acceptance open, without rebaselining.
  Candidate repeat warm-to-settled private deltas are **+258,048 / -909,312 / -1,470,464 bytes**;
  the repeated baseline includes +376,832 bytes in its first repetition. No monotonic growth was
  demonstrated over this bounded fixture; no zero-allocation/CPU claim is made.
- [Complete comparison and receipts](../../../TestRuns/4cb089111a23/Commands/2026-09-19_225208-i26-graph-labels-candidate-repeat/README.md)
  bind final workspace `8d490be258745f83819ebae82bafe7d756a477c5d195433bae2dc69927347717`,
  receipt `a3700fca2b600e2bdad885729701305fb36547d030c920f7192bc484f4fa8e1f`. Both final resource
  runs pass 3/3, zero skips, zero warnings/errors and zero runner audit issues. Raw data/source/logs
  remain in `C:/RedSalamander.Perf/evidence/i26-ui`; exact S-drive runs are preserved under its
  evidence directory. Final candidate sources were restored and their saved hashes verified after
  the baseline repeat. Both plans and the recurring continuation remain ACTIVE.
- **Previous next-action record (superseded above), Astra High:** exact-path keyboard/full visible detail, then remaining
  adaptive conflict/Issues slices. Investigate an isolated local checkout for Release qualification:
  independently launched PID 76260 still owns the primary Release output and must be preserved;
  that exact-path lock need not block a separate checkout. Keep Full, physical mixed-DPI,
  publication/current-main CI/explicit qualified DxUi pin adoption and the preserved library
  developer-local-settings validator blocker open. No ownership or performance gate is waived.
- Slice closeout: pinned clang-format 22.1.3, staged whitespace and spec inventory pass (zero
  blocking findings). Eight explicit archives pass for 86 files; whole archive inventory passes
  for 5,049 files. All selected metric rows match their digest-bound raw lines, all six PNG hashes
  match their raw captures, and the three final C++ hashes match the restored tested source.
  DxUi skills/specs/format pass; its unchanged dependency-validator blocker is not retried or waived.

- **Latest qualification: prepared-frame recovery and hidden client rendering.** Test-only popup
  hooks use the existing pinned DxUi diagnostics; no production renderer, library source or pin
  changed. Three consecutive losses per repetition reject the old frame, replace the context,
  restore French WGC pixels and retain disclosure actions. The initial case passes 3/3; the final
  expanded run passes **6/6, zero skips**, including the progress-contract case, each repeated three
  times. All nine final recovery captures pass (95 foreground / 4,906 background witness pixels).
  Nine real-message hidden intervals (359–375 ms) show zero hosted render-count increments;
  showing resumes rendering and the final image remains byte-identical to the reviewed frame.
- Receipts and all recovery/hidden samples are retained in
  [initial evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-19_213110-i26-device-recovery/README.md)
  and [final evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-19_213557-i26-device-recovery-idle/README.md).
  Both builds report zero warnings/errors and both runner audits report zero issues. Raw runs/logs
  remain under `C:/RedSalamander.Perf/evidence/i26-ui/device-recovery*`; exact completed S-drive runs
  were moved into its evidence directory. The gallery records final capture provenance without
  duplicating identical PNG bytes. This is test-only qualification of the previous implementation,
  not a newly paired rendering-performance comparison or a memory/lifetime soak pass.
- **Next Astra High work:** measure prepared-frame/host resource retention and hidden/idle CPU
  with a matched fixture, then fix graph-label/plot separation and exact-path keyboard access.
  Keep native mixed-DPI, Release/Full, hardware removal, publication/current-main CI/explicit pin
  adoption and the preserved DxUi developer-local-settings validator blocker open. The hidden
  timer intentionally serves the normative taskbar contract; do not remove it as a rendering fix.
  Both plans remain ACTIVE and the recurring continuation remains enabled.
- Closeout for this qualification: final workspace `e2418b58bfc337a33de6387e09db541dc0bab3378974eff09beafd7780b0ead0`,
  receipt `3884060edf5ce3794af78d117d53de89061c1bf81d96b8cc02a9d8811f695eae`; pinned formatting,
  whitespace and spec inventory pass. Explicit archives pass for 22 files; whole inventory passes
  for 4,963 files. Selected metric rows/digests and all four final gallery capture identities match
  their raw sources. DxUi skills/specs/format pass; its unchanged dependency blocker is preserved.

- **Latest implemented slice: one client presentation.** The popup now records its existing
  chrome/text into one Direct2D command list on the pinned WindowHost device, then replays it
  beneath the retained controls. It no longer creates an independent HWND client render target.
  Target/DPI/transform/text state is restored after preparation; replaced frames and device-bound
  resources are released. Zero-sized/minimized hosts suspend instead of entering permanent failure UI.
  Changed production paths are `FolderWindow.FileOperations.Popup.{h,cpp}`; tests are in
  `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`. No library source/API/pin or conflict policy changed.
- Diagnosis retained a readable 3/3 instrumented run, then a 0/3 failed run: all 1,050 legacy draws
  returned S_OK with post-draw occlusion, while target/client sizes matched and hosted painting
  remained fast. Failure began before the new owned cover appeared. This identifies the failing
  presentation boundary, not the underlying Windows/driver cause. Probes/retry limits were not weakened.
- Final Debug validation passes **24/24, zero skips**, eight explicitly named cases repeated three
  times. It covers French WGC pixels, covered/uncovered capture, absent-color rejection, synthetic
  zero-client recovery, paths/UIA, progress, both native failure surfaces, presentation/taskbar,
  conflict metadata/close/Escape and completed-group/navigation. An intermediate run passed the
  capture case 3/3 but failed nine shared-fixture checks: its obsolete three-action expectation and
  overflow Skip invocation were updated to the already-approved four-action contract.
- [Final evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-19_211202-i26-paint-single-surface-qualified/README.md)
  binds workspace `cc3c77980efddd97007a3320b223bb97fd2a907aabb06b1b1f56483261dcd836` and receipt
  `9dfcdf7e1686fd3f536fe54fe6d8acea5d1b62f697597045a2463e1eaec51728`. The clean final build has
  zero warnings/errors. Earlier compile and zero-case-selector failures are recorded, not counted
  as tests. The runner's one advisory was the preserved failed-build directory; after moving that
  exact owned run into evidence, the separate sandbox audit is clean. Original reports remain intact.
- [Seven reviewed images](implementation/2026-09-19-single-surface/README.md)
  include the rejected baseline and readable final French 480/640/760-DIP captures at native 144 DPI.
  Both broader candidate runs use identical production source. Final eight-redraw medians are
  66,462/66,438.5/66,466 us; p95 is 66,720/66,805/66,674 us. All samples and diagnostic baselines
  remain archived. The expanded test fixture and rejected baseline prevent claiming complete paired
  performance acceptance; no Release, FPS, allocation or idle-resource gate is waived.
- Raw data and source are under `C:/RedSalamander.Perf/evidence/i26-ui/paint-*`, with completed runs
  preserved under `S:/RedSalamander.Perf/evidence`. Normative rendering/performance docs and gallery
  are updated. **Next Astra High slice:** qualify prepared-frame device recreation and hidden/idle
  resource lifetime, then graph-label/plot separation and exact-path keyboard access. Retain the
  independent Release PID 76260, Full/native mixed-DPI gates, publication/current-main CI and explicit
  DxUi adoption as open. Both plans remain ACTIVE; the recurring continuation remains enabled.
- Closeout: all three C++ digests match the final tested snapshot; pinned clang-format 22.1.3,
  whitespace, spec inventory (zero blocking), and explicit archive validation (46 files) pass.
  DxUi skills/specs/format pass. Its dependency validator separately rejects the pre-existing,
  ignored `src/DxUi.vcxproj.user` (last modified September 13); that unrelated developer-local file
  is preserved, and the dependency check is not claimed green. Whole archive inventory passes
  for 4,941 files. All selected metric rows match their digest-bound raw originals and all seven
  PNG digests match the reviewed manifest; runner trailing whitespace/blank tails alone are normalized.

- **Previous implemented slice: ordinary full-path semantics.** Running/Delete/completed path rows
  now attach the complete localized label and exact Unicode value to the existing tooltip/UIA
  region on the painted path rectangle. In-flight row slots are bounded and refreshed from the
  current snapshot. Running destination-folder and completed destination-file meaning stays distinct;
  empty/replaced/retired values leave the current tree. No operation policy, action or library pin changed.
  Source changes are `FolderWindow.FileOperations.Popup.cpp` and `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`.
- The identical fixture (`502a384745e01955f981f4a0e5ba81fe990eac773e8518feea4c7bafa497b097`)
  fails the original baseline's missing full-path UIA witness, then passes 3/3 in each of two candidate
  runs. Actual names/help/row IDs and positive visible screen bounds pass at 480/640/760/480 DIP,
  along with two in-flight files, supplementary Unicode, replacement/retirement, Delete, completed
  paths and empty values. Existing disclosure/settings/action/capture checks also pass. All four
  builds have zero warnings/errors and clean sandbox audits; no test was skipped.
- Final workspace `ab16b687ae661c332ae2bee5a553c18d75f19a7a410d12aecc1cd1f121089c9d` is bound by
  receipt `d6910257339812ab48665e54a7d14a8d9fb9dacfaa987aefe0c70cce59d4efb8`.
  [Four-run evidence](../../../TestRuns/4cb089111a23/Commands/2026-09-19_200946-i26-path-access-candidate-repeat/all-series.json)
  retains every timing sample; compact metrics retain every selected row in original order with
  raw byte/row counts, SHA-256 and line-number binding. Original baseline median/p95 is 66,520/67,643 us.
  Final candidate medians are 66,328/66,327/66,572.5 us; p95 is 67,210/66,985/67,229 us. Earlier
  candidate p95 values 133,511/67,286/131,318 us remain flagged, not filtered or pooled.
- **Capture regression remains unresolved.** The repeated unchanged baseline failed 0/3 because
  legacy text/background disappeared while hosted controls remained. All three repetitions returned
  capture `0x8007000D` with zero foreground/background witnesses; 1,020 recorded legacy EndDraw results
  remained S_OK. Its approximately 374-ms redraw medians cannot clear performance acceptance.
  The same source hashes recover in the final candidate, without a renderer/harness change. Do not
  attribute the intermittent cause or call the failing repeat a valid comparison. Three final
  readable captures plus the rejected baseline are reviewed in the
  [path-access gallery](implementation/2026-09-19-path-access/README.md).
- Raw logs/data and hash-preserved source are under `C:/RedSalamander.Perf/evidence/i26-ui/path-access-*`;
  whole completed runs were moved with validated paths to `S:/RedSalamander.Perf/evidence`.
  Screenshot pixels still show abbreviated paths and graph-label/plot overlap. This slice establishes
  semantic access, not keyboard copying, visible complete detail, hover or screen-reader qualification.
- Closeout: the production source still matches the final tested digest. Pinned clang-format 22.1.3
  only normalized one test-file line ending after the runs; the newline-normalized texts are identical.
  [The exact before/after hashes](../../../TestRuns/4cb089111a23/Commands/2026-09-19_200946-i26-path-access-candidate-repeat/format-normalization.json)
  distinguish that final formatting change from the common tested fixture. Formatting, whitespace,
  spec inventory (zero blocking), DxUi skills/specs and explicit pre-stage validation (38 files) pass.
  Whole-inventory validation covers 4,895 files. Selected metric digests normalize record terminators
  to LF for portability while preserving row text/order; raw source digests bind the original bytes.
  All four published images are visually reviewed and their manifest digests match.
- **Next dependency-ready slice: Astra High.** Diagnose the recurring legacy-layer loss against the
  retained failing capture and S_OK paint receipts; instrument the existing legacy/hosted composition
  boundary before another controlled comparison, without loosening probes or repeating unchanged runs.
  Graph-label separation and exact-path keyboard access remain independent implementation work.
  Keep Release (independent PID 76260 preserved), Full, physical mixed DPI, final performance/resource
  gates and publication/main-CI/explicit DxUi adoption open. Both plans remain ACTIVE.

- **Previous implemented slice: custom-speed dialog reflow.** Creation now targets the owner's monitor
  before DPI sizing/centering. The existing wrapped-text helper measures Body/Small label and
  help/error height, and layout grows the client when necessary after resize or validation changes.
  Correction does not shrink the dialog. Growth respects monitor bounds, ignores minimized/empty
  clients and guards synchronous resize reentrancy. No parser, accepted speed, default action,
  task lifetime or shutdown policy changed. Code is limited to `FolderWindow.FileOperations.Popup.{h,cpp}`
  and `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`; the DxUi pin is unchanged.
- [Five reviewed harness images](implementation/2026-09-19-speed-dialog-layout/README.md)
  show the failing constrained baseline and four final French states. The live witness reports
  baseline help bottom 132 DIP versus action top 92 DIP in the requested 300x142-DIP client.
  Original baseline/candidate pass counts are 0/1 and 1/1; repeated baseline/final counts are 0/3
  and 3/3, zero skips. All four builds have zero warnings/errors and clean sandbox audits. Final
  repetitions also pass the existing UIA surface/value and nested modal-shutdown checks. The
  harness's existing warning/desktop lease/restoration governs foreground use; no desktop screenshot
  or fake native-DPI claim was used.
- All runs retain fixture SHA-256 `aa40029bddbb4454a9caf8a421bce868e739f2fe032d22fae476d5cf230f7f26`.
  Final workspace is `20e929da89cbba329de8af7f88ed73376ecdd25b57e07fb4323987c07fc5a68f`, receipt
  `4c4d04b8bfdf3ecc948c8bb1cddab676faf040347e1b8f6c86c17c893ed340a2`. The initial pair's p95
  rose 102,916 to 118,309 us; this flag is retained. Sixty samples per repeated variant have median
  66,509.5/66,514.5 us and p95 121,362/112,266 us. The rise did not repeat, and the baseline itself
  varied; no consistent degradation is established by these Debug diagnostics. This is not Release,
  FPS, allocation/idle or final performance acceptance. See the
  [four-run comparison](../../../TestRuns/4cb089111a23/Commands/2026-09-19_194500-i26-speed-layout-candidate-repeat/README.md).
- Logs/artifacts remain at `C:/RedSalamander.Perf/evidence/i26-ui/speed-layout-*`, with whole completed
  runs moved using checked paths to `S:/RedSalamander.Perf/evidence`. Hash-verified instrumented baseline
  source and final source are retained under `speed-layout-source/baseline` and `speed-layout-source/final`
  in that C: evidence folder, for targeted follow-up without rediscovery. Intermediate `candidate`
  source predates the final minimized/empty-client guard; do not confuse their receipts.
- Closeout checks: pinned clang-format 22.1.3 and staged whitespace pass, and all three C++ digests
  match the final tested snapshot. Spec inventory reports zero blocking findings; DxUi skills/specs
  pass. Four archives pass explicit pre-stage validation (33 files) and whole-inventory validation
  (4,857 files). All five published images were visually reviewed and their manifest digests verified.
- **Next dependency-ready slice: Astra High.** Address long-path exact access or graph-label/plot
  separation through existing qualified primitives, with live witnesses and paired evidence. Keep
  speed-dialog native 96/144 transition, extremely short-work-area fallback, live locale changes
  and full keyboard qualification open. Existing publication/main-CI permission, explicit immutable
  DxUi adoption and fresh mixed-DPI gates remain pending. Release PID 76260 remains independently
  launched and preserved, so this slice has no Release result. No build remains running; both plans
  remain ACTIVE. Do not treat this bounded repair as all of P1 or completion of either plan.

- **Previous completed slice: French footer action fit.** `LayoutChrome` no longer shrinks Pause/Resume
  and Cancel to 72 DIP below 560 DIP. It measures the hosted Body font, adds 24-DIP padding, reserves
  the larger Pause/Resume width and retains the 112-DIP minimum. A bounded three-label/format cache
  avoids unchanged-paint measurement. The independent runtime witness checks actual text/control
  padding, active action/Queue/Options targets and no overlap at 480/640/760/480 DIP. Changed code:
  `FolderWindow.FileOperations.Popup.{h,cpp}` and `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`.
- [Reviewed before/after captures](implementation/2026-09-19-footer-action-fit/README.md)
  show complete running-copy French footer labels at all three widths. The
  [baseline](../../../TestRuns/4cb089111a23/Commands/2026-09-19_184300-i26-footer-actions-baseline/README.md)
  fails 0/1 with -23.463135-pixel remaining label padding at 480 DIP; the
  [final candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-19_184300-i26-footer-actions-final/README.md)
  passes 1/1, zero skips. Both have zero warnings/errors and zero sandbox audit issues. Fixture hash
  `bb286a8b2b377c20c0afe2bcb2fc2ea598de270fd078b145b60516eb8beb2238` is identical. Final workspace
  `e52e828dc55849b449f521445a514c68775c116f8d6799c8ae45fda76b7f046e` is bound by receipt
  `9e2a6253d91cb47f5f666b61816b3cf686b62388e1b6632d87aeae345ba0af1e`. The initial candidate's
  four repeated C5246 diagnostics were fixed with aggregate braces before final validation.
- Twenty eight-redraw medians are 66,405.5/66,561 us; nearest-rank p95 is 67,320/67,160 us.
  The final candidate has two setup cache fills (107/80 us, three labels each), neither in the measured
  redraw sequence. This is a focused Debug diagnostic, not broad non-regression or resource acceptance.
  Raw logs/copies remain in `C:/RedSalamander.Perf/evidence/i26-ui/footer-actions-*`; completed runs
  were preserved under `S:/RedSalamander.Perf/evidence` with checked absolute paths before another run.
- Closeout checks: pinned clang-format 22.1.3 and staged whitespace pass; all three source digests
  match the final tested workspace. Spec inventory has zero blocking findings, DxUi skills/specs
  pass, and the archives pass explicit pre-stage validation (17 files) and whole-inventory validation
  (4,824 files). All four published images were reviewed and hash-bound in the capture manifest.
- **Next dependency-ready slice: Astra High.** Characterize and repair the 96-DPI custom-speed
  help/error overlap with live layout assertions and the deterministic harness, using existing
  qualified primitives where sufficient. Long-path exact access, graph label separation, general
  wrapping/measured footer height and paused/Resume full interaction coverage remain open. Existing
  publication/main-CI permission, explicit pin adoption and physical mixed-DPI gates remain pending.
  Independent Release PID 76260 was preserved, so this slice has no Release qualification. No DxUi
  code/pin changed, no build remains running, and both plans remain ACTIVE.

- **Previous completed slice: French discovery action width.** The fixed 84-DIP slot clipped
  `Ignorer la découverte` even at wide widths. Production now measures the actual hosted Body font,
  adds 24 DIP padding and caches one label/retained format identity. The runtime guard measures the
  live control independently and checks fit, active Skip/limit/Cancel targets and no overlap at
  480/640/760/480 DIP. Three candidate images and the 480-DIP baseline were visually reviewed in the
  [before/after gallery](implementation/2026-09-19-discovery-action-fit/README.md).
  The earlier description of the task row clipping off the left edge was imprecise: its centered
  label was clipped inside an undersized button. Narrow footer labels and long-path access remain open.
- Retained Debug [baseline](../../../TestRuns/4cb089111a23/Commands/2026-09-19_181400-i26-discovery-actions-baseline/README.md)
  fails 0/1 as intended (184.48828-pixel text in a 126-pixel control); the
  [candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-19_181400-i26-discovery-actions-candidate/README.md)
  passes 1/1, zero skips. Both builds have zero warnings/errors and identical fixture SHA-256
  `e60006635e0c358fae90c729dfc5a0df863ed574dcccd47e655fb2fc27707c2e`. The 20 eight-redraw batch
  medians are 66,430/66,422.5 us; nearest-rank p95 is 67,065/67,135 us. Candidate records one
  14-us production measurement across all widths. This is focused Debug evidence, not full performance
  acceptance. Candidate workspace is `e27ee2e16e041028a512209b4f81ae2436a5342334a253b2dae2a9bad5c8f30f`;
  receipt is `cf1fdd8325ca006ce167d60b78540e819798efd560f1aa02173ac023898e7ad6`.
- The candidate runner reports one sandbox issue for this task's preserved named baseline. Both
  completed run directories were moved, with validated paths, from `S:/RedSalamander.Perf/runs` to
  `S:/RedSalamander.Perf/evidence`; the separate post-archive audit is clean. The original audit is
  retained unchanged. Full logs/copies are in `C:/RedSalamander.Perf/evidence/i26-ui/discovery-actions-*`.
  An earlier autogenerated baseline directory was deleted by runner cleanup; it is not used for the
  pair. The retained repeat precedes restoring the candidate and uses the warning-corrected fixture.
- Closeout checks: all three C++ digests still match the tested candidate; pinned clang-format
  22.1.3 and staged whitespace pass. Both archives pass explicit pre-stage validation (18 files),
  and whole-inventory validation passes (4,807 files). Spec inventory has zero blocking findings;
  DxUi skills/specs validation passes. All four published PNGs were reviewed and hash-bound.
- **Next dependency-ready slice: Astra High.** Add the remaining footer/long-path and 96-DPI
  speed-dialog fit witnesses, then repair using existing qualified primitives where sufficient.
  General action-flow adoption still requires publication/current-main CI and the supported pin
  update; the prior permission question is pending. Do not bypass it. Fresh physical mixed DPI
  remains pending (both displays were 144 DPI). Release qualification for this slice is blocked by
  independently launched PID 76260 at `.build/x64/Release/RedSalamander.exe`, observed again at closeout;
  preserve it. No library code/pin changed. Both plans remain ACTIVE and no build is left running.

- **Previous completed slice: localized open-discovery progress.** The production renderer now formats
  `IDS_FMT_FILEOPS_DISCOVERY_PROGRESS` through the active satellite instead of forcing the embedded
  source-language sentence. The retained French baseline fails with `2,00 MB of 8,00 MB so far, 3
  of 13 items`; the candidate semantic contract passes with `pour l'instant` and `éléments` and rejects
  `so far`. The same shared helper is used by paint and the runtime assertion. Changed code is limited
  to `FolderWindow.FileOperations.Popup.{h,cpp}` and the Commands File Operations fixture; the
  consumer DxUi pin is unchanged.
- [Two reviewed French application captures](implementation/2026-09-19-discovery-localization/README.md)
  cover open discovery at 480/760 DIP, long paths, dark theme and native 144 DPI. Their content probes
  report 1,952/387,172 and 2,408/636,219 expected foreground/background pixels. The counter sentence
  is readable, but the images also preserve a concrete P1 failure: the task action row clips off the
  left edge and the long source path truncates aggressively. These captures qualify localization and
  harness integrity only; they do not qualify adaptive layout or fresh mixed DPI.
- Baseline run `discovery-localization-baseline` fails 0/1 exactly at the new French assertion; initial
  candidate runs `discovery-localization-candidate-progress` and
  `discovery-localization-candidate-capture` pass 1/1 each with zero skips and zero build warnings/errors.
  After the code-format and documentation closeout, `discovery-localization-final-progress` also
  passes 1/1 with zero skips and a zero-warning/error Debug rebuild; its workspace snapshot is
  `2990ab4b4d840ffb734e5ee18bf30a3e57bb6a9555e0905e91a6002baab54457` and build receipt is
  `230ca41121d88c9ab6280e86b9f48c5bd80971d81b8e8d6ab12005759a7d9020`.
  Baseline workspace snapshot is `a1cd622c37f2f9df9abd3e3c9dc4e562a6cff4e140d044e2a05bfdef8dc47f8e`;
  initial candidate snapshot is `bb6399d87d6a248d116aaf2f8898abf68edb16e354eaa659e36e0d954a3b154c`.
  The C: audits list 226/227 unrelated preserved run directories; they are not clean-sandbox claims.
  Raw logs and artifacts remain under `C:/RedSalamander.Perf/evidence/i26-ui/discovery-localization-*`
  and `C:/RedSalamander.Perf/runs/discovery-localization-*`.
- **Next dependency-ready slice: Astra High.** Turn the observed long-French action/path clipping into
  direct measured geometry failures, then implement the smallest responsive consumer layout using the
  qualified shared action-flow API once its immutable pin can be adopted. While publication/main CI is
  pending, continue independent P0 characterization for long French fit and the 96-DPI speed-dialog
  overlap rather than bypassing the updater. Retain exact-path access and whole-button wrapping as
  separate assertions. Keep one writer/build lane and queue at most one model-routed continuation.

- **Previous completed slice: capture content validation.** The canonical test-harness helper now
  accepts client-pixel color witnesses, preserves rejected PNGs and performs at most three
  event-driven frame attempts, repainting only its owned window without activation after rejection.
  The throughput case requires French label foreground/background and tests an absent-color rejection.
  Changed code: `Tests/TestSupport/WindowScreenshot.h/.cpp` and
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`. No product/library code or pin changed.
  [Eight reviewed Debug/Release captures](implementation/2026-09-19-throughput-capture-validation/README.md)
  contain readable legacy text and the corrected destination path. The 480-DIP footer still clips.
- Final Debug run `20260919T150817Z-42640-3ae9c86a7496424fa396524f225ab1fd` passes 3/3 repetitions;
  Release `20260919T151219Z-26152-fef8dbf0834e4dcda73aa5fbb9197516` passes 1/1. Both have zero skips,
  zero sandbox issues and zero build warnings/errors. Source snapshot is
  `3860bf5dceaa9fab857b1d13cc5f617efe7e24c8ebbdcac63fc5dd74a7e9212f`; receipts are
  `28cbb9e2a7a85e2f64d3d305b32e39513f7005ea6f8b2aac204d0e811365c9f8` (Debug) and
  `a32989defe7d81e1061db01ddc1e23594233158993814c026c8fc4488aecc1ab` (Release).
  All 16 positive witness rows are 35 foreground/4859 background pixels. Every negative control
  correctly fails with a retained diagnostic image. Full logs/raw artifacts use `capture-layer-*`
  under `C:/RedSalamander.Perf/evidence/i26-ui`. Curated baseline/Debug/Release archives are linked
  from the gallery. The first probe build had an undeclared DIP helper; corrected to the fixture's
  existing MulDiv pattern. That failed build is diagnostic only. No build remains running.
- The original layer loss is **intermittent and unattributed**, not a proven Debug-only renderer bug.
  Before retries were implemented, probes alone passed 3/3 cases and all 12 images. Current capture
  integrity can be verified independently, allowing documentation work to continue, but do not claim
  reproduced transient recovery or drop inspection/Full visual gates. Previous rejected PNGs remain
  intact. Other callers without witnesses still require review; any future rejection is actionable.
- **Next dependency-ready slice: Sol High**, repair the discovery progress resource lookup in
  `FolderWindow.FileOperations.Popup.cpp`: `FormatEmbeddedStringResource` forces the source-language
  sentence for `IDS_FMT_FILEOPS_DISCOVERY_PROGRESS`, while the French satellite already translates it.
  Use the localized loader, add a meaningful French runtime assertion/capture using the established
  harness, update specs and retain focused evidence. Avoid unrelated string rewrites. After this
  bounded slice, route the next complex layout/policy/input work back to Astra High. Keep one writer,
  one build lane and at most one queued continuation. Pending publication approval/main-CI/pin and
  fresh physical mixed-DPI gates remain unchanged; continue independent work without repeating them.
- Capture-slice closeout: documentation/localization contracts pass 27/27 with zero skips; spec
  inventory reports zero blocking findings. Changed C++ files pass pinned clang-format; staged
  whitespace and all eight PNG digests pass. Three new archives pass explicit pre-stage validation
  (21 files), and the whole archive inventory passes (4789 files). DxUi skills/specs pass; its
  coordination-only update is `5650cc2`. Both implementation plans remain ACTIVE.

- Latest shared-library slice is committed as `26a3d044843396ce77c347593c04fe0a6016c00f` in
  canonical DxUi. Final x64 Debug/Release/ASan cover all 18 suites (Menu capability skips recorded);
  all three ARM64 cross-builds pass. Final five-theme gallery is published/reviewed. Clean worktree
  `C:/Users/eric/.codex/worktrees/localized-layout-qualified/DxUi` passes dependencies, skills and
  specs validation, preserving the personal ignored project file in the original checkout.
  Ten alternating paired common-fixture runs retain every flag: no systematic regression was
  confirmed, deterministic resource budgets remain unchanged, and clean-memory uncertainty is
  documented in DxUi `Measurements/LocalizedAdaptiveLayout/2026-09-19/interleaved-final/README.md`.
  Native product/mixed-DPI and long-run gates remain open; this is not product acceptance.
- Consumer pin remains `f5c7fc7403e352425b8e82b7a4da1c042d942ede`. The supported update command
  requires successful current-main CI. Library publication is an explicit lifecycle action under
  `Specs/Build/Build_Toolchain.md`; an asynchronous user question asks permission to push the exact
  tested branch and open a draft PR. No publication/merge has been performed and no approval may
  be inferred from elapsed time. Continue independent work while awaiting the answer.
- The bounded Sol High settings slice implements `fileOperations.popupThroughputExpanded` with default
  `true` in the typed model, schema, parser, canonical writer and non-default pruning. Its focused
  Commands test covers absent/true/false/invalid values and proves `false` survives as the sole
  difference. The unchanged-production witness failed 0/1 for the missing retention; candidate run
  `20260919T133051Z-76472-2b4998908dba40dcbf2937cadee166cd` passes 1/1 with zero skips and a
  zero-issue sandbox audit. Both builds had zero warnings/errors. Documentation contracts pass 18/18
  and the spec inventory has zero blocking findings. Curated [baseline](../../../TestRuns/4cb089111a23/Commands/2026-09-19_152341-i26-throughput-settings-baseline/README.md)
  and [candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-19_153051-i26-throughput-settings-candidate/README.md)
  receipts retain exact provenance. Direct standalone `SettingsSchemaTests.exe` launches hit unrelated
  SettingsStore CAS/read-sharing environment guards and are diagnostic only; do not repeat unchanged.
  The final governed Full runner remains required.
- Throughput live-card slice now implements a hosted disclosure on the existing pin, independent
  per-card state, remembered defaults for new/recreated cards, a collapsed rate/ETA summary, retained
  task actions, removed hosted graph controls and preserved outer geometry. Explicit changes queue
  the existing settings save; painting and default synchronization do not write settings. Native
  tests cover actual file reload, multiple cards and compact/footer-only suppression. The canonical
  strings and cs-CZ/fr-FR/ja-JP/sk-SK satellites are updated. Production changes are in Popup.cpp/.h,
  State.Runtime.cpp, State.Private.h and FileOperationsInternal.h; tests are in Commands.SelfTest.cpp
  and Commands.SelfTest.FileOps.cpp. Affected C++ files have the pinned clang-format 22.1.3 formatting.
- [Baseline/first candidate](../../../TestRuns/4cb089111a23/Commands/2026-09-19_161917-i26-throughput-ui-first/README.md)
  retain the missing-control failure then a 1/1 state pass. Expanded Debug redraw medians are
  374500 -> 373680 us (-0.22%); no performance improvement is established. The baseline's one
  stale direct-launch scratch folder was preserved under governed evidence, not deleted. Candidate
  sandbox audits are clean. The [test-enabled Release](../../../TestRuns/4cb089111a23/Commands/2026-09-19_163216-i26-throughput-ui-release/README.md)
  and [final Debug](../../../TestRuns/4cb089111a23/Commands/2026-09-19_164036-i26-throughput-ui-debug/README.md)
  each pass 1/1 with zero skips and zero build warnings/errors. Release receipt is `34621d0175fee7287a3b83ea536f1f322a7a5a70d3c19a808b2e1eca68ad8369`;
  Debug receipt is `62748c0626f03df8c63249ea3e848a85f01177dbf71743358cdf51f841afe02c`.
  A first Release launch omitted `RSBuildEnableTests=true` and had no case results; do not count it.
- Prior capture defect, retained as evidence: [capture review](implementation/2026-09-19-throughput-disclosure/README.md)
  accepts four readable Release images as partial evidence, but rejects all four final Debug images:
  legacy Direct2D background/text is missing while hosted controls remain visible. No failed EndDraw
  HRESULT is recorded. Applying resolved dark theme/French before popup creation did not fix Debug.
  Release's synthetic destinationFolder was empty; the final Debug fixture supplies it. Do not claim
  those Release captures contain the later fixture correction. Raw runs/logs remain under
  `C:/RedSalamander.Perf/evidence/i26-ui/throughput-ui-*`; the latest governed run is
  `20260919T144036Z-56576-209e4ac8dfe145399bbf8508d5b1bee1`. No build is left running.
- The subsequent capture-content guard above replaces first-frame encoder-only acceptance for this
  fixture. If legacy loss returns, relevant product paths remain Render/EnsureTarget/OnSize (HWND
  render target plus hosted composition). Occlusion/presentation readiness remains a hypothesis.
  Do not repeat the theme-only attempt or assert a library cause without a reproducer.
- Keep P3 unchecked: failed save, actual external hot reload/process restart, both placement keys,
  physical keyboard/UIA and final resource checks remain open. The current disclosure label is static
  beside its chevron; newer UIA ExpandCollapse requires qualified pin adoption. Legacy fixed card
  heights and expanded zero/unknown graph labels (library positive-marker paint condition) are open
  consumer/shared work. After the rendering defect, route the bounded discovery-language repair to
  **Sol High**: Popup.cpp uses `FormatEmbeddedStringResource` for `IDS_FMT_FILEOPS_DISCOVERY_PROGRESS`
  although the French satellite already contains the full translation. Add a focused localization
  assertion/capture instead of inventing strings. Publication/main-CI/pin adoption and fresh physical
  mixed-DPI remain blockers for their dependent acceptance; do not repeat pending approval questions.
- Slice closeout checks: documentation/localization contracts pass 27/27, zero skips; consumer spec
  inventory has zero blocking findings. All seven changed C++ files pass clang-format 22.1.3, staged
  diff whitespace checks pass, and all eight PNG digests match. The four new archives pass explicit
  pre-stage validation (28 files); whole TestRuns inventory passes (4768 files). DxUi skills/specs
  validation passes after its coordination update, committed as `f44a932`. These checks do not waive
  the rejected Debug capture or any required implementation/acceptance gate.

- Current continuation after `9d3318382`: repeat scope now uses the existing pinned DxUi Checkbox,
  with model acknowledgement via `SetChecked` and eligibility via `SetEnabled`. Removed the
  overlapping manually drawn indicator/label. Debug layout snapshots expose actual control type/state.
  The retained unchanged-product witness fails missing UIA Toggle. First candidate passes Toggle,
  visible acknowledgement and repeated synchronization, but fails Space: the popup's task pause
  shortcut intercepted it. Routing now lets the Checkbox handle Space; the corrected native case
  passes 1/1 with zero skips, including pointer activation/release-outside and keyboard parity.
  Full baseline/first-candidate evidence is retained under
  `C:/RedSalamander.Perf/evidence/i26-ui/checkbox-baseline` and `checkbox-candidate-first`.
  Builds had zero warnings/errors. The
  [paired archive](../../../TestRuns/4cb089111a23/Commands/2026-09-19_120441-i26-checkbox-candidate/README.md)
  passes the archive validator; matched Debug redraw median is 66516 -> 66993 us (+0.72%).
  No Checkbox capture or completed acceptance is claimed yet.
- Shared DxUi work adds measured action flow, optional multiline Buttons and disclosure UIA.
  Its own plan records validation and the unresolved common-performance comparison; consumer pin
  remains unchanged. Next consumer action: obtain deterministic French Checkbox captures, then
  apply measured consumer layout after qualified immutable pin adoption. Preserve the user's
  independently running Release app (PID 89896, rechecked alive) and use no competing build lane.

- Last completed: P0 source/ownership map, canonical DxUi common Release baseline, and the ordinary
  action-row repair with a matched failing/passing native witness. Debug build: zero warnings/errors;
  native Setup, popup smoke and Cleanup: 3/3 passed. See
  [paired evidence](../../../TestRuns/4cb089111a23/FileOps/2026-09-19_103033-i26-actions-candidate-v2/README.md).
- Captures: [300 French application images / 119 static scenarios](implementation/2026-09-19-ordinary-actions/README.md),
  harness passed 1/1. Both displays now report 144 DPI; fresh 96-DPI acceptance remains open. Old
  captures remain intact. Footer label clipping and other adaptive defects are not resolved by this fix.
- Next: independent DxUi L0 localized fixture/baseline and remaining P0 Checkbox/French geometry
  witnesses. The new 480-DIP capture is a concrete whole-button wrapping failure. Use Astra High. Do not repeat completed
  ordinary-action baseline collection without a fixture or implementation change that requires it.
- Last completed continuation: conflict-primary placement witness in
  `Phase10_DeferredConsentAndRecycleEscalation`, plus rendered Commands expectations. The policy test
  exercises every conflict class and all 64 combinations of its six authority/scope switches, checks
  unique allowed-action partitioning and safe default/Escape, and measures five identical 16-pass
  matrix batches after one full warm-up. The unchanged-placement baseline failed as expected and is
  retained at `Specs/TestRuns/4cb089111a23/FileOps/2026-09-19_105159-i26-policy-baseline` (median 33219 us).
  Candidate placement removes the three-primary cap, orders collision choices Keep Both,
  replacement, Skip, Cancel, and applies Replace/Remplacer labels. Setup/policy/Cleanup passed 3/3;
  [paired evidence](../../../TestRuns/4cb089111a23/FileOps/2026-09-19_105819-i26-policy-candidate/README.md).
  The identical fixture's Debug median increased 3.1% to 34250 us per 24576 constructions. This is
  diagnostic evidence, not Release performance acceptance; required paired Release/resource gates remain open.
  [Commands integration](../../../TestRuns/4cb089111a23/Commands/2026-09-19_110240-i26-policy-integration/README.md)
  passed 1/1: exact order/default/loading, keyboard/UIA, stale overflow rejection, direct Skip and
  destination preservation, plus existing settings/taskbar/reduced-motion/high-contrast assertions.
  Both builds had zero warnings/errors. Allowed-action order/eligibility, cached consent and dispatch
  remain unchanged. The bounded live-output class also publishes its previously capped Skip directly;
  its order/default are explicitly tested.
- Latest captures: [300 French PNGs / 119 scenarios](implementation/2026-09-19-conflict-actions/README.md),
  gallery passed 1/1. At 760 DIP four choices and Remplacer fit; at 480 DIP the longer actions clip.
  Both displays remain 144 DPI. Adaptive wrapping, scope Checkbox, short-work-area scrolling and
  fresh mixed-DPI qualification remain unfinished. Static capture used no activation; Commands used
  the existing notified desktop lease. The capture receipt is `2128d4c7009fe019ba5dabfd5043d411b16a99c2830941d8d85fa2a567688495`.
- Changed paths this continuation: `FolderWindow.FileOperations.State.cpp`, canonical/fr-FR resources,
  `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`; policy metric contract in Testing_PerformanceValidation,
  UI/FileSystem implementation status, this checkpoint and linked evidence. Consumer pin is unchanged.
- Execution notes: freeze all consumer edits, including documents, during builds. Focused runner
  receipts target the application; `-SkipBuild` requires a full-solution receipt and was rejected.
  Use the normal governed runner. Raw logs and the full candidate metric stream are retained under
  `C:/RedSalamander.Perf/evidence/i26-ui`; curated metrics bind the raw digest and row indices.
- Open environment gate: sandbox audit lists 225 unrelated pre-existing run directories; preserve
  them. User-owned Release app was running at P0, so recheck exact-path ownership before Release builds.
- Sandbox resolution for final qualification: initialized the previously absent exact
  `S:/RedSalamander.Perf` root with `Initialize-RSTestSandboxRoot -AllowExternalInitialization` on
  a fixed local drive (about 115 GB free). Use it for the next integration/Full run and retain its
  actual audit result; no clean Full result is claimed yet. Earlier paired measurements stay on C:.
  The user-owned Release app PID 89896 is still running and must not be stopped by automation.
- Slice validation: both spec inventories passed (consumer zero blocking findings); all four curated
  native-run archives passed the archive contract; all 300 PNG hashes/entries and the five paired
  metric rounds were verified. Curated `Specs/TestRuns` folders are ignored by default: retain them
  and use the repository's explicit reviewed force-add workflow when preparing a commit.
- Conflict-slice closeout: both spec inventories pass again (consumer zero findings; DxUi 67 Markdown,
  20 authority files, one active plan). Three new policy/integration archives pass (20 retained files);
  all 300 new PNG hashes and the baseline/candidate/current fixture hash match. Both repositories pass
  `git diff --check`. Logs use `i26-policy-*` under the evidence directory above. No Release/full-suite,
  fresh mixed-DPI, final resource or library acceptance gate is claimed by these focused Debug results.
- Initial model: `gpt-6-astra`, reasoning `high`.
- Scheduled continuation: `implement-file-operations-ui-to-done` (hourly heartbeat in the existing
  task). Use this exact automation ID to disable it after verified closeout; never create a duplicate.
- Existing repository findings: DxUi dependency validation rejects `src/DxUi.vcxproj.user`; formatting
  check flags unchanged `src/Controls/DxUi.cpp:33`. Recheck before treating these as current blockers;
  do not delete personal settings or claim a clean gate without evidence.
- Preserve prior documentation edits and historical 437-image evidence. Keep future commands, receipts,
  changed paths, remaining failures and the next bounded action here as execution proceeds.

