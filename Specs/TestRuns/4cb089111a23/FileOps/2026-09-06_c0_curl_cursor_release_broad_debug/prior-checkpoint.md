# File Operations: reliable core and clear feedback

> **NON-NORMATIVE WORKING PLAN — I14.** Priority: P1 reliability. Owner: the agent
> executing this plan in coordination with the product owner. Reviewed against
> `58f22d0e` on 2026-09-05. Current behavior remains defined by `Specs/<Domain>/`
> and the executable provider contracts; target changes land with code, tests, and
> their authoritative spec updates. This document has one execution status table.

The goal is simple: Copy, Move, Delete, and Rename do what the user requested when
the provider permits it; otherwise the application stays responsive, protects data,
and clearly says what happened, what is uncertain, and what the user can do next.
“Always working” cannot promise an available network, device, permission, or disk.
It requires an honest and usable outcome for each of those failures.

On 2026-09-05 the owner explicitly moved persistent searchable history,
collect-then-review conflicts, link retargeting, and the Shell Copy/Move experiment
to [a later plan](Operation_FileOperations_LaterFeatures_2026-09-05.md). They no
longer block this core milestone. Literal **Copy links unchanged** remains required.
This scope change supersedes the old M3 dependency on those four features.

The previous decisions, activation cards, and execution receipts are preserved in
[the historical record](Operation_FileOperations_ReliabilityFirst_ExecutionRecord_2026-09-05.md).
Its old states and timing claims must not be copied back as current evidence. The
referenced conversation, **Review Reliability Plan Code**, informed this review;
its suggestions were checked against source, not treated as implementation proof.

## 1. Scope and execution status

Supported writable destinations remain Local fixed volumes, UNC/mapped SMB,
FTP/SFTP/SCP, ordinary S3, Microsoft Drive/SharePoint, Google Drive, and writable MTP.
Read-only profiles, IMAP, archive restrictions, S3 Tables, and provider-specific
unsupported operations keep their documented limits. Do not make a provider
permanently read-only merely to close this plan. An unproved exact route must be
contained, repaired, or visibly refused; a capability flag cannot substitute for proof.

States: `LANDED` means implementation exists with historical/focused evidence, but
still participates in the final gate; `READY` means the bounded work below is
authorized; `VERIFY` means prove the implementation and fix a demonstrated gap;
`WAITING` means its listed predecessors are unfinished; `DONE` requires the row's
exit evidence. Update status only in this table. Instructions and evidence records
must not maintain competing state lists, milestone checkboxes, or activation cards.

| ID | State | Work / former IDs | Depends on | Exit evidence / next action |
|---|---|---|---|---|
| B1 | LANDED | Typed plans, immutable terminal results, publication/cleanup, clipboard barrier, rename chain and name contracts; E0–E7, R0a–R0e, R1a–R1d, R2/R3, R7-A10/A15 | — | Historical receipts; revalidate affected behavior through C0/C6. No new outcome model or bridge extraction. |
| B2 | LANDED | Discovery/progress, writable-provider containment, exact same-host Queue/Run/output guards; R0f, R4, R6-A02/A08 | B1 | `20c64170`, `e96e0226`, `31fe0b8c` plus historical receipts. C0 verifies provider walker coverage and the complete overlap matrix. R6-A02 has no second UI implementation. |
| B3 | LANDED | Routine F5/F6 and explicit Shift+F5/Shift+F6 options; R6-A20 | B1 | `c775b9ed`; command/resource tests and zero-warning Debug/Release receipts exist. C5/C6 cover final integration. |
| B4 | LANDED | Name-only artifact exemptions and Properties explanation; R6-A09 | B1 | `58f22d0e`; 4 focused Commands cases, Phase10 9/9, source contracts 188/188, localization 9/9. C5 owns remaining provider-error/refresh behavior. |
| C0 | VERIFY | Review corrections and an explicit route/walker evidence matrix | B1/B2 | Lookup 273/273 and CRUD 3/3 in Debug/Release. Broad Release Copy remains 134/142: four deep-Move source-control 10048 failures; all four exact failed pairs later appear in TIME_WAIT. Both wide copies pass, but prior broad Debug remains 141/142 (serial 3,828/4,097 at timeout). Exclusive fixture ownership is corrected, not a causal transport repair. Next: release completed listing borrows before descent and prove nested-cursor connection reuse. Live FTP/SFTP/SCP-profile listing/S3 pass; IMAP is TLS-policy blocked. Section 4 retains Move proof memory, IMAP, FTP/SSH limits, route matrix and host Rename. |
| C1 | READY | Responsive ingress, Preparing, cancellation, and shutdown; R1d-OR1 / FOS-05 residual | B1/B2 | Command-ingress delay, blocked-reader, close, and lifetime tests; no UI-thread provider waits. |
| C2 | READY | Proof-gated repeated explicit Retry; R6-A07 | B1 | Existing mutation facts drive eligibility; host/provider fault tests show one request on unknown commit and repeated explicit safe retries. |
| C3 | READY | Literal Preserve without a retarget graph; core part of R7-A03 | B1 | Literal link/payload matrix, deep/wide trees, >4,096 links, no-follow and Managed Move source-retention tests. |
| C4 | READY | One hosted popup and a graphics-independent failure surface; R6-A12 | B1 | Attachment and D2D failure tests, keyboard/UIA, safe Cancel All/Close, no duplicate graph/action owner. |
| C5 | READY | Human-readable results and Properties failure paths; remaining R6-A09/UX | B3/B4 | Missing/unavailable/malformed/refresh Properties cases plus the journey matrix in section 2. |
| C6 | WAITING | Consolidation, integrated tests, and core retirement; R8/E8 | C0–C5 | Authoritative specs and inventories current; archived evidence; exact Fresh Full passes and includes inherited I17 qualification. |

These rows replace the repeated activation/exit-card machinery. Before editing a
row, reconcile its named files against the review commit, record the actual starting
commit and intended tests in that row's evidence, and ensure no other owner edits
the same area. The user's instruction already authorizes the bounded core fixes.
Routine drift reconciliation and test failure diagnosis do not require another
product decision. A change to the safety or user-choice rules does.

## 2. The user experience to deliver

One operation has one stable task surface from Preparing through completion. Routine
work starts from one command with accepted defaults; slow preparation becomes
visible and cancellable. Options, conflicts, and material risks use the existing
decision owner. Progress does not steal the operating-system foreground.

| Situation | Human feedback | Safe actions / required behavior |
|---|---|---|
| Routine Copy/Move | Requested verb and source/destination; current item and useful counters | F5/F6 start without a generic confirmation. Shift+F5/F6 deliberately open the existing options surface. |
| Slow preparation or provider | `Preparing…` or `Waiting for the device/server…` with the affected endpoint | Cancel remains responsive. No recursive scan, mutation, cut-list consumption, or breadcrumb before acceptance. |
| Another live task overlaps | Name both tasks/locations and the concrete consequence | Queue after (default), Run at the same time, Don't start. A queued Delete explicitly warns that it can remove outputs after the producer finishes. |
| Name collision or invalid name | State which item cannot be written and why | Existing exact Replace/Keep Both/Skip/Cancel choices only where valid; no automatic rename or overwrite. |
| Temporary failure proved not committed | `Could not copy/move/delete/rename…` plus useful cause and attempt count | Explicit Retry, Skip, Cancel. Each Retry revalidates the same intent/authority. Never silently turn a second Retry into Skip. |
| Server may have committed | `Final state uncertain` and `The server may have completed this operation` | Inspect/refresh/check the destination. Stop dependent destructive work. Never offer a replay as if it were safe. |
| Move copied but kept source | `Copied — source kept`, reason, and both locations | Inspect either side. Do not report `Moved`, consume a new cut list, or offer a generic source-delete shortcut. |
| Some items fail or are skipped | Counts distinguish completed, failed, skipped, canceled/unattempted, and uncertain | Failed/retained locations and reasons are reachable. A parent/aggregate cannot say all succeeded. |
| Cancel | `Stopping…`, then what completed and what remains | Stop new primary work, drain bounded calls and exact cleanup. Repeated Cancel does not restart cleanup or erase known success. |
| Cleanup fails after success | `Copied/Moved; cleanup item retained` with available location and explanation | Preserve primary truth; no deletion by guessed temporary name. |
| Popup cannot initialize graphics | Plain localized text explaining whether work continues or needs a decision | Native accessible Cancel All and Close. No unrendered conflict is accepted. |
| Possible `.rs_*` artifact | `Possible interrupted-operation artifact — name only` | Read/open/edit/preview/copy/export stay available. Manager-owned mutation has one exact-object Cancel-default warning. Properties explains known and unavailable facts. |
| Application close with live work | One aggregate choice: keep open, or cancel operations and exit | Keep open is default. Accepted exit visibly drains; UI stays responsive and provider lifetime stays owned. |

Result text answers six questions: requested action; definite changes; possible
changes; what remains and where; why work stopped/changed; safe next action. Primary
text uses familiar verbs. HRESULT, strategy, identity, phase, and receipt details
belong in expandable diagnostics. Localize all new strings and preserve the maintained
English/cs-CZ/fr-FR/ja-JP/sk-SK resources, including positional placeholders.

Escape cancels the currently actionable decision; it never means Cancel All.
Outside a decision it hides the popup. Caption Close hides without resolving work.
`cmd/app/showFileOperations` / Ctrl+Shift+J restores the existing surface. At most
one actionable decision owns focus; attention remains discoverable when another
application has foreground. In the minimal failure surface Escape only hides because
no ordinary conflict actions are rendered there.

Before discovery closes, show exact discovered/processed counters and indeterminate
task progress. A supported estimate says `still discovering` and may change in
either direction; hide it when paused/stalled or sampling is insufficient. Closed
totals may show percentages. A compatible closed-total cohort may show **Known work**;
open totals are excluded and counted separately. Do not average percentages or mix
bytes/items. The taskbar remains indeterminate while any included total is open.

Cover pane commands, captured clipboard Paste, internal/external drag/drop, Find,
Compare synchronization, F2, Batch Rename/Change Case, F7, Recycle/Permanent Delete,
and archive delete-after-success. Existing preview/editor/Run consent is reused.
Permanent Delete always has one Cancel-default confirmation; Recycle escalation is
exact-item and never Apply-to-all. External drag/drop reports Move only after proof.
Navigation while preparing/running cannot change the captured provider or paths.

## 3. Reliability rules and existing owners

The linked review's proposed four-state mutation enum and extra scheduling phases
are explanatory models. Reuse the existing typed model and preparation boundary;
do not add a second outcome enum or lifecycle just to match that prose.

| Rule | Required guarantee | Existing owner / durable spec |
|---|---|---|
| FO-LIFE-01 | Capture intent, prepare off the UI thread, obtain required consent, consume the exact clipboard sequence once, then release execution. No mutation on admission/cancel failure. | `FolderWindow.FileOperations.cpp`, `State.cpp`, `State.Runtime.cpp`; `Specs/FileSystem/FileSystem_FileOperations.md`. |
| FO-AUTH-01 | Provider capabilities describe executable routes. Mutation authority, comparable identity, and cheap overlap evidence are distinct. Missing/zero identity is never fabricated; HRESULT 996 is not automatically Unsupported. | `Common/PlugInterfaces/FileSystem.h`, `Common/FileSystemRouteContract.*`, `FileSystemRouteProviderBase.h`; `Specs/Plugins/Plugins_VirtualFileSystem.md`. |
| FO-ITEM-01 | Exactly one validated immutable terminal result per selected item, including failed preparation and partial trees. Unknown commit stays unknown. Null destructive receipt and impossible receipt combinations cannot become success. | `FolderWindow.FileOperationsInternal.h` (`FileOperationItemResult`, `SourceItemResultBuilder`), `State.cpp` finalizer; File Operations/VFS specs. |
| FO-PUBLISH-01 | Clean up only exact objects created/owned by this operation. Overwrite and Managed Move use staged owned publication. Ordinary new-name Local Copy may retain its exclusive final-leaf object; incomplete content is never complete. | Existing `PublicationTransaction` in `FolderWindow.FileOperations.State.cpp` and Local/remote writer adapters; `Specs/Core/Core_FileSystemBridge.md` and provider specs. |
| FO-MOVE-01 | Delete source only after exact destination commit, source stability, required verification, and conditional cleanup. Native Move and Managed Move keep distinct proof paths. Weak/explicit Verified routes require content proof; size/time equality is not verification. | Existing typed transfer plan/receipts and bridge; File Operations/VFS/provider specs. |
| FO-DELETE-01 | Preserve exact selected real root, canonical virtual-folder boundary, or fixed object/version selection. No followed link/mount, sibling prefix, replacement root, or guessed-generation cleanup. | Existing provider Delete adapters and bindings; File Operations and provider specs. |
| FO-DECISION-01 | One immutable conflict/action snapshot; changed occupants invalidate destructive decisions. Retry requires proved no-commit and current exact authority, never an error-code guess. | `ConflictArbiter` in `FolderWindow.FileOperationsInternal.h`, `BuildConflictActionPolicy`, provider callbacks and typed receipts in `FolderWindow.FileOperations.State.cpp`. |
| FO-OVERLAP-01 | Publish prepared scopes atomically with same-host comparison so simultaneously preparing tasks cannot both miss each other. Queue/Run consent is scheduling/consequence consent, never object authority. | Runtime admission, `MutationInterlockScope` / `ActiveMutationInterlock` in `FolderWindow.FileOperationsInternal.h`, `State.Queue.cpp`, and the exact live-output receipt index; File Operations spec. |
| FO-DISCOVERY-01 | One bounded iterative traversal streams safe work; no silent incomplete size/list/tree success. Backpressure or JIT when the queue fills. No failure merely because aggregate work exceeds an in-memory queue budget. | Host bridge, Local/cloud/device walkers, `Common/FileOperationTraversalPolicy.h`; File Operations/bridge/provider specs. |
| FO-TRUTH-01 / FO-UX-01 | Preserve independent publication/source/verification/cleanup truth and render section 2. History, breadcrumbs, name hints, and paths grant no mutation/recovery authority. | Diagnostics, popup, Issues, Properties, artifact classifier; `UI_FileOperationsPopup.md`, `UI_FolderWindow.md`, File Operations spec. |
| FO-LINK-01 / FO-RENAME-01 / FO-NAME-01 | Preserve literal links; one typed acyclic RenamePlan; provider-scoped names and collision policy at every name ingress. No Batch Rename journal/replay. | Local/bridge link adapters, Batch Rename/Change Case, route name contract; File Operations, Batch Rename and FolderView specs. |

Real-container Delete may encounter new descendants inside the exact bound root;
a changing writer can cause a bounded residual failure. S3 virtual-folder Delete
uses the just-observed current-key ETag, bounded re-list/convergence, and no VersionId
for ordinary current visibility; exact historical-version cleanup remains separate.
A complete empty listing proves that instant, not immunity from a later writer.
Graph ordinary Delete is Recycle by admitted stable item ID; unrelated revisions of
that same ID do not require a root ETag, and Permanent Delete is not invented.

While two same-host tasks are live, an exact newly published output without covering
Run consent parks destructive work for Skip, Queue, explicit consequence-named
destruction, or Cancel. Queue ends concurrency and later uses ordinary membership;
only Skip promises preservation. Entries end with the live relation. No history or
cross-process coordination reconstructs protection. Run promises admitted concurrent
tasks, with any narrow provider session serialization honestly reported.

## 4. Remaining implementation steps

All paths below are repository-relative. Scope includes each named implementation,
its direct callers when required, focused tests, resources, and the owning specs.
Use the existing C++/WIL, callback/payload, error, async, and performance guidelines.
Read `Specs/Core/Core_SharedHelpers.md` before adding a helper. No raw owning COM or
Win32 handles, detached unowned workers, catch-all exceptions, or allocation recovery.

### C0 — Prove the repaired invariants across the reachable routes

The seven original findings are repaired or tracked, rather than seven new projects:
Google Drive mutation retries and depth-64 traversal, simultaneous preparation,
contradictory plan status, provider-native recursive walkers, fake MTP conditional
deletion, and exception/handle ownership. `20c64170` contains the main corrections;
`e96e0226` completes the overlap implementation. This rewrite removes status drift.

Create one evidence matrix in a normal `Specs/TestRuns/` archive, with rows for
Local/SMB, Curl, S3, Graph, Google Drive, MTP, and relevant Dummy/bridge adapters.
Separate host discovery, provider-native copy/merge, deletion, and size traversal.
For each row record the reachable method, test ID, cancellation bound/mechanism,
commit/identity evidence, and result. Do not mark all recursion removed from a host-only test.

### C0 current evidence and next work — 2026-09-06

The lookup admission and connection-cache repairs are implemented, but C0 remains
**VERIFY**. The latest broad run is still red; these focused results cannot close it.

The first row is the latest checkpoint. Earlier rows and chronological notes
retain failed/superseded evidence; their historical assertion counts and proposed
experiments do not override the current command/count or next action below.

| Latest and historical evidence | Result |
|---|---|
| [Exclusive fixture ownership](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_exclusive_lookup_release/README.md), [broad Release and exact TCP pairs](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_exclusive_broad_release/README.md) | Lookup 273/273 and CRUD 3/3 in both builds; policy 197/197. Debug build is warning-free; Release retains one MSB3061 shared `brotlidec.dll` Clean warning. Broad 15/1/1, Copy 134/142: four source-control 10048 failures and zero deep-Move publications. Both wide copies pass with no pending retirements. Each failed numeric pair later has a PID-zero TIME_WAIT row; the sequential observation does not attribute ownership or prove the cause. Five digest-bound archives retain the evidence and tested patch. |
| [Actual socket trace](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_socket_trace_lookup_release/README.md), [broad Release](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_socket_trace_broad_release/README.md) | Release lookup 271/271, CRUD 3/3, policy 197/197; zero-warning full solution. Broad 15/1/1, Copy 132/142: five 10048 attempts produce seven numeric socket rows, identifying three source-control failures and two other-port/data-phase failures. Both wide copies publish all 4,097 files, but serial destination retires one pending passive listener and fails the unchanged fixture guard. Diagnostic-only checkpoint; Debug has not been rebuilt for this delta. |
| [Fresh-loopback connection diagnosis](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_connect_probe_lookup_debug/README.md), [broad Release](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_connect_probe_broad_release/README.md) | Lookup 252/252 and CRUD 3/3 in both builds; zero-warning solutions share one source snapshot; policy 197/197. Broad Release remains 15/1/1, Copy 132/142: all 12 implicit/explicit probes succeed after six 10048 failures. No evidence for a production explicit-bind workaround. Prior broad Debug timeout remains open. Five archives retain exact results and digest-bound metrics. |
| [Listing allocation optimization](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_listing_lookup_debug/README.md) | Lookup 249/249 and CRUD 3/3 in both configurations; both solution builds warning-free, same source snapshot; policy 197/197. Fixed Debug listing witness 918,520 → 589,488 us for identical bytes. Broad Debug/Release remain 15/1/1; Copy 141/142 and 133/142 respectively. Seven archives preserve baseline and failures. Fixture-only improvement, not a production/transport repair. |
| [Split-handle lookup](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_split_handles_focused/README.md) | Debug 201/201; four pure/mixed Unix/DOS probes prove connection reuse, complete metadata and bounded lookup storage. All eight cancellation/reborrow routes pass. |
| [Measured C0 baseline](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_copy_phase_accounting/README.md) | Debug 15 passed, 1 failed, 1 not reached. Copy 141/142; all 46 scenarios finish. Parallel wide Copy passes; serial publishes 3,210/4,097 before timeout. Delete 127/127, size 85/85, Move preflight 97/97 pass. |
| [Bounded framing and final cross-configuration checks](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_framing_lookup_debug/README.md) | Lookup 233/233, CRUD 3/3 and zero-warning test-enabled solution builds in Debug/Release. Broad Debug remains 15/1/1: serial wide Copy publishes 3,675/4,097 before timeout. Broad Release remains 15/1/1: wide copies complete, but deep-tree cases fail (Copy 131/142). Same-byte isolated Release Copy then passes 142/142; it does not replace the broad failure. All nine runs, including initial count/warning failures, are archived. |
| [Numeric connection-failure diagnosis](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_connect_diagnostics_lookup_debug/README.md) | Final lookup 237/237 and CRUD 3/3 in Debug/Release, both final solution builds warning-free; policy 197/197. Each refusal control captures cursor 1 / size-probe 3 / quote 1 with Windows 10061. Broad Release remains 15/1/1, Copy 130/142: eight cursor failures carry Windows 10048. Six runs preserve the initial control failure and broad failure; the earlier MSB3061 build warning and denied overlapping build are retained. |
| [Rejected loopback-address candidate](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_loopback_topology_broad_release/README.md) | Two broad Release runs: 15/1/1, Copy 134/142, four 10048 failures each; all 46 scenarios finish and both wide copies publish 4,097 files. Candidate reverted. Exact patch/receipt, two owned socket samples and rejected comma-filter invocation retained. Policy 197/197; spec inventory zero blockers. No repair or throughput claim. |
| [Configured providers and host policy](../../TestRuns/4cb089111a23/CompareDirectories/2026-09-05_c0_configured_remote_host_policy/README.md) | Debug 11 pass / 1 fail / 0 skip: seven host/security controls and live FTP, SFTP, SCP-profile SFTP listing, S3 comparisons pass. IMAP fails `0x80070005` at the unchanged automation TLS policy. Only five approved isolated profiles have Hello off; normal settings unchanged. Zero-warning build. Earlier missing-host and warning-bearing evidence remains linked in the archive. |
| [Pure/headless control](../../TestRuns/4cb089111a23/CompareDirectories/2026-09-05_c0_compare_host_headless_final/README.md) | Debug 2/2, zero skips, zero-warning build: 21 selection checks, 2 required-host failure-classification checks and ordinary `unique` comparison. Same source snapshot as the hosted run. |
| [Initial pool candidate](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_entry_pool_focused/README.md), [mixed-workload failure](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_entry_pool_mixed_red/README.md) | Isolated lookups reused connections, but alternating blocking/multi APIs destroyed the private transfer cache. These superseded runs stay archived. |

Split-handle implementation source snapshot:
`00de4818d995672d276fa3a3a66458c7735d2159074102bd2484bd0f96d8071a`.
Both split-handle builds are warning-free. Their
[wide baseline](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_split_handles_wide_red/README.md)
is preserved. The new measurement snapshot is
`f915a122a1f7811e1f1805a8f3a88b60d7d07f5f1b6a87c72695712839d3059d`.
Its two builds are also warning-free. Correct connection accounting makes parallel
Copy pass (4,097 files, 115.64 s, 13 listener creations), but serial still times out
(3,210 files, 120.01 s, 4 creations). The throughput failure is not connection churn.

**Next: diagnose broad Release deep-tree refusals and remove the measured Debug
wide-copy bottleneck.** The Compare host repair is validated in focused Debug;
IMAP remains a separate TLS prerequisite. Preserve
the 4,097-file corpus, 120-second deadline, 64-creation cap,
full-listing/collision validation, source-size proof and exact contents. Record
per-endpoint LIST/RETR/STOR first-use/reuse counts. Measure repeated full-parent
lookup/parsing before optimizing; do not restore whole-directory caches or invent
source survival after failed admission.

The measured candidate adds accepted LIST/NLST/RETR/STOR first-use/reuse accounting
and bounded file-phase/list-generation timings. Serial staging-name and promotion
phases consume 52.26 s and 54.07 s; contained fixture listing generation/mutex wait
consumes 49.30 s across 6,423 LISTs. Do not add these nested timings together.
Optimize one measured cost at a time without relaxing complete-list validation;
keep fixture improvement separate from production/real-server claims.

Implemented after `809daa66`: bounded block scanning/copying keeps the existing
parser, buffers, pool, checkpoint and successful-EOF rules. The fresh unchanged
lookup baseline passes 201/201; final Debug/Release pass 233/233 with 32 new
framing checks. Host exact-count guards now match lookup 233 / Copy 142 and report
expected/observed counts. Initial C5246/stale-count failures are preserved, not
reclassified. Both final builds bind snapshot
`673426f0da99bcb2f72a90557ff98fa35f8d7574212dc5fbf31bbae243ac4530`.

Fixed-list timing samples vary; do not claim consistent throughput improvement.
Debug parallel wide Copy takes 107.38 s; serial times out with 3,675 publications,
including 53.06 s of nested fake listing generation/mutex wait. Release wide Copy
passes at 7.80/23.16 s, but subsequent depth-128 Move and depth-129 refusal cases
fail with partial-copy/connection-refused HRESULTs. The same-byte isolated Copy
passes all 142 checks. Capture the failing Curl/OS error and owned route/session
lifecycle, repair the cause, then repeat the predecessor cluster and broad C0.
Machine-wide TIME_WAIT observations are diagnostic only, not proof of exhaustion.
No network settings, test limits or fake-server generation changed in this step.

Implemented after `4a42e3cb`: capture numeric Curl/OS error and connection metadata
before cursor/probe/quote handle reset, without URLs, credentials or changed retry
policy. Add four deterministic checks using an exclusively bound, non-listening
loopback endpoint (237 lookup checks total). The refusal-only fixture binding does
not tune production/listening sockets or address the wide-throughput failure.
Validate the diagnostic oracle, then reproduce the broad Release/predecessor order
and use the actual failing transport evidence to choose a causal repair.

Initial diagnostic evidence: Release broad C0 again fails (15 pass / 1 fail / 1
not reached; Copy 130/142). All eight captured failures are `CURLE_COULDNT_CONNECT`
with Windows `WSAEADDRINUSE` (10048), not `WSAECONNREFUSED` (10061). This establishes
local address/port allocation failure on the affected transport, but does not by
itself prove which endpoint lifecycle consumes the resource. The new refusal
control initially fails its three operation checks with a 500 ms connect limit;
an independent owned bound-only socket probe observes 10061 after 2,064 ms,
consistent with the control timing out first. Correcting only that control to
5 s connect / 7 s operation makes all 237 checks pass in both configurations.
The metric now retains actual HRESULTs, and both failed runs remain archived.
Do not change production/traversal timeouts, network settings, corpus or pool policy
to conceal the broad failure. Reported easy-handle ports can describe the control
connection or be unavailable, not necessarily identify the failed data socket.

Final diagnostic builds bind snapshot
`2afb2008ab09ed5d40e950d4c6020f36f04e328c52663586e86e56c3ce70e428`.
Debug receipt `2b2adacb8a58d3313c4722645f5c09a568baac2b7f0e3ef05d4c0fb106b5a9ad`
and Release receipt `10e81ea099b865401a5a72d83377a42534d8087a3dc6011a48c9cbabe17eaf52`
are test-enabled full-solution, zero-warning builds. Each exact lookup captures
all five expected connection failures, with three read attempts and one namespace
attempt; both basic CRUD runs pass. These are focused diagnostic proofs, not broad
or final-HEAD qualification. The first Release build's shared-DLL Clean warning
is preserved; a later warning-free rebuild does not close C6's ownership issue.
An overlapping Debug build was rejected by the shared dependency lock, then retried
only after Release completed. Normal settings, TLS/SSH and live mutation authority
remain unchanged. The archive contains the exact next commands and provenance.

Rejected after `a46a5a57`: distinct `127.0.0.2` source / `127.0.0.3` destination
control/passive listeners did not cure the failure. Two broad Release runs finish
all 46 scenarios at 134/142 checks, with four depth-128 Move failures and four
captured 10048 rows each; depth-129 refusal cases pass. Both wide copies publish
all 4,097 files. Source, corpus, order, deadlines and passive-listener cap are
otherwise unchanged. The candidate is reverted, including diagnostic control-port
values in `.Begin`; its exact patch, source-byte hashes, receipt and failures are
archived. Neither a repair nor a throughput improvement is claimed.

The actual test's owned socket samples confirm client `127.0.0.1` to the two
server addresses. The first sample excludes wildcard/BOUND states. The second
captures 1,092 owned rows, including 362 BOUND, during serial-wide Copy; this is
not an instantaneous lifecycle trace and does not prove a leak/exhaustion.
Both sampling intervals overlap runtime measurements. Include BOUND states in
future process-scoped sampling, bind PID/start/executable, and never attribute
PID-0 TIME_WAIT rows from machine-wide totals to the test.

The archived four-case comma-filter command is invalid for FileOps: its native
phase selector rejects it before cases run, despite the shared matcher supporting
comma-separated names. The failed invocation is retained, not counted as predecessor
coverage. Use the working `-CaseFilter C0_` broad selection until that selector has
a tested, fail-closed multi-case implementation; do not split the predecessor
sequence into separate processes or claim an unexecuted cluster. Both subsequent
listing-candidate builds restore the original `127.0.0.1` topology.

Implemented after `d63fbe53`: borrowed parent/leaf views and direct-output formatting
remove temporary fixture strings. Twelve exact-byte checks pass against the old
generator and the candidate; nine compile-time slice/self-prefix checks also pass.
Full scans, lock scope, ordering, unknown/overridden sizes and malformed overrides
are unchanged. No cache, topology, corpus, deadline or listener-cap relaxation.
The fixed Debug witness drops from 918,520 to 589,488 us for 64 complete renders /
14,159,232 bytes. Candidate Release takes 23,748 us; no same-build Release generator
baseline was collected, so do not infer a Release speedup from Debug timings.

Both candidate builds bind snapshot
`80931b760b2b6757c44172031d2ed2f143530d37d6497481791b2d0e1b0b576d`:
Debug receipt `15c07d7e8de2a32e9b3d33023e48c93c858e8685cec91686004e1f9e9b8ef21e`,
Release receipt `ad73628fbd5c9796ee1e377429641c52ba0d91c37010b8c097b2dd510500f967`.
Lookup 249/249, CRUD 3/3 and policy 197/197 pass. Broad Debug is still red:
parallel Copy completes in 71.42 s, serial times out at 3,828/4,097. Serial
staging-name/promotion phases consume 51.06/53.42 s, including 38.62 s of nested
fixture generation/mutex wait. Broad Release completes both wide copies in
7.26/22.20 s but has five cursor 10048 failures: four depth-128 Moves and one
depth-129 refusal. Copy is 133/142. All 46 scenarios finish in both builds.

Next, reduce the remaining measured full-parent work without caching/truncation;
keep generation and production parsing comparisons separate. Diagnose the actual
Release socket lifecycle, not machine-wide counts or an isolated retry. The
archive preserves the exact tested delta before these evidence-only plan updates;
new runs after this checkpoint/commit need fresh build attestation. No C row closes.

Current diagnostic step after `dbc75c7d`: observe a fresh wildcard/ephemeral bind
immediately after each captured 10048 failure, only in test-enabled builds. One
new control brings lookup to 250 checks and proves allocation plus preservation
of the caller's Winsock error. This does not connect, retry, change socket/network
policy or alter operation HRESULTs. Read-only machine inspection found all five
concrete TCP templates at dynamic range 49152/16384 and auto-reuse range 0/0;
the Automatic template has unavailable values. That is context, not exhaustion
proof.

This step is now verified: [Debug lookup](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_allocation_probe_lookup_debug/README.md)
and [Release lookup](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_allocation_probe_lookup_release/README.md)
are 250/250; unchanged CRUD is 3/3 in each build. Both full-solution test-enabled
builds have zero warnings/errors and share source snapshot
`3a6f609f3361388a361e02c57d3f1dea590403b3ddbb414f3aa34908be75a74d`.
Source/localization policy is 197/197 with no skips. The
[broad Release repeat](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_allocation_probe_broad_release/README.md)
remains 15 passed / 1 failed / 1 skipped, Copy 134/142 across all 46 scenarios.
Four depth-128 Move failures report 10048 and zero publications; all four same-thread
fresh binds immediately afterward succeed (56–624 us). Wide serial/parallel copies
both publish all 4,097 files; depth-129 refusal checks pass. This observation does
not prove the failed tuple, exclude transient pressure, or establish a leak. Broad
Debug was not repeated; its prior timeout remains open. All build/test handles are
terminal. Next distinguish implicit connect from explicit binding using bounded,
owned-loopback diagnostics; do not change system settings or mutation/retry policy.
Five archives preserve exact results and digest-bound metrics; Release lookup also
contains the tested delta before these evidence-only plan/attribute changes.

Next diagnostic step after `70d41c23`: add two owned-loopback connection controls
(lookup now expects 252). On each test-enabled 10048 failure, compare implicit
connect and explicit wildcard bind/connect using fresh listeners, no protocol
traffic, one 100 ms pending-connect wait per probe and no retry or system changes.
Capture failure stage, mode, local port and status. These are not the failed peer
or tuple; preserve that limitation and repeat broad Release before selecting a fix.

**Latest verified checkpoint (2026-09-06):** both builds pass lookup 252/252 and
CRUD 3/3, policy 197/197 with no skips, zero-warning full-solution receipts at
snapshot `71cf0be399fa0a06a45529a4a01eb4ddd39c6d22d825b8925a6a46efb2b7b926`.
All build/test handles are terminal. Broad Release remains 15/1/1 and Copy 132/142
across all 46 scenarios. Six captured 10048 failures each have a successful fresh
bind and successful implicit/explicit connects (12/12); every connection probe
reports `connected`. The result does not justify changing the production bind
policy. One depth-128 Move publishes a destination and returns partial-copy;
three others publish none. Two serial depth-129 Copy cases get connection refusal
instead of the depth error. Move's `truthful=false` here is the success-only
receipt oracle failing, not independent proof of a dishonest partial-result receipt.
Wide copies still publish and byte-check all 4,097 files. Broad Debug was not rerun;
its previous timeout remains open. Neither old failure is superseded by controls.

**Current checkpoint after `2283f0cf`:** exclusive loopback ownership and both
negative-bind controls pass in Debug/Release. This fixture correction stays, but
is rejected as a sufficient explanation/repair of the broad failure. Four failed
source-control pairs still produce 10048; later exact-pair TCP queries find
TIME_WAIT for every pair. The 74-second sequential observation is not an atomic
snapshot or proof of process ownership; PID-zero rows are never attributed to
the test process. Fresh bind/implicit/explicit probes continue to succeed.
The prior pending-retirement failure is preserved; its absence in this candidate
does not establish a causal fix. No client/retry/machine setting was changed.

**Next bounded production step:** `CurlDirectoryCursor::Next` currently yields a
row as soon as a chunk is available, before draining the LIST terminator; its
`Impl` also holds the pool borrow until cursor destruction even after completion.
Each traversal frame owns a cursor, so descent retains parent connections and
deep preflight/copy creates avoidable connection pressure. Reuse the current
bounded chunk/pause machinery to finish a small listing before yielding, then
detach, clear callback state and return a completed borrow while retaining only
buffered entries. Large listings must still pause with bounded memory; do not
drain an unbounded directory into RAM or add a new walker/spool framework.
First prove nested small cursors reuse a connection while the parent's remaining
entries remain intact. Cover empty/fragmented/multi-chunk listings, bad final
reply, cancellation, teardown and reborrow; do not weaken the full-LIST admission
checks in `GetEntryInfo`. Then rerun exact lookup, broad C0 and CRUD in both builds
with unchanged workload/deadlines, retaining failures and connection counts.
This is a code-supported candidate, not yet an implemented or verified repair.

Candidate after `96c5c5f8`: the existing chunk buffer now supports bounded
read-ahead; completed/failed/canceled cursors return their detached borrow before
destruction. Eleven nested/empty/wide/failure/cancel/reborrow controls bring exact
lookup to 284 checks, and broad scenarios emit authenticated-session counts.
Source review and Debug/Release focused/broad validation are pending; the previous
red gates stay authoritative until actual replacement results are archived.

The latest Release lookup archive retains the tested four-file delta and hashes
before final documentation/attribute changes. Debug shares the same source
snapshot. Source/spec/index are frozen during runs; subsequent code work needs
a fresh compatible full-solution receipt. Broad Debug was not repeated for the
fixture-only correction; its previous wide timeout remains open.

Strict lookup retains only the current row and requested candidate. The existing
pool now keeps a cursor multi cache and distinct cursor/blocking easy handles per
exclusive entry, resetting both callbacks before idle return. Pure and mixed
controls prove this small ownership fix; they do not prove whole-operation memory
or real FTP/SSH server quotas. Source/localization guards pass 197/197; spec inventory
has zero blockers. Older continuation details and run links are preserved in
[the prior checkpoint](../../TestRuns/4cb089111a23/FileOps/2026-09-05_c0_curl_split_handles_wide_red/prior-checkpoint.md).

Continue in this order, committing each coherent code/test/spec/evidence step:

0. Validate the available `FileOpsSelfTest FTP`, `SFTP`, `SCP`, `IMAP`, and `S3`
   profiles on this machine. Seed only their reviewed non-secret definitions into
   each governed run's isolated settings; names alone do not import normal app
   profiles. First run the five read-only `remote_file_*` Compare smoke cases,
   then use uniquely owned remote children for separately authorized mutation
   witnesses. Saved-secret checks and fake loopback tests are not live coverage.
   Preserve per-provider failures/skips and record SCP's SFTP listing route.
   **Host repair validated after `190847e3`:** declared host-dependent selections
   now use the existing owned, non-activating application host; pure engine/UNC-only
   selections remain headless. Missing required host errors fail instead of skip.
   Both final runs bind snapshot
   `3315b326efc7436b78c61137511c6b493e1a4d608ada393eb8b61bb89d084fed`.
   FTP, SFTP, SCP-profile SFTP listing and S3 read-only comparisons passed twice.
   **IMAP [blocked]:** its saved `ignoreSslTrust=true`, absent per-profile TLS
   acknowledgement and absent automation TLS allowance are rejected by the existing
   host policy before network dispatch. Hello permission does not cover TLS.
   An optional question asks whether to enable certificate verification in the
   isolated IMAP copy; it is unanswered at this checkpoint. Do not infer a TLS/SSH
   bypass or modify normal settings. Keep all five results visible in fresh runs.
   Seven host/Hello/OAuth/cloud-gate controls pass; source/resource policy is
   197/197, zero skips. Startup metrics are archived, but hosted WM_CREATE timings
   include selftest work and are not pure startup/remote-throughput measurements.
   The archive records exact roots, scoped seeding, normal-settings hash checks,
   commands, receipts and failed predecessors. Do not request the already-granted
   Hello permission again. Current Release/broad qualification remains open.
1. Diagnose and repair the broad Release deep-tree failure with bounded, secret-free
   Curl result/OS-error and session-lifecycle evidence. Use `-CaseFilter C0_` to keep
   the predecessors and Copy together; the documented comma-list is rejected by
   FileOps and must not be used until its selector is repaired and tested. Repeat
   broad C0 after a causal correction, not just a successful isolated retry.
   Separately reduce the measured fake listing serialization cost, with literal-byte
   oracles proving unchanged rows/faults; distinguish fixture speed from production.
   Complete Debug wide Copy; rerun `C0_CurlEntryLookupTruth` (273), `C0_` including Copy
   (142 assertions / 46 scenarios), and `R0fCurl_FakeFtpReadWriteCreateDelete` in both builds.
   Use the governed runner with unique run IDs. Do not use `-SkipBuild` with a
   project-only receipt. Preserve/restore `RSBuildEnableTests` for Release repeats.
2. Freeze source/spec/index during runs, preserve failures and digest-bound metrics,
   then run source/localization, spec inventory, explicit/archive-inventory checks
   and patch hygiene before committing. Both broad C0 configurations still lack a green gate.
3. Bound the O(files) Move commitment map without losing membership/size proof or
   imposing a 4,096-file aggregate cap. Address IMAP buffering, real FTP/SSH quotas
   and opaque transport memory; finish the provider/route/walker and host Rename matrix.
4. Continue C1–C6. The four deferred extras remain HOLD. Exact Fresh Full, inherited
   I17, non-test Release SQLite staging and shared DLL Clean ownership still block Done.

Raw evidence stays beneath the initialized `C:/RedSalamander.Perf`. The untracked
repository-root `last_run/` is user-owned: never inspect, clean or stage it.
Do not delete retained runs to hide disk-audit findings. No push or destructive
cleanup is implied. Git staging/commit require the normal approved escalation
because this worktree's Git metadata lives outside the workspace.

Google Drive limits transient HTTP retries to `AuthorizedRequestRetry::ReadOnly`
and its fake-drive test copies/counts 80 levels. Retain those guards and the passing
host `C0_GDriveCommittedCopyResponseLost` / `C0_GDriveCommittedDeleteResponseLost`
witnesses: one mutation request, an actual backend side effect, no unsafe Retry,
and an uncertain result. A test that only returns 503 before mutation is insufficient.
Apply the same safety contract to reachable Graph/S3/Curl/MTP mutation adapters;
backend-idempotent retry is allowed only with the same proved identity/precondition.
Read retries remain bounded and cancellable. Cancellation after request dispatch
does not by itself prove the server did nothing.

Cover exact MTP PUID delete under replacement, stale occupancy and failed original
delete, Local foreign-occupant survival, S3 per-key conditions, and simultaneous
preparation in both publication orders. Include deep/wide traversal, unreadable or
vanishing descendants, late writers, out-of-scope entries, and overflow/counter truth.
Existing provider loops that are bounded/adapters may remain; remove reachable
unsafe recursion/cutoffs with a behavioral witness, not a repository-wide rewrite.

FOS-01–04 and FOS-06–12 map to B1/B2 and this matrix; FOS-05's remaining responsiveness
work is C1. A failed row stays open with its exact reason. Conditional R5 isolation
is considered only if a measured residual route cannot be contained safely in process.

### C1 — Move provider waits behind visible preparation and owned shutdown

At the review commit, `AdmitOperation` still calls `QueryChildNameContract` once
per transfer leaf (`FolderWindow.FileOperations.cpp:2581`) and binds Permanent Delete
roots (`:2623`) before `StartOperation`. Inline Rename also queries providers
(`:2748–2762`). Later worker Preparing tests do not prove those command paths responsive.
`FileOperationState::Shutdown` cancels then synchronously joins workers
(`State.Runtime.cpp:1831–1855`). A provider timeout bounds a call but does not make
a window waiting for that timeout responsive.

Capture immutable intent, pane/selection epochs, retained provider ownership, and
already-cached facts on the UI thread. Preserve containment admission before any
task/card/worker creation: both endpoints' exact path-and-operation routes must be
proved `bounded` or `providerWatchdog`; reject `uncontained` without provider I/O.
Only then publish the existing task and move potential provider/alias/name/bind I/O
to worker-owned Preparing under that proved cancellation contract. Admission queries
must themselves have a proved nonblocking contract; if existing cached facts cannot
establish it, repair the route contract rather than merely moving an uncontained
adapter onto a worker. Keep validation buckets and one stable failure surface.
Audit actual F5/F6/Paste/Delete/F2/F7/Find/Compare/rename callers, including nested
artifact/overlap prompts; a supposedly cached query needs evidence that it cannot do I/O.

Before consent/release, publish prepared overlap scope under the existing atomic
coordination boundary. Pane navigation and later clipboard changes must not alter
captured intent. Cancel before acceptance leaves source, destination, cut sequence,
and Move breadcrumb untouched. Do not recursively pre-scan or read payload content.

Reuse the existing completion/reaper ownership boundary to drain cancellation
asynchronously while the UI remains alive and shows Stopping. Retain tasks,
providers, and module pins until callbacks are quiet; then drain registered payloads
and release on the correct threads. Do not detach, kill a thread, unload executing
code, or claim cancellation acknowledgement merely because a token was set.

Tests start from real command ingress with deterministic delayed name/bind/read
fixtures. Prove command return, UI heartbeat and navigation while the fixture remains
blocked, then release/cancel through its proved route mechanism. Preserve the existing
500-ms card reveal deadline and fake-provider visible cancel acknowledgement target
of at most 2 seconds (`FileSystem_FileOperations.md`, Preparing/performance sections).
The exact Local pending-reader Release witness must retain `ERROR_CANCELLED`, zero
bytes, drained ownership, and every sample plus p95 below 500,000 us under
`Testing_PerformanceValidation.md`; that bound is not a universal shutdown deadline.
Measure actual provider quiet point separately against its route-specific contract.
Check cancel-and-exit, zero rejected-route provider calls/tasks, late completion/payload
delivery, and final ownership counts. Use `R1d_PreparingLifecycle`, R0e admission,
and Commands lifetime tests as patterns, extending their earlier boundary. MTP/S3
stalled-read and shutdown evidence must name the actual reader path. Verify zero
pre-acceptance mutations and no deadlock/use-after-free.

### C2 — Make Retry safe first, then remove the one-retry limit

`IsRetryableConflictBucket` (`State.cpp:4041`) currently permits Unknown;
`allowRetry` (`:19470`) depends on that bucket/count, and `:19550` turns another
Retry into Skip. Similar limits exist in provider, inline/batch Rename, bridge,
and cleanup loops. Lifting the count check alone would increase risk.

Use existing `FileSystemItemMutationResult.outcomeKnown/mutationCommitted`, typed
item axes, and phase-aware pre-mutation facts. The conflict bucket explains the
problem; it never proves non-commit. Keep one eligibility decision in the engine
snapshot. Permit a Retry only for a proved no-commit attempt whose exact authority,
source stability, scope, and required compensation state remain valid. Unknown
publication/source/cleanup never becomes a primary replay via Skip, Cancel, or Retry.

Then replace the boolean/one-shot limit with an attempt count. Each click requests
exactly one attempt; no automatic loop, cached Retry/Apply-to-all, or silent Skip.
Revalidate just before retry and return changed conflicts to the same arbiter.
Exact cleanup retries remain bounded compensation, not a restart of primary work.

Replace `Phase9` retry-cap expectations with: several no-commit failures then
explicit success; source/destination replacement; cancel between attempts;
committed-but-error; null/inconsistent receipt; timeout/disconnect; failed cleanup;
serial/parallel and native/bridge routes. Assert UI action availability, provider
request counts, immutable result axes, and human attempt text, not source spelling.

### C3 — Copy links unchanged with no dependency graph

Local link conversion still invokes `TryRetargetPathIntoDestination`
(`Plugins/FileSystem/FileSystem.FileOps.cpp:8757`) and rewrites a mapped target;
Local deferred-link state can fail at the 4,096-entry shared queue limit (`:3229`).
The bridge repeatedly scans deferred link mappings (`State.cpp:15478`). The current
authoritative File Operations spec also describes that old semantic Preserve behavior.

Make default Preserve copy the literal no-follow link object/payload, including
relative spelling and flags, without resolving/requiring the target to exist.
Remove retarget mapping/deferred state from this route and its redundant policy
branches. Do not remove provider ABI needed by remaining callers without a caller
inventory; no new Retarget option or ABI is needed for this core change.
Unsupported representation must report a precise retained/skipped/failed outcome,
never follow the target, silently materialize content, or corrupt link bytes.

Test relative/absolute symbolic files/directories, junctions, missing/external
targets, link cycles, Keep Both destination names, cross-provider support/refusal,
partial failure/cancel, and more than 4,096 links. Literal copying may leave a link
pointing to the old absolute location; say so where relevant and do not silently
repair it. Managed Move source deletion still requires its existing exact proof.
Update Local/bridge tests and `FileSystem_FileOperations.md` in the same change.
The optional transform and its indexed graph move wholly to the later plan.

### C4 — Remove the full fallback popup; keep a usable failure surface

The attachment-failure route reaches the full painter and actions in
`FolderWindow.FileOperations.Popup.cpp:6524–6570`; `Render` returns without drawing
when D2D resources are missing (`:6162`). Commands currently expects a fallback
throughput graph, and `UI_FileOperationsPopup.md:304–309` requires it. Replace all
three together; changing only the plan cannot retire this implementation.

Retain ordinary shared composition: `Render` still owns noninteractive text, card
geometry, and hosted-control descriptors; DxUi owns its existing interactive controls,
progress/graph pixels, focus, and accessibility. This is not an all-DxUi layout rewrite.
Remove only attachment-failure action/graph/progress branches and fallback-only state.
The failure surface uses native accessible text/buttons independent of D2D/DirectWrite: a concise
operation/count/context/status explanation, explicit Cancel All, and Close. Use WIL
resource ownership. No task graph, queue/speed menu, Replace/Retry action, or second
conflict policy. Waiting tasks remain waiting; running tasks continue unless canceled.
If the native window itself cannot be created, report through the existing host error
surface and retain access to explicit cancellation; never leave an invisible decision.

Test forced hosted attachment failure and forced D2D target/resource creation
failure separately. Verify readable status, keyboard Tab/Enter/Escape, UIA names and
Invoke, high contrast/DPI, close/reopen, teardown, and that no hidden mutation action
can fire. Remove obsolete graph-fallback tests and fallback-only painter/action code;
preserve shared composition and the canonical hosted graph painter.

### C5 — Complete human feedback, including Properties errors

The latest artifact change is implemented, but `OnLoadComplete`
(`FolderWindow.ItemProperties.cpp:2017–2026`) returns on provider/JSON failure before
appending its available host explanation. `RefreshDocumentFromIo` (`:2144`) calls
the provider synchronously and reuses `_artifactExplanation`. The happy-path tests
do not establish missing-object or refresh truth.

Keep the host File Operations explanation independently renderable when ordinary
properties fail or are malformed. Run refresh I/O/classification on the existing
worker with a generation token and teardown-safe payload ownership. Each refresh
uses at most one bounded classifier query for a Possible name, zero for ordinary
names; it shows Present/Missing/Unavailable from that observation and does not
invent a claim. Only show semantic fields supplied by a valid source; absent task,
phase, publication, source, cleanup, verification, or recovery facts say Unavailable.

Extend Commands Properties tests for missing/unreadable object, malformed provider
JSON, replacement during refresh, dialog close during load/refresh, and stream
removal Cancel/accept/revalidation. Audit stream mutation responsiveness with C1.
Retain prompt-free opens/copies/external actions and exactly one warning for actual
manager-owned mutation. Never present `.rs_*` naming as proof of ownership.

For section 2, exercise routine work, disk full, permission denied, device removal,
connection/authentication loss, invalid names, conflict changes, Copy-only Move,
partial results, uncertain commit, cleanup debt, and blocked cancellation. Reuse
existing fake-provider/Commands cases. Add only missing behavior coverage. Confirm
mouse, keyboard-only, UIA, screen-reader wording, high contrast, RTL/long translations,
reduced motion, minimum width and 100/150/200% DPI. Record human observations
separately from automation; neither may be claimed without having been performed.

### C6 — Consolidate and qualify the core once

Resolve the reproduced shared-output CoreClean warnings as part of build
qualification, together with the already-recorded non-test Release SQLite staging
transition. Multiple project FileListAbsolute records claim the same plugin DLLs
(yyjson, sqlite3, libcrypto; the size repeat run additionally witnesses z.dll); the partial-tree and size archives retain the warnings and
ownership evidence. Repair and test cleanup ownership/order, not warning suppression
or manual deletion. Later warning-free builds do not explain an intermittent race.
The exclusive-listener Release build additionally witnesses MSB3061 for
`brotlidec.dll`; its complete log is retained in the 2026-09-06 lookup archive.
The configured-provider smoke build again reports three MSB3061 cleanup warnings
for `libcrypto-3-x64.dll`, `zd.dll`, and `brotlicommon.dll`; its archived build log
is another C6 witness, not a clean-build qualification.

Merge landed behavior into the owning specs in section 3. Keep provider matrices in
provider/VFS contracts, tuning constants in code/performance policy, and run logs
in `Specs/TestRuns/`. Update `Specs/NormativeConsistency.json` only after reviewing
the actual changed sections, implementation/test anchors, and source identity.
Do not bulk-stamp old global `VERIFIED_MATCH` records as freshly reviewed.

I14 owns these core fixes and inherited I17 qualification. I3
(`FileOperations_RemainingStressValidation_2026-08-04.md`) owns its five stress rows;
reuse their evidence and correct any old hard-cap expectation displaced by streaming
traversal. I12 (`Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md`) owns its
completion/reaper/focus witnesses. Coordinate shared changes and record which rows
the final gate covers; unrelated Terminal work is not absorbed or declared complete.
Remove stale “only E0 may begin” and inactive-slice statements from the WIP index.
Coordinate I0's stale Google Drive read-only milestone claims with its owner using
the current provider spec and C0 evidence; do not reimplement already delivered writes
or claim that its separate OAuth/remote-validation work is complete.

The old I17 plan remains frozen. Final FileOps, Commands, resource, spec, and Fresh
Full evidence must explicitly qualify its delivered popup/decision baseline together
with this core. No second implementation or reopened historical plan is required.

## 5. Verification and evidence

Before code work, run `git status --short` and
`git diff 58f22d0e -- Common RedSalamander Plugins Tests Tools Specs`. Reconcile
overlapping edits; leave the pre-existing untracked repository-root `last_run/`
untouched. All runtime scratch/logs use an initialized owned fixed-drive
`X:\RedSalamander.Perf` root. `.build` is only for build outputs and documented
inventory reports. Never launch two interactive suites concurrently or kill an
unrelated executable to acquire an output path.

| Purpose | Command / expected result |
|---|---|
| Debug affected build | `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` (plus affected standalone tests); 0 warnings/errors. Project builds are iteration evidence, not full-solution attestation. |
| Test-enabled Release | `$env:RSBuildEnableTests='true'; .\build.ps1 -ProjectName RedSalamander -Configuration Release`; 0 warnings/errors and a valid receipt with tests enabled. Restore the prior process environment afterward. |
| Focused Commands/FileOps | `.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Debug -FailFast -CaseFilter <exact-case> -RunId <unique-id>` (use `-Suite FileOps` for that suite, and repeat required cases with `-Configuration Release`); expected cases selected and all pass. Focused runs are not repository qualification. |
| Case discovery | The matching test-enabled `RedSalamander.exe --selftest-list-cases --commands-selftest --fileops-selftest` emits native inventory JSON. Wait for exit and capture stdout beneath the owned test root; select actual IDs and reject zero-case runs. `Get-TestInventory.ps1 -Format Json` lists harnesses/counts, not all native case names. |
| Final qualification | `.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh`; exact final implementation/test/tool state, zero failures, complete expected counts, explicit allowed environment skips. |
| Source/resource policy | `Invoke-Pester -Script '.\Tools\Tests\TestHarnessSourceContracts.Tests.ps1' -PassThru` and `ResourceLocalizationContracts.Tests.ps1`; inspect `FailedCount`, do not rely on process exit alone. |
| Spec/index hygiene | `.\Tools\Get-SpecInventory.ps1 -FailOnFindings`; zero blocking findings. |
| Evidence archives | `.\Tools\Test-TestRunArchive.ps1 -Inventory`; all required archives valid. |
| Patch hygiene | `git diff --check`; no findings. |

Use the governed runner where possible; the focused command above builds its selected
profile. Add `-SkipBuild` only with a compatible full-solution receipt and byte-identical
artifacts accepted by the runner, never on the strength of a project-only build.
If an exact native case is needed for
diagnosis, use `Start-Process -Wait -PassThru`, inspect its result JSON and selected
case count, and keep the same sandbox/serialization contract. Direct invocation of
a GUI-subsystem executable may return before the test completes. Never treat that
shell return as a passed test. Existing structural source guards remain for ownership,
ABI, and lifecycle rules; runtime fault tests prove behavior. I4 owns any broader
source-test migration.

Each row's receipt records commit/source digest, build receipt, command/filter,
run ID and archive, counts/skips, fault assertions, and measured resources. Preserve
failed evidence, classify its cause, repair, and rerun the affected scope. Final
Fresh Full follows all relevant source changes; older runs, Resume, Affected, or
isolated reruns cannot replace it. A flake is an unresolved test problem until its
cause is addressed and the required gate passes. No disabling checks to get green.

Performance evidence is proportional to the changed path: UI heartbeat/reveal and
cancel/quiet-point latency for C1; attempts/request counts for C2; first mutation,
deep/wide/link throughput and retained memory for C0/C3; redraw/UIA/fallback cost for
C4; classification/provider-call counts for C5. Preserve the accepted discovery
reservation/JIT scheduler unless a measured causal regression justifies a change.
Reuse concurrency 1/4/16 and independent-device/SMB/MTP witnesses where applicable.
Queue bounds cause backpressure, never silent omissions. Record host/provider,
topology, corpus, configuration, and before/after results. No general limiter,
disk work spool, or new benchmarking framework is required.

Latest evidence at this review is useful but does not close the plan:

- Fresh Full `20260905T051943Z-46440-7e8a5e538718446d95a4e12a5f0190a3` on
  `c775b9ed`: 2,087 total; 2,033 passed; 1 failed; 53 skipped. ViewerImgRaw scheduler
  was the sole failure; atomic replacement correction `9f29ec6d` followed, with
  50 isolated successful repetitions. A new final Fresh Full is still required.
- B4 Debug receipt `c7d889ad181a30b6cb47d1dd5eee187e3b287bd3c46e13016d438cb45b9c1d1d`:
  0 warnings/errors. Runs under `C:\RedSalamander.Perf\runs\`:
  `r6-a09-final-routine-20260905`, `r6-a09-final-surface-20260905`,
  `r6-a09-final-streams-20260905`, `r6-a09-final-contract-20260905` each 1/1;
  `r6-a09-final-phase10-20260905` 9/9. Durations: 1,415/1,175/2,261/1,045/6,500 ms.
- Test-enabled Release receipt visible after the interrupted build:
  `2b0007216d3cb81c19737209cf6165f1f7dca89a84afbd8f4e7b3d4d886c5c31`.
  Inspect receipt provenance and required focused Release results during closeout;
  the interrupted tool call itself is not test evidence.

These local receipts need the normal reviewed archive before retirement. Live
credentials/devices unavailable to automation remain explicit coverage limitations;
deterministic simulations do not prove real hardware execution. Do not bypass a
release-critical missing witness by calling it an expected skip. Where a human or
real-endpoint check is still required, identify the exact check and owner.

## 6. Definition of Done and stop boundaries

The core plan may move to Done only when every C row has exit evidence, all required
behavior is implemented once and described by its current authoritative owner,
Debug and test-enabled Release checks pass, required fault/resource/human evidence
is archived, inventories and patch hygiene pass, and one exact final Fresh Full
qualifies the core plus inherited I17 scope. The final receipt must state remaining
provider/environment limitations. Update the WIP index and move this plan and its
historical record through the existing archive process; the later-features plan
stays HOLD. Archive-only receipt/link changes still require inventory/hygiene checks.
They must not conceal unvalidated source/test/tool changes after the final gate.

Stop the affected change and explain the evidence if it would weaken identity or
source-retention proof, replay unknown mutation, silently skip data, force a second
policy owner, widen consent, claim an unenforceable cancellation guarantee, or
require destructive cleanup of existing user/recovery data. Continue independent
authorized work. Request a product choice only when the existing decisions do not
resolve a real tradeoff. Do not return for routine activation paperwork.

Keep the design small: existing typed plans and results, one conflict owner, one
publication transaction, one ordinary popup renderer, and bounded provider adapters.
No universal Undo/Resume/WAL, distributed coordinator, compatibility parser beside
typed authority, persistent overlap graph, eager whole-tree preparation, or new
framework solely to describe the current implementation. Remove a displaced owner
with its migration; do not maintain duplicate safety paths for a hypothetical future.

The historical review also noted legacy plaintext credential settings in Curl/IMAP.
That finding needs a separate current security review and credential-storage owner;
it is preserved in the later plan's routed-risk note. This plan does not clear or
migrate user secrets, and file-operation diagnostics must never expose them.
