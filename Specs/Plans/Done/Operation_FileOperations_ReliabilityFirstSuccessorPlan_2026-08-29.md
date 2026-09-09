# File Operations: reliable core and clear feedback

> **NON-NORMATIVE RETIRED WORKING-PLAN RECORD — I14.** Moved from `Specs/Plans/WIP/`
> on 2026-09-07 by product-owner decision: C1–C11 landed and Fresh Full gate #21 is
> recorded in C6; the C5 human observation pass was outstanding at the move, and the
> owner reopens this plan if that pass finds the delivered behavior invalid. **No
> remaining implementation action belongs to this file.** Priority was P1
> reliability. Owner was the agent
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
to [a later plan](../WIP/Operation_FileOperations_LaterFeatures_2026-09-05.md). They no
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
| B2 | LANDED | Discovery/progress, writable-provider containment, exact same-host Queue/Run/output guards; R0f, R4, R6-A02/A08 | B1 | `20c64170`, `e96e0226`, `31fe0b8c` plus historical receipts; live-output index correction after the 2026-09-06 review (on top of `12344657`): rename publications are indexed under their parent scope instead of degrading the host to overflow, and an overflowed task pays no lock per later item (`R4A02_RenamePublishKeepsIndexPrecise`). C0 verifies provider walker coverage and the complete overlap matrix. R6-A02 has no second UI implementation. |
| B3 | LANDED | Routine F5/F6 and explicit Shift+F5/Shift+F6 options; R6-A20 | B1 | `c775b9ed`; command/resource tests and zero-warning Debug/Release receipts exist. C5/C6 cover final integration. |
| B4 | LANDED | Name-only artifact exemptions and Properties explanation; R6-A09 | B1 | `58f22d0e`; 4 focused Commands cases, Phase10 9/9, source contracts 188/188, localization 9/9. C5 owns remaining provider-error/refresh behavior. |
| C0 | VERIFY | Review corrections and an explicit route/walker evidence matrix | B1/B2 | Focused Debug `C0_` on 2026-09-06 passed 20/20 (`20260906T132305Z-2364-47c7631340054c48b1fa9c3db25f59b7`): Local known no-commit receipts; IMAP listing 369/369 and loopback transport 15/15; lookup 298/298; Move 250/250; Copy 142/142; directory-size 85/85. Walkers finish a listing before child Open; DOS listings require a real timestamp; Move source-delete consults preflight membership. Closeout 2026-09-07: the route/walker and host Rename matrix is archived at `../../TestRuns/4cb089111a23/Continuation/2026-09-07_144220_reliability_first_gate21/README.md` against Fresh Full gate #21; the native Move commitment map is bounded (item 1 below); the IMAP summary workspace is bounded to one 200-UID fetch chunk (plugin self-test asserts the peak) while the `UID SEARCH ALL` reply is still one whole-mailbox string and libcurl transport memory is opaque, both waiting on the pending bounded-batches product question. Owner/environment items stay open: real FTP/SSH session quotas, live IMAP (TLS policy refuses the isolated profile), Google Drive live mutations. Coverage correction 2026-09-07: the three C0 receipt steps and the thirteen C0 Curl truth steps were outside every FileOps family, so no Fresh Full through gate #20 executed them (`not executed` in the archives); they now form `FileOpsFamily_C0MutationReceipts` and `FileOpsFamily_C0CurlTruth`, and gate #21 is the first Fresh Full that runs them. `C0_CurlCopyTraversalTruth` carries one open scenario (C11). |
| C1 | LANDED | Responsive ingress, Preparing, cancellation, and shutdown; R1d-OR1 / FOS-05 residual | B1/B2 | Command-ingress delay, blocked-reader, close, and lifetime tests; no UI-thread provider waits. Landed first (2026-09-06): the per-leaf destination child-name contract moved from `AdmitOperation` into Preparing (`Task::PrepareTransferDestinationNames`, RC3-9 in task-result form, Commands guard retargeted); then the inline F2 admission query moved there too (pending join filled in Preparing, `C1_InlineRenameNameRefusedOnCard`, guard extended); Batch Rename already queried and bound on its worker admission; then application exit drains cancellation off the UI thread (`FolderWindow::CancelAllFileOperationsThenClose` defers the close until the last task is reaped, `C1_ExitCloseDeferredUntilTasksQuiet`); then the Local Native Move shape probe left admission (listing-snapshot shapes, provisional Native with the prepared Managed fallback, `Task::PrepareNativeMoveShapes` qualifying on the task thread and recording per-item fallbacks, `C1_MoveShapeProbeRunsInPreparing`); last, the Permanent Delete ingress bind left admission and its confirmation moved to the card after the Preparing bind, with the post-confirmation re-check keeping continuity (`Task::RunPermanentDeleteConfirmation`, `C1_PermanentDeleteConfirmsOnCard`, the replacement-race case relocated to the re-check window). Every admission path is now free of provider I/O; the Fresh Full gate for the row is C6. |
| C2 | LANDED | Proof-gated repeated explicit Retry; R6-A07 | B1 | Existing mutation facts drive eligibility; host/provider fault tests show one request on unknown commit and repeated explicit safe retries. Landed 2026-09-07: eligibility stays the receipt classifier's proved no-commit plus the typed-bucket rule; the one-shot caps in the per-item loop, the provider-callback path and inline/batch Rename became attempt counts (each click one attempt, revalidated by the loop it re-enters, no silent Skip); the prompt carries the count ("Attempt N failed.") and keeps Retry; the Managed Move exact cleanup retry stays bounded. `Phase9_ConflictPrompt_RetryCap` now injects three proved no-commits, retries three times and completes on the fourth attempt. |
| C3 | LANDED | Literal Preserve without a retarget graph; core part of R7-A03 | B1 | Landed 2026-09-07 in two commits: the Local provider copies links literally (stored text, kind, relative flag) with its retarget helpers, deferred link queue, and Keep Both mapping removed; the bridge's sparse map, failed prefixes, deferred queue, dependency classification, held cleanups, and eight `FileOps.Bridge.Link*` counters removed. Phase12/Fairstream/plugin self-tests prove literal text, a skipped target beside a literal link, and a 4,200-junction copy with no ceiling. |
| C4 | LANDED | One hosted popup and a graphics-independent failure surface; R6-A12 | B1 | Landed 2026-09-07: host-attach or Direct2D failure switches the popup to native text + Cancel all + Close (Tab/Enter/Escape, UIA from the native classes, high contrast and DPI, hide and reopen, refused actions); the legacy no-host painting branches, fallback graph metrics, and legacy mouse routes are removed. Two Commands guards force each failure separately. |
| C5 | LANDED (automation); human pass owner-deferred at Done | Human-readable results and Properties failure paths; remaining R6-A09/UX | B3/B4 | Landed 2026-09-07: provider failure, malformed JSON, and missing object keep the File Operations explanation; refresh re-reads through the worker; close during load/refresh is safe (three Commands cases). The section 2 journey matrix in section 5 maps every situation to its automated cases and names the three causes no automation injects (mid-transfer disk full, device removal, authentication loss). The human observation pass (mouse, keyboard-only, screen reader wording, RTL/long translations, minimum width, 100/150/200% DPI) has not been performed; on 2026-09-07 the owner moved the plan to Done with that pass outstanding and will reopen it if the pass is not valid. |
| C6 | GATE RECORDED; DONE by owner decision 2026-09-07 | Consolidation, integrated tests, and core retirement; R8/E8 | C0–C5 | Shared-DLL Clean and diagnostic accounting repairs verified below; production Release transition 21/21. Fresh Full gate #21 `20260907T111154Z-110364-496172ac310a45f7b4934393bf924d4a` on `03d851fa` (2026-09-07, after C7–C10): 2,155 total; 2,089 passed; 6 failed; 60 skipped; classification: `R4A19_DiscoveryIndependentVolumes` is the environment gate (no authorized alternate-volume test root); the other five are run-window timing artifacts, not code: four Commands UI cases (`cmd_preferences_dialog_category_tree_keyboard_expand_collapse_and_child_entry`, `cmd_preferences_dialog_category_switches_do_not_churn_tree_host`, `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete`, `cmd_pane_filter_prompt_uses_dxui_surface`) ran two to six times slower than in gate #20 while the median case ratio between the two runs is 1.06, and `Phase7_ParallelCopyMoveKnobs` took 113 s against 54 s and did not observe more than one in-flight entry; each of the five passed alone twice on `03d851fa` right after the gate (one churn rerun was skipped by the harness because the desktop foreground changed during its measurement, which is the same disturbance); no case that C7–C11 touch failed, and `C0_CurlCopyTraversalTruth` passed in 214 s within its 480 s budget with the C11 bound holding; the `Phase16_Remote*` live-profile cases are skipped because the gate environment has no connection profiles. Test-enabled Release build: `-Rebuild` with tests enabled on `03d851fa`, 0 warnings, 0 errors, artifacts attested by receipt `073da4f1b9401b1c468c20d036c3c9ab4692e8d524dbd368518747d2a3cbdfed` (`tests_enabled` true). Gate #20 (`20260907T050325Z-126352-bc9b07a54595434397c2c7daa237249a` on `37957c69`, 2,152/2,084/1/67) preceded C7–C10 and is superseded. Spec inventory 0 blocking findings; archive inventory passes; Pester 188/188 and 9/9; `git diff --check` clean over every commit since `e48369bc`. Outstanding at the move to Done (owner decision 2026-09-07), each named for its owner: the human observation pass (C5), the C0 owner/environment items, the bounded-batches product question, the C11 Curl fixture bound, and the intermittent ViewerImgRaw latest-wins regression (outside this plan; see the gate record). The inherited I17 popup/decision baseline ran inside this same gate. |
| C7 | LANDED | Spec truth after C1–C6: descriptors, VFS matrix, FileOps domain sentences | C6 | Landed 2026-09-07: the VFS summary and strategy matrix state the routes and operations the plugins declare (SMB bounded, Curl/GDrive/Graph/S3 provider watchdog and admitted, one `google-drive` profile); Local declares `retargetInTree=false` under literal Preserve and `FileOps_ProviderCapabilityMatrix` pins it; FileOps spec: Known-work aggregate, native failure surface instead of a fallback painter, R0f policy in present tense. |
| C8 | LANDED | S3 reader cancel on its own thread | C7 | Landed 2026-09-07: `S3RangedFileReader` implements `IFileReaderOperationControl`, `Read` opens the per-thread options scope and checks the control before each request; scenario 4 of `RunS3StalledRequestCancelSelfTests` (RED before: no control interface) proves a stalled reader `GET` returns `ERROR_CANCELLED` within the declared bound after Cancel. |
| C9 | LANDED | MTP decision-time replace identity | C7 | Landed 2026-09-07: both temp-swap commits resolve the destination PUID before the upload and delete only that identity; the writer implements `IFileWriterExpectedReplacement` (live resolve at Set, early refusal of a gone or changed occupant); `FileSystemMtp` implements `IFileSystemAtomicWriter` for the overwrite flag, so the host now offers Replace on device destinations it used to withhold. Witnesses (RED before): `mtp_overwrite_refuses_occupant_replaced_during_upload`, `mtp_copy_overwrite_refuses_occupant_replaced_during_temp_copy`, `mtp_atomic_writer_offers_conditional_replace_only`. |
| C11 | OPEN AT DONE (needs a new owner) | Curl copy traversal: the Wide serial scenario | C7 | Exit: `C0_CurlCopyTraversalTruth` passes in a Fresh Full. Landed so far (2026-09-07): the host case's worker budget is 480 s (the export's Wide matrix, 4,096 files in eight variants, takes about four minutes; the old 180 s budget reported a timeout before the export finished), and the export's verdicts now print their detail (status, expected status, both servers, contents, cooperative cancel, fixture bound). Open: the Wide scenario's fixture bound fails intermittently (two of three direct runs, once at concurrency 1 and once at 4; the harness run passed): the destination endpoint ends with `pendingRetirements=1` and 12,290 passive-listener reuses against 12,289 data-transfer reuses, so one PASV on the destination is never followed by a data connection; source accounting is exact (3 listeners, 4,096 reuses, 4,097 payload transfers), contents and statuses are right, and every cancel scenario passes. Reproduced on `4bc5c987` before C10. Not loosened: the extra PASV needs its origin (plugin or libcurl) before the bound can change. Frozen with the plan on 2026-09-07; a follow-up needs a new active owner. |
| C10 | LANDED | Identity-pinned Permanent Delete | C9 | Landed 2026-09-07 through a narrow optional contract (`IFileSystemIdentityDelete`) instead of full object binding, which would have rerouted Copy, Rename and publication: Preparing pins each root's identity, the card re-check compares it, execution deletes only while it holds. S3 pins the version id or ETag and reuses its conditional delete; Google Drive pins the file id; Curl keeps the path delete and the card says so. Witnesses in the final phases of `R0fS3_`, `R0fGDrive_` and `R0fCurl_` (RED before). |

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

Sections C0–C6 below describe the code as reviewed on 2026-08-29 and the steps that then
landed. Once a row in section 1 reads LANDED, its section here is historical narrative with
pre-landing file and line anchors, not an instruction; the same holds for the frozen
execution record and its old status words. Sections C7–C10 are the post-closeout corrections
added on 2026-09-07 after the independent review of `1f665a75`; C11 is the pre-existing failure
that surfaced when the C0 Curl steps first ran in a family.

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

C0 remains **VERIFY**: broad Debug and Release now pass, but the remaining
provider/route/memory requirements below are not closed. No C1–C6 gate is closed.

Update 2026-09-07: C1–C5 are LANDED and Fresh Full gate #20 is recorded in C6. The
code-side C0 requirements are closed by that gate and the archived route/walker matrix
(see the C0 row and "C0 closeout" above). What keeps C0 at VERIFY are owner/environment
items only: real FTP/SSH session quotas, live IMAP (TLS policy refuses the isolated
profile), Google Drive live mutations, and the bounded-batches product question.

| Current evidence | Result / limitation |
|---|---|
| [Move-mode Debug](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_movemode_candidate_debug/README.md) and [Release](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_movemode_candidate_release/README.md) | 247/247 each; all 48 mode and two mixed-bulk scenarios pass. Exact eight-file tested patch/raw hashes, matching snapshot `049d8724b6ba200a243b870ce24bd6edc7c1b1c32925b225e823518c415afa5f`, full build receipts and logs are archived. Debug: zero warnings/errors; Release: one MSB3061 for shared `Plugins/z.dll`, zero errors. Source/resource policy 197/197, no skips. |
| [Move-mode broad Debug](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_movemode_broad_debug/README.md) and [Release](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_movemode_broad_release/README.md) | 17/0/0 each, including deep Moves, Copy 142/142 and lookup 294/294. Wide Copy completes 4,097/4,097 at concurrency 4/1 in Debug 62.673838/99.079379 s and Release 6.324826/17.585574 s. No timeout increase. Companion CRUD archives pass 3/0/0 each. These supersede the parser checkpoint for affected runtime qualification, not its shared-converter evidence. |
| [Move-mode red baseline](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_movemode_baseline_debug/README.md) | 153/241: 80 assertions expose ignored modes; eight expose the fake server's file-only rename limitation. Both repaired; six mixed-bulk assertions added. Exact failed evidence is retained. File-scenario comparison shows native-only cross-endpoint mutations 16 to zero over four requests, invalid-option commands 356 to zero over 16 requests, with unchanged valid native command counts. |
| [Debug lookup/build](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_lookup_debug/README.md) and [Release lookup/build](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_lookup_release/README.md) | Lookup 294/294 in both configurations. Full test-enabled solution builds have zero warnings/errors and share snapshot `b59e77718ac62ce6212984e1fa92efc59f2aa92283b4b9f1fd64b123fa44dbf4`. Source/resource policy: 197/197, no skips. Debug/Release receipts are retained in the archives. |
| [Broad Debug](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_broad_debug/README.md) | 17/0/0; Copy 142/142 across 46 scenarios and lookup 294/294. Parallel/serial wide Copy publish all 4,097 files in 60.027837 / 108.325842 seconds. All deep Moves pass; the 120-second deadline is unchanged. |
| [Broad Release](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_broad_release/README.md) | 17/0/0; Copy 142/142 and lookup 294/294. Parallel/serial wide Copy publish all 4,097 files in 6.184419 / 18.162982 seconds. All deep Moves pass. |
| [Debug CRUD](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_crud_debug/README.md) and [Release CRUD](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_crud_release/README.md) | 3/0/0 each. These are loopback checks, not live endpoint mutation evidence. |
| [Debug converter](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_theme_debug/README.md) and [Release converter](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_theme_release/README.md) | Native DxUi Theme passes, including storage reuse, Unicode/NUL, malformed/empty input and recovery. Debug's observer initially checked the wrong output stream; original logs and a digest-bound correction are retained. |
| [Pre-repair Debug baseline](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_baseline_debug/README.md) | 288/294: precisely six malformed DOS numeric-size controls fail. Four parser-only controls pass. Its build has an unnumbered vcpkg app-local warning despite a zero-warning summary; retain this C6 staging/diagnostic-accounting witness. |
| Live configured profiles | FTP, SFTP, SCP-profile SFTP listing and S3 read-only comparisons passed twice. IMAP remains TLS-policy blocked. Read-only/loopback results do not prove live mutations or real server quotas. |

The repair following `3917fef5` reuses only filename storage through the canonical
strict UTF converter, resets every row's metadata, and rejects incomplete DOS
integer tokens. It preserves full-listing admission and timestamps. Four parser
controls each inspect 262,208 rows (64 x 4,097). Focused Debug Unix present/missing
cost falls from 834,239 / 812,691 to 626,350 / 552,519 us; DOS from
744,995 / 732,893 to 484,903 / 476,961 us. These same-machine single-run
improvements are directional. Callback/framing and construction are included;
FTP and fixture generation are excluded. No stale cache, deadline increase,
fixture reduction, retry or machine-policy change was introduced.

The preceding committed Debug run copied only 4,041/4,097 files before timing out;
the completed serial run is not an equal-work throughput comparison. The earlier
completed-listing connection-lifetime repair remains intact: completed transports
release their pool borrow while unread bounded rows survive. Historical failures,
exact live-profile evidence, and build receipts remain in the
[prior checkpoint](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_parser_broad_debug/prior-checkpoint.md).
Its snapshot hash matches the tested-source manifest; historical links use the
original WIP base. The Debug lookup archive retains the exact ten-file candidate
patch/hashes. Post-test plan, provider-spec and archive-attribute changes do not
claim another runtime result.

Continue in this order, committing each coherent code/test/spec/evidence step:

Work after `f0710f9b`: Curl ignored `moveMode`; a host native-only Move could enter
the cross-connection copy/delete fallback. The repair validates modes in the
existing progress owner before endpoint resolution, captures native-only admission
immutably, and refuses cross-endpoint relay with the original per-item no-commit
evidence. Same-endpoint file/directory rename and default operations remain working.
The fake server atomically renames exact directory subtrees, including nested/empty
directories, preserving prefix siblings. Debug/Release focused, broad and CRUD
validation above is complete; baseline and source/resource evidence are archived.
Post-test plan/consistency notes do not claim another runtime result. No runtime
handles remain active.
This does not close the whole-tree memory issue. A nonblocking product
question asks whether bounded batches with clear partial results may replace
the existing all-tree-before-mutation requirement; do not assume that change
while the answer is pending. Continue independent authorized C0 work if that
product decision blocks a particular memory-policy change.

IMAP checkpoint after `f5858fbe`: the prior cancellation/unknown-size repair and
200-summary workspace remain intact. Structural FETCH parsing now borrows complete
top-level values, rejects duplicate/damaged framing without resynchronizing into
message literals, repairs incomplete metadata within the existing request budget,
and preserves unavailable flags. Diagnostics retain counts, not message payloads.

| IMAP structural-attribute evidence | Result |
|---|---|
| [Debug failing baseline](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_imap_listing_attributes_baseline_debug/README.md) | 110/134 assertions pass, 24 fail; clean test completion. |
| [Debug candidate](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_imap_listing_attributes_candidate_debug/README.md) and [Release candidate](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_imap_listing_attributes_candidate_release/README.md) | 208/208 each, retaining all baseline controls; 12 standalone helper functions pass per configuration; source/resource policy 197/197. |
| [Broad Debug](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_imap_listing_attributes_broad_debug/README.md) and [broad Release](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_imap_listing_attributes_broad_release/README.md) | 18/0/0 each, including IMAP 208, Copy 142, lookup 294, Move 247 and both 4,097-file Copy scenarios; no retries or increased deadlines. |

Full test-enabled builds have zero warnings/errors. Both candidates bind snapshot
`57bc7d9f6a1ddc2e80539eec07b74433272555dedc3d9cc4d842d78a580c756c`: Debug receipt
`a3b75f8ef4e4590e2f4fb84014dd9839ba368322e37cdd2cb071b35485a2d319`, Release
`c38aa434016840c80c9f10fdc99067005f64010deb8e33b3883156bbf243e373`. Exact patches,
working-byte hashes, full receipts, failed baseline and ordered metrics are archived.
For complete 1/201/4,097-message listings, requests remain 1/2/37 and map peaks
1/200/200. Same-machine Debug durations are 1,069/5,992/115,815 us before and
1,012/5,418/113,100 us after: single observations, not a speed/percentile claim.
Malformed persistent metadata can require more repair requests; the comparison
records that cost explicitly. No live connection or certificate setting changed.

Whole buffered-response and opaque transport memory, live TLS/cancellation, and
the remaining route matrix below are still open. UID SEARCH parser qualification
is now complete for the focused Debug/Release and broad Release lanes, but this
checkpoint does not close C0. All candidate processes are terminal.
Post-test plan, consistency, FETCH-range wording and archive annotations are
documentation-only; a subsequent source/test/tool change needs a new compatible
receipt. Continue committing coherent steps; the four extras stay HOLD.

The implementation checkpoint at `5ed9032c`: the UID SEARCH parser accepts numeric prefixes,
zero/out-of-range identities and malformed partial results, and mistakes missing
or lowercase SEARCH replies for complete empty mailboxes. The existing fixture
adds 35 four-check response scenarios (348 total) before repair: valid empty and
case variants, exact range/token errors, duplicate/multiple result rows, missing
or truncated replies, and opaque literal/quoted framing. The currently unimplemented
ESEARCH form must report Unsupported, never successful empty; this is not an
IMAP4rev2 feature claim. Preserve a failing baseline, then qualify the correction
through Debug/Release focused, helper, broad C0 and policy checks. Source and index
remain frozen while build/test processes run; C0 is not closed by this checkpoint.
The Debug baseline now completes with 260/348 passing and 88 failing checks;
its eleven C5246 warnings came only from the newly added fixture aggregate
braces. The candidate corrects those braces, reuses exact numeric/borrowed
framing, deduplicates UIDs and adds nine one-shot control checks (357 total).
Candidate test-enabled Debug build and focused C0 qualification pass (357/357;
three runner cases), standalone FileSystemCurlTests pass (12 helpers), and the
197/197 source/resource policy run passes. The candidate Debug archive preserves
those results and the receipt-bound source hashes. The focused test-enabled
Release qualification also passes (357/357; three runner cases), and the broad
Release lane passes 18/18 C0 cases with zero failures or skips. Both Release
results, the shared zero-warning/error build receipt, helper trace and policy
evidence are archived in the UID candidate Release and broad Release archives.
The earlier fail-closed residual-process/WMI guard blocker is cleared; no
process-safety bypass was used.
The follow-up governed retry `c0-imap-uids-candidate-release-20260906-r2` stopped
at the same pre-build guard before producing any Release evidence; its explicit
elevated retry was refused by the environment usage limit. That failed attempt
is historical only and is superseded by the successful `r3` Release run. All
candidate processes are terminal. The accidentally committed repo-root
`last_run/trace.txt` was removed in `6f3686d4`; no generated repo-root trace
remains.

1. `[x]` (2026-09-06) The O(files) native Move commitment map is bounded:
   `PreflightDirectorySourceSizes` keeps its membership and size proof as 64-bit
   path digests in vectors finalized once after preflight (16 bytes per file, 8 per
   directory; `FileOps.Curl.MovePreflight.RetainedDigestBytes` reports the
   whole-operation figure), the delete pass proves the source root is a member before
   deleting anything (`ERROR_INVALID_DATA` otherwise), and the late-writer rule is
   unchanged. No aggregate file cap. The plugin's Move-preflight self-test exercises
   the digest containers (membership, size proof, retained bytes for a 4,096-entry
   tree, parity through the shared join). Landing evidence: Move-preflight 256/256,
   C0 Curl family green except `C0_CurlCopyTraversalTruth`, which sits at its 180 s
   budget by construction (the broad Debug run above recorded 182.8 s wall time; its
   two wide 4,097-file copies dominate and do not consult the digests) and passes on
   this tree with `-TimeoutMultiplier 1.5`; C6 owns the budget. GDrive commit-loss
   rows stay deferred.
2. Finish IMAP buffering, real FTP/SSH session limits, opaque transport memory,
   and the provider/route/walker plus host Rename matrix. Do not rerun completed
   parser diagnostics merely to restate their result; rerun affected coverage
   after the next causal code change. Working-tree native counts become lookup
   298, Move 250, Copy 142, and IMAP transport 15 once the next compatible
   receipt is archived. Serial 4,097-file Fake FTP Copy keeps a 180s item
   deadline so the last publications cannot miss the old 120s clock. Broad
   selection is `-CaseFilter C0_`;
   comma-separated FileOps selectors are rejected. CRUD selection is
   `R0fCurl_FakeFtpReadWriteCreateDelete`.
3. Live profiles remain separately scoped. Seed only reviewed non-secret
   definitions into governed isolated settings. Hello bypass permission covers
   only those isolated FileOpsSelfTest connections; normal settings are untouched.
   FTP root: `/home/RedSalamander-selftest`; SFTP/SCP:
   `/share/CACHEDEV1_DATA/homes/eric/RedSalamander-selftest`; S3:
   `/redsalamander-selftest`; IMAP: `/INBOX/RedSalmander-selftest`
   (saved spelling). SCP's read-only witness uses its SFTP listing route.
   IMAP's saved `ignoreSslTrust=true` lacks the required per-profile and
   automation acknowledgements, so host policy refuses it before dispatch.
   The question about enabling certificate verification in the isolated IMAP
   copy remains unanswered; Hello permission is not TLS/SSH permission.
   Live mutation witnesses require separately authorized uniquely owned children.
   Keep all five provider outcomes visible. Hosted startup metrics include
   selftest work and are not pure startup or remote-throughput measurements.
4. Continue C1–C6 after C0's remaining requirements. The four deferred extras stay
   HOLD. Exact Fresh Full and inherited I17 remain open. Shared-DLL Clean
   ownership, diagnostic accounting and the non-test Release SQLite transition
   are verified in the independent C6 checkpoints below.

C0 closeout (2026-09-07): C1–C5 landed after this checkpoint (see their rows). The
route/walker and host Rename evidence matrix lives in the gate #21 archive named in the
C0 row; each cell counts the cited case IDs by the status recorded in that run's results.
Facts recorded without a code change: the IMAP summary workspace peak is at most one
200-UID fetch chunk (`FileSystemCurl.Imap.cpp` `kFetchChunkSize`, asserted by the plugin
listing self-test); the `UID SEARCH ALL` reply is buffered as one whole-mailbox string
(about seven bytes per UID) and libcurl's transport memory is opaque. Bounding those two
needs UID-range searches in bounded batches, which is the pending product question;
nothing was assumed. Real FTP/SSH quotas, live IMAP TLS, and Google Drive live mutations
remain owner/environment items and are listed as such in the final gate record.

All candidate build/test handles are terminal. A further source/test/tool change
requires a fresh compatible full-solution receipt. Keep source/spec/index frozen
during each runtime run; preserve failures and ordered/digest-bound metrics.
Use governed runners, unique IDs and initialized `C:/RedSalamander.Perf`.
Never inspect, clean or stage user-owned repository-root `last_run/`, delete old
runs to hide disk-audit findings, bypass ownership checks, or kill an unrelated
process. The first broad Debug launch was refused by sandbox process-inspection
permissions; the approved retry kept the guard unchanged. No push or destructive
cleanup is implied; Git metadata requires the normal approved escalation.

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

Landed (C3a, 2026-09-07): `ReadBoundLocalLink` returns the stored target text and relative flag
with the outside-root mapping and no source-relative component. `ResolveReparseTargetAbsolute`,
`TryRetargetPathIntoDestination`, `SemanticCopyLinkState`, the Keep Both mapping registration, and
`DrainDeferredSemanticCopyLinks` are removed; links publish inline in provider order through the
unchanged exclusive-stage route. The ABI keeps its shape: the other mapping values and the
transform's component mappings are documented as reserved and ignored. Tests:
`Phase12_ReparsePointPolicy` (copied loop and moved absolute in-tree junction keep source text),
`Fairstream_MovedTreeKeepsLiteralLinks` (Keep Both beside literal links, skipped target beside a
literal link with one prompt, 4,200 junctions with no ceiling), and the Local plugin self-test
(literal payload matrix; component mappings never change a payload).

Landed (C3b, 2026-09-07): the bridge reads the literal payload and publishes each link in provider
order. Removed from `FolderWindow.FileOperations.State.cpp`: `SparseLinkComponentMappingRecord`,
`ActiveLinkTraversalFrame`, `HeldDirectoryCleanupRecord`, the semantic-link mutex and its vectors,
`RegisterSparseLinkComponentMapping`, `RegisterFailedLinkDependencyPrefix`,
`ClassifyLinkDependency`, `HasPendingLinkDependencyForSource`, `HoldDirectoryCleanup`,
`FinalizeHeldDirectoryCleanups`, `QueueDeferredLink`, `DrainDeferredLinks`,
`HasRetainedDeferredLinks`, the parallel-traversal deferral flag, the unmappable Target-conflict
route, and the eight `_bridgeLink*` counters with their rows. Keep Both on a nested child still
chooses the sibling in that child's parent; it records nothing. The PluginConfig bridge guard and
the Pester bridge contract now pin the machinery as absent. The ABI keeps its shape (mapping enum,
component mappings, source-relative buffers) as documented reserved fields; trimming it is a
separate decision recorded for the owner.

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

Landed (2026-09-07): `EnterFailureSurface` runs on host-attach failure (real or forced) and when
`Render` finds no Direct2D target or brushes (real or forced through the new
`DebugFailNextFileOperationsD2DTargetAttempts`). It detaches the host, discards device resources,
and creates the three native children with WIL font ownership; the popup then paints only its
ground, refreshes the text on its timer, routes `WM_COMMAND` to the existing Cancel all confirmation
or hide-only Close, handles Tab/Enter/Escape through a plain window-procedure subclass on the two
buttons, and refuses every other action in `OnActivatedHit` and the self-test invoke. Removed: the
`! _controlsRoot` progress/chevron painting branches, `DrawVerificationProgressBar`, the fallback
body of `DrawBandwidthGraph` and its brushes, the fallback graph metrics, snapshot fields, and seed
constant, `OnMouseMove`/`OnMouseLeave`/`OnLButtonDown`/`OnLButtonUp` with the hover/pressed state,
and `TestFileOperationsPopupFallbackGraphUsesSharedBoundedPainter`. The layout snapshot reports
`failureSurfaceActive`, `failureSurfaceChildCount`, and `failureSurfaceText`. Spec:
`UI_FileOperationsPopup.md` hosted contract and the new Failure Surface section.

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

Landed (Properties, 2026-09-07): `ApplyPropertiesLoadFailure` appends the File Operations section
from the completed load's classification and folds it into the copyable text; `RebuildCards` and
`LayoutCards` place the failure card above the sections. `RefreshDocumentFromIo` is gone; a stream
removal calls `BeginReload`, which runs `StartLoadPropertiesAsync` (worker, fresh classification,
generation token, window-lifetime check). The test fault hook
`DebugSetNextItemPropertiesLoadFault` forces a provider failure or malformed JSON. Cases:
`cmd_pane_itemProperties_window_failure_keeps_artifact_explanation`,
`cmd_pane_itemProperties_window_refresh_reads_fresh_object`,
`cmd_pane_itemProperties_window_close_during_load_is_safe`.

### C6 — Consolidate and qualify the core once

The independent [shared-DLL Clean repair](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c6_shared_runtime_clean/README.md)
started from `96e03773`. Seven deterministic pre-repair failures reproduce deletion,
MSB3061 and stale cleanup ownership; all nine final cases and 73 focused policy
checks pass. Root targets now migrate explicit Clean histories and exclude shared
DLLs from incremental cleanup/recording after late vendor registrations. Exact
current project targets/private intermediates remain cleanable. Runtime copying,
normal parallelism, manifest-owned obsolete cleanup and receipt checks stay enabled.
The owning build spec is updated; historical failed runs remain intact.

Debug rebuild receipt `b69a108f2dd328eb9f134feabe2b2f868f40706f2849593253b96910ff2c5a2f`,
Release rebuild `45f63abadf35472b8da15a4dd0dc5743f30cc70f39822a45b40ca52f9b327e7f`,
and incremental Release `a6fece6675db82bf539822d2caf671a416ad4918b27e0c406522c637dff2576d`
share snapshot `3754f54aea44516549094354eff5053a2588e68eb2fc6d939f6fec34518b7335`.
All have zero diagnostics and a subsequent 3/0/0 governed loopback CRUD smoke.
All 28 captured cleanup histories per profile exclude the monitored shared DLLs.
No build/test handles remain live. Plan/archive-only edits after testing do not
claim a new runtime result; the next code/tool change needs a compatible fresh receipt.

The [diagnostic-accounting checkpoint](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c6_diagnostic_accounting/README.md)
starts from `0b2ef88e`. One shared, culture-invariant header classifier now counts
and colors optional-code diagnostics without counting message text again. The
original 16/30 regression baseline and cold-locale 30/31 baseline are preserved;
the final candidate passes 31/31 and the surrounding policy suite passes 98/98,
with no skips. Replaying the unchanged historical vcpkg log counts one warning,
not zero. Exact tested inputs and same-machine timing samples are archived; this
tool-only repair does not claim another native runtime or full-solution result.

The [production Release transition](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c6_production_release_transition/README.md)
then ran at committed `0ffe4e18`. All 177 prior test-enabled artifacts were verified
before the full `RSBuildEnableTests=false` rebuild. It finished with zero warnings
and errors in 164,299 ms and published receipt
`6af048447fb32e79182634085a2ee3367f0234f31c3225567e2588c57eef6341`, snapshot
`dd4b8a8b40736ba27199adad66d858f7d91b31a6507fb9233522a511eb255652`.
All 164 production artifacts verify. The 21/21 transition checks include retiring
13 exact test-only binaries, retaining the Search Service's attested SQLite closure,
retaining the non-shipping package host and both SQLite copies, Search Service help,
and the production package smoke (1,365 checks, no failures/skips). Raw before/after
receipts, commands, build/smoke logs and measurements are retained. This is an
in-place output-profile smoke, not clean ZIP extraction or Fresh Full evidence.

Return to the remaining C0 work above, then C1–C5 and final C6 consolidation/Fresh
Full/inherited I17. This historical transition produced test-disabled artifacts;
the later IMAP checkpoint above rebuilt and qualified both test-enabled profiles.
All transition processes are terminal. Post-test plan/archive/attribute updates
do not claim a new runtime result. The earlier smoke runner's 183–185 disk-audit
issues remain preserved, not removed for closeout. The four deferred extras remain
outside core; no live remote mutations or ordinary connection changes occurred.

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

### C7 — Spec truth after the landed rows

Landed 2026-09-07. The independent review of `1f665a75` found the VFS strategy matrix still
describing the pre-R0f world (network routes `uncontained/0` and rejected), the S3 summary row
uncontained, a Google Drive read-only profile that no longer exists as a separate profile id, the
Local row promising an in-tree link retarget, and three FileOps domain sentences contradicted by the
code and the popup spec (aggregate stays indeterminate; a legacy fallback painter; R0f "still open").
Each row now states what its plugin descriptor declares. The Local descriptor itself declared
`retargetInTree=true` after C3; it now declares `false`, and the provider-matrix step pins that.
No behaviour changed.

### C8 — S3 reader cancel on its own thread

`S3RangedFileReader` implements only `IFileReader`. Its GetObject arms the request control, but
`ArmS3RequestControl` reads the per-thread current options, which the pipeline reader thread never
establishes, so the continue handler is not installed and cancel is seen only between Reads.
Landed 2026-09-07: the reader implements `IFileReaderOperationControl` (the host already hands
the options to any reader that exposes it), `Read` opens `S3OperationOptionsScope` from the stored
pointer and checks the control before each request, and scenario 4 of the stalled-request self-test
holds a `GET` and proves `ERROR_CANCELLED` within the bound. `FileSystem_S3.md` and the VFS row
state the new contract.

### C9 — MTP decision-time replace identity

`CommitWriterOverwriteWithTempSwap` reads the destination PUID after the temp upload is verified and
deletes that object. An occupant replaced during a slow upload is therefore deleted although the
user agreed to replace the earlier one; `FileSystem_Mtp.md` documents that as the rule. Landed 2026-09-07. Reading the host showed a second fact: MTP implemented neither object binding
nor `IFileSystemAtomicWriter`, so File Operations never offered Overwrite on a device destination
(the bridge logged `bridge.replace.expectationUnavailable`); the race lived in the provider-native
paths and in direct plugin users. The change makes the decided occupant the only delete identity in
both temp-swap commits (resolved before the upload; `DeleteItemByIdentity` re-resolves live and
refuses a different object), gives the writer the expected-replacement contract with an early
refusal, and declares the overwrite swap atomic-final so the host offers Replace and carries the
occupant through the writer. The fake backend gained a replace-during-upload fixture; the specs
state the new order and the witnesses.

### C10 — Identity-bound Permanent Delete

Providers without object binding are marked `nativeAuthority` at admission and deleted by path at
execution. Google Drive and S3 declare `delete: true` without binding; Curl has no identity at all;
Graph and MTP declare `delete: false` (Graph Recycle already consumes a stable item ID). Landed 2026-09-07. Implementing `IFileSystemObjectBinding` on S3 and Google Drive was rejected:
the host routes Rename, Copy publication and owned stages through bound objects, so a partial
binding would have changed those routes. The narrow contract `IFileSystemIdentityDelete`
(`ResolveDeleteIdentity` / `DeleteIfIdentity`) carries only the delete identity: the host pins it
for every native-authority root while the task prepares, re-resolves it after the card's answer,
and deletes through it at execution; the `nativeAuthority` path delete survives only for a provider
without the contract, and the card then carries the by-name sentence. S3 already deleted
conditionally on the revision it observed inside one call; it now pins that revision from Preparing.
Google Drive already deleted by file id; it now refuses a name that gained another id. The R0f
fake-provider cases carry the witnesses in their final phases.

### C11 — Curl copy traversal: the Wide serial scenario

Surfaced 2026-09-07 when the C0 Curl truth steps first ran in a family. The host case reported
"native safety worker timed out": its 180 s budget was shorter than the export it hosts, whose Wide
matrix (4,096 files, Copy and Move, single and bulk entry, concurrency 1 and 4) runs about four
minutes on loopback. Called directly, the export finishes with one failing verdict,
`shape=8;CopyItem;concurrency=1 preserves exact contents, discovery failure and cooperative
cancellation`, whose components (status, servers, contents, fixture bound) it did not print. Landed:
the budget is 480 s and the verdict prints its detail plus the per-endpoint listener numbers. The
detail shows the only failing component: the destination endpoint ends with one pending passive
retirement (12,290 passive-listener reuses against 12,289 data-transfer reuses), so one PASV is
issued without a following data connection, intermittently and only for the Wide matrix; the
source endpoint, contents, statuses and cancel scenarios are exact. Open: the origin of that extra
PASV (plugin publication path or libcurl connection close) and the fix; the bound is not loosened
until then. Reproduced on `4bc5c987` before C10, so C7–C10 did not introduce it;
the window opens at the last green focused C0 run (2026-09-06 13:23) and includes the Curl commits
`3f2fa6bf` and `e7a2776a`.

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

- Fresh Full `20260907T111154Z-110364-496172ac310a45f7b4934393bf924d4a` on `03d851fa` (2026-09-07, after C7–C10: spec truth, S3 reader cancel, MTP decision-time identity, identity-pinned Permanent Delete): 2,155 total; 2,089 passed; 6 failed; 60 skipped. Suite verdicts: 33 of 35 harness entries promoted (CompareDirectories, all DxUi, both ViewerPE and ViewerSqlite groups, the standalone executables, perf, and tooling Pester); the Commands entry (890 passed, 4 failed, 2 skipped) and the FileOperations entry (163 passed, 2 failed, 29 skipped) are the two incomplete entries. Classification: `R4A19_DiscoveryIndependentVolumes` is the environment gate (no authorized alternate-volume test root); the other five are run-window timing artifacts, not code: four Commands UI cases (`cmd_preferences_dialog_category_tree_keyboard_expand_collapse_and_child_entry`, `cmd_preferences_dialog_category_switches_do_not_churn_tree_host`, `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete`, `cmd_pane_filter_prompt_uses_dxui_surface`) ran two to six times slower than in gate #20 while the median case ratio between the two runs is 1.06, and `Phase7_ParallelCopyMoveKnobs` took 113 s against 54 s and did not observe more than one in-flight entry; each of the five passed alone twice on `03d851fa` right after the gate (one churn rerun was skipped by the harness because the desktop foreground changed during its measurement, which is the same disturbance); no case that C7–C11 touch failed, and `C0_CurlCopyTraversalTruth` passed in 214 s within its 480 s budget with the C11 bound holding; the `Phase16_Remote*` live-profile cases are skipped because the gate environment has no connection profiles. Archive: `../../TestRuns/4cb089111a23/Continuation/2026-09-07_144220_reliability_first_gate21/`.
- Fresh Full `20260907T050325Z-126352-bc9b07a54595434397c2c7daa237249a` on `37957c69` (2026-09-07, after C1–C5 and before C7–C10, superseded by gate #21: Preparing/shutdown, proof-gated Retry, literal links, the popup failure surface, Properties failures): 2,152 total; 2,084 passed; 1 failed; 67 skipped. Suite verdicts: 34 of 35 harness entries promoted (Commands, CompareDirectories, all DxUi, both ViewerPE and ViewerSqlite groups, the standalone executables, perf, and tooling Pester); the FileOperations entry is the one incomplete entry (157 passed, 1 failed, 36 skipped). Classification: the single failure is `R4A19_DiscoveryIndependentVolumes`, the environment gate (no authorized alternate-volume test root); no code-attributed failure; the `Phase16_Remote*` live-profile cases are skipped because the gate environment has no connection profiles. Archive: `../../TestRuns/4cb089111a23/Continuation/2026-09-07_082250_reliability_first_gate20/`.
- Fresh Full `20260907T032305Z-17432-12d46bd54eb943bf9f4fcd0a9b7e434b` on `af4a62ea` (2026-09-07, superseded by gate #20): 2,152 total; 2,081 passed; 3 failed; 68 skipped. Classification: `cmd_app_prompt_uses_alert_overlay_window` read the option snapshot after one pump while the overlay's UIA Invoke posts the cycle (bounded wait added in `37957c69`); `ViewerPETests.Noninteractive` failed in `TestViewerImgRawLatestWinsExactReaderAndCloseSafety` (middle request read while the first decode was blocked), the intermittent ViewerImgRaw latest-wins regression recorded on 2026-07-15 under `../../TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_d_ci_gate/`, plugin untouched on this branch, group passes alone; `R4A19_DiscoveryIndependentVolumes` is the environment gate.
- Fresh Full `20260907T014431Z-109112-4c9427c7cc0b4fb2896366aff7053042` on `5e1c4b88` (2026-09-07, superseded by gate #19): 2,152 total; 2,082 passed; 3 failed; 67 skipped. Classification: the `RunAllTestsPlan` stale-run sweep test counted the run's own finished harness directories (fixture-scoped in `af4a62ea`); `cmd_pane_find_dialog_result_shortcuts_use_shell_clipboard_and_file_actions` expected the pre-C1 modal permanent-delete prompt (updated in `af4a62ea`); `R4A19_DiscoveryIndependentVolumes` is the environment gate (no authorized alternate-volume test root).
- Fresh Full `20260906T153040Z-88480-5c83013221c74bf39b28011749f0e908` on `3f2fa6bf` (2026-09-06, after the Curl, IMAP, Local receipt
  and same-host guard work): 2,142 total; 2,064 passed; 6 failed;
  72 skipped. Suite verdicts: every DxUi, viewer, contract, monitor, Curl and performance suite passed; ToolsPesterTests, Commands and FileOperations carried the six failures. Classification: `cmd_connection_manager_window_pointer_click_toggles_visible_dx_toggle` and `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible` passed 1/1 isolated on the same build (UI timing under the loaded desktop); `R4A19_DiscoveryIndependentVolumes` is the known environment refusal (alternate-volume test root not authorized); the Tools Pester sweep test `sweeps stale TestSandbox run directories whose owner process is dead` (`RunAllTestsPlan.Tests.ps1:932`, expected 3 targets, got 5) is environmental: it scans the real TestSandbox runs root, which held 293 named run directories left by the other sessions, and counted two more dead-owner siblings than its fixture; outside the harness the plan tests cannot run at all (empty test root), so the test needs an isolated root (noted for C6); `R0fSmb_BlockedSynchronousCallCancelReturns` and `Phase5_DiscoveryCancelReleasesSlot` fail deterministically in isolation: a task canceled while its call was still blocked now ends with `ERROR_IO_INCOMPLETE` (0x800703E4, "result unknown") instead of `ERROR_CANCELLED`, a regression since the c775b9ed gate, bisected on the same two cases to `cbf19004` (its parent `58f22d0e` passes both 3/3, `cbf19004` fails both): a canceled item without a provider receipt now carries publication Unknown, and the task aggregate turned any unknown axis into an indeterminate task; fixed by keeping a canceled item's unknown axis on the item (visible, Retry still blocked) while the task ends Canceled in `07908056`, after which the two witnesses pass and the R0f-SMB, Phase 5, Preparing-lifecycle, inline-F2, Phase 9, Phase 10 and R4Overlap families, the Commands fileops contracts and the source-contract Pester file are green. The gate therefore does not close C0; the cancel regression is fixed first, then a new Fresh Full follows.
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

### Journey matrix evidence (section 2)

Automated coverage per situation, by existing self-test case; the cause column says how the
failure is produced. "Surface only" means the typed failure/decision surface is proven by the
listed cases but no automation injects that particular cause.

| Situation | Automated cases | Cause / gap |
|---|---|---|
| Routine Copy/Move | `Phase7_CrossPaneRelocateLocal`, `Phase7_CopyRecursiveParallelismMatrix`, `FileOps_CopyMergeIntoExistingFolder`, `FileOps_MoveMergeIntoExistingFolderSameVolume`, `cmd_pane_fileops_*` (F5/F6 admission and popup) | real Local/Dummy transfers |
| Slow preparation or provider | `R1d_PreparingLifecycle`, `C1_MoveShapeProbeRunsInPreparing`, `C0_CurlMovePreflightTruth`, `R0fSmb_BlockedSynchronousCallCancelReturns`, `cmd_pane_fileops_local_blocked_reader_cancel_is_bounded` | blocked reader / preflight fakes |
| Another live task overlaps | `R4A02_*` (11 cases) | real overlapping tasks |
| Name collision or invalid name | `Phase9_ConflictPrompt_*`, `FileOps_ProviderCapabilityMatrix` (invalid rename leaf), `Phase8_InvalidDestinationRejected`, `Causeway_BridgeRejectsHostileChildNames`, `cmd_pane_rename_*` | real collisions and invalid names |
| Temporary failure proved not committed | `Phase9_ConflictPrompt_RetryCap` (attempt counts), `C0_LocalKnownNoCommitReceipts`, `C0_NativeCopyFailurePreservesKnownAxes` | injected known-no-commit failures |
| Server may have committed | `C0_CurlCommittedMutationResponseLost`, `C0_CurlHostCommittedDeleteResponseLost`, `C0_GDriveCommittedCopyResponseLost`, `C0_GDriveCommittedDeleteResponseLost`, `Cinderstar_LegacyWriterEventuallyConsistentMove` | fake providers losing the response |
| Move copied but kept source | `Floodgate_CrossFsDirectoryCopyOnlyRetainsSource`, `Floodgate_CrossFsCopyOnlyNeverEntersSourceCleanup`, `Floodgate_CrossFsMoveGetSizeFailurePreservesSource`, `Riptide_MoveDistinctSameSizeFilePreservesSource`, `Fairstream_MovedTreeKeepsLiteralLinks` (skipped target) | Copy-only profiles and verification refusals |
| Some items fail or are skipped | `FileOps_CrossVolumeMovePartialFailureStatus`, `C0_CurlPartialTreeFailure`, `C0_CurlHostPartialTreeFailure`, `Riptide_BridgeSequentialContinueOnErrorCopiesSiblings`, `Riptide_BridgeNestedDirVsFileSkipContinuesSiblings`, `Phase10_TypedResultsAndConsumers` | partial trees and skips |
| Cancel | `Phase5_CancelQueuedTask`, `Phase5_DiscoveryCancelLatencyLocal`, `Phase5_DiscoveryCancelReleasesSlot`, `R4A02_DontStartCancelsBeforeMutation`, `Cinderstar_LegacyWriterCancelDuringBackoff`, `C1_ExitCloseDeferredUntilTasksQuiet` | real cancellation at each phase |
| Cleanup fails after success | `Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete`, `Riptide_LiveFinishedSnapshotCarriesDiagnostics` | surface only: retained artifact after an abort failure; no case injects a cleanup failure after a successful publication |
| Popup cannot initialize graphics | `cmd_pane_fileops_popup_failure_surface_on_host_attach_failure`, `cmd_pane_fileops_popup_failure_surface_on_d2d_target_failure` | forced attach / target failure |
| Possible `.rs_*` artifact | `Phase10_ArtifactTouchGuard`, `cmd_pane_changeAttributes_recursive_artifact_guard`, `cmd_pane_itemProperties_window_streams_can_remove`, `cmd_pane_itemProperties_window_failure_keeps_artifact_explanation` | real artifact names |
| Application close with live work | `C1_ExitCloseDeferredUntilTasksQuiet` | real close with a running task |
| Disk full | `Phase10_DeferredConsentAndRecycleEscalation` (insufficient-space and space-unknown consent before the first write) | surface only for a mid-transfer `ERROR_DISK_FULL`; the not-committed failure card is proven by the temporary-failure rows |
| Permission denied | `Phase10_ClipboardAdmissionAndRetainedActions`, `R1d_PreparingLifecycle`, `FileOps_ProviderCapabilityMatrix`, `Phase12_ReparsePointPolicy` (deny-listed junction) | real ACL denials at admission, Preparing, and traversal |
| Device removal | `TestFolderViewRenderingErrorOverlayRequiresPersistence` (pane only) | surface only: no file-operation case removes a device mid-transfer |
| Connection / authentication loss | `Causeway_BridgeSchedulingAndResourceContracts` (`ERROR_UNEXP_NET_ERR`), `Phase16_Remote*` (real services, environment-gated) | surface only for authentication loss mid-transfer |
| Blocked cancellation | `R0fSmb_BlockedSynchronousCallCancelReturns`, `Riptide_HostPerItemSchedulerShutdownWaitsForBlockedWorker`, `Cinderstar_LegacyWriterCancelDuringBackoff`, `cmd_pane_fileops_local_blocked_reader_cancel_is_bounded` | blocked readers and workers |

Human observations (not performed in this closeout; owner checklist): mouse and keyboard-only
paths through a conflict card and the failure surface; UIA tree and screen-reader wording of the
card, footer, and failure surface; high contrast and reduced motion; RTL and the cs-CZ/fr-FR/ja-JP/
sk-SK translations at minimum popup width; 100/150/200% DPI. Automation already covers keyboard,
UIA names/Invoke, high contrast, reduced motion, and a DPI change for the popup and the failure
surface; the human pass confirms wording and layout, which automation does not judge.

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

Owner decision 2026-09-07: the plan moved to Done with the C5 human observation pass
outstanding; the owner reopens it if that pass is not valid. The other outstanding
items are listed in C0, C6, and C11 for their owners, not as work in this file.

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
