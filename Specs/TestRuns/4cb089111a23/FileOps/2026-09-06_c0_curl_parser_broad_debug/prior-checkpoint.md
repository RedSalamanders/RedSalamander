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
| C0 | VERIFY | Review corrections and an explicit route/walker evidence matrix | B1/B2 | Lookup 284/284 and CRUD 3/3 in Debug/Release. Completed-listing borrow release fixes broad Release: 17/0/0, Copy 142/142, all deep Moves pass. Debug remains 141/142: serial wide Copy publishes 4,041/4,097 at its unchanged timeout. Next: reduce measured repeated strict-listing cost without relaxing admission proof. Live FTP/SFTP/SCP-profile listing/S3 pass; IMAP is TLS-policy blocked. Section 4 retains Move proof memory, IMAP, FTP/SSH limits, route matrix and host Rename. |
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

C0 remains **VERIFY**. The completed-listing connection-lifetime repair passes broad
Release, but Debug still times out on serial wide Copy. No C1–C6 gate is closed.

| Current evidence | Result / limitation |
|---|---|
| [Release lookup/build](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_cursor_release_lookup_release/README.md) and [Debug lookup/build](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_cursor_release_lookup_debug/README.md) | 284/284 native checks and 3/0/0 suite results in both builds. Full solution builds have zero warnings/errors and share source snapshot `87965dd7ee6ceb0b341a8d6016ead438d846e6e6c216a05abeba6d45b8ad0829`. Source/resource policy: 197/197, no skips. |
| [Broad Release](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_cursor_release_broad_release/README.md) | 17/0/0; Copy 142/142 across 46 scenarios and lookup 284/284. All four formerly failing depth-128 Moves pass; no Windows 10048 errors. Parallel/serial wide Copy both publish exactly 4,097 files in 6.219373 / 18.936925 seconds. |
| [Broad Debug](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_cursor_release_broad_debug/README.md) | 15/1/1; Copy 141/142, lookup explicitly not reached. Deep Copy/Move pass with two authenticated sessions per endpoint. Parallel wide Copy passes 4,097/4,097 in 64.002003 seconds; serial publishes 4,041/4,097 before its unchanged 120-second deadline. Zero connection failures, zero pending listener retirements. |
| [Release CRUD](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_cursor_release_crud_release/README.md) and [Debug CRUD](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_cursor_release_crud_debug/README.md) | 3/0/0 each. Separate passing CRUD never replaces a failed broad run. |
| [First candidate](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_cursor_release_initial_lookup_release/README.md) | Lookup 284/284. Its C4457 test-name shadow was fixed before final builds; its MSB3061 shared `yyjson.dll` Clean warning remains a C6 witness. |
| Live configured profiles | FTP, SFTP, SCP-profile SFTP listing and S3 read-only comparisons passed twice; IMAP is blocked by saved TLS policy. Read-only/loopback results do not prove live mutations or real server quotas. Details and exact roots remain in the preserved checkpoint below. |

The production change reuses the existing fixed cursor buffer and pool. Small
listings finish before their first row is yielded; completed, failed and canceled
transports detach and return their borrow while cursor objects retain only buffered
rows and parse state. Nested parent/child controls prove one authenticated session
for two LISTs while unread parent rows survive. Empty, fragmented, 4,097-row,
abandoned, canceled and late-final-reply failures retain their exact truth.
Large listings still stream with bounded pause/redelivery. No new pool, whole-tree
cache, client bind policy, automatic retry or machine setting was introduced.

**Next bounded step:** reduce repeated strict listing/row materialization cost.
The latest serial Debug run spends 51.907702 seconds in staging-name lookup and
54.154806 seconds in promotion; nested fixture generation/lock wait accounts for
39.371795 seconds across 8,084 destination LISTs. Measure pure production parsing
separately from the existing exact-byte fixture generator before choosing the
change. Preserve all malformed/duplicate/missing/timestamp/final-reply checks and
the current deadline; no stale listing cache may become mutation authority.
Re-run broad C0 in both builds after the repair, not only an isolated wide case.

Work after `3917fef5`: a no-network parser benchmark now covers four Unix/DOS
present/missing scans (64 x 4,097 rows each), and six fragmented DOS numeric-size
regressions reject `4junk`/`4.5`. The current DOS parser accepts those integer
prefixes. Expected lookup count is 294; collect the pre-repair Debug baseline,
then repair numeric validation and measured per-row filename allocation. These
new tests initially yielded 288/294: only the six numeric-size controls failed.
The four parser-only cases validated 262,208 rows each in 0.73-0.83 seconds.
The candidate reuses filename capacity with the shared strict UTF output overload
and rejects incomplete DOS numeric tokens. The existing DxUi Theme UTF-policy
test covers storage reuse and malformed/empty/Unicode recovery. Candidate runtime
verification is pending. The baseline build has an unnumbered vcpkg app-local
dependency warning despite its summary reporting zero warnings; retain it as a
C6 staging and diagnostic-accounting witness, not a warning-free build.

The same-machine prior Release run had four deep-Move 10048 failures and Copy
134/142; prior Debug serial Copy stopped at 3,828/4,097. New session-lifetime
evidence proves the bounded ownership repair, not global TCP exhaustion or a
universal FTP/SSH quota. Single-run timings are directional; failed and completed
serial runs are not equal-work throughput comparisons.

The full diagnostic history, exact live-profile roots, previous build receipts and
failed experiments are preserved byte-for-byte in
[the prior checkpoint](../../TestRuns/4cb089111a23/FileOps/2026-09-06_c0_curl_cursor_release_broad_debug/prior-checkpoint.md).
Its `plan-snapshot.json` binds that snapshot; historical relative links retain their
original WIP base. Final lookup Release retains the seven-file tested patch/hashes.
All build/test handles are terminal. Documentation/attribute changes follow the
tested snapshot; any further code/test/tool change requires a fresh compatible
full-solution receipt. C0 still needs whole-operation Move commitment memory, IMAP,
real FTP/SSH quotas, opaque transport memory, and the provider/route/host Rename matrix.

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
1. Finish the serial Debug wide-copy path using the measured strict-listing costs;
   the Release deep-tree failure is repaired. Use `-CaseFilter C0_` to keep
   the predecessors and Copy together; the documented comma-list is rejected by
   FileOps and must not be used until its selector is repaired and tested. Repeat
   broad C0 after a causal correction, not just a successful isolated retry.
   Distinguish production parsing from fake listing serialization cost, with exact
   row/byte oracles proving unchanged admission and faults.
   Complete Debug wide Copy; rerun `C0_CurlEntryLookupTruth` (284), `C0_` including Copy
   (142 assertions / 46 scenarios), and `R0fCurl_FakeFtpReadWriteCreateDelete` in both builds.
   Use the governed runner with unique run IDs. Do not use `-SkipBuild` with a
   project-only receipt. Preserve/restore `RSBuildEnableTests` for Release repeats.
2. Freeze source/spec/index during runs, preserve failures and digest-bound metrics,
   then run source/localization, spec inventory, explicit/archive-inventory checks
   and patch hygiene before committing. Broad Release is green; broad Debug remains red.
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
