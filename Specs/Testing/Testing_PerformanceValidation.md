# Performance Validation Specification

## I19 accepted static-library adoption costs

On 2026-09-13 the user explicitly accepted the measured `.lib` adoption trade-off
and requested closeout. Retain the [matched Preferences record](../TestRuns/Local-x64/DxUiAdoption/2026-09-13-LocalQualification/README.md)
and the original package comparisons. All 40 Preferences round trips pass; the
candidate's median observed peak private bytes increase by 1,046,528 bytes (0.384%),
while working set decreases 1.071%. CPU measurements change in opposite directions
across the pairs and establish no CPU improvement or regression.

The accepted resource envelope is the measured I19 adoption cost on the recorded
test-enabled x64 Release fixture and portable package, approximately 1 MiB additional
peak private memory and 1.4% additional packaged disk space. Canonical source/test
ownership and the corrected shared-control behavior justify that bounded cost.
Preserve the baseline, original failed attempts and existing thresholds; this
decision neither admits additional growth nor establishes normal idle, GPU,
hardware-presentation, DLL-versus-library isolation or long-run resource behavior.
RedXe and DxUi retain their own measured trade-offs in their resource contracts.

Final local package verification at `1232d181` records 415,535
additional compressed bytes (1.37%) and
1,140,347 additional expanded bytes (1.45%).
The [final A/B/A receipt](../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Packaging1232d181/README.md)
verifies all 11 packaged DxUi identities and identical baseline payloads before/after.
These are the final artifact measurements for the accepted approximate adoption cost;
earlier package results remain retained. This package check does not remeasure memory,
CPU, GPU or frame latency and does not authorize future growth.


## File Operations residual and UI focus regression witnesses

- `Beeline_CopySkipLinksKeepsPlaceholders` requires a real non-name-surrogate placeholder on NTFS.
  If the primary filesystem is not NTFS, reuse only the explicitly authorized,
  marker-owned alternate FileOps root and verify its filesystem. Without that
  capability, record a real skipped case with its exact reason. The case verifies
  all 65,536 copied bytes and the absent junction, with no hydration/conflict prompt
  for the fully hydrated fixture.
  It fixes automatic concurrency at four and requires exactly one worker admission.
  `FileOps.SelfTest.CopySkipLinksKeepsPlaceholders` records duration in microseconds,
  success in `value0`, worker admissions in `value1`, and the source volume in detail.
  The fixture uses a fully hydrated Cloud Files placeholder with ALWAYS_FULL hydration
  and population, no account, callbacks or network. It verifies the actual reparse
  attribute and Cloud tag before host admission; `FileOps.SelfTest.PlaceholderFixture`
  records tag/attributes in `value0`/`value1`. The local sync-root registration is WIL-owned
  across asynchronous ticks and unregistered before scratch cleanup, including failure.
  The fixture explicitly exposes placeholders in its isolated process and UI thread,
  records the previous modes in `FileOps.SelfTest.PlaceholderVisibility`, and restores
  both through the same WIL owner. Otherwise Windows compatibility disguise can
  silently bypass the reparse-kind branch. WOF compression alone is insufficient
  evidence where its reparse attribute is absent.
- `BR5_CurlHostDirectoryMove` and `BR5_CurlHostDirectoryMoveRefused` run host admission
  and execution against the loopback FTP fixture, including nested files, an empty
  directory and unchanged siblings. `FileOps.SelfTest.CurlHostDirectoryMove` carries
  elapsed microseconds and RNFR/RNTO counts in `value0`/`value1` (exactly one each).
  Both cases prohibit RETR/STOR/DELE/RMD/MKD fallback. A refused attempted rename
  remains conservatively Indeterminate where the provider supplies no authoritative
  non-commit receipt; fixture-observed source retention is not host retry authority.
- `cmd_reliability_preferences_keyboard_search_reacquires_input` displaces native
  focus while logical search focus remains. The wait must reacquire the live page
  input HWND before edit messages; `selftest.preferences.keyboard.input_target`
  records elapsed microseconds and the correct-target bit in `value0`.
- `cmd_reliability_navigation_popup_close_survives_reentrant_restore` includes the
  focus-follows-pointer predecessor, delivers any prematurely queued restore during
  popup destruction, then clears retiring native focus. Escape must still restore
  the FolderView within the existing three-second bound after destruction returns.
  `selftest.navigation.popup_close_focus` records elapsed microseconds, premature
  delivery in `value0`, and successful restoration in `value1`.
- Retain controlled failing evidence and repeated repaired witnesses. Original
  Keyboard/popup cases must also pass in predecessor context and canonical Commands;
  isolated passes do not replace final Fresh Full.

## Overview

Performance work in RedSalamander is a **normative engineering requirement**, not an optional cleanup step.

Any new feature, behavior change, or optimization that can affect:

- UI responsiveness,
- startup time,
- folder enumeration,
- rendering,
- search,
- Compare Directories,
- File Operations,
- plugin I/O,
- memory retention,
- background queueing,

MUST define its validation path up front and MUST ship with measurable coverage from the beginning.

File Operations evidence records endpoint topology and selected strategy, discovery-ahead depth and
queue bounds, publication bytes, verification bytes reread, source bytes read only for digesting,
digest algorithm and CPU time, verification outcome, and link kind/policy/action. Recording only
"same filesystem" is invalid. Validation must prove discovery is not starved by transfer work,
bounded memory holds under deep/wide trees, Skip discovery releases reserved discovery capacity to
transfer, cancellation remains responsive, and verification traffic is not double-counted as copy
throughput.

`BR3_CancelQueuePausedTransfer` runs two real Local transfers, changes Parallel to
Queue, observes both workers parked, and keeps the first task user-paused while
cancelling the queue-paused task. The latter must complete within three seconds
without releasing its predecessor or changing the mode.
`FileOps.SelfTest.QueuePausedCancelLatency` records microseconds, the pass bit in
`value0`, and the 3,000,000-us limit in `value1`. Failure-only fixture recovery
releases Queue mode and drains both tasks; that recovery cannot qualify the bound.
`Causeway_BridgeFailureStatusAndPausedReader` also holds a peer between its false
conflict predicate and sleep. The cancel notifier must wait for that predicate's
mutex, then release the peer within two seconds. `FileOps.SelfTest.ConflictCancelWake`
records the pass bit in `value0` and premature notification in `value1`; emergency
conflict release is failure-only recovery. The test-only gate is absent from production.
`Phase7_ParallelCopyMoveKnobs` admits each parallel transfer with its bandwidth
limit already set and retains it until both parallel in-flight entries and per-call
sharing are observed. Retry counts must remain bounded; fast completion without
those observations cannot pass. `FileOps.SelfTest.CopyMoveKnobObserved` records
elapsed microseconds, in-flight count in `value0`, sharing observed in `value1`,
and configured concurrency/active calls in detail before releasing the limit.
In-flight File Operations test observations use `_inFlightFilesMutex` for the
file-progress table and `_perItemInFlightCallsMutex` for active provider calls.
`_progressMutex` protects neither table; assertions that combine domains must
acquire all relevant owners together with `std::scoped_lock`.

Cross-phase exact-object/link hardening protects literal link preservation, object-kind classification,
Permanent Delete confirmation-to-mutation continuity, and conflict-cache metadata latency. Evidence
uses the task-terminal `FileOps.PermanentDelete.IdentityMismatchCount`,
`ConditionalAttemptCount`, `RemovedCount`, and `IndeterminateCount`; the literal-link self-test rows
below; and the Local provider's
`FileOps.Rename.CaseTempEntropyFailureCount`, `CaseTempCollisionCount`, and `CaseTempSuccessCount`.
Authoritative-kind source, serial/parallel traversal-limit convergence, and cached metadata reads
remain bounded correctness dimensions in the associated selftest/archive scenario rather than
unbounded per-item trace rows.
No throughput improvement is claimed; the protected regressions are wrong-object mutation, guessed
link publication, dead destructive actions, and avoidable repeated metadata I/O when the complete
typed cache scope is already authoritative.

Typed File Operations admission uses `fileops.plan.construct_us` and `fileops.plan.admit_us` as the
envelope-build and confirmation/admission latency rows. `fileops.plan.source_count` and
`fileops.plan.payload_bytes` expose selection scale and retained immutable-plan size;
`fileops.plan.rejection_bucket` is a bounded enum counter. These are one metric set per attempted
top-level plan, contain no provider paths, and must not become per-item or per-descendant emissions.

`C0_CurlImapListingTruth` exercises the production IMAP mailbox, UID, summary and
repair orchestration with a synchronous, connection-owned request fixture. Its 357
checks require cancellation at bulk/repair/later-chunk boundaries to stop further
requests and publish no successful listing, preserve unknown versus explicit-zero
sizes in listing and targeted lookup, retain bounded ordinary metadata repair,
and complete 1/201/4,097-message mailboxes with at most 200 retained summaries.
Deadline and one-shot callback failures remain failures rather than metadata repair;
unsolicited unrelated UIDs cannot grow the map, and FLAGS-only updates preserve
previously fetched immutable metadata. Overflowing/signed sizes remain unknown.
The pre-repair 38-check baseline is retained separately from the expanded candidate.
The subsequent attribute-parser checkpoint retains those 70 checks and adds 64:
thirteen quoted/literal/reordered/malformed FETCH shapes and three partial-summary
repair scenarios. Message text must never establish UID, size or flags. Preserve
its separate failing baseline; a green earlier numeric-token check does not qualify
structural parsing. The structural-parser candidate adds 74 more assertions:
thirteen further four-check shapes (including truncated/overflowing literals,
embedded fake FETCH lines, partial header fields, valid empty headers and unknown
flags), three four-check incomplete-summary repairs, and ten numeric-range controls.
Existing fixture metrics retain request count and workspace peak.
The UID SEARCH checkpoint adds 35 four-check response scenarios before repair
(348 total). Empty, case-insensitive and duplicate/multi-row results are distinct
from missing/malformed replies, invalid UID tokens and opaque literal text.
Enumeration failure must publish no entries and issue no FETCH/repair requests;
unimplemented ESEARCH is explicitly unsupported, not an empty mailbox. Retain
the failing baseline and exact UID/order/request/workspace checks with the candidate.
The UID candidate retains those 348 assertions and adds nine for one-shot control
failure after empty/single-message SEARCH and during a 4,097-UID parse. Check the
exact failure, checkpoint count, zero FETCH requests, empty output and zero summary
workspace; parser duration remains observable through the existing UID/listing metrics.
`filesystem.imap.summary_workspace_peak_entries` records that map high-water mark;
`FileOps.Curl.Imap.ListingFixture` records duration, summary request count and peak
summary count for each scenario. These counts do not include buffered wire responses,
the mailbox/UID vectors, returned pane entries or opaque transport state. They are
not a whole-operation memory bound or live IMAP/TLS/cancellation qualification.
Preserve the failed baseline and candidate rows together; no throughput claim is
implied by this correctness/workspace comparison.

`C0_GDriveCommittedCopyResponseLost` protects provider-native Copy against replay after
server commit and response loss. `FileOps.SelfTest.C0GDrive.CommittedCopyResponseLost`
records one row per faulted task: elapsed microseconds, mutation-request count in
`value0`, and actual backend commit count in `value1`. Both counts must equal one;
the backend object must exist, the source must survive, no replay/Skip prompt may be
published, and the terminal result plus task summary must remain uncertain. Use the
same-machine red and repaired runs as correctness/request-amplification evidence,
not as a throughput or percentile claim. The existing fake-Drive phase timeout bounds
the fixture; this does not establish a live-network or universal cancellation deadline.
The companion `C0_GDriveCommittedDeleteResponseLost` emits
`FileOps.SelfTest.C0GDrive.CommittedDeleteResponseLost` with the same request/commit
axes, requiring one request/removal and unknown source disposition. The Local host
case `C0_NativeCopyFailurePreservesKnownAxes` reuses publication/cleanup fault hooks;
`FileOps.SelfTest.C0Native.UnknownCleanupPromptCount` and
`FileOps.SelfTest.C0Native.PublishedAttributesFailurePromptCount` each require one
ordinary overwrite prompt and no post-failure decision. Assert source/destination
bytes and all independent result axes alongside those counters.
The existing Local final-attributes fault must fire after actual publication and
backup cleanup, with a committed receipt and replacement bytes at the final path.
Injecting the same error before `PublishAs` changes the tested failure phase and is
not evidence for committed-error handling. The direct Local overwrite test also
asserts those replacement bytes before resetting its fixture.

`C0_MutationReceiptPrefixBoundary` protects the actual host completion callback and
shared classifier against reading an absent V1 suffix, accepting truncated headers,
or inheriting earlier proof after malformed/missing completion. Its guard-page fixture
requires no file/provider I/O. `FileOps.SelfTest.C0MutationReceipt.PrefixBoundary`
records elapsed microseconds, five unsupported-header checks in `value0`, and five
host callbacks in `value1`. Also prove full-stage copying, future-suffix normalization,
invalid-stage rejection, and an immutable uncertain result when a provider ignores
the malformed-callback error. This is bounded ABI/correctness evidence, not a
throughput or universal callback-latency claim.

`C0_CurlNativeDeleteGuards` runs the native FTP fixture on an owned worker, with the
DLL retained until join. Its 29 checks cover single/bulk Delete at concurrency 1/4,
recursive/nonrecursive root rejection, successful selected-subtree controls, and
single/bulk cancellation during a stalled first LIST. `FileOps.Curl.NativeDelete.RootGuard`
emits eight rows: elapsed microseconds, destructive command count (must be zero), and
retained file count (two). `FileOps.Curl.NativeDelete.SelectedSubtree` emits eight
rows: elapsed microseconds, one DELE and two RMD commands, with a byte-checked sibling
retained. `FileOps.Curl.NativeDelete.ListCancel` emits two rows: cancel latency in
microseconds (<2,000,000), destructive command count (zero), and cooperative return
(one). A fixture-forced unblock fails the cancellation check. Preserve red/candidate
rows; these are route correctness and bounded loopback cancellation evidence, not
throughput, deep/wide memory, or live-endpoint qualification.

`C0_CurlCommittedMutationResponseLost` has 33 native checks over eight loopback FTP
scenarios: single/bulk Delete/Rename with ordinary acknowledgement or a committed
mutation followed by connection loss and concurrent source replacement.
`FileOps.Curl.MutationReplyLoss` records duration in microseconds, mutation request
count in `value0` (one), and injected loss count in `value1` (zero/one). Assert
exact file bytes, exactly one callback with honest mutation classification, and no
incidental NLST transfer. Keep the pre-fix replay failures alongside candidate rows;
the shorter failure path is not a general throughput claim.
`C0_CurlHostCommittedDeleteResponseLost` exercises the real host task and records
`FileOps.SelfTest.C0Curl.CommittedDeleteResponseLost`: duration, one Delete request,
one backend commit. Require preserved replacement/sibling bytes, an immutable
Indeterminate result, localized unknown summary, and zero conflict prompts.

`C0_CurlPartialTreeFailure` records `FileOps.Curl.PartialTreeFailure` for single/bulk
Delete at concurrency 1/4: elapsed microseconds, mutation commands, and denied CWD
count. Both concurrency paths must preserve uncertain completion after earlier child
deletion; explicit selected-root refusal must prove zero mutations and retain no-commit.
`FileOps.Curl.IndependentItemProof` covers two selected roots at concurrency 1/4,
with one deleted and one refused without mutations; completion proof is item-local.
`FileOps.Curl.CollisionNoCommit` covers single/bulk Move/Rename preflight refusal:
zero mutation commands, unchanged source/occupant, explicit no-commit receipt.
Shared reporter controls reject error-code-only proof for Delete/Move/Rename.
Do not preserve whole-tree preflight just to keep a zero-mutation control green:
when migrating traversal, place an explicit first-listing refusal before any
mutation and separately inject failure after a child mutation at each concurrency.
`C0_CurlHostPartialTreeFailure` records `FileOps.SelfTest.C0Curl.PartialTreeFailure`:
two mutation commands, one deleted leaf, retained later child/sibling, immutable
uncertain result, localized unknown summary and zero Retry/conflict prompts.
Keep the failed baseline alongside the repaired rows. These bounded loopback
measurements do not retire native-walker or real-endpoint qualification.

`C0_CurlNativeDeleteLateListing` supplies 49 assertions over twelve native paths:
single/bulk Delete, concurrency 1/4, and complete/refused/canceled late listings.
The loopback server gates a later LIST on an exact earlier DELE without holding its
command mutex. A five-second gate timeout is a failed ordering witness, not an
accepted provider timeout. `FileOps.Curl.NativeDelete.LateListing` records elapsed
microseconds, released-gate count (one), and gate-timeout count (zero).
`FileOps.Curl.NativeDelete.LateListCancel` records cancellation microseconds
(<2,000,000), exact deleted-leaf count (one), and cooperative return (one).
Require preserved sibling bytes, exact remaining children, and one truthful terminal
receipt; partial changes cannot claim retryable no-commit. Fixture shutdown is only
an emergency unblock and must not pass cancellation. This small-tree ordering test
does not prove wide-tree memory bounds, iterative depth safety, host UI latency, or
live FTP/SFTP/SCP behavior; retain those separate qualification requirements.

`C0_CurlDeleteTraversalBounds` adds 127 assertions over single/bulk Delete at
concurrency 1/4: 4,097 files, 80 nested directories, shared depth-limit refusal,
literal arrow names, no-follow symlink entries, malformed listing rows, out-of-root
child names, an overlong record, and nested names with spaces/quotes/percent/Unicode.
The 36 transport scenarios check exact remaining bytes and terminal receipts with
a 30-second operation deadline. Eighteen pure command-operand checks cover FTP,
SFTP and SCP literal base paths, protocol quoting, percent identity, and rejection
of CR/LF/NUL/invalid Unicode; these do not substitute for live SSH tests. The late-listing
fixture releases its command mutex before transmitting listing data so client
backpressure cannot manufacture a server-side cross-session deadlock.
`FileOps.Curl.NativeDelete.BoundsFixture` records duration, expected leaf count,
and actual deleted leaves. For each native tree, `FileOps.Curl.NativeDelete.Traversal`
records peak active frames (at most 129 including the depth-zero root) and queued
leaves (one at concurrency 1, otherwise at most the shared discovery target).
`FileOps.Curl.NativeDelete.Retained` records peak owned path bytes and metadata
reservations; verify at most 16 MiB and 8 MiB respectively in the archived rows.
Those reservations include fixed cursor buffers and paused-chunk allowance, not
opaque TLS/socket state or total process working set. The wide successful cases
must delete all 4,097 leaves despite the smaller queue ceiling. A refused depth or
record boundary must remain explicit failure, not a silently incomplete success.

`FileOps.Curl.NativeDelete.BoundsFixture` also records actual/expected operation
status, fake-server status, exact-content and receipt checks separately. A loopback
session closed by a refusing/cancelled consumer may end with EOF, connection reset,
or `WSAECONNABORTED`; those are session-disconnect observations, not independent
server defects. Every scenario still requires its exact operation verdict, contents,
and mutation truth. `FileOps.Curl.FakeFtp.PeerDisconnect` records the observed aborted
connection without including paths, commands, credentials, or payloads.

`C0_CurlDirectorySizeTruth` is the directory-size qualification gate (85 assertions
over 26 loopback FTP scenarios, plus four ABI argument checks and two handle-pool
lifetime checks). It covers known/probed/unknown file and child sizes,
empty and shallow directories, 4,097 files, depth 80 and explicit depth-129 refusal,
pre-cancel before network I/O, flat-list progress cancellation, cancellation during
a stalled LIST, root-error precedence, denied-child sibling continuation, malformed/
out-of-root/overlong records, arithmetic overflow, and literal arrow names. The
fixture verifies unchanged server contents, matching return/result status, latest
totals, cookie identity, progress-before-cancel ordering, and one final successful
report. Unknown sizes must not become zero-byte success. The required unknown-child
outcome is partial totals; an unmeasurable file root remains an explicit unsupported
size rather than a fabricated total. File counts include observed unknown-size files.
`FileOps.Curl.DirectorySize.TruthFixture` records elapsed microseconds and final
file/directory counts; diagnostic detail includes actual/expected HRESULTs, bytes,
server status, and feedback checks. `FileOps.Curl.DirectorySize.ListCancel` records
cancel latency, reached-list count, and cooperative return; the owned emergency
server shutdown does not count as cancellation success. Require return within two
seconds after cancellation of an observed blocked LIST. This is a correctness and
bounded-loopback-latency gate, not real FTP/SSH or IMAP qualification. Archive red
baseline and repaired candidate before claiming this gate passed.

The pool checks inspect a live idle handle before re-borrow to prove per-call
options were cleared on return, not deferred until reuse. The content-free
`FileOps.Curl.DirectorySize.IdleHandleReset` metric records that postcondition.
Idle eviction, overflow, and shutdown must never retain expired callback storage.

The additional cases cover a wide parent lookup, a late malformed row after known
bytes, cancellation from final progress, and an access-denied callback that must
remain global rather than becoming a recoverable descendant error. Inspect the
provider's `FileOps.Curl.DirectorySize.Traversal` / `.Retained` rows for admitted
frame/path/metadata peaks as well as fixture status and feedback. Do not represent
these reservations as whole-process memory or IMAP buffered-facade qualification.
Two final-feedback cases accumulate unknown-child partial totals before the final
callback cancels or fails. The global callback verdict must win over accumulated
partial status while retaining known totals; prior global provider failures still win.

`C0_CurlMovePreflightTruth` protects Move admission with 247 assertions: the
original 97 assertions over 32 recursive loopback scenarios, plus 144 over 48
Move-mode scenarios and six over two mixed bulk scenarios. Recursive cases cover singular/bulk APIs at concurrency 1/4,
shallow/depth-80 success, depth-129 refusal, malformed/out-of-root/overlong LIST
records, an unknown-size tail after 4,096 known files, and a blocked LIST cancel.
Rejection must precede payload download and every destination/source mutation;
source/destination contents and one no-commit completion receipt are independently
checked. The wide case must probe all 4,097 files, not hit an aggregate queue cap.
`FileOps.Curl.MovePreflight.TruthFixture` records duration, mutation-command count,
and source SIZE-command count; detail separates actual/expected status, endpoint
health, contents and receipt truth. `.ListCancel` requires cooperative return
within two seconds after cancellation of an observed preflight LIST wait; owned
emergency server shutdown never counts as a pass. Archive the red baseline before
repair. This gate alone does not qualify bounded commitment storage, Copy discovery,
source-identity races, or real FTP/SSH session limits. `.Begin` identifies the next
scenario if a process terminates before its completion row; it is not success
evidence. Provider `.Traversal` records preflight duration, peak frames and retained
file commitments; `.Retained` records cursor-frame path/metadata reservations only.
The whole-tree commitment map and opaque transport allocations remain outside that
reservation metric and must not be described as bounded by it.

`FileOps.Curl.MovePreflight.ModeFixture` records the native-only/invalid-option
matrix for files/directories, single/bulk APIs and concurrency 1/4. Same-endpoint
native Move must rename without RETR/STOR; a cross-endpoint native-only request
must refuse before the copy/delete relay and preserve both endpoints with one
no-commit completion. Unknown Move modes and nondefault modes on Copy/Delete/
Rename must return E_INVALIDARG before provider I/O or item callbacks. Value0
counts mutation commands; value1 counts all endpoint commands. Detail records
status, exact-content checks, relay exclusion, invalid-option I/O and callback
truth. Preserve the failing pre-repair witness. Default legacy Move remains
covered by the original recursive cases; this guard does not bound its commitment
map or qualify source-identity cleanup. The fake endpoint's native directory rename
atomically moves exact descendants, including nested/empty directories, while
preserving prefix siblings; it refuses existing targets rather than merging them.
`.MixedModeFixture` covers two selected items with Continue on error and configured
concurrency 1/4: one same-endpoint native success and one cross-endpoint refusal.
It requires partial aggregate status, exact contents, zero relay requests and
independent per-item committed/no-commit receipts. Value0 counts mutation commands
and value1 all commands, with elapsed time and explicit content/receipt checks.
The pre-repair baseline has 241 assertions, before the six mixed-bulk controls and
the strengthened directory fixture were added; compare equal scenario subsets.

`C0_CurlEntryLookupTruth` protects single/bulk Copy, Move, Delete and Rename
admission with 294 assertions (201 route/setup, 38 framing, four refused-endpoint,
12 fixture-serialization checks, three allocation/connect controls and 19 actual-
socket diagnostic controls, two endpoint-ownership, 11 cursor-lifetime and four
pure-parser controls).
The serializer oracles call the same fake
FTP generator as network LIST, with hand-authored byte expectations for empty/root/
nested listings, direct-child ordering, literal UTF-8/percent/arrow names, zero/
unknown/overridden sizes, a changed generation and unchanged custom malformed LIST/
NLST payloads. `FileOps.Curl.FixtureListing.Bytes` records actual/expected lengths;
`.Wide` records elapsed time, 64 renders and total bytes for 4,097 files. Every
render compares the complete payload; these fixture timings are not real-provider
throughput. Retain the same-byte generator baseline before optimizing it.
The generator borrows normalized UTF-8 parent/leaf slices only while their input
and state lock remain live, and formats into the owned output instead of creating
temporary rows. Compile-time checks also protect root/literal slices and CDUP's
self-prefix assignment. Keep the full scan, mutex boundary and override precedence;
do not cache or truncate generated listings to pass the performance gate.
Its 64 route scenarios cover literal percent/arrow
names, missing names, malformed rows before/after the match, duplicate matches,
out-of-parent names, oversized rows and stalled-LIST cancellation. Unsafe admission
must preserve every fixture object, emit no mutation command, and retain one honest
terminal receipt. Cancellation must return cooperatively within two seconds, not
through the fixture's emergency unblock. `FileOps.Curl.EntryLookup.TruthFixture`
records elapsed microseconds, mutation count and LIST count; `.ListCancel` records
cancel latency. `.WideRepeated` measures 16 complete lookups through a 4,097-row
Unix/DOS parent and verifies the requested size/name/timestamp. Keep same-machine
red/candidate rows; these are native-route witnesses, not host admission or real
FTP/SSH endpoint qualification. Existing wide Copy keeps its corpus and deadline.
The candidate also compares the requested timestamp exactly with the ordinary
one-row parser and requires 4,097 inspected rows per lookup. `.Retained` records
peak owned path bytes and fixed row/cursor metadata reservations (16 MiB/8 MiB
ceilings); opaque libcurl session storage is explicitly outside those metrics.
The pooled-cache candidate additionally requires exactly one passive listener and
15 reuses across each 16-query wide probe (`.Connections`), and proves a successful
same-endpoint lookup after every canceled admission. These strengthen the existing
route checks without relaxing cancellation time, exact contents or full-row validation.
Two additional Unix/DOS controls alternate each strict lookup with a blocking SIZE
probe. They require 16 successful probes, exactly two authenticated control sessions
and the same one-listener/15-reuse LIST counts. `.ControlSessions` records USER/SIZE
counts; pure lookup controls require one session. These detect cache destruction when
one easy handle alternates the blocking and multi APIs.

`FileOps.Curl.CursorLifetime` protects bounded read-ahead and early pool return:
retain parent/child cursor objects and unread parent entries while two small
listings share exactly one authenticated session. Empty completion, a 4,097-row
multi-buffer listing, abandoned paused transfer, cancellation, failed LIST final
reply and healthy reborrow are required controls. A small failed listing must
publish neither a cursor row nor an exact-lookup result. The callback/framing
seam drives the same buffer-fill/pause/whole-chunk-redelivery path, including all
existing byte splits. `FileOps.Curl.NativeCopy.ControlSessions` records source and
destination USER counts for each broad scenario; its S_OK denotes measurement,
not operation success or a general bound on opaque connection memory.

`FileOps.Curl.EntryLookup.ParseOnly` measures 64 scans of a 4,097-row Unix/DOS
listing with either a present final leaf or a missing leaf. It drives the existing
cursor callback/framing/parser without sockets or fixture LIST generation, counts
all 262,208 inspected rows, and retains only matches. Matched name/size/attributes
and timestamp must agree with the ordinary one-row parser. The reported time
excludes fixture construction and assertions, but includes the bounded simulated
transport-chunk copies and cursor construction; it is not pure CPU instruction
profiling or a replacement for native full-LIST/final-reply tests. Retain a
same-machine baseline before changing production row allocation/conversion.

The framing checks feed the existing callback/consumer with 1-byte, 7-byte and
`CURL_MAX_WRITE_SIZE` chunks: empty input, ignored rows, mixed Unix/DOS and UTF-8,
LF/CRLF/exact-limit/EOF records, overflow, embedded NUL and malformed late rows.
DOS numeric-size suffixes/fractions must fail, not become a successfully parsed
integer prefix; retain the initially failing regression witness before repair.
They also test every two-chunk split of the mixed stream and reject an oversized
callback before copying. `.Framing` records scenario/expected HRESULT, elapsed
microseconds, payload bytes in `value0` and submitted chunks in `value1` (summed
across the all-splits scenario). These are deterministic framing correctness
timings, including the test adapter, not transport throughput. Production changes
are compared using the unchanged `.WideRepeated` probes and full wide-Copy corpus;
do not change fake-server generation in the same causal comparison.

The refused-endpoint control keeps an exclusively bound, non-listening loopback
socket alive while cursor lookup, size probing and namespace dispatch return
`ERROR_CONNECTION_REFUSED`; it never tests a released port that another server
could acquire. Its test-only 5-second connection / 7-second operation limits allow
Windows to complete the refused connect (about two seconds on this machine),
instead of mistaking an earlier timeout for the expected refusal. Production and
traversal-test deadlines are unchanged. `.RefusedEndpoint` records the three actual
HRESULTs and owned port. The
unchanged read policy makes up to three size-probe attempts; namespace dispatch
still executes once. `FileOps.Curl.ConnectionFailure` records each failing
`CURLE_COULDNT_CONNECT` attempt at cursor completion, size probe and quote dispatch:
`value0` is the Curl code, `value1` is the available positive OS error (zero otherwise),
and detail carries only fixed phase, OS-error availability, numeric local/remote
ports and response code. Nonpositive ports are unavailable; reported positive
ports can belong to the established FTP control connection, not the failed data
socket. Do not infer failed-socket identity from them. Capture before handle reset/return.
No URL, profile, credential, control reply or error-buffer text is retained.
Expected refusal rows keep their failing HRESULT; the control's own result must
pass. These diagnostics do not authorize replay or prove that a mutation did not
commit, and machine-wide socket counts alone do not establish exhaustion.

Test-enabled builds additionally run a fresh wildcard/ephemeral bind probe only
after a captured 10048 connection failure. `FileOps.Curl.ConnectionAllocationProbe`
records the original OS error in `value0`, the observed allocated port (zero on
failure) in `value1`, probe duration and probe HRESULT. The socket closes before
emission; no listener, connection, reuse option, network-setting change, retry or
mutation is performed. The helper preserves the caller's Winsock error. Its exact
lookup control proves successful allocation, a nonzero port and error preservation.
This is a contemporaneous allocation observation, not proof of global exhaustion
or failed-tuple identity; concurrent port releases and the probe's small resource
cost must remain explicit. Normal production builds contain no probe.

After that observation, test-enabled builds compare one implicit-connect socket
and one explicitly wildcard-bound socket, each connecting once to a fresh owned
loopback listener. `FileOps.Curl.LoopbackConnectProbe` records mode (0 implicit,
1 explicit) in `value0`, observed local port in `value1`, duration/HRESULT, and a
fixed phase/stage. Stages distinguish listener/socket setup, nonblocking mode,
bind, connect, wait, completion and connected. Each pending nonblocking connect
uses one select wait of at most 100 ms and checks SO_ERROR; both socket owners
close and the caller's Winsock error is restored before return. No accept worker,
protocol command, original-peer retry, reuse option or machine setting is involved.
Two exact-lookup controls require connected/S_OK/nonzero port/error preservation.
The fresh peers and sequential observations cannot identify the original failed
tuple or exclude concurrent releases; timeout is diagnostic, not a repaired verdict.

Test-enabled native cursors install a suppressing libcurl debug callback before
enabling verbose output. It recognizes only the pinned loopback `connect to ...
port ... from ... port ... failed` diagnostic and retains numeric ports, never
addresses, error text, headers, credentials or payload. `SocketConnectFailure`
under `FileOps.Curl` records remote/local port in value0/value1 (zero when local
port is unavailable); fixed detail retains signed local port and cursor/connect
stage. This is the socket diagnostic, not CURLINFO_PRIMARY_PORT's control socket.
Its HRESULT is the existing normalized connection-failure category, not a parsed
OS error; correlate the same attempt with `ConnectionFailure` for the actual code.
The callback ignores internal handles, preserves Winsock error, and owns no sockets
or global observer. Cursor teardown disables verbose then clears borrowed callback
state before member destruction; ordinary pool reset remains mandatory. Eighteen
parser/privacy controls and one actual refused-cursor control guard the exact path.
The grammar deliberately accepts only remote 127.0.0.1 and local loopback/wildcard/
unavailable addresses; it is not generic real-provider connection qualification.
`FileOps.Curl.NativeCopy.EndpointPorts` records owned source/destination control
ports at scenario start, for correlation with the actual socket diagnostic. It
does not prove a data-listener mapping or concurrent port ownership after teardown.

The fake FTP endpoint factory sets SO_EXCLUSIVEADDRUSE before bind for both
listening and bound-only modes, retaining port-zero allocation on loopback only.
`FileOps.Curl.FakeFtp.ExclusiveEndpoint` guards the installed option and refusal of
a same-address/port competing SO_REUSEADDR bind in both modes. The challenger is
an owned negative-test socket, never a production socket policy. This ownership
contract does not promise immediate port reuse after close or qualify a broad run.

`C0_CurlCopyTraversalTruth` has 142 assertions: 46 native scenarios and three
passive-listener lifecycle checks, plus Winsock setup. The host's exact-count
guard must agree with the current plugin corpus (294 lookup / 142 Copy) and
report expected/observed counts on failure; S_OK alone is not complete coverage. Both
singular/bulk Copy APIs run serial/parallel shallow, literal-name, malformed,
out-of-root, overlong, blocked LIST, late refusal/cancel and depth-128/129 cases;
both Move APIs additionally prove accepted depth 128. Two wide Copy cases publish
4,097 real files, beyond the aggregate queue ceiling. Wide-copy fixture deadlines are
180 seconds in Debug/Release and 360 seconds in ASan Debug; the explicit instrumented
allowance preserves the identical corpus and bounds fixture execution. It is not a
throughput acceptance threshold. Other shape deadlines and the two-second observed
cancellation bound remain unchanged. Exact source/destination
bytes, outside siblings and terminal status are checked independently. A late
failure must retain an earlier subtree's published file; an invalid first listing
must not create children. Copy must not fabricate a destructive-source receipt.
`FileOps.Curl.NativeCopy.TruthFixture` records elapsed time, source file count and
publication count with status, endpoint health, contents and completion truth in
detail. `.ListCancel` measures return after an observed wait (under two seconds);
owned emergency endpoint shutdown never passes. `.Begin` is diagnostic only.
When diagnosing local port allocation with an OS socket snapshot, bind the sample
to the observed selftest PID, creation time and exact executable. Include wildcard
`BOUND` states when assessing retained sockets, record the collection interval,
and retain the raw rows. Do not attribute PID-0 TIME_WAIT rows to that process or
infer exhaustion/leaks from one point-in-time sample. Such observation may affect
fixture timings; disclose it instead of claiming a clean throughput comparison.
Archive baseline failures before changing the recursive implementation. These
fixtures do not establish real FTP/SSH quotas or bounded Move commitment storage.
Provider `.Traversal` records duration, peak ancestry frames and queued files;
`.Retained` records owned frame/batch path bytes and metadata reservations, excluding
the separate Move commitment map and opaque transport memory. `.FileFailure`
identifies the content-free failed publication stage without logging paths or
credentials. Wide/deep timings must retain baseline failures, not imply throughput
improvement from a smaller or incomplete corpus.
Wide Copy runs parallel before serial to distinguish transport failures from
preceding long-run pressure while retaining the identical 4,097-file corpus.
Copy completion consistency is checked against the actual call HRESULT; the
independent scenario assertion still requires the expected result and exact bytes.
`FileOps.Curl.UnmappedTransportFailure` retains only the otherwise-unmapped CURLcode
in value0; it does not change the returned HRESULT or log endpoint information.
The wide corpus also repeats source and staged-size probes thousands of times:
metadata output must remain independent of stdout and retain no response bytes.
The metadata sink has compile-time zero/ordinary/overflow count checks; native
integration must still prove all 4,097 publications and both endpoint states.
The fake FTP server reuses a control session's passive listener only after its
advertised data connection has been accepted. A fresh EPSV/PASV retires an unused
listener, including a queued connection behind a refused RETR/LIST. A raw-client
regression leaves that stale connection open, then requires two exact transfers
on the replacement/reused listener. Empty idle control sessions close with 421;
partial command and active data timeouts remain failures. Wide Copy additionally
requires at most 64 listener creations per endpoint, exactly 4,097 accepted RETR
transfers at the source and STOR transfers at the destination, no NLST fallback
or abandoned passive request, and complete data-command accounting. For each
endpoint, accepted first-use/reuse totals must equal listener creation/reuse totals;
every LIST/NLST/RETR/STOR command must have one corresponding accepted transfer.
Multiple healthy sessions legitimately have fewer than 4,096 payload reuses:
N transfers over K sessions have N-K reuses. Do not hard-code one-session math.
`.PassiveListeners` reports aggregate creations/reuses and `.PassiveRetirements`
reports abandoned listeners/idle closures. These are fixture resource checks,
not production connection quotas; corpus, operation deadline and bytes stay fixed.
`.EndpointPassiveListeners` separates both endpoints and `.DataTransferUses`
separates LIST/NLST/RETR/STOR first uses (`value0`) and reuses (`value1`). The refused
RETR control proves that an unaccepted transfer increments neither count; pure
and mixed lookup controls require one first-use and 15 reused LIST transfers.

`.FilePhase` emits nine bounded summary rows per native tree after all file workers
drain: phase wall microseconds in duration and visits in `value0`, without paths
or per-file rows. The phase accumulator is included in owned metadata reservations.
These timings include the work and waits between the named phase boundaries;
parallel file timings overlap and are neither CPU time nor whole-operation elapsed
time. `.ListPayloadBuild` reports the fixture's cumulative listing serialization
and mutex-wait microseconds, LIST count in `value0`, and generated payload bytes in
`value1`. Keep client phase and server generation metrics separate: speeding up a
fake server does not establish a production parser or real-server improvement.

Compare selftest startup-policy coverage emits `compare.selftest.host_selection`
with the number of deterministic filter checks in `value0` and required-host
failure-classification checks in `value1`. The owned-host context
case emits `compare.selftest.host_context` for one non-secret missing-profile
metadata request; it must reach profile lookup without credentials or network I/O.
Archive the existing `App.Startup.*` timings and startup trace with both an exact
pure/headless control and a host-dependent run. Host setup cost is not remote I/O
performance, and enabling the required host must not disable no-activation policy.
Hosted Compare executes within `WM_CREATE`; its enclosing startup/InitInstance
timings include selftest work and cannot be compared as pure UI startup latency.

Create Directory admission emits one `fileops.create_directory.admit_us` row per qualification
attempt. Detail is bounded to `accepted-local`, `accepted-provider`, or `rejected`; `value0` is the single candidate count,
and `value1` is only the leaf UTF-16 length; no parent, candidate, provider path, or root identifier
is emitted. The deterministic capability-matrix scenario covers accepted Local and Dummy profiles,
read-only rejection, empty-root contract failure, provider-separator joining, and Local
`GLOBALROOT` rejection. This is a prompt/admission-latency and fail-closed correctness baseline, not
a directory-creation-throughput claim.

Batch Rename separates UI task publication from worker qualification. The complete immutable-capture,
task/card publication, completion-registration, and worker-release call emits
`batchrename.admission.ui_thread_us`; 1,024 Local and deterministic delayed-provider rows have a 50,000
us reference-profile ceiling. This budget was derived from the E4 cold attested full-call maximum of
28.926 ms with about 1.7x headroom. `batchrename.admission.wall_us` records end-to-end qualification;
`batchrename.admission.worker_wall_us` spans worker-only capability lookup,
provider-parent derivation, no-follow source binding, one-domain proof, duplicate-object rejection,
acyclic schedule construction/cycle rejection, validation, and immutable plan publication. The same attempt emits
`batchrename.admit.us`, `batchrename.admission.ui_capture_us`, capability/bind call counts, retained
bindings (required to be zero), and cancel-request-to-terminal latency. The deterministic
`cmd_pane_batchRename_admission_worker_owned_cancellable` scenario covers 1, 64, and 1,024 Local and
delayed rows and proves no row-proportional provider call occurs on the UI thread.

Change Case uses the protected `cmd_pane_changeCase` discovery scenario and
`cmd_pane_changeCase_central_rename_plan` admission/execution scenario. The deterministic Local
discovery fixture contains 1,024 files and measures discovery/planning separately from central
RenamePlan admission and execution. The self-test emits `Commands.SelfTest.E5.ChangeCaseDiscoveryUs`;
production emits
`changecase.discovery.us`, `changecase.admission.ui_thread_us`, the remaining
`changecase.admission.*` counters, and `changecase.execute.us`. Metrics contain aggregate counts and
bounded scenario labels only; they never contain source paths or leaf names. The scenario requires
exactly 1,024 immutable operations, at least 1,024 scanned entries, zero mutation during planning,
and completion inside the timeout-scaled five-second correctness ceiling. Release closeout runs use
five independent processes per scenario and archive the receipts under
`Specs/TestRuns/<MachineHash>/Commands/2026-09-01_e5_change_case_{discovery,central_plan}_release/`.
The first Release characterization recorded discovery p95 `3,206,236 us`, central UI-admission p95
`26,216 us`, central worker-admission p95 `2,791 us`, and central execution p95 `2,378 us`; every
sample remained below `5,000,000 us`. Because E5 introduces the first separated Change Case
instrumentation, these archives are characterization baselines rather than before/after optimization
claims.

`batchrename.schedule.build.us`, `.layers`, `.cycle_members`, and `.retained_index_bytes` measure the
provider-free immutable schedule. `batchrename.execute.us` plus the rows/completed/failed/mutations
counters measure exact conditional-mutation dispatch; `batchrename.execute.journal_writes` is a
retirement drift alarm and MUST be zero. The deterministic scheduler scenario uses 1,024 independent
rows in one runnable layer, requires exactly one mutation callback per row and zero journal writes,
and applies a timeout-scaled five-second ceiling. Direct cyclic cases require exact cycle-member
reporting, `ERROR_CIRCULAR_DEPENDENCY`, zero mutation callbacks, and unchanged source bytes.

R7-A10 Release evidence uses five independent processes for the exact
`cmd_pane_batchRename_execution_engine_large_independent_perf` case. The accepted baseline archive is
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r7_a10_acyclic_no_journal_baseline_release/`; its
`batchrename.execute.us` p95 is 197,661 us and the candidate regression ceiling is 296,492 us (1.5x).
The matching candidate archive is
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r7_a10_acyclic_no_journal_candidate_release/`.
Its five execution samples are 1,801, 1,832, 1,723, 1,790, and 1,776 us (p95 1,832 us); its schedule
build samples are 1,257, 1,292, 1,212, 1,281, and 1,258 us (p95 1,292 us). Every candidate process
proves 1,024 mutations, zero journal writes, and a duration below the five-second scenario bound.
These are latency/boundedness baselines, not a throughput-improvement claim; no path, leaf, identity,
or provider payload is emitted.

Batch Rename destination-collision preparation emits
`batchrename.preview.collision_index_retained_bytes`. The deterministic
`cmd_pane_batchRename_collision_name_index_memory_gate` scenario builds 65,536 names of 240 UTF-16
units, requires provider-canonical folded/exact lookup semantics, and keeps the complete index at or
below 64 MiB. The index retains canonical keys only; duplicating every raw listing name is forbidden.
This is a retained-memory and comparison-semantics gate, not an admission-throughput claim.

R7-A15 protects provider-name consumption with
`cmd_pane_batchRename_engine_large_preview_perf`: five independent test-enabled x64 Release
processes build the same 10,000-row preview while the candidate applies a deterministic typed
provider name contract. Candidate p95 must be no more than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)` and every sample must remain below
5,000,000 us. In addition to existing build-plan rows, the candidate emits
`batchrename.preview.provider_name_validation.us`, `.provider_name_queries`,
`.provider_name_arena_fallbacks`, `.provider_name_rejected`, and
`.provider_name_key_bytes`, plus `.destination_directory_listings` and
`.collision_index_retained_bytes`; values are aggregate counts/bytes only and never retain provider
paths or leaf text. `file_system_provider_name_policy_ingresses` is the offline correctness
control for Inline F2, Batch preview/worker drift, Change Case, and Create Directory, including
exact zero-mutation failures and one-enumeration suffix selection. The 65,536-name/64 MiB gate
remains independently blocking over provider-canonical keys.

Batch Rename journal write/load/recovery metrics are retired with the journal implementation. New
production emission of `batchrename.journal.*` or `batchrename.recovery.*` is a contract regression;
the only retained journal-named metric is the zero-valued
`batchrename.execute.journal_writes` drift alarm.

Interrupted Move bookkeeping emits `fileops.move_breadcrumb.persist.us` for atomic `create`,
`advance`, and `terminal` transitions. `value0` is the bounded durable-phase enum and `value1` is the
serialized byte count; paths, roots, identities, and provider payloads are never emitted. Startup
loading emits `fileops.move_breadcrumb.load.us` with bounded detail `loaded` or `partial`, interrupted
record count in `value0`, and scanned record count in `value1`. The deterministic scenario persists
17 qualified roots as a total plus exactly 16 fixed samples, proves failed phase persistence rolls
memory back to the last durable phase, rejects a malformed sibling without deleting it, and verifies
that load/dismissal do not mutate source or destination sentinels. The same case projects the record
through the host and popup, verifies pane-faithful Open-location state, the Indeterminate result,
strategy/root summary, absence of synthetic Issues actions, and Dismiss-only JSON acknowledgement.
These are startup/durability and
constant-retention baselines, not Move-throughput or recovery-success claims.

Artifact startup discovery emits `fileops.artifact.registry.load.us` with bounded detail `loaded` or
`partial`, valid claim count in `value0`, and canonical record count in `value1`. The scenario includes
a malformed sibling beside valid claims and proves partial quarantine does not discard the valid
snapshot. `fileops.artifact.index.visibility.us` records the never-hidden indexed-query latency with
returned candidate count in `value0` and required artifact-shape count in `value1`. Classification
and Search/index visibility measurements must contain counts only—never
paths, identities, intended names, or provider payloads—and make no throughput-improvement claim.
Folder enumeration emits `fileops.artifact.folder_projection.us` and
`fileops.artifact.folder_projection.probe_candidates`; Find emits
`fileops.artifact.find_projection.us` and `fileops.artifact.find_projection.lookup_rows`. Actual
projection binding emits `fileops.artifact.capture.bind_count`. After journal retirement production
has no durable claims or keyed claim index. The deterministic
`cmd_fileops_artifact_name_shape_projection_scaling` scenario records
`fileops.artifact.{folder|find}_name_shape_lookup_10k.us` for 10,000 ordinary names and 100 recognized
Possible names using the same `HasPossibleArtifactName` function as FolderView and Find. `detail`
contains only those counts, `value0` is the ordinary-name probe count, and `value1` is the Possible-
name probe count. Every run requires `value0=0`, `value1=100`, and `hr=0`. The scenario characterizes
bounded name-shape filtering only; it must not manufacture private claim keys, hard-code endpoint
results, or be cited as production claim-index evidence.
Typed task admission records `fileops.artifact.touch.capture.us`; worker revalidation records
`fileops.artifact.touch.revalidate.us`; external Shell/viewer/editor launch guarding records
`fileops.artifact.touch.external.us`. `value0` is the bounded exact-set item count. The deterministic
external case proves ordinary paths remain prompt-free, artifact Cancel is the safe default, and an
accepted exact identity is reloaded/rebound/revalidated before launch. These are warning-path
latency and boundedness baselines, not copy, launch, or search throughput claims.

Direct-admission prompt teardown evidence uses the test-enabled Release Commands case
`cmd_fileops_start_operation_artifact_prompt_shutdown_during_nested_pump`. It emits
`FileOps.SelfTest.ArtifactPromptNestedShutdownUs` with detail
`start-operation-direct-admission`, `value0=1` when `FileOperationState` remained alive until the
nested prompt unwound, `value1=1` when the selected source remained untouched, and
`hr=0x8007045B` (`ERROR_SHUTDOWN_IN_PROGRESS`). Capture five independent processes beneath
`Specs/TestRuns/<MachineHash>/Commands/`; every sample must be below 2,000,000 us and all three
correctness fields must match exactly. This is a correctness and bounded-teardown characterization;
it makes no baseline or latency-improvement claim.

Single-pass File Operations discovery emits one bounded terminal metric set per task:
`FileOps.Discovery.OpenUs`, `CallbackCount`, `CallbackUs`, `LockWaitUs`, `MaxQueueDepth`,
`StarvationCount`, `FirstMutationBeforeClose`, `FirstMutationUs`,
`BytesCompletedWhileOpen`, `MutationsCompletedWhileOpen`, `SkipReleaseUs`, `GrowthAfterClose`, and
`Closed`.
`FirstMutationUs` is elapsed time from discovery start to the first host-proved completion observed
while discovery remains open. `BytesCompletedWhileOpen` and `MutationsCompletedWhileOpen` are
saturating task-local totals; they are measurement only and must not feed scheduling. Provider
progress contributes real deltas. Exact Local permanent Delete contributes its typed `Removed`
selected-root result once, with that root's discovered bytes, because that exact path intentionally
has no provider progress callback. `MaxQueueDepth` is compared with the fixed
discovery-ahead target 256, not with tree size. `FirstMutationBeforeClose=1` is required in the
wide-tree overlap scenario and proves that discovery did not become a preflight barrier; it is not a
universal requirement, because a selected root that is one known file closes its exact total before
any payload moves. `GrowthAfterClose` counts discovery reports that added work after the task's
totals were already final; closure is one-way, so every scenario requires exactly 0. The
`DiscoveryScope_` family carries the paired witnesses: with the leaf, nested-cleanup and
provisional-rename rules disabled it reads 5 and 6 on the managed-directory and rename-merge
scenarios and the leaf checkpoint is never reached, and it reads 0 with them in place.
`Phase5_DiscoveryCancelLatencyLocal` records bounded cancellation while traversal and transfer are
both live; `Phase5_DiscoverySkipContinues` and
`Phase5_SwitchParallelToWaitDuringDiscovery` prove the one-way Skip state and its separation from
global Queue admission. The popup source contract proves that totals/ETA stay indeterminate until
closure. These are scheduler/correctness baselines, not throughput-improvement claims.

`R4A19_DiscoveryMeasurementFacts` is the deterministic Local measurement fixture: Copy uses
128 files of 256 KiB at configured concurrency 4; permanent Delete uses 1,024 files of 1 KiB at
configured concurrency 8. Both must complete below 30 seconds, close discovery, retain at most 256
ready records, and prove the first completion before closure and within 1 second. Copy must report
nonzero open-discovery bytes/mutations and exact 32 MiB/128-file output. Delete must remove the exact
selected tree, report exact 1 MiB/1,024-file discovery, exactly 1 MiB and one selected-root mutation
while open, and zero provider progress callbacks. Release qualification uses five independent
test-enabled x64 processes per side. Candidate total-duration p95 must be at most
`max(1.35 * baseline, baseline + 250,000 us)`; first-mutation p95 must be at most
`max(1.50 * baseline, baseline + 100,000 us)` and 1 second. A baseline missing the observation fact
is retained as expected RED evidence and is never relabeled green.

The same `FileOpsFamily_R4A19DiscoveryBaseline` process also owns the unchanged-scheduler provider
and independent-volume qualification. The deterministic Dummy control copies 8 directories with
8 files of 8 KiB through `latencyMs=5`, `streamChunkLatencyMs=2`, and a 1-MiB/s virtual limit. The
test-only fake-MTP control copies 8 ordinary-name files of 64 KiB through one serialized session
with a 25-ms backend delay, then cancels a second task while discovery and mutation are both live.
The fixed-volume control copies four 4-MiB files independently on marker-authorized C and D roots at
16 MiB/s, then repeats both concurrently. It emits
`FileOps.SelfTest.R4A19.{ProviderControl,ProviderCancel,IndependentVolumes,IndependentVolumeCIsolatedUs,IndependentVolumeDIsolatedUs}`.
The fake-MTP fixture, whose route the host bridge traverses, requires first mutation before discovery
closure and within 1 second, nonzero open-discovery byte/mutation facts, zero discovery starvation,
and exact bytes/items. The Dummy fixture runs on the provider's direct route, where Dummy already
holds its subtree before it mutates and therefore closes its record on the exact total before the
first byte moves; it requires that closed record (512 KiB, 64 files, 9 directories), zero
open-discovery bytes/mutations, zero growth after closure, zero discovery starvation, and exact
bytes. Fake MTP requires maximum backend concurrency exactly 1 and cancellation within 500 ms. Concurrent C/D
wall time must be at most 1.35 times the slower isolated control and both tasks must have overlapping
progress. Ready queue, bridge admission queue, retained entries, queued path bytes, metadata, and
limit-hit gates remain 256, 256, 4,096, 16 MiB, 8 MiB, and zero. Qualification is five independent
test-enabled x64 Release processes. When an evidence-only card produces no scheduler candidate, it
must report the absolute/p95 characterization and keep the relative candidate/JIT throughput gate
unclaimed rather than fabricate a second timing side; the behavioral Skip/JIT cases remain mandatory.
Repository archives may retain only `[Ff]ileOps.*` perf rows when the complete process perf stream
would exceed the TestRuns per-file ceiling; the unfiltered governed source run remains under the
exact `X:\RedSalamander.Perf\runs\<run-id>` root named by the receipt.

MTP provider memory evidence uses `mtp.writer.private_payload_buffer_high_water_bytes` (required
zero), `mtp.writer.memory_high_water_bytes` (bounded upload buffer),
`mtp.reader.reusable_request_bytes` (at most 8 MiB), `mtp.verify.memory_high_water_bytes` (at most
8 MiB), and `mtp.transfer.bridge_buffer_high_water_bytes` (at most 4 MiB). The deterministic
`mtp_public_writer_and_reader_memory_is_bounded` case transfers 12 MiB through the fake backend,
writes in 1-MiB calls, requests one 12-MiB read, and proves the provider chunks it through bounded
reusable state. Fake-backend object bytes are fixture storage and are excluded from the production
provider-buffer claim.

MTP exact-recovery identity evidence uses the test-enabled Release case
`mtp_r0c_exact_recovery_identity`. It enumerates and formats 4,096 full PUID identities, including
one pair that collides under the retired case-folded hash, and resolves one recovery record with one
backend enumeration vector plus constant match state. The candidate emits
`FileOps.Mtp.R0c.{ExactIdentity.SelfTestUs,RecoveryEnumeratedCount,FullIdentityCompareCount,AmbiguousMatchCount,QuarantinedJournalCount}`.
Capture five sequential independent processes per side beneath
`Specs/TestRuns/<MachineHash>/FileOpsMtp/`. If the RED baseline fails before the candidate metric can
be emitted, retain that failure and compare the same runner's whole-case duration instead of
fabricating a metric. Acceptance requires candidate whole-case p95 no greater than
`max(1.50 * baseline p95, baseline p95 + 50 ms)`, every candidate metric below 2,000,000 us, exact
4,096 enumeration/compare counts, and no second PUID/path candidate vector. This is deterministic
identity/recovery resource evidence, not live-WPD latency.

R0d receipt/capability-honesty evidence uses the test-enabled Release case
`FileOps_ProviderCapabilityMatrix`. Each process performs 4,096 receipt classifications, 8,192
flag-aware host admission decisions, and 12,288 capability queries while retaining one capability
document and constant typed-result state. It emits
`FileOps.SelfTest.R0d.{CapabilityReceiptHonestyUs,CapabilityQueryCount,AdmissionDecisionCount,ReceiptClassificationCount}`
once per process. Capture five sequential independent processes per side beneath
`Specs/TestRuns/<MachineHash>/FileOps/`. Acceptance requires candidate p95 no greater than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)`, every duration below 2,000,000 us, and exact
12,288/8,192/4,096 counts. Debug and ASan Debug correctness runs retain the same loop, counts and
validity assertions with a five-second diagnostic ceiling, matching the adjacent R0e profile;
they do not establish Release performance qualification. A RED baseline may fail the receipt assertion after emitting the measured
loop; retain that failure rather than substituting a synthetic control. This is in-process
admission/classification and constant-resource evidence, not live-provider latency.

R0e cancellation-route containment evidence uses the same test-enabled Release case. Each process
performs 4,096 operation-specific capability parses and 4,096 fixed-volume Local admission
decisions, then attempts one uncontained never-returning-provider admission. It emits
`FileOps.SelfTest.R0e.{RouteContainmentUs,CapabilityParseCount,AdmissionDecisionCount,RejectedUncontainedCount,ProviderCallCount}`
once per process. The uncontained witness must be rejected before task/card/worker creation and its
provider-call count must remain zero; the test must never invoke the blocking method. Capture five
sequential independent processes per side beneath `Specs/TestRuns/<MachineHash>/FileOps/`.
Acceptance requires candidate p95 no greater than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)`, every Release duration below 2,000,000 us,
exact 4,096 parse/decision counts, exactly one rejected witness, zero provider calls, and constant
retained state. Debug correctness runs use a wider five-second diagnostic ceiling and are not
production performance qualification. This is admission/containment evidence, not live-network,
SMB, provider-watchdog, or shutdown-throughput evidence.

R0e-OR1 fixed-volume Local reader cancellation evidence uses the test-enabled Release Commands case
`cmd_pane_fileops_local_blocked_reader_cancel_is_bounded`. One real Local `Win32FileReader` issues an
overlapped one-byte read against a connected outbound byte pipe whose server produces no bytes. The
case flips the same operation-control callback captured by an exact bound reader and emits
`FileOps.SelfTest.R0eOr1.LocalBlockedReadCancelUs` with abort-check count in `value0`, bytes reported
after Cancel in `value1`, and the terminal reader HRESULT. GREEN requires kernel-pending observation,
at least two total operation-control checks, `ERROR_CANCELLED`, zero bytes, cleanup of the reader
thread/request/pipe, and every sample plus p95 below 500,000 us. A 500-ms fail-safe server close keeps
the RED fixture bounded but cannot satisfy the pending/cancel assertions.

The same process reads one deterministic 8-MiB fixed-local payload completely through both the
legacy and exact-bound Local readers, seeks both to offset 17, and compares the replay bytes. It emits
`FileOps.SelfTest.R0eOr1.OrdinaryReadUs` with the combined 16-MiB byte count in `value0` and seek-match
truth in `value1`. Capture five independent test-enabled x64 Release processes before and after under
`Specs/TestRuns/<MachineHash>/Commands/2026-09-01_r0e_or1_local_reader_cancel_{baseline,candidate}_release/`.
Candidate ordinary-read p95 must be no greater than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)`, with exact bytes and seek replay in every
sample. These rows prove Local pending-read cancellation and ordinary reader overhead only; they do
not qualify disk throughput, writers, remote providers, a generic UI join deadline, or routes still
classified `uncontained`.

R2 typed-route evidence uses the same test-enabled Release case
`FileOps_ProviderCapabilityMatrix`. Its protected loop performs 4,096 canonical
typed queries over the offline Dummy route and emits
`FileOps.RouteFacts.{QueryUs,QueryCount,JsonParseCount,TypedQueryCount,ArenaFallbackCount,RejectedCount,ResultBytes}`.
Acceptance requires exact query/typed counts of 4,096, zero JSON parses, zero normal
arena fallbacks, zero rejected results, and constant result bytes. Capture five
independent processes per side beneath
`Specs/TestRuns/<MachineHash>/FileOps/2026-09-01_r2_typed_route_{baseline,candidate}_release/`.
Candidate p95 must be no greater than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)`, and every sample must remain
below 2,000,000 us. The provider contract suite separately covers all eight shipped
DLLs/14 IDs, record/arena boundaries, allocation failure, malformed output, full-ID
peer/name/join/collision rules, and receipt classifications. This is typed
query/validation and constant-resource evidence, not provider or network latency.

R1a terminal/clipboard-truth evidence uses the test-enabled Release case
`Phase10_ClipboardAdmissionAndRetainedActions`. Each process initializes and phase-finalizes 4,096
selected-root builders, attempts one duplicate terminal store, and evaluates 1,024 no-provider
readiness/clipboard gate decisions. It emits
`FileOps.SelfTest.R1a.{TerminalFunnelUs,BuilderCount,TerminalStoreCount,DuplicateStoreRejectCount,ClipboardGateDecisionCount,ProviderCallCount}`
once per process. Capture five sequential independent processes per side beneath
`Specs/TestRuns/<MachineHash>/FileOps/`. Acceptance requires candidate p95 no greater than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)`, every duration below 2,000,000 us, exactly
4,096 builders and terminal stores, one rejected duplicate, 1,024 decisions with exactly 512
mutation permits, zero provider calls, no terminal `E_PENDING`, and constant retained state. This is
result-funnel/gate overhead evidence, not provider or clipboard-service throughput.

R1c cleanup-truth evidence uses the test-enabled Release case
`FileOpsFamily_ClearflowPhase07_ProviderMatrix`. Each process reduces 4,096 provider-free committed
primary results carrying exact `Retained` cleanup debt and emits
`FileOps.CleanupDebt.{ReduceUs,ResultCount,CompletedCount,RetainedDebtCount,IndeterminateCount,RetainedBytesPerResult}`.
Capture five sequential independent processes per side beneath `Specs/TestRuns/<MachineHash>/FileOps/`.
Acceptance requires candidate p95 no greater than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)`, every duration below 2,000,000 us, exactly
4,096 results/Completed/Retained, zero Indeterminate, and a fixed 256-byte retained-results bound per
input. This measures typed result reduction and retained storage, not provider cleanup or network I/O.

R1b Change Case task-transport evidence uses the test-enabled Release case
`cmd_pane_changeCase_task_payload_truth`. Each process resolves 4,096 tokenized task updates through
one shared creation receipt without invoking a provider. It emits
`Commands.SelfTest.R1b.{ChangeCaseTaskDispatchUs,PayloadCount,TaskCreateCount,TaskUpdateCount,ProviderCallCount,OutstandingReceiptCount}`
once per process. Capture five sequential independent processes per side beneath
`Specs/TestRuns/<MachineHash>/Commands/`. Acceptance requires candidate p95 no greater than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)`, every duration below 2,000,000 us, exactly
4,096 payloads, one task creation, 4,095 updates, zero provider calls, and zero outstanding receipts
after drain. This measures task-receipt and payload-reduction overhead, not traversal or provider
rename throughput.

R1d common-preparation evidence uses the test-enabled x64 Release family
`FileOpsFamily_R1dPreparingLifecycle`. Its protected offline loop projects 4,096 top-level Local
transfer items and 8,192 selected-root interlock scope facts without descendants or content I/O. It
emits `FileOps.Preparing.{BuildSnapshotUs,SelectedRootReadinessUs,UiAdmissionUs,SelectedRootCount,ScopeFactCount,StrategyFactCount,CopyOnlyCount,MutationBeforeReadyCount,ClipboardBeforeReadyCount,BreadcrumbBeforeReadyCount,RetainedBytes}`.
Capture five sequential independent processes per side under
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r1d_preparing_{baseline,candidate}_release/`.
Candidate p95 for `BuildSnapshotUs` and `SelectedRootReadinessUs` must be no greater than
`max(1.50 * baseline p95, baseline p95 + 50,000 us)`; `UiAdmissionUs` p95 may regress by at most
25,000 us and must not include synchronous popup construction/show work. Every timed sample must
remain below 2,000,000 us, selected-root/scope/strategy facts must
remain exactly 4,096/8,192/1 for the protected plan, retained snapshot state must remain below
64 MiB, and all three before-Ready counters must be zero. Behavioral cases in the same family prove
routine no-prompt Copy, common-gate slow/fail/cancel truth, clipboard-after-preparation ordering, and
dense source-retained terminal results. This is top-level preparation and immutable-fact overhead,
not recursive traversal, provider latency, or content throughput evidence.

S3 multipart evidence includes four simultaneous known-size 4-KiB writers and requires their
aggregate payload reserve to remain at or below 16 KiB. Multipart concurrency and the existing
four-payload/256-MiB worst-case ceiling remain unchanged.

S3 ordinary virtual-folder Delete evidence uses the test-enabled Release command
`PluginContractTests.exe --s3-r0b-delete-selftests`. Each independent process deletes 4,096 listed
4-KiB fake objects. The operation emits the coalesced
`FileOps.S3.VirtualFolderDelete.{ElapsedUs,PassCount,ObservedObjectCount,ConditionalRequestCount,RevisionMismatchCount,ResidualObjectCount}`
family and the fixture emits `FileOps.S3.VirtualFolderDelete.SelfTestUs`. Acceptance requires one
serialized ETag-conditioned request per observation, two listings in the no-writer case, zero
residuals, no duplicate bare-key vector, every sample below 2,000,000 us, and candidate p95 no
greater than `max(1.50 * baseline p95, baseline p95 + 50,000 us)`. Archive five independent
processes per side beneath `Specs/TestRuns/<MachineHash>/FileOpsS3/`; this is deterministic
request-shape/resource evidence, not a live endpoint latency claim.

FBP-10 bridge-pipeline evidence uses the deterministic
`Phase11_BridgePipelineDummyToDummyPerf` File Operations case in test-enabled Release. It runs the
same four 32-MiB Dummy inputs first with the reader/writer pipeline disabled and then enabled, with
30-ms provider chunk latency, a 4-MiB configured buffer, and the expected 8-MiB endpoint-resolved
buffer. The run emits `FileOps.SelfTest.BridgePipelineBaseline`,
`FileOps.SelfTest.BridgePipelineCandidate`, and `FileOps.SelfTest.BridgePipelineImprovement` and
requires the candidate to take at most 80 percent of the baseline duration. The accepted
2026-08-27 same-machine run measured 130.484 seconds versus 0.250 seconds and is archived under
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-27_011215/`. This closes the measure-first decision by
retaining the current bounded pipeline; it does not authorize more than 16 admitted pumps, more
than 256 MiB of aggregate host pump buffers, or another thread/buffer ownership layer.

Local permanent recursive Delete additionally emits `FileOps.DeleteTraversal.MaxDepth`,
`MaxBatchEntries`, `MaxTerminalFailures`, and `MaxFailurePathBytes`. The observed values must remain
within 128 levels, 256 queued children, 4,096 retained terminal failures, and 16 MiB of retained
failure-path text. `Phase6_DeleteBytesMeaningful` proves discovery and completed-byte accounting from
the same traversal; `Phase7_ParallelDeleteKnobs` and
`Fairstream_ParallelDeleteContinuesPastLockedChild` protect bounded parallel execution and
continue-on-error convergence. These rows are bounds/correctness evidence, not delete-throughput
improvement claims.

Recycle Bin batch validation records `FileOps.SelfTest.RecycleBinBatchBaseline`,
`RecycleBinBatchCandidate`, and `RecycleBinBatchImprovement` as live-environment timing evidence.
Those elapsed times are advisory because `IFileOperation` completion includes variable Windows
Shell broker, Defender, and indexing latency; a single baseline/candidate wall-clock pair is not a
hard regression gate. Test-enabled Debug and Release instead hard-gate the process-local route:
the fixture uses the same explicit `BulkItems` admission mode as the pane Recycle command,
batch size `1` performs zero batch calls, 384 sibling inputs at batch size `500` perform exactly one
384-item batch, and 768 sibling inputs at batch size `256` perform exactly three batches with 768
requested and observed items, zero failed items, zero fallbacks, and maximum batch size `256`.
Each invocation uses run-unique leaf names so repeated cases and pre-existing Recycle Bin contents
cannot divert the measurement into the conflict-prompt path.
`FileOps.SelfTest.RecycleBinBatchRoute` archives the route evidence, while
`RecycleBinBatchTimingAdvisory` records a candidate sample outside the retained noise band without
turning externally controlled Shell latency into a correctness failure.

Executable transfer routing uses `fileops.operation.strategy`. The bounded `detail` is
Copy/Move x Native/Managed/Copy-only x same-root/cross-root, `value0` is the child-plan count in
that bucket, and `value1` is the selected top-level item count. The task emits at most one row per
non-empty bucket and never includes provider paths or emits per item/discovered descendant. A Move
admitted through the Copy export/import pair must appear as `move.copy-only.*`; it must not appear
as Managed and must retain its source.

`fileops.plan.admit_us` performs no top-level shape probe and acquires no `IFileSystemIO` interface
for shape; Native admission is endpoint plus capability only and never enumerates. A destination
directory that appears after admission is first attempted through the provider's one-mutation
Native boundary; only its known non-commit may emit
`fileops.execute.native_directory_race_requalified_after_noncommit`, and the same item then
continues as a rename merge with no copy. There is no missing-fallback outcome. Deterministic
coverage must exercise the rename-merge continuation.

`FileOps.SelfTest.SameVolumeTreeMoveIsRename`, `FileOps.SelfTest.SameVolumeRenameMerge`, and
`FileOps.SelfTest.NativeDirectoryRaceContinuesAsRenameMerge` are the same-endpoint Native rows: zero
bridge bytes, unchanged `FILE_ID_INFO` for every relocated descendant, `move.native.same-root` in
`fileops.operation.strategy`, and, for the merge, one conflicting child prompted and one skipped
child retaining its ancestors as `Moved; source folder kept`.
`FileOps.SelfTest.CopySkipLinksKeepsPlaceholders` proves that Copy `Skip links` skips name-surrogate
links only. `FileOps.SelfTest.LoopbackShareMoveIsRename` proves the `local-win32-smb` profile takes
the same Native route when both paths live on the loopback administrative-share alias of the sandbox.
Managed-route proofs that need two Local endpoints on one machine
(`Phase10_MetadataPreservationAndSourceRetention`, the Move sub-steps of `Phase10_ContentVerification`)
target that alias as their destination and skip with a log line when it is unreachable. The archived throughput comparison for a same-volume directory Move (Managed copy
baseline versus rename candidate) lives under `Specs/TestRuns/<commit>/FileOps/`.

`FileOps.SelfTest.LocalProviderCrossVolumeMoveNoFallback` is the Local native-strategy boundary:
when a second writable fixed volume exists, the direct provider Move returns
`ERROR_NOT_SAME_DEVICE` without destination publication or source cleanup. The environment-gated
absence of a second volume is a documented skip. `FileOps.SelfTest.FairstreamMoveWindowPreserved`
records the deterministic rename-merge membership race with
`detail=new-file-before-exact-cleanup`; it requires the `Moved; source folder kept` outcome without a
cleanup prompt, a recopy, or a second delete. Both rows are correctness/strategy evidence and make no throughput-improvement claim.

Link preservation is literal and publishes in provider order; the bridge keeps no link queue,
mapping table, or dependency graph, so no `FileOps.Bridge.Link*` rows exist.
`FileOps.SelfTest.FairstreamMoveLiteralLinks`
covers backward publication, a forward target renamed by nested Keep Both, and a two-link cycle, all
keeping their stored target text. `FairstreamMoveLiteralLinkSkippedTarget` covers a skipped target
and requires a literal published link plus exact source retention. `FairstreamCopyWideLiteralLinks`
copies 4,200 junctions with no link queue. These rows prove literal preservation only; they make no
throughput-improvement claim.

Phase 3 conflict-authority runs also record
`FileOps.Conflict.ExpectedDestinationBoundCount`,
`FileOps.Conflict.ExpectedDestinationUnavailableCount`, and
`FileOps.Conflict.ExpectedDestinationReturnedCount`. The Local Replace-link and Native Move
Overwrite scenarios require one returned exact receipt and zero unavailable receipts. These are
bounded safety counters, not throughput claims; endpoint, path, identity, and artifact text are not
emitted.

Deferred File Operations consent records `FileOps.Consent.WaitUs` and
`FileOps.Consent.PromptCount` separately from `FileOps.Conflict.WaitUs` and
`FileOps.Conflict.PromptCount`. The per-wait row measures only the user-decision wait; the terminal
task rows aggregate consent duration and prompt count. They contain no endpoint, path, identity,
receipt, or artifact text. The deterministic Phase 10 matrix covers all seven risk layouts, known
item/byte facts, safe Cancel ordering, off-layout action rejection, task-local receipt creation, and
Recycle's no-Apply-to-all rule. Its live Local cases cover successful exact-object permanent-delete
escalation and replacement before the fresh bind. These are correctness and prompt-wait baselines,
not throughput-improvement claims.

Exact metadata transfer emits `FileOps.Metadata.InspectUs` once per inspected regular object and
`FileOps.Metadata.ApplyUs` for each prepare/finalize phase. Details are bounded phase/feature masks;
values report feature masks or logical bytes and never paths, stream names, identities, or security
descriptors. `FileOps.SelfTest.MetadataMatrix` proves actual Local MOTW/ADS/basic/sparse preservation,
ordered placeholder/EFS/sparse decisions, exact per-class loss diagnostics, and published-copy/source-
retained behavior. These rows are correctness/latency baselines, not throughput claims.

Local absent-name regular-file Copy emits `FileOps.Local.DirectFinalCopyUs` once per transfer with
bounded detail `shape=new-name-local-regular-file`; duration is the retained-reader/writer route,
`value0` is logical content bytes, and `value1` is the per-transfer pump buffer size. The buffer must
remain at or below 4 MiB; the implementation uses 1 MiB. The deterministic
`Floodgate_LocalCopyNewNameConcurrentReplacementSurvives` case emits
`FileOps.SelfTest.R0aLocalDirectFinalCopyUs` for one 64-MiB Local Copy and also proves metadata/ADS,
exact receipt truth, operation-control cancellation after visible progress, foreign-replacement
survival, exact partial removal, and absence of a sibling stage. Same-machine
test-enabled x64 Release validation uses five independent processes before and after the route
change. Candidate p95 must be at most `max(1.25 * baseline p95, baseline p95 + 25,000 us)` and every
sample must complete within 2,000,000 us. This is a safety-stabilization budget, not a throughput-
improvement claim.

`FileOps.SelfTest.CinderstarLegacyWriterReverify` and
`FileOps.SelfTest.CinderstarLegacyWriterCancelBackoff` cover the legacy atomic-final destination
verifier only through a Copy-only Move. Their bounded details identify transient/permanent/wrong-size
verification and cancellation, while the result must retain the source and emit no Managed cleanup
or commit-size-proof fallback. These are compatibility latency/cancellation baselines, not authority
for destructive Move.

Exact-object qualification uses `fileops.identity.bind_us` and
`fileops.identity.revalidate_us`. `value0` is provider-path payload bytes, `value1` is the bounded
typed binding/revalidation state, and `hr` is the provider/contract result. These metrics contain no
path or identity bytes. Scenarios compare local regular-file, hard-link alias, reparse object, and
missing/changed outcomes so identity safety does not silently add unbounded admission latency.

Role-aware mutation interlocks use `FileOps.Interlock.PrepareUs` once per task and exactly three
`FileOps.Interlock.ScopeCount` rows per task. The bounded details are `read-source`,
`write-source`, and `publish-destination`; `value0` is the matching role count and `value1` is the
total scope count. No path, identity, selected-item, ancestor, or discovered-descendant text is
emitted. Evidence must show that role classification does not increase retained authority beyond
unique governed provider paths plus the 1,024-node path-depth bound, that read/read overlap is not
serialized, and that an overlapping writer or publisher cannot proceed while a source is being
read. `FileOps.Interlock.WaitUs` remains the wait/cancel latency evidence.

Alias-safe final-leaf admission additionally emits `FileOps.Interlock.BindCount` once per prepared
task plus `FileOps.Interlock.CheckUs`, `FileOps.Interlock.ScopeComparisons`, and
`FileOps.Interlock.ActiveCandidates` for each admission
scan. The check row carries comparison count and active-candidate count as bounded numeric values;
no provider path, suffix, anchor identity, or selected-item text is emitted. Test-enabled x64
Release evidence covers distinct two-file Queue/Parallel admission, same-leaf early exit, alias
serialization, and a 256-by-256 disjoint final-publication-leaf scan. An unexplained same-machine p95
regression above 10 percent in PrepareUs or CheckUs, a timeout, UI starvation, or unbounded provider
probing is a stop-and-review trigger. The initial exact Cartesian scan reached that trigger at 65,536
comparisons and 2,600,470 microseconds in
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-29_022850/`. The accepted retained-anchor/suffix index
baseline is 8,704 probes and 2,689 microseconds in
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-29_024402/`. The indexed path must preserve the exact
Cartesian predicate as a fallback for conservative, legacy, or unkeyable scopes; an optimization may
not convert incomplete identity into path authority.

Explicit same-host concurrent admission and live-output protection use
`FileOps.Overlap.RunConcurrent`, `FileOps.LiveOutput.IndexHighWater`,
`FileOps.LiveOutput.IndexOverflow`, `FileOps.LiveOutput.GateCount`, and
`FileOps.LiveOutput.WaitUs`. Values are numeric only and contain no provider path or identity text.
The deterministic `R4A02_RunConcurrentSameDestinationStaysSafe` case proves that one disclosed
identity-bound relation admits both tasks without an interlock wait while a third undisclosed task
still asks independently. `R4A02_LiveOutputGuardChoices` pauses a Local publisher immediately after
one exact item is published and pauses a permanent Delete immediately before the live-output guard;
it covers Skip, exact-publisher Queue, explicit invalidation, terminal release, and a forced 256-entry
overflow fallback that withholds the destructive choice. The protected case must complete without
recursive enumeration or provider I/O added solely for correlation. The index high-water must not
exceed 256, overflow must retain protection, and Queue wait duration is observation rather than a
throughput claim because it intentionally includes the named publisher's remaining lifetime.

`FileOps.SelfTest.InterlockLargeSelectionRetainedAuthorityNodes` records preparation duration in
`duration_us`, unique retained authority nodes in `value0`, and the fixed 256-item selection in
`value1`. `FileOps.SelfTest.InterlockLargeSelectionMaxAuthorityDepth` records observed maximum
chain depth in `value0` and the 1,024-node bound in `value1`. The scenario is a retention baseline,
not a throughput claim; it must prove common ancestors are interned and that no discovered
descendant contributes a retained node.

Just-in-time transfer containment uses `fileops.identity.guard_us` for initial no-follow
classification and `fileops.identity.guard_revalidate_us` at the provider mutation boundary.
`value0` on classification is the combined source/destination path payload size (never content),
and the bounded state/HRESULT distinguish Ready, same-object, same-folder Move, subtree,
ancestor-link, changed, unsupported, indeterminate, and provider-contract outcomes. Validation must
compare direct file, same-path Keep Both, hard-link alias, directory ancestry, and swap cases and
must confirm the gate does not become a whole-selection preflight.

Related documents:

- `Specs/Testing/Testing_SelfTests.md`
- `Specs/TestRuns/README.md`
- `Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md`

## Validation-planner scenario contract

Operation Startrail is itself performance-sensitive tooling. Its implementation must
measure these named scenarios:

- `validation.plan.warm_explain_ms`: warm affected-set planning plus explanation;
- `validation.plan.cold_snapshot_ms`: cold workspace acquisition, hashing, and planning;
- `validation.plan.fileops_resume_avoided_ms`: elapsed work avoided by a valid controlled
  File Operations resume compared with an equivalent Fresh run.

Measurements must expose enough component timings to attribute acquisition, hashing,
plan construction, evidence lookup, verification, serialization, and runner-launch cost.
The archive must retain the sample count and distribution or percentiles needed to judge
variance. Timing comparisons are authoritative only on the same machine/profile and are
archived under `Specs/TestRuns/<MachineHash>/Validation/<RunId>/`; cross-machine results
are directional only. Estimated time savings may inform planning output but never trust,
reuse, or Full-gate decisions. The implementation reports Git discovery, content hashing,
canonicalization, impact matching, schema validation, evidence lookup, receipt binding,
and entry fingerprinting separately. If independent Commands payload proof is unavailable,
`validation.plan.fileops_resume_avoided_ms` is archived as explicitly blocked with zero
Full avoidance rather than manufacturing a savings claim. Fresh remains final-closeout
authority even though local exact Resume and Affected iteration are active.

## Required Development Contract

### 1. Every perf-sensitive change MUST name its scenario

Before implementation is considered complete, the owning change MUST identify:

- the user-visible scenario being protected,
- the primary metric or metric family,
- the authoritative selftest or deterministic repro path,
- the expected direction of change.

Examples:

- “Find Files progress storm while results are visible”
- “Compare Directories content-progress queue churn”
- “Cold folder open with repeated icon indices”
- “FileOps Skip-discovery release latency”
- “Parsed ViewerText diff open, side-by-side scroll repaint, unchanged-text expansion, range-bounded referenced-file hydration, viewport rehydration, and unresolved placeholder fallback”

### 2. Every perf-sensitive change MUST have measurable evidence

A change MUST satisfy one of these:

- extend an existing metric family, or
- add new instrumentation, or
- explicitly justify why existing metrics already cover the scenario.

For optimizations, “seems faster” is not sufficient. The change MUST be supported by archived measurements or by a documented blocked reason.

### 3. New features MUST integrate tests and perf measurement from the start

When a new feature introduces a new hot path, queue, async pipeline, render surface, large-list path, or repeated callback flow:

- a deterministic selftest or deterministic repro harness MUST be added with the feature,
- at least one performance metric relevant to that path MUST be emitted with the feature,
- the expected archive path under `Specs/TestRuns/` MUST be part of the validation plan for the feature.

This requirement applies even when the first landing only establishes a baseline.

### 4. Optimizations MUST preserve correctness coverage

An optimization is incomplete if it improves a metric but lacks behavioral protection.

Any performance change MUST keep or add:

- correctness selftests,
- perf instrumentation,
- archived run evidence when the scenario is runnable on the current machine.

### 5. Claims of improvement MUST be grounded

A claimed performance improvement MUST state:

- the baseline run,
- the candidate run,
- whether the comparison is same-machine and same-suite,
- the metrics that improved,
- any material caveats.

If same-machine evidence is unavailable, the claim MUST be labeled directional rather than definitive.

## Perf Gate Template

Every perf-sensitive PR MUST include this gate in the PR description, review notes, or linked closeout spec before merge:

- Scenario: the exact user-visible path being protected.
- Subsystem and change type: feature, optimization, stabilization, regression-fix, or instrumentation-only.
- User-visible risk protected: the latency, throughput, responsiveness, queueing, memory, or correctness/perf interaction being guarded.
- Metric keys: the emitted metric family or the existing metrics that cover the path, including units and sample grain.
- Instrumentation: existing instrumentation reused or new instrumentation added.
- Deterministic validation: the selftest, focused harness, or deterministic repro that exercises the path, including exact command, case filter, timeout multiplier, perf budget path when applicable, and required environment variables.
- Build flavor: Debug diagnostic, test-enabled Release, or another explicitly justified flavor. Final throughput, latency, percentile, or budget claims require Release evidence unless the owning spec allows otherwise.
- Environment/root matrix: machine hash, CPU/load notes when material, exact `REDSALAMANDER_TEST_ROOT=X:\RedSalamander.Perf`, filesystem capability reason, per-machine resource capability digest when applicable, and whether the run is final evidence or diagnostic only. `X:` may be any selected fixed local drive; repository-local `.build\TestSandbox`, historical `RSPerf`, `%LOCALAPPDATA%`, and process-temp roots are not valid evidence locations.
- Archived evidence: the `Specs/TestRuns/<MachineHash>/<Area>/<RunId>/` folder containing `results.json`, `trace.txt`, `run-all-tests-results.json` when the runner created one, and `perf_metrics.jsonl` when metrics are emitted.
- Analyzer and sample quality: the `Tools/Show-PerfRuns.ps1` command or owning analyzer, `-FailOnQuality` status for percentile claims, and explicit sample count sufficiency.
- Before/after delta: baseline run, candidate run, same-machine status, same-suite status, changed metric values, and caveats.
- Authoritative spec updates: the owning durable spec or guidance files that were updated, or the exact blocker if the update is deferred.

If any field is blocked, the PR MUST state the blocker explicitly and identify the follow-up owner/task. A perf-sensitive PR is not complete with only manual timing notes or unarchived terminal output.

## Instrumentation Rules

### Metric design

Metrics SHOULD be:

- scenario-specific,
- deterministic enough for repeated selftest use,
- attributable to one stage of work,
- named consistently with the owning subsystem.

Examples:

- `find.ui.*`
- `compare.ui.*`
- `FileOps.*`
- `render.*`
- `icons.*`
- `viewer.diff.*`

The Monitor retained-state family separates streaming/open publication and asynchronous export stages:
`monitor.file_open.total_us`, `monitor.file_open.peak_retained_text_bytes`, `monitor.file_open.publish_us`,
`monitor.file_open.cancel_us`, `monitor.file_open.close_us`, `monitor.file_save.snapshot_us`,
`monitor.file_save.snapshot_lock_us`, `monitor.file_save.ui_return_us`, `monitor.file_save.total_us`,
`monitor.file_save.throughput_bytes_per_sec`, `monitor.file_save.cancel_us`, and conditional
`monitor.file_save.close_us`. Immutable capture also reports `monitor.file_save.retained_text_bytes`,
`retained_line_count`, `shared_block_count`, `shared_block_bytes`, `copied_text_bytes`,
`active_tail_copied_bytes`, and `peak_additional_snapshot_bytes`. Its Release gate runs 512, 4,096, and 16,384
records five times each and emits `monitor.file_save.baseline_deep_copy_us` for the faithful removed algorithm.
Candidate text-copy and active-tail bytes must remain zero; additional snapshot bytes must be bounded by the
record-handle descriptors rather than retained UTF-16 bytes.
Search-at-cap and ingestion pressure remain observable through `monitor.search.match_update_us`,
`monitor.search.match_rebuild_us`, `monitor.etw.queue_high_water_mark`, and `monitor.etw.dropped_count`.

The embedded Terminal Kitty pipeline uses `terminal.kitty.capture_lock_us` for the bounded Ghostty-owned copy,
`terminal.kitty.submit_us` for UI-to-worker admission, `terminal.kitty.queue_bytes` for latest-only queue pressure,
`terminal.kitty.convert_us` for worker CPU conversion, `terminal.kitty.active_close_us` for the joined active
maximum-size shutdown path, `terminal.kitty.upload_us` for UI-thread D2D publication, the aggregate upload
retry/stable-failure/device-recovery/repaint counters for bounded failure behavior, `terminal.kitty.resize_us`
plus retained-bitmap count/bytes for ordinary resize reuse, and `terminal.kitty.memory_high_water_bytes` for
resident pipeline storage. End-to-end rendering and responsiveness
remain observable through `terminal.frame_total_us`, `terminal.input_to_frame_us`, and
`terminal.output_to_frame_us`. The deterministic Release path is
`PluginContractTests.exe --terminal-selftests`; it must cover a 4096x4096 RGBA generation, raw/PNG input,
replacement, saturation, stale-result rejection, queued and active-conversion shutdown, injected upload
dispositions, ordinary resize retention, and unload. Product capture/upload/resize metrics require a real
Kitty-producing lifecycle and a real D2D render target; they must not be inferred from the CPU-only fixture.

### Terminal VT upgrade admission

An admitted Ghostty lib-vt change requires a same-machine, same-commit,
test-enabled x64 Release Product/candidate pair produced by
`PluginContractTests.exe --terminal-vt-upgrade-corpus <archive-directory>`.
Each run retains exactly 200 ordered raw samples for `vt_write`, `snapshot`,
and `hosted_render`. Percentiles use nearest rank; the validator recomputes p50
and p95 from the raw rows and rejects a recorded summary that disagrees.

The pair must have identical branch, configuration, environment, corpus, and
machine identity, but distinct runtime identity and build receipt. All control
PTY-output and semantic-snapshot digests must match exactly; candidate main
digests may differ from its control only for the corpus-declared expected delta.
For every metric, candidate p95 must be no greater than
`ceil(Product p95 * 1.10)`. A run's self-recorded
`candidateMaximumP95` is checked for internal consistency but is never the
comparison authority.

Before an admission claim, run:

```powershell
.\Tools\Test-TestRunArchive.ps1 `
    -BaselineRunPath <product-run> `
    -CandidateRunPath <candidate-run>
```

The command validates `TerminalVtUpgradeEvidence.schema.json`, parses every
JSONL row, verifies the result/archive/sample bindings and exact 200-sample
order for all three metrics, recomputes the percentiles, and enforces the paired
digest and 1.10 gates.

The selected-path manifest scenario is `fileaction/selected-paths-10000`, exercised by the Commands case
`file_action_selected_paths_streaming_perf`. It measures the synchronous launch-plan build/write path over 10,000
Unicode paths with `fileaction.selected_paths_file.scenario_us`, `build_us`, `validate_us`, `create_us`, `write_us`,
`path_bytes`, `write_calls`, `peak_additional_bytes`, and `selection_copy_bytes`. The companion
`file_action_selected_paths_manifest_metrics.json` artifact also records record count, result HRESULT, and whether an
aggregate payload was used. A successful 10,000-record run requires the exact `ceil(totalBytes / 64 KiB)` bounded-buffer
write count, zero selection-copy bytes, no aggregate payload, and at most 64 KiB of serialization storage. The case also
emits a test-only `reference_*` family for the reviewed vector-copy, aggregate-payload, two-write algorithm; that
reference uses a safe delete-on-close sandbox file and never recreates the reviewed pathname race. Final
performance evidence MUST use a test-enabled x64 Release build and five independent processes with separate selftest
roots, archived under `Specs/TestRuns/<MachineHash>/FileActions/<RunId>/`. Sandbox-root length changes `path_bytes`, so
before/after review compares the write/copy/peak shape and stage timings rather than requiring identical byte totals.
Unless the scenario is extended through a real 10,000-selection command dispatch, `scenario_us` is synchronous
launch-plan build/write evidence and MUST NOT be described as end-to-end visible UI command latency.

The selected-path crash-recovery scenario is the Commands case `file_action_selected_paths_recovery_bounds`. It
creates old reclaimable manifests plus fresh, locked/live, directory, reparse, shared-temp-shaped, and out-of-root
controls, then proves zero-time cutoff, missing-cursor reset, per-pass entry/delete bounds, and deterministic
multi-pass convergence. The test emits `fileaction.selected_paths_recovery.selftest_us`, `selftest_passes`,
`selftest_skipped`, `selftest_errors`, and `selftest_bound_hits`, and writes
`file_action_selected_paths_recovery_metrics.json`. Product recovery emits
`fileaction.selected_paths_recovery.elapsed_us`, `inspected`, `deleted`, `skipped`, `errors`, `bound_hits`, and
`passes`. Closeout evidence MUST use a test-enabled x64 Release build and five independent processes with separate
selftest roots, archived under `Specs/TestRuns/<MachineHash>/FileActions/<RunId>/`. The pre-repair shared-temp scan
had no accepted recovery instrumentation, so this repair's first archive is an invariant/scaling baseline and MUST
state that limitation instead of inventing before numbers; acceptance requires all reclaimable entries deleted,
all protected controls preserved, zero unexpected errors, observed bounds, and bounded convergence.

Current ViewerText diff baselines include `viewer.diff.open_to_first_visible_us`, `viewer.diff.visible_rows`, `viewer.diff.semantic_row_paint_us`, `viewer.diff.visible_styled_rows`, `viewer.diff.visible_context_rows`, `viewer.diff.visible_banner_rows`, `viewer.diff.visible_split_rows`, `viewer.diff.theme_switch_repaint_us`, `viewer.diff.scroll_repaint_us`, `viewer.diff.hunk_jump_to_visible_us`, `viewer.diff.expand_context_us`, `viewer.diff.viewport_rehydrate_us`, `viewer.diff.viewport_backtrack_us`, `viewer.diff.deferred_rows`, `viewer.diff.referenced_bytes_read`, `viewer.diff.viewport_referenced_bytes_read`, `viewer.diff.viewport_referenced_bytes_delta`, `viewer.diff.viewport_backtrack_referenced_bytes_delta`, `viewer.diff.placeholder_rows`, `viewer.diff.placeholder_bands`, and `viewer.chrome.paint_us`.
The `viewer_text_diff_perf` artifact also records semantic-row paint timing, context-row visibility, pane-local side-by-side layout activation, visible split-row counts, pane column widths, rainbow-mode theme-switch repaint timing, root-shell chrome repaint timing, hunk-jump latency, and expand-time, post-viewport, and post-backtrack referenced-byte counts so clickable hidden-banner reveal, calmer base-background side-by-side structure, combo-sync behavior on scroll repaint, root-shell typography changes, horizontal-scroll presentation switches, bounded referenced-file growth, and cache reuse can be reviewed alongside the metric stream.
Preserve-viewport diff rehydrate/backtrack rebuilds should reuse the existing section-combo model and avoid full combo repopulation or header relayout unless the visible section set actually changes.

The ViewerText encoding-picker scenario is `viewer_text_encoding_picker_open_and_filter_full_catalog`. Its exact test-enabled x64 Release invocation is `.build\x64\Release\ViewerPETests.exe viewer_text_encoding_picker_open_and_filter_full_catalog`. It emits `viewer.encoding_picker.catalog_load_us`, `viewer.encoding_picker.open_to_ready_us`, `viewer.encoding_picker.filter_us`, `viewer.encoding_picker.catalog_items`, `viewer.encoding_picker.filtered_items`, `viewer.encoding_cycle.catalog_cache_build_us`, `viewer.encoding_cycle.catalog_cache_items`, and `viewer.encoding_cycle.command_us`. Release evidence must use at least five independent processes and the full localized catalog (at least 100 selectable encodings), prove exactly one catalog load and one cycle-cache build per process despite repeated picker/F8 input, keep candidate open-to-ready p95 below 50 ms, filter p95 below one 16.67 ms frame, and encoding-cycle command p95 below one frame, and archive the raw JSONL plus summary under `Specs/TestRuns/<MachineHash>/Viewers/<RunId>/`. Because the legacy native cascade did not emit an accepted open-latency metric, the first searchable-picker archive is an invariant/scaling baseline; it must not invent a numerical before measurement. The first accepted archive is `Specs/TestRuns/4cb089111a23/Viewers/20260828_192400_menu_encoding_picker_release/`: five processes exercised 140 catalog items with open-to-ready p95 25,150 us and filter p95 12,112 us. The catalog-cache candidate is `Specs/TestRuns/4cb089111a23/Viewers/20260829_115300_menu_encoding_catalog_cache_release/`: five expanded processes retained 140 items with one catalog load/cache build per process, open-to-ready p95 19,474 us, filter p95 4,383 us, and repeated-F8 command p95 4,624 us.

### Preferred measurements

When applicable, instrument:

- queue depth or queue drain behavior,
- message coalescing and skipped work,
- input-to-visible or progress-to-visible latency,
- rebuild/repaint counts,
- bytes processed or rewritten,
- lock hold time,
- pre-calc / sort / refresh stage timings,
- bounded visible-work counters for large lists or grids,
- DirectWrite/glyph text-layout creation counts and timing (the `dwrite.text_layout.*` family in FolderView; see `Specs/UI/UI_FolderView.md`). 2026-06-19 same-machine evidence found FolderView text-layout creation non-material (about 1.2–1.4% of `render.layout_items_us`), so this family is reusable measurement infrastructure and not, by itself, a justification to add a text-layout cache without fresh evidence.
- Per-phase decomposition of a hot pass into named sub-stage timings (e.g. the `folder.layout.*_us` family decomposes `render.layout_items_us` in FolderView; see `Specs/UI/UI_FolderView.md`). 2026-06-19 this decomposition exposed that an apparent FolderView layout "bottleneck" (~82.6% in `UpdateItemTextLayouts`) was ~96% a *measurement artifact* — the JSONL sink opening/closing the file per metric row — not real work; the sink was fixed to keep the handle open (`Common/Common/PerfJsonl.cpp`). A worked example of why you decompose before optimizing: the dominant cost was outside the suspected mechanism.

### Anti-patterns

Avoid relying only on:

- ad hoc logging with no archived metric,
- one-off manual stopwatch measurements,
- broad “total duration” metrics when the decision requires stage attribution.
- repeated hot-path rows that only report pointer coordinates and sub-millisecond callback duration, such as per-hit-test/per-pointer-move traces, when the user-visible scenario needs input-to-visible latency, queue counts, scroll-apply cost, or paint cost instead.
- high-frequency per-item or per-creation emits that distort the very pass that hosts them. The JSONL sink keeps its file handle open (fixed 2026-06-19, `Common/Common/PerfJsonl.cpp`, so a row is a single `WriteFile` rather than an open/close), but each row still has a cost and grows the file, so coalesce hot-path telemetry into per-pass/per-frame aggregates (cf. the FolderView `dwrite.text_layout.frame_create_*` pattern).
- file-backed selftest trace writes inside the measured render, draw-item, icon-apply, or present hot path. Use `Debug::Perf` counters/scopes for measured paths, and keep `SelfTest::AppendSelfTestTrace(...)` outside the timed gesture or behind a diagnostic mode that is disabled for perf-budget runs. A 2026-06-29 FolderView closeout pass found that trace writes around `Present1` and per-batch icon apply distorted the huge quick-search and relayout matrices even though the product work was below budget.
- self-test-local perf sinks configured through `Debug::Perf::ConfigureJsonlOutput(...)` must immediately update the process-local sink cache. A previous "no sink configured" fast path must not suppress metrics after a later self-test configures `perf/perf_metrics.jsonl`; `ClearJsonlOutput()` must likewise clear the cache so following tests do not append to a stale sink.

## Test and Archive Requirements

### Selftests

Perf-sensitive features SHOULD prefer:

- `--commands-selftest` for window/UI behavior,
- `--compare-selftest` for Compare/search engine behavior,
- `--fileops-selftest` for File Operations.
- test-enabled `RedSalamanderMonitor.exe --chrome-selftest --wait-instance --perf` for Monitor chrome, bounded
  ingestion/search, streaming open/move-publication, exact immutable-snapshot asynchronous export, transactional
  thread admission, canonical output bytes, three-size/five-repeat snapshot scaling, and retained-state metric
  presence. The selftest-owned I/O fixtures must be removed before the case completes.

If none of the existing suites can exercise the scenario, a new deterministic selftest case or equivalent harness MUST be added.

### Archival

Meaningful validation runs MUST be archived under `Specs/TestRuns/` per `Specs/TestRuns/README.md`.

Before curating or force-adding a generated run, validate the complete run directory with
`Tools\Test-TestRunArchive.ps1 -RunPath <run-folder>`; the ordinary no-argument invocation is only the fast
changed-set gate. Repository closeout also requires `Tools\Test-TestRunArchive.ps1 -Inventory`, which covers every
checked-in archive added or modified since contract commit `8b9e36191835cc71b5ee0dfeea4dddd5d942dd5a`. Both
modes enforce the 2 MiB per-file and 5 MiB per-run limits and return nonzero with the exact violation. A failed archive
gate is failed evidence, regardless of the scenario result. Every nonblank JSONL row is parsed. Governed area/schema
formats additionally bind summary, trace, raw rows, archive identity, and machine profile; a Terminal VT upgrade claim
also requires the paired validator described above.

High-frequency raw metrics that exceed the limits may be compacted only when their full ordering is not load-bearing,
or when a bounded deterministic summary preserves the original digest, byte/row counts, behavioral authority, and
claim-bearing rows in original order. Do not raise archive limits or discard raw ordering that a claim needs.

If the repo archive path is unavailable, the blocked reason MUST be documented explicitly.

### Baselines

Once a scenario becomes important for ongoing work, a same-machine archived baseline SHOULD be recorded and reused for future comparisons.

### Evidence quality gates

Frame-time and latency claims MUST report the actual selftest build flavor carried
by the JSONL `build` field. Release runs are not acceptable evidence when the
artifact is mislabeled as Debug or when the build flavor is unknown.

For percentile-based decisions, `Tools/Show-PerfRuns.ps1` is the required
first-pass analyzer. Use PowerShell 7+ for this script. Any p95 claim SHOULD have
at least 200 samples for the requested metric unless an authoritative per-metric
budget declares a smaller `minimumSamples`; p99 SHOULD have at least 1000
samples. Automation that gates a p95 claim MUST pass `-FailOnQuality`, which
exits non-zero when a quality-gated requested metric has fewer than its p95
sample minimum. An explicitly supplied `-MinimumSamplesForP95` overrides budget
minimums; otherwise `-BudgetPath` selects the budget document used for those
minimums. If that selected budget file exists but cannot be parsed, the analyzer
MUST warn, latch quality failure, and exit non-zero when `-FailOnQuality` is set.
`-FolderViewPreset` uses the p95 rows in
`Specs/Testing/FolderViewPerfBudgets.json5` as its quality-gated metric set; the
remaining preset rows are still displayed but are informational for
`-FailOnQuality` unless requested explicitly with `-Metric`. In `-CompareRun`
mode, `-FailOnQuality` gates the candidate run only; a sparse baseline remains
visible in the comparison table but does not fail the command. Use
`-FolderViewPreset` for the core FolderView frame, scale/cold/slow, scroll,
icon, thumbnail, and refresh metric family when reviewing FolderView runs.
FolderView scroll/dirty-region reviews MUST check that the preset summarizes
`folder.scroll.product_paint_render_count`,
`folder.scroll.product_paint_full_client_count`, and
`folder.scroll.product_paint_dirty_rect_area_px` in addition to
`folder.scroll_input_to_paint_us` and the `folder.frame.*` rows. No-op scroll
requests that do not change the viewport must report zero product paint frames;
viewport-changing scrolls are expected to remain full-client paints unless a
separately gated DXGI scroll-rect implementation exists.
FolderView icon-pipeline reviews MUST check that the preset summarizes
`FolderView.ExecuteEnumeration.IconIndex.QueryExtensions`,
`FolderView.ExecuteEnumeration.IconIndex.QueryPerFileIcons`,
`FolderView.ExecuteEnumeration.IconIndex.BuildPerFilePaths`,
`FolderView.IconLoading.ProcessQueue`, `FolderView.IconLoading.BatchUpdate`,
`FolderView.IconLoading.BitmapConversion`, `icons.queue_wait_to_dequeue_us`,
`icons.extract_us`, `icons.batch_update_scan_us`, `iconcache.shgetfileinfo_us`,
`iconcache.lock_wait_slow_us`, `iconcache.lock_hold_slow_us`, and
`icons.recall_avoided_count`.

FolderView same-folder refresh reviews MUST check the aggregate `folder.refresh.*`
family. `preserve_count`, `rebuild_count`, `selection_preserve_count`,
`rename_transfer_count`, `debounce_delay_ms`, and `enumeration_count` are
one-row-per-refresh counters; `request_to_paint_us` is a request-to-present
latency row emitted only after the refresh result has reached the UI thread and
the pane has presented. Refresh latency MUST use the dedicated FolderView refresh
pending slot so input-to-paint and navigation-to-paint rows cannot overwrite or
misattribute it.

The FolderView current/selection state scenario is `folderView_perf_focus_selection_state`. It uses the dummy provider with at least 10,000 displayed items and at least 200 accepted model-change samples for every p95 claim. Its phases cover current survives/disappears, multiple-neighbor disappearance, keyed sort and Sort None repair, sparse/dense selection, filter hide/show, selection preservation, one real pointer current/selection transition through the existing aggregate input-to-paint metric, and bounded focus-memory churn.

The scenario reuses the one-row-per-refresh `folder.refresh.*` family and `folder.refresh.request_to_paint_us`, and adds `folder.focus.resolve_us` exactly once per accepted model rebuild with `value0 = displayed item count` and `value1 = current-resolution reason`. Focus-memory observations use scenario-checkpoint aggregates only: `folder.focus_memory.entry_count`, `folder.focus_memory.payload_bytes`, and `folder.focus_memory.eviction_count`. The artifact also records resolution-reason counts, maximum entries/payload, evictions, hits/misses/no-cache reasons, and invariant/cap violations. It MUST NOT emit per-item or per-focus-move JSONL rows.

`folderView_perf_focus_selection_state` remains visible but unbudgeted until repeated stable same-machine, test-enabled Release evidence supports a hard threshold. A single baseline/candidate pair, or sparse frame/request-to-paint observations from this correctness-oriented scenario, MUST NOT be used to invent a machine budget. The deterministic invariant and dual-cap assertions remain blocking regardless of whether a timing budget exists.

Final evidence uses test-enabled x64 Release, same-machine baseline/candidate runs with identical scenario parameters, archive validation through `Test-TestRunArchive.ps1`, and comparison through `Show-PerfRuns.ps1 -FolderViewPreset -FailOnQuality -ShowBuildFlavor`. A hard row may be added to `FolderViewPerfBudgets.json5` only after repeated stable same-machine samples justify it; until then the scenario remains visible and explicitly unbudgeted rather than receiving an invented threshold.

FolderView frame-producing perf cases that make frame p95 claims MUST archive a
`metricQuality` object with `folderFrameTotal.count`,
`folderFramePresent.count`, `samplesEnoughForP95`, `samplesEnoughForP99`, and
`buildConfiguration`. `samplesEnoughForP95` is a hard case gate for
`folder.frame.total_us` and `folder.frame.present_us` when the case claims p95
evidence, and strict budget runs still enforce the budget file's
`minimumSamples`. `folderView_perf_overlay_invalidation_stress` records
`samplesEnoughForP95` as advisory because animation cadence and host load can
reduce frame samples without proving an overlay correctness failure; cite p95 for
that case only when the analyzer quality gate passes. `samplesEnoughForP99` is
advisory until a case deliberately captures at least 1000 frame samples. Other
event-scoped rows exposed by the FolderView preset, such as scroll input latency
or one-shot icon/thumbnail counters, may report `P95Quality=fail`; do not cite a
p95 for those rows unless that exact metric also has enough samples.

Representative FolderView perf coverage MUST include:

- `folderView_perf_scroll_render_stress`: 1,600-item normal-mode scroll/render
  coverage across Brief, Detailed, and Extra Detailed modes. It must produce at
  least 200 `folder.frame.total_us` and `folder.frame.present_us` rows for p95
  claims, emit scroll input latency and product-paint delta metrics, assert
  repeated no-op boundary scrolls do not repaint, and archive whether each
  scroll step changed the viewport.
- `folderView_perf_huge_folder_scale`: synthetic dummy-provider scale coverage at
  10,000 items for routine runs and 50,000 items only when
  `REDSALAMANDER_FOLDERVIEW_HUGE_PERF=1` is set. Artifacts record item count,
  extension count, enumeration time, first visible paint, sort toggle,
  quick-search keystroke-to-paint, select-all scroll brush guard, and working-set
  / private-byte samples including bytes per item.
- `folderView_perf_cold_first_visit`: local first-visit coverage with unique paths
  and extensions after clearing the application `IconCache`. The artifact records
  first enumeration, first paint, icon-index lookup metric count, icon bitmap
  queue count, icon-settle time, and a deterministic forced first thumbnail
  fallback. The artifact MUST state that the OS shell cache is not controlled by
  the test.
- `folderView_perf_slow_virtual_provider`: deterministic slow-provider coverage
  for dummy-provider enumeration latency, icon extraction latency, live-path
  failed icon lookup negative-cache behavior (`icons.repeated_failed_lookup_count`
  must stay bounded), provider-allowed thumbnail lookup latency, and the paste
  shortcut latency coverage supplied by the ShellCommands paste-shortcut cases.
- `folderView_perf_icon_pipeline_cold_slow`: focused icon-pipeline coverage for
  per-file icon types (`.exe`, `.dll`, `.ico`, `.lnk`, `.url`, `.cpl`, `.scr`,
  `.msc`, `.ocx`), one delayed HICON extraction, analyzer-visible queue/extract/
  convert/apply metrics, and synthetic offline/recall placeholder avoidance.
  The artifact must prove at least one visible bitmap icon resolves before the
  delayed extraction finishes, `IconPathLiveLookup` is not consumed for the
  offline/recall dummy fixture, `icons.recall_avoided_count` emits for those
  items, and the thumbnail pass avoids provider-allowed shell I/O.
- `folderView_perf_relayout_churn_while_scrolled`: 10,000-item scrolled-pane
  relayout coverage that alternates DPI, size, light/dark/high-contrast theme
  triggers, emits `folder.relayout_to_paint_us`, archives repaint-burst sample
  quality and `fontRelayoutCovered=false`, and asserts focus/scroll survival.

Every representative FolderView scale/cold/slow/relayout artifact MUST include an
`environmentMatrix` object with build flavor, active DPI, display refresh rate,
display scale percent, local-console/RDP status, WARP availability, whether WARP
was actually executed, adapter name, driver-version availability, high-DPI run
status, and notes for matrix dimensions that require separate hardware or
process-level runs. FolderView WARP coverage is opt-in for selftest evidence:
set `REDSALAMANDER_FOLDERVIEW_FORCE_WARP=1` before launching the Commands
selftest so FolderView creates its D3D device with `D3D_DRIVER_TYPE_WARP`; the
archive must then report `environmentMatrix.warpRunExecuted=true`.

Release perf evidence that depends on selftest cases MUST use a test-enabled
Release build, because normal Release binaries may omit selftest entrypoints:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
```

### Automated FolderView budget gates

The authoritative FolderView perf budget file is
`Specs/Testing/FolderViewPerfBudgets.json5`. It is intentionally machine-keyed
through a top-level `machines[]` array: each entry has a `machineHash` and its own
`budgets[]` list. Hard thresholds apply only when the current selftest
`machineHash` matches one of those machine entries, and each hard threshold MUST
cite its source archive, measured value, maximum, statistic, build flavor, and
`minimumSamples`.

Unknown machines are never silent. The native harness emits a visible warning with
a scaffold entry shape to fill, e.g. `{ "machineHash": "<current>", "budgets": [] }`.
The default behavior keeps non-applicable budgets as warnings so ad-hoc Debug runs
can proceed; strict runs add `--selftest-require-perf-budgets` (or
`Run-AllTests.ps1 -RequirePerfBudgets`) to fail when the budget path is missing, no
current-machine entry exists, or no hard entry matches the current build flavor.

Run the focused strict gate through the normal runner with:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -CaseFilter folderView_perf_scroll_render_stress -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5 -RequirePerfBudgets -TimeoutMultiplier 8
```

For an already-built test-enabled Release binary, the native equivalent is:

```powershell
.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-perf-budget=Specs\Testing\FolderViewPerfBudgets.json5 --selftest-require-perf-budgets --selftest-timeout-multiplier=8
```

For grouped validation, pass the comma-separated budgeted case list. The current
budgeted set is `folderView_perf_scroll_render_stress`,
`folderView_perf_overlay_invalidation_stress`,
`folderView_perf_huge_folder_scale`, `folderView_perf_slow_virtual_provider`,
`folderView_perf_relayout_churn_while_scrolled`, and
`folderView_thumbnail_cached_only_no_close_stall`.

Expected-failure smoke checks SHOULD use a local scratch budget with an
impossible maximum and MUST NOT commit that scratch file.

Selftest-only latency injection points MUST be implemented through
`RedSalamander/SelfTest/Common/SelfTestLatencyHooks.h/.cpp` and compiled as
no-ops outside `ENABLE_TESTS`.
These hooks exist to make slow shell/provider/file-system behavior deterministic;
they must not become production delays or runtime configuration.

## Spec Ownership Requirement

The owning normative spec for a subsystem MUST describe:

- the important user-visible performance scenarios,
- the expected test entrypoints,
- the authoritative metric families when they exist.

Performance requirements MUST NOT live only in a WIP plan when the subsystem has already adopted them as standard practice.

When a WIP plan is finished:

- the plan MUST move to `Specs/Plans/Done/`,
- any durable performance contract, verification entrypoint, or workflow requirement discovered in that work MUST be merged into the authoritative subsystem spec or repo-level guidance,
- the finished plan remains historical evidence, not the authoritative contract.

## Completion Criteria

A feature or optimization is performance-complete when:

1. the scenario is named,
2. the metrics exist or were shown to already exist,
3. deterministic coverage exists,
4. archived evidence exists or the block is documented,
5. the owning authoritative spec notes the lasting result when the change establishes or updates a baseline, and any finished WIP plan has been moved to `Specs/Plans/Done/`.
