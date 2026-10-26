# Earlier C0 execution history — non-normative

Moved from the active plan during the Move-preflight repair closeout. This preserves
older observations and provenance, not active instructions. In particular, historical
"next"/"pause"/"continue" directions below are superseded by the current WIP plan.
Durable behavior belongs to the authoritative domain specifications.

Focused receipt progress: `c0-native-receipt-matrix-verified-20260905` passed 5/5
(three behavior cases plus setup/cleanup), Debug receipt
`667f6840edaeddcc00bec93e28c2cd6550f3728e35a119b60724da9144104c44`, zero warnings/errors.
The normal machine `4cb089111a23` FileOps archive
`2026-09-05_c0_native_receipts_verified_debug` retains raw evidence and the prior
fault-location correction. Copy/Delete response loss and independent Local
publication/cleanup truth are covered. Subsequent Debug regressions all passed:

| Run ID (all 2026-09-05) | Cases | Build receipt | Machine FileOps archive |
|---|---|---|---|
| `c0-conflict-regression-repaired-20260905` | 11/11 | `a7f0414334e2c51ee67417c2f4383ea81ed6394c3940d0975360d9a41b32cfad` | `2026-09-05_c0_conflict_regression_repaired_debug` |
| `c0-owned-publication-regression-20260905` | 7/7 | `167683804c5f713ab2845ffc11459409a91d27ef68482d9d88506ae353893145` | `2026-09-05_c0_owned_publication_regression_debug` |
| `c0-local-overwrite-regression-20260905` | 3/3 | `f8134788c8e746e61a4785969ee51f835f51c03dbeb39deb69295fe76b5786b1` | `2026-09-05_c0_local_overwrite_regression_debug` |

All have zero warnings/errors and no skips. These runs are focused evidence, not the
complete route matrix; the existing RetryCap test remains baseline, not C2 acceptance.
Large metric streams retain digest-bound exact selected rows in the curated archives.
Disk audits flag retained older runs outside each invocation's allow-list, not a clean
test-root audit; no old evidence was deleted.

Receipt review additionally found full-struct copying from accepted 16-byte V1 buffers
and stale proof after malformed completions. Shared prefix validation and bounded
snapshotting now replace the host/test copies; malformed completion invalidates prior
proof even if the provider ignores `E_INVALIDARG`. `C0_MutationReceiptPrefixBoundary`
adds an inaccessible-page witness, malformed/null/full/future records, and actual host
callback/terminal checks. That six-case matrix passed without skips in both profiles:

- Debug `c0-receipt-prefix-debug-20260905`, receipt
  `6a2e447ae39f17f3ec187024c842d93e2b9eb25f91bb46ef5324df83f13535a2`,
  archive `2026-09-05_c0_receipt_prefix_debug`.
- Explicitly test-enabled Release `c0-receipt-prefix-enabled-release-20260905`, receipt
  `b1cb2ab29e8ca3c34ccb3f3205d05444ebdd11eef1ee502519f04a073f0c9bc4`,
  archive `2026-09-05_c0_receipt_prefix_release`.

Both builds have zero warnings/errors. The prefix witness measures 249/90 us in
Debug/Release (five rejected headers, five callbacks); differing configurations do
not establish a performance improvement. Source/localization Pester passes 197/197;
spec inventory has zero blockers; archive inventory passes 2,083 tracked files and
each newly curated local archive passes explicit pre-stage validation.

The first Release invocation omitted `RSBuildEnableTests=true` and failed receipt
collection because the profile-root SQLite DLL was missing after compilation.
`2026-09-05_c0_release_staging_failure` retains that observation. No runtime tests or
valid receipt came from it. The accepted test-enabled rebuild does not prove the
non-test Release staging transition repaired; C6 must qualify that transition too.

Curl still recurses in serial Copy, parallel Copy discovery and Move size preflight.
Native Delete, Move source cleanup and directory sizing now use bounded outer
walkers qualified below. Use the fake-FTP fixture to prove/fix each remaining
route, including deep/wide resource behavior and partial results. The earlier
root/cancel repair alone was not iterative-traversal evidence.

`C0_CurlNativeDeleteGuards` reproduced and repaired root rejection after child
deletion and missing single-item cancellation context. Native single/bulk calls at
concurrency 1/4 now reject root requests without mutation, delete valid selected
subtrees while preserving siblings, and cancel stalled admission LISTs in about one
second. The fixture passes 29/29 checks in Debug and test-enabled Release. Its
original run details and build receipts live in the validated machine FileOps archives:

- [Pre-fix failures](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_native_delete_red_debug/README.md): serial root Delete removed both fixture files; single Delete did not cancel cooperatively.
- [Debug C0 matrix](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_native_delete_debug/README.md): 7/7, zero skips/warnings/errors.
- [Release C0 matrix](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_native_delete_release/README.md): 7/7, zero skips/warnings/errors; preceded by a safely declined overlapping build launch.
- [FTP host regression](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_host_regression_debug/README.md): 3/3. Source/localization checks 197/197; spec inventory zero blockers; whole archive inventory 2,083 tracked files plus explicit checks for the new curated archives. Retained-run disk-audit limitations remain recorded, not erased.

The next C0 iteration, also based on `58f22d0e` plus working changes, reproduced
unsafe quote replay: all four single/bulk Delete/Rename lost-reply cases sent two
mutations, destroyed or moved a concurrent replacement, and reported success.
`CurlPerformQuote` now executes once and suppresses the unrelated directory
transfer. `C0_CurlCommittedMutationResponseLost` passes all 33 checks: eight native
scenarios, exact file bytes, one request/callback, honest mutation classification,
and no incidental NLST. `C0_CurlHostCommittedDeleteResponseLost` additionally proves
an immutable uncertain result, localized `Outcome unknown` summary and zero conflict
prompts. The fake endpoint's client/start/stop paths retain RAII ownership through
thread-start failure and named exception boundaries. Validated archives:

- [Pre-fix lost-reply failures](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_reply_loss_red_debug/README.md): 21 native checks passed, 12 failed; four actual unsafe replays.
- [Debug C0 matrix](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_reply_loss_debug/README.md): 9/9; lost-reply paths 3.6–5.6 ms instead of 219–221 ms with replay; bounded loopback fault timing only.
- [Test-enabled Release C0 matrix](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_reply_loss_release/README.md): 9/9; native guards 29/29 and lost-reply checks 33/33; host witness 63 ms.
- [FTP workflow regression](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_reply_loss_host_regression_debug/README.md): 3/3, ordinary mutations remain usable. All four builds had zero warnings/errors. Source/localization checks 197/197; spec inventory zero blockers; new archives explicitly validated, tracked archive inventory 2,083 files. Retained-run disk-audit limitations remain explicit.

Partial-tree correction, starting from `58f22d0e` plus recorded C0 working changes:
the shared completion owner now requires explicit selected-item evidence, never
a final-HRESULT guess. Child mutations invalidate no-commit; untouched items and
native Move/Rename collision preflight retain proof. Results and diagnostics live
in the reviewed archives rather than another execution log in this plan:

- [Original red baseline](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_partial_tree_red_debug/README.md): 11/13 native checks; both serial paths falsely claimed no commit after two mutations.
- [Fixture compile correction](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_partial_tree_build_failure/README.md) and [misplaced control fault](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_partial_tree_control_red_debug/README.md): retained failures, not qualification.
- [Passing Debug](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_partial_tree_debug/README.md) and [test-enabled Release](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_partial_tree_release/README.md): 11/11 each, native partial-tree/proof checks 37/37; host uncertain result and no unsafe conflict actions. Native timings and limitations are retained.
- [Ordinary FTP regression](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_partial_tree_ftp_regression_debug/README.md): 3/3; zero-warning Debug build. All three passing runs bind source snapshot `37858cb4eb796e8d3dd2c8ded2ee0276c1abb32913f61e649aad77fdff57a6ac`. Final edits only consolidate documentation/evidence.

Next complete native iterative route/walker and host Rename journey evidence.
The next traversal step starts from checkpoint `cbf19004` and introduces
`C0_CurlNativeDeleteLateListing`: native single/bulk Delete at concurrency 1/4
must delete already discovered work before a later directory listing completes.
The fake server gates that listing on an exact earlier deletion, then completes,
refuses, or waits for cancellation. Run the focused Debug baseline before changing
the walker; retain failed ordering/partial-state receipts. This small-tree ordering
witness is not deep/wide bounded-memory or UI responsiveness qualification.
The [Debug red baseline](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_late_listing_red_debug/README.md)
reproduces all six parallel ordering failures (31 assertions pass, 18 fail), while
serial normal/late-failure/cancel controls pass. Parallel cases time out at the
fixture's five-second gate; serial cancellation is about one second. This is an
intentional failing regression, not a passing C0 gate. Next replace eager collection
without preserving unbounded per-directory vectors or adding an aggregate queue cap.
The bounded Delete step starts at `5132d866`: native Delete/Move-cleanup recursion
is replaced with an incremental cursor and explicit stack, reusing the current
Delete scheduler. `C0_CurlDeleteTraversalBounds` adds 127 assertions: 36 single/bulk
transport scenarios at concurrency 1/4 and 18 pure FTP/SFTP/SCP command-operand
checks. A discovered shared bug is also repaired: namespace commands previously
sent URL-escaped path spellings, failing on spaces and other literal bytes.
The existing literal base path and strict conversion now feed the appropriate
FTP control-line or libcurl SFTP operand grammar; control injection is rejected.

Qualification archives under machine `4cb089111a23` / FileOps:

- [Debug C0](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_streaming_delete_debug/README.md): 13/13, late listing 49/49, traversal/operands 127/127; all existing C0 cases pass.
- [Test-enabled Release C0](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_streaming_delete_release/README.md): the same source snapshot, 13/13, no skips.
- [Ordinary FTP workflow](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_streaming_delete_ftp_regression/README.md): 3/3; Copy/Create/Rename/Replace/Delete remain usable.

All three builds have zero warnings/errors. Each C0 profile has 65 checked
traversal/retained metric pairs: peaks are 129 frames, 128 queued leaves, 130,592
path bytes, and metadata reservations below 6.32 MiB (within 16/8-MiB policy).
The 4,097-file cases succeed; the 80-level cases succeed; depth/record boundaries
fail explicitly without mutation in the rejection fixtures. Late blocked-list
cancel acknowledgements are 61.8–92.6 ms Debug and 60.3–62.1 ms Release on loopback;
these are focused cancellation observations, not whole-application timing claims.
Source/localization checks pass 197/197. Full raw streams or digest-bound exact
selected rows are retained; disk-audit limitations are not erased.

The [compile failure](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_streaming_delete_build_failure/README.md),
[root-error regression](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_streaming_delete_partial_failure/README.md),
[peer-disconnect diagnosis](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_streaming_delete_disconnect_diagnostics/README.md),
and [literal-name failure](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_streaming_delete_literal_name_failure/README.md)
remain archived rather than being represented as qualification passes.
Copy/Move discovery and directory-size walkers remain separate open routes.
The cursor retains one transport owner per active ancestry frame; real FTP/SSH
session limits and opaque transport memory still need explicit route qualification.
The reservation counters are not total process-memory proof.
The current concurrency-4 discovery-first observation is not a requirement to
pre-scan whole trees. During streaming migration, use an explicit first-listing
refusal for zero-mutation controls and an after-mutation fault for partial results
at each concurrency; preserve proof truth without preserving eager collection.
These focused checks do not retire walkers, C1–C6, or live SSH qualification.

The next bounded size slice starts at `78a61e7b`. Its test-first gate is
`C0_CurlDirectorySizeTruth`: known/probed/unknown totals, wide/deep traversal,
structural refusal, root-versus-child failure, cancellation before and during I/O,
and ordered human-feedback callbacks. Run and archive the pre-fix Debug case before
changing the production walker, then compare the repaired Debug and test-enabled
Release C0 families. Native Copy/Move preflight and real-session qualification stay
separate. An intentionally failing baseline is not evidence that this route is done.

#### C0 directory-size repair and current continuation — 2026-09-05

The pause at `1337dee1` was explicitly superseded by the user's `continue`.
The current implementation uses the existing streaming FTP/SSH cursor, including
root lookup, with iterative outer traversal, coalesced feedback, unknown-size
truth, checked counters, eligible partial-child continuation and global-error
precedence. The pool clears per-call options on return, before idle publication
or disposal; final feedback is terminal during transport teardown. The lasting
contract is in the VFS, FTP/SSH, IMAP, shared-helper and performance specs.

| Evidence | Result and limitation |
|---|---|
| [Red size baseline](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_size_truth_red_debug/README.md) | 28/61 native checks; old buffered/recursive implementation. |
| [First iterative Debug](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_size_iterative_first_debug/README.md) | 77/77; blocked LIST cancels in 242 ms. Single wide/deep Debug samples are slower than red; no throughput gain claimed. |
| [Debug C0 candidate](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_size_candidate_debug/README.md) | 14/14 host cases, 79/79 size checks including idle-handle reset. |
| [First Release failure](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_size_release_failure/README.md) | 12 passed, one failed, size case skipped. Original failure retained. |
| [Release diagnostic](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_delete_release_diagnostic/README.md) | Malformed listing has correct status, contents and receipt; fake-server health wrongly rejects Winsock abort 10053 during disconnect. |
| [Release after fixture repair](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_size_release_repeat3_intermediate/README.md) | 42/42 across three C0 repetitions; 79/79 size and 127/127 Delete bounds each time. One witnessed abort with unchanged exact assertions. Shared z.dll Clean warning is tracked by C6. |

The final **85-check gate** covers 26 size scenarios, four ABI checks, two pool
lifetime checks and startup. Final cancellation/callback failure wins over
accumulated skipped-child partial totals without overwriting an earlier global
failure. This size-repair substep is qualified:

- [Debug C0](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_size_85_debug/README.md): 14/14, size 85/85.
- [Release C0](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_size_85_release_repeat3/README.md): 42/42 across three repetitions, size 85/85 in each.
- [Ordinary FTP](../../../4cb089111a23/FileOps/2026-09-05_c0_curl_size_85_ftp_regression_debug/README.md): 3/3.

All three builds are warning-free and bind source snapshot
`3663194d725a23644d1ff88b3eeb10eb4d3955600ed513571a41e0a791c190df`.
Final blocked-LIST cancellation is 246 ms in Debug and 242–250 ms in Release.
Final source/localization checks pass 197/197; spec inventory has zero blockers.
Checkpoint archive validation covers all 53 newly curated files and 2,287 files
in the complete governed inventory; patch/index whitespace checks pass. Metric
selection manifests bind the original raw stream and line order, not checkout-
dependent newline hashes.
Original failures and the intermediate z.dll warning remain archived. C0 remains
VERIFY for its other routes, not for a missing size-repair candidate.

IMAP still uses its buffered mailbox/UID facade. Applying operation control to its
request seam and charging returned storage do not qualify internal/transient
buffering or real transport cancellation. The retained-size metrics measure admitted
frame/cursor reservations, not whole-process or opaque TLS/socket memory. FTP/SSH
session limits remain unqualified, including a live cursor per ancestry frame.
