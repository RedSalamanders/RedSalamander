# Cross-Domain Correctness, Data-Safety, Validation-Integrity, and Architecture Remediation Plan

| Field | Value |
|---|---|
| Status | **COMPLETE** -- moved to `Done` on 2026-08-17 by explicit user direction after the five Fresh-run Commands failures were repaired and focused-green; the user explicitly waived a second Fresh Full run. |
| Priority | **P1** -- twenty-two P1 correctness, data-integrity, destructive-filesystem, ABI-boundary, build-attestation, and verdict-integrity defects (nine original plus thirteen from the 2026-08-14 follow-up), with additional P2/P3 reliability, performance, simplification, proof, and specification work, including the 2026-08-15 post-commit extension |
| Planned | 2026-08-10 at `83ab72733ecd9d3aa457db1d14465dd40ab498af` |
| Last audited | 2026-08-15 at `e9c76e59aa2b2993170ea1e6c88e7d15362954c7` |
| Closeout record | **Closed 2026-08-17 by explicit user direction.** Packages F0-A through F0-K and Packages A-J are complete and accepted. Package K's Fresh Full completed 1,901 cases with 1,844 passed, five failed Commands cases, and 52 skipped; every other suite passed. Those five failures were repaired and are focused-green on the final Debug binary. Tool/spec inventories, archive validation, localization contracts, and `git diff --check` are green. The user explicitly directed that no second Full run be started and accepted this documented validation exception before directing the plan to `Done`. |
| Package D checkpoint | **Complete and accepted 2026-08-16.** `CurlPublicationResult` plus its atomic accumulator preserve primary, cleanup, source-deletion, committed, rollback-eligibility, preserved-artifact, debt-mask, and debt-count truth across writer and single/recursive/batch Copy/Move/Rename paths. Committed cleanup-only debt remains successful while one content-free metric/diagnostic/localized modeless alert exposes retained artifacts; primary and source-deletion failures remain truthful. Upload/verification/prepare/promotion and safe-staging-cleanup failures are covered, the second writer Commit is idempotent, and armed rollback-delete failures remain unused with zero unsafe `DELE`. Final Debug receipt `3a5ac49657c6df9f5e229a330fc46be7f9a970fda4f4bc331b45b029a1f4ac5a`; focused result `221/0`. Five Release runs passed `221/0`; `CleanupDebtGuard` is 59–62 ms (60-ms median), 320 commands, and nine retained rollback items per run. L4-CURL-02/07 are closed; do not reopen without contradictory evidence. |
| Review window | `f48ac8855b2e42ad450caf66ced7bdddb627062d..83ab72733ecd9d3aa457db1d14465dd40ab498af` (the four-day review baseline through the reviewed `HEAD`) |
| Follow-up review window | `83ab72733ecd9d3aa457db1d14465dd40ab498af..e9c76e59aa2b2993170ea1e6c88e7d15362954c7`, including the committed Startrail workspace (`d24c2bde0`), the sibling UI/File Operations repair (`a210613a2`), and their merge (`e9c76e59a`); the reviewed worktree was clean |
| Ownership | Seventy-two findings/corrections: the original twenty-four production defects/architecture pitfalls plus three proof/specification corrections, and forty-five follow-up build/tooling, Chordforge, File Operations, FolderView, DxUi, test, performance, simplification, and policy findings. The same-IID Terminal item is explicitly not a current mixed-binary defect under the user-confirmed lockstep build/staging invariant. |
| File history | This is the original `CodeReview_LastFourDays_Remediation_2026-08-10.md` plan, renamed in place after the review expanded beyond the original four-day scope. There is one plan and no compatibility copy or tombstone. |
| Size / risk | **XL / CRITICAL**. The follow-up adds external-junction deletion/overwrite, false-green/full-verdict, source-identity, receipt/portable-attestation, and output-buffer overwrite paths. Curl conditional publication, Terminal placement accounting, graph rendering, WindowHost shutdown, and Monitor immutable snapshots remain separate architecture changes. |
| Original primary implementation commit | `6de1e8fa6` (`Complete Operation Cinderstar remediation`) |
| Follow-up implementation commits | `1ab7a3b40`, `a4faca9aa`, `7ca29d34a`, `bae5f9546`, `b11d54a6c`, `d24c2bde0`, `a210613a2`, and merge `e9c76e59a` |

## Executor contract

Read this entire plan before changing code. Reproduce every RED proof before its
corresponding repair, preserve the invariants below, and keep each task package
independently reviewable. Do not combine the Monitor storage redesign with the
smaller Curl, Terminal, or selected-path fixes in one implementation commit.

Before implementation, recheck drift:

```powershell
git diff --stat 83ab72733ecd9d3aa457db1d14465dd40ab498af..HEAD -- `
  RedSalamander Plugins/FileSystemCurl Plugins/Terminal RedSalamanderMonitor `
  Common Tests Tools Specs/FileSystem Specs/Plugins Specs/Terminal Specs/Core Specs/UI
git diff --stat 83ab72733ecd9d3aa457db1d14465dd40ab498af..HEAD -- build.ps1 Directory.Build.props Installer .github Specs/Build Specs/Testing
```

If relevant code has drifted, refresh line anchors and rerun the focused RED
proofs before editing. Later code is authoritative; this plan records reviewed
behavior and intended invariants, not permission to overwrite newer decisions.

Completed work is not closed merely because tests pass. Merge durable behavior
into the named authoritative specs, archive required performance evidence under
`Specs/TestRuns/`, move this file to `Specs/Plans/Done/`, and reconcile
`Specs/Plans/WIP/README.md` only after every done criterion is satisfied.

## How to read and execute this plan

This is one master ledger, not seventy-two equally urgent production bugs. Every
finding keeps its detailed evidence and repair below, but execution must use the
classification and package order in this section. A finding cannot be promoted
to a stronger class without the proof named in its detailed section.

### Truth classifications

| Classification | Finding IDs | Required action |
|---|---|---|
| **CONFIRMED DEFECT / RISK** | `L4-BUILD-01..06/08`; `L4-STAR-01..10/12/13`; `L4-CURL-01..08`; `L4-TERM-01..04/07`; `L4-FILEACTION-01..06`; `L4-MON-01..03`; `L4-CHORD-01..08`; `L4-DXUI-03/04`; `L4-FILEOPS-09/11..15` | Add a deterministic RED reproduction, implement the repair, update the authoritative spec, and run the focused plus final gates. |
| **RUNTIME PROOF REQUIRED** | `L4-DXUI-01/02` | Preserve the static finding, but do not claim user-visible impact until the COM-identity or owner-thread teardown test reproduces it. |
| **MEASUREMENT FIRST** | `L4-TERM-05`; `L4-FILEOPS-10` | Archive the deterministic baseline before selecting or implementing the fix. TERM-05 needs code changes only if active-conversion close latency fails its bound. |
| **CONTRACT / GOVERNANCE / SIMPLIFICATION** | `L4-BUILD-07`; `L4-CURL-09`; `L4-TERM-06`; `L4-FILEACTION-07/08`; `L4-STAR-11/14/15/16`; `L4-CHORD-09` | Correct the contract, diagnostics, dependency ownership, tests, or shared-helper use. Do not describe these as reproduced production failures. |
| **NOT A CURRENT DEFECT** | `L4-CHORD-10` | Document and test the confirmed lockstep host/plugin staging boundary. Do not build legacy mixed-binary compatibility unless the product boundary changes. |

### Ordered execution map

| Order | Existing package(s) in this file | Scope | Dependency/gate |
|---:|---|---|---|
| 1 | `F0-A` | Destructive ZIP and evidence-prune containment | Must land before other validation-evidence work. |
| 2 | `F0-B` | Terminal-result, promotion, aggregate, and exit truth | Requires contained evidence roots. |
| 3 | `F0-C` | Receipt, source snapshot, Full closure, and test-surface truth | Requires contained artifact paths. |
| 4 | `F0-D` | MSIX restore, portable payload provenance, and toolchain identity | Requires semantic receipt authority. |
| 5 | `F0-E`, `F0-F` | Selected-profile epochs, immutable evidence, canonical reads, origin safety, diagnostics, and dependency closure | Requires orders 2-4. |
| 6 | `F0-G`, `F0-H` | Chordforge ABI, command state, shortcut scope, Unicode identity, and message governance | Independent product lane. |
| 7 | `F0-I`, `F0-J` | DxUi identity/lifecycle/tooltip/DPI and File Operations graph/focus | Independent lane; measure graph before optimizing. |
| 8 | `A`-`D` | Curl capability and transaction integrity | Independent provider lane; fail closed where endpoint semantics are unproven. |
| 9 | `E`, `F` | Terminal Kitty cache, bounds, recovery, resize, and shutdown proof | Independent Terminal lane. |
| 10 | `G`, `H` | Selected-path manifest identity, grammar, streaming, and tests | Independent application lane. |
| 11 | `I`, `J` | Monitor immutable snapshot, thread admission, and canonical export | Independent Monitor lane. |
| 12 | `F0-K`, `K` | Specification ownership, WIP/Done reconciliation, Fresh Full, and exact Resume closeout | Runs only after every prior package is done or explicitly rejected. |

The detailed package definitions, verification matrix, checklist, STOP
conditions, and closeout rules near the end of this file remain authoritative.
This summary only makes their ordering and finding type explicit.

## Independent re-audit verdict (2026-08-10)

I agree with the outcome of all six original findings. The re-audit found four
material corrections to the first draft: the host already queries
`IFileSystemAtomicWriter`; the Curl cleanup defect affects every copy/move
publication path, not only `IFileWriter::Commit`; the Kitty repair must be a
bounded visible-working-set/delta cache rather than an all-images generation
copy; and selected-path cleanup does not need a live no-delete guard. The source
evidence also exposes twenty-one adjacent items: eighteen production
defects/pitfalls and three proof/specification corrections.

The four-day window is the review/discovery boundary, not a claim that every
adjacent defect was introduced there. `6de1e8fa6` introduced or materially
changed the Curl writer/import mismatch, asynchronous Kitty cache, selected-path
scavenger, and Monitor save snapshot/export path. L4-FILEACTION-01 was
pre-existing but was touched/entrenched by that commit. L4-CURL-03/04/05/06 and
parts of cleanup debt predate the window (some were already recorded in older
review ledgers) but remain unfixed at the reviewed `HEAD`; they are owned here
because this re-audit explicitly followed the modified transaction code and its
shared helpers/callers.

### Verdict on the original findings

| ID | Priority | Finding | Effort | Fix risk | Confidence |
|---|---:|---|---:|---:|---:|
| L4-CURL-01 | P1 | **AGREE, explanation corrected.** Imports are advertised but deterministically rejected because Curl lacks the host's atomic-writer opt-in and rejects the fallback no-overwrite writer. | S | Low/Medium | High |
| L4-CURL-02 | P1 | **AGREE, scope broadened.** Post-publication backup deletion can falsely fail writer, recursive/single/batch Copy, and single/batch Move; Move may already have deleted its source. | M | Medium | High |
| L4-TERM-01 | P1 | **AGREE, repair changed.** Geometry-only reveal can blank a valid Kitty image; use a generation-keyed bounded delta/working-set cache. | M | Medium | High |
| L4-FILEACTION-01 | P2 | **AGREE, repair simplified.** Path-only cleanup can delete a replacement; compare identity and disposition through one cleanup handle, without a persistent guard. | M | Medium | High |
| L4-FILEACTION-02 | P2 | **AGREE.** Shared `%TEMP%` filename shape is not ownership and can delete another workflow's old `rsaNNNN.tmp`. | M | Medium | High |
| L4-MON-01 | P2 | **AGREE.** Save As can synchronously copy up to 4 GiB / 5 million strings under the UI-thread document lock. | L | High | High |

### Additional finding inventory

| ID | Priority | Kind | Adjacent finding | Effort / risk | Confidence |
|---|---:|---|---|---|---:|
| L4-CURL-03 | P1 | Data integrity | Native Curl Copy/Move verifies the staged upload against the downloaded temp size, but never proves that download matches a known source size; Move can delete the complete source after publishing a short copy. | M / Medium | High |
| L4-CURL-04 | P1 | Data integrity | No-overwrite Curl publication is probe-then-unconditional-rename; a late creator can be overwritten and a Move can then delete its source. | L / High | High |
| L4-CURL-05 | P1 | Data integrity | Failure rollback deletes the current destination path/tree without immutable identity and can delete a concurrent replacement or child. | L / High | High |
| L4-CURL-06 | P2 | Contract | A known-size `CurlStreamingReader` can return successful premature EOF or bytes beyond its committed size. | M / Medium | High |
| L4-CURL-07 | P2 | Durability/privacy | Cleanup debt is warned about but has no safe reconciliation or explicit user recovery outcome; remote staged/old bytes can remain indefinitely. | M/L / Medium | High |
| L4-CURL-08 | P3 | Contract | `ALLOW_REPLACE_READONLY` without `ALLOW_OVERWRITE` returns `ERROR_NOT_SUPPORTED`, not required `E_INVALIDARG`. | S / Low | High |
| L4-CURL-09 | P3 | Architecture | Atomic-writer support is probed with raw task flags rather than the effective per-item grants later passed to the writer. Current S3 ignores the flags, so this is a future-provider pitfall. | S / Low | Medium/High |
| L4-TERM-02 | P1 | Resource safety | The claimed 4,096 Kitty-placement bound counts only accepted visible placements; iterator work and engine placement retention remain unbounded. | M/L / High | High |
| L4-TERM-03 | P2 | Performance/architecture | Capture deduplication and three-layer drawing use repeated linear image searches, producing quadratic UI-thread work. | M / Medium | High |
| L4-TERM-04 | P2 | Correctness | `CreateBitmap` failure neither schedules a bounded retry nor routes `D2DERR_RECREATE_TARGET` into device recovery, so an image can remain blank. | S/M / Medium | High |
| L4-TERM-05 | P3 | Test gap | Shutdown tests cancel queued work only; active maximum-size conversion has no stop polling and no close-latency proof. | S test, M fix / Low/Medium | High |
| L4-TERM-06 | P3 | Specification | The authoritative Terminal spec cites absent Kitty archive `2026-08-09_150404` while the actual final archive is `2026-08-09_165700`. | S / Low | High |
| L4-TERM-07 | P2 | Performance | Ordinary `WM_SIZE` destroys all valid Kitty bitmaps/render resources even though the HWND render target supports `Resize`, forcing progressive reupload. | S/M / Medium | Medium/High |
| L4-FILEACTION-03 | P2 | Data safety | `GetTempFileNameW` reserves then closes a path which is reopened with `CREATE_ALWAYS`; a replacement in that gap is truncated and filled with selected paths. | S / Low | High |
| L4-FILEACTION-04 | P1 | Record injection | Raw CR/LF/NUL-bearing provider paths are written into a CRLF record format and can become extra/truncated paths in an external action. | S / Low | High |
| L4-FILEACTION-05 | P2 | Performance | The UI command path copies the complete selection into another vector and then one aggregate string before writing. | S/M / Low | High |
| L4-FILEACTION-06 | P2 | Performance/recovery | Scavenging runs synchronously before every launch, can consume 20 ms, restarts at the first entry, and can starve eligible leftovers beyond the same first 128 entries. | M / Low | High |
| L4-FILEACTION-07 | P3 | Simplification | FileActionLauncher duplicates the canonical Windows argument-quoting grammar in `Common/ProcessCommandLine.h`. | S / Low | High |
| L4-FILEACTION-08 | P3 | Test architecture | Current source/runtime tests require the unsafe prefix scavenger, check only BOM/substrings, and never make the child parse the manifest. | M / Low | High |
| L4-MON-02 | P1 | Reliability | `std::thread`/`std::jthread` creation can throw through the Win32 command path after active state is published, terminating the process or stranding state. | S/M / Low | High |
| L4-MON-03 | P2 | Data correctness | Canonical exported text is not idempotent: opening `BOM + "a\n"` creates `{"a", ""}` and the exporter writes `"a\n\n"`; each open/save adds a blank record. | S/M / Low | High |

## Review and validation baseline

The audit covered 13 commits including two merges, 306 changed files, 27,061
additions, and 20,643 deletions in the review range. The following baseline was
green unless qualified below:

- Full x64 Debug rebuild: 0 warnings, 0 errors, 154.2 seconds; log
  `.build/logs/msbuild-20260810_193034_494.log`.
- Changed/source-contract Pester selection: 285 passed, 0 failed in 16.5 seconds.
- `FileSystemCurlTests.exe`: passed.
- `PluginContractTests.exe`: Curl 130/0 and Terminal 49/0. The first aggregate
  invocation had one transient Microsoft Drive internal failure (212/1); the
  immediate identical rerun passed 213/0 and the process exited 0.
- `Tools/Run-AllTests.ps1 -SkipBuild`: **incomplete, not a pass**. The outer shell
  limit stopped it after roughly 20 minutes. Commands ran for roughly 16 minutes,
  then the aggregate progressed through File Operations family 4/18. The traces
  showed progress and no new crash dump, but this run cannot be used as a release
  gate. Evidence is under
  `.build/TestSandbox/runs/20260810T173321Z-67752-f9553418c9504404a43d29d2c9009ecc/`.
- `git diff --check`: clean at the reviewed commit.

The existing green tests do not cover the production interleavings above. Each
production repair starts with a focused deterministic regression that fails on
the reviewed commit for the stated reason. Test/spec-only items start with a
link, inventory, exact-byte, or active-lifecycle proof that demonstrates the
current credibility gap; source-shape assertions alone are not closure.

## Non-negotiable architecture invariants

1. **Provider capabilities are executable truth.** A cross-filesystem import may
   be advertised only if the complete host-to-writer path implements it. A
   no-overwrite request must never be weakened into probe-then-replace or a
   racy existence check.
2. **Publication has a point of no return.** After the new object has replaced the
   destination successfully, failure to remove rollback debris is cleanup debt,
   not transaction failure. The operation reports success, warns once, retains
   diagnostics/metrics, and never encourages an unsafe retry.
3. **Curl source proof and rollback are identity safe.** A known source size is
   a commitment. Destructive Move cannot delete the source without an exact
   downloaded-byte proof, and failure recovery cannot delete a remote path/tree
   that may now belong to another actor. No-overwrite publication is conditional
   or fails closed; a probe is never permission to replace.
4. **Kitty cache completeness survives geometry changes.** Storage content
   generation and viewport geometry are different axes. Scrolling/resizing may
   change visible placements without changing Ghostty storage generation; every
   newly visible image must already be available or trigger bounded preparation.
5. **Kitty bounds cover retained state and work.** Count all engine placements,
   including virtual/offscreen entries, before filtering. CPU/GPU caches use a
   keyed bounded working set; limits are not raised to hide incomplete accounting.
6. **Temporary-file ownership is identity based.** A pathname or filename prefix
   is not ownership. Cleanup must delete only the exact file object created by
   RedSalamander and must not follow a replacement or reparse target.
7. **Crash scavenging has a private ownership boundary.** Never enumerate the
   shared `%TEMP%` root and infer ownership from `rsa????.tmp`. Scavenge only an
   application-owned directory with bounded, no-follow inspection.
8. **Selected-path manifests have a strict external format.** UTF-16LE BOM,
   selection order, one CRLF-terminated record per path, and rejection of
   NUL/CR/LF are explicit. Creation is one `CREATE_NEW` transaction, and large
   selections stream without an aggregate copy.
9. **Save As preserves an exact start snapshot without an O(retained-bytes) UI
   copy.** Later ingestion is excluded, the worker owns stable immutable content,
   and UI lock/return latency remains bounded as retained byte size grows.
10. **Worker admission is exception safe.** Catch only `std::system_error` around
    thread construction, publish active state only with rollback, and keep
    `std::bad_alloc` fatal under repository policy.
11. **Canonical Monitor export is idempotent.** Opening and re-exporting one
    canonical UTF-8+BOM/LF file cannot synthesize an additional blank record.
12. **Shared-helper and RAII rules still apply.** Before adding a helper, search
   `Specs/Core/Core_SharedHelpers.md`, `Common/`, and `Tests/TestSupport/`. Use WIL
   owners for handles/COM/resources; do not add manual cleanup or `catch (...)`.

## Finding L4-CURL-01 -- advertised Curl imports cannot execute

### Evidence and failure chain

- `RedSalamander/FolderWindow.FileOperations.State.cpp:9426-9433` correctly
  queries `IFileSystemAtomicWriter` and calls `SupportsAtomicWriterCommit` for
  the final path. Curl does not expose that interface, so the host selects its
  CSPRNG sibling fallback at lines 9435-9449, strips overwrite grants, and calls
  `CreateFileWriter` for a no-overwrite temporary destination.
- `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:1560-1567` rejects every
  writer creation that does not set `FILESYSTEM_FLAG_ALLOW_OVERWRITE`.
- `Plugins/FileSystemCurl/FileSystemCurl.h:83-89` exposes the stream provider but
  not `IFileSystemAtomicWriter`.
- Static FTP, SFTP, and SCP capabilities at
  `Plugins/FileSystemCurl/FileSystemCurl.h:223-310` advertise
  `crossFileSystem.import.copy` and `.move` as `["*"]`.

The host therefore enables an operation that the destination rejects before
transfer. Passing `ALLOW_OVERWRITE` is not an acceptable shortcut: the host temp
name is intended to be created without replacing an existing object, and an
existence probe followed by upload/rename cannot prove fail-if-exists on these
protocols.

### Selected safe design

The immediate repair is capability truthfulness: FTP/SFTP/SCP must advertise
empty cross-filesystem import copy/move policies until a separately reviewed
provider/ABI design supplies a remote publication primitive with the required
no-replace semantics. Keep supported export and same-provider operations
unchanged for this finding. Keep `operations.write=true`: explicitly
overwrite-authorized direct writers exist. Do not add a cosmetic
`IFileSystemAtomicWriter` implementation whose endpoint behavior cannot satisfy
the contract. L4-CURL-04 separately decides the unsafe no-overwrite semantics of
native same-provider mutations; disabling bridge imports does not close it.

### Implementation tasks

1. Add a host/provider integration regression using the real capability parser
   and cross-filesystem planner. The fixture must model the Curl writer contract:
   import was advertised, atomic-writer QI returns `E_NOINTERFACE`, the host
   selects a no-overwrite sibling, and writer creation rejects it. Prove the
   reviewed commit admits the task and then fails before transfer.
2. Change the three static capability documents to the schema-supported empty
   import policy. Verify the parser retains required `crossFileSystem` objects and
   disables COPY and MOVE into Curl before task creation.
3. Add direct capability-contract assertions for FTP, SFTP, and SCP. Retain their
   existing export assertions and protocol-specific same-provider tests.
4. Update user-visible enablement/routing tests so disabled imports cannot be
   invoked through a stale command state.
5. Update `Specs/FileSystem/FileSystem_FtpSftpScp.md`, the built-in provider table
   and capability examples in `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
   the host planning rules in `Specs/FileSystem/FileSystem_FileOperations.md`.
   State that import is deliberately disabled pending a provable no-replace
   primitive; do not describe it as an incidental implementation gap.

### Done criteria

- FTP/SFTP/SCP capability JSON remains valid v1 and returns `S_OK`.
- Cross-filesystem import COPY/MOVE into all three protocols is rejected before
  task creation; no endpoint command or destination mutation occurs.
- Existing cross-filesystem export and same-provider Curl behavior remain green.
- No code path passes overwrite authorization merely to work around the rejected
  writer call.

## Finding L4-CURL-02 -- successful overwrite can report failure

### Evidence and failure chain

- `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp:1081-1110`
  (`PromoteStagedFileToDestination`) renames the staged object to the final path,
  then returns `FinalizeOverwriteTarget(...)` directly when the caller does not
  request the backup path.
- The new writer path calls this mode at approximately
  `FileSystemCurl.CopyMove.cpp:2125` and propagates the cleanup HRESULT.
- `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:269-318`
  (`TempFileWriter::Commit`) leaves `_committed == false` when that HRESULT fails,
  even though the destination already contains the new bytes.
- `CompleteSuccessfulOverwrite` at
  `FileSystemCurl.CopyMove.cpp:1003-1012` already encodes the intended
  warning-only post-publication cleanup behavior.
- Recursive Copy reaches the same helper at approximately lines 1472-1481 and
  1759-1768; single and batch Copy reach it at lines 2257-2266 and 2854-2863.
- Single and batch Move retain the backup through source deletion, then return
  `FinalizeOverwriteTarget` directly at lines 2417-2445 and 3203-3231. In these
  paths the source is already gone when backup deletion turns success into
  failure.
- The only direct proof at lines 905-910 invokes
  `CompleteSuccessfulOverwrite` in isolation, so it can pass while production
  call sites bypass the helper.

This is a transaction truthfulness defect: callers can retry or report failure
after the operation has crossed its publication point of no return.

### Implementation tasks

1. Extend the Curl fake transport with deterministic rollback-sibling deletion
   failure after successful final promotion and, for Move, after successful
   source deletion.
2. RED-prove writer, single Copy, recursive Copy, batch Copy, single Move, and
   batch Move. Assert exact destination bytes, remaining backup, callback/result,
   notification behavior, source presence/absence, and cleanup-debt observation.
3. Represent primary publication and cleanup as separate named outcomes, or route
   every post-publication deletion through one success-preserving helper. Do not
   suppress staging upload, source-size proof, promotion, rollback, restoration,
   or source-deletion errors that occur before the point of no return.
4. On post-publication cleanup failure, preserve primary `S_OK`, mark writer or
   item committed, publish committed size, issue the normal mutation notification
   once, log once, and record structured cleanup debt. Batch aggregation and
   per-item callbacks must observe the same primary success.
5. Confirm destructor/abandonment and retry logic cannot republish or remove the
   successfully published destination after the warning-only result.

### Done criteria

- Injected backup cleanup failure produces successful writer/Copy/Move outcomes
  with exact new destination bytes and one observable warning/debt record per
  affected item; Move's source remains deleted after its successful deletion.
- A second `Commit()` follows the existing committed-state contract and cannot
  republish data.
- Pre-publication failures still fail and restore the prior destination exactly.
- Existing command-count and transactional-writer performance evidence is either
  unchanged or refreshed if the observable command sequence changes.

## Finding L4-CURL-03 -- native Copy/Move lacks source-size proof

### Evidence and failure chain

- `CopyFileViaTemp` receives `expectedSizeBytes` at
  `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp:1140-1150`, downloads to a
  local temporary file, then reads the actual temp size at lines 1211-1218.
  It never compares those values.
- Upload and remote staged verification use the temp file's actual size at lines
  1230-1280, so a short local download is internally self-consistent at the
  destination.
- `CurlDownloadToFile` treats `CURLE_OK` as success at
  `FileSystemCurl.Shared.cpp:3154-3219` without an independent expected-byte
  check.
- Copy/batch task construction retains `sourceInfo.sizeBytes` but discards
  `sourceInfo.sizeKnown` around lines 2675-2684 and 2988-2997. Move can delete
  the source after destination publication at lines 2431-2444 and 3217-3230.

### Repair and proof

1. Propagate `sizeKnown` with each native Curl file task. Prefer the targeted
   per-file size probe over a directory-listing estimate when available.
2. When the source size is authoritative, compare the completed local temp size
   before any upload. A mismatch returns `ERROR_PARTIAL_COPY`, publishes no
   destination, and preserves the source.
3. A destructive Move requires an authoritative size or equivalent immutable
   content proof before destination mutation/source deletion. If the protocol
   cannot provide one, fail closed before upload. Copy may retain an explicit
   source-size-unknown path only because the source remains; document that weaker
   diagnostic rather than pretending it is a destructive proof.
4. Add fake-transport clean-short RETR and stale/unknown-size cases for single,
   recursive, and batch Copy/Move. Assert zero promotion/source deletion on a
   known-size mismatch and no source deletion when proof is unavailable.
5. Reconcile this with L4-CURL-06 so native copy and direct `IFileReader`
   consumers share one known-size commitment rule.

### Done criteria

- Known-size short downloads cannot be uploaded or reported successful.
- Native Move never deletes a source whose downloaded bytes were not positively
  proven against an authoritative source commitment.
- Unknown-size behavior is explicit in the Curl/File Operations specs and tests;
  no zero/default size is silently treated as authoritative.

## Findings L4-CURL-04/05 -- conditional publication and rollback identity

### Evidence and failure chains

- `PrepareOverwriteTargetForRename` at
  `FileSystemCurl.CopyMove.cpp:923-940` implements no-overwrite as an existence
  probe; publication occurs later through `RemoteRename` at lines 1027/1095.
  FTP uses RNFR/RNTO and SFTP/SCP use libcurl `rename`
  (`FileSystemCurl.Imap.cpp:3864-3884`). The fake FTP RNTO replaces an existing
  destination at `FileSystemCurl.SelfTest.Ftp.cpp:445-457,789-795`.
- The writer implementation itself acknowledges that these transports lack a
  portable no-replace primitive at `FileSystemCurl.DirectoryOps.cpp:1560-1563`.
- `RestoreOverwriteTargetAfterFailure` at
  `FileSystemCurl.CopyMove.cpp:963-990` deletes whatever now occupies the final
  path before restoring a backup. `RollbackMovedFileDestination` at lines
  1037-1055 does the same after a source-delete failure, and recursive Copy
  rollback deletes the destination tree at lines 1072-1078.
- The host bridge already states the correct safety principle at
  `FolderWindow.FileOperations.State.cpp:9998-10021`: after publication, a
  by-path delete can remove a concurrent replacement, so uncertain content is
  retained and the task fails partial.

### Required architecture decision

There is no safe check-then-rename/delete repair. Choose one reviewed option per
protocol before implementing ordinary no-overwrite mutation:

1. Add a provider/transport primitive that atomically publishes only if the
   destination is absent and conditionally deletes/restores only the immutable
   object/revision owned by the operation; or
2. fail closed for no-overwrite Curl copy/move/rename paths and make capabilities
   and UI enablement truthful until that primitive exists.

For rollback without immutable identity, the safe fallback is preservation:
retain source, published destination, and rollback/staging artifacts as
applicable; return a precise partial result and diagnostic; never delete a path
or recursively remove a tree merely because this operation previously used the
same pathname. Endpoint serialization inside one plugin instance is insufficient
because external clients remain concurrent.

### RED proof and implementation tasks

1. Extend the fake transport with barriers that create/replace the destination
   after the absence probe, after publication, and before file/tree rollback.
2. Prove a no-overwrite single/batch Copy, Move, and Rename currently replaces a
   late destination; prove rollback can delete a concurrent file or child.
3. Record a protocol matrix for FTP, SFTP, and SCP: atomic create-if-absent,
   conditional rename, immutable identity/revision, conditional delete, and
   directory-tree behavior. Do not infer semantics from command names.
4. Implement the selected option behind explicit provider APIs/capabilities.
   Passing overwrite authorization to a random host temp, serializing a mutex, or
   adding a second existence probe is forbidden.
5. Make failure recovery a structured result distinguishing primary mutation,
   source deletion, rollback eligibility, and preserved cleanup debt. Callbacks,
   batch aggregation, refresh notifications, and UI diagnostics must reflect the
   actual state.
6. Update `Specs/FileSystem/FileSystem_FtpSftpScp.md`,
   `Specs/FileSystem/FileSystem_FileOperations.md`, and
   `Specs/Plugins/Plugins_VirtualFileSystem.md` with the selected protocol/ABI
   contract. Any capability ABI change needs its own compatibility decision.

### Done criteria

- A late destination creator always wins a no-overwrite race; its bytes are never
  replaced.
- Failure rollback never removes a concurrent replacement or independently added
  directory child.
- Move source deletion occurs only after conditional publication/source proof;
  uncertain recovery preserves data and reports partial state truthfully.
- Every enabled Curl capability has a tested endpoint primitive that enforces the
  advertised semantics.

## Finding L4-CURL-06 -- known-size reader can return successful premature EOF

`CurlStreamingReader::GetSize` commits `_sizeBytes` when `_sizeKnown` is true
(`FileSystemCurl.DirectoryOps.cpp:415-424`), but `Read` breaks on `_eof` and
returns whatever is buffered without comparing `_positionBytes` to that
commitment (lines 507-566). `WorkerMain` sets `_eof=true` for every `CURLE_OK`
result at lines 775-803. Existing tests cover a normal full request and an
unknown-size tail, not a known-size short/overlong response. This violates
`IFileReader`'s committed-length contract in `Common/PlugInterfaces/FileSystem.h`
and `Specs/Plugins/Plugins_VirtualFileSystem.md`.

Repair the reader so transport completion is validated against the committed
start offset/length, latching `ERROR_PARTIAL_COPY` for premature EOF and rejecting
overrun/inconsistent bodies. Seek/retry generations must reset and revalidate the
right range without leaking stale completion. Add known-size exact, short,
overlong, seeked-range, and generation-restart tests through the real reader
boundary. Direct viewers/search/compare must never observe `S_OK + 0` before the
promised length.

## Finding L4-CURL-07 -- cleanup debt has no safe recovery outcome

`CompleteSuccessfulOverwrite` says failed backup deletion is left for "later
maintenance" (`FileSystemCurl.CopyMove.cpp:1003-1012`), while several staging
cleanup calls discard their HRESULT and the generated names contain no durable
object identity. There is no production reconciler. Old overwritten or staged
bytes may therefore remain indefinitely, consuming storage and retaining
possibly sensitive content.

1. Extend the structured publication result from L4-CURL-02 with a content-free
   cleanup-debt identifier/count and primary/cleanup HRESULTs.
2. Surface a precise localized warning/manual recovery outcome. Metrics and
   archived evidence must not contain endpoint URLs, credentials, usernames, or
   path content.
3. Automatic delayed deletion is allowed only after L4-CURL-05 supplies immutable
   conditional identity. Never scavenge remote names by marker/prefix alone.
4. Correct the FTP/SFTP/SCP spec's unconditional cleanup language and remove the
   unsupported promise of unspecified "later maintenance".

Done means cleanup failure preserves primary success, is observable/actionable,
and cannot silently retain bytes while claiming cleanup completed. If immutable
identity remains unavailable, explicitly document/manualize the debt rather than
performing unsafe path-based reconciliation.

## Findings L4-CURL-08/09 -- flag-contract cleanup

- Validate `FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY` without
  `FILESYSTEM_FLAG_ALLOW_OVERWRITE` before Curl's generic no-overwrite rejection
  and return required `E_INVALIDARG`. Add one direct provider contract case.
- In the host bridge, compute the effective final-writer flags from per-item
  overwrite/read-only grants before atomic probing. Pass the same effective flags
  to `SupportsAtomicWriterCommit` and the final-path `CreateFileWriter`; compute
  the stripped sibling flags only after the provider declines atomic publication.
  Add a flag-sensitive fake atomic provider. Current S3 ignores probe flags, so
  this closes an architecture/test pitfall rather than a current S3 failure.

## Finding L4-TERM-01 -- scrolling can reveal a blank Kitty image

### Evidence and failure chain

- `Plugins/Terminal/Terminal.cpp:4154` submits image conversion only when the
  Ghostty storage generation changes.
- The capture loop skips placements outside the current viewport at
  `Terminal.cpp:4194-4199` before copying image bytes, and image bytes are copied
  only for submitted work at approximately `Terminal.cpp:4216`.
- `consumeKittyReady` replaces `_kittyImages` with the newly converted visible
  set at `Terminal.cpp:4317-4330`.
- `drawKittyLayer` skips a placement whose image is absent from that map at
  `Terminal.cpp:4381-4389`.
- The tracked Ghostty candidate contract states that geometry-only scroll/resize
  does not bump storage generation
  (`External/TerminalEngine/CandidatePatches/ghostty-rs-v4.patch:2376-2379`).

Thus a content generation captured while image A is visible may cache only A.
Scrolling to already-stored image B changes placement visibility but not content
generation, no conversion is submitted, and B is skipped during draw.

### Selected bounded design

Use a generation-keyed visible-working-set/delta cache. Every capture recomputes
visible placement geometry, counts all iterator entries before filtering, and
discovers required `{imageId,imageGeneration}` keys missing from the CPU cache or
in-flight work. Missing keys may submit a bounded delta even when Ghostty storage
generation is unchanged. Accepted results merge into a capped keyed cache rather
than replacing the complete cache; current-visible keys are pinned during
eviction and nonvisible least-recently-used entries leave first.

Storage generation remains the content-mutation/stale-result fence, while image
generation is the pixel identity. One missing/oversized image must degrade only
its placements, not clear every Kitty image. The 64-MiB source/converted/storage,
16-million-pixel, latest-only queue, and one-upload-per-frame ceilings remain
hard. The current 4,096 check is **not** a real total-placement bound and is fixed
under L4-TERM-02.

Do not copy every referenced image on each storage-generation change. A valid
64-MiB source store can expand beyond 64 MiB as BGRA (especially grayscale), and
all-image copying would increase engine-lock time and reject otherwise renderable
visible content.

### Implementation tasks

1. Add a runtime selftest with two distinct Kitty images separated by scrollback:
   capture/render with A visible, scroll until B is visible without mutating
   terminal storage, capture/render again, and assert B has pixels and A/B
   generation identities remain correct. Record that storage generation was
   unchanged across the geometry change.
2. Add resize/reflow coverage where visibility changes without image content
   mutation. Include an image deletion/content-generation transition so stale
   images do not live forever.
3. Add one keyed cache/index used for missing-key discovery, work deduplication,
   result merge, lookup, and eviction. Avoid a second parallel vector-scan
   mechanism. Copy a missing image's bytes under `_terminalMutex`, then convert
   outside the lock as today.
4. Preserve latest-only worker replacement, stale result rejection, device-loss
   CPU reuse, cancellation, and unload quiet point. Do not read Ghostty-owned
   bytes after releasing `_terminalMutex`.
5. Track/cache-test aggregate source, active, ready, converted, cached, pinned,
   evicted, and uploaded bytes plus missing-key and per-image rejection counts.
6. Replace or narrow any source-shape-only assertion in
   `Tools/Tests/TerminalPluginSourceContracts.Tests.ps1` that would bless the old
   ordering. Runtime behavior is authoritative; retain only structural guards for
   ownership, limits, and teardown.
7. Update the core runtime and Operation Cinderstar Kitty sections in
   `Specs/Terminal/Terminal_EmbeddedPlugin.md` to state the geometry/content
   generation distinction and cache-completeness invariant.

### Performance proof

Use the existing Kitty maximum-size test and archived Release procedure as the
baseline. Add a deterministic scroll-reveal scenario recording at least:

- `terminal.kitty.capture_lock_us`;
- unique referenced/visible image counts;
- source, queued, active, ready, converted, and cached CPU bytes;
- bitmap upload count/bytes per frame;
- worker submit/convert latency and stale-result count.

Run five repeats on the same machine/configuration before and after. Archive the
scenario, exact command, commit/configuration, summaries, and result JSON under a
new `Specs/TestRuns/<machine>/Terminal/` directory. No regression may be hidden by
raising the 64-MiB or 16-million-pixel ceilings.

### Done criteria

- Scrolling/resizing to an already-stored image renders it without requiring a
  content mutation.
- Storage mutation/deletion cannot render a stale generation.
- Simultaneously visible content beyond policy degrades deterministically per
  image; it never clears an otherwise valid generation or raises the ceilings.
- Capture never uses Ghostty-owned bytes after releasing the session lock.
- Memory, latest-only queue, per-frame upload, cancellation, and unload limits
  remain green with archived same-machine evidence.

## Finding L4-TERM-02 -- placement bound does not cover iterator/storage work

The check at `Plugins/Terminal/Terminal.cpp:4161-4165` uses
`placements.size()`, but that vector grows only after virtual entries are skipped
at lines 4172-4184 and offscreen entries at 4194-4199. The pinned engine limits
image bytes but its placement insertion path has no count/allocated-byte
admission. Host paint can therefore scan arbitrarily many virtual/offscreen or
repeated placements while holding `_terminalMutex`, and engine memory can grow
independently of the 64-MiB image cap.

1. RED-test more than 4,096 placements referencing one one-pixel image, with
   visible, offscreen, virtual, and repeated variants. Record total iterated,
   accepted-visible, rejected, engine-retained bytes/count, and lock duration.
2. Count every iterator result before any filter and terminate/degrade
   deterministically at the host work ceiling. Never describe a visible-vector
   length as total admission.
3. Add placement count and allocated-byte admission/accounting to the pinned
   Ghostty runtime, or stop for the required runtime/ABI patch decision. Define
   deterministic protocol rejection/eviction semantics and preserve image-byte
   limits separately.
4. Add boundary and +1 tests plus hostile-output/close/unload coverage. Archive
   same-machine Release paint-lock and retained-memory evidence.

Done means terminal output cannot create unbounded placement storage or
unbounded per-frame iterator work, including when every placement reuses one tiny
image.

## Finding L4-TERM-03 -- quadratic image lookup and deduplication

Capture linearly scans `work.images` for every placement at
`Terminal.cpp:4218-4223`; drawing walks placements for three z layers and
linearly scans `_kittyImages` for each match at lines 4366-4389. Near the nominal
placement/image limits this is millions of UI-thread comparisons.

Close this as part of the L4-TERM-01 keyed cache: use one
`{imageId,imageGeneration}` index for work deduplication, cache lookup, merge, and
LRU state; avoid raw pointers/iterators that become stale across vector eviction.
Add high-cardinality/repeated-image tests and before/after comparison counts,
capture-lock time, and frame time. The keyed design must simplify the current
parallel scans rather than adding another cache alongside them.

## Finding L4-TERM-04 -- bitmap upload failure has no recovery state

`ensureKittyBitmap` sets `_kittyUploadDeferred` only for the per-frame budget
(`Terminal.cpp:4346-4350`). `CreateBitmap` failure at lines 4352-4363 returns
false without a retry/recreate disposition; the frame schedules another paint
only for budget deferral. A transient failure can stay blank, and a
`D2DERR_RECREATE_TARGET` from `CreateBitmap` never reaches the existing
post-`EndDraw` recovery path.

Return a structured upload disposition/HRESULT. Route recreate-target failures
to a post-frame device-resource discard, schedule bounded/backed-off retry for
approved transient failures, and latch a stable per-image failure for persistent
resource exhaustion so paint cannot spin. Add injected one-shot failure,
persistent OOM, and recreate-target tests with exact repaint/upload counts.

## Finding L4-TERM-05 -- active conversion shutdown proof is missing

`TerminalKittyImagePipeline::Convert` has no stop token and its pixel loop
(`TerminalKittyImagePipeline.cpp:248-284`) cannot observe shutdown; `Stop()` joins
the worker. The current test pauses work before the worker takes it, proving only
queued cancellation. Add an active maximum-size conversion barrier, request
close/unload after conversion starts, and measure join/quiet latency. If the
measured supported-hardware bound is not acceptable, pass the stop token into
conversion and poll at bounded row/chunk boundaries. This is a proof gap until
the active measurement demonstrates a violation; do not claim cooperative
cancellation without it.

## Finding L4-TERM-06 -- authoritative evidence link is broken

`Specs/Terminal/Terminal_EmbeddedPlugin.md:756-768` cites absent archive
`Specs/TestRuns/4cb089111a23/Terminal/2026-08-09_150404/`. The later final section
correctly cites existing `2026-08-09_165700`. Remove or label the preliminary
non-archived paragraph and retain one authoritative final closeout narrative.
Add an evidence-link inventory check so a normative archive reference cannot
silently point to a missing directory.

## Finding L4-TERM-07 -- ordinary resize discards reusable GPU resources

`WM_SIZE` calls `discardDeviceResources` at `Terminal.cpp:5919-5923`, resetting
every Kitty bitmap and render/text resource. Paint already calls the supported
HWND render-target `Resize` at lines 4853-4858, while uploads are capped at one
bitmap per frame. Interactive resize therefore causes avoidable texture churn
and progressive blanking.

Split ordinary size handling from DPI/theme/device-loss invalidation. Preserve
the render target and Kitty bitmaps on `WM_SIZE`, resize the target, recompute
terminal geometry/scrollbar, and reserve full discard for DPI/resource-policy
changes or actual recreate-target failure. Add a resize-storm test recording
bitmap creation/upload count, retained CPU/GPU bytes, final geometry, and visual
presence; retain existing DPI/device-loss tests.

## Findings L4-FILEACTION-01/02 -- selected-path lease ownership

### Evidence and failure chains

- `RedSalamander/FileActionLauncher.h:41-59` stores only a pathname in the
  launch-lifetime lease.
- `RedSalamander/FileActionLauncher.cpp:507` closes the creation handle before the
  launched process uses the list. Later cleanup calls
  `std::filesystem::remove(file)` by path at `FileActionLauncher.cpp:321-338`.
  If the child deletes, renames, or replaces that path, RedSalamander can delete
  the replacement.
- Ownership for crash scavenging is inferred only from `rsa` plus four hex digits
  plus `.tmp` at `FileActionLauncher.cpp:341-357`.
- The code enumerates the shared temp directory and deletes old matches at
  `FileActionLauncher.cpp:368-416`, invoked at approximately lines 450-453.
- The current Commands selftest intentionally creates `rsa0001.tmp` and expects it
  to be deleted
  (`RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp:3173-3187`),
  which codifies the unsafe shared-namespace heuristic.

### Selected design

Create manifests only beneath a private RedSalamander root derived from
`AppDataPaths::GetLocalAppDataPath()` (with a deterministic test-sandbox
override). Validate the root as an application-owned non-reparse directory and
use a CSPRNG/GUID name with bounded `CREATE_NEW` retry. Record the created
`FILE_ID_INFO` in the live move-only lease, then close the creation handle before
launch for arbitrary-reader interoperability.

Do **not** keep a persistent no-delete guard. At cleanup, open the candidate once
with `DELETE | FILE_READ_ATTRIBUTES`, `OPEN_EXISTING`,
`FILE_FLAG_OPEN_REPARSE_POINT`, and `FILE_SHARE_READ | FILE_SHARE_WRITE` (omit
delete sharing for the short compare/disposition interval). Compare identity and
delete through that same handle with `SetFileInformationByHandle`. If the path is
absent, busy, changed identity, a directory/reparse point, or unqueryable, leave
it untouched and record a bounded diagnostic. Never check identity and then call
a pathname delete.

Crash scavenging is allowed only inside the validated private root. A restart no
longer has the in-memory expected identity, so its contract must be explicit:
the private manifest namespace is disposable after the age threshold, or a
durable identity record must be added before claiming exact-object recovery.
Never imply that a filename alone proves cross-restart identity. Enumeration is
entry/time bounded, no-follow, rate-limited/off the launch-critical path, and
must make progress across successive passes.

### Implementation tasks

1. Search `Common/PathUtils.h`, `Common/LocalFileTransaction.h`,
   `Common/DeleteOnCloseTemporaryFile.h`, and test-sandbox helpers before adding
   primitives. Put reusable file-identity/no-follow disposition logic in
   `Common/` only if its semantics match another consumer; otherwise document why
   the launch-lifetime lease is intentionally FileAction-local.
2. Introduce a test-overridable private selected-path root using the existing
   `AppDataPaths` owner. Validate/revalidate its identity and reparse status;
   create a long-path-safe GUID/CSPRNG candidate directly with `CREATE_NEW`.
   Remove the 16-bit `GetTempFileNameW` reservation/reopen sequence.
3. Store expected `FILE_ID_INFO` (and volume identity if needed) in the live
   lease. Close the creation handle after complete flush/write; do not retain a
   guard that imposes undocumented child share policy.
4. Implement exact-object cleanup with a no-follow opened handle and identity
   revalidation immediately before handle-based disposition. Treat absent or
   identity-mismatched paths as safe non-deletion, not a retry loop.
5. Rewrite scavenging to enumerate only the private root, with maximum entries,
   deletions, elapsed time, and a continuation/progress policy. Run once at
   startup or rate-limit it off the launch-critical path. Record
   inspected/deleted/skipped/error/bound counts; skip reparse points and folders
   and leave uncertain objects untouched.
6. Replace the unsafe `rsa0001.tmp` expectation with these deterministic tests:
   - normal live lease cleanup deletes the exact manifest;
   - a same-path replacement with a different file identity survives cleanup;
   - an old matching name in shared `%TEMP%` survives;
   - an old regular manifest in the private root is scavenged;
   - fresh, live/locked, directory, junction/symlink/reparse, and out-of-root
     entries survive;
   - enumeration/deletion/time bounds stop excessive work;
   - launch failure and process-close paths release every WIL handle once.
7. Make the representative child actually open and parse the manifest while the
   launched process owns it; test exclusive/ordinary reader share modes and a
   child that deliberately replaces the path. Identity-safe cleanup must preserve
   the replacement without requiring a persistent guard.
8. Update `Specs/Core/Core_SharedHelpers.md` for the relationship to the existing
   delete-on-close helper. Add the selected-path manifest ownership/lifetime
   contract to `Specs/UI/UI_CommandMenuKeyboard.md` unless repository review
   identifies a more specific current authoritative File Actions runtime spec.

### Done criteria

- Cleanup deletes the original file object and cannot delete a same-path
  replacement, reparse target, directory, or out-of-root object.
- No product or selftest code scans shared `%TEMP%` for `rsa????.tmp` ownership.
- Crash leftovers in the private root are reclaimed within explicit bounds while
  fresh/live/uncertain entries remain untouched.
- Representative external launch consumption succeeds without a persistent
  RedSalamander handle, and cleanup's short handle interval is fail-safe when the
  child still has an incompatible share.
- Commands selftests and teardown handle-leak checks pass repeatedly.

## Finding L4-FILEACTION-03 -- reservation/reopen can truncate a replacement

`GetTempFileNameW` creates and closes a 16-bit-namespace file at
`FileActionLauncher.cpp:456-464`; line 467 reopens the pathname with
`CREATE_ALWAYS`. A same-user replacement in the gap is truncated and populated
with selected paths. The private-root `CREATE_NEW` design above closes this
finding, but it needs its own RED collision barrier: replace the reservation
before reopen on the reviewed code, prove the replacement is truncated, then
prove bounded exclusive creation never opens/truncates an existing object. Cover
long roots, candidate collision exhaustion, and cleanup of only the successfully
created identity.

## Finding L4-FILEACTION-04 -- selected-path record injection

`FileActionLauncher.cpp:475-483` appends raw provider paths plus CRLF.
`FolderView.Enumeration.cpp:434-454` accepts provider names without excluding
CR/LF/NUL, and selection turns them into full paths at
`FolderView.Selection.cpp:953-963`. The external manifest contract currently
defines neither encoding nor record validity. A crafted path can therefore become
multiple apparent records or be truncated by a consumer.

1. Make the authoritative format exact: UTF-16LE BOM, original selection order,
   one CRLF-terminated record per nonempty path, no normalization/deduplication,
   and no embedded NUL/CR/LF.
2. Validate every record before creating the manifest or launching a process. An
   unrepresentable record returns a precise validation HRESULT/localized action
   failure and leaves no manifest/process. Do not invent escaping that existing
   external tools do not understand.
3. Test exact bytes for empty/single/multiple/non-ASCII/long paths plus NUL, CR,
   LF, and CRLF rejection. Include a fake provider-selected path through the real
   command route and make the child parse records exactly.

## Finding L4-FILEACTION-05 -- aggregate selection copies block the UI path

The command already materializes selected paths; `BuildExternalLaunchPlan`
copies the vector again at lines 757-763, then `CreateSelectedPathsFile` copies
every path into one aggregate `std::wstring` at lines 475-503. This happens
synchronously from `FolderWindow.Viewers.cpp:1822-1852`.

Pass a span/view of the existing selection, handle the focused-item fallback with
a bounded one-element view, write the BOM once, then validate and stream each path
and CRLF via `Common::HandleIo::WriteAll`. Use checked byte accounting and retain
the transactional cleanup lease; do not add a second full payload. Add a 10,000-
selection deterministic scenario reporting UI command/build/write time, path
bytes, write calls, and peak additional bytes. Archive same-machine before/after
Release evidence because this changes a user-visible synchronous path.

## Finding L4-FILEACTION-06 -- scavenging hitches and can starve leftovers

Scavenging currently runs before every manifest at lines 450-454, may consume 20
ms, counts every directory entry, ignores per-delete failures, and restarts from
the directory beginning. In a populated root, eligible files beyond the same
first 128 entries can starve indefinitely.

Use the private-root design, run recovery once at startup or behind a monotonic
rate gate outside launch-critical work, and retain a cursor/deterministic rotation
so successive bounded passes make progress. Tests must cover a dense prefix of
fresh/nonmanifest entries, eligible entries beyond the first bound, delete/share
failures, time cutoff, and multi-pass convergence. Metrics remain content-free.

## Finding L4-FILEACTION-07 -- duplicate argument quoting

`FileActionLauncher.cpp:39-71` reimplements the Windows argv grammar already
owned by `Common::Process::QuoteWindowsCommandLineArgument` in
`Common/ProcessCommandLine.h`. Replace the ordinary unwrapped macro path with the
canonical helper and retain its existing conformance corpus. The content-only
operation for macros already wrapped by template quotes has different semantics:
either add a clearly named shared append-content primitive with tests/catalog
entry, or keep the local variant with a definition comment and reviewed
source-contract exception. Do not force unlike semantics through one misleading
helper.

## Finding L4-FILEACTION-08 -- tests enforce the unsafe implementation

- `Tools/Tests/DeleteOnCloseTemporaryFileSourceContracts.Tests.ps1:50-58`
  requires the `rsa????.tmp` prefix scavenger.
- Commands serialization tests check a BOM and substring presence rather than
  exact records (`Commands.SelfTest.Settings.cpp:3014-3028`).
- The SearchService child only calls `GetFileAttributesW` on the manifest around
  `RedSalamanderSearchService/Main.cpp:3412-3415`; it never opens/parses it.

Replace those assertions with runtime boundaries for private-root exclusive
creation, strict bytes/records, child read compatibility, original/replacement
identity, no-follow cleanup, root reparse/long-path behavior, and bounded recovery
progress. Keep only legitimate structural guards. Coordinate any broader
source-contract migration with Operation Astrolabe; this plan owns the assertions
directly invalidated by the repair.

## Finding L4-MON-01 -- Save As deep-copies retained text on the UI thread

### Evidence and failure chain

- `RedSalamanderMonitor/RedSalamanderMonitor.cpp:2877-2890` captures the complete
  document synchronously before starting the save worker.
- `RedSalamanderMonitor/Document.cpp:950-959` holds the document lock and copies
  every retained `std::wstring` into a new snapshot.
- `Common/Common/SettingsStore.cpp:3392-3395` permits 5,000,000 retained lines and
  4 GiB of retained text.
- `Specs/Core/Core_RedSalamanderMonitor.md:188-191` correctly requires the exact
  export boundary, while lines 220-221 already separate snapshot and UI-return
  timings. The implementation satisfies exactness by doing O(retained bytes) work
  and allocation on the UI thread.

At configured limits this can freeze interaction, block ingestion behind the
document lock, temporarily approach a second copy of retained text, and fail due
to transient allocation pressure before asynchronous I/O begins.

### Target architecture

Replace the deep string snapshot with an immutable block-level snapshot. The
document should retain text in bounded blocks/chunks that become immutable when
sealed. `MonitorTextSnapshot` owns shared immutable block handles plus first/last
slice offsets and exact line/byte metadata. Snapshot capture under the document
lock copies only block handles and freezes/copies at most the bounded active tail;
the worker serializes those immutable slices after the UI releases the lock.

Retention advances the front slice/removes document ownership without invalidating
an in-flight snapshot. Appends after capture use a new/mutable tail, so the worker
cannot observe later ingestion. Avoid a permanent parallel full-text journal or
per-save full-string duplication. Choose a block size from measured UI capture,
append, lookup, and memory overhead rather than intuition.

This is an architecture slice, not permission to silently change visible line
indexing, metadata, filtering, selection, search, or the exact newline format of
saved files.

### Implementation tasks

1. Add a deterministic RED stress selftest using a practical large retained
   dataset (large enough to expose scaling without requiring a 4-GiB CI fixture).
   Record retained bytes/lines, snapshot lock time, UI return time, copied text
   bytes, shared block/handle count, active-tail copied bytes, and peak retained
   snapshot bytes. On the reviewed commit, assert copied text bytes scale with the
   entire retained payload.
2. Add functional concurrency fixtures:
   - capture, append more lines, and verify the file contains exactly the starting
     snapshot;
   - capture, trigger retention eviction/clear/replacement, and verify the worker's
     snapshot remains exact and valid;
   - cancel/close while saving and verify bounded shutdown and no late HWND use;
   - preserve empty document, blank-line, embedded-newline, structured metadata,
     and final newline behavior.
3. Introduce a block snapshot model with explicit maximum active-tail size and
   overflow-checked line/byte accounting. Reuse blocks for UI and export storage;
   do not keep a second permanent full representation solely for Save As.
   Move the neutral snapshot/storage types out of `MonitorFileReader.h` into a
   reader-independent header so `Document` and the exporter do not depend on an
   input adapter merely to share the model.
4. Adapt `Document` indexing/mutation behind its existing public behavior. Keep
   read/write locking correct and make snapshot creation allocation-failure policy
   explicit. `std::bad_alloc` remains fatal per repository policy; do not catch it
   as a normal save failure.
5. Update the save worker to stream immutable blocks/slices and preserve the
   existing progress, cancellation, committed output, payload registry, and
   bounded close semantics.
6. Add/extend metrics:
   - `monitor.file_save.snapshot_us` and `ui_return_us`;
   - snapshot lock-hold time;
   - retained bytes/lines at capture;
   - shared block count/bytes and active-tail copied bytes;
   - peak additional snapshot bytes;
   - worker total/write throughput and close/cancel latency.
7. Run the existing Monitor scrollback and ETW burst gates because changing text
   storage can regress indexing, layout, ingestion, and append-to-visible latency.
8. Update `Specs/Core/Core_RedSalamanderMonitor.md`: retain exact-start semantics,
   replace the deep-copy implication with immutable block/slice ownership, state
   the bounded active-tail rule, and document the new metrics and evidence.

### Performance proof

Create an archived same-machine before/after Release run under
`Specs/TestRuns/<machine>/Monitor/` with at least small, medium, and large retained
datasets. Five repeats per size are required. Report p50/p95/max snapshot lock and
UI-return latency, copied tail bytes, shared block count, additional snapshot
memory, save throughput, append-to-visible latency, and ETW batch drain latency.

The primary scaling gate is architectural: snapshot copied text bytes must be
bounded by the configured active block/tail, not total retained bytes. UI capture
may scale with the number of shared block handles but must not scale with total
string bytes. Establish numeric latency thresholds from the reviewed same-machine
baseline and record them in the authoritative spec/evidence; do not invent a
cross-machine microsecond limit in advance.

### Done criteria

- Save As exports the exact starting snapshot while concurrent append, retention,
  clear, and replacement proceed safely.
- UI-thread copied text is bounded by one documented active block/tail, not the
  complete retained payload.
- Snapshot lifetime does not double retained text, leak blocks after completion,
  or block ingestion for worker I/O duration.
- Existing Monitor interaction, search, filtering, selection, scrolling, ETW
  batching, cancellation, and teardown tests remain green.
- Same-machine Release evidence demonstrates the intended scaling and no material
  regression in append-to-visible or ETW batch-drain metrics.

## Finding L4-MON-02 -- thread-start failure escapes the Win32 command path

### Evidence and repair

- `DoFileOpen` publishes `g_monitorFileOpenActive=true` and then constructs a
  `std::jthread` at `RedSalamanderMonitor.cpp:2666-2667`.
- `DoFileSaveAs` publishes `g_monitorFileSaveState` and constructs a
  `std::thread` at lines 2888-2890.
- These functions are called directly from the main command/window dispatch
  around lines 4416-4417 and 4671-4700. A thread constructor can throw
  `std::system_error`; there is no named exception boundary, so it can escape a
  Win32 callback and terminate the process. State was already marked active.

1. Isolate each throwing constructor in a small Monitor-local `noexcept` start
   helper. Catch only `const std::system_error&`, log once, map to an appropriate
   HRESULT, and leave `std::bad_alloc` fatal. Do not add `catch (...)`.
2. Publish/rollback active state transactionally: failed open clears the active
   flag; failed save resets the shared state/thread owner; no completion is
   posted; the generation may remain monotonically consumed.
3. Show one localized launch failure and leave the document/destination
   unchanged. The failed Save snapshot must release normally on the UI thread.
4. Add fail-next-open-thread and fail-next-save-thread seams through the real
   command path. Assert no process termination, stuck busy state, joinable worker,
   payload token, file mutation, or blocked subsequent retry.
5. Reconcile the open/save lifecycle wrappers only where stop/detach semantics
   actually match; do not force `jthread` and detached shared-state workers into a
   misleading generic Common abstraction.

## Finding L4-MON-03 -- canonical open/save adds blank records

`SnapshotBuilder` publishes on every LF at
`MonitorFileReader.cpp:75`, then `Finish` unconditionally publishes the current
line at lines 98-107. Canonical exporter output always appends LF after each
snapshot line at `MonitorFileExporter.cpp:141`. Therefore exported
`BOM + "a\n"` reopens as `{"a", ""}` and re-exports as `"a\n\n"`; each cycle
adds another blank record. The retained-state Chrome check accepts
`lineCount >= inputCount` and never compares exact output bytes, masking the bug.

Keep the existing monitor-record policy of one final LF per exported record, but
make reader finalization delimiter-aware: an LF publishes that record; EOF
publishes a remaining unterminated record only when one exists. EOF must not add
a synthetic record after a terminal delimiter. Define empty/BOM-only input as
zero records and ensure the exact line budget does not charge an extra terminal
empty line.

Add exact read/export/read byte-and-record matrices for empty/BOM-only, one
unterminated record, one canonical terminated record, CRLF normalization,
consecutive blank records, multiple trailing blank records, Unicode, and exact
line-budget boundary/+1. Tighten the Chrome retained-state drill to exact line
count and canonical output, not `>=`/existence. Update
`Specs/Core/Core_RedSalamanderMonitor.md` with the record/final-terminator policy.

## Follow-up re-audit expansion (2026-08-14)

This follow-up reviews the complete committed interval after the original
baseline, not only the lines changed in each commit:

- 1ab7a3b40 (Terminal pointer-follow), a4faca9aa (Operation Chordforge),
  7ca29d34a (File Operations popup), bae5f9546 (tooling governance/build
  coordination), and b11d54a6c (Operation Startrail plan), including their
  merge/refactor context.
- The surrounding callers, shared helpers, schemas, normative specs, tests,
  workflows, packaging scripts, and the then in-progress Startrail/DxUi worktree.
- 266 committed changed files (36,710 additions, 7,962 deletions) plus the
  pre-plan-edit worktree snapshot (67 tracked changed files, 4,282 additions,
  577 deletions, and 22 untracked paths).

The 2026-08-15 extension then reviewed the materialized workspace and its sibling
UI repair through merge `e9c76e59a`: 142 files changed from `b11d54a6c`, with
12,363 additions and 1,851 deletions. Line anchors below describe that committed
HEAD and must be refreshed if implementation drifts. A green source-contract
test is not treated as proof where the test encodes the defective shape.

### Post-commit extension verdict (2026-08-15)

- `d24c2bde0` committed the previously reviewed Startrail/DxUi workspace and
  moved the Startrail plan to `Done`; it did not repair L4-BUILD-01..08,
  L4-STAR-01..11, or L4-DXUI-01/02. The Done plan is historical implementation
  evidence, while this active plan supersedes its closeout assertions for the
  unresolved defects below.
- `a210613a2` is orthogonal to the existing Terminal, accessibility-root,
  WindowHost-shutdown, graph-bound, aggregate-wrap, and localization findings;
  those findings remain. It adds retained-tooltip, graph-color, delete-focus,
  and DPI-relayout defects documented as L4-DXUI-03/04 and
  L4-FILEOPS-13/14/15.
- Merge `e9c76e59a` contains no additional semantic conflict resolution. The
  reviewed tree was clean, and `git diff --check b11d54a6c..e9c76e59a` passed.
- The post-commit tooling checks remained superficially green: focused
  BuildEvidence/ValidationFingerprint/ValidationEvidence/ValidationImpact/
  RunAllTestsPlan Pester passed 93/93; tool inventory classified 172/172; spec
  inventory reported zero blockers. The new findings explain why those green
  checks are insufficient, including an impact rule and inventory validator
  that both encode or miss the defective closure.

### Follow-up validation baseline

- Tools/Get-ToolInventory.ps1 -FailOnFindings: 172 discovered/classified files,
  zero findings.
- Tools/Get-SpecInventory.ps1 -FailOnFindings: zero blocking findings.
- Focused Startrail/runner Pester: 93 passed, zero failed.
- Tooling-governance/migration Pester: 121 passed, zero failed.
- Additional focused BuildEvidence, ValidationEvidence,
  ValidationFingerprint, ValidationImpact, RunAllTestsPlan, and
  ToolingGovernance batches were green.
- Run-AllTests -Suite Full -ExplainPlan -ImpactBase 83ab72733... planned all 18
  Full entries as affected; plan digest
  9806dc1098d8ce234af929877a47524142802d3ad494b33790fe5e9927e96a82.
- A read-only receipt probe changed receipt_id to 64 zeroes and was still
  accepted. A second probe reduced the 319-output receipt to only
  RedSalamander.exe, retained the false ID, and was also accepted.
- git diff --check found no whitespace error (only existing line-ending
  conversion warnings).
- No Fresh Full aggregate was run for this review. The focused green tests do
  not cover the destructive reparse, process-interruption, malformed-result,
  cross-profile, high-cardinality render, or cross-thread teardown cases below.

### Follow-up finding inventory

| ID | Priority | Kind | Finding | Confidence |
|---|---:|---|---|---:|
| L4-BUILD-01 | P1 | Receipt truth | tests_enabled is inferred from configuration instead of the effective RSBuildEnableTests property, lying in both Debug-production and Release-test directions. | High |
| L4-BUILD-02 | P1 | Attestation | Receipt reads do not verify receipt_id, the immutable record, output-ID uniqueness, runtime-closure semantics, or complete Full artifacts. | High |
| L4-BUILD-03 | P1 | Attestation | Rebuild recursively blesses every surviving output file, so stale/removed binaries can bootstrap into a new trusted receipt. | High |
| L4-BUILD-04 | P1 | Packaging | MSIX version stamping mutates tracked source and normally makes the final receipt barrier reject the completed package. | High |
| L4-BUILD-05 | P1 | Profile isolation | Full's mandatory deployment writer always deletes/rebuilds x64 Debug even for Release or ARM64 validation. | High |
| L4-BUILD-06 | P1 | Supply chain | Portable producer attestation hashes live output without proving it is the byte set named by the receipt. | High |
| L4-BUILD-07 | P2 | Supply chain | Receipt toolchain identity does not bind the actual compiler/linker/resource-compiler/SDK byte identity claimed by workflow evidence. | High |
| L4-BUILD-08 | P1 | Filesystem safety | ZIP publication can follow a repository junction and force-overwrite an external path. | High |
| L4-STAR-01 | P1 | Source identity | Renames across an excluded-path boundary can erase the included production-side deletion/addition from the workspace snapshot. | High |
| L4-STAR-02 | P1 | Filesystem safety | Evidence pruning can recurse through an ancestor junction and delete external data. | High |
| L4-STAR-03 | P1 | Verdict integrity | Missing, malformed, or structurally incomplete selftest JSON can exit green and, when a file exists, be promoted as reusable evidence. | High |
| L4-STAR-04 | P1 | Verdict integrity | The v2 summary can turn schema-readable or even missing promotion data into repository PASSED without semantic/digest verification. | High |
| L4-STAR-05 | P1 | CI contract | The documented 0..5 exit taxonomy is reduced to 0/1, allowing NOT_EVALUATED/Explain outcomes to look like a repository gate pass. | High |
| L4-STAR-06 | P2 | Crash consistency | Artifact-writer receipt invalidation metadata and the read barrier are unused during the real writer launch. | High |
| L4-STAR-07 | P2 | Immutability | Re-attestation replaces an already-published decision record that the spec declares immutable. | High |
| L4-STAR-08 | P2 | Evidence integrity | Result artifacts are path-only, mutable, copied non-atomically, and limited per file rather than per entry/run. | High |
| L4-STAR-09 | P2 | Parser contract | Read-side validation accepts noncanonical JSON and omits promised node/string resource bounds. | High |
| L4-STAR-10 | P2 | Confinement | Resume accepts nested or ancestor-junction origins and does not bind the directory leaf to state.run_id. | High |
| L4-STAR-11 | P3 | Simplification/architecture | expected_build_action is computed after the unconditional build and is display-only, so Affected cannot honor none/targeted/full planning. | High |
| L4-STAR-12 | P2 | Impact correctness | The Specs/** catch-all classifies executable schemas, themes, budgets, and machine data as prose with no affected tests or build work. | High |
| L4-STAR-13 | P2 | Record integrity | Resume persists a Fresh validation plan even though run-state says Resume, while the supposedly shared plan digest includes that mode. | High |
| L4-STAR-14 | P3 | Diagnostics | Aggregate fingerprint mismatch returns SOURCE_INPUT_CHANGED before component comparison, making specific invalidation reasons unreachable. | High |
| L4-STAR-15 | P3 | Governance/simplification | Tool consumer/import closure and runner dependency lists are hand-maintained, incomplete, and not reconciled by inventory validation. | High |
| L4-STAR-16 | P2 | Specification ownership | Startrail is marked Done with every gate passed while this active plan still owns live release-blocking Startrail findings. | High |
| L4-CHORD-01 | P1 | ABI boundary | ITerminalActions::GetActionState writes a full output record without validating the caller's sizeBytes. | High |
| L4-CHORD-02 | P2 | Command architecture | Command Palette discards CommandStateSource, presents disabled/stale actions as enabled, and has incorrect registry state ownership for two commands. | High |
| L4-CHORD-03 | P2 | Shortcut correctness | Preferences can persist Application/Function Bar physical-key bindings that runtime lookup can never dispatch. | High |
| L4-CHORD-04 | P2 | Settings validation | Terminal-only cmd/shortcut/passthrough is accepted in every shortcut scope with inconsistent runtime effects. | High |
| L4-CHORD-05 | P2 | Dispatch truth | Numbered Preview/Terminal tab commands report handled when their target is unavailable. | High |
| L4-CHORD-06 | P2 | Callback contract | SetCallback accepts non-null callback/null cookie although the pair must be both null or both non-null. | High |
| L4-CHORD-07 | P2 | Unicode correctness | Suggestion deduplication combines locale-sensitive towupper hashing with ordinal-ignore-case equality. | High |
| L4-CHORD-08 | P3 | UI truth | Palette shortcut chips include global aliases overridden by the active context. | High |
| L4-CHORD-09 | P3 | Architecture | Floating Terminal declares two WM_APP messages outside the central registry. | High |
| L4-CHORD-10 | P3 | Policy/spec | Same-IID slot insertion conflicts with the general plugin ABI rule, but is not a current defect under confirmed atomic lockstep build/staging. | High |
| L4-DXUI-01 | P2 | Accessibility | The new root cache is synchronized, but several root-returning/event paths still create distinct COM identities for one HWND. | High |
| L4-FILEOPS-09 | P2 | UI correctness | Hosted throughput graphs enable hue bands for one stream in the default theme, contrary to spec and fallback behavior. | High |
| L4-FILEOPS-10 | P2 | Performance/architecture | The 100-ms graph path groups float hues quadratically and can create 2,864 D2D geometries per frame. | High |
| L4-DXUI-02 | P2 | Threading | Process-exit shutdown may detach a live secondary-thread WindowHost and TSF state from the wrong thread. | High |
| L4-FILEOPS-11 | P3 | Numeric correctness | Global progress totals use unchecked uint64_t addition and can wrap backwards. | High |
| L4-FILEOPS-12 | P3 | Localization/simplification | New count clauses produce singular grammar such as “1 need attention,” while tests bless the resource and an older combined format remains dead. | High |
| L4-DXUI-03 | P2 | Pointer/tooltip correctness | Passive badge tooltip controls cannot win hit testing; the shared route also bypasses the specified show delay and display lifetime. | High |
| L4-DXUI-04 | P2 | DPI/layout correctness | Menu DPI relayout clamps origin but not oversized width/height, so content can extend beyond a smaller monitor work area. | High |
| L4-FILEOPS-13 | P2 | Visual contract | Production row and graph colors diverge by theme/path, while the new helper-only parity test bypasses those real gates and fallback. | High |
| L4-FILEOPS-14 | P2 | Result/focus correctness | Delete-focus tracking uses aggregate HRESULT and pre-excludes every requested path instead of resolving exact per-source removals with provider path identity. | High |
| L4-FILEOPS-15 | P2 | Interaction correctness | A delayed/stale/coalesced delete refresh can steal newer user focus and consumes only the first pending removal intent. | High |

### Additional non-negotiable invariants

13. **Receipts are verified content-addressed claims, not schema-readable hints.**
    Every consumer recomputes receipt identity, resolves the matching immutable
    record, validates unique outputs and exact closure, and refuses unmanifested
    or stale files.
14. **Build profile and test surface are explicit identity.** Platform,
    configuration, effective test-hook property, writer mutations, and
    re-attestation all name the same profile. Configuration name alone never
    implies test availability.
15. **Evidence cannot manufacture a verdict.** Only normalized, schema- and
    outcome-validated terminal results can promote. Missing/malformed/partial
    output, corrupt provenance, or evidence-publication failure produces the
    documented nonzero category.
16. **Repository-owned paths are resolved component by component.** Recursive
    delete or force-overwrite requires a direct governed child with no reparse
    component from the trusted root to the target.
17. **Command state has one resolver.** Menus, palette, shortcuts, plugin
    actions, and tab availability consume the same typed state/ownership truth
    and revalidate session identity immediately before mutation.
18. **One accessible HWND has one root provider identity.** All child/root/event
    paths retain or query the target-owned canonical root, and teardown
    disconnects that exact identity.
19. **WindowHost teardown is owner-thread work.** Global process shutdown
    orchestrates/join-drains UI owners before resetting shared graphics/TSF
    resources; it never tears down a foreign live host directly.
20. **Graph work is bounded by governed slots, not historical identities.**
    Rendering is linear in bounded samples and palette slots, and default
    single-stream rendering remains the classic accent path.
21. **Executable specification data is an input, not documentation.** Schemas,
    themes, budgets, corpora, normative machine maps, and other consumed data
    select every dependent validation/build surface before the prose catch-all.
22. **Validation records tell one mode/identity story.** A Resume run cannot
    persist a plan that calls itself Fresh. Reusable coverage identity is
    explicitly separated from invocation mode and cross-record checks compare
    the correct identity.
23. **Supplemental tooltips are pointer-discoverable without becoming actions.**
    Passive status regions participate in a separate deepest-hover query, keep
    click-through semantics, use the governed show delay, auto-hide lifetime,
    and survive tree rebuild only through lifetime-validated references.
24. **Delete focus follows observed removals and user intent.** Exact per-source
    outcomes plus provider path identity determine survivors; a later user focus
    or navigation epoch wins, stale snapshots retain the intent until proof or
    expiry, and one accepted refresh resolves all applicable completions.
25. **Popup containment is re-established after every DPI transition.** Initial
    placement and relayout share one work-area size-and-position resolver, and
    the final reachable viewport/scroll range derives from the clamped surface.
26. **File-operation color identity is end to end.** When concurrent bands are
    active, a row and its hosted/fallback graph band resolve the same stable
    stream hue through one color conversion in every non-high-contrast theme.

## Follow-up build and packaging findings

### Finding L4-BUILD-01 -- receipt test-surface identity is inferred, not measured

build.ps1:732 sets receiptTestsEnabled to Configuration != Release.
Directory.Build.props:69-70 and 160-181 instead compile ENABLE_TESTS from the
effective RSBuildEnableTests property; the Full environment forces that
property true in Tools/Modules/Testing/TestInvocation.psm1:169-180. Therefore:

- an ordinary Debug build can omit project-specific test hooks while claiming
  tests_enabled=true and can be accepted by Full -SkipBuild;
- a valid test-enabled Release build claims false and is rejected initially and
  after the mandatory writer re-attestation; and
- build_arguments do not bind the effective property, so the two binary shapes
  can otherwise share receipt compatibility.

Repair by resolving the effective property once, passing it explicitly through
MSBuild, receipt projection, compatibility, re-attestation, and portable
attestation. Full requires every executable/plugin in its plan to be built with
the required hooks; a partial project exception makes the whole Full receipt
test-disabled. Add ordinary/test-enabled Debug and production/test-enabled
Release matrices, including Full -SkipBuild and post-writer re-attestation.

### Finding L4-BUILD-02 -- receipt pointer identity and Full closure are not verified

Tools/Modules/Build/BuildEvidence.psm1:440-450 schema-reads the mutable pointer
but does not recompute receipt_id or require equality with
receipts/<receipt_id>.json. Test-RSBuildReceiptForArtifactUse at 499-529 hashes
only outputs the pointer chooses to list and checks required paths supplied by
the caller; Run-AllTests currently requires only RedSalamander.exe. It does not
validate unique artifact IDs, runtime-closure members, closure completeness, or
the full executable/plugin plan.

The review reproduced both defects without editing source: a 64-zero receipt ID
was accepted, and a receipt truncated from 319 outputs to RedSalamander.exe with
the false ID was also accepted.

Create one semantic receipt loader used by build, runner, workflow, packaging,
and Resume. It must recompute the projection digest, load and byte/semantic-match
the immutable content-addressed record, require case-insensitive unique IDs and
paths, validate every closure reference, and prove the complete selected plan
closure before artifact use. Tests mutate every projection field, pointer/record
pairing, IDs/paths, closure entries, and required Full output.

### Finding L4-BUILD-03 -- stale files bootstrap into new receipts

Get-RSBuildArtifactFiles at BuildEvidence.psm1:156-187 recursively selects every
surviving file in the output root (apart from receipt files and a narrow
project-build exclusion). Get-RSBuildArtifactRecords at 189-208 blesses all of
them and gives every executable the entire scanned output as its runtime
closure. A Rebuild does not guarantee deletion of outputs from removed/renamed
projects or arbitrary injected files, so those bytes can become newly attested
and later executed by a fixed plan path. This contradicts the Build_Toolchain
claim that pre-existing outputs cannot bootstrap themselves.

Derive outputs and per-launchable closure from the governed project/runtime
manifest, or clean/quarantine the exact profile root before the first complete
attestation and reject extras afterward. Add tests that pre-place an unknown
EXE/DLL and a removed-project binary; the build must remove or reject them and
the runner must never launch them.

### Finding L4-BUILD-04 -- MSIX stamping invalidates every ordinary receipt

build.ps1 captures the source snapshot at 734, invokes
Installer/msix/UpdateManifestVersion.ps1 at 1004-1009, and the stamper rewrites
the tracked Package.appxmanifest at 101-114. build.ps1:1220-1223 then compares
the final tree to the pre-operation snapshot and refuses the receipt whenever
the generated version/architecture differs from the committed manifest.

Generate/stage a temporary manifest into the packaging project, or preserve the
exact original bytes and restore them in a finally block before the final
barrier. Restoration failure marks contamination and publishes no receipt.
Exercise a nonmatching version success, injected package failure, and
interruption; success restores a clean source snapshot and publishes a receipt,
while failure leaves either the exact original or an explicit recoverable
contamination record.

### Finding L4-BUILD-05 -- Full's writer mutates the wrong build profile

Tools/Tests/RedSalamanderPluginDeployment.Tests.ps1:7-10 and 84-100 hardcode
.build/x64/Debug and invoke a Debug x64 build. TestSuitePlan.psm1:214-226 inserts
that Pester writer first for every Full plan, while Run-AllTests:1615-1655 holds
and re-attests the user-selected platform/configuration.

Parameterize the writer through the plan contract and isolated environment so
it deletes/rebuilds/re-attests only the selected profile; if a profile is not
supported, reject it before planning. Add fake and real contract tests for x64
Release and ARM64 Debug/Release proving no Debug sibling mutation and no
cross-profile nested lock ordering.

### Finding L4-BUILD-06 -- portable attestation is not derived from the receipt

The reusable workflow reads a receipt, then scans the live output at
.github/workflows/build-reusable.yml:451-474. New-RSPortableBuildAttestation at
BuildEvidence.psm1:571-627 accepts those arbitrary artifact records and copies
source/toolchain identity from the old receipt without comparing payload bytes
or membership. A post-receipt executable mutation can therefore receive a new
portable digest under stale provenance. The workflow's later
toolchain-identity.json also demonstrates the need for an explicit auxiliary
file contract.

Construct the payload from a freshly byte-verified receipt output set plus a
versioned allowlist/manifest of producer-added metadata, or issue a new receipt
after every addition. Refuse changed receipt bytes, missing closure members, and
unreviewed extras. Cover all three plus the allowed identity record.

### Finding L4-BUILD-07 -- receipt toolchain identity is weaker than claimed

Get-RSBuildIdentityBundle at BuildEvidence.psm1:288-320 hashes MSBuild but uses
possibly empty VCToolsVersion/WindowsSDKVersion environment text and does not
bind actual cl, link, rc, or SDK bytes. The workflow creates a stronger
toolchain-identity.json at build-reusable.yml:397-449 only after the receipt and
does not include it in receipt toolchain_identity_digest.

Resolve the effective toolset through MSBuild before compilation, emit one
canonical governed toolchain record, and use its digest in local receipt,
portable attestation, cache key, and workflow output. Tests must invalidate
identity when compiler/linker/RC/SDK seams change, even if environment labels do
not.

### Finding L4-BUILD-08 -- ZIP publication follows repository junctions

Installer/zip/build-zip.ps1:46-48 and 269-274 constructs
.build/AppPackages, creates it with Force, then Compress-Archive -Force writes
the ZIP. It does not use the guarded repository path resolver already required
by MSI/Winget. A pre-existing .build or AppPackages junction can redirect the
overwrite to an external location.

Use guarded component-by-component creation, reject reparse ancestors, write a
unique same-directory temporary archive, revalidate immediately before atomic
replacement, and preserve the previous complete ZIP on failure. Tests place
junctions at both levels and prove an external sentinel is unchanged.

## Follow-up Startrail and runner findings

### Finding L4-STAR-01 -- excluded-boundary renames disappear from source identity

ValidationFingerprint.psm1:613-624 records a rename only if both old and new
paths pass Test-RSValidationWorkspacePathIncluded. The exclusions at 396-405
include .build, Specs/TestRuns, and .log. Renaming an included production
x.h to x.log records neither the deletion nor the addition; BuildEvidence then
uses that snapshot for receipt identity and Run-AllTests can accept stale
artifacts.

Classify the two sides independently. An included old side always contributes a
deletion; an included new side always contributes an addition; rename linkage
is supplemental only when both are included. Cover staged and unstaged
included-to-.log, included-to-.build, and excluded-to-included changes and
verify source snapshot, impact widening, and -SkipBuild rejection.

### Finding L4-STAR-02 -- pruning can delete through an ancestor junction

Remove-RSValidationEvidenceRuns at ValidationEvidence.psm1:479-492 checks a
string prefix and only the final target's reparse bit before Remove-Item
-Recurse. A path such as runs/link/<valid-run-id>, where link is an external
junction and the final directory is ordinary, passes and recursively deletes
external data.

Require the canonical parent to be exactly runs, verify every component from
the trusted evidence root is non-reparse, bind apply to the exact previously
planned file identity/token, and refuse drift. A destructive-safety test creates
an external sentinel behind an ancestor junction and proves prune fails without
changing it.

### Finding L4-STAR-03 -- malformed/missing terminal results can false-green

Run-AllTests.ps1:1731-1778 warns on JSON parse failure and leaves a zero-failure
suite initialized from the process exit. Coverage runs only when parsed JSON
contains a cases array at 1780-1854. It then publishes evidence at 1883-1885.
Test-RSValidationResultPassed checks only exit code, Failed, and classification.
Consequences:

- exit 0 plus malformed JSON, or valid JSON with no matching cases/suite, can
  promote because the artifact exists and Failed remains zero;
- the aggregate-suite loop at Run-AllTests.ps1:1801 accepts the first suite
  whose name matches **or merely has a cases property**, so a preceding wrong
  suite can be consumed as the requested result; and
- Pester handling at Run-AllTests.ps1:545-547 derives success only from
  FailedCount, so zero discovered tests can pass and promote;
- exit 0 plus no result file makes evidence fail, but
  Complete-RSValidationEvidenceRun's failed status at 1934-1939 is not copied
  into overallExitCode, so the CLI may still return zero; and
- skipped/partial/unexpected outcomes are not evaluated against outcome_contract.

Make a normalized schema-validated terminal-result object mandatory before
promotion for every entry kind. Parse/discovery, coverage, outcome, allowed-skip,
and required-artifact verdicts must feed both suite and process exit. Evidence
publication takes an explicit validated verdict rather than inferring from three
loose fields. Test missing, malformed, no-cases, wrong-suite-first, zero Pester
discovery, all-skipped, unallowed-skip, count mismatch, and aggregate mismatch.

### Finding L4-STAR-04 -- aggregation trusts state labels instead of evidence

New-RSValidationRunSummaryV2 at TestRunSummary.psm1:229-305 schema-reads
promotion/evidence but never recomputes their digests or verifies duplicated
origin, entry, fingerprint, snapshot, receipt, epoch, coverage, outcome, and
artifact bindings. It even maps state status promoted directly to promoted when
promotion_digest is null, allowing a schema-valid state edit to contribute to
repository PASSED. Resume has stronger but still incomplete checks.

Build one verified promotion/evidence loader for Resume, summary, archive,
duration reporting, and pruning provenance. Require digest filename equality,
all duplicated fields to agree, complete artifact set equality (not subset),
content-bound result artifacts, directory/run identity, and acyclic provenance.
Cross-record plan checks must compare the mode-independent coverage-plan
identity introduced by L4-STAR-13 rather than equating a Fresh invocation record
with a Resume invocation record.
Null promotion, digest tamper, cross-entry/cross-run record, coverage/outcome
mismatch, and partial artifact map must yield BLOCKED/corrupt-evidence exit.

### Finding L4-STAR-05 -- the runner does not implement its exit taxonomy

Testing_ValidationEvidence.md:58-64 and 220-225 assigns 0 PASSED, 1 FAILED,
2 NOT_EVALUATED, 3 invalid CLI/plan, 4 blocked attestation, and 5 corrupt
evidence. Run-AllTests Explain exits zero, Affected prints NOT_EVALUATED but the
final path emits only overallExitCode 0/1, and thrown categories collapse to the
host default.

Introduce typed terminal states and one finalizer that maps them once after
summary publication while preserving pre-binding CLI failures. Subprocess tests
cover every category and CI/release policies must reject 2 as a final repository
gate even though it remains a valid scoped diagnostic result.

### Finding L4-STAR-06 -- writer invalidation has no crash-safe production path

TestRunPlan.Common.psm1 declares artifact writes and receipt invalidation.
Run-AllTests launches the writer at 1615-1620 without suspending the pointer or
calling the existing artifact read barrier. The deployment test deletes outputs
at RedSalamanderPluginDeployment.Tests.ps1:154-167; its child build suspends the
receipt only later. A kill in that interval leaves a current pointer describing
removed bytes.

Before writer launch, atomically transition the selected profile/epoch to
mutating and suspend its receipt; block all readers; on completion re-attest and
publish the new epoch in one ordered transition. Crash injection after deletion
must leave no usable receipt and later continuation must repair deterministically.

### Finding L4-STAR-07 -- an immutable decision is replaced in place

New-RSValidationEvidenceRun publishes the initial decision. After the writer,
Run-AllTests:1661-1680 writes a recomputed decision to the same path through
Write-RSAtomicCanonicalJsonFile, whose existing-file path deliberately replaces.
Testing_ValidationEvidence.md says every record except run-state is immutable.
A crash can also split the overwritten decision from the state that references
it.

Separate create-once immutable publication from mutable state. Store each
epoch's decision content-addressed and update only a run-state pointer after the
record is durable. A second different write to one immutable path must fail;
crash points before/after pointer switch remain reconstructible.

### Finding L4-STAR-08 -- result artifacts are mutable and effectively unbounded

Copy-RSValidationResultArtifacts at Run-AllTests.ps1:272-294 checks 64 MiB per
file, not per entry, has no one-GiB run accumulator, copies non-atomically, and
does not recheck/hash the destination. ValidationEntryEvidence stores only path
strings and Resume checks existence. Up to 4,096 files can bypass the intended
entry budget, while later modification/truncation remains reusable.

Publish immutable records of path, size, and SHA-256 after atomic copy/flush and
reopen verification. Enforce accumulated entry and run quotas and reject
reparse/nonregular inputs. Resume and aggregation rehash. Test two 40-MiB files,
run-budget crossing, source mutation during copy, destination truncation, and
post-promotion modification.

### Finding L4-STAR-09 -- read-side canonical and resource contracts are absent

Read-RSValidationJsonFile at ValidationFingerprint.psm1:877-907 checks 16 MiB,
BOM/LF, duplicate keys, schema, and depth, but accepts pretty/reordered/
overescaped/NFD bytes and never enforces 100,000 total nodes or one-MiB strings.
The normative spec requires exact canonical persisted records and these bounds.

Use a bounded streaming parse, enforce node/string/path/depth limits before
materialization, canonicalize with the shared writer, and require the original
bytes to equal canonical semantic bytes plus one LF. Cover whitespace, key
order, alternate escapes, NFD, normalized-key collision, excessive nodes, and
oversized strings. Migrate and canonically rewrite checked-in
Tools/validation-impact.json before enabling the strict reader so the repair
does not strand the runner on its current noncanonical input.

### Finding L4-STAR-10 -- Resume origin confinement is only a prefix check

Run-AllTests.ps1:1332-1350 accepts any path whose string starts with runs plus a
separator and checks only the final directory's reparse bit. It does not require
a direct child, verify directory name equals state.run_id, or reject an ancestor
junction. Get-RSValidationResumeCandidates likewise trusts the supplied root.

Use the same direct-child, component-safe resolver as prune and duration reads;
require leaf/state ID equality and bind plan/decision/state origins. Reject
nested ordinary directories, junction-backed runs, copied/renamed runs, and
case/path aliases before reading evidence.

### Finding L4-STAR-11 -- expected build action is dead planning metadata

Run-AllTests performs its build at 1173-1205, but constructs/publishes the impact
decision only around 1460. ValidationImpact computes expected_build_action as
none, targeted, full, or blocked; the runner only prints it. Affected therefore
cannot skip or target build work even when the plan says none/targeted, and the
field reads like an enforced decision despite being advisory.

Either move snapshot/impact planning before artifact mutation and execute a
governed build dispatcher that honors the selected action, or rename/version the
field as advisory until that phase exists. Centralize build selection so
Explain, Affected, Fresh, and Resume cannot describe a different action from the
one performed. Subprocess tests cover none, targeted, full, and blocked with
spy build seams and assert no pre-plan build occurs.

### Finding L4-STAR-12 -- executable Specs inputs are classified as prose

Tools/validation-impact.json:31 matches every Specs/** path as
documentation.specs with no tags and build_requirement none. ValidationImpact's
first matching explicit rule suppresses the conservative fallback. Direct
reproduction shows BuildReceipt/Validation schemas, SettingsStore.schema.json,
theme JSON5, performance budgets/corpora, and normative machine maps selecting
no dependent runtime or tooling entry even though code, tests, packaging, or
workflows consume them. This contradicts Testing_ToolingGovernance.md:263-265.

Add ordered machine-contract, runtime-data, performance-data, test-corpus, and
governance-map rules before the prose-only fallback. Derive referenced
executable Spec inputs from consumer manifests/imports where practical, and
make an unclassified consumed Spec path widen rather than disappear. Mutation
tests cover receipt/evidence schemas, settings schema, themes, budgets, corpora,
NormativeConsistency.json, and a genuinely prose-only Markdown file.

### Finding L4-STAR-13 -- Resume stores a Fresh plan with a Resume state

Run-AllTests.ps1:1264-1265 always creates New-RSValidationPlan with mode Fresh.
New-RSValidationEvidenceRun publishes that plan but writes the supplied Resume
mode into run-state at ValidationEvidence.psm1:72-98. The plan digest itself
includes validation_mode at TestRunPlan.Common.psm1:286 and 297-303. The checked-
in final campaign consequently has a Resume state beside a plan that calls
itself Fresh while both invocations claim one digest.

Split stable reusable coverage-plan identity from invocation/execution mode.
Persist a truthful Fresh or Resume invocation record, reference the shared
coverage identity explicitly, and make semantic loaders compare the intended
identity rather than contradictory fields. Test Fresh-to-Resume reuse, direct
record inspection, digest stability, wrong-mode tamper, and cross-record mode
consistency.

### Finding L4-STAR-14 -- fingerprint diagnostics collapse before comparison

ValidationFingerprint defines specific reason codes for build, logical input,
shared contract, runner, environment, arguments, coverage, outcome policy,
capability, and epoch changes. Get-RSValidationResumeCandidates at
ValidationEvidence.psm1:300-301 compares only the aggregate fingerprint first
and returns SOURCE_INPUT_CHANGED before any component can be identified. Most
diagnostic reason codes are therefore unreachable for authentic drift.

Persist a versioned component-digest projection and compare components in one
documented precedence order before falling back to an aggregate/corruption
reason. Mutation tests change one component at a time and assert the exact
reason, including simultaneous changes and an unknown future component.

### Finding L4-STAR-15 -- dependency closure is incomplete and hand-maintained

Tools/tool-inventory.json omits direct BuildEvidence workflow/runner consumers
and ValidationFingerprint importers. Run-AllTests.ps1:1290-1310 also omits
TestRunSummary.psm1 and RunAllTestsSummary.schema.json from its explicit runner
dependency list although it consumes them around 1978-1984. The inventory still
reports 172/172 and zero findings. The complete workspace snapshot currently
prevents this metadata drift from becoming an independent trust bypass, but the
claimed dependency closure and future cache narrowing are unsafe.

Derive PowerShell imports and workflow/script consumers, reconcile them against
usedBy, and derive the runner dependency manifest from the actual imported
module/schema closure. Add omission/extra-edge tests and require every imported
behavioral dependency to affect the appropriate cache/fingerprint.

### Finding L4-STAR-16 -- Done status contradicts active remediation ownership

Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md
is now under Done and claims every gate passed, while all eight L4-BUILD and
eleven pre-existing L4-STAR findings remain and this plan previously said the
Startrail plan could not close. That leaves two active-looking authorities with
opposite completion truth.

Keep the Done file as immutable implementation history and explicitly mark this
active remediation plan as the superseding owner for L4-BUILD-01..08 and
L4-STAR-01..16, or reopen the operation and its index atomically. Update
cross-links and inventory tests so a Done plan cannot be the sole active owner
of unresolved findings.

## Follow-up Chordforge findings

### Finding L4-CHORD-01 -- GetActionState overwrites undersized output

TerminalActionState is a size-versioned public record at
Common/PlugInterfaces/Terminal.h:255-264. Terminal::GetActionState at
Plugins/Terminal/Terminal.cpp:7016-7050 checks only state != null, then assigns
the full result without reading state->sizeBytes. An undersized caller buffer is
overwritten, unlike GetViewState's explicit validation.

Capture and validate the caller size before any write; return E_INVALIDARG for
an undersized record, populate only the reviewed prefix, and preserve oversized
tail bytes. Add null, undersized canary, current, and oversized-tail plugin
contract tests.

### Finding L4-CHORD-02 -- palette ignores typed command state

CommandStateSource exists in CommandRegistry.h:25-31, but
ShortcutCommandCatalog drops it, PaletteRow has no enabled state, and
CommandPaletteWindow::Activate always closes and dispatches. Terminal Copy with
no selection, missing Preview/Terminal tabs, and a replaced/exited Terminal
therefore appear actionable and close on a no-op. contextMenu and close also
claim TerminalActions state in CommandRegistry.cpp:434-453 although the plugin
allowlist excludes them and FolderWindow handles them.

Add a shared typed state resolver used by palette/menu/shortcut surfaces. Correct
registry owner/state metadata, capture origin HWND plus Terminal
instance/session generation, render localized disabled state, and revalidate
immediately before activation. Disabled Enter/click remains inert and preserves
focus. Test no-selection Copy, missing tabs, host-owned close/context menu, and
terminal replacement while the palette is open.

### Finding L4-CHORD-03 -- persisted physical shortcuts can be dead

The settings schema, parser, and Preferences allow keyPosition in every scope,
and ShortcutManager loads it into every map. FindApplicationCommand and
FindFunctionBarCommand at ShortcutManager.cpp:156-175 form only VK keys; their
runtime callers do not pass scan/extended identity. An Application number-row
plus/minus binding saves, displays, and reloads but can never match.

Reuse the physical-aware FindCommand path and pass original scan/extended data
for Application. Either implement the same for Function Bar or reject physical
identity for that scope in UI, parser, and schema. Add per-scope round-trip plus
runtime dispatch under a non-US mapping.

### Finding L4-CHORD-04 -- pass-through validation is not scope aware

Core_SettingsStore states cmd/shortcut/passthrough is Terminal-only.
ParseShortcuts at Common/Common/SettingsStore.cpp:4666-4705 validates only the
cmd/ prefix through one parser for Application, Function Bar, Folder, and
Terminal. Today Application silently yields input while Folder/Function Bar may
show an unknown-command dialog.

Make scope part of semantic validation and Preferences import. Reject and
recover only the invalid section outside Terminal; preserve the sentinel in
Terminal. Add strict/recovery/import tests for all four scopes and align the
schema if its structure can express the rule.

### Finding L4-CHORD-05 -- numbered tab selection lies about handling

FolderWindow.Interaction.cpp:366-376 returns false from selectContent when
Preview or Terminal is unavailable, but the numbered command path at 397-405
discards that result and returns true. The key is consumed while selection and
focus do not change.

Return selectContent directly and expose that same availability through the
typed command-state resolver. Test identities 1-3 through direct dispatch,
shortcut, and palette in both available/unavailable states.

### Finding L4-CHORD-06 -- SetCallback accepts an invalid half-pair

Terminal_EmbeddedPlugin requires callback and cookie to be both non-null or both
null. Terminal::SetCallback at Terminal.cpp:7148-7160 rejects only null callback
plus non-null cookie. Non-null callback plus null cookie registers, but host
cookie validation drops all exit events, so tabs need not auto-close.

Use XOR nullness validation, preserve the synchronous clear barrier, and test
all four pair combinations plus an actual event delivery/clear-drain sequence.

### Finding L4-CHORD-07 -- suggestion hash and equality disagree for Unicode

TerminalCommandSurfaceModel.cpp:122-142 hashes each wchar_t through
locale-sensitive std::towupper but compares keys with CompareStringOrdinal
ignore-case; the pair is used by unordered_map at 319. Equal keys are required
to have equal hashes. Non-ASCII case variants can therefore occupy different
buckets and evade deduplication/frequency ranking.

Canonicalize once using one documented ordinal relation and use exact
hash/equality on the canonical value, or implement a hash with precisely the
same relation. Cover accented Latin, Greek, Cyrillic, locale changes, and mixed
current/persisted ranking.

### Finding L4-CHORD-08 -- shortcut chips show overridden aliases

ShortcutCommandCatalog.cpp:48-75 blindly appends Application and active-context
bindings without resolving runtime precedence. A Terminal pass-through or
different-command override for global Ctrl+Shift+P still leaves that chip on the
Command Palette row although Terminal-first lookup makes it ineffective.

Resolve the effective chord map for the active context first, then aggregate
aliases from that map. Cover pass-through, unassigned, different-command, and
same-command overrides plus duplicate aliases.

### Finding L4-CHORD-09 -- Floating Terminal bypasses message governance

FloatingTerminalWindow.cpp:37-38 defines WM_APP+392 and +393 locally despite the
central registry contract in Common/WindowMessages.h. No current collision was
found, so this is preventive architecture debt.

Move both constants into WndMsg, use the central names, and extend the
repository-wide uniqueness/source-contract test so future collisions are
visible.

### Finding L4-CHORD-10 -- clarify the lockstep same-IID policy

SetCallback was inserted before Close while retaining ITerminal's IID. That
conflicts with the general plugin-callback rule and the spec sentence that an
incompatible vtable change gets a new IID. The user explicitly confirms that
today's supported product always compiles and stages host plus Terminal.dll
atomically from this source tree, with no mixed-binary deployment. Therefore
this is not a current crash/compatibility defect and must not be presented as
one.

Choose and document one policy: explicitly waive the same-IID rule for this
closed lockstep interface and test atomic staging/complete receipt closure, or
use a queried new IID before any future independently deployable compatibility
promise. Do not add old/new fixture work unless that product boundary changes.

## Follow-up File Operations and DxUi findings

### Finding L4-DXUI-01 -- root-provider identity is not canonical everywhere

The new WindowHostAccessibilityTarget rootProvider cache is protected by
GetAccessibilityTargetMutex, normal WM_GETOBJECT reuse is synchronized, and
unregister detaches/disconnects safely. The suspected com_ptr data race/UAF is
therefore rejected. However root navigation at DxUi.Accessibility.cpp:7320-7327,
collapsed-root hit/focus at 5533-5568, the debug factory at 7820-7860, and
notification creation at 7863-7889 still construct fresh root providers. One
HWND can expose multiple IUnknown identities; events may be raised on a provider
other than the cached/disconnected root.

Create one target-owned lazy canonical-root factory under the target mutex and
QI/AddRef it for every root, child FragmentRoot, hit/focus, notification, and
debug path. Runtime tests compare IUnknown identity across all paths, then
detach/force HWND reuse to prove the old provider is unavailable and the new
HWND has a distinct identity.

### Finding L4-FILEOPS-09 -- hosted graph colors a default single stream

Both File Operations hosted call sites enable perStreamBands unconditionally.
DxUi.Controls.cpp:3411-3518 gates only on that flag and high contrast, so even a
one-weight history renders hue bands. The legacy fallback checks for at least
two streams and the authoritative specs require default-theme bands only once
history contains two concurrent streams.

Move the semantic gate into the shared control: Rainbow may color one stream;
otherwise per-stream bands require multi-stream history, and the fallback uses
the same helper. Add hosted-versus-fallback pixel tests for default one stream,
default two-plus streams, Rainbow one stream, and high contrast.

### Finding L4-FILEOPS-10 -- graph grouping is quadratic and geometry-heavy

History allows 180 samples by 16 weights. Each item transition receives a new
golden-angle float hue. Shared rendering linearly searches exact floats for each
band, emits up to 2,864 quads, then creates/fills one D2D path geometry per
historical hue on every 100-ms frame; worst-case grouping is roughly four
million comparisons. The fallback duplicates the algorithm, and the hosted bug
above activates it even for sequential history. Existing tests use only two
samples and no required performance archive landed.

Store a bounded palette/color-slot ID (at most the admitted stream slots), group
with fixed arrays in O(samples times slots), cap geometry count by palette size,
and share one renderer with the fallback. Remove the unreachable final-index
branch. Instrument samples, quads, unique geometries, and render microseconds;
run deterministic 180x16 churn five times in test-enabled Release and archive
before/after evidence under Specs/TestRuns.

### Finding L4-DXUI-02 -- process shutdown detaches foreign-thread hosts

WindowHost::Attach records owner thread. ShutdownAllWindowHostsForProcessExit at
DxUi.WindowHost.cpp:1187-1204 copies raw hosts and invokes DetachForProcessExit
on the caller. Detach tears down capture, TSF, retained tree, and devices;
native text input uses thread-affine ITfThreadMgr operations, while final
shutdown covers only the current thread. The app closes its known splash thread
first, but the shared API does not enforce that quiet point for another live
secondary UI host.

Provide owner-thread shutdown/marshal orchestration, join worker UI threads,
then make the final global sweep assert the registry is empty before global
resource reset. Never detach a live foreign host directly. Add a test that
leaves a worker-thread host live during process-exit orchestration and verifies
all TSF/host teardown runs on its attachment thread.

### Finding L4-FILEOPS-11 -- aggregate progress totals can wrap

BuildGlobalStatusSummary at FileOperations.Popup.cpp:1007-1015 adds provider
uint64_t byte/item totals with plain += and later derives progress/taskbar state
from the wrapped values. Multiple extreme provider totals can produce a smaller
or zero aggregate and regress progress.

Use saturating addition or a wider/scaled accumulator for completed and total
values, preserve completed <= total, and define the displayed saturated state.
Exercise UINT64_MAX + 1 and several near-maximum tasks through the existing
debug summary helper; add “aggregate never wraps” to the progress spec.

### Finding L4-FILEOPS-12 -- count grammar and a dead format should be simplified

The new attention-only clause formats English “1 need attention” and analogous
satellite plural risks. Code has one count form and the selftest compares output
to the same malformed resource, so it cannot detect grammar. The older combined
IDS_FMT_FILEOPS_GLOBAL_STATUS_SUMMARY remains unused in base and satellites.
The post-commit badge strings extend the same exposure: for example, French
uses plural adjective forms for every badge count, while parity tests validate
only IDs and placeholder tokens.

Prefer count-neutral localized labels such as “Needs attention: {0:L},” or add
the repository's reviewed plural mechanism. Test representative 1/2 values per
satellite independently of the production resource under test. Remove or
explicitly reserve the dead combined resource and update localization/source
contracts.

### Finding L4-DXUI-03 -- passive tooltip regions are unreachable

File Operations creates completed-group badge tooltip regions as DxUi::Label at
FolderWindow.FileOperations.Popup.cpp:5105-5112 and updates them at 5194-5202.
Label::HitTest unconditionally returns null at DxUi.Controls.cpp:1864-1872,
while WindowHost::UpdateHover dispatches OnMouseMove only to the normal
HitTestControl result at DxUi.WindowHost.cpp:4009-4086. The badge therefore
publishes accessibility HelpText but can never show the required pointer
tooltip. The new test calls region->OnMouseMove directly and bypasses this
failure.

The shared Control::OnMouseMove path also calls immediate cursor-following
SetTooltip at DxUi.cpp:490-502 even though the authoritative design requires a
500-ms show delay and 5,000-ms display lifetime; the existing delayed API is not
used, and TooltipLayer has no display-duration expiry while the pointer stays.

Add a separate deepest-visible supplemental-tooltip query so passive regions
can be hover targets without becoming activation targets or blocking the
underlying header. Route it through SetTooltipDelayed, lifetime-validate the
target across retained-tree rebuilds, and implement the specified visible
lifetime. Test real host WM_MOUSEMOVE before/after delay, stationary auto-hide,
leave, text/bounds change, rebuild during delay, click-through, and accessibility
parity. Direct virtual calls are not sufficient proof.

### Finding L4-DXUI-04 -- DPI relayout can escape the monitor work area

Initial menu placement at DxUi.Menu.cpp:2980-3002 caps desired width and height
to the selected monitor work area. WM_DPICHANGED relayout instead recomputes raw
content/minimum-root dimensions and calls ComputePopupSurfaceRectFromTopLeft at
3184-3205, which clamps only x/y and returns x+width/y+height without capping
the extents. A menu moved to a smaller or higher-scale monitor can leave its
right/bottom content unreachable; the new minimum invoker-width option can make
the width path fail independently.

Share one work-area extent-and-position resolver between initial creation and
DPI relayout, then derive viewport, scrollbar, backdrop, region, and stored DIP
dimensions from the clamped surface. Test oversized width and height, long
localized items, a minimum root width larger than the destination work area,
and monitor/DPI transition; assert the final rect is contained and the last
command remains scroll-reachable.

### Finding L4-FILEOPS-13 -- row/graph color parity is helper-only

UI_FileOperationsPopup.md requires each active row mini bar to use the exact
assigned color of its concurrent graph band through one conversion. Production
rows use assigned hue only in Rainbow mode and otherwise use the base accent at
FolderWindow.FileOperations.Popup.cpp:7162-7169. Hosted graphs render canonical
per-stream hues in ordinary themes, while fallback graph code at 4411-4425 and
4694 uses a different saturation/value conversion. The new parity test directly
compares RainbowProgressColor with ThroughputGraphColorFromHue and never
executes either production theme gate or the fallback renderer.

Resolve active-row color from the stable stream cookie/ID/source guard whenever
bands are actually active, keep the one-stream default accent required by
L4-FILEOPS-09, and make hosted/fallback share the canonical conversion and band
renderer. Add runtime pixels/state tests for default one stream, default
concurrent streams, Rainbow, high contrast, and forced hosted-control failure;
assert row-to-band identity rather than helper equality.

### Finding L4-FILEOPS-14 -- delete focus ignores exact item outcomes

All three File Operations completion paths pass SUCCEEDED(task HRESULT) to
CompleteRemovalFocusTracking at FolderWindow.FileOperations.cpp:523,1129,1188.
The engine already records sparse per-source statuses and publishes them at
1437-1444; exact removal requires per-source S_OK, while S_FALSE, skip,
cancellation, and partial copy are not proof. BeginRemovalFocusTracking also
pre-excludes every requested name before outcomes exist, and its case-insensitive
sets/survival tests ignore providers whose declared component identity is
ordinal case-sensitive.

For sorted A/B/C/D with focused B and request {B,C}, B may delete while C fails.
The aggregate ERROR_PARTIAL_COPY discards the token and leaves focus unset; a
naive aggregate-success repair would preselect D even though surviving C is the
nearest item. SUCCEEDED(S_FALSE) can commit a removal that never happened, and
case-only siblings can be conflated on S3/Curl/MTP-style identities.

Pass source paths and exact outcomes into focus completion, fail unknown
outcomes closed, and resolve the nearest actual survivor against the refreshed
list using the provider's path-identity profile. Test focused-success/other-
failure, focused-failure/other-success, S_FALSE, missing outcome, cancellation,
bulk partial, all-success, and case-only siblings.

### Finding L4-FILEOPS-15 -- delayed delete refresh can steal newer focus

PendingRemovalFocus stores token/folder/removed/replacement/expiry only; it has
no user-focus, sort, provider, path-visit, or enumeration ownership generation.
FolderView.Enumeration.cpp:1664-1686 takes the first committed same-folder entry
and can overwrite the current surviving focus with its old replacement. If the
user moves from B to E while deletion is pending, a later refresh moves focus
back to C. A stale snapshot where B still exists erases the intent prematurely,
and coalesced parallel completions consume only one find_if result, potentially
stranding the newer intent with no later accepted refresh.

Capture focus ownership and provider/path/sort epochs; invalidate the automatic
handoff when the user or navigation changes ownership. On an accepted refresh,
evaluate all applicable committed intents, keep an unreflected intent until a
newer snapshot or expiry, and choose the newest deterministic applicable result.
Test refocus-before-completion, navigation/provider/sort change, stale-then-
reflected snapshots, two completions in both orders collapsed to one generation,
and ordinary sequential success.

## Considered and rejected during re-audit

- **Not a current defect:** mixed old/new Terminal binaries. The user confirms
  host and Terminal.dll are compiled and staged atomically from this source
  tree. L4-CHORD-10 owns only the contradictory general rule/waiver and the
  atomic-staging proof; it does not claim a supported mixed-binary crash.
- **Rejected:** a data race/UAF in the new cached accessibility root. All current
  rootProvider reads/writes occur beneath the recursive target mutex through
  reviewed call paths, unregister detaches under that mutex, and the retained
  target becomes host-null/snapshot-empty. L4-DXUI-01 is the narrower duplicate
  COM-identity/event-routing defect.
- **Rejected:** the Terminal pointer-follow patch focuses Preview or hidden
  content. The new gate requires an open, selected Terminal tab; focused tests
  and the authoritative behavior agree.
- **Rejected:** posted-payload, Terminal callback-drain, module-unload, and
  floating-host ownership defects. Init/drain ordering, callback barrier,
  instance retention/module pin transfer, WIL ownership, and callback clearing
  were consistent in the reviewed paths.
- **Narrowed:** unknown paths still conservatively widen and ordinary
  two-included-side renames work. L4-STAR-12 is the distinct underwidening path:
  the explicit Specs/** rule suppresses that fallback for executable machine
  contracts and runtime data.
- **Rejected:** Resume always starts at artifact epoch zero. The runner loads
  origin run-state before fingerprinting and seeds its terminal epoch.
- **Rejected:** portable consumers fail to verify downloaded bytes. Consumer
  verification rejects reparse, size/hash mismatch, and extra files; the defect
  is the producer's missing receipt-to-payload binding in L4-BUILD-06.
- **Rejected:** File Operations hosted-tree focus/capture teardown, footer
  accessible name, change-case focus restoration, WM_SIZE propagation, async
  menu lifetime, payload drain, or WIL/D2D resource ownership as additional
  defects. Tooltip routing, graph theme/color semantics, delete-focus ownership,
  and menu DPI containment are the narrower confirmed findings above.
- **Rejected:** HostPromptPresentation creates a mixed-binary ABI defect. It
  consumes one reserved uint32_t without changing HostPromptRequest size, and
  the supported host/plugin product is compiled and staged atomically.
- **Rejected:** the host does not query `IFileSystemAtomicWriter`. It does; Curl
  lacks the interface and therefore takes the incompatible sibling-writer path.
- **Rejected:** pass `ALLOW_OVERWRITE` to the host's random temporary Curl writer,
  or add a cosmetic Curl atomic-writer interface. Both weaken capability truth
  without supplying conditional publication.
- **Rejected:** set Curl `operations.write=false`. Explicitly overwrite-authorized
  writers exist; the current lie is import eligibility and no-overwrite mutation.
- **Rejected:** treat cleanup failure as primary transaction failure. Publication
  is already irreversible; cleanup is separate debt. Also rejected: automatic
  remote prefix scavenging without immutable identity.
- **Rejected:** Curl's overwrite rollback makes no-overwrite safe. It still probes
  before an unconditional rename and later deletes by path.
- **Rejected:** copy every Kitty image referenced by a storage generation as the
  preferred fix. Source bytes can expand beyond the converted cap and capture
  work would scale with offscreen storage; use a bounded visible delta cache.
- **Rejected:** a Kitty borrowed-pixel UAF, stale-generation draw, device-loss CPU
  loss, or unload race as separate current defects. Pixel bytes are copied under
  `_terminalMutex`, generation gates reject stale results, device discard retains
  CPU BGRA, and close joins before runtime unload.
- **Rejected:** reuse the delete-on-close temporary helper for selected-path
  manifests. Arbitrary child processes cannot be required to inherit its share
  contract. Also rejected: stronger `rsa????.tmp` naming in shared `%TEMP%`,
  check-then-path-delete, or a proprietary length-prefixed external format.
- **Rejected:** require a persistent selected-path no-delete guard. It is not
  needed when identity comparison and disposition use one cleanup handle, and it
  can break supported exclusive readers/replacement behavior.
- **Not reopened:** the ten-minute selected-path fallback; it is an explicit
  product decision. The repair makes its eventual cleanup identity safe.
- **Rejected:** move Monitor's current deep copy to a worker while keeping the
  shared lock. That transfers the stall into UI ingestion/append lock waits and
  does not provide bounded snapshot capture. Also rejected: catching
  `std::bad_alloc`; repository policy keeps it fatal.
- **Not a finding:** Monitor export excluding metadata prefixes. Current tests and
  authoritative behavior intentionally export raw text; richer metadata remains
  a separate product decision.

## Execution order and integration boundaries

Implement the follow-up containment packages before trusting a receipt, Full
verdict, evidence prune, or package publication. Keep destructive-path,
attestation, runner-verdict, UI/runtime, and performance changes independently
reviewable:

1. **Package F0-A -- destructive local-path containment (complete 2026-08-15):** L4-STAR-02 and
   L4-BUILD-08 direct-child/component-safe resolvers, junction RED proofs, and
   atomic publication/deletion plans. Land before any local evidence cleanup or
   ZIP overwrite.
2. **Package F0-B -- fail-closed verdicts (complete 2026-08-15):**
   L4-STAR-03/04/05/13 normalized terminal results, mode-independent
   coverage-plan identity, shared verified evidence loader, complete semantic
   binding, and typed exits. Do not consume Startrail PASSED/Resume evidence
   produced before this package.
3. **Package F0-C -- source/receipt/impact truth (complete 2026-08-15):** L4-STAR-01/12 and
   L4-BUILD-01/02/03 rename-side identity, effective test surface, semantic
   receipt validation, governed output membership, exact runtime closure, and
   executable-Spec impact mappings.
4. **Package F0-D -- packaging/portable provenance (complete 2026-08-15):** L4-BUILD-04/06/07
   transactional MSIX input, receipt-derived portable payload, and one resolved
   toolchain identity.
5. **Package F0-E -- profile/epoch crash consistency (complete 2026-08-15):** L4-BUILD-05 and
   L4-STAR-06/07 selected-profile writer parameterization, pre-mutation receipt
   suspension, reader barrier, and immutable epoch decisions.
6. **Package F0-F -- evidence bytes/confinement/planning (complete 2026-08-15):**
   L4-STAR-08/09/10/11/14/15 atomic content-bound result artifacts, aggregate
   quotas, canonical bounded reads, shared direct-origin resolution, actionable
   build planning, component diagnostics, and derived dependency closure.
7. **Package F0-G -- Chordforge runtime correctness (complete 2026-08-15):** L4-CHORD-01/02/05/06
   output-buffer validation, typed state, tab handling, callback pair contract,
   and runtime plugin/palette tests.
8. **Package F0-H -- shortcut/model governance (complete 2026-08-15):** L4-CHORD-03/04/07/08/09/10
   scope-aware physical/pass-through semantics, Unicode canonicalization,
   effective aliases, message registry, and explicit lockstep ABI waiver/staging
   proof.
9. **Package F0-I -- DxUi/File Operations interaction correctness (complete 2026-08-16):**
   L4-DXUI-01/02/03/04 and L4-FILEOPS-09/11/12/14/15 canonical accessibility
   identity, owner-thread shutdown, passive delayed tooltips, DPI containment,
   single-stream graph semantics, saturating totals, localization, exact delete
   outcomes, and user-owned focus epochs.
10. **Package F0-J -- graph architecture/color/performance (complete 2026-08-16):** L4-FILEOPS-10/13
    bounded color slots, shared linear hosted/fallback renderer, exact row-band
    colors, deterministic counters, Release before/after evidence, and
    authoritative work-bound spec.
11. **Package F0-K -- Startrail governance closeout (complete 2026-08-16):** L4-STAR-16 reconciles the
    historical Done record, active ownership, cross-links, and indexes only
    after F0-A..F0-F are complete.

### Package F0-A completion evidence (2026-08-15)

- RED: the portable-package source contract failed because ZIP compression still
  force-overwrote the final path; the prune regressions failed because apply accepted
  arbitrary literal paths and plans carried no directory identity.
- Repair: ZIP construction now rejects reparse ancestors before lock/output mutation,
  compresses and smoke-tests a unique same-directory staged archive, revalidates the
  complete path, and publishes through guarded atomic move/replace. Evidence prune
  plans now contain only direct-child canonical paths plus Windows volume/file identity;
  apply revalidates the trusted chain and exact identity immediately before recursion.
- Destructive-safety proof: `.build`, `AppPackages`, final-file, and nested prune
  junction/symlink fixtures all failed closed; external sentinels and prior final bytes
  survived. Plan/apply directory replacement was also rejected by identity drift.
- Green: `BuildReproducibility.Tests.ps1` 19/19,
  `ValidationEvidence.Tests.ps1` 11/11, all five new guarded-path/publication cases in
  `BuildOutputProcess.Tests.ps1`, `ToolInventory.Tests.ps1` 16/16,
  `ToolingGovernance.Tests.ps1` 12/12, tool inventory 172/172 with zero findings, and
  specification inventory with zero blocking findings.
- Durable contracts reconciled: `Build_Toolchain.md`, `Installer_PortableZip.md`,
  `Testing_ValidationEvidence.md`, `Testing_ToolingGovernance.md`,
  `Core_SharedHelpers.md`, `Tools/README.md`, and `Tools/tool-inventory.json`.

### Package F0-B completion evidence (2026-08-15)

- RED: malformed, wrong-suite, zero-discovery, all-skipped, count-mismatched, or
  artifact-incomplete results could remain green; aggregation trusted readable
  labels without one semantic/digest verifier; exit states collapsed to 0/1; and
  Resume persisted a Fresh-mode plan identity.
- Repair: every adapter now emits one schema-validated normalized terminal result;
  promotions bind its digest; Resume, aggregation, history, and prune provenance use
  one verified-evidence loader; native skips carry an explicit reviewed class; and
  the runner maps PASSED/FAILED/NOT_EVALUATED/invalid/blocked/corrupt to exits 0..5.
- Identity repair: the stable coverage-plan digest excludes invocation mode while
  the invocation-plan digest includes the actual Fresh, Resume, or Affected mode;
  entry fingerprints bind the stable coverage identity, so truthful Resume records
  can reuse compatible Fresh evidence.
- Green: serial full Debug rebuild completed with 0 warnings and 0 errors;
  `ValidationEvidence.Tests.ps1`, `ValidationFingerprint.Tests.ps1`, and
  `RunAllTestsPlan.Tests.ps1` passed 84/84; tooling-governance tests passed 28/28;
  tool inventory classified 172/172 files with zero findings; and specification
  inventory reported zero blocking findings.
- Durable contracts reconciled: `Testing_ValidationEvidence.md`,
  `Testing_SelfTests.md`, `Tools/README.md`, `Tools/tool-inventory.json`, all
  affected validation schemas, and the native selftest JSON contract.

### Package F0-C completion evidence (2026-08-15)

- RED: included-to-excluded renames erased the included source-side change;
  executable schemas/data below `Specs/` matched a prose-only no-validation rule;
  receipt test-surface identity came from configuration rather than compiled
  projects; the mutable pointer could disagree with its immutable record; and a
  recursive output scan attested stale or unrelated binaries.
- Source/impact repair: staged and unstaged rename sides are classified
  independently, with an included deletion/addition retained across `.log` and
  `.build` boundaries. Ordered impact rules now map build/validation schemas,
  settings/theme data, performance budgets, terminal corpora, and governance maps
  to their real consumers; unknown `Specs` JSON/JSON5 widens to Full and reviewed
  Markdown prose remains inert.
- Receipt repair: `RSBuildEnableTests` is resolved once and passed explicitly;
  per-project MSBuild manifests record the effective value, target, vcpkg write
  tlog, and native runtime inputs. Receipt output/closure is derived only from
  current declarations, full-solution builds reject undeclared EXE/DLL files, and
  targeted builds never attest unrelated survivors. The shared loader recomputes
  receipt identity, byte/semantic-matches the immutable record, and validates
  case-insensitive IDs/paths plus exact closure references.
- Full closure proof: CI/Full now requires every selected file-backed executable
  and plugin at initial `-SkipBuild` acceptance and after writer re-attestation.
  A production Release `RedSalamander` rebuild published receipt `16e86448...`
  with `tests_enabled=false` and 119 declared outputs. A test-enabled full Release
  rebuild published `51a09aa8...` with `tests_enabled=true`, 174 outputs, and all
  13 file-backed artifacts in the 18-entry Full plan verified. The matrix also
  removed an erroneous test-only guard from the production accessibility provider
  dispatch and bounded the supported toolchain's oversized test-enabled Release
  accessibility TU with object-embedded debug information.
- Green: full Debug and test-enabled full Release builds completed with zero
  errors; the focused Startrail/build/runner/inventory suite passed 152/152;
  `BuildReproducibility.Tests.ps1` passed 19/19 after the Release matrix hardening;
  tool inventory classified 172/172 files with zero findings; and specification
  inventory reported zero blocking findings.
- Durable contracts reconciled: `Build_Toolchain.md`, `Testing_SelfTests.md`,
  `Testing_ValidationEvidence.md`, `Testing_ToolingGovernance.md`,
  `Tools/README.md`, `Tools/tool-inventory.json`, and the validation-impact schema
  and policy data.

### Package F0-D completion evidence (2026-08-15)

- RED: MSIX versioning rewrote the tracked manifest before the final source barrier;
  portable publication rescanned mutable live output instead of consuming receipt
  membership; and the receipt identity omitted the actual compiler, linker, resource
  compiler, and representative SDK bytes later claimed by a separate workflow record.
- Transactional MSIX input: the tracked manifest is now read-only. Version and
  architecture stamping writes a unique CreateNew manifest beneath guarded
  `.build\AppPackages` staging, and both integrated and standalone packagers pass its
  exact `RSAppxManifestPath` to the WAP project and remove only that validated stage in
  `finally`. Duplicate-output failure leaves the source byte-identical. A failed full
  Release matrix attempt left the source unchanged, zero stage directories, and no
  success receipt; the subsequent valid test-enabled full Release build packaged the
  nondefault `7.0.184.0` identity successfully.
- Portable authority: producer attestation now revalidates every receipt output's
  regular-file status, size, and SHA-256, validates the receipt-bound toolchain record,
  and copies only that exact set into a unique payload root. The only added payload file
  is the schema-validated attestation. Tests reject post-receipt mutation and missing
  files, exclude an unreviewed live-output extra, reject consumer extras, and rebind the
  exact downloaded set locally.
- Toolchain authority: `ToolchainIdentity.schema.json` governs one MSBuild-resolved
  record covering the actual MSBuild, `cl`, `link`, `rc`, `Windows.h`, `kernel32.lib`,
  and `ucrt.lib` byte seams. Its digest binds CI cache keys, local receipts, portable
  evidence, and workflow outputs; CI makes `build.ps1` consume the same explicitly
  selected MSBuild path used by the pre-cache probe. Each individual byte-seam mutation
  changes identity in focused tests.
- Real proof: Debug targeted receipt `91ec9560...` contained 120 outputs and a matching
  canonical toolchain digest. Full x64 Release plus MSIX completed with zero errors and
  published receipt `5cef07aa...` with 203 outputs, one `7.0.184.0` MSIX, and one matching
  toolchain record. The tracked manifest remained SHA-256 `0aa3458e...`, and the manifest
  staging root was empty after success.
- Green: `BuildEvidence.Tests.ps1` plus `BuildReproducibility.Tests.ps1` passed 34/34;
  tool inventory classified 172/172 files with zero findings; specification inventory
  reported zero blocking findings. Durable contracts were reconciled in
  `Build_Toolchain.md`, `Tools/README.md`, and `Tools/tool-inventory.json`.

### Package F0-E completion evidence (2026-08-15)

- RED: the Full deployment writer resolved and deleted `.build/x64/Debug` regardless
  of the selected runner profile; the runner neither suspended the selected receipt nor
  enforced its artifact-read barrier before launch; and re-attestation replaced the
  already-published decision path in place. Moving profile validation to file scope also
  initially made the read-only Pester lane fail while excluding the tagged writer.
- Selected-profile repair: the runner supplies `RS_VALIDATION_PLATFORM` and
  `RS_VALIDATION_CONFIGURATION` only around the isolated writer, and the deployment test
  validates those values inside its tagged `BeforeAll`. Output selection, lock scope,
  deletion, and child `build.ps1` arguments all use that exact profile. Plan-contract
  tests cover x64 Release plus ARM64 Debug/Release and reject an x64 Debug sibling path.
- Crash-consistency repair: before the writer, the runner verifies the stable barrier,
  suspends the selected receipt, then persists a new `mutating` epoch and invalidates
  intersecting promotions. Readers require stable state plus matching epoch/receipt.
  Re-attestation publishes a verified same-profile receipt, then one run-state replacement
  switches snapshot, receipt, decision digest, epoch, and mutation state back to stable.
  A dead mutating checkpoint repairs to interrupted and remains unreadable.
- Immutable-decision repair: each epoch decision is stored create-only at
  `decisions/<decision-digest>.json`; concurrent publication accepts only identical
  canonical content. State points at the active digest, and focused crash/tamper tests
  prove that the old bytes survive and no split state becomes reusable.
- Contained writer launch: the exclusion-safe profile change exposed a hidden test
  dependency on ambient module imports and a CIM-denied immediate-parent query. Optional
  lock hooks are now invoked through their exact discovered commands, and direct-child
  authentication uses a local native parent query while retaining kill-on-close job,
  one-use token, and no-second-hop checks. The focused containment suite passed 28/28.
- Green: the complete current read-only tooling profile passed 613/613; the focused
  Startrail fingerprint/evidence/runner set passed 88/88 before the added create-only
  regression, and `ValidationFingerprint.Tests.ps1` then passed 20/20; tooling inventory
  and governance tests passed 28/28. Durable contracts were reconciled in
  `Build_Toolchain.md`, `Testing_SelfTests.md`, `Testing_ValidationEvidence.md`,
  `Tools/README.md`, `Tools/tool-inventory.json`, and `ValidationRunState.schema.json`.

### Package F0-F completion evidence (2026-08-15)

- RED: copied result artifacts were path-only mutable files, copied non-atomically,
  bounded only per file, and trusted by existence on Resume. The JSON reader accepted
  pretty/reordered/alternately escaped/decomposed input and constructed object graphs
  before node/string bounds. Resume used a prefix check rather than one safe run resolver.
- Artifact repair: result artifacts now publish by flushed create-only move and persist
  `{path,size_bytes,sha256}`. Source and destination are reopened and rehashed; evidence,
  Resume, and aggregate loading rehash again. The 64 MiB entry and 1 GiB run quotas fail
  before publication. Tests cover two 40 MiB files, run-quota crossing, source mutation,
  destination truncation before promotion, and post-promotion tamper.
- Canonical/confinement repair: a streaming token scan enforces depth, nodes, strings,
  duplicate normalized keys, integers, and supported tokens before materialization; the
  parsed value must then reproduce the exact canonical payload. One resolver now accepts
  only an exact direct non-reparse child whose leaf equals `run-state.run_id`; nested,
  path/case alias, copied-renamed, and junction origins fail closed.
- Planning truth: `expected_build_action` was renamed to `advisory_build_action`, and the
  authoritative spec states that it explains impact but does not dispatch the independent
  runner build. Entry fingerprints persist `validation-entry-components.v1`; Resume
  compares build, logical, shared, runner, environment, arguments, coverage, outcome,
  capability, epoch, then source components before aggregate fallback. Single/simultaneous
  mutation and unknown-future-component tests prove exact stable reasons.
- Derived closure/governance: runner inputs are recursively derived from
  `Tools/Run-AllTests.ps1`, including `TestRunSummary.psm1` and
  `RunAllTestsSummary.schema.json`. Dependency-critical inventory entries opt into
  `consumerClosure: derived-references`; removing a real import edge from `usedBy` is a
  blocking finding. Tool inventory remains 172/172 with zero findings.
- Performance/correctness evidence: the optimized 52-file derived closure passed its
  deterministic recursive fixture and current-runner assertions. Same-machine diagnostic
  samples were 420.19..520.01 ms (median 452.49 ms) versus a directional single 8.2 s
  pre-index observation; the three-file archive passed pre-stage validation at
  `Specs/TestRuns/4cb089111a23/Validation/2026-08-15_154800_f0f_dependency_closure/`.
  Focused fingerprint/evidence tests passed 46/46 after closure coverage; impact passed
  7/7, ToolInventory passed 17/17, and RunAllTestsPlan passed 53/53.

### Package F0-G completion evidence (2026-08-15)

- ABI/callback repair: `GetActionState` validates the captured caller
  `sizeBytes` before writing and preserves oversized tails; `SetCallback` accepts
  only both-null or both-non-null pairs. Plugin contracts cover null, undersized
  full-buffer canary, current/oversized output, all four callback/cookie pairs,
  unload quiet point, and 59/59 deterministic Terminal selftests.
- Runtime-state repair: `CommandRegistry` state ownership flows through the
  shared `CommandRuntimeState` resolver to Terminal menus, direct/shortcut
  dispatch, and Command Palette rows. The exact invocation HWND plus Terminal
  instance/session identity is captured and revalidated before activation.
  Host-owned Close/Context/Session commands no longer query the plugin action
  allowlist; no-selection Copy and unavailable numbered tabs render disabled.
- Dispatch truth: numbered embedded tab selection returns the real
  `selectContent` result. A dedicated live-Terminal case proves Folder/Preview/
  Terminal identities 1..3 through state query, direct dispatch, shortcut, and
  palette paths; disabled Enter remains inert and closing the captured Terminal
  makes a formerly enabled row stale. Enabled palette Enter was also repaired
  to run before the focused text field consumes it and now opens Display
  Shortcuts through the real command path.
- Green: final focused Commands artifact passed 6/6, including the dedicated
  state case and both embedded/floating real-shell exit auto-close cases;
  `PluginContractTests --terminal-selftests` passed; targeted Debug receipt
  `8ce8c937...` built with zero warnings/errors.

### Package F0-H completion evidence (2026-08-15)

- Scope repair: Application lookup now carries original scan/extended identity;
  Folder View and Terminal retain physical lookup; Function Bar rejects physical
  identity in schema, loader, import, and normalizes capture to virtual key.
  `cmd/shortcut/passthrough` is accepted only in Terminal and startup recovery
  discards only the invalid shortcuts section.
- Model repair: suggestion dedup/ranking uses one `CompareStringOrdinal`
  ignore-case ordering relation rather than a locale-sensitive hash/equality
  mismatch. Accented Latin, Greek, Cyrillic, current/persisted ranking, and a
  guarded Turkish-locale switch pass. Palette chips now derive from the
  effective context precedence and suppress pass-through, unassigned,
  different-command, and duplicate same-command shadows.
- Governance/policy repair: Floating Terminal private messages moved into
  `Common/WindowMessages.h`; the source contract rejects local declarations and
  duplicate central literal IDs. Terminal and Plugin API specs explicitly
  document the retained-IID callback insertion as one closed lockstep waiver,
  not mixed-binary compatibility. Receipt `8ce8c937...` has 120 governed outputs;
  the app closure contains `Terminal.dll`, private runtime DLL and identity, four
  Terminal satellites, and five required notices.
- Green: Terminal source contracts passed 28/28, localization contracts 5/5,
  SettingsSchemaTests passed, the final Commands gate passed 6/6, and all
  authoritative contracts were reconciled in `Terminal_EmbeddedPlugin.md`,
  `Plugins_PluginAPI.md`, `Core_SettingsStore.md`, and
  `UI_CommandMenuKeyboard.md`.

### Package F0-I completion evidence (2026-08-16)

- UIA root creation now uses one retained-target canonical COM identity across
  root acquisition, focus/hit fallback, fragments, events, and repeated calls.
  Unregister empties the snapshot, clears the host, disconnects the exact root,
  and retires the HWND map. The exact-same-HWND detach/reattach regression proves
  a distinct new identity while the retained old provider routes neither focus
  nor point hits into the new attachment.
- `WindowHost::Detach()` rejects a foreign attachment thread, and process-exit
  shutdown marshals foreign-owned detach/TSF teardown to the HWND owner thread.
  Passive supplemental tooltips use the deepest retained hover path, shared
  delay/lifetime, retained-tree invalidation, and click-through behavior. Menu
  DPI relayout bounds both popup extents to the destination work area and keeps
  oversized content scroll-reachable.
- File Operations global totals use saturating addition, and all five resource
  roots use separate count-neutral Running/Waiting/Needs-attention clauses with
  the retired combined resource removed. Exact per-source delete outcomes now
  prove removal only for `S_OK`; provider-sensitive identities and focus/path/
  sort/provider epochs defer replacement selection to the nearest actual row in
  a newer accepted enumeration, with stale retention and newest-intent wins.
- Focused proof: Debug DxUiTests receipt
  `b0b5caca3385cf7a0ed00f361ad235a5246c4b2393857653011cd55b2ec0b785`;
  Accessibility and Rendering passed. RedSalamander Debug receipt
  `699e00c4788f8878909a24ead536443e1cd095abd5f683356ddfbbe0260c4a87`;
  exact focus archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-08-16_070621/` passed 1/1,
  and exact real-popup fallback archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-08-16_071803/` passed 1/1.
- Authoritative contracts are reconciled in `UI_DxUiSharedGrid.md`,
  `UI_FileOperationsPopup.md`, and `UI_FolderView.md`. Localization contracts
  passed 6/6; Tool Inventory passed 17/17; Tooling Governance passed 12/12;
  tool and spec inventories report zero blocking findings.

### Package F0-J completion evidence (2026-08-16)

- History is bounded to 180 samples with 16 stable reusable stream-color slots;
  aggregate-only samples do not expand the palette, and the shared semantic gate
  handles high contrast, Rainbow single-stream, and default concurrent-stream
  behavior. Hosted and fallback paths use the same canonical band painter and
  row-color conversion.
- The accepted implementation performs bounded `O(samples * slots)` raster work
  into one fixed 180x64 premultiplied-BGRA bitmap, with one clipped geometry and
  one draw. The real fallback case proves 180 samples, 16 slots, 2,864 logical
  quads, one geometry, 16 production row-color matches, and zero mismatches.
- Same-machine accepted Release evidence compares
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-15_175346_graph_before/`
  with `Specs/TestRuns/4cb089111a23/FileOps/2026-08-16_064500_graph_after/`.
  Five-run median improved from 18,493 us to 16,583 us (10.3%) and mean from
  18,768.4 us to 16,593.2 us (11.6%); after rows report painter median 81 us and
  the exact 180/16/2,864/one-geometry bound. Both four-file evidence archives
  passed archive validation, and the lasting architecture/metric contract is in
  `UI_FileOperationsPopup.md`.

### Package F0-K completion evidence (2026-08-16)

- The Startrail Done plan now identifies itself as non-normative implementation
  history and links this active cross-domain plan as the sole post-completion
  owner of `L4-BUILD-01..08` and `L4-STAR-01..16` until remediation closeout.
- `Build_Toolchain.md`, `Testing_ValidationEvidence.md`, and
  `Testing_ToolingGovernance.md` carry the same explicit ownership link while
  remaining the authoritative behavior contracts. The WIP index already lists
  this plan as active rank I0 and sole implementation owner, so no competing
  active authority remains.
- `git diff --check` passed; Tooling Governance passed 12/12; spec inventory
  reports zero blocking findings.

### Continuation checkpoint -- F0-I/F0-J complete; remaining packages resume (2026-08-16)

This checkpoint is stored in this original, renamed-in-place plan so a later
session can resume without creating a second plan or a new plan folder. The old
`CodeReview_LastFourDays_Remediation_2026-08-10.md` path is absent; Git reports
one rename to this cross-domain name. The required `Specs/TestRuns/...` paths
below are test/performance evidence, not replacement plans.

The worktree is intentionally dirty with 124 status entries at detached `HEAD`
`e9c76e59aa2b2993170ea1e6c88e7d15362954c7`. Do not reset, clean, stash, or
replace it before reconciling the changes below. Do not make a bulk checkpoint
commit: the tree includes an unrelated user change in
`.agents/skills/improve/SKILL.md`, which must be preserved and excluded from this
remediation scope. The currently untracked
`RedSalamander/CommandRuntimeState.h`,
`Specs/Build/ToolchainIdentity.schema.json`, and
`Specs/Testing/ValidationTerminalResult.schema.json` are intentional outputs of
completed F0-B/F0-C/F0-G work, not disposable files.

F0-I implementation saved on disk:

- DxUi accessibility root acquisition now uses one target-owned canonical COM
  identity across root, hit-test, focus, fragment, event, and debug paths.
- WindowHost process shutdown marshals detach/TSF teardown to the attachment
  thread and rejects foreign-thread detach rather than tearing down from the
  caller.
- Passive tooltip tracking, delayed activation, retained-tree invalidation,
  click-through behavior, and shared menu DPI/work-area containment have focused
  DxUi coverage.
- File Operations aggregate totals saturate. The five resource files use three
  count-neutral Running/Waiting/Needs-attention strings, and the dead combined
  format was removed.
- Delete completion carries exact per-source results. FolderView focus handoff
  records provider/path/sort/focus ownership epochs, treats only exact `S_OK` as
  proven removal, retains stale intents, and coalesces to the newest applicable
  intent.
- Refreshed-nearest-survivor handoff is now resolved only after a newer accepted
  enumeration. Each source stores an exact-`S_OK` `provenRemoved` bit; the
  resolver searches the actual refreshed list forward and then backward around
  the captured focus index, skips proven removals with provider-sensitive path
  identity, chooses the refreshed display name, retains stale intents, and lets
  the newest applicable commit sequence win. The regression covers the case
  where the pre-delete next item also disappears but a farther refreshed
  survivor remains. The final RedSalamander Debug receipt is
  `699e00c4788f8878909a24ead536443e1cd095abd5f683356ddfbbe0260c4a87`
  with zero warnings/errors. The exact Commands case
  `cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs` passed
  twice after this repair; the latest archive is
  `Specs/TestRuns/4cb089111a23/Commands/2026-08-16_070621/` (1/1, 336 ms).
- The exact-same-HWND UIA lifecycle regression is implemented in
  `Tests/DxUiTests/DxUiTests.Accessibility.cpp`. It acquires the old canonical
  root identity, detaches and reattaches the same saved HWND, proves the new
  attachment has a different canonical identity, proves repeated new
  acquisition remains canonical, and proves old-provider focus and hit-test
  routing return null. Debug DxUiTests receipt
  `b0b5caca3385cf7a0ed00f361ad235a5246c4b2393857653011cd55b2ec0b785`
  built with zero warnings/errors; `DxUiTests.exe --suite=Accessibility` passed.
- The real-popup fallback graph proof is implemented in
  `FolderWindow.FileOperations.Popup.{h,cpp}` and
  `Commands.SelfTest.FileOps.cpp`. A test-only fail-next-host-attach seam leaves
  production logging unchanged, the existing popup fixture injects deterministic
  180-sample/16-slot data, and the snapshot reports actual fallback painter and
  active-row production-branch metrics. Exact case
  `cmd_pane_fileops_popup_fallback_graph_uses_shared_bounded_painter` passed 1/1
  in archive `Specs/TestRuns/4cb089111a23/Commands/2026-08-16_071803/`.
  It proved no DxUi host, zero hosted graphs, active fallback bands, 180 samples,
  16 slots, 2,864 logical quads, one rendered geometry, 16 exact row-color
  matches, and zero mismatches. The Debug Rendering suite then passed in full,
  including the graph-layering and hue-churn performance scenarios.
- `Tools/Tests/ResourceLocalizationContracts.Tests.ps1` now includes the
  count-neutral five-file contract. Its new assertions were made compatible
  with the repository-supported Pester 3.4 surface; the file passed 6/6.
- F0-I authoritative specification reconciliation is complete in
  `Specs/UI/UI_DxUiSharedGrid.md`, `Specs/UI/UI_FileOperationsPopup.md`, and
  `Specs/UI/UI_FolderView.md`; focused localization/tooling checks and both
  inventories are green.

F0-J implementation and measurement state saved on disk:

- Throughput samples are bounded to 180 and use stable reusable color slots
  `0..15`. Hue-weight deduplication is by slot. Aggregate-only samples do not
  expand the palette. The shared semantic gate disables bands for high contrast,
  enables Rainbow for one admitted stream, and enables default per-stream bands
  only for at least two concurrent slots.
- Hosted and fallback paths call the same `PaintThroughputGraphBands` renderer;
  active-row mini bars and graph bands use the canonical hue-to-color helper.
  Control/Rendering tests cover the semantic gate and the deterministic 180x16
  hue-churn scenario.
- Accepted legacy baseline evidence is at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-15_175346_graph_before/`.
  The five Release capture durations are `19451`, `19184`, `18225`, `18489`,
  and `18493` microseconds (median `18493`, mean `18768.4`), with `2864`
  legacy unique geometries. Raw `perf_metrics.jsonl` SHA-256 is
  `dcf9ade3f67f42030d5de0af479cc971ff8d94dbc608d363c294350c5600c7c6`;
  the four-file archive passed `Tools/Test-TestRunArchive.ps1` validation.
- Rejected, unaccepted renderer experiments were: sixteen multi-figure path
  geometries, sixteen D2D meshes, sixteen continuous polygons, and 2,864
  `FillRectangle` submissions. Although some CPU construction counters were
  small, captured frame time was about 33 ms, so none is closure evidence and
  none should be archived as the accepted after result.
- The accepted implementation in `Common/DxUi/DxUi.Controls.cpp` is a fixed
  180x64 premultiplied-BGRA raster, uploaded as one `ID2D1Bitmap`, clipped by one
  area geometry/layer, and drawn once with nearest-neighbor interpolation. The
  Rendering assertions expect one geometry.
- Test-enabled Release DxUiTests built with receipt
  `061eec7fb714bceeeb752fd2e0b24616aec12f0bd0478aea6175a45e3b04aa87`
  and zero warnings/errors. One Rendering-suite invocation executed its one
  warmup and five internal measured captures; all Rendering tests passed.
- Accepted after capture durations are `16583`, `16700`, `16526`, `16559`, and
  `16598` microseconds (median `16583`, mean `16593.2`). Against the accepted
  baseline this improves the median by 10.3% and mean by 11.6%.
  Band-painter durations are `81`, `76`, `82`, `71`, and `84` microseconds
  (median `81`, mean `78.8`). Every measured row reports 180 samples, 16 active
  slots, 2,864 logical quads, and one rendered geometry.
- Accepted after archive:
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-16_064500_graph_after/`.
  Raw `perf_metrics.jsonl` SHA-256 is
  `0c5bcf708a0bfd8f3e3d289d9a42e5cd925b16a83633a7f83498f7e48b4c7e19`.
  Its `perf_metrics.jsonl`, `README.md`, `results.json`, and `trace.txt` passed
  `Tools/Test-TestRunArchive.ps1` validation. `Specs/TestRuns` is ignored by
  default, so force-add this accepted evidence only when intentionally preparing
  the final scoped commit.
- F0-J authoritative specification reconciliation is complete in
  `Specs/UI/UI_FileOperationsPopup.md`, including the semantic gate, stable
  0..15 slots, 180x16 work bound, one raster/geometry, shared renderer,
  deterministic runtime/metric guards, and accepted before/after evidence.

F0-I/F0-J/F0-K resume gate is complete: `git diff --check`, localization 6/6, Tool
Inventory 17/17, Tooling Governance 12/12, tool inventory, and spec inventory
all passed. Do not repeat this lane unless a later aggregate failure points back
to it. Package A is closed below; continue with Package B;
the final Fresh Full gate remains mandatory before moving this plan to Done.

### Package A completion evidence (2026-08-16)

Package A is complete. Its production repair, real-parser/planner and
flag-sensitive test seams, focused runtime proof, and lasting specification
contracts agree.

Implemented state:

- `Plugins/FileSystemCurl/FileSystemCurl.h` and
  `FileSystemCurl.Shared.cpp` no longer advertise host imports for FTP, SFTP, or
  SCP. Same-provider operations and Curl-to-local export remain advertised.
- `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp` returns
  `E_INVALIDARG` when `ALLOW_REPLACE_READONLY` is supplied without
  `ALLOW_OVERWRITE`, before its generic unsupported/no-overwrite result.
- `FolderWindow.FileOperations.State.cpp` computes the effective per-item writer
  flags once, uses those exact flags for `SupportsAtomicWriterCommit(...)` and
  final-path `CreateFileWriter(...)`, and strips overwrite/read-only flags only
  after the atomic route declines and a random sibling path is selected.
- Test-only wrappers in `FolderWindow.FileOperationsInternal.h`,
  `FolderWindow.FileOperations.cpp`, and
  `FolderWindow.FileOperations.State.cpp` exercise the real host capability
  parser/planner and writer-route calculation.
- `FolderWindow.FileOperations.SelfTest.cpp` contains the flag-sensitive atomic
  writer fake. `FolderWindow.FileOperations.SelfTest.Phases05_06.cpp` verifies
  the exact probe/final flags, fallback stripping, real FTP/SFTP/SCP capability
  JSON, absence of the atomic-writer interface, contradictory-flag validation,
  rejected local-to-Curl import, and retained Curl-to-local export.

Validation state:

- Integrated RedSalamander Debug build passed with zero warnings/errors. Receipt
  SHA-256:
  `88aceb55d175bbd9e077137304b77ed17aa72123a31ce025cb766455209cde93`.
- Direct focused command:
  `.build\\x64\\Debug\\RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix`
  passed Setup, the focused case, and Cleanup. Archive:
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-16_073557/`; focused duration
  422 ms.
- `Tools/Run-AllTests.ps1 -Suite FileOps -SkipBuild` was not a failed product
  test: it correctly refused the targeted-project receipt because `-SkipBuild`
  requires a qualifying full-solution receipt. Do not record that refusal as a
  regression.
- `Specs/FileSystem/FileSystem_FtpSftpScp.md`,
  `Specs/FileSystem/FileSystem_FileOperations.md`,
  `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
  `Specs/Core/Core_FileSystemBridge.md` now own the lasting import-policy,
  contradictory-flag, host-planning, and effective atomic-writer flag
  contracts. `git diff --check` passed and spec inventory reported zero
  blocking findings after reconciliation.

### Package B completion state (accepted 2026-08-16)

Package B is **complete and accepted**. L4-CURL-03/06 are closed by the focused
Curl selftests, recursive/batch coverage, authoritative specs, and archived
same-machine Release performance evidence below. Package C is the next package.

Current implementation state:

- `FileSystemCurl.Internal.h` adds an optional expected-size commitment to
  `CurlDownloadToFile(...)`.
- `FileSystemCurl.Shared.cpp` checks the downloaded file size after a successful
  transfer, observes the real data-stream terminator for committed downloads,
  and returns `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)` on a committed-size
  mismatch. A committed `CURLE_PARTIAL_FILE` is normalized before transient
  retry when the staged length proves the mismatch.
- `FileSystemCurl.CopyMove.cpp` resolves a targeted remote size before using a
  directory-listing size, carries the commitment through native Copy/Move, and
  preflights recursive destructive Move before destination-tree mutation. Copy
  may proceed when the size is unknown because the source is retained; Move
  fails closed when it cannot prove the source size. Targeted probe
  unavailability may fall back to listing/unknown policy, while cancellation,
  authentication, disappearance, and transport failures now propagate instead
  of being collapsed into an ordinary unknown-size result.
- `FileSystemCurl.Internal.h` and `FileSystemCurl.Shared.cpp` now own one
  `ResolveCurlSourceSizeCommitment(...)` helper used by native Copy/Move and
  direct readers. FTP 500/502/504 SIZE responses become explicit primitive
  unavailability, FTP 530 becomes `ERROR_LOGON_FAILURE`, and other hard probe
  failures propagate. FTP control reply capture is deliberately narrow: the
  libcurl debug callback retains only numeric status classification for the
  current probe, never payloads, commands, or credentials, and does not log the
  control stream. A successful targeted content-length proof wins over earlier
  non-fatal control replies; otherwise authentication, disappearance (`550`),
  other hard failures, primitive unavailability, and transport failures are
  classified without hiding them behind listing metadata. This consolidation
  and its focused tests are built and green.
- `FileSystemCurl.DirectoryOps.cpp` tracks the expected and received bytes for
  each `CurlStreamingReader` generation, rejects known-size premature EOF and
  overrun, withholds a read that reaches the committed end until terminal
  validation, and permits same-position restart after a failed generation. It
  issues FTP `REST <offset>` through the FTP-only
  `CURLOPT_PREQUOTE` lifetime contract, retains
  `CURLOPT_RESUME_FROM_LARGE` for SFTP/SCP, and independently validates the
  terminator and byte count. Its committed-end wait now explicitly refuses to
  await terminal completion when the remaining range exceeds the 1 MiB ring
  buffer, with a focused deadlock-regression assertion. A later static audit
  also corrected the zero-length seam: `startOffset == committedSize` now
  performs and validates the transfer, so a nonempty body cannot hide behind
  immediate successful EOF; only offsets beyond the commitment fail locally.
- `FileSystemCurl.SelfTest.Ftp.cpp` extends the fake FTP server with complete,
  strict-prefix, and overlong RETR delivery; adds directory-aware LIST/CWD/MKD/
  RMD state and stale listing-size injection; covers direct real-reader cases
  plus single, recursive, and batch Copy/Move across known-short,
  known-overlong, unknown-size, stale-listing, and targeted-authentication
  inputs; and emits
  `FileOps.Curl.SourceCommitmentGuard`. The collection matrix is compiled and
  green. Unknown-size files use a structurally valid
  Unix LIST record with a nonnumeric size token; this prevents the production
  mixed directory/file parser from discarding a bare filename after it has
  parsed the first directory record. The real-reader matrix includes
  exact-zero and zero-commitment-overlong bodies. It also makes the seek proof
  assert exactly one `REST 137`, one RETR, and the exact remaining byte count,
  and makes unknown-size recursive Move explicitly assert zero STOR in addition
  to zero RETR/MKD/RMD/DELE. All pre-transfer failure checks reject every STOR,
  not only a command whose path happens to contain the expected staging marker.
  The final matrix also covers a targeted `550` disappearance for single Copy,
  single Move, and the direct reader, with no RETR/STOR/DELE side effects. The
  complete matrix is compiled and passes.
- `Tests/PluginContractTests/PluginContractTests.cpp` adds the focused
  `--curl-selftests` entry point. Keep this narrow deterministic rerun surface.

Latest successful build receipts:

- Integrated RedSalamander Debug:
  `5bc02d07d79af874acda1d96199e2b618854b2212ab13c147baa0cf5b791fb85`.
- FileSystemCurl Debug after the focused diagnostic helper:
  `91edf470fdf62d9055d165673c96cb9e5c39041d7948b3330baa15c8e8101d7a`.
- PluginContractTests Debug with `--curl-selftests`:
  `cbbc4a64de52a1d1747e61a6eb7d73e03e88ec3eea7b0161a56ccb1a916a0fda`.
- Successive FileSystemCurl Debug receipts while reducing the reader failures:
  `366c35b7afc8f3bc9a52b585698e114a1ebc130a7bec3c01c9b904b15c576a02`,
  `edd83c1fd65d66f4e1e2c4ebf128a29908ae3ee1174a35f66c2986a1ea46093e`,
  `17cf168424b11f7c9c42bb6cd6bae25ef728c9c9ae7116e1663a62292027d6bf`,
  and latest verified `df8f3a01d188335ebca50f4777688a628afc9e39527a56a561e6257219019d20`;
  each completed with zero warnings/errors.
- Final Package B focused Debug build receipt:
  `817f5262c9fae2da58d97985b8716997608a6532b2edaf7832857ff33450f861f`;
  zero warnings and zero errors.
- Test-enabled x64 Release receipts: FileSystemCurl
  `47bd49a72bf0bb552e1be107ee8fa370944d3e426950e96b09e666fbcf10838a`
  and PluginContractTests
  `375cf500f5765eaaf32e102158442097a5d4cb9e7998114d460a9e51c316beaf`;
  both zero warnings and zero errors.

Latest focused results:

- `.build\\x64\\Debug\\RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix`
  still passes; archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-16_075633/`, focused
  duration 422 ms. This host case does **not** execute the Curl DLL debug
  selftests and is not Package B proof.
- The original Package B run was `155` passed / `5` failed. After normalizing
  committed partial transfers and waiting for terminal reader validation, native
  short Copy/Move, reader premature EOF, overlong body, and failed-generation
  restart all pass.
- The former `159/1` seek failure proved that
  `CURLOPT_IGNORE_CONTENT_LENGTH` suppressed FTP's SIZE-dependent resume route
  when paired with `CURLOPT_RESUME_FROM_LARGE`. The repair keeps length ignoring
  for FTP but supplies `REST 137` as a pre-transfer FTP command whose
  `curl_slist` remains alive through `curl_easy_perform`; SFTP/SCP retain their
  normal resume option. Static protocol review against the
  official `CURLOPT_PREQUOTE` and libcurl error-code documentation confirmed
  that PREQUOTE is FTP-only, the command list must outlive the transfer, and an
  FTP quote response is an error only at 400 or above; the fixture's standard
  `350` REST response is therefore valid. References:
  <https://curl.se/libcurl/c/CURLOPT_PREQUOTE.html> and
  <https://curl.se/libcurl/c/libcurl-errors.html>.
- The final **verified**
  `.build\\x64\\Debug\\PluginContractTests.exe --curl-selftests` run exits `0`
  with `199` passed / `0` failed (`hr=0x00000000`), and its module-unload quiet
  point passes. This includes the seeked-range repair, the complete
  single/batch/recursive matrix, exact-zero/overlong validation, targeted auth
  propagation, and targeted disappearance propagation.
- A broad `PluginContractTests.exe` run also reports the existing volatile
  per-user registry-fixture failure. Treat it separately; use
  `--curl-selftests` while closing Package B so that unrelated broad-suite state
  cannot obscure the Curl result.
- The post-checkpoint static audit completed on 2026-08-16: declaration/definition
  and call-site searches agree on `ResolveCurlSourceSizeCommitment(...)`; the
  reader generation, buffer, zero-range, REST-option lifetime, and fake-server
  command paths were inspected; and focused `git diff --check` passed with only
  the repository's expected LF-to-CRLF warnings. This review evidence
  complements the now-green focused executable matrix.
- Five separate test-enabled Release processes each passed `199/199`, returned
  `hr=0x00000000`, and reached the plugin unload quiet point. The accepted
  archive is
  `Specs/TestRuns/4cb089111a23/FileOps/20260816_085809_curl_source_commitment_release/`:
  140 valid JSONL rows, five claim rows, raw SHA-256
  `d40e99beeb400d869998c4392094c701a8c965c903297d4c430fd9043c7a9d13`.
  `FileOps.Curl.SourceCommitmentGuard` measured
  `4000,6000,6000,5000,6000` us (median `6000` us), exactly 22 source
  commands and 4,098 deliberately short bytes per run, and
  `0x8007012B` before destination upload or source deletion. This establishes a
  bounded same-machine loopback safety-cost baseline; no apples-to-apples
  historical metric exists, and the writer archive is explicitly not compared.
- `Tools/Test-TestRunArchive.ps1 -RunPath ...` passed for the four-file curated
  run, whole-inventory validation passed for 525 files, staged and unstaged
  `git diff --check` passed, and `Tools/Get-SpecInventory.ps1 -FailOnFindings`
  reported zero blockers after reconciliation into
  `FileSystem_FtpSftpScp.md`, `FileSystem_FileOperations.md`, and
  `Plugins_VirtualFileSystem.md`.

Accepted constraints to preserve:

- `ResolveCurlSourceSizeCommitment(...)` distinguishes targeted SIZE/stat
  unavailability from hard failures. The focused fake matrix confirms both the
  fallback and hard-failure paths; preserve that distinction during spec work.
- `PreflightDirectorySourceSizes(...)` intentionally performs an extra read pass
  before destructive recursive Move. Retain the before-mutation guarantee and
  include this command/performance cost in the Package B evidence.
- Temporary FTP stderr/control-stream diagnostics were removed. Keep the
  production capture status-only and content-free; failure assertions may report
  HRESULT and command counts but must never print credentials or server payloads.

Exact resume point: begin Package E. Packages A-D are accepted below and must not
be reopened without contradictory evidence. Re-read L4-TERM-01/02/03 and audit the
current Terminal Kitty capture/conversion/publication implementation before adding
the bounded keyed visible-delta cache, all-placement admission/accounting, and
high-cardinality performance proof.

#### Package C completion state (accepted 2026-08-16)

Package C is complete and accepted. This section records the selected design,
runtime RED proof, implementation, focused Debug/Release GREEN proof,
authoritative-spec reconciliation, and archived same-machine evidence.

Verified dependency and command surface:

- The installed x64 dependency manifest contains curl `8.21.0` with the SSH
  feature through libssh2 `1.11.1`.
- `RemoteRename(...)` in `FileSystemCurl.Imap.cpp` sends FTP `RNFR`/`RNTO` and
  sends libcurl quote command `rename <source> <destination>` for SFTP/SCP.
- FTP RFC 959 defines the rename command sequence but does not provide this
  plugin with an atomic create-if-absent publication, an immutable object token,
  or a conditional delete/restore operation:
  <https://datatracker.ietf.org/doc/html/rfc959>.
- In curl `8.21.0`, the SFTP `rename` quote implementation calls
  `libssh2_sftp_rename_ex(...)` with `OVERWRITE | ATOMIC | NATIVE`; the current
  libcurl quote surface therefore does not let this plugin request no-overwrite:
  <https://raw.githubusercontent.com/curl/curl/curl-8_21_0/lib/vssh/libssh2.c>.
  The underlying libssh2 API can omit `OVERWRITE`, but that control is not
  exposed by the current plugin/libcurl command path, and `ATOMIC`/`NATIVE` are
  request flags rather than a complete immutable rollback contract:
  <https://libssh2.org/libssh2_sftp_rename_ex.html>.
- `CURLOPT_QUOTE` passes FTP commands directly and documents only the supported
  SFTP command grammar; a command name alone is not proof of conditional
  publication or conditional recovery:
  <https://curl.se/libcurl/c/CURLOPT_QUOTE.html>.

| Current provider surface | Atomic publish-if-absent | No-replace rename control | Immutable object/revision for rollback | Conditional delete/restore | Safe recursive rollback by path | Package C disposition |
|---|---|---|---|---|---|---|
| FTP `RNFR`/`RNTO` | Not proven | Not exposed/proven | None in current plugin contract | None | No | Fail closed |
| SFTP libcurl quote `rename` | No through current curl `8.21.0` quote path | No; current curl supplies `OVERWRITE` | None in current plugin contract | None | No | Fail closed |
| SCP using the plugin's non-FTP quote path | Not proven | Not exposed/proven | None in current plugin contract | None | No | Fail closed |

Saved engineering recommendation: select architecture option 2 for the current
dependency/ABI surface. FTP/SFTP/SCP no-overwrite Copy, Move, and Rename must
fail before endpoint mutation; recovery that cannot prove immutable ownership
must preserve source, destination, backup/staging artifacts, and concurrent
directory children, then return partial-state truth. Do not add a second probe,
an in-process mutex, an unconditional rename, overwrite authorization for a
supposedly absent destination, or delete/recursive rollback by pathname.

The product/ABI representation decision is closed after tracing the host parser,
task planner, standard File Operations entry points, Batch Rename parser, and
Package A test seams. Capability JSON version 1 already defines the coarse
`operations.copy`, `operations.move`, and `operations.rename` booleans, and the
host rejects false Copy/Move operations before task creation. Batch Rename
already requires `rename: true`; the general File Operations parser currently
omits that field and incorrectly returns `true` for every parsed Rename. Package
C therefore uses no new field and no COM/JSON version change: set all three
booleans to `false` for FTP/SFTP/SCP, add `renameOperation` to the general parser,
and gate `FILESYSTEM_RENAME` on it plus stable path identity. Direct Curl methods
remain responsible for returning `ERROR_NOT_SUPPORTED` before endpoint work,
because capability JSON is UI/planning truth rather than an authorization
boundary. If a later product requirement demands enabled same-provider Curl
Copy/Move/Rename, open a new versioned provider/transport design with a proven
conditional-publication primitive instead of weakening this decision.

Exact unsafe production surfaces to re-open first:

- `PrepareOverwriteTargetForRename(...)`: absence probe followed later by an
  unconditional remote rename.
- `RestoreOverwriteTargetAfterFailure(...)`: deletes the current destination by
  path before restoring the backup.
- `RollbackMovedFileDestination(...)`: deletes the published destination by
  path after source-delete failure.
- `TryRollbackCopiedDirectory(...)`: recursively deletes the destination tree
  by path.
- `PromoteStagedFileToDestination(...)`, `RenameWithOverwriteRollback(...)`,
  single/batch Copy/Move call sites, and `RenameItem(...)`/`RenameItems(...)`:
  all must use the selected fail-closed and preservation result consistently.

Resume order:

1. **Complete:** capability/UI tracing chose existing version-1 booleans with no
   ABI/schema extension; the general Rename gate defect is part of the repair.
2. **Complete:** fake-transport barriers cover the absence/publication/rollback
   race windows, and the deterministic runtime RED proof is saved below.
3. **Complete:** fail-closed no-overwrite behavior, preservation-only uncertain
   recovery, capability truth, the general host Rename gate, focused tests, and
   the complete Debug GREEN matrix are saved on disk.
4. **Complete:** `FileSystem_FtpSftpScp.md`,
   `FileSystem_FileOperations.md`, and `Plugins_VirtualFileSystem.md` carry the
   same contract; targeted contradiction searches found no stale Curl claim.
5. **Complete:** staged/unstaged `git diff --check` passed and
   `Tools/Get-SpecInventory.ps1 -FailOnFindings` exited cleanly.
6. **Complete:** test-enabled x64 Release FileSystemCurl and
   PluginContractTests built with zero warnings/errors. Five independent
   `PluginContractTests.exe --curl-selftests` processes each passed 211/211,
   returned `hr=0x00000000`, and reached the unload quiet point.
7. **Complete:** the four-file same-machine archive under
   `Specs/TestRuns/4cb089111a23/FileOps/20260816_123000_curl_conditional_publication_release/`
   has exactly five rows for each new guard metric and passes targeted plus
   changed-inventory archive validation. L4-CURL-04/05 are closed.

Accepted Release/capture command template (the completed lane used run id
`20260816_123000_curl_conditional_publication_release` and its archive path for
all five processes):

```powershell
$env:RSBuildEnableTests = 'true'
.\build.ps1 -ProjectName FileSystemCurl -Configuration Release
.\build.ps1 -ProjectName PluginContractTests -Configuration Release

$env:REDSALAMANDER_PERF_JSONL_PATH = '<absolute-archive-path>\perf_metrics.jsonl'
$env:REDSALAMANDER_PERF_JSONL_SCENARIO = 'Curl conditional publication and identity-safe recovery contract'
$env:REDSALAMANDER_PERF_JSONL_BUILD = 'Release x64 ENABLE_TESTS'
$env:REDSALAMANDER_PERF_JSONL_BRANCH = 'detached-worktree'
$env:REDSALAMANDER_PERF_JSONL_COMMIT = 'e9c76e59aa2b2993170ea1e6c88e7d15362954c7'
$env:REDSALAMANDER_PERF_JSONL_MACHINE_HASH = '4cb089111a23'
$env:REDSALAMANDER_PERF_JSONL_RUN_ID = '<run-id>'

1..5 | ForEach-Object {
    & .build\x64\Release\PluginContractTests.exe --curl-selftests
    if ($LASTEXITCODE -ne 0) { throw "Curl Release process $_ failed with exit $LASTEXITCODE" }
}
```

The accepted capture was inspected before curating `README.md`, `results.json`,
and `trace.txt`; it has exactly five rows for each new guard metric. Targeted and
changed-inventory archive validation pass. The capture/build variables were
process-scoped and do not persist into Package D.

Useful first retrieval commands (read-only):

```powershell
rg -n 'operations\.(copy|move|rename)|"copy"|"move"|"rename"|GetCapabilities' `
  RedSalamander Common Tests Plugins/FileSystemCurl Specs/Plugins
rg -n 'PrepareOverwriteTargetForRename|RestoreOverwriteTargetAfterFailure|RollbackMovedFileDestination|TryRollbackCopiedDirectory|PromoteStagedFileToDestination|RenameWithOverwriteRollback' `
  Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp
```

Package C deterministic RED evidence captured on 2026-08-16 before production
repair:

- Test-enabled Debug FileSystemCurl built with zero warnings/errors, receipt
  `83edce45475b5e095c94128da15f78a8bd3372af00204ecf0aae9183d52242fb`;
  PluginContractTests built with zero warnings/errors, receipt
  `0c32a15d602c2ba59543b860c5c6976ecebc049a528a05570e4a05ff697de838`.
- `.build\\x64\\Debug\\PluginContractTests.exe --curl-selftests` returned the
  expected failure (`200` passed / `11` failed, `hr=0x80004005`) while still
  reaching the DLL unload quiet point.
- The eleven failures cover concurrent-final preservation after writer
  promotion failure (two assertions), no-overwrite single/batch Copy and Move
  (four), no-overwrite single/batch Rename (two), post-publication Move source
  delete failure and concurrent replacement preservation (two), and recursive
  concurrent-child preservation (one). This is the required runtime RED proof;
  the subsequent `211/0` GREEN run turned these same assertions green without
  removing their deterministic server-side injection barriers.

Package C implementation and focused GREEN evidence saved on 2026-08-16:

- `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp` validates the options
  prefix and rejects no-overwrite single/batch Copy, Move, and Rename with
  `ERROR_NOT_SUPPORTED` before settings or endpoint work. Explicitly
  overwrite-authorized direct calls remain a non-advertised compatibility path.
- Uncertain recovery is preservation-only. Rename/promotion failure no longer
  deletes the current final or restores a backup by pathname; Move source-delete
  failure no longer deletes the published destination; recursive failure no
  longer deletes a destination tree that may contain concurrent children.
  Partial-state truth is returned when ownership cannot be proven, while a
  successful publication is not changed into failure solely by safe retained
  cleanup debt.
- `Plugins/FileSystemCurl/FileSystemCurl.h` and
  `FileSystemCurl.Shared.cpp` advertise `operations.copy`, `operations.move`, and
  `operations.rename` as false for FTP/SFTP/SCP and publish
  `caseOnlyRename: notApplicable`.
- `RedSalamander/FolderWindow.FileOperations.cpp` parses
  `operations.rename`; same-provider `FILESYSTEM_RENAME` now requires that
  capability and stable path identity. The focused provider matrix validates
  Copy/Move/Rename rejection and retained Delete support in
  `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Phases05_06.cpp`.
- `Plugins/FileSystemCurl/FileSystemCurl.SelfTest.Ftp.cpp` retains deterministic
  barriers for late destination creation, post-publication replacement,
  source-delete failure, concurrent recursive children, and writer-promotion
  failure. It emits `FileOps.Curl.ConditionalPublicationGuard` and
  `FileOps.Curl.IdentitySafeRecoveryGuard` for the mandatory Release archive.
- FileSystemCurl Debug builds after the repair completed with zero
  warnings/errors; receipts include
  `4edfabcefd342b96245bbdb8e4d64d44a43b8738eece07bdef50aa424b435a55`
  and latest metric-enabled
  `d51819ea3d2a36a4722708b187e62fdfc1cd670354613d7cadd17aa0b0096b8d`.
- The integrated test-enabled RedSalamander Debug build completed with zero
  warnings/errors, receipt
  `0bc26eb22a05eb63f098543e87f0262aeb80b41f8531c53d850f0570a1a5d32d`
  and log `.build/logs/msbuild-20260816_091957_921-pid75352-26e12a8f.log`.
- `.build\\x64\\Debug\\PluginContractTests.exe --curl-selftests` now exits `0`
  with `211` passed / `0` failed, `hr=0x00000000`, and a passing unload quiet
  point. The exact host command
  `.build\\x64\\Debug\\RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix`
  also exits `0`.
- Authoritative reconciliation is complete in `FileSystem_FtpSftpScp.md`,
  `FileSystem_FileOperations.md`, and `Plugins_VirtualFileSystem.md`. Targeted
  contradiction searches, staged/unstaged `git diff --check`, and spec inventory
  pass.
- Final test-enabled Release builds completed with zero warnings/errors:
  FileSystemCurl receipt
  `c4383f8a109f7f74fb007bff9862cb7ed6a7ed8ee66b5c4478a359406116456e`
  and PluginContractTests receipt
  `136b4b668cbc0a24bf5f2d222e0c9ae5fa138c9c3d6796d3af17af73c71592db`.
  Five separate Release processes each passed 211/211 with clean unloads.
- Accepted archive:
  `Specs/TestRuns/4cb089111a23/FileOps/20260816_123000_curl_conditional_publication_release/`.
  It contains 175 valid JSONL rows across five processes, raw SHA-256
  `37dfd98b361f4febda93bf856b76b5bd1c4a99e179dafeb8447e219ff1db1550`.
  `ConditionalPublicationGuard` proves six no-overwrite shapes issue zero
  endpoint commands/race injections and return `0x80070032`.
  `IdentitySafeRecoveryGuard` measures 18,000–22,000 us (20,000-us median), 103
  deterministic loopback commands, three preserved artifacts, and
  `0x8007012B`. The archive makes no live-endpoint or speedup claim.

Worktree preservation on resume:

- The repository remains at detached `HEAD`
  `e9c76e59aa2b2993170ea1e6c88e7d15362954c7` with 156 status entries at this
  saved checkpoint. Preserve all existing user/remediation changes; do not
  reset, clean, stash, or create a bulk checkpoint commit.
- Preserve the unrelated `.agents/skills/improve/SKILL.md` modification and
  exclude it from remediation commits.
- Git now reports this plan as `R `: the old-path deletion and all current
  new-path content are staged as one in-place rename, with no unstaged plan diff.
  The old LastFourDays path is absent; do not recreate it or add a compatibility
  copy. If a later session edits this plan, stage only this new plan path to fold
  those edits into the same rename, then inspect `git diff --cached`.
- Keep the intentional untracked outputs
  `RedSalamander/CommandRuntimeState.h`,
  `Specs/Build/ToolchainIdentity.schema.json`, and
  `Specs/Testing/ValidationTerminalResult.schema.json`.
- Package B closeout verification on 2026-08-16 had 145 status entries after its
  four-file archive. Package C added its separate four-file accepted archive,
  producing the then-current 149-entry state; the saved Package D implementation
  brought that checkpoint to 156 entries; Package D's accepted four-file archive
  brings the current checkpoint to 160 entries. There is still one matching WIP plan
  path (this CrossDomain file; no LastFourDays copy), staged and unstaged
  `git diff --check` pass, zero-blocker spec inventory, and whole-inventory
  TestRuns success. Diff output contained only expected line-ending warnings.

#### Package D completion state (accepted 2026-08-16)

Package D is complete and accepted. The production result, focused tests,
authoritative specifications, localization/static contracts, test-enabled Release
five-run capture, archive validation, and final L4-CURL-02/07 audit all agree. Do
not repeat Packages A-D or weaken their accepted invariants without contradictory
evidence.

Saved production implementation:

- `FileSystemCurl.Internal.h` now defines `CurlCleanupDebtKind`,
  `CurlPublicationResult`, and `CurlPublicationAccumulator`. The result carries
  `primaryMutationHr`, `cleanupHr`, `sourceDeletionHr`, `primaryCommitted`,
  `immutableRollbackEligible`, `artifactsPreserved`, `cleanupDebtMask`, and a
  saturating `cleanupDebtCount`. Its merge and `OperationResult()` rules retain
  the first relevant failure, return `S_OK` for a committed destination with
  cleanup-only debt, and return failure/partial truth for primary or
  source-deletion failure. The accumulator merges recursive and concurrent
  worker results atomically.
- `FileSystemCurl.CopyMove.cpp` no longer contains `RemoteRecoveryResult` or
  `CompleteSuccessfulOverwrite`. Structured outcomes now flow through
  overwrite preparation/finalization, staging promotion, server rename,
  `CopyFileViaTemp(...)`, recursive Copy, single/batch Copy/Move, and
  single/batch Rename. Each public operation owns one accumulator and observes
  its aggregate debt once. Committed cleanup-only debt preserves successful item
  callbacks and normal destination notifications; source-deletion failure
  remains `ERROR_PARTIAL_COPY` with source/destination/recovery artifacts
  represented truthfully.
- Retained rollback siblings are recorded as `ERROR_NOT_SUPPORTED` cleanup debt
  and are never deleted by pathname. Destination-tree preservation and failed
  operation-owned staging cleanup are distinct content-free debt kinds. No
  delayed scavenger, prefix cleanup, compatibility delete, or unsupported
  "later maintenance" claim was added.
- `PublishCurlWriterTransaction(...)` now returns the same structured result.
  `TempFileWriter::Commit()` observes cleanup debt after the first publication,
  sends its normal notification once for committed publication, and returns
  immediately on a second successful `Commit()`. Writer metric details were
  changed from `_pluginPath` to fixed content-free text.
- `FileSystemCurl` queries and stores `IHostAlerts`. Its
  `ObserveCurlCleanupDebt(...)` emits one content-free warning, one
  `FileOps.Curl.CleanupDebt` metric with fixed detail, and at most one localized,
  application-scoped, modeless warning per aggregate public operation/first
  writer commit. The user message contains only the retained-item count; it does
  not include endpoint, username, credential, or path content.
- Resource IDs `50009` and `50010` define the cleanup-debt title/message in the
  source RC and all four satellites (`cs-CZ`, `fr-FR`, `ja-JP`, `sk-SK`). Every
  message preserves the positional `{0}` placeholder. The resource localization
  contract passes all 6 tests.

Selected Package D result contract:

1. Introduce one internal structured publication/recovery result shared by all
   affected helpers. It must separately represent primary mutation HRESULT,
   cleanup HRESULT, source-deletion HRESULT, whether primary publication
   committed, immutable rollback eligibility, preserved artifacts, a
   content-free cleanup-debt kind/mask, and debt count.
2. The operation HRESULT remains failure/partial for a failed primary mutation
   or failed post-publication source deletion. A committed destination plus only
   retained cleanup debt remains `S_OK`; writer/item size, callback, normal
   notification, and destination bytes describe the committed result.
3. Aggregate cleanup debt across recursive and batch operations without losing
   the first relevant cleanup HRESULT. Emit one content-free cleanup-debt metric,
   one diagnostic, and at most one localized modeless warning per public
   operation/first writer commit. The warning states that the operation
   completed, gives the retained recovery-item count, and tells the user to
   inspect/remove only identifiable backup or staging items manually.
4. Automatic delayed cleanup, rollback-sibling prefix scavenging, unconditional
   pathname deletion, and claims of future maintenance remain forbidden without
   immutable identity. Package C's preservation-only decision has precedence.
5. For a retained rollback sibling, the deterministic cleanup result is the
   policy failure `ERROR_NOT_SUPPORTED` and zero unsafe delete commands. If a
   fake deletion failure is armed for that sibling, the test must prove it was
   not consumed. Use actual injected delete failure only for a cleanup path that
   is still identity-safe for the operation-owned staging object.

Saved test-first and focused validation state:

- Fake host-alert, operation-callback, and directory-watch observers plus
  writer/single/recursive/batch scenario runners were added before the
  production repair. The accepted deterministic RED was exactly
  `passed=211, failed=6, hr=0x80004005`: one missing cleanup-debt observation
  each for writer, single Copy, recursive Copy, batch Copy, single Move, and
  batch Move. All prior assertions stayed green and plugin unload was clean.
  This RED is behavioral evidence only and must not be archived as an accepted
  run.
- The fixture now has marker-based generated-staging deletion rejection. The
  cleanup-failure control proves verification returns the primary
  `ERROR_PARTIAL_COPY`, the old destination remains intact, one staging sibling
  is retained, one safe `DELE` attempt is rejected, there is one content-free
  alert, and there is no normal notification or rollback-sibling deletion.
- Writer success proves its second `Commit()` performs no second upload,
  notification, warning, or alert. The committed writer plus the single,
  recursive, and batch Copy/Move shapes assert final bytes, recovery siblings,
  source state, item/callback outcome, committed size, exactly-once normal
  notification, and exactly-once aggregate cleanup-debt observation.
- The latest test-enabled Debug FileSystemCurl build completed with zero
  warnings and zero errors. Receipt:
  `3a5ac49657c6df9f5e229a330fc46be7f9a970fda4f4bc331b45b029a1f4ac5a`;
  log: `.build/logs/msbuild-20260816_202617_514-pid78940-c04fba58.log`.
  `.build\x64\Debug\PluginContractTests.exe --curl-selftests` exits `0` with
  exactly `passed=221, failed=0, hr=0x00000000`, followed by a passing plugin
  unload quiet point.
- Upload rejection after deterministic storage, verification, prepare, promotion,
  and failed safe staging-cleanup controls are explicit and green. The upload and
  verification controls retain their primary errors, preserve the old final, emit
  no normal notification, and expose staging debt only when operation-owned safe
  cleanup is rejected. Rollback-delete failure markers are armed for committed
  writer/Copy/Move scenarios and remain unused with zero rollback `DELE` commands.

Package D closure evidence:

1. `FileSystem_FtpSftpScp.md`, `FileSystem_FileOperations.md`, and
   `Plugins_VirtualFileSystem.md` define the same structured-result,
   committed-cleanup-success, source-deletion-partial, aggregate-observation,
   privacy, idempotent-writer, and no-unsafe-cleanup contract.
2. `ResourceLocalizationContracts.Tests.ps1` passes 6/6;
   `Tools/Get-SpecInventory.ps1 -FailOnFindings` reports zero blockers; focused
   `git diff --check` and targeted contradiction searches pass.
3. Final test-enabled x64 Release builds have zero warnings/errors. FileSystemCurl
   receipt `77383e4563a8e026c8fbdbaa1bc03744a4ed105320d4da5898d28b8bc9a03bb9`;
   PluginContractTests receipt
   `2248d99f39c924b815dd9781beab9b03ac470b5398c518ed65f354d73fc5e391`.
4. Five independent Release processes each pass 221/221 and the unload quiet
   point. Accepted archive:
   `Specs/TestRuns/4cb089111a23/FileOps/20260816_202500_curl_cleanup_debt_release/`.
   It contains 350 valid content-free JSONL rows, SHA-256
   `4f8cb519d6de71dd727cbf4e8c0b3311000590434da10958ff1d193968008247`.
   The guard is 59,000–62,000 us (60,000-us median), 320 commands, and nine
   retained rollback items per run; the aggregate observer emits 19 rows carrying
   22 retained-item counts per process.
5. Exact pre-stage archive validation passes for four files, and whole-inventory
   validation passes for 533 files. L4-CURL-02/07 are complete.

Exact resume sequence: begin Package E. Re-read L4-TERM-01/02/03, inspect the
current Kitty capture/storage/conversion/publication paths and existing tests,
capture the required deterministic RED proof, then implement the bounded keyed
visible-delta cache and all-placement accounting with Release evidence. Do not
move this plan to `Done`; Packages E-K and final Fresh Full remain mandatory.

Checkpoint integrity:

- Saved at detached `HEAD`
  `e9c76e59aa2b2993170ea1e6c88e7d15362954c7` with 160 porcelain entries.
- The original plan remains a single staged in-place rename; there is no
  `CodeReview_LastFourDays_Remediation_2026-08-10.md` duplicate and no new
  handoff directory.
- Package D added only the four curated files in its accepted TestRuns archive;
  they are staged intentionally. Preserve the three already-known intentional
  untracked files listed in the earlier worktree checkpoint and the unrelated
  `.agents/skills/improve/SKILL.md` modification.
- Package D's saved edits are confined to the existing FileSystemCurl source,
  header, selftest, and five RC files plus this original plan. No commit, stash,
  reset, cleanup, duplicate plan, or handoff folder was created.
- Focused `git diff --check` over FileSystemCurl, the three authoritative specs,
  the archive, and this plan passes; only expected LF-to-CRLF warnings are printed.

Do not move this plan to `Done` at Package D. Packages E-K and the final Fresh
Full closeout remain mandatory.

#### Package E continuation checkpoint -- deterministic RED saved (2026-08-16)

This is the exact later-session handoff for L4-TERM-01/02/03. The Package E
audit and deterministic test-first RED are complete. Only the focused test seam
and failing runtime fixtures have been added; no Package E production behavior,
runtime overlay, authoritative-spec contract, governed inventory change, GREEN
proof, or accepted evidence archive exists yet. Continue from the RED rather
than reproducing or weakening it, and do not infer that the selected
architecture below is already implemented.

Current verified defect state:

- `Terminal::captureKittyGraphicsLocked(...)` submits image work only when
  Ghostty's aggregate image-storage generation differs from
  `_kittyTargetStorageGeneration`. It recomputes visible placements, but copies
  only visible/nonvirtual image data and uses `std::ranges::any_of(work.images)`
  for deduplication. Geometry-only scrolling can therefore reveal a different
  valid image while the storage generation is unchanged, without submitting
  that image for conversion.
- `Terminal::consumeKittyReady()` replaces the full `_kittyImages` vector rather
  than merging keyed results. A missing, oversized, stale, or failed image can
  consequently invalidate unrelated usable images.
- `Terminal::drawKittyLayer(...)` performs a linear
  `std::ranges::find_if(_kittyImages)` for each placement and repeats that work
  for the three Kitty layers. Capture deduplication and drawing are quadratic at
  high cardinality.
- The current 4,096 guard tests `placements.size()` before accepting another
  visible/nonvirtual placement. Off-screen and virtual iterator entries are not
  represented in that size, and Ghostty's `ImageStorage.PlacementMap` has no
  product hard count/allocated-byte admission limit. The host check therefore
  does not bound iterator work or engine-retained placement state.
- Existing ceilings remain non-negotiable: 64 MiB source, converted, storage,
  and per-frame upload budgets; 16 million pixels; latest-only publication; and
  at most one bitmap upload per frame. Do not raise a limit to close these
  findings and do not introduce an all-image copy.

Relevant current files and clean/dirty boundaries:

- Host capture/publication/draw state is in `Plugins/Terminal/Terminal.cpp` and
  `Plugins/Terminal/Terminal.h`; conversion is in
  `Plugins/Terminal/KittyImagePipeline.{h,cpp}`. `Terminal.h` now has only the
  Package E test-method declaration; both pipeline files remain clean.
- `Terminal.cpp` already has an unrelated 51-addition/1-deletion working-tree
  diff for `<clocale>`, Unicode/Turkish command-suggestion selftests,
  `GetActionState` caller-size validation, and strict callback/cookie pairing.
  Package E adds the host visible-delta RED fixture, the 4,097-placement raw
  runtime RED, and the call from the existing exported selftest. Preserve all
  of these edits and build on that file; never replace it from `HEAD`.
- Related existing dirty command-surface work also spans
  `TerminalCommandSurfaceModel.cpp`, `FloatingTerminalWindow.{h,cpp}`, the
  authoritative Terminal spec, and
  `Tools/Tests/TerminalPluginSourceContracts.Tests.ps1`. Reconcile rather than
  overwrite those changes.
- The product Ghostty builder is `Tools/Build-GhosttyTerminalRuntime.ps1`; the
  active overlay is
  `Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch`, currently
  SHA-256
  `bc9ac4cbae85d8a491e32471096e43f2042c03f4558728bfb067ddb8c36cf2ac`.
  The pinned inspected source is
  `.build/TerminalEngine/Sources/ghostty-b2d44625908eb56aee299d8511d803bc11ab79fc`
  at commit `b2d44625908eb56aee299d8511d803bc11ab79fc`. Its two already-deleted ignored
  fuzz-corpus files are cache state, not remediation targets. Do not edit the
  historical `External/TerminalEngine/CandidatePatches` evaluation.
- The pinned C ABI currently ends Terminal options at 30 and Kitty graphics data
  at 2. The intended private-overlay additions are a versioned placement-limits
  Terminal option at 31 and Kitty graphics queries for retained count, retained
  allocation bytes, count limit, and byte limit at 3 through 6. Reconfirm the
  exact pinned declarations while applying the overlay; do not edit generated
  product headers by hand.

Selected host repair shape:

1. Key image state by `{imageId, imageGeneration}` with a hash index used for
   missing detection, work deduplication, ready merge, draw lookup, and LRU
   eviction. Do not add a second persistent vector requiring linear scans.
2. Recompute geometry and count every Ghostty placement iterator entry before
   virtual/off-screen filtering on every capture, even when aggregate storage
   generation is unchanged. The aggregate storage generation is a content stale
   fence; the per-image generation is pixel identity.
3. Use a bounded two-pass visible-working-set capture under the terminal mutex:
   collect visible placements and unique keyed image handles first, then copy
   only currently visible keys that are not ready. When a newer missing key
   supersedes pending latest-only work, include all still-visible pending keys in
   the replacement request so they cannot be stranded.
4. Copy Ghostty-owned image bytes while holding its mutex, but perform pixel
   conversion outside it. A same-storage-generation missing key must submit a
   bounded delta request.
5. Merge accepted results into the keyed cache. Pin current-visible entries;
   evict nonvisible entries by LRU under the existing byte ceiling. A missing,
   oversized, rejected, stale, or upload-failed key degrades only its placements
   and must not clear unrelated ready keys. Content mutation/deletion must make
   obsolete keys evictable so stale content cannot remain live indefinitely.
6. Remove the broad all-or-nothing storage-generation draw gate. Each placement
   draws only when its exact key is ready. Preserve latest-only cancellation and
   the one-upload-per-frame rule.
7. Extend runtime statistics with actual source, active, ready, converted,
   cached, pinned, evicted, uploaded, missing, rejection, total-iterated,
   visible, engine-retained-count/bytes, and Ghostty-lock-duration evidence.
   Runtime assertions are authoritative; source-contract checks are secondary.

Selected Ghostty runtime-bound direction, subject only to compile-time ABI
details discovered while applying the pinned overlay:

- Add a versioned Kitty-placement-limits option to the private product ABI and
  expose actual placement count, allocated bytes, count limit, and byte limit
  through terminal graphics data. Apply the configured limits to all screens.
- Keep hard product defaults of no more than 4,096 placements and a finite
  placement-map byte budget; callers may lower, never raise, those hard limits.
  Before inserting a new placement, reject when the next count or predicted
  retained map allocation exceeds its ceiling. Replacing an existing placement
  does not consume another count and must deinitialize the replaced value.
- Compute retained placement-map bytes from the pinned Zig hash-map allocation
  layout/capacity, not from only `count * sizeOf(value)`. Rejection must be
  deterministic and leave count/retained bytes unchanged; Ghostty's existing
  error encoding can report the failed command without retaining it.
- Configure/query the runtime limit from the real product host and exercise it
  through the actual private runtime. If the pinned ABI makes this exact shape
  invalid, document the explicit alternative in this plan and the authoritative
  Terminal spec before implementation; do not silently fall back to a host-only
  iterator guard.

Deterministic RED to add before production repair:

1. Under `ENABLE_TESTS`, add `Terminal::DebugRunKittyCacheSelfTests(...)` and call
   it from the existing exported `RedSalamanderTerminalDebugSelfTests` entry
   point. Initialize the real Terminal engine, transmit two explicit one-pixel
   Kitty images separated by enough scrollback, publish the initially visible
   image, then scroll to reveal the second image without mutating storage.
   Record/assert the unchanged aggregate storage generation and the absent newly
   visible keyed image. The current implementation must fail this exact runtime
   assertion before the repair; after the repair it must publish/merge the second
   image while retaining the first image identity. Extend the same fixture for
   resize/reflow and deletion/stale eviction.
2. In the existing raw `TerminalRuntimeLoader` selftest fixture, transmit one
   one-pixel image and 4,097 explicit placements, including visible, off-screen,
   virtual, and repeated-image cases. The pre-repair runtime must demonstrate
   retained placement state beyond the intended 4,096 ceiling. After the private
   ABI repair, assert boundary success, `+1` rejection, unchanged count/bytes on
   rejection, actual limit/count/byte queries, and clean hostile close/unload.
3. Add a high-cardinality/repeated-image runtime case proving keyed unique work
   and lookup counts rather than timing alone. Capture total iterated, visible,
   unique, rejected, retained bytes/count, lock duration, and frame/upload work.
4. Build and run the RED with:

   ```powershell
   .\build.ps1 -ProjectName Terminal -Configuration Debug
   .\build.ps1 -ProjectName PluginContractTests -Configuration Debug
   .\.build\x64\Debug\PluginContractTests.exe --terminal-selftests
   ```

   Save the exact receipts and pass/fail counts in this subsection. Do not archive
   a RED run as accepted evidence. If the runner syntax differs, use the existing
   focused Terminal entry point rather than creating a new public tool.

Package E deterministic RED captured on 2026-08-16 before production repair:

- The test-only host fixture uses the actual `Terminal` object and pinned private
  runtime. Its first viewport capture is a valid control: diagnostic flags are
  `63`, meaning scrollback, capture, submission, nonzero storage generation,
  image 101 copying, and image 101 placement visibility all succeeded.
- Scrolling to the bottom reveals image 202 while the aggregate storage
  generation remains unchanged. Diagnostic flags are `19`: capture, same
  generation, and visible image-202 placement succeed, but submission and image
  copying are both absent. This is the exact L4-TERM-01 runtime failure.
- A separate real-runtime terminal transmits one one-pixel image and 4,097
  unique explicit placements. The placement iterator returns 4,097 entries, so
  the desired exact 4,096 assertion fails. This is the L4-TERM-02 engine-retention
  RED; it is independent of the host's visible-placement vector.
- The final pre-repair focused run is exactly `passed=61, failed=2,
  hr=0x80004005`. All localization and sized-record contracts remain green, and
  the Terminal DLL reaches its unload quiet point. The two failures are the
  same-generation visible-delta assertion and the 4,096 retained-placement
  assertion; no unrelated Terminal assertion fails.
- The test-first Terminal Debug build completed with zero warnings/errors,
  receipt `f3b014c730ee590e8b34fb3afdb8b92cb0b4fcfd8f422b18a73b5e337adbeeaf`;
  log `.build/logs/msbuild-20260816_205131_548-pid69648-0374f131.log`.
  The current-source PluginContractTests Debug receipt is
  `e74546c4547b9f5aee4efbd82f77952b5e354c88b48e7d1403769df49d43b02a`;
  log `.build/logs/msbuild-20260816_205315_913-pid90396-2a9273b5.log`.
- The exact saved test-first working files hash as follows before production
  repair: `Plugins/Terminal/Terminal.cpp` =
  `e5b0435e6f8ec1340aea37e9dd12362bce0193ad92a4e4de272074dbcb954018` and
  `Plugins/Terminal/Terminal.h` =
  `d26fc1a5627126dce7e5af9b38a67e646fc0e8535d443f154180dcd11da2d5f4`
  (SHA-256). `Terminal.cpp` intentionally includes the earlier unrelated
  command-experience changes described above; do not treat the entire file diff
  as Package E-only.
- The diagnostic JSONL used only to disambiguate the fixture is under ignored
  `.build` state and is not accepted/archive evidence. Do not add a RED archive
  under `Specs/TestRuns/`.

Exact resume sequence from this save:

1. Preserve the two failing assertions and their controls. The temporary
   `terminal.kitty.visible_delta_*_probe` events may be replaced by the durable
   required Package E metrics during GREEN implementation, but must not be
   mistaken for accepted performance evidence.
2. Extend only the active
   `Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch`: add hard
   placement count and actual map-allocation-byte admission, replacement-safe
   ownership, the private limits option, and the four graphics queries. The
   product builder's overlay hash is already part of its cache identity. Avoid
   editing the pinned source cache directly; its two pre-existing deleted fuzz
   corpus files are unrelated cache state.
3. Extend the real runtime loader/host initialization and the raw selftest to
   configure/query the limits and prove 4,096 boundary success, `+1` rejection,
   unchanged retained count/bytes after rejection, repeated-ID replacement, and
   hostile close/unload. This must turn the placement RED green through engine
   admission, not merely an iterator break in the host.
4. Replace the host vector-scan/all-or-nothing generation behavior with the one
   keyed visible-working-set/delta cache specified above. Keep exact-key lookup,
   latest-only replacement, current-visible pinning, nonvisible LRU eviction,
   per-key degradation, 64-MiB/16-million-pixel ceilings, and one upload per
   frame. Extend the runtime fixture for same-generation reveal, resize/reflow,
   deletion/stale eviction, and high-cardinality/repeated-image work.
5. Reconcile the active overlay inventory metadata, focused source/runtime
   contracts, `Specs/Terminal/Terminal_EmbeddedPlugin.md`, and this plan. Then
   build Debug, run the focused GREEN suite, collect five independent
   test-enabled x64 Release processes, curate/validate one content-free Terminal
   archive, and only then close L4-TERM-01/02/03 and proceed to Package F.

Implementation and closeout obligations after the RED:

- Update the active Ghostty product overlay, host ABI declarations/loader, host
  cache/pipeline/draw path, focused runtime tests, and
  `Specs/Terminal/Terminal_EmbeddedPlugin.md` together. Changing the overlay
  invalidates the builder cache identity automatically; verify that the build
  consumes the new hash.
- Because the active file under `Tools/TerminalEngine/Patches/` is governed
  tooling, audit/update its `Tools/tool-inventory.json` metadata/dependency
  closure as needed and run the focused Terminal runtime/source-contract,
  Release-workflow/cache-closure, Tool Inventory, Tooling Governance, and spec
  inventory gates. No new top-level command or README map entry is expected.
- Read the performance-validation instructions before creating evidence. Capture
  five independent test-enabled x64 Release processes with a deterministic
  before/after comparison for reveal correctness, high-cardinality work, retained
  placement bounds, Ghostty lock duration, and frame/upload counts. Curate only
  the accepted content-free archive under `Specs/TestRuns/...` and validate it
  before intentionally staging its exact files.
- Reconcile every L4-TERM-01/02/03 sentence against production, executable tests,
  the Terminal spec, and the archive. Only then mark those checklist rows
  complete and proceed to Package F. Packages F-K plus final
  `.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` remain required
  before moving this single plan to `Specs/Plans/Done/`.

#### Package E acceptance checkpoint -- complete and accepted (2026-08-17)

The production core, high-cardinality correctness proof, current-source Debug
gate, five-process test-enabled Release evidence, authoritative Terminal spec,
governed source contract/inventory, and archive validation for
L4-TERM-01/02/03 are complete. Package E is accepted; do not reopen it without
contradictory evidence.

Implemented runtime placement bound:

- The only durable Ghostty source change remains the existing active overlay,
  `Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch`; no second
  patch or replacement plan was created. Its current SHA-256 is
  `f2aa0f100e6d03a03fab433ee7d5ec72d58fef82663326732eddd6547781999b`.
- The pinned private C ABI now exposes the sized
  `GhosttyKittyPlacementLimits` option at Terminal option 31 and Kitty graphics
  count, exact allocated bytes, count limit, and byte limit at data keys 3..6.
  Null resets the hard defaults; a non-null sized record may lower but not raise
  the 4,096-count or 1-MiB allocation limits, applies atomically across screens,
  and is rejected if existing retained state does not fit.
- Placement admission predicts the pinned Zig hash-map's actual single
  allocation layout/capacity before insertion. The 4,096 boundary succeeds;
  `+1` is rejected without changing count/bytes; explicit-ID replacement
  releases the previous placement pin, preserves count, and cannot trigger a
  grow-before-replacement allocation.
- The overlay's focused Zig tests pass with
  `zig build test-lib-vt --seed 0 -j1 -fno-incremental -Demit-lib-vt
  -Doptimize=Debug -Dsimd=false`. The product offline builder consumed the new
  overlay and passed for x64. The raw DLL SHA-256 is
  `5b18a0b2f3775baac7ec6a1beab98a331855d2bfa00ed60c0a2dd3450fd16e8d`,
  the normalized diagnostic hash is
  `9443d9b916e6bfa947df0d6527dc450008f72e2c1f75f4f8af12f2b09134e82d`,
  and the 180-export set hash remains
  `596894f7af28c573cf8a45429c8b800f92a52708b6eb94857cbc81ea118868dd`.
  `git apply --check --whitespace=error-all` passes against pinned source
  `b2d44625908eb56aee299d8511d803bc11ab79fc`.
- The original pinned source cache still has only its two unrelated deleted
  fuzz-corpus entries. The ignored authoring worktree at
  `.build/TerminalEngine/PackageEOverlay` is temporary implementation scratch;
  it is not a durable source, handoff, plan, or product patch. Continue from the
  active overlay file above.

Implemented host keyed cache:

- `Terminal` now keys one `std::unordered_map` by
  `{imageId,imageGeneration}`. That index owns missing/deferred/pending/ready/
  rejected state, visible epoch, CPU BGRA data, bitmap, exact work request, and
  eviction data. Capture deduplication, result merge, and draw lookup no longer
  use persistent vector scans.
- Every capture recomputes placement geometry and counts iterator entries before
  filtering. It queries and enforces the engine-retained count/allocated-byte
  contract, discovers missing exact visible keys even when storage generation is
  unchanged, copies Ghostty-owned bytes only under `_terminalMutex`, and submits
  conversion outside the lock. Aggregate storage generation fences stale work;
  per-image generation remains pixel identity.
- Ready results merge per exact key. Current-visible ready bytes are pinned;
  nonvisible entries are sorted once by visible epoch and evicted first when the
  64-MiB converted/cache ceiling needs room. Intrinsic invalid/oversized keys are
  rejected individually and aggregate-budget overflow is deferred individually;
  neither path clears unrelated ready images. Storage mutation purges obsolete
  nonvisible keys, while device loss retains CPU data and drops only bitmaps.
- The broad generation-wide draw gate is removed. Each placement performs one
  keyed lookup and draws only its exact ready image. Existing 64-MiB source,
  converted, storage, and upload limits, 16-million-pixel ceiling,
  latest-only replacement, and one-upload-per-frame rule remain unchanged.
- Pipeline statistics now expose queued, active-source, active-converted, ready,
  resident, and high-water bytes. Bounded frame metrics also expose total and
  visible/off-screen/virtual placements, unique visible keys, retained
  placement count/bytes, source/cache/pinned/evicted bytes, missing/deferred/
  rejected keys, exact image-map lookups, frame upload count/bytes, and the
  existing capture/submit/convert/upload durations.
- The real-runtime high-cardinality fixture retains 1,024 off-screen and 1,024
  virtual references to one repeated one-pixel image plus 1,024 visible distinct
  one-pixel images and 1,024 visible repeated-image placements, then rejects
  placement 4,097. Capture asserts exactly 4,096 iterated/retained placements,
  598,040 engine bytes, 2,048 visible drawable placements, 1,025 exact visible
  keys/work items, and 3,075 copied source bytes. A second unchanged-geometry
  capture submits no work and proves 4,100 cached/pinned BGRA bytes with zero
  missing/deferred/rejected keys.
- The test-only model of the removed vector algorithm uses the runtime's actual
  unordered iterator result, asserts order-independent exact bounds, and records
  529,915 capture plus 530,940 draw comparisons. The three production draw
  layers perform exactly 2,048 keyed lookups. Zero uploads are explicitly
  CPU-only evidence and make no L4-TERM-04 product upload claim.

Focused evidence saved at this checkpoint:

- An intermediate run after only the runtime placement repair was exactly
  `passed=62, failed=1`; the retained-placement RED was green and only the
  same-storage-generation reveal RED remained. Terminal Debug receipt
  `ca74173f36a30d2925d88c5cc9fa781f2ddb93bbcc1b1cf41bd0a4a2f8db5710`
  and PluginContractTests receipt
  `63228bc167dbbc5eb44c3261111c25475c960c8d802bc26438869883db0a6b3c`
  both had zero warnings/errors.
- The pre-high-cardinality Terminal Debug build completed with zero warnings and
  zero errors. Receipt:
  `f40a9a7a8a5b8ba35561000649f951d6738c7cfe12e8991ad3cdaef7b0c35620`;
  log: `.build/logs/msbuild-20260817_094049_424-pid43024-e57e5f35.log`.
- The last PluginContractTests Debug build completed with zero warnings and zero
  errors. Receipt:
  `93d436c3ac0f23f3c055edde0593a5bde9042c411d03179bf54f779de84a8f85`;
  log: `.build/logs/msbuild-20260817_094224_422-pid22556-e15e4bb6.log`.
- The final current-source Terminal Debug build completed with zero warnings and
  zero errors. Receipt:
  `4bb3ae276c273677621e578e5cdb088c7c7f7e0f0a1a1606c77a9cc2ecccddeea`;
  log: `.build/logs/msbuild-20260817_101910_425-pid69920-db1b32e5.log`.
  `.build/x64/Debug/PluginContractTests.exe --terminal-selftests` then passed
  exactly `passed=69, failed=0, hr=0x00000000`; localization, sized-record
  contracts, runtime placement boundary/replacement/`+1`, same-generation A-to-B
  reveal, resize/reflow, deletion/stale purge, mixed 4,096-placement/1,025-key
  work, ready-cache reuse, comparison bounds, and unload quiet point are green.
- The final test-enabled Release x64 Terminal build completed with zero warnings
  and zero errors. Receipt:
  `4d6f3e62c66f9d64981f3a9f71a9aaad439f6095fd0c716a4a448bb9f1e244d18`;
  log: `.build/logs/msbuild-20260817_102054_210-pid77792-364b88e7.log`.
  Five independent processes each passed `69/0`, exited 0, emitted no stderr,
  and reached the unload quiet point.
- The accepted content-free Release archive is
  `Specs/TestRuns/4cb089111a23/Terminal/20260817_102300_kitty_visible_delta_release/`.
  It contains 445 raw JSONL rows; SHA-256
  `b89a35e4b44a8d019efae590ea27cce8d4b92430cd6b73f1985db6bc1a8aff61`.
  Explicit `Test-TestRunArchive -RunPath` passes for all four files, the
  whole-inventory validator passes for 537 files, and
  `Show-PerfRuns -FailOnQuality -MinimumSamplesForP95 5` passes all 39 metric
  families. Missing-key capture is 414-594 us (454-us median); ready-cache
  recapture is 205-230 us (219-us median). Paired legacy-versus-keyed draw work
  is 530,940 comparisons versus 2,048 lookups; median lookup time is 117 us
  versus 24 us. The maximum-size conversion median is 29,098 us and resident
  high-water remains exactly 134,217,728 bytes.
- `Tools/Tests/TerminalPluginSourceContracts.Tests.ps1` passes `28/0` and now
  guards the exact-key `std::unordered_map`, `try_emplace`/`find` behavior, the
  absence of the old persistent vector scans, overlay option 31/data keys 3..6,
  hard 4,096/1-MiB limits, and the expanded metric family.
- The surrounding governance suites pass: ReleaseWorkflowPolicy `18/0`,
  ToolInventory `17/0`, and ToolingGovernance `12/0`.
- `Tools/Get-ToolInventory.ps1 -Validate` exits 0. `Tools/Get-SpecInventory.ps1
  -FailOnFindings` exits 0 with `Blocking findings: 0`.
- `Tools/tool-inventory.json` now describes the active overlay's placement-count
  and exact-allocation ABI, names `Plugins/Terminal/Terminal.cpp` as a consumer,
  and names the focused Terminal source-contract suite. The authoritative
  `Specs/Terminal/Terminal_EmbeddedPlugin.md` now owns the all-placement
  admission/accounting, exact-key cache, geometry/content generation, LRU and
  pinned-entry, per-key degradation, one-lookup draw, and device-loss contracts;
  it also labels the missing 2026-08-09 archive as non-authoritative and points
  to the surviving 2026-08-09_165700 evidence.
- Current SHA-256 identities are:
  `Plugins/Terminal/Terminal.cpp` =
  `e5e32dcd01d56fba0a1315ec3bc85f3f6e03989e8a3c5b896e7ccdc2e4c5cdce`,
  `Plugins/Terminal/Terminal.h` =
  `5b376045ee9b5a2274e5e0c133a7d013389206ce71489c625f8ca6248e6d5aeb`,
  `Plugins/Terminal/TerminalKittyImagePipeline.cpp` =
  `9a75439c053cf36683a8795813eb3c08f1a0a44c0b0c0c887e72fb209d17f5e2`,
  and `Plugins/Terminal/TerminalKittyImagePipeline.h` =
  `663d59f79321531f710c5b84db7e845fbfcdbcb64e1613db785badbf145a52e9`.
  Supporting identities are: active Ghostty overlay =
  `f2aa0f100e6d03a03fab433ee7d5ec72d58fef82663326732eddd6547781999b`,
  Terminal source contracts =
  `92453955086b5d11bbc85f861aa94087e34ff097f38bfe6f494f6b65b535c439`,
  tool inventory =
  `c37d1330ba484b4613bac2a75defe6a928e6b358a4ccdafab274b7bb710576de`,
  and authoritative Terminal spec =
  `44d1bb842895a759dd557d747af62ead69609e5c4704192f5fca4281a6a9da67`.
  Unstaged and staged `git diff --check` both pass with only expected
  LF-to-CRLF notices.

Package E acceptance and exact next sequence:

1. L4-TERM-01/02/03 are complete and accepted. Preserve the validated archive,
   current Terminal spec, one active Ghostty overlay, `69/0` runtime suite, and
   `28/0` structural guard; do not reopen Package E without contradictory
   evidence.
2. Begin Package F by re-reading L4-TERM-04/05/07 and the corresponding Terminal
   spec/runtime paths. Capture deterministic RED evidence for bitmap upload
   disposition/recovery, active-conversion close latency, and ordinary resize
   resource reuse before changing production.
3. Close Package F with focused Debug/runtime tests, test-enabled Release
   evidence, authoritative spec reconciliation, and archive validation. Then
   continue Packages G-K and final Fresh Full before moving this plan to `Done`.

Checkpoint integrity at this save remains detached `HEAD`
`e9c76e59aa2b2993170ea1e6c88e7d15362954c7` with 168 rename-aware porcelain
entries (169 at Git's default 50% rename threshold because this heavily expanded
plan is now 48% similar to its `HEAD` source). `git status
--short --find-renames=1%` and `git diff --cached --find-renames=1% --summary`
both show the one staged in-place rename; the old filesystem path remains absent.
Both staged and focused unstaged `git diff --check` pass. Do not create another
handoff artifact.

#### Package F acceptance checkpoint -- complete and accepted (2026-08-17)

The production implementation, deterministic RED-to-green tests, Terminal and
performance contracts, current-code Debug gate, five-process test-enabled
Release evidence, and archive/governance validation for L4-TERM-04/05/06/07 are
complete. Package F is accepted; do not reopen it without contradictory
evidence.

Preserved RED and baseline evidence:

- The source-contract RED was exactly `passed=28, failed=2`: ordinary `WM_SIZE`
  still discarded all device resources, and the authoritative Terminal spec
  referenced missing/stale evidence paths. Both tests are now green in the last
  source-contract run, which passed `30/0`.
- The runtime RED was exactly `passed=72, failed=4,
  hr=0x80004005`. The four failures were one-shot retry scheduling, persistent
  OOM latching, recreate-target post-frame recovery, and resize-storm bitmap
  retention. The test-only active 4,096-by-4,096 conversion-close case already
  passed in that RED run.
- The ignored RED diagnostic is
  `.build/codex-runs/terminal-package-f-red-20260817_104500/`; its performance
  JSONL SHA-256 is
  `17975b9cd6f85fc8ffa7f88bdb4e6c73174ae26dd97fdfd21dcc4c731de2659`.
  It recorded a 432,510-us Debug active close with 67,108,864 source bytes and
  67,108,864 converted bytes. Recovery counters were correctly zero before the
  repair. This diagnostic is not an accepted `Specs/TestRuns/` archive.
- The Package F pre-repair test-seam builds completed with zero warnings/errors:
  Terminal receipt
  `7f07a67b6f0bb7f9cb87c4d2a241bb17b150b71cfc6e6511761d7e8e49625e8b`
  (`.build/logs/msbuild-20260817_104158_017-pid66280-337cda76.log`) and
  PluginContractTests receipt
  `dda49aae489cf944128aa4d5f19ed2cdbc2e6cd5ce40c4048a52a8a7b2800164`
  (`.build/logs/msbuild-20260817_104358_193-pid28276-bd713a36.log`).

Implemented behavior:

- Bitmap upload now returns a structured disposition plus `HRESULT`. Only
  `HRESULT_FROM_WIN32(ERROR_RETRY)` and `D2DERR_DISPLAY_STATE_INVALID` are
  transient; retry delays are 16 ms then 64 ms, and the third failed attempt
  latches a stable per-image/per-device failure. `E_OUTOFMEMORY` and every other
  non-approved failure latch immediately, preventing paint-spin retry loops.
- `D2DERR_RECREATE_TARGET` records a recreate disposition and performs device
  discard/recreation only after the frame. A new device generation makes an
  exact image eligible again. Repaint requests and retry-timer arming are
  centralized and measured.
- The shared retry timer tracks the earliest due image. A pending image whose
  deadline is still in the future re-arms the timer without incrementing its
  failure/schedule count. The final test in `Terminal.cpp` exercises two images
  with different deadlines and guards this edge. Its first version used a
  forbidden linear `_kittyImages` scan, which the source-contract gate rejected
  at `29/1`; the repaired test derives a known fixture key and uses exact map
  lookup. Current runtime and source-contract gates are green.
- Ordinary `WM_SIZE` calls `ID2D1HwndRenderTarget::Resize` through
  `handleResize`/`resizeRenderTargetToClient` and retains compatible Kitty
  bitmaps. DPI/theme/device-loss paths retain the full-discard contract. Resize
  count, recreate count, retained bitmap count, and retained bytes are emitted.
- The active-close fixture enters a real maximum-size conversion before calling
  `Stop()`, then records `terminal.kitty.active_close_us` and proves the queues,
  active bytes, ready bytes, worker count, and notification state drain. The
  production conversion loop remains non-cooperative. The accepted Release
  maximum is 51,873 us, so row/chunk stop polling is not required; the 5-second
  executable safety bound remains.
- `Specs/Terminal/Terminal_EmbeddedPlugin.md` owns the structured upload,
  backoff/latching, post-frame recovery, resize-retention, active-close, metric,
  and executable evidence-link contracts. The preliminary missing Kitty
  narrative and a second nonexistent Commands evidence bullet were removed.
  `Tools/Tests/TerminalPluginSourceContracts.Tests.ps1` now verifies every
  authoritative `Specs/TestRuns/<machine>/<area>/<run>` path actually exists.
- `Specs/Testing/Testing_PerformanceValidation.md` requires the Package F metric
  family and real Kitty-producing lifecycle plus a real Direct2D render target;
  CPU-only fixtures cannot substantiate upload or resize behavior.

Final GREEN and Release evidence:

- The final code-bearing Terminal Debug build completed with zero warnings and
  zero errors. Receipt:
  `567daeb203f8c3d0ef9dc46c0307a5da49232fc6c8f3a7e827c6814e3a6619fa`;
  log `.build/logs/msbuild-20260817_110914_708-pid1252-33199b69.log`.
  `.build/x64/Debug/PluginContractTests.exe --terminal-selftests` passed exactly
  `passed=77, failed=0, hr=0x00000000` and reached the unload quiet point.
- An earlier Debug diagnostic at
  `.build/codex-runs/terminal-package-f-green-debug-20260817_105200/`
  recorded active close 418,905 us, retry schedules 1, stable failures 1,
  device recoveries 1, repaint requests 2, resize count 32, resize recreations
  0, retained bitmaps 1, and retained bitmap bytes 4.
- Terminal source contracts pass exactly `30/0`, including the executable
  evidence-link inventory. The transient `29/1` result after adding the final
  two-image test proved that the no-linear-map-scan contract remains effective;
  no production behavior failed.
- The final test-enabled Release Terminal build completed with zero warnings and
  zero errors. Receipt:
  `b42be952023a1337233005cb8b08d2fcc4aaabc2817e5e8fe2f1392cf56197d0`;
  log `.build/logs/msbuild-20260817_111111_827-pid57648-357428ce.log`.
  PluginContractTests Release also completed with zero warnings/errors, receipt
  `2c5109d7e2bacdf4c662f999503bf32ab7476e21e4c6679ac69ec3747ab8c35f`;
  log `.build/logs/msbuild-20260817_111256_629-pid50776-60841c80.log`.
- Five independent Release processes each passed `77/0`, exited 0, emitted no
  stderr, and reached the unload quiet point. Active close was 51,873, 44,907,
  43,229, 39,514, and 44,172 us: 39,514-us minimum, 44,172-us median, and
  51,873-us maximum. Every process drained queued/active/ready bytes and workers
  to zero and suppressed notification, so L4-TERM-05 requires no production
  stop-polling change.
- Every Release process emitted eight real-render upload rows in the exact
  sequence `ERROR_RETRY`, success, `ERROR_RETRY`, `ERROR_RETRY`, success,
  `E_OUTOFMEMORY`, `D2DERR_RECREATE_TARGET`, success. Final aggregates were one
  retry schedule, one stable failure, one device recovery, and two repaint
  requests. Successful uploads were 12-378 us.
- Each process performed 32 ordinary real render-target resizes and retained the
  same one-pixel bitmap/four bytes with zero recreation or reupload. Across 160
  resize samples, duration was 264-1,081 us (344.5-us median).
- The accepted content-free archive is
  `Specs/TestRuns/4cb089111a23/Terminal/20260817_111500_kitty_recovery_resize_release/`.
  Its JSONL contains 710 rows/338,024 bytes and has SHA-256
  `eb44ec281e2d92f2b0352e6e615c706a963cc54f716a1808bf1c60c341a0891b`.
  Explicit pre-stage validation passes for all four files, whole-inventory
  validation passes for 541 files, and `Show-PerfRuns -MinimumSamplesForP95 5
  -FailOnQuality` passes all 50 metric families.

Package F acceptance and exact next sequence:

1. Preserve the RED, final `77/0` runtime, `30/0` source contract, structured
   recovery semantics, non-cooperative-but-bounded active close, resize-retained
   bitmap behavior, authoritative Terminal contract, and validated archive. Do
   not reopen L4-TERM-04/05/06/07 without contradictory evidence.
2. Run the surrounding ReleaseWorkflowPolicy, ToolInventory,
   ToolingGovernance, tool/spec inventory, archive, and diff gates recorded by
   this package checkpoint; repair any regression before entering Package G.
3. Begin Package G at L4-FILEACTION-03/04/05/07. Capture its deterministic RED
   for private-root exclusive creation, strict record grammar, bounded streamed
   writing, canonical quoting, and exact consumer bytes before production edits.
4. Packages G-K and final
   `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` remain mandatory
   before moving this single plan to `Specs/Plans/Done/`.

Checkpoint integrity:

- Detached `HEAD` is
  `e9c76e59aa2b2993170ea1e6c88e7d15362954c7`. Before this acceptance edit,
  status contained 173 rename-aware entries at the 1% threshold and 174 entries
  at Git's default threshold. The old plan path is absent; staged history is one
  in-place rename to this file.
- Current Package F identities are:
  `Terminal.cpp` =
  `468c9765ce2221d78c9b3a771d119a79383719efa7d0f53ce5dfbc42b20ebb3d`;
  `Terminal.h` =
  `5a83a5eae5dfb192fb8104a541f0d2cf7c10c7837b169aa7ae80fba990525dd6`;
  `TerminalKittyImagePipeline.cpp` =
  `119151e31137dc3e64e993bd1e46aa76ae7e60bc1647da8613ae462a09118538`;
  `TerminalKittyImagePipeline.h` =
  `8e331d01b3afd7fb343fb5b65e63da5f841d8ba059c0cf46bc9c157fb5162645`;
  Terminal source contracts =
  `7d471c5fbeb804d15fa233a1e3a6ea65c10cd6eb96ce0b9350b711a62e0bebd5`;
  Terminal spec =
  `94e53fe548a11c89f5f0172a98d5d353a24b75d8fd96ab37c405f74c9fbb13f7`;
  performance spec =
  `285880adefd4309d06ad1883dd4c2e60273647612f00571d22204109aca2d106`.
- Preserve the intentionally dirty worktree, including
  `.agents/skills/improve/SKILL.md` and the untracked
  `RedSalamander/CommandRuntimeState.h`,
  `Specs/Build/ToolchainIdentity.schema.json`, and
  `Specs/Testing/ValidationTerminalResult.schema.json`. Do not reset, clean,
  stash, bulk-commit, or disturb unrelated package work.
- `git diff --check` passes; its output contains line-ending conversion notices
  only. No new handoff file or plan directory was created.

#### Saved continuation state -- Package G has not started production edits (2026-08-17)

Resume from this subsection, in this same renamed-in-place plan. Package F is
closed. Package G is in the audit-before-RED phase; no Package G production file
has been edited yet and no Package G result may be called GREEN. The old
`CodeReview_LastFourDays_Remediation_2026-08-10.md` filesystem path is absent.

Verified starting state:

- `RedSalamander/FileActionLauncher.h` and `.cpp`,
  `RedSalamanderSearchService/Main.cpp`,
  `Tools/Tests/DeleteOnCloseTemporaryFileSourceContracts.Tests.ps1`,
  `Common/ProcessCommandLine.h`, `Common/HandleIo.h`, and
  `Common/DeleteOnCloseTemporaryFile.h` are clean relative to `HEAD` at this
  checkpoint. Do not attribute the existing edits in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `Specs/Core/Core_SharedHelpers.md`, or `Specs/UI/UI_CommandMenuKeyboard.md` to
  Package G; preserve and edit around them carefully.
- Starting SHA-256 identities are: `FileActionLauncher.h` =
  `bc32019a5ffbe8a080e04e25774b236153cb809a1467a86a7ba09643b057781e`;
  `FileActionLauncher.cpp` =
  `e1d2db2e0e4c2a2489b94de616b78348bd950ba844b68cd0a630f17511b0b3ff`;
  Commands settings selftest =
  `257d9464d01903a36fce836962e088658404cfd7023c3f1e428ef4dd4535aebd`;
  SearchService `Main.cpp` =
  `82c569a971320696e43467f37c9ec5a3f8cba76949c3f78c91fbe203f8f61cca`;
  selected-path source contracts =
  `91cb211adb2b87688cc6b1e2c00d2b684da1a1060e53cf20d64546a12f0046e0`;
  `ProcessCommandLine.h` =
  `2ff18b16b4a9ad262b8276dc2991c75c9dcc17be6236880a516645610522b271`;
  `HandleIo.h` =
  `6da8c015f7a21b72607317617ba627a8ed3dda8736eacfa8f889d6af00c97b5d`;
  and `DeleteOnCloseTemporaryFile.h` =
  `e43d31927da1cbe06c85f026fc609fbf2479460758432c9778c97afe44b8c7bd`.
- The current unsafe implementation still reserves a system-temp `rsa????.tmp`
  name with `GetTempFileNameW`, reopens it with `CREATE_ALWAYS`, copies the
  selection vector, builds a second full aggregate UTF-16 string, and writes
  that aggregate after the BOM. Cleanup uses path-based
  `std::filesystem::remove`; scavenging restarts at the beginning and calls
  `DeleteFileW`; the lease stores only a path. These are preserved RED targets,
  not accepted behavior.
- Command-line macro expansion still uses local
  `AppendWindowsQuotedArgument*` helpers even though the canonical ordinary
  argument helper is `Common::Process::QuoteWindowsCommandLineArgument`.
  Template-wrapped quote content needs separately named/documented semantics;
  it must not silently reuse an ordinary full-argument quoting contract.
- The current SearchService lifecycle child checks the manifest only with
  `GetFileAttributesW`; it does not open and parse exact UTF-16LE/BOM/CRLF
  records. The current source-contract test still recognizes the unsafe
  `rsa????.tmp` convention. Both tests must be converted to executable boundary
  proof rather than retained as implementation assertions.

Exact resume order:

1. Re-read L4-FILEACTION-03/04/05/07 and inspect the complete current
   `FileActionLauncher` implementation, existing app-data/no-follow/file-identity
   helpers, command selftest registration, and SearchService test-child parser
   before choosing seams. Reuse shared helpers where semantics match.
2. Add only the minimum test seams and deterministic tests needed to capture RED
   for private app-owned-root `CREATE_NEW` collision handling, all-record
   prevalidation, exact UTF-16LE BOM plus one CRLF-terminated record per nonempty
   path, 10,000-path bounded streaming/accounting, canonical ordinary quoting,
   and a real consumer parse. Build and record the exact failing results before
   changing production behavior.
3. Implement Package G without taking Package H's identity/recovery shortcut:
   use a validated private root, exclusive bounded random-name creation, validate
   every record before file creation, stream through `Common::HandleIo::WriteAll`
   without a copied selection or second full payload, and route ordinary quoting
   through the shared helper. Preserve selection order and do not normalize or
   deduplicate paths.
4. Close Package G only after focused Debug GREEN, five independent test-enabled
   x64 Release runs for the 10,000-path scenario, content-free performance
   evidence under the existing `Specs/TestRuns/<machine>/FileActions/<run>/`
   convention, authoritative UI/core/performance spec reconciliation, archive
   validation, tool/spec inventories, source contracts, and `git diff --check`.
   Then begin Package H at L4-FILEACTION-01/02/06/08 for recorded file identity,
   same-handle disposition, and bounded crash recovery.

#### Package G deterministic RED checkpoint -- preserved before production repair (2026-08-17)

The minimum test-only root/collision observation seam and four isolated Commands
cases were added without changing Release behavior. The reviewed selected-path
implementation remains deliberately unfixed at this checkpoint.

- `Tools/Tests/DeleteOnCloseTemporaryFileSourceContracts.Tests.ps1` is exactly
  `passed=3, failed=3`. The three failures require private app-data
  `CREATE_NEW`, all-record validation plus streamed writes, and canonical ordinary
  argument quoting. The pre-existing delete-on-close/provider/lifecycle checks
  remain green.
- The test-enabled x64 Debug `RedSalamander` build completed with zero warnings
  and zero errors. Receipt:
  `735d31aaa6380ab5df380bbb4a244cee7de081f3eb1899d10182bb26aae49e47`;
  log `.build/logs/msbuild-20260817_113055_324-pid65448-1b56db57.log`.
- The exact four-case Commands run is `passed=0, failed=4`, exit code 1, under
  `.build/codex-runs/fileaction-package-g-red-20260817_113500/last_run/`:
  - creation returned success after one attempt, reused the injected collision
    pathname, and replaced the sentinel (`samePath=true`,
    `collisionSurvived=false`);
  - an embedded-NUL record returned success and created a file after one attempt;
  - the 10,000-record path reported 4,400,000 path bytes, two writes, 4,458,198
    peak additional bytes, 8,800,000 selection-copy bytes, and
    `aggregate_payload=true`;
  - the real SearchService child returned success but wrote only `present`
    instead of parsing and reporting the two exact records.
- The Debug RED 10,000-record build took 27,207 us on machine
  `4cb089111a23`. Its diagnostic performance JSONL is 172,102 bytes with SHA-256
  `471abc799086b293affcf993a97e28859930fb5fb266d0324abba268dbdf584c`.
  This `.build` run is preserved diagnostic baseline evidence, not an accepted
  `Specs/TestRuns/` archive and not a Release comparison.

Production repair may now begin. Do not weaken or delete these four runtime
boundaries or the three source-contract groups to obtain GREEN.

#### Saved continuation state -- Package G compiled; expanded GREEN is next (2026-08-17)

This is the authoritative stop point requested for later continuation. It
replaces the stale "partial and unbuilt" checkpoint while retaining the earlier
starting-state and deterministic-RED sections as history. No new folder,
duplicate plan, commit, stash, cleanup operation, or compatibility copy was
created. `HEAD` remains detached at
`e9c76e59aa2b2993170ea1e6c88e7d15362954c7`; the old
`Specs/Plans/WIP/CodeReview_LastFourDays_Remediation_2026-08-10.md` filesystem
path remains absent, and its staged delete plus this file's staged add remain the
intentional in-place rename history.

Current Package G implementation saved in the worktree:

- `Common/PathUtils.h` provides a five-argument
  `CreateUniqueFileInDirectory(...)` overload with a deterministic candidate
  factory and bounded exclusive `CREATE_NEW` attempts; the four-argument public
  wrapper continues to generate GUID tokens.
- `RedSalamander/FileActionLauncher.h` exposes selected paths as
  `std::span<const std::filesystem::path>` and retains only test-enabled creation
  and streaming observations/options. All synchronous callers and Commands tests
  now keep their owning arrays/vectors alive across plan construction.
- `RedSalamander/FileActionLauncher.cpp` uses
  `Common::Process::QuoteWindowsCommandLineArgument(...)` for ordinary
  arguments and keeps the distinctly named, documented local content-only
  quoting helper for already-open template quotes. It prevalidates every record
  before directory/file creation, rejects NUL/CR/LF with
  `HRESULT_FROM_WIN32(ERROR_INVALID_DATA)`, preserves original order without
  normalization or deduplication, uses the private local-app-data
  `RedSalamander/SelectedPaths` root and bounded exclusive creation, and streams
  the UTF-16LE BOM plus every nonempty path and CRLF via
  `Common::HandleIo::WriteAll(...)`. Metrics cover validate/create/write/build
  time, path bytes, write calls, peak additional bytes, selection-copy bytes,
  and aggregate-payload truth.
- `RedSalamanderSearchService/Main.cpp` now opens the representative manifest
  with ordinary read/write/delete sharing, bounds it to 64 MiB, requires an even
  byte count and `FF FE` BOM, enforces nonempty exact CRLF-terminated UTF-16
  records with no NUL/bare CR/bare LF and a final CRLF, and writes
  `parsed:<count>` or `invalid`.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` retains the
  four RED-to-GREEN cases, adds empty/single/multiple/Unicode/duplicate and
  greater-than-600-character exact-record proof, upgrades the lifecycle and real
  User Menu route to the exact SearchService parser, and registers
  `file_action_provider_record_rejected_in_command_route` to prove a provider
  CRLF record is rejected by the public command route with a localized Warning
  containing exact HRESULT `0x8007000D` and no child launch.
- `Tools/Tests/DeleteOnCloseTemporaryFileSourceContracts.Tests.ps1` has six
  current Package G/source-contract groups. The obsolete system-temp prefix and
  minimum-age scavenger assertions were removed instead of preserving a dead
  constant; Package H owns the replacement recovery contract.

Verified evidence before this stop:

- Preserved RED remains immutable: source contracts `passed=3, failed=3`; four
  runtime cases `passed=0, failed=4` under
  `.build/codex-runs/fileaction-package-g-red-20260817_113500/last_run/`.
- After the core repair and exact child parser, the focused Debug run under
  `.build/codex-runs/fileaction-package-g-green-debug-20260817_115100/last_run/`
  passed `4/0`: creation, record grammar, 10,000-record streaming, and child
  parse. Its diagnostic artifact reports 10,000 records, 4,560,000 path bytes,
  20,001 writes, four peak additional bytes, zero selection-copy bytes, no
  aggregate payload, and success. This is Debug diagnostic evidence, not the
  required accepted Release archive.
- Adjacent launch-plan and selected-path lifecycle cases passed `2/0` under
  `.build/codex-runs/fileaction-package-g-regression-debug-20260817_115500/last_run/`.
- The current Pester 4 invocation of
  `Tools/Tests/DeleteOnCloseTemporaryFileSourceContracts.Tests.ps1` passed
  `6/0`. Use `Invoke-Pester -Script <path> -PassThru`; installed Pester 4 does
  not accept the Pester 5 `-Output` or `-Show` syntax used in two discarded
  invocations.
- The latest test-enabled x64 Debug build includes the subsequently added long,
  single-record, real User Menu parse, and provider-rejection proofs and is
  GREEN: exit 0, zero warnings/errors, receipt
  `54a5a33305229549b771647029c5eaf60f6ec20c5d25d740f6e17fceb6af7501`,
  source snapshot
  `f06429654e1dbd5c79a8e80f84f49cb11089d3ad92636989ddff7c976f294621`,
  log `.build/logs/msbuild-20260817_115445_652-pid16032-431eb9ab.log`, and
  receipt `.build/x64/Debug/build-receipt.json`. The process session expired
  after completion, but the receipt records `exit_code=0`, start
  `2026-08-17T09:54:45.4493011Z`, and end
  `2026-08-17T09:57:18.9151139Z`.

The current source identities are:

- `Common/PathUtils.h`:
  `219667b027dc7a27446437c7422fce0711d4e9b010d1443a9042095c7038861a`;
- `RedSalamander/FileActionLauncher.h`:
  `12728c9dbc9589bf6acc39c5e6b9715e53b3a99189a28ff794b4ae70db2ac168`;
- `RedSalamander/FileActionLauncher.cpp`:
  `04f7043413ea531cf82344e2fb63e534d92358c2a3c06ab6fbf9081bc2b622b1`;
- `RedSalamander/FolderWindow.Viewers.cpp`:
  `6767b1cddbda03c5f1582b504da337c8b1401f0fea0b470cb07653f2249718a9`;
- Commands settings selftest:
  `440976e8c0c90cb2f06c18a6ba439c3adcd159309187a430a985a072f8806c82`;
- SearchService `Main.cpp`:
  `4ff49efed97eee4f6c6893b6555f77f2f872216e11da3fd346d11b688504d777`;
- selected-path source contracts:
  `ca2f8035aa1e184f42c083e214eda2eb9493a2957f02dbd1dd26a5508daa6420`.

Exact next actions:

1. Do not rebuild first. Run the just-built Debug executable with one fresh
   `REDSALAMANDER_SELFTEST_ROOT` and this exact eight-case filter:
   `file_action_selected_paths_creation_contract,file_action_selected_paths_record_contract,file_action_selected_paths_streaming_perf,file_action_selected_paths_child_parse,file_action_provider_record_rejected_in_command_route,user_menu_populates_and_dispatches_configured_actions,file_action_external_launch_plan_macros,file_action_selected_paths_file_lifecycle`.
   Use `--commands-selftest --selftest-case=<filter>`, wait for process exit, and
   inspect `last_run/results.json`. The first four and last two have prior GREEN;
   the expanded record assertions, real User Menu parser, and provider-rejection
   route compile but have not yet run. Fix any failure without weakening exact
   byte, HRESULT, no-process, or parser boundaries.
2. Re-run the Pester source-contract file and record `6/0`. If source changes
   follow, rebuild and rerun the eight cases; never treat the prior receipt as
   attesting later source.
3. Reconcile Package G's durable contract in
   `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/Core/Core_SharedHelpers.md`,
   `Specs/Testing/Testing_PerformanceValidation.md`, and
   `Specs/Testing/Testing_SelfTests.md` (plus TestCoverage only if its ownership
   requires it). Preserve unrelated edits already present in those files. State
   the exact UTF-16LE BOM/CRLF grammar, pre-create rejection, order/no-dedup/no-
   normalization rule, private exclusive-create root, span/focused fallback,
   streamed writes, quoting distinction, exact child parser, real-route test,
   metrics, and scenario ownership.
4. After source and specs are stable, build test-enabled x64 Release and run
   `file_action_selected_paths_streaming_perf` in five independent processes
   with separate selftest roots. Archive only content-free evidence under the
   existing `Specs/TestRuns/4cb089111a23/FileActions/<run>/` convention, validate
   each archive with `Tools/Test-TestRunArchive.ps1`, and run the relevant tool
   inventory, spec inventory, source-contract, and `git diff --check` gates.
   Describe the 10,000-record scenario as the synchronous UI-thread build/write
   path unless a full 10,000-selection command-route measurement is added; do
   not overclaim end-to-end UI command time. Path bytes vary with sandbox-root
   length, so compare write/copy/peak shape and timings rather than expecting the
   RED and GREEN byte totals to match.
5. Close Package G only when the expanded Debug selection, current source
   contracts, authoritative specs, five Release runs, content-free archives,
   archive validation, inventories, and diff check are all accepted. Then begin
   Package H at L4-FILEACTION-01/02/06/08. The current path-based
   `CleanupSelectedPathsFile`, test-only legacy scavenger seam, and incomplete
   root-identity/recovery protection are not Package G claims: Package H must
   add recorded live identity, same-handle disposition, replacement-resistant
   cleanup, and bounded off-hot-path recovery.
6. Continue Packages I, J, and K, then run the required final Fresh Full,
   reconcile every authoritative spec, move this one plan to `Specs/Plans/Done/`,
   and update `Specs/Plans/WIP/README.md`. Do not mark the overall plan complete
   before those gates.

#### Earlier saved continuation state -- Package G complete; Package H was next (2026-08-17)

This checkpoint is retained as Package G history and is superseded by the
Package H checkpoint below. It originally superseded the earlier
“Package G compiled; expanded GREEN is next” action list. It is stored in this
same renamed-in-place plan. No new plan folder, duplicate plan, compatibility
copy, commit, stash, reset, or cleanup operation was created. `HEAD` remains
detached at `e9c76e59aa2b2993170ea1e6c88e7d15362954c7`. Preserve all existing
tracked and untracked worktree changes, including `.agents/skills/improve/SKILL.md`,
`RedSalamander/CommandRuntimeState.h`,
`Specs/Build/ToolchainIdentity.schema.json`, and
`Specs/Testing/ValidationTerminalResult.schema.json`.

Package G work now present in the worktree:

- The implementation and tests described by the preceding Package G checkpoint
  remain, but `RedSalamander/FileActionLauncher.cpp` now uses one bounded
  64-KiB buffer. It validates before creation, writes the BOM/record/CRLF stream
  without copying the selection or aggregating the full payload, flushes bounded
  chunks through `Common::HandleIo::WriteAll(...)`, and reports the exact write
  count as `ceil(totalBytes / 64 KiB)`.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` now includes a
  safe test-only Release reference that reproduces the reviewed algorithm's
  redundant selection copy, aggregate UTF-16 payload, and two writes. It does
  not reproduce the old unsafe shared-temp pathname race. Candidate execution
  precedes reference execution, and the artifact states that comparison order.
- `Tools/Tests/DeleteOnCloseTemporaryFileSourceContracts.Tests.ps1` contains six
  current source-contract groups, including the 64-KiB buffer assertions.
- Durable Package G requirements have been reconciled in
  `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/Core/Core_SharedHelpers.md`,
  `Specs/Testing/Testing_PerformanceValidation.md`,
  `Specs/Testing/Testing_SelfTests.md`, and
  `Specs/Testing/Testing_TestCoverage.md`. They cover the exact UTF-16LE
  BOM/nonempty-record/CRLF grammar, pre-create rejection, original order with no
  normalization or deduplication, private exclusive-create root, span/focused
  fallback, bounded streaming, quoting distinction, exact child parser, real
  route tests, metrics, and scenario ownership.
- The test-only large Fairstream phase dispatcher has a narrowly scoped C4883
  suppression in
  `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp`;
  the first Release build exposed that warning, and all later accepted Release
  builds were zero-warning/zero-error.

Evidence already completed:

- The expanded eight-case Debug selection passed `8/0` at
  `.build/codex-runs/fileaction-package-g-expanded-debug-20260817_120100/last_run/`.
  This predates the later bounded-buffer production edit and is therefore useful
  diagnostic evidence, not the final current-source Debug acceptance gate.
- The current source-contract invocation passed `6/0` after the bounded-buffer
  edit using Pester 4's `Invoke-Pester -Script <path> -PassThru` form.
- The current test-enabled x64 Release build completed with zero warnings and
  zero errors. Receipt ID
  `e5b8e0400b84927399c04638569da9761c88abda2bbabb22402fdf5a9b3d09dd`,
  source snapshot
  `bb24dc3764a4cdab21982b445354899e4926951f2a08d8bd805b1033886bdb85`,
  log `.build/logs/msbuild-20260817_122218_290-pid74180-454dde5f.log`, and
  receipt `.build/x64/Release/build-receipt.json`. This checkpoint-only plan edit
  occurred afterward; do not represent the receipt as attesting later production
  or spec edits if any are made.
- A single bounded proof passed `1/0` at
  `.build/codex-runs/fileaction-package-g-buffered-proof-20260817_122622/last_run/`.
- Five independent test-enabled Release processes passed `1/0` each under
  `.build/codex-runs/fileaction-package-g-buffered-release-20260817_122637/`
  (`r01` through `r05`). Candidate metrics were: validation 503-679 us
  (544-us median), create 125-172 us (155-us median), write 4,019-8,209 us
  (4,084-us median), build 4,809-9,216 us (5,102-us median), and scenario
  4,855-9,268 us (5,149-us median). Every candidate processed 10,000 records,
  used 73 writes and a 65,536-byte peak buffer, copied zero selection bytes,
  and created no aggregate payload. Every safe reference used two writes,
  7,165,218 peak additional bytes, 9,480,000 selection-copy bytes, and a
  4,780,000-byte aggregate payload; reference build time was 4,716-12,500 us
  (5,157-us median). The content-free claim rows and summaries are now curated
  under
  `Specs/TestRuns/4cb089111a23/FileActions/20260817_102637_selected_paths_streaming_release/`.
  Explicit pre-stage validation passed for all four files, whole-inventory
  validation passed for 541 files, spec inventory reported zero blocking
  findings, tool inventory classified 172/172 files with zero blocking
  findings, current source contracts passed `6/0`, and `git diff --check`
  reported no whitespace errors.
- The earlier unbuffered paired diagnostic at
  `.build/codex-runs/fileaction-package-g-paired-proof-20260817_122026/`
  exposed 20,001 candidate writes and motivated the bounded-buffer repair. Do
  not archive or present that superseded candidate as the accepted result.
- Final current-source acceptance is GREEN. The fresh test-enabled x64 Debug
  rebuild completed with zero warnings/errors, receipt
  `63985661e22e5a4246313003b3b2e4e6c09e4760e08e21dad0495c32e634c13b`,
  source snapshot
  `4ebe29d501e70973343e14f8feb09da176b93f4d0e2ecd1513d5269cb2011594`,
  and log `.build/logs/msbuild-20260817_134046_973-pid86340-a59fcf9e.log`.
  The exact eight-case process then passed `8/0`, with zero failures/skips and
  exit 0, under
  `.build/codex-runs/fileaction-package-g-final-debug-20260817_134407/last_run/`.
  Package G is complete and accepted as of 2026-08-17.

Current Package G source identities for change detection:

- `Common/PathUtils.h`:
  `219667b027dc7a27446437c7422fce0711d4e9b010d1443a9042095c7038861a`;
- `RedSalamander/FileActionLauncher.h`:
  `12728c9dbc9589bf6acc39c5e6b9715e53b3a99189a28ff794b4ae70db2ac168`;
- `RedSalamander/FileActionLauncher.cpp`:
  `ddfa60debc4ba64e1da9a5aa0acdd1d5861c72167e9bb6954d772410f9e7decf`;
- Commands settings selftest:
  `d7d5967d8445abd755cea9d9422e001d7221057fa5fc34de9342b1e89dd6dfe1`;
- Commands selftest dispatcher:
  `bb6834d47c8f63db7112c0c87b515fdba0cc13db13b94256a1f12add9c25807c`;
- SearchService `Main.cpp`:
  `4ff49efed97eee4f6c6893b6555f77f2f872216e11da3fd346d11b688504d777`;
- selected-path source contracts:
  `0e99e4b20ff757422206a621815d32456742e9fe9cbeaf8943b26018f9c768d7`.

Exact resume order:

1. Package G is complete. Do not regenerate its archive unless Package G
   behavior or claim-bearing evidence changes.
2. Begin Package H. Replace the remaining
   path-based `CleanupSelectedPathsFile`, old `rsa` filename ownership, and
   test-only legacy scavenger behavior with recorded live `FILE_ID_INFO`,
   creation-handle close before child launch, cleanup through one
   `DELETE | FILE_READ_ATTRIBUTES` `OPEN_EXISTING` handle opened with
   `FILE_FLAG_OPEN_REPARSE_POINT` and no delete sharing, identity comparison plus
   same-handle disposition, and bounded private-root recovery off the launch hot
   path. Prove exact cleanup, same-path replacement survival, shared-temp `rsa`
   survival, old private-root regular-file recovery, fresh/live/directory/reparse/
   out-of-root survival, bounded entries/deletes/time, launch-fault handle truth,
   and a real child read/replacement case.
3. Continue Packages I, J, and K. The overall plan remains active until the final
   Fresh Full passes, all authoritative specs are reconciled, this single plan is
   moved to `Specs/Plans/Done/`, and `Specs/Plans/WIP/README.md` is updated.

#### Latest saved continuation state -- Package H complete; Package I is next (2026-08-17)

This is the newest authoritative resume point. Package H closes
L4-FILEACTION-01/02/06/08 without creating a second plan or directory.

Accepted implementation and proof:

- `SelectedPathsFileLease` records the manifest and private-root `FILE_ID_INFO`,
  closes creation/root handles before child launch, and cleans only after
  reopening the same non-reparse root and same file identity. The candidate is
  opened once with `DELETE | FILE_READ_ATTRIBUTES`, `OPEN_EXISTING`,
  `FILE_FLAG_OPEN_REPARSE_POINT`, and read/write sharing without delete sharing;
  disposition occurs through that verified handle. Replacements, root
  replacement, directories, reparse points, missing/busy files, and uncertain
  objects survive.
- Shared `%TEMP%` `rsa????.tmp` ownership/scanning is gone. Recovery enumerates
  only `%LOCALAPPDATA%\RedSalamander\SelectedPaths`, is scheduled once per
  process on a threadpool callback after the first production-root resolution,
  applies a 24-hour age threshold, and uses 128-entry, 16-delete, 20-ms,
  32-pass bounds with a continuation cursor and content-free metrics.
- SearchService now opens and parses the exact manifest and can deliberately
  replace it. Commands coverage proves exact cleanup, regular-file replacement,
  root replacement, reparse preservation, shared-temp/out-of-root preservation,
  zero-time cutoff, missing-cursor reset, entry/delete bounds, and multi-pass
  convergence. The legacy lifecycle fixture now uses its governed per-case
  sandbox instead of constructing a 261-character artifact-tree marker path.
- Durable behavior is reconciled in `Specs/UI/UI_CommandMenuKeyboard.md`,
  `Specs/Core/Core_SharedHelpers.md`,
  `Specs/Testing/Testing_PerformanceValidation.md`,
  `Specs/Testing/Testing_SelfTests.md`, and
  `Specs/Testing/Testing_TestCoverage.md`.

Accepted evidence:

- Current source contracts pass `9/0`.
- The current test-enabled Debug targeted build completed with zero warnings and
  errors: receipt
  `43cd505f6c3a06637f349643b49d5a58d49c7017719a37e72e2bbb4ced47d0e6`,
  source snapshot
  `043fccd0dca918d0f6b20e993969475348513513b671b76c61e0f0b7177d7ba6`,
  log `.build/logs/msbuild-20260817_140742_591-pid85384-ec9967f3.log`.
  The exact ten-case Debug process passed `10/0`, exit 0, under
  `.build/codex-runs/runs/fileaction-package-h-debug-20260817-141100/artifacts/selftest/last_run/`.
- The test-enabled x64 Release solution build completed with zero errors:
  receipt
  `5d9b10ac85b6cc0b822fbd368399ebba4214578375ee41745aace77c9782d4302`,
  source snapshot
  `c895e744f9eaf8582fcdb82329c88281d26225e3400ea553302e1da80bbe152a`,
  log `.build/logs/msbuild-20260817_141227_655-pid66932-d850205a.log`.
  Its 11 warnings are pre-existing Release-only unused internal helpers in
  `Tests/ViewerPETests/ViewerPETests.cpp`, not Package H diagnostics.
- Five independent Release processes passed `1/0` each. Every process converged
  in eight passes, inspected 18 entries, deleted all six reclaimable manifests,
  preserved seven protected controls, reported zero errors, and observed seven
  bound hits. Elapsed recovery time was 37,493-42,067 us (38,367-us median).
  The removed shared-temp scan had no accepted instrumentation, so this is an
  explicit invariant/scaling baseline rather than an invented before comparison.
- Curated evidence is under
  `Specs/TestRuns/4cb089111a23/FileActions/20260817_141706_selected_paths_recovery_release/`.
  Explicit archive validation passed for four files, whole-inventory validation
  passed for 541 files, tool inventory classified 172/172 with zero findings,
  and `git diff --check` found no whitespace errors.

Exact resume order:

1. Package H is complete. Do not regenerate its archive unless Package H
   behavior or claim-bearing evidence changes.
2. Implement Package I at L4-MON-02/03: transactional named-exception
   thread-start helpers and canonical reader/export record idempotence, including
   real command-path failure seams and exact byte/record matrices.
3. Implement Package J at L4-MON-01 with immutable block snapshots and the
   required Release scaling/regression evidence.
4. Execute Package K: final Fresh Full, inventories and archive audit,
   authoritative-spec closeout, move this one plan to `Specs/Plans/Done/`, and
   update `Specs/Plans/WIP/README.md`.

#### Latest saved continuation state -- Packages I/J complete; Package K is next (2026-08-17)

This is the newest authoritative resume point. Packages I and J close
L4-MON-01/02/03 without creating a second plan or directory.

Accepted implementation and proof:

- `MonitorTextSnapshot` and its immutable `MonitorTextBlock` moved from the
  reader adapter into reader-independent `MonitorTextSnapshot.h`. `Document`
  owns those same immutable record allocations. Capture copies ownership
  descriptors only; later append, eviction, clear, and replacement cannot alter
  or invalidate an in-flight snapshot. The mutable active-tail bound and copied
  UTF-16 payload are both exactly zero.
- Reader EOF is delimiter-aware: LF publishes one record, EOF publishes only a
  remaining nonempty unterminated record, and empty/BOM-only input has zero
  records. Export remains BOM plus one final LF per record, so canonical bytes
  are idempotent without losing real consecutive/trailing blank records.
- Monitor-local `noexcept` open/save start helpers catch only
  `std::system_error`, map/log once, transactionally roll back active/shared
  state, and post no completion on failure. `std::bad_alloc` remains fatal. The
  production admission helpers have fail-next seams; the Chrome drill proves
  no active flag, shared state, joinable worker, completion post, or destination
  mutation and then proves both retries start.
- `MonitorTest --document-model-selftest` covers the exact canonical matrix,
  line-budget boundary/+1, tail mutation after capture, retention/clear/
  replacement lifetime, practical-large zero-copy capture, transactional
  preservation, and existing document/search/ETW shutdown behavior.
- Durable behavior is reconciled in
  `Specs/Core/Core_RedSalamanderMonitor.md`,
  `Specs/Testing/Testing_PerformanceValidation.md`,
  `Specs/Testing/Testing_SelfTests.md`, and
  `Specs/Testing/Testing_TestCoverage.md`. One localized launch-failure string
  is present in the base resource and all four satellites; localization
  contracts pass `6/0`.

Accepted evidence:

- Current test-enabled x64 Release `MonitorTest` built with zero warnings/errors
  and receipt
  `bd973fa04445f6080f3935516dca30c52c0ac7122db1b50d306ecf958d2a8d33`;
  log `.build/logs/msbuild-20260817_144034_631-pid81712-02ccc84f.log`.
  `MonitorTest.exe --document-model-selftest` passed with exit 0.
- Current test-enabled x64 Release `RedSalamanderMonitor` built with zero
  warnings/errors and receipt
  `5e8f6314a135163c68a3dcd256e4bdaa3e0813741b14032d692615f6823ea07c`;
  log `.build/logs/msbuild-20260817_144143_585-pid16300-d2f95d43.log`.
- The exact Chrome/ETW/scrollback command passed. Its scaling drill produced
  five repeats each at 512, 4,096, and 16,384 records. At 16,384 records /
  8,060,928 UTF-16 bytes, legacy deep-copy p95 was 3,395 us and candidate
  snapshot-lock p95 was 1,008 us. Every candidate reported zero copied text and
  active-tail bytes; peak additional snapshot storage was 262,144 bytes. ETW
  append-to-visible p95 was 18,303 us and batch-drain p95 was 2,405 us.
- Curated evidence is under
  `Specs/TestRuns/4cb089111a23/Monitor/2026-08-17_144304/` with 1,419 raw JSONL
  rows, results, trace, and summary README.

Exact resume order:

1. Packages I/J are complete. Do not change Monitor code or regenerate their
   archive unless validation exposes a real regression.
2. Execute Package K: run final test-enabled Debug Fresh Full, then tool/spec
   inventories, whole TestRuns archive validation, resource contracts, and
   `git diff --check`.
3. Reconcile any final receipt/result links, mark the checklist complete, move
   this one file to `Specs/Plans/Done/`, and remove its entry from
   `Specs/Plans/WIP/README.md` without recreating the old filename.

#### Final Package K closeout -- failures repaired; final Fresh rerun explicitly waived (2026-08-17)

This is the final closeout record and supersedes the Package I/J resume order
immediately above. No new plan or folder was created. The original plan was
renamed in place when its scope expanded and was then moved to
`Specs/Plans/Done/` by explicit user direction after the validation exception
below was recorded.

Fresh Full result and diagnosis:

- The completed Fresh run is checkpoint
  `.build/ValidationEvidence/runs/20260817T141650Z-89604-c46c4ab82ab6400db75e6d5a782585ac`
  with isolated test root
  `C:\RSPerf\runs\20260817T141650Z-89604-c46c4ab82ab6400db75e6d5a782585ac`.
  It finished in 56m57s with 1,901 total, 1,844 passed, five failed, and 52
  skipped. Every non-Commands suite passed, including both Pester lanes,
  Compare Directories, File Operations, DxUi, plugins/viewers, Monitor ETW, and
  PerformanceTests2.
- The five failures were
  `pane_view_options_toggle_preview_pane_tabs_and_selection`,
  `cmd_preferences_dialog_category_tree_keyboard_expand_collapse_and_child_entry`,
  `cmd_app_shortcuts_long_run_open_close_stays_stable`,
  `cmd_app_shortcuts_row_tooltip_tracks_hovered_cell`, and
  `cmd_pane_createDirectory_prompt_long_run_open_close_stays_stable`.
- `WindowHost::UpdateSupplementalTooltipTarget` treated an empty weak pointer as
  expired even when there was no supplemental target, so the empty
  supplemental pass cleared the tooltip just established by the interactive
  control. The lifetime check is now conditional on an existing supplemental
  target. A host-routed DxUi regression covers the empty-target pass and the
  existing Preview/Shortcuts tooltip scenarios cover both production consumers.
- The Shortcuts reopen scenario now waits for cleared search text instead of
  sampling a transient prior value. Create Directory now waits for pane
  enumeration through the canonical `WaitForPaneItems` helper. Preferences now
  sends real key lParams through the production debug route, does not equate a
  plugin detail page's independent scroll position with the parent page, waits
  for Plugins page initialization before establishing tree focus, and records
  focused diagnostic state on failure.

Accepted focused evidence on the final binary:

- Final x64 Debug `RedSalamander` build passed with zero warnings/errors,
  receipt `457a61997fc1112d9a005469e00935eaed75e5ef6d6b9c8b85a7c26f3ef4e8b1`,
  and log `.build/logs/msbuild-20260817_174513_905-pid37468-95821138.log`.
- Preferences passed 20/20 serial repetitions under run
  `focused-prefs-final-20260817-affd7a1aba764856a4fb18b389915f45`.
  On that same final binary, Preview, Shortcuts long-run, Shortcuts tooltip, and
  Create Directory each passed 5/5 serial repetitions under runs beginning
  `focused-final-binary-` in `C:\RSPerf\runs\`.
- The focused `DxUiTests.exe --suite=Tooltip` regression passed. The focused
  TestHarness source contracts passed 175/175, RunAllTestsPlan passed 53/53,
  the complete non-build tooling Pester cohort passed 634/634, and resource
  localization contracts passed 6/6.
- `Get-ToolInventory.ps1 -FailOnFindings` reports 172 discovered/classified and
  zero blockers. `Get-SpecInventory.ps1 -FailOnFindings` reports zero blockers.
  Whole-inventory TestRuns validation passed for 541 files, the Monitor archive
  passed explicit validation for all four files, and `git diff --check` passed.

Closeout decision:

1. Do not reopen the five repaired cases unless new focused evidence fails.
2. The user explicitly declined a second Full run after the focused repairs and
   accepted that the required final Fresh aggregate was therefore not rerun.
3. The user then explicitly directed this same plan to `Specs/Plans/Done/`, its
   WIP index entry was removed, and no compatibility copy or old filename was
   recreated.

After those trust and containment gates, continue the original independent
packages:

1. **Package A -- Curl capability/flag truth (complete 2026-08-16):**
   L4-CURL-01/08/09 RED proofs, truthful imports, correct flag
   validation/probing, focused runtime proof, and authoritative contracts.
2. **Package B -- Curl read/source integrity (complete 2026-08-16):**
   L4-CURL-03/06 known-size seams, direct-reader and native Copy/Move repairs,
   destructive-Move proof, authoritative specs, and five-run Release evidence.
3. **Package C -- Curl conditional publication decision (complete 2026-08-16):**
   L4-CURL-04/05 protocol matrix and ABI/product decision, fail-closed
   no-overwrite behavior, capability/host gating, preservation-only recovery,
   deterministic RED-to-green proof, three authoritative specs, and accepted
   five-run Release evidence.
4. **Package D -- Curl transaction/cleanup truth (complete 2026-08-16):** L4-CURL-02/07 structured
   publication result, all writer/Copy/Move call sites, cleanup-debt outcome, and
   focused performance/command evidence.
5. **Package E -- Terminal bounded cache/storage (complete 2026-08-17):**
   L4-TERM-01/02/03 keyed visible-delta cache, real host/runtime placement
   admission, high-cardinality paired comparison proof, current-source Debug
   gate, and validated five-process Release evidence.
6. **Package F -- Terminal recovery/resize/proof (complete 2026-08-17):** L4-TERM-04/05/06/07 upload
   recovery, active-close evidence, evidence-link correction, resize resource
   retention, and final spec reconciliation.
7. **Package G -- selected-path format/creation (complete 2026-08-17):**
   L4-FILEACTION-03/04/05/07
   private-root `CREATE_NEW`, strict records, bounded streamed writing, shared
   quoting, exact consumer bytes, and synchronous UI-path performance evidence.
8. **Package H -- selected-path identity/recovery (complete 2026-08-17):** L4-FILEACTION-01/02/06/08
   live file identity, same-handle disposition, private-root bounded-progress
   recovery, real child replacement/read tests, and spec/source-contract repair.
9. **Package I -- Monitor small correctness (complete 2026-08-17):** L4-MON-02/03 thread-start
   transactionality and canonical reader/export idempotence.
10. **Package J -- Monitor immutable snapshots (complete 2026-08-17):** L4-MON-01 design/test slice,
    implementation, full functional/performance closure, and spec change.
11. **Package K -- aggregate closeout:** complete full verification, reconcile
    overlaps, move this plan to `Done`, and update the WIP index.

Coordination rules:

- Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md
  is now a historical Done implementation record. This active remediation plan
  supersedes its closeout assertions and owns L4-BUILD-01..08 plus
  L4-STAR-01..16. Durable repairs still land in Build_Toolchain,
  Testing_ValidationEvidence, Testing_ToolingGovernance, and their machine
  contracts; do not rewrite the historical plan as if the defects were absent.
- Operation_Chordforge_TerminalCommandExperience_2026-08-11.md is already Done;
  these new Chordforge defects are owned here and their durable corrections go
  to the Terminal/UI/settings specs rather than reopening history in that plan.
- `FileOperations_RemainingStressValidation_2026-08-04.md` may consume the new
  deterministic Curl/host evidence but does not own these behavior changes.
- `Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md` owns the
  repository-wide source-shape migration. This plan may replace only assertions
  directly invalidated by its repairs and must prefer runtime boundaries.
- `Operation_PerfMeasurementContract_2026-07-06.md` owns reusable measurement
  infrastructure. This plan owns its Terminal/Monitor scenarios and archived
  before/after evidence and should reuse the established record format.

## Verification matrix

Run focused tests after each package and the authoritative aggregate before
closeout. Adjust an executable path only if the current build layout changed; do
not silently substitute a narrower suite for the full gate.

```powershell
.\build.ps1 -ProjectName FileSystemCurl -Configuration Debug
.\build.ps1 -ProjectName Terminal -Configuration Debug
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug

.\.build\x64\Debug\FileSystemCurlTests.exe
.\.build\x64\Debug\PluginContractTests.exe --terminal-selftests
.\.build\x64\Debug\MonitorTest.exe

.\Tools\Get-ToolInventory.ps1 -FailOnFindings
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\build.ps1 -Rebuild -Configuration Debug -Platform x64
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh

git diff --check
git status --short
```

For L4-BUILD/L4-STAR, add and run focused Pester/subprocess fixtures for receipt
projection and immutable-record tamper, stale/extra output, effective test
surface, MSIX restoration, selected Release/ARM64 writer profile, portable
producer mutation, compiler/SDK seam identity, included/excluded renames,
missing/malformed/partial result JSON, every documented exit category,
promotion/state cross-binding, crash-before-receipt suspension, immutable
decision epochs, result quota/content mutation, noncanonical/oversized records,
none/targeted/full/blocked build-action dispatch, executable-Spec impact
mutations, Fresh/Resume mode/coverage identity, per-component invalidation
reasons, derived import/consumer closure, and nested/junction prune/Resume/ZIP
targets. All junction tests use disposable
test roots and external sentinels; they never target user data.

For L4-CHORD, run PluginContractTests plus Commands runtime cases for output
canaries, callback pair/delivery, typed palette state, tab availability,
physical-key dispatch, scope-aware pass-through, Unicode deduplication, effective
shortcut aliases, message uniqueness, and the chosen atomic lockstep staging
invariant. Source-contract checks remain secondary and must not require an
unsafe vtable/output shape.

For L4-DXUI/L4-FILEOPS, run DxUiTests and Commands File Operations cases for
cross-path IUnknown equality, HWND reuse, live secondary-thread shutdown,
real passive-region WM_MOUSEMOVE tooltip delay/lifetime/click-through, retained-
tree rebuild during pending hover, oversized menu DPI relayout containment,
hosted/fallback graph and row-band pixels, saturating totals, independent
localized 1/2 strings, exact mixed delete outcomes, provider-sensitive case-only
names, user refocus, stale enumeration, and coalesced delete completions. The
180x16 hue-churn scenario must assert the linear work/geometry bound, not merely
render a screenshot.

Also run every new Curl reader/native Copy/Move/publication/rollback case, host
bridge capability/flag case, Commands selected-path case, Terminal
scroll/placement/upload/resize/active-close case, and Monitor
snapshot/thread-start/round-trip case directly through registered runtime test
names. Run the focused FileAction and Terminal source contracts only as secondary
structural guards. Record exact names/commands in the authoritative specs and
archived evidence. The full suite must finish; the incomplete 2026-08-10 baseline
is not closure evidence.

For performance-sensitive packages, build test-enabled Release using the current
`perf-validation` workflow, run each deterministic scenario five times before
and after on the same machine, and archive raw JSON plus the summary/readme under
`Specs/TestRuns/`. This is mandatory for Terminal placement/cache/resize,
selected-path 10,000-record/scavenger work, the File Operations 180x16
hue-churn graph, Monitor snapshot/ingestion, and any Curl change that alters
endpoint commands, buffering, or retry behavior.

## Final closeout checklist

- [x] Every production finding has a deterministic RED proof on the reviewed
      behavior; L4-TERM-05/06 and L4-FILEACTION-08 have direct proof/link/consumer
      credibility checks instead of invented runtime failures.
- [x] L4-CURL-01/08/09: import capability, contradictory flags, and effective
      atomic-probe flags are truthful before mutation.
- [x] L4-CURL-03/06: native transfers and direct readers enforce one known-size
      commitment; destructive Move never deletes an unproven source.
- [x] L4-CURL-04/05: no-overwrite is conditional or disabled, and rollback cannot
      delete concurrent file/tree ownership.
- [x] L4-CURL-02/07: every writer/recursive/single/batch Copy/Move call site
      separates committed success from observable, safe cleanup debt.
- [x] L4-TERM-01/02/03: geometry reveal refills a bounded keyed cache, all
      placements are admitted/accounted before filtering, and quadratic scans are
      removed without raising memory/pixel limits.
- [x] L4-TERM-04/05/07: bitmap failures recover without repaint loops, active
      conversion close is measured/bounded, and ordinary resize retains reusable
      GPU resources.
- [x] L4-TERM-06: all authoritative evidence links exist and one final Kitty
      narrative remains.
- [x] L4-FILEACTION-01/02/03: creation and cleanup are exclusive/identity-based,
      replacements/reparse targets survive, and shared `%TEMP%` is never scanned.
- [x] L4-FILEACTION-04/05/06/07: strict exact records stream with bounded UI
      memory, recovery makes bounded progress off the hot path, and canonical
      quoting is reused.
- [x] L4-FILEACTION-08: the real child reads/parses bytes and tests no longer
      enforce the defective prefix/path-only design.
- [x] L4-MON-02/03: forced worker-start failure is recoverable and canonical
      open/save is record-idempotent at empty/boundary/trailing-blank cases.
- [x] L4-MON-01 preserves exact-start export with active-tail-bounded UI copied
      bytes and no ingestion/search/render regression.
- [x] L4-BUILD-01/02/03: effective test surface is receipt identity; the pointer
      matches a verified immutable record; unique, governed outputs and exact
      runtime closure reject omitted, stale, or extra binaries.
- [x] L4-BUILD-04: MSIX never mutates tracked source; failure and success retain exact
      source bytes, and successful package output is included in the final receipt.
- [x] L4-BUILD-05: every Full writer mutates/re-attests only the selected
      platform/configuration.
- [x] L4-BUILD-06/07: portable payload is derived from verified receipt bytes
      plus reviewed auxiliaries, and one resolved compiler/linker/RC/SDK identity
      binds local and workflow evidence.
- [x] L4-BUILD-08 and L4-STAR-02: ZIP overwrite and recursive prune reject every
      ancestor/final reparse path; disposable external sentinels survive.
- [x] L4-STAR-01: included-side rename deletion/addition always changes source
      identity and widens impact, even when the other side is excluded.
- [x] L4-STAR-03/04/05: only normalized complete outcomes promote; wrong-suite
      JSON and zero-discovery Pester cannot pass; every promotion/evidence
      binding is verified by one loader; PASSED, NOT_EVALUATED, invalid,
      blocked, and corrupt exit distinctly.
- [x] L4-STAR-06/07: writer mutation suspends receipt/opens an epoch before byte
      deletion, blocks readers, and publishes immutable content-addressed epoch
      decisions with crash-recoverable state transitions.
- [x] L4-STAR-08/09/10/11: result bytes are atomic, hashed, and quota-bound;
      persisted JSON is exact canonical bounded input; Resume/prune/reporting use
      one direct-child non-reparse run resolver with leaf/state identity; the
      advertised build action is either executed or explicitly advisory.
- [x] L4-STAR-12: executable Spec data selects its dependent validation/build
      surface.
- [x] L4-STAR-13: Fresh/Resume persist truthful invocation modes over one explicit
      mode-independent coverage identity.
- [x] L4-STAR-14/15: component invalidation reasons are reachable and stable;
      module/schema/workflow dependency closure is derived and inventory/cache
      governance fails when a real consumer edge is missing.
- [x] L4-STAR-16: this active plan and authoritative specs supersede the
      historical Startrail Done assertions until every L4-BUILD/L4-STAR item is
      closed; indexes and links expose one unambiguous owner.
- [x] L4-CHORD-01/06: output canaries and all callback/cookie pair/delivery cases
      pass without overwrite or silently dropped exit events.
- [x] L4-CHORD-02/05/08: one truthful state/precedence resolver drives palette,
      shortcuts, host/plugin ownership, tab availability, enabled rendering, and
      effective shortcut chips with pre-dispatch generation revalidation.
- [x] L4-CHORD-03/04/07/09: physical keys dispatch in supported scopes,
      pass-through is Terminal-only, Unicode deduplication uses one relation, and
      all custom messages are centrally governed.
- [x] L4-CHORD-10: the lockstep same-IID exception and atomic host/Terminal
      staging invariant are explicit; no unsupported mixed-binary crash is
      claimed.
- [x] L4-DXUI-01/02/03/04: every accessibility root path returns one HWND
      identity, old identities become unavailable across reuse, each live
      WindowHost/TSF teardown executes on its attachment thread, passive
      tooltips obey hit/delay/lifetime/click-through contracts, and DPI-relayout
      menus remain work-area-contained with reachable content.
- [x] L4-FILEOPS-09/10/11/12/13: single-stream/default rendering matches
      fallback; concurrent row/band colors match through one renderer; 180x16
      churn stays within the linear geometry/time bound; aggregates saturate;
      singular/plural resources and dead formats are reconciled.
- [x] L4-FILEOPS-14/15: exact provider-sensitive per-source deletion outcomes
      choose the nearest actual survivor, while later user focus/navigation wins
      and stale/coalesced refreshes resolve every intent deterministically.
- [x] Focused tests, the rebuilt x64 Debug tree, and spec inventory are green;
      the completed Fresh aggregate found five Commands failures that are now
      focused-green. The user explicitly declined and waived a second Full run
      before directing closeout; this exception is recorded above.
- [x] Required same-machine Release evidence is archived for all named
      performance-sensitive packages.
- [x] Durable behavior is merged into every named authoritative spec; this plan
      contains no unique normative rule.
- [x] This plan is moved to `Specs/Plans/Done/` and the WIP index is reconciled at
      the closing commit.

## STOP conditions

Stop and request an explicit architecture/product decision rather than weakening
an invariant if any of the following occurs:

- Product requires Curl import/no-overwrite mutation to remain enabled but a
  protocol lacks atomic create-if-absent and immutable conditional rollback.
  Open the focused provider/ABI decision; do not pass overwrite flags, serialize
  only this process, or probe-then-rename/delete.
- Safe Curl recovery requires deleting a remote path without stable object
  identity. Preserve artifacts and request the identity/revision design; never
  trade partial clutter for possible concurrent data loss.
- The Kitty fix needs a higher 64-MiB/16-million-pixel limit, retains unbounded
  generations/placements, or the pinned engine cannot enforce placement
  admission without an unapproved runtime patch. Use the bounded delta design or
  make the runtime/ABI decision explicitly.
- Existing external-action compatibility requires accepting NUL/CR/LF records or
  treating a private-root filename as exact cross-restart identity. Decide a
  versioned manifest/recovery contract; do not silently escape records or delete
  by path.
- The Monitor design requires a permanent second full-text representation,
  changes exact-start export, or cannot bound UI copied bytes by the active block.
  Produce a measured storage-design decision before implementation continues.
- A build/portable design requires trusting files not named by a governed output
  or auxiliary manifest, permits a receipt pointer to self-assert identity, or
  infers test surface from configuration. Stop and define the manifest/property
  boundary; do not weaken the attestation claim.
- Evidence compatibility requires accepting missing/partial terminal results,
  mutable path-only artifacts, noncanonical records, or a schema-readable state
  label without full semantic provenance. Stop and version the evidence format;
  never map the legacy shape to PASSED.
- Product begins supporting independently staged/mixed Terminal host/plugin
  binaries. The current lockstep waiver then ends: freeze the old interface and
  introduce/query a new IID before shipping that boundary.
- A live secondary UI-thread WindowHost cannot be owner-thread drained before
  global teardown. Define an explicit marshal/join ownership protocol; do not
  call TSF/host teardown from the main thread as a fallback.
- The graph fix requires more than 16 active color slots, unbounded historical
  hue identity, or per-frame geometry construction proportional to item history.
  Produce a measured rendering/storage decision rather than raising the bound.
- Relevant production code/specs drift after the planned commit such that a RED
  proof no longer reproduces or another active owner has already changed the same
  invariant.
- Verification requires live destructive endpoints, credentials, publishing,
  pushing, or release actions not explicitly authorized by the user. Use bounded
  fake/loopback tests and stop before external side effects.
