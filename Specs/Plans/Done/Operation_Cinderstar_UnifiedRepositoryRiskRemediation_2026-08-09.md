# Operation Cinderstar — Unified Repository Risk Remediation

> **Executor instructions:** This is the sole live implementation ledger for the
> accepted remainder of the August 8 review, the July 21–August 4 residual review,
> Operation Emberglass, and Operation Floodgate. Execute one numbered package at a
> time, run its focused proof, update its checkbox/evidence here, and merge any lasting
> behavior into the named authoritative specification. Do not reopen the archived
> predecessor plans as parallel queues.
>
> **Drift check (run first):**
> `git diff --stat 68d2b3bc7b6f8f4e09163735124bd9d35c45890a..HEAD -- .github Common Plugins RedConfigure RedSalamander RedSalamanderMonitor Tests Tools Specs`
> If a cited symbol or failure interleaving changed, re-run the focused audit before
> implementation. STOP rather than applying a shape-only fix to drifted code.

## Status and ownership

- **State:** DONE — all owned packages, focused proof, archive/spec governance, and the uninterrupted Full closeout gate are green
- **Priority:** P0 master; contains P0 through P3 packages in exact rank order below
- **Effort / risk:** XL aggregate / HIGH; each package carries its own estimate and risk
- **Planned at:** `68d2b3bc7b6f8f4e09163735124bd9d35c45890a`, 2026-08-09
- **Review range:** `0f473e67bd547bc33a72cc9c3d7f39c99e8ca81a..68d2b3bc7b6f8f4e09163735124bd9d35c45890a`, plus revalidation of inherited rows at HEAD
- **Source plans:** archived under `../Done/` with their original evidence, refutations,
  checkpoints, and historical measurements intact
- **Size waiver:** the user requested one prioritized list. This file keeps the exact
  interleavings, dispositions, dependencies, proof gates, and STOP conditions needed to
  avoid duplicate remediation ownership; splitting it would recreate the four queues it
  replaces.
- **Implementation boundary:** packages 1–11 and 13–20 are owned here. Package 11 is
  closed by the maintainer-selected source-tree lockstep policy; package 12 is
  implementation-routed to the CI formatter plan. That routed row stays in this ranked
  view only so risk ordering remains complete.
- **Out of scope:** no production fix is implemented by this review; no historical perf
  archive is rewritten merely to improve a claim; no public ABI/product decision is inferred;
  no unrelated active WIP plan is absorbed.

## Audit verdict

The current code supports **39 accepted work items** represented by the 20 ranked packages
below. The plugin ABI-generation question is resolved as current source-tree lockstep; the
formatter publication risk remains visible here but retains its existing sole owner. The review also found two adjacent
cases not completely captured by the predecessor ledgers:

1. Curl's no-overwrite writer path has the same final-name publication race as overwrite:
   a concurrent creator can win after the probe and then be overwritten by the upload.
2. Terminal application-mouse state tracks only one pressed button, so a button chord can
   leave the TUI with a permanently pressed button.

All executable August 8 findings remain present at the planned commit. Earlier S3 proof
continues to establish source-revision capture and conditional reads/copy only; it does not
establish rollback isolation, versioned MOVE path hiding, or the real AWS request header.

**Concurrent-worktree note (2026-08-09):** after this audit started, uncommitted changes not
authored by this review appeared in `Tools/Build-GhosttyTerminalRuntime.ps1`,
`Tools/Tests/TerminalPluginSourceContracts.Tests.ps1`, and
`Specs/Terminal/Terminal_EmbeddedPlugin.md`. They attempt uucode source→canonical Zig-cache
materialization and LF-normalized overlay hashing. The targeted source-contract suite passed
20/20, but no clean runtime build, offline cache proof, locale proof, or import-closure proof has
closed package 8. Preserve that work and re-audit it; do not overwrite or count it as DONE.

## One prioritized list

Status values are `TODO`, `IN PROGRESS`, `DONE`, `BLOCKED`, and `REJECTED`. Priority describes
impact; rank also accounts for proof dependencies.

| Rank | Status | Package / source findings | Priority | Effort / risk | Concrete problem and strongest current evidence |
|---:|---|---|---:|---|---|
| 1 | DONE | **Immutable S3 MOVE identities** (`A8-S3-01`, `A8-S3-02`) | P0/P1 | L / HIGH | Discovery was valid. Final publication, backup, and delete-marker identities are journaled; rollback removes only operation-owned revisions and preserves concurrent replacements. Debug and test-enabled Release provider contracts, full Debug build/contract loading, focused FileOps MOVE, and bounded Release evidence are green. |
| 2 | DONE | **Transactional Curl writer publication** (`A8-CURL-01` plus adjacent no-overwrite race) | P1 | L / HIGH | Discovery was valid. Writer creation is local-only, and Commit uses a unique remote sibling, exact-size verification, overwrite backup/restore, and promotion. Portable no-overwrite and IMAP publication fail closed. Debug and test-enabled Release contracts, deterministic fault/cancellation/owner-lifetime coverage, focused source contracts, and Phase 11 FileOps are green. |
| 3 | DONE | **Host service ABI, marshaling, and modal ownership** (`A8-HOST-02`, `A8-HOST-01`, `A8-HOST-03`) | P1/P2 | M / MED | Discovery was valid. Result capacity is validated before writes, synchronous calls use a dedicated sender-owned opaque-token registry, and modal ownership resolves to a valid root with disable/restore coverage. Debug build, focused worker-thread ABI/UI coverage, and source contracts are green. The current full prefix remains explicit without preempting Atlas. |
| 4 | DONE | **Bounded canonical WSL discovery** (`R4-ARCH-01`) | P1 | M / MED | Discovery was valid. WSL key ownership/open and typed DWORD reads use `wil::reg`; `Common::Registry::ReadBoundedStringValue` is limited to the pre-allocation byte cap and strict query/read stability policy absent from WIL's allocating string getter. `Common::Wsl::EnumerateDistributions` owns stable name/GUID/base-path/version records, exclusions, and ordering for app and plugin consumers. Volatile-registry contracts, a zero-diagnostic full Debug rebuild, fresh plugin contracts, 155/155 source contracts, and an archived measured baseline are green. |
| 5 | DONE | **Visible-identity rectangular paste** (`R4-RC-01`) | P1 | M / MED | Discovery was valid. The UI now derives ordered `(owner, resource ID)` destinations from the filtered/sorted projection and culture IDs from pinned/reordered visible columns. The session resolves and validates the complete rectangle before its single Undo snapshot; missing/duplicate/out-of-bounds/hidden or placeholder-invalid targets fail unchanged. RedConfigure product/tests build with zero diagnostics, the full RedConfigure test executable is green, and source contracts pass 156/156. |
| 6 | DONE | **Monitor retained-state correctness and responsiveness** (`R4-MON-01` through `R4-MON-04`) | P1/P2 | L / HIGH | Discovery was valid. Open now strictly streams into one move-published decoded line snapshot under encoded/decoded-byte and line caps; capped search owns an eviction-stable scan frontier; Show IDs resets/rebuilds display-coordinate state; Save As captures one exact snapshot and converts/writes transactionally on an owned cancellable worker. Open/save close is bounded and payload-safe. Focused boundary/fault/cancel tests, 156/156 source contracts, zero-warning Release builds, and archived Debug/Release retained-state metrics are green. |
| 7 | DONE | **Terminal input and typed pane-location correctness** (`A8-TERM-01` through `A8-TERM-04`, new mouse-chord finding) | P1/P2 | M / MED | Discovery was valid. One-shot translated-character suppression preserves unselected Ctrl+C ETX; mouse chords retain/release exact buttons; canonical typed path/provider state eliminates display parsing and default-root fallback; real mounted `7z` control proves plugin-backed location. Debug rebuild, focused lifecycle, source contracts, localization, and archived metrics are green. |
| 8 | DONE | **Checkout-, locale-, and cache-invariant Terminal build** (`A8-BUILD-01` through `A8-BUILD-05`) | P1/P2/P3 | M / HIGH | Discovery was valid. Source/canonical uucode transport, archive, root, and license locks are distinct; all frozen text inputs are LF-pinned; tar preflight ignores localized dates while retaining safety bounds; actual imports are checked at reuse/build/final acceptance; duplicated helpers are gone. Full Terminal-engine Pester, clean Debug x64, Release x64/ARM64 Terminal, empty-online, and warm-offline proofs are green. |
| 9 | DONE | **Atomic release matrix and resumable Winget publication** (`A8-RELEASE-01`, `CR-RES-WINGET-ADOPTION` / `R4-CI-01`) | P1/P2 | M / MED | Discovery was valid. Explicit x64-only releases now skip Winget; release and normalized package/version publication concurrency prevent automation races. A manifest-digest state machine creates only from proven empty state, resumes one exact automation-owned PR, rejects all ambiguous/external/closed/drifted states, recovers post-create failures through bounded direct listing, and shares one PATCH-and-verify path. Stub/source tests pass 46/46 with no production token. |
| 10 | DONE | **Bounded asynchronous Kitty preparation** (`CR-RES-TERM-KITTY`) | P1 | L / HIGH | Discovery was valid. Ghostty bytes are copied into immutable generation work under ownership, then one latest-only `jthread` converts outside paint/`_terminalMutex`; stale results are rejected and UI upload is one bitmap/64 MiB per frame. Maximum raw/PNG, replacement, saturation, cancellation, device-loss retention, close/unload, metrics, zero-warning builds, 21/21 source contracts, 45/45 Release selftests, and archived 64 MiB/128 MiB evidence are green. |
| 11 | DONE | **Plugin/Terminal ABI generation and prefix decision** (`A8-ABI-01`, `CR-RES-TERM-ABI`) | P1 decision | S / MED | The maintainer selected current source-tree/release lockstep: the host and all shipped plugins change atomically, current methods require the full current record, and `sizeBytes`/retained IIDs make no pre-transition or mixed-release compatibility promise. Atlas records the decision; Plugin API, Viewer, Terminal, and public-header comments own the durable boundary. A future compatibility promise requires explicit pre-invocation generation negotiation/rejection, new IIDs/adapters, and frozen old/new x64/ARM64 fixtures. |
| 12 | ROUTED | **Formatter checkout/publication safety** (`A8-CI-01`) | P1 | M / MED | `.github/workflows/ci.yml:33-94` combines ambiguous PR checkout, persisted write credential, floating formatter acquisition, `git add -A`, and an unqualified push without stale-head/concurrency protection. **Sole implementation owner:** `CI_FormatAutocommitRace_2026-08-25.md`. |
| 13 | DONE | **Stable DxUi tab and text-unit semantics** (`A8-DXUI-01`, `A8-DXUI-02`) | P2 | M / MED | Discovery was valid. Generic `TabControl` reordering remains available, while the fixed Folder/Preview/Terminal strip explicitly disables it and preserves semantic indices through all drag permutations and hidden-tab keyboard navigation. Paragraph/Page normalize upward to Document through the shared expand/move/endpoint path. Focused Control, Accessibility, and Commands tests pass; the Debug x64 host build is warning/error free. |
| 14 | DONE | **Single escaping boundary for S3 CopySource** (`A8-S3-03`) | P2 | S–M / MED | Discovery was valid. The provider now passes logical UTF-8 bucket/key/VersionId text and the pinned AWS SDK owns the sole `URI::URLEncodePath` boundary. Both SDK serializers pass exact header checks for seven required input classes; reserved Unicode/versioned fixtures stay on direct single and multipart routes with zero relay. Release tests pass 157/157. Live endpoint timing is explicitly blocked by unavailable reviewed credentials; existing publication metrics remain the timing contract. |
| 15 | DONE | **Sensitive and delete-on-close temporary ownership** (`R4-FILEACTION-01`, `R4-ARCH-02`) | P2/P3 | M / MED | Discovery was valid. External actions transfer one move-only selected-path lease through launch/wait/deferred cleanup with a ten-minute hard fallback and exact aged-prefix scavenger. Curl/S3/Microsoft Drive now share one RAII reservation helper with provider policy inputs. Debug and test-enabled Release lifecycle/native contracts plus the focused source contract are green. |
| 16 | DONE | **Whole-inventory TestRun archive enforcement** (`R4-TEST-01`) | P2 | S / LOW | Discovery was valid for the July 20 archive and incomplete for pre-contract history. The 7,009,000-byte raw stream is replaced by a digest-bound ordered summary; explicit pre-stage, fast changed-set, and contract-era whole-inventory modes now enforce the 2 MiB/5 MiB limits without deleting grandfathered evidence. Focused Pester and all three live wrapper modes pass. |
| 17 | DONE | **Terminal release proof, localization, executable seams, then routing refactor** (`CR-RES-TERM-EVIDENCE`, `-LOCALIZATION`, `-EXECUTABLE-SEAMS`, `-WINDOWPROC`) | P2/P3 | L / MED | Terminal strings and UIA names are resource-owned; real plugin objects characterize scheduler, payload, renderer, accessibility, teardown, and unload behavior; the large WindowProc switch is routed through explicit handled/result seams. Release x64/ARM64 builds, runtime contracts, and bounded evidence are green. |
| 18 | DONE | **One truthful ADS removal command** (`R4-TOOL-01`) | P2 | S / LOW | `remove-ads.ps1` is the sole `SupportsShouldProcess` implementation with distinct matched/preview/removed/skipped/failed accounting; `remove-stream.ps1` forwards for compatibility. Command discovery and 8/8 live NTFS Pester cases pass. |
| 19 | DONE | **Bridge MOVE skipped-child behavioral proof** (`FG-P1-3`) | P2 tests | S / LOW | Real serial and configured-parallel dummy-to-local MOVE cases prove a skipped colliding child remains at source while the verified sibling source is removed. The production cleanup manifest now permits only this narrow file-conflict partial-success cleanup. |
| 20 | DONE | **Bounded transient re-verification for legacy bridge writers** (`FG-P0-2-RESTAT-FALSEPOS`, narrowed) | P3 | M / MED | Legacy publication verification is bounded to three cancellation-aware attempts for the narrow transient set. Pre-commit writer-position proof rejects false write counts while staging is still operation-owned; post-publication path-only deletion remains forbidden without immutable identity. Five legacy cases, the proof-capable control, and provider under-consumption coverage pass. |

### Exact source-ID crosswalk

This crosswalk is the completeness check for the transfer; ranges in the ranked prose do not
silently drop an inherited row.

| Package | Exact accepted/routed IDs |
|---:|---|
| 1 | `A8-S3-01`, `A8-S3-02` |
| 2 | `A8-CURL-01` (including the adjacent no-overwrite publication interleaving) |
| 3 | `A8-HOST-01`, `A8-HOST-02`, `A8-HOST-03` |
| 4 | `R4-ARCH-01` |
| 5 | `R4-RC-01` |
| 6 | `R4-MON-01`, `R4-MON-02`, `R4-MON-03`, `R4-MON-04` |
| 7 | `A8-TERM-01`, `A8-TERM-02`, `A8-TERM-03`, `A8-TERM-04`, `CS-NEW-TERM-MOUSE-01` |
| 8 | `A8-BUILD-01`, `A8-BUILD-02`, `A8-BUILD-03`, `A8-BUILD-04`, `A8-BUILD-05` |
| 9 | `A8-RELEASE-01`, `CR-RES-WINGET-ADOPTION` / `R4-CI-01` remainder |
| 10 | `CR-RES-TERM-KITTY` |
| 11 | `A8-ABI-01`, `CR-RES-TERM-ABI` — maintainer-resolved source-tree lockstep |
| 12 | `A8-CI-01` — implementation-routed |
| 13 | `A8-DXUI-01`, `A8-DXUI-02` |
| 14 | `A8-S3-03` |
| 15 | `R4-FILEACTION-01`, `R4-ARCH-02` |
| 16 | `R4-TEST-01` |
| 17 | `CR-RES-TERM-EVIDENCE`, `CR-RES-TERM-LOCALIZATION`, `CR-RES-TERM-EXECUTABLE-SEAMS`, `CR-RES-TERM-WINDOWPROC` |
| 18 | `R4-TOOL-01` |
| 19 | `FG-P1-3` |
| 20 | `FG-P0-2-RESTAT-FALSEPOS` (narrowed) |

## Package instructions and proof gates

### 1. Immutable S3 MOVE identities

- Make final publication return an immutable destination revision identity. Store that identity,
  the captured source identity, and any created delete-marker identity in the rollback journal.
- Restore a deleted source only from the exact revision published by this operation. Remove only
  the exact published revision/marker. If exact identity cannot be established, preserve all
  potentially authoritative objects and return partial failure.
- Ordinary versioned MOVE must hide the key, not merely delete captured `v2` and reveal `v1`.
  Use a key-level conditional removal/delete marker that is valid only while the captured revision
  is current; do not widen to unconditional delete.
- Add deterministic barriers for destination replacement after publication, rollback after
  replacement, old-version exposure, source replacement, cancel, and successful rollback. Assert
  exact version sets and visible keys, not only request counts.
- Update `Specs/FileSystem/FileSystem_S3.md`; preserve the older archived source-pinning metrics but
  narrow their claim. Because this is a hot remote path, define counters and capture test-enabled
  Release before/after evidence under `Specs/TestRuns/`.
- **Focused proof:** build `FileSystemS3` and `PluginContractTests`; run the complete S3 provider
  contract in Debug and test-enabled Release, then the focused FileOps MOVE family.
- **STOP:** the backend/API cannot condition key-level hiding on the captured current identity. Do
  not emulate safety with a pathname re-stat or unconditional delete; report the required API or
  product semantic decision.

Implementation evidence (2026-08-09):

- [x] Revalidated AWS SDK 1.11.840 support for conditional current-key delete, destination
  ETag/VersionId capture, and created delete-marker VersionId capture.
- [x] Replaced key-only rollback journal entries with exact publication, backup, source-marker,
  and original-revision identities; ambiguous successful responses preserve objects and return
  partial failure.
- [x] Added deterministic source-history hiding, source replacement, destination replacement,
  later-delete rollback, cancel, exact version-set, and operation-owned cleanup assertions.
- [x] Debug x64 `FileSystemS3` build: zero warnings/errors.
- [x] Focused Debug and test-enabled Release S3 plugin contract: 139 / 0 directory-transfer
  checks in each configuration; the subsequent full Debug `PluginContractTests` run also passed.
- [x] Focused host FileOps MOVE proof: `FileOpsFamily_ClearflowPhase01_DirectoryMerge`, 5 / 0,
  clean TestSandbox audit.
- [x] Test-enabled Release before/after evidence and bounded archive at
  `Specs/TestRuns/4cb089111a23/FileOps/20260809_105500_s3_immutable_move_identities_release/`.

### 2. Transactional Curl writer publication

- `CreateFileWriter` must not delete or mutate the final target. Stage locally or at a unique
  remote sibling and publish only from successful `Commit` after exact-size verification.
- Destruction/cancellation removes staging only. Overwrite must retain/recover the prior final on
  failed promotion; no-overwrite must make publication conditional so a creator after preflight wins.
- Keep protocol-specific promotion capabilities explicit. If FTP/SFTP/WebDAV/IMAP cannot provide
  the needed atomic condition, fail closed or document a narrower capability; never stream directly
  to the final name after an existence probe.
- Test abandoned writer, short/failed upload, promotion failure, overwrite sentinel, concurrent
  no-overwrite creator, cancellation, and unload. Reuse the existing bounded Curl fault transport.
- Update the Curl and File Operations specs. Capture request count, staged bytes, commit latency,
  cancellation, and cleanup evidence.
- **STOP:** a protocol has no recoverable/conditional publication mechanism and current ABI cannot
  express that limitation. Escalate the capability/semantic choice.

Implementation evidence (2026-08-09):

- [x] Reused the provider copy/move transaction: unique remote sibling, exact-size proof,
  overwrite backup/restore, promotion, and operation-owned cleanup.
- [x] `CreateFileWriter`/`Write` are local-only; abandonment contacts no endpoint. FTP/SFTP/SCP
  no-overwrite and all IMAP writer requests fail closed because their current rename/append
  surfaces cannot guarantee the required conditional publication.
- [x] Deterministic loopback FTP coverage includes commit, abandonment, short upload,
  promotion failure with sentinel restore, no-overwrite fail-closed, owner lifetime, and cleanup.
- [x] Debug x64 provider build and complete plugin contract: zero warnings/errors, Curl 130 / 0.
- [x] Test-enabled Release proof and bounded metrics archive at
  `Specs/TestRuns/4cb089111a23/FileOps/20260809_111300_curl_transactional_writer_release/`;
  source contracts passed 154 / 154 and isolated Phase 11 FileOps passed 9 / 0 with a clean
  TestSandbox disk audit after a full Debug rebuild.

### 3. Host service ABI, marshaling, and modal ownership

- Read and validate the caller's result capacity before every result write; accept only reviewed
  required prefixes plus oversized tails and leave rejected buffers byte-for-byte untouched.
- Separate synchronous `SendMessageW` ownership from `PostMessagePayload` tokens. Keep sender-owned
  storage alive across synchronous dispatch or add a dedicated synchronous helper; never release a
  raw pointer into `TakeMessagePayload`.
- Resolve a valid root/host HWND once for Connection Manager fallback and use it for disable/restore.
- Add worker-thread Prompt and Connection Manager tests, undersized/exact/oversized result fixtures,
  allocation/leak accounting, invalid/stale token controls, destroyed-window teardown, and owner
  modality checks. Update `Specs/Plugins/Plugins_PluginAPI.md` and Host ABI tests.
- **STOP:** required prefix sizes or the supported old/new binary boundary are unclear; coordinate
  package 11 before accepting more than the current safe prefix.

Implementation evidence (2026-08-09):

- [x] `ShowConnectionManager` rejects undersized results byte-for-byte unchanged, accepts exact
  and oversized records, and preserves unknown tail bytes while retaining the current full prefix.
- [x] Prompt and Connection Manager worker calls keep caller-owned storage live across a dedicated
  synchronous opaque token; stale, invalid, wrong-kind, and leaked registrations are covered.
- [x] Null, destroyed, invalid, and explicit owners resolve to a valid root; the unowned modal façade
  disables and restores that root on success and cancellation.
- [x] Debug x64 build passed with zero warnings/errors; focused Commands case
  `host_services_worker_marshaling_and_connection_manager_abi` passed 1 / 0 in 2.3 seconds with a
  clean disk audit; source contracts passed 154 / 154.
- [x] The first test attempt exposed a worker-to-DxUi test violation, preserved at
  `Specs/TestRuns/4cb089111a23/Continuation/20260809_114633_host_services_worker_modal_hang/`;
  automation now posts terminal window commands exclusively and the exact repro is green.

### 4. Bounded canonical WSL discovery

- First recheck `Specs/Core/Core_SharedHelpers.md`, `Common/`, and `Tests/TestSupport/`; no canonical
  registry-string/WSL catalog helper exists at the planned commit.
- Add the lowest-layer shared bounded reader using `RegGetValueW` or a bounded two-pass read. Use the
  returned byte count; reject wrong types/odd bytes; normalize terminators; cap allocation; retry a
  bounded number of times when the value changes.
- Return stable WSL records (name, GUID, base path, version) through one filter/sort policy and
  migrate all three consumers. Record the helper in `Core_SharedHelpers.md` and the owning WSL/VFS spec.
- Test terminated/nonterminated, long, empty, odd-byte, wrong-type, changing values, exclusions,
  case-insensitive sort, and WSL1/2 using a volatile per-user key.
- **STOP:** app/plugin layering cannot consume the proposed Common location; select and document the
  lowest existing shared library rather than create another consumer-local copy.

Implementation evidence (2026-08-09):

- [x] Rechecked `Core_SharedHelpers.md`, `Common/`, and `Tests/TestSupport/`; no matching shared reader
  or catalog existed, while app, standard filesystem plugin, and contract tests already shared `Common.dll`.
- [x] Added a bounded registry-string reader with a 64 KiB default / 1 MiB hard cap, three default /
  eight hard-maximum attempts, exact byte-count/type revalidation, safe nonterminated strings, trailing-
  terminator normalization, and rejection of odd-byte, wrong-type, embedded-null, empty-by-default,
  over-cap, or continually changing values.
- [x] Added the canonical 1,024-subkey-bounded WSL catalog and migrated `WSLDistro`, the standard
  filesystem navigation menu, and ShellNew/file-type registry strings to the shared layer. Stable records
  retain name, GUID, base path, and WSL1/2 state; utility exclusions and ordering are case-insensitive.
- [x] `PluginContractTests.exe --registry-wsl-catalog` passed 25 / 0 against a volatile per-user fixture;
  the fresh full plugin contract passed, source contracts passed 155 / 155, and full Debug rebuild
  `.build/logs/msbuild-20260809_120926_225.log` completed with zero warnings/errors.
- [x] Added `wsl.catalog.enumerate_us` and archived the Debug diagnostic baseline (136 us, five inspected
  subkeys, two retained records) at
  `Specs/TestRuns/4cb089111a23/Navigation/20260809_121500_wsl_catalog_debug/`; the analyzer resolves the
  metric and no percentile or Release-latency claim is made from the single sample.

### 5. Visible-identity rectangular paste

- At the UI boundary, build ordered destination row identities from the visible review projection
  and ordered culture IDs from the visible/pinned column order. Pass identities, not one raw start index.
- Resolve and validate the entire rectangle before mutation; use one undo unit and all-or-nothing
  placeholder validation. Missing/hidden identities fail without change.
- Test a 2x2 paste with stored cultures `fr,de,ja`, visible `ja,fr,de`, non-contiguous sorted/filtered
  rows, hidden culture, invalid placeholder, missing identity, and undo/redo.
- Update `Specs/Core/Core_RedConfigure.md` and `Specs/UI/UI_RedConfigure.md`; coordinate culture
  identity vocabulary with Rosetta Lantern without transferring correctness ownership.
- **STOP:** product rejects atomic invalid-placeholder behavior. Record the chosen partial behavior
  explicitly before implementation; do not leave the current accidental skip semantics.

Implementation evidence (2026-08-09):

- [x] Added stable row identity `(owner name, resource ID)` and visible paste-target projection; the UI
  uses the current filtered/sorted row sequence and `LanguageColumnModel` pinned/reordered culture order.
- [x] Session preflight resolves every row and culture uniquely, requires a rectangular in-bounds matrix,
  and validates every placeholder before `RecordUndoSnapshot`; rejected input changes neither data nor
  Undo/Redo history, while success is exactly one Undo unit.
- [x] Focused coverage models stored cultures `fr-FR,de-DE,ja-JP`, visible order
  `ja-JP,fr-FR,de-DE`, non-contiguous projected rows, a 2x2 paste, filtered neighbors, hidden culture,
  invalid placeholder, missing row identity, and complete Undo/Redo restoration.
- [x] `RedConfigureTests.exe` passed; RedConfigureTests build
  `.build/logs/msbuild-20260809_122143_511.log` and product build
  `.build/logs/msbuild-20260809_122322_662.log` completed with zero warnings/errors; source contracts
  passed 156 / 156.

### 6. Monitor retained-state correctness and responsiveness

Execute in this order: decoded snapshot, search frontier, Show IDs contract, asynchronous export.

- Replace whole-file decode with a strict streaming decoder that budgets retained UTF-16 bytes and
  lines, carries UTF-8/UTF-16 boundary state, polls cancellation during CPU work, and publishes one
  immutable worker-built line snapshot without reparsing it on the UI thread.
- Track an explicit search scan frontier through eviction. When slots reopen, scan the oldest
  retained unindexed range before the new tail; preserve sort/order and avoid full rebuild per append.
- Choose one Show IDs rule. Prefer reset/rebuild of search, match index, selection, and caret; if
  preservation is required, add one logical-to-display mapping and use it everywhere.
- Move Save As conversion/write to an owned cancellable worker with exact start-snapshot semantics,
  transactional publication, registered posted payloads, bounded close, and no document lock across I/O.
- Tests: decoded expansion at/over limit; split multibyte/surrogate boundaries; invalid/truncated input;
  mid-decode cancel/close; repeated match-cap eviction/refill; F3 ordering; Show IDs with metadata and
  ordinary lines; export during ingestion; cancel/fail keeps target unchanged; payload drain on close.
- Metrics: peak retained bytes, open/publish/cancel/close latency, append/search time at cap, save UI
  return time, ingestion queue/drop counters. Archive same-machine before/after evidence.
- Update `Specs/Core/Core_RedSalamanderMonitor.md` and performance/testing authority.
- **STOP:** the model cannot retain stable immutable line identity/snapshot without a second full
  document copy. Design the ownership representation before layering more offset arithmetic.

Implementation evidence (2026-08-09):

- [x] `MonitorFileReader` incrementally decodes 64 KiB chunks, carries split UTF-8 scalars and UTF-16
  bytes/surrogates, enforces encoded bytes, retained UTF-16 bytes, and lines before growth, and polls stop during
  CPU decode. Its worker-owned `MonitorTextSnapshot` moves line buffers directly into `Document`; no second whole
  encoded/decoded payload or UI reparse remains.
- [x] Search owns an explicit line/offset frontier. Oldest-first eviction shifts both matches and frontier, newly
  opened match capacity resumes at the oldest retained unscanned position, and matches stay ordered for F3/wrap.
  Show IDs uses the selected reset rule: selection/caret go to zero and matches/frontier rebuild for the new display
  offsets.
- [x] Save As captures one exact line snapshot under a shared lock, releases the lock before I/O, then performs strict
  streaming UTF-16-to-UTF-8 conversion through `LocalFileTransaction` on an owned worker. Later ingestion is excluded;
  cancellation/write/flush/encoding failures preserve the prior target and remove the temporary sibling.
- [x] File-open and Save As progress/completion use `PostMessagePayload`; `WM_NCDESTROY` drains the registry. Open
  and save cancellation use stop requests plus `CancelSynchronousIo`, wait at most two seconds on close, and detach
  only worker-owned state after that bound. Metrics cover open/read, peak retained bytes, publication, cancel/close,
  save snapshot/UI return/total/close, search update/rebuild, and ingestion queue/drop state.
- [x] `MonitorTest.exe --document-model-selftest` covers decoded expansion at/over limit, UTF-8 and UTF-16 chunk
  boundaries, truncated/invalid input, mid-decode cancellation, sparse >2 GiB rejection, move-publication, exact
  export snapshot during ingestion, cancellation/fault preservation, and temporary cleanup. Release build
  `.build/logs/msbuild-20260809_125011_568.log` completed with zero warnings/errors and the focused executable exited
  0.
- [x] Test-enabled Release Monitor build `.build/logs/msbuild-20260809_125226_672.log` completed with zero
  warnings/errors; `--chrome-selftest --wait-instance --perf` passed with correct `build: Release` labeling at
  `Specs/TestRuns/4cb089111a23/Monitor/2026-08-09_125240/`. The archive contains same-machine before/after caveats,
  all retained-state/frame metrics, deterministic fixture cleanup, and final Release observations. Source contracts
  passed 156 / 156 and `git diff --check` was clean.
- [x] Lasting behavior is merged into `Specs/Core/Core_RedSalamanderMonitor.md`,
  `Specs/Testing/Testing_PerformanceValidation.md`, and `Specs/Testing/Testing_TestCoverage.md`.

### 7. Terminal input and typed pane-location correctness

- Represent consumed keydown-to-character suppression explicitly and one-shot; later modifier state
  must not decide whether a queued translated character leaks. Preserve unselected Ctrl+C as ETX.
- Track a set of reported mouse buttons. Release the exact button on up and every still-pressed
  button on capture loss/teardown; do not clear unrelated state.
- Use `Common/PathUtils` canonical classes: drive and extended drive are `WindowsLocal`, UNC and
  extended UNC are `WindowsUnc`, device/unsupported paths fail closed.
- Never substitute the default Windows root for unsupported plugin namespaces. Publish typed
  `TerminalLocation` state derived from provider/`PaneState`, not display/history strings. When a
  namespace is unsupported, publish a valid detached/unsupported generation so old follow state dies.
- Add executable exact-byte tests for selected/unselected Ctrl+C and modifier release; mouse
  single/chord/mismatched-up/capture-loss sequences; extended drive/UNC/device cases; local→plugin
  generation advance; unsupported open; supported plugin backing-location control.
- Update `Specs/Terminal/Terminal_EmbeddedPlugin.md` and relevant UI/navigation spec. All user-facing
  rejection text must be localized.
- **STOP:** a plugin display string is the only available identity. Extend the typed pane state; do
  not parse a presentation string back into an ABI location.

- [x] Keydown-to-character suppression is explicit and one-shot; selected Ctrl+C cannot leak a
  translated byte, while unselected Ctrl+C emits exactly one ETX byte even if modifiers change before
  the queued character message.
- [x] Application-mouse reporting tracks left/middle/right buttons independently, releases the exact
  button on up, preserves unrelated chord state, and releases every retained button on capture loss
  and close. Exact Ghostty SGR byte assertions cover single, chord, mismatched-up, and capture-loss
  sequences.
- [x] Terminal locations use `Common::Paths` canonical path classes and typed pane/provider state.
  Unsupported providers publish a detached generation, retain no stale path/follow identity, show
  localized rejection, and never fall back to the default Windows root. A real mounted `7z` instance
  proves validated plugin-backed directory control without losing namespace identity.
- [x] The test-enabled Debug full rebuild `.build/logs/msbuild-20260809_132911_364.log` and follow-up
  host build `.build/logs/msbuild-20260809_133337_301.log` completed with zero warnings/errors. The
  focused lifecycle passed at `Specs/TestRuns/4cb089111a23/Commands/2026-08-09_133558/`, including
  archived open/reuse and contextual/full/current-directory insertion metrics. Final Release proof is
  intentionally owned by package 17.
- [x] Durable input, typed-location, navigation-publication, localization, and coverage rules are in
  `Specs/Terminal/Terminal_EmbeddedPlugin.md`, `Specs/UI/UI_NavigationView.md`, and
  `Specs/Testing/Testing_TestCoverage.md`.

### 8. Checkout-, locale-, and cache-invariant Terminal build

- Freeze uucode URI, cache filename, archive SHA-256, internal root, and license identity separately.
  Prove the selected artifact rather than changing only the current failing hash.
- Define one semantic identity for text fixtures: pin LF in `.gitattributes` for the exact inputs or
  hash a documented canonical Git-blob/LF representation. Generator and evaluator must agree.
- Replace localized verbose-tar parsing with a guaranteed invariant representation or native header
  parser while retaining traversal, duplicate, link, special-file, count, and expanded-size defenses.
- Re-derive the actual DLL import set at cache reuse and final acceptance; compare with policy and
  persisted identity. Keep exactly one definition of each duplicated helper.
- Tests: wrong hash/root/license, empty online cache, warm offline cache, CRLF/LF-equivalent checkout,
  localized tar fixtures, malicious archive corpus, tampered cached import, unexpected final import,
  exact-one helper definitions, x64/ARM64 PE identity.
- **Focused proof:** `Invoke-Pester Tools/Tests/TerminalEngine*.Tests.ps1`; clean Debug x64 and
  test-enabled Release Terminal builds; then ARM64. Expected: zero failures/warnings and identical
  identity from equivalent checkouts/locales.
- Update Terminal engine/build specs before collecting package 17 evidence.
- **STOP:** the selected archive cannot be independently reproduced or verified; do not bless a
  workspace cache as the source of truth.

- [x] uucode source transport, source cache leaf/SHA/root, canonical Zig cache
  leaf/SHA/root, and normalized license identity are independently named, validated,
  persisted in schema 5, and documented. A real empty-uucode-cache online bootstrap produced
  source SHA `7e76fc7f...94bc9e` and canonical SHA `3c571e2e...555ed0`; immediate warm-offline
  reuse returned the same AMD64 DLL SHA `7b5cdb2e...464592`.
- [x] The complete Terminal-engine frozen raw-byte surface is `eol=lf`; 88 existing
  `core.autocrlf`-converted files were mechanically renormalized. The aggregate
  `TerminalEngine*.Tests.ps1` run passed 212/217 with zero failures and five intentional skips,
  including transition, fixture, archive,
  CRLF/LF identity, malicious corpus, license substitution, and import-policy cases.
- [x] Tar preflight uses numeric owners and parses only invariant leading metadata fields.
  Actual PE import modules are re-derived at cache reuse, build acceptance, and final acceptance;
  the persisted list is an independent check. The two formerly duplicated builder helpers each
  have exactly one definition.
- [x] The clean full Debug x64 rebuild
  `.build/logs/msbuild-20260809_140204_735.log`, test-enabled Release x64 Terminal rebuild
  `.build/logs/msbuild-20260809_140522_729.log`, and warm test-enabled Release ARM64 Terminal
  build `.build/logs/msbuild-20260809_141856_870.log` each completed with zero warnings/errors.
  The ARM64 runtime is a real `ARM64` PE with raw SHA `ef671a5a...1c9643`; its initial cold vcpkg
  bootstrap also compiled successfully but retained eight host-installation `CS1668` warnings and
  is not the cited zero-diagnostic proof.
- [x] Durable checkout, dependency/archive/license, locale, import, and test-matrix rules are in
  `Specs/Build/Build_Toolchain.md`, `Specs/Terminal/Terminal_EmbeddedPlugin.md`,
  `Specs/Terminal/TerminalEngineDecision.md`, and `Specs/Testing/Testing_TestCoverage.md`.

### 9. Atomic release matrix and resumable Winget publication

- Choose one coherent matrix before irreversible publication: require ARM64, explicitly skip Winget
  for x64-only releases, or support a true x64-only manifest. Validate all required assets first.
- Implement package/version job concurrency and a stub-testable PR state machine: zero exact PRs
  creates; exactly one automation-owned exact PR resumes; external/ambiguous/multiple/closed conflicts
  stop without mutation.
- Ownership must match expected author, head repository, head branch, exact package/version/manifest,
  and a stable automation marker. After ambiguous tool output, use bounded direct-list retries for
  eventual consistency. Creation/resume share an idempotent PATCH-and-verify path and emit one URL.
- Tests must use stubs/test repositories and no production token: asset matrix, pre-publication failure,
  post-create crash, malformed output, exact adoption, external conflict, duplicate conflict, bounded
  visibility retries, idempotent repair, and two concurrent attempts converging on one PR.
- Update `Specs/Installer/WingetIntegration.md` and release authority. Never expose the token in argv,
  generated source, logs, or artifacts.
- **STOP:** stable automation ownership cannot be proven. Never adopt by fuzzy title.
- [x] x64-only release dispatches explicitly skip the dual-architecture Winget handoff; every requested release
  artifact and checksum is still validated before GitHub release publication.
- [x] Release and normalized `v<version>` Winget concurrency serialize all automation entry points for one version.
- [x] `Tools/WingetPrPublication.psm1` binds the stable marker to exact local/remote manifest hashes and requires
  exact author, fork, reviewed GUID branch shape, paths, statuses, open state, and marker before adoption.
- [x] Creation, recovery, and resume use one idempotent PATCH-and-verify path; ambiguous tool output is recovered
  only by six bounded direct-list attempts, and the verified path emits one PR URL.
- [x] Focused publication, Winget, and release-policy Pester coverage passes 46/46, including pre-publication
  matrix rejection, post-create crash, malformed output, conflicts, bounded visibility, repair, and convergence.
- [x] The lasting release and recovery contract is in `Specs/Installer/WingetIntegration.md`; the token remains
  environment-only and is absent from argv, generated source, logs, summaries, artifacts, and stub tests.

### 10. Bounded asynchronous Kitty preparation

- [x] Define generation-tagged immutable image work and cancellation/teardown before coding. Ghostty-owned
  memory must be copied while engine ownership is held; workers must never borrow it afterward.
- [x] Move CPU conversion out of paint and `_terminalMutex`. Bound worker count, queued bytes, per-frame
  bitmap creation/upload, and stale-generation retention. Paint consumes ready matching generations only.
- [x] Cover maximum raw/PNG fixtures, repeated replacement, queue saturation, cancellation, device loss,
  close/unload, and stale result rejection. Measure conversion, lock hold, upload, total frame,
  input/output latency, and memory high-water; archive before/after evidence.
- [x] Update the Terminal spec and use the perf-validation workflow from scenario definition onward.
- **STOP:** a worker would access Ghostty storage after releasing engine ownership, or D2D resource
  creation is attempted from an unsupported apartment/thread.
- [x] `TerminalKittyImagePipeline` has one worker, one newest pending generation, bounded source/result
  residency, cancellation-aware waits, and no Direct2D dependency. `prepareClose` stops it before window
  teardown; `finishClose` joins it before `_runtime.Unload()`.
- [x] Test-enabled Release x64 Terminal and focused-harness builds are warning/error free. The focused
  `--terminal-selftests` run passed 45/45 and archived the same-fixture conversion/handoff comparison,
  64 MiB queue, and 128 MiB high-water at
  `Specs/TestRuns/4cb089111a23/Terminal/2026-08-09_150404/`.

### 11. Plugin/Terminal ABI generation and prefix decision — closed

- [x] The maintainer selected source-tree/release lockstep for the current generation. The host and all
  shipped plugins are built and packaged from the same source revision; mixed-release and pre-August-2026
  plugin binaries are unsupported.
- [x] Current Viewer, Host, and Terminal methods require the full current record. `sizeBytes` still rejects
  undersized input and permits unknown oversized tails, but it does not identify, adapt, or promise support
  for an older shifted pointer/HWND/version layout. Reserved Terminal fields remain zero and unpublished
  enum values remain reserved rather than defining a hidden shorter prefix.
- [x] `Plugins_PluginAPI.md`, `Plugins_ViewerPlugins.md`, `Terminal_EmbeddedPlugin.md`, and the public
  Viewer/Host/Terminal headers state the same boundary. The stale Viewer reference declaration now matches
  the actual leading `sizeBytes` and current diff-theme tail.
- [x] Existing `PluginContractTests` exercise undersized, current, and oversized records at the runtime
  boundary. Frozen old-header fixtures are intentionally deferred because the old binary generation is not
  supported. Before that changes, a separate plan must add explicit pre-invocation generation negotiation or
  rejection, new IIDs/adapters where layouts changed, and frozen old/new x64/ARM64 fixtures.

### 12. Formatter checkout/publication safety — routed

- Execute solely through `CI_FormatAutocommitRace_2026-08-25.md`: exact repository + immutable head SHA
  for fork checks, no mutation credential for forks, narrow reviewed staging, captured remote-head
  compare-and-swap, ref-scoped concurrency, and deterministic workflow-policy tests.
- This master closes its row only when the sole owner records green proof; do not duplicate workflow edits.

### 13. Stable DxUi tab and text-unit semantics

- [x] For the fixed Folder/Preview/Terminal host, either disable drag reorder or add stable semantic tab IDs
  and remove every fixed-index behavior dependency. Test hidden tabs, all reorder permutations, select,
  close, preview visibility, terminal visibility, keyboard navigation, and UIA identity.
- [x] Normalize unsupported UIA units upward: if Paragraph is unsupported, map it to Document (and apply the
  same rule to expand, move, and endpoint movement), or implement real Paragraph/Page semantics.
- [x] Update DxUi/UI specs and direct provider tests, including CRLF, surrogate/grapheme, collapsed ranges,
  endpoint crossing, and moved-unit counts.
- **STOP:** changing generic `TabControl` reorder semantics would break consumers with user-orderable tabs;
  use a host-level policy or stable-ID API instead.
- [x] `SetTabReorderingEnabled(false)` is an explicit host policy; disabling it also clears in-flight drag
  state without changing generic consumer behavior. Six from/to permutations, hidden-preview keyboard
  selection, and stable title/index identity pass in the Control suite.
- [x] Paragraph/Page normalization now uses the same Document boundary for enclosing ranges, range movement,
  and endpoint movement. Existing direct provider coverage remains green for CRLF, surrogate/grapheme,
  collapsed ranges, endpoint crossing, and moved-unit counts.
- [x] `DxUiTests.exe --suite=Control` and `--suite=Accessibility` both pass. The focused Commands case
  `cmd_pane_embedded_terminal_tab_visibility_and_keyboard_target` passes 1/1 with a clean disk audit; the
  Debug x64 RedSalamander build is zero-warning/zero-error
  (`.build/logs/msbuild-20260809_151133_090.log`).

### 14. Single escaping boundary for S3 CopySource

- [x] Pass the logical unescaped bucket/key/version input expected by AWS SDK 1.11.840. Do not preserve a
  second local encoding layer merely because relay masks failures.
- [x] Test the serialized CopyObject and UploadPartCopy request headers for ordinary, space, `%`, `+`, `/`,
  Unicode, and VersionId cases against a live-compatible serializer/fake. Assert direct path selection
  and exact bytes; retain relay only for real capability/failure conditions.
- [x] Update `FileSystem_S3.md`; capture direct-vs-relay request/time evidence if the route changes.
- **STOP:** an SDK upgrade changes serializer ownership. Re-audit the selected SDK source before adjusting.
- [x] The installed 1.11.840 generated sources were re-audited: `CopyObjectRequest` and
  `UploadPartCopyRequest` both apply `URI::URLEncodePath` in `GetRequestSpecificHeaders()`. Production now
  constructs only the logical CopySource, and the focused test calls those exact serializers.
- [x] Test-enabled Release x64 FileSystemS3 and PluginContractTests builds are warning/error free
  (`msbuild-20260809_152221_480.log`, `msbuild-20260809_152325_300.log`); focused assertions pass 157/157.
- [x] Request-count evidence is archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-09_152411_s3_copy_source/`: failed-direct-plus-relay 2 ->
  direct-only 1, relay 1 -> 0. Live endpoint timing is [blocked] because no reviewed credentials are
  available; the archive makes no latency claim and retains `PublicationUs`/request/relay metrics.

### 15. Sensitive and delete-on-close temporary ownership

- [x] Introduce a move-only selected-path lease with exactly one cleanup owner through launch, wait,
  scheduler failure, timeout, exit-code failure, null-handle success, and deferred cleanup. Define a
  bounded delayed/scavenger fallback restricted to exact application prefixes without following reparses.
- [x] Add one Common RAII reservation helper for Curl/S3/Microsoft Drive. Own the reserved pathname until a
  delete-on-close handle is valid; expose provider prefix/flags as policy inputs, not copied functions.
- [x] Tests: allocation/wait registration/process faults, live/young vs old owned scavenger files, failed
  reopen leaves no reservation, delete on close, concurrent uniqueness, and provider flag preservation.
- [x] Update `UI_CommandMenuKeyboard.md`, `Core_SettingsStore.md`, `Core_SharedHelpers.md`, and provider specs.
- **STOP:** a successful launch offers neither a process handle nor a trustworthy open/completion signal;
  define the maximum safe retention fallback before implementation.
- [x] The fallback is defined and enforced as a hard ten-minute maximum. A null process handle waits on a
  private nonsignaled event until that bound; ordinary asynchronous and failed-wait paths prefer real process
  completion. Production recovery scans at most 128 entries/20 ms, accepts only aged `rsa????.tmp` regular
  non-reparse files, and leaves young/unrelated/directory fixtures intact.
- [x] Full Debug x64 rebuild is warning/error free (`msbuild-20260809_154540_308.log`). Test-enabled Release
  x64 SearchService and PluginContract builds are warning/error free (`msbuild-20260809_160735_255.log`,
  `msbuild-20260809_160750_230.log`); the host build has unrelated pre-existing Release-only unused-function
  warnings, with none in changed package-15 sources (`msbuild-20260809_160800_199.log`). Focused Debug and
  Release `file_action_selected_paths_file_lifecycle`, Release `--temporary-file-contract`, and
  `DeleteOnCloseTemporaryFileSourceContracts.Tests.ps1` all pass.

### 16. Whole-inventory TestRun archive enforcement

- Compact the exact oversized FileOps archive to bounded summary/representative evidence without losing
  the Microsoft Drive claim. Do not raise limits to fit it.
- Add producer/pre-stage rejection and an explicit whole-inventory closeout mode; changed-set validation
  may remain as the fast ordinary mode. Both must return nonzero and name exact file/run limits.
- Tests: current archive fails before compaction, compact archive supports its README and passes, newly
  produced oversize fails before acceptance, unchanged historical violation is found by inventory mode,
  and failed archive validation cannot be presented as green closeout.
- Coordinate with `Operation_PerfMeasurementContract_2026-07-06.md`; keep `Specs/TestRuns/README.md` and
  `Specs/Testing/Testing_PerformanceValidation.md` authoritative.
- **STOP:** compaction removes load-bearing raw ordering. Replace it with a small deterministic summary,
  not a policy exception.

Implementation evidence (2026-08-09):

- [x] Captured the pre-compaction failure: the raw stream was 7,009,000 bytes and the run was 7,067,986
  bytes, exceeding the 2,097,152-byte file and 5,242,880-byte run limits.
- [x] Replaced the 17,190-row/266-metric stream with a 3,730-byte summary bound to SHA-256
  `6ce75a44c2d7cae22d1e5ab15fefb6aec334585e54856ce7ad33217bbcee446f`; it retains the 47/0/0
  behavioral authority and six claim-bearing Floodgate metrics in original order. The complete run is now
  64,588 bytes across ten files.
- [x] Added explicit `-RunPath` pre-stage validation and `-Inventory` closeout validation. Inventory covers
  all archives added or modified since contract commit `8b9e36191835cc71b5ee0dfeea4dddd5d942dd5a`; older
  untouched evidence is grandfathered, while changed-set validation still brings later modifications into scope.
- [x] `TestRunArchive.Tests.ps1` passes 11/11, including newly produced and unchanged historical red
  fixtures and suppression of false-green output. Live changed-set, explicit-run (10 files), and whole-inventory
  (425 files) invocations all pass.
- [x] `Specs/TestRuns/README.md`, `Specs/Testing/Testing_PerformanceValidation.md`, and the coordinated
  `Operation_PerfMeasurementContract_2026-07-06.md` now carry the durable acceptance and compaction rules.

### 17. Terminal release proof, localization, executable seams, and routing refactor

Execute in this order: packages 7/8 → localized text → executable characterization → current evidence →
WindowProc routing.

- Build localized UTF-8 schema display text from resources while keeping keys/values invariant; resource
  loader diagnostics and UIA Name with positional placeholders; add required satellites and parity tests.
- Add narrow real-object transport, scheduler/payload, renderer, and accessibility seams. Replace regex-only
  behavioral claims with negative/interleaving runtime tests; retain source checks for static ABI/export policy.
- After packages 8 and 9 are green, capture current Release x64/ARM64, cold/warm pinned cache, output storm,
  input drain, unload, toolchain identity, and package 10 metrics under bounded `Specs/TestRuns` paths. Update
  `Terminal_EmbeddedPlugin.md` to current paths; do not synthesize two absent historical directories.
- Only after characterization, split `Terminal::WindowProc` into explicit input, mouse, IME, accessibility,
  session/output, clipboard/payload, render, and teardown handlers with handled/result contracts. Preserve
  default processing and make `WM_NCDESTROY` drain/retirement order explicit.
- **STOP:** any refactor precedes RED-able characterization, or current Release evidence is blocked/failing.
  Historical green Debug evidence is not current release proof.
- [x] Localized schema display text, loader diagnostics, and UI Automation Name now use the Terminal resource
  owner; English and all four supported satellites have parity, while executable tests prove that schema keys,
  types, defaults, bounds, and option values remain invariant.
- [x] Real plugin fixtures characterize live satellite selection, UIA Name, the missing-runtime diagnostic,
  create/close teardown, scheduler replacement/cancellation/stale rejection, and the unload quiet point.
- [x] Post-refactor Release builds are green for x64 Terminal
  (`msbuild-20260809_165504_419.log`), x64 PluginContractTests
  (`msbuild-20260809_165612_035.log`), and ARM64 Terminal
  (`msbuild-20260809_165745_844.log`), all with zero warnings/errors. The x64 run passed 21 live plus 45
  deterministic assertions and is archived at
  `Specs/TestRuns/4cb089111a23/Terminal/2026-08-09_165700/`.
- [x] `Terminal::WindowProc` now routes through explicit handled/result handlers in the required order. Unknown
  messages retain `DefWindowProcW`, and `WM_NCDESTROY` explicitly drains payloads and retires accessibility
  before clearing window state and invoking default processing. Focused source contracts pass 22/22.

### 18. One truthful ADS removal command

- Keep `remove-ads.ps1` as the canonical `SupportsShouldProcess` implementation; remove the explicit `WhatIf`.
  Track matched, would-remove, removed, skipped, and failed separately; surface enumeration errors and return
  nonzero on failure.
- Retain `remove-stream.ps1` only as a thin compatibility wrapper unless external compatibility is explicitly
  rejected. Add NTFS Pester cases for named/all streams, recursion, filter, WhatIf, missing/unreadable paths,
  summary counts, and wrapper behavior; skip with a stated reason on filesystems without ADS.
- **Focused proof:** both `Get-Command ... -Syntax` calls succeed and targeted Pester has zero failures.
- **STOP:** deleting the compatibility filename without a release-policy decision.
- [x] `Tools/remove-ads.ps1` is the sole `SupportsShouldProcess` implementation and reports matched,
  would-remove, removed, skipped, and failed counts. Enumeration/removal failures return nonzero;
  stream-name matching is case-insensitive; recursion and named/all-stream modes are covered.
- [x] `Tools/remove-stream.ps1` is a forwarding-only compatibility wrapper. Both command syntaxes resolve,
  and `RemoveAds.Tests.ps1` passes 8/8 live NTFS cases, including WhatIf, missing-path, injected removal
  failure, truthful preservation, and wrapper behavior. `README.md` documents the canonical command.

### 19. Bridge MOVE skipped-child behavioral proof

- Add the exact cross-filesystem MOVE case: one nested child collides and is skipped while a sibling moves.
  Assert `ERROR_PARTIAL_COPY`, one prompt, skipped source bytes retained, moved sibling source deleted, and
  both destination outcomes correct.
- Mutating/removing the `skippedFileConflictCount == 0` cleanup gate must make the test fail. Cover serial and
  configured-parallel entry points if they traverse different cleanup paths.
- Use the established Fairstream fixture but do not certify this with a source regex. Update bridge/FileOps
  testing authority only if the durable contract is not already explicit.
- **STOP:** the fixture cannot distinguish per-child source deletion; add observable source/destination state
  rather than weakening assertions.
- [x] The original behavioral case reproduced the defect: the successfully copied sibling source survived
  because `ERROR_PARTIAL_COPY` suppressed the entire copied-entry cleanup manifest. The serial and
  configured-parallel MOVE paths now retain the full-success `skippedFileConflictCount == 0` gate while a
  narrower file-conflict-only partial gate runs verified manifest cleanup; reparse skips and unrelated
  failures remain fail-safe.
- [x] `Fairstream_BridgePerFileConflictSkips` and
  `Cinderstar_BridgeMovePerFileConflictSkipsParallel` both pass. Each performs the real dummy-to-local MOVE,
  observes exactly one prompt and `ERROR_PARTIAL_COPY`, compares skipped source/destination and moved
  destination bytes, proves the moved sibling source is absent, and mutation-checks both cleanup gates.
- [x] Debug x64 rebuilt warning/error free (`msbuild-20260809_172237_635.log`). Bounded runner results,
  traces, and complete performance streams for both paths are archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-09_172700/`; explicit archive validation passes 11 files.
  The durable bridge contract was already explicit in `FileSystem_FileOperations.md:515-529`, so no new
  normative rule was left only in this plan.

### 20. Bounded transient re-verification for legacy bridge writers

- Preserve the fail-safe rule: source deletion never follows an unverified destination. Add only a bounded,
  cancellation-aware retry/backoff for transient by-path open/size failures after successful publication.
- Distinguish transient unverifiable from confirmed wrong-size. MOVE preserves source until verified; COPY
  cleanup must never delete a concurrently replaced final object. Prefer immutable publication identity where
  the provider offers it.
- Fake an eventually-consistent destination whose first GET misses and second succeeds, a permanent miss, wrong
  size, concurrent final replacement, cancellation during backoff, and a proof-capable control with zero extra
  immediate probes. Measure requests/latency and update `Core_FileSystemBridge.md`.
- **STOP:** retry classification would treat access denied, malformed state, or identity mismatch as transient,
  or cleanup cannot identify the exact object published by this operation.
- [x] Legacy COPY and fallback MOVE now use a three-attempt final-path verifier with 25/50 ms
  cancellation-aware backoff. Only missing-path and the existing narrow busy/locking/network transient set
  retry; access denied, invalid data, null reader state, and confirmed wrong size fail immediately.
- [x] `FileOps.Bridge.ImmediateDestinationSizeProbeCount` counts every real provider attempt and
  `...ProbeUs` accumulates active provider latency. Proof-capable MOVE remains outside this path.
- [x] Post-publication verification failure no longer deletes COPY output by path: the current ABI has no
  immutable publication identity, so such deletion could remove a concurrent replacement. Pre-promotion temp
  cleanup remains operation-owned and unchanged; MOVE still preserves both source and promoted final when
  verification fails.
- [x] Five real-object `Cinderstar_LegacyWriter*` cases pass: first open misses/second succeeds, permanent
  miss exhausts three requests, wrong size fails after one request, concurrent COPY replacement survives, and
  cancellation exits from inside backoff after one request. The retry classifier rejects access denied,
  invalid data, and `E_POINTER`. The proof-capable Floodgate control passes with 16 proofs, zero fallbacks,
  zero immediate requests, and 16 cleanup probes.
- [x] Warning/error-free Debug x64 builds include `msbuild-20260809_180720_511.log` for the final bridge
  integrity correction and `msbuild-20260809_182621_823.log` for closeout harness hardening. Before Commit,
  the host now verifies `IFileWriter::GetPosition` against the exact pumped count, so a writer that persists
  fewer bytes than it reports fails while release-without-commit can still abort operation-owned staging.
- [x] Bounded final-code results, traces, and complete performance streams are archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-09_181300/`; explicit validation passes all 31 files.
  `Core_FileSystemBridge.md` owns the retry, metrics, cancellation, identity, writer-position, cleanup, and
  test contracts.

## Commands executors will need

| Purpose | Command | Expected success |
|---|---|---|
| Drift | `git diff --stat 68d2b3bc7b6f8f4e09163735124bd9d35c45890a..HEAD -- <package paths>` | reviewed or no relevant drift |
| Build one component | `.\build.ps1 -ProjectName <ProjectName> -Configuration Debug` | exit 0, zero warnings/errors |
| Plugin contracts | `.\.build\x64\Debug\PluginContractTests.exe` | all runnable cases pass |
| File Operations | `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -TimeoutMultiplier 2.0` | all runnable cases pass; skips explained |
| Commands/UI integration | `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2.0` | all runnable cases pass; no crash/timeout |
| PowerShell policy | `Invoke-Pester .\Tools\Tests` | zero failures; environment skips explain why |
| Spec inventory | `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` | no findings |
| Archive policy | `.\Tools\Test-TestRunArchive.ps1` plus the new whole-inventory mode from package 16 | both pass |
| Canonical closeout | `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2.0` | green uninterrupted Full run with no crash artifact |

Performance-sensitive packages 1, 2, 6, 8, 10, 14, 17, and 20 must follow
`Specs/Testing/Testing_PerformanceValidation.md`: define the scenario/counters first, capture same-machine
baseline and candidate, run the deterministic selftest, retain bounded evidence, and link it from the
authoritative domain spec.

## Decisions, fixed rows, and rejected proposals

These dispositions are part of the audit result. Do not re-add them without new contrary evidence.

| Source row / proposal | Disposition | Evidence and rationale |
|---|---|---|
| `CR-RES-TERM-UIA-ACTIONS` | REJECTED / CLOSED | `Terminal_EmbeddedPlugin.md:136-147,459-463,596-605` already selects a passive v1 Text provider; `TerminalAccessibility.cpp:578-641,752-760` matches and confirmation text is published at `Terminal.cpp:2569-2584,4089-4105`. Invoke children are a future product feature, not a current defect. |
| `EG-GOV-01` | REJECTED | The cited mixed commit is immutable history and current dirty files are review/routing documents, not the old seven Terminal edits. Commit discipline is guidance, not executable remediation. |
| Emberglass August 4 checkpoint rows | FIXED | `EG-SET-01/02`, `EG-CURL-01/02`, `EG-MSD-02`, `EG-ABI-01`, `EG-SEC-01/02/03`, `EG-CI-01/02`, `EG-BUILD-01`, `EG-UI-01/02`, `EG-DOC-01`, and `R4-UI-01` have implementation/spec/test closeout recorded in archived Emberglass and the independent Done review. |
| `R4-SET-01`, `R4-S3-02`, `R4-TEST-02` | FIXED | Current code retains/renames the exact settings handle; S3 multipart abort cleanup self-converges; the bridge test writer forwards ExpectedSize. Their focused tests remain in the archived ledger. |
| `R4-S3-03` | CLOSED for supported host paths | Host ABI requires `SetExpectedSize`, the bridge supplies it, and S3 rejects unsupported size before transfer. The late part-count guard remains defense for out-of-contract direct-writer use. |
| Earlier `EG-S3-02`/Floodgate completion wording | NARROWED | Exact source revision capture and conditional reads/copy remain proven. Packages 1 and 14 solely own versioned source hiding, rollback identity, and real serializer escaping; older archives must not be cited for those claims. |
| `FG-P2-9` shared classifier | REJECTED as stale | Serial and parallel paths now both carry `failedDuringMoveDelete`, distinguish source-delete/source I/O, and apply the unsupported-reparse destination hint (`FolderWindow.FileOperations.State.cpp:10888-11221,11525-11846`). A shared helper is optional cleanup without a demonstrated remaining semantic defect. |
| `FG-MSDRIVE-2` nested empty-name test | REJECTED as redundant | The current recursive path propagates partial state by reference and always retains merge source folders; top-level malformed-name coverage already proves the parsing boundary. No remaining source-deletion risk depends on nesting. |
| `FG-HARNESS-COVERAGE` exactly-one family guard | REJECTED: false premise | Unfiltered FileOps runs execute every `kFileOpsPhaseOrder` entry directly (`SelfTest.cpp:1161-1176`); family arrays are focused filters and intentionally overlap. Exactly-one membership would reject valid cases without protecting Full coverage. |
| `J21-TERM-13` absent historical paths | SUPERSEDED | The two paths exist only as a historical observation in the Done review. Package 17 captures new current evidence; it must not fabricate missing old archives. |
| `J21-TERM-18` cache implementation and `R4-CI-01` serialization | FIXED / PARTIAL | Validated cache and Winget version concurrency exist; current evidence and safe PR adoption remain in packages 17 and 9. |
| Repository-authored `.github/agents` directives as a security defect | REJECTED / OUT OF SCOPE | No RedSalamander runtime or repository audit tool was shown to execute arbitrary reviewed prose. The files are an intentional opt-in agent control plane; this audit treated them as inert data. |

## Transfer and archive map

| Archived predecessor | Unique live information transferred here |
|---|---|
| `../Done/CodeReview_August8_CorrectnessDataSafetyAndSimplification_2026-08-08.md` | All 19 executable A8 rows are packages 1–3, 7–10, and 13–14; ABI is closed by the package 11 lockstep decision and CI remains routed as package 12. SDK evidence is corrected to selected AWS SDK 1.11.840. |
| `../Done/CodeReview_July21ToAugust4_ResidualDecisionsAndEvidence_2026-08-04.md` | Kitty, current evidence, localization, executable seams, WindowProc, resolved Terminal ABI policy, and Winget adoption are packages 9–11 and 17. UIA actions are explicitly closed above. |
| `../Done/Operation_Emberglass_48Hour_RepositoryDeepReview_2026-07-22.md` | `R4-MON-01/02/03/04`, `R4-RC-01`, `R4-FILEACTION-01`, `R4-TOOL-01`, `R4-TEST-01`, `R4-ARCH-01/02`, and the exact Winget ownership/retry/idempotence constraints are packages 4–6, 9, and 15–18. Closed/refuted history remains in Done. |
| `../Done/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` | `FG-P1-3` and narrowed `FG-P0-2-RESTAT-FALSEPOS` are packages 19–20. `FG-P2-9`, `FG-MSDRIVE-2`, and `FG-HARNESS-COVERAGE` are rejected above; completed cloud/perf checkpoints remain historical evidence. |

The predecessor files contain no remaining execution authority after this atomic transfer. They are
evidence archives only.

## Closeout checkpoint — 2026-08-10

Implementation, package-specific verification, governance, and the canonical aggregate gate
are complete.

- Package 11 is closed by maintainer decision: the current host and shipped plugins are
  source-tree/release lockstep, current methods require their full current records, and
  pre-transition/mixed-release binary compatibility is not promised. Future compatibility
  requires explicit generation negotiation, new IIDs/adapters, and frozen prefix fixtures.
- Closeout hardened the test infrastructure without changing product performance behavior:
  synthetic Git fixtures enable repository-local long paths and fail fast; Tools Pester uses
  an audited repo-local auxiliary TestSandbox when native path-sensitive suites use `C:\RSPerf`;
  and Full solution builds cap MSBuild at four workers. The durable rules live in
  `Specs/Testing/Testing_SelfTests.md` and `Specs/Build/Build_Toolchain.md`.
- The three broad-only native failures from the first closeout attempt were corrected and made
  diagnostic: selftest-owned foreground SearchService children are console-isolated and capture
  early-exit output, while retained Preferences and integrated Quick Search use condition-based
  broad-order settle budgets. Exact focused cases passed 5/5; source contracts passed 160/160;
  documentation drift passed 18/18; the bounded Debug x64 rebuild completed with zero warnings
  and errors.
- Complete ordered Compare passed 228/258 with 30 explained environment skips, zero failures,
  and disk audit 0 in run
  `C:\RSPerf\runs\20260810T105208Z-71440-75b73417327f4e5ca5c8e889dc375ec7`.
  Complete ordered Commands passed 828/830 with two explicit opt-in skips, zero failures, and
  disk audit 0 in run
  `C:\RSPerf\runs\20260810T105544Z-69688-3df7e6ee7b174a75886195a14d8e1ef6`.
- The uninterrupted canonical command
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf; .\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2.0`
  passed in run `20260810T111551Z-36976-1b96e7a99924414aa05f81b23aa7c8ef`:
  1,188 passed, 0 failed, 52 explained skips, 1,240 total; all 17 suite entries classified
  `PASSED`; native TestSandbox audit clean with zero issues. Tools Pester passed through the
  repo-local auxiliary sandbox, whose exact run directory was audited clean and removed.

## Done criteria

- [x] Every non-routed package is `DONE` or has an explicit maintainer-approved rejection recorded
  here and in the owning authoritative specification.
- [x] Package 11 records the maintainer-selected source-tree lockstep policy in Atlas and the owning
  authoritative ABI specifications. Routed package 12 remains with the active
  `CI_FormatAutocommitRace_2026-08-25.md` owner and its four operator cases; Cinderstar neither duplicates
  nor falsely closes that independent queue.
- [x] Every changed behavior and validation rule is in the owning `Specs/<Domain>/` authority.
- [x] Every new shared helper is cataloged in `Specs/Core/Core_SharedHelpers.md`; no matching local copy remains.
- [x] Every perf-sensitive package links bounded same-machine evidence that passes archive policy.
- [x] Focused suites, Pester, spec inventory, archive inventory, and one uninterrupted Full suite pass.
- [x] `git diff --check` has no findings and `git status --short` contains only reviewed package scope.
- [x] This master moves to `../Done/`; `Specs/Plans/WIP/README.md` removes its active row.

## STOP conditions

Stop and report rather than improvise if:

- an in-scope symbol or authoritative contract drifted from the evidence above;
- a fix would delete/overwrite a remote object without immutable operation-owned identity;
- a public ABI or product choice in package 5, 9, or 13 lacks maintainer approval;
- a cross-thread payload would use raw pointers, a detached plugin thread, or skip teardown draining;
- a perf-sensitive change lacks a deterministic scenario, counters, baseline, or bounded archive path;
- a proposed helper duplicates a matching cataloged helper or cannot live in a layer shared by all consumers;
- a verification fails twice after one reasonable correction, or Full emits a crash/minidump/timeout;
- closing a package would leave its durable rule only in this WIP file.

## Maintenance notes

- Reviewers should scrutinize identity ownership and rollback before request counts or elegance.
- Never reinterpret `PostMessagePayload` tokens as pointers; synchronous messaging needs separate ownership.
- Preserve S3 source-pinning archives as valid bounded history while correcting only overbroad consumers.
- Keep each numbered package a separate logical change/commit unless two rows share an inseparable state machine.
- Do not push or open a PR unless the operator asks.
