# Operation Emberglass — Whole-Repository Deep Review and July 15-to-HEAD Remediation

> **ARCHIVED 2026-08-09 — SUPERSEDED AS AN EXECUTION QUEUE.** The ten unique live rows,
> Winget recovery constraints, S3 contrary-evidence boundary, exact tests, performance gates,
> and STOP conditions are transferred to
> `../WIP/Operation_Cinderstar_UnifiedRepositoryRiskRemediation_2026-08-09.md`. Completed,
> refuted, and historical rows remain here as evidence only.

**Status:** Archived — live remainder transferred to Operation Cinderstar
**Review date:** 2026-07-22; expanded 2026-07-30 (Europe/Paris)
**Original reviewed HEAD:** 91484dd315277feeafc4ee357749548a3d4c19ea
**Rolling 48-hour cutoff:** 2026-07-20T10:56:43.4306975+02:00
**First in-window commit:** f243e77ecca44480769d5f645d22c38f87a8864f
**Comparison base:** ea55ff11c3bb5cf626f9f710fdd61fee542e89e0
**Comparison window:** ea55ff11c3bb5cf626f9f710fdd61fee542e89e0..91484dd315277feeafc4ee357749548a3d4c19ea
**July 15 refresh base:** 98146d41de8df4846c31df7dd7425653ee81a8ab
**July 15 refresh HEAD:** a00ec99a43fe6daa936d6f05eaa08e438b1bdd46
**July 15 refresh window:** 98146d41de8df4846c31df7dd7425653ee81a8ab..a00ec99a43fe6daa936d6f05eaa08e438b1bdd46
**Historical primary owner:** this plan at review time; Operation Cinderstar now owns the accepted remainder
**Review mode:** source-read-only; this report and its WIP routing-index entry are the only intended repository changes

> **Terminal routing update — 2026-08-03:** Terminal v1 is complete. References
> below to modified `Specs/Plans/WIP/Terminal_*.md` files describe the reviewed
> historical worktree and intentionally remain snapshot facts. The plans are
> now archived under `Specs/Plans/Done/`; current Terminal authority is
> `Specs/Terminal/Terminal_EmbeddedPlugin.md` and the current executive section
> of `Specs/Terminal/TerminalEngineDecision.md`.

## 1. Outcome

### 2026-08-04 independent-review remediation checkpoint

The July 21-to-August 4 independent review revalidated this ledger and completed the
following routed implementation rows: `EG-SET-01`, `EG-SET-02`, `EG-CURL-01`,
`EG-CURL-02`, `EG-MSD-02`, `EG-ABI-01`, `EG-SEC-01`, `EG-SEC-02`, `EG-SEC-03`,
`EG-CI-01`, `EG-CI-02`, `EG-BUILD-01`, `EG-UI-01`, `EG-UI-02`, `EG-DOC-01`, and
`R4-UI-01`. The repairs include retained-handle/long-path settings recovery, strong
Curl writer ownership and per-connection limiting, fail-closed ranged responses and
packed records, least-privilege/direct release publication, pinned toolchain provenance,
visual-order focus, bounded accessibility state, and shared animated geometry.

`R4-CI-01` is only partially complete: version-scoped serialization is implemented,
but adopting and resuming an exact-version open Winget PR remains owned solely by
`CodeReview_July21ToAugust4_ResidualDecisionsAndEvidence_2026-08-04.md`. This checkpoint
supersedes the live labels for those exact rows below; unrelated Emberglass and inherited
rows retain their existing state.

The original Emberglass refresh reviewed every first-party code/build path changed in the
rolling 48-hour window and then read the surrounding publication, ownership,
cancellation, teardown, UI, test, specification, and workflow scopes in detail. The
2026-07-30 continuation expanded that review to every commit from July 15 through
`a00ec99a4`, the current worktree, and the enclosing existing code. The combined ledger
contains **20 distinct Emberglass findings**:

- 8 P1 findings: possible remote/settings data loss or unintended overwrite, incorrect
  file-operation semantics, lifetime use-after-free, and CI credential/trust-boundary
  exposure.
- 10 P2 findings: settings recovery completeness, accessibility state/tab order,
  release atomicity/automation, privileged tool pinning, build-identity drift,
  Microsoft Drive response validation, plugin-ABI bounds validation, and documentation
  correctness.
- 2 P3 findings: per-connection Curl delete throughput and commit/worktree review
  granularity.
- No P0 finding was confirmed.

Four Emberglass P1 rows are closed (EG-MSD-01, EG-S3-01, EG-S3-02, and EG-FOPS-01). The live
Emberglass queue therefore contains 16 rows: 4 P1, 10 P2, and 2 P3, in addition to the
explicitly retained inherited R4 rows.

Of the original 15 rows, 7 were introduced or materially widened in the window and 8
were pre-existing defects found in the code immediately surrounding a changed path.
The July 30 continuation adds one process defect introduced in the expanded window and
four pre-existing defects in touched surrounding code. The historical four-day ledger
remains below so no open work is lost. One earlier closeout,
R4-SET-01, is contradicted by a newly proven external-writer interleaving and is
reopened with remediation owned by EG-SET-02. R4-S3-01 remains closed; EG-S3-01 is a
separate direct-transfer publication path. R4-SEARCH-01 closed on 2026-07-22 with
store-scoped writer ownership and deterministic security/lifetime proof.
EG-MSD-01 closed on 2026-07-22 with mutation-time path identity validation plus an
eTag-conditional backup PATCH; its durable contract and proof are routed to the
Microsoft Drive spec and Operation Floodgate while the remaining Emberglass rows stay live.
EG-S3-01 closed on 2026-07-22 with conditional direct-transfer publication.
EG-S3-02 closed on 2026-08-04 with exact ETag/VersionId propagation through copy, relay,
and MOVE deletion plus bounded test-enabled Release evidence.
EG-FOPS-01 closed on 2026-07-30 with current-conflict cache eligibility, deterministic
sequential/parallel COPY and MOVE proof, and a durable File Operations contract.

**2026-08-08 contrary-evidence correction:** EG-S3-02 remains closed only for exact source
revision capture and conditional reads/copy parts. Its former MOVE-delete completion statement
is not proof that the ordinary source path disappears: deleting the captured `VersionId` can
make an older version current. The same follow-up review also proved that rollback journals only
destination keys and can restore/delete a concurrent writer's current object. New executable
ownership is A8-S3-01 and A8-S3-02 in
`CodeReview_August8_CorrectnessDataSafetyAndSimplification_2026-08-08.md`; do not reopen or
duplicate the already-correct source-read pinning slice here.

The refresh also adds materially new evidence to four already-owned rows without
duplicating implementation ownership:

- R4-CI-01 also needs per-version serialization and an adopt/resume state machine.
- R4-TOOL-01 is reproduced by the duplicate common WhatIf parameter generated by
  SupportsShouldProcess plus an explicit WhatIf declaration.
- R4-TEST-01 is reinforced by a single approximately 7 MB archived JSONL payload.
- CI format drift, partial formatter failure, and new-branch/force-push base handling
  are routed to active plan 012-format-autocommit-race.md.
- EG-S3-02 proves that destination-safe publication is insufficient when COPY/MOVE
  does not pin the source ETag or VersionId through every copy part and final deletion.
- EG-MSD-02 proves that a ranged reader can accept shifted, empty, or inconsistent
  `206` content without validating `Content-Range`.
- EG-ABI-01 proves that the public packed `FileInfo` path-list helper reads ABI-owned
  fields before the containing record is known to be in bounds.
- EG-CURL-02 proves that one connection's delete limit unnecessarily serializes
  unrelated connections despite the existing per-connection keyed limiter.
- EG-GOV-01 records the reviewability boundary exposed when a focused runtime repair
  is committed with unrelated Terminal plans and workspace configuration.

The central simplicity problem is still not a shortage of abstractions. It is that
several paths have a check, identity, or owner that stops one boundary too early:

- cloud conflict checks precede the actual mutation/publication;
- S3 transfer plans previously retained a source key and size but not the immutable source revision; EG-S3-02 now closes that boundary;
- a returned writer retains the producer as a raw pointer instead of a COM owner;
- public packed records and ranged responses are trusted before their full structural
  contract is validated;
- a cached file-operation choice bypasses the eligibility rule that governed the
  original prompt;
- settings recovery records only a successful read stamp, not the existence and
  identity of a source that failed while being read;
- write authority is broad at workflow scope and checkout credentials persist before
  the publishing step;
- hosted UI controls are created in descriptor order rather than visual navigation order.

The repairs below therefore favor one mutation-time condition, one durable owner, one
eligibility function, one explicit workflow state machine, and one visual ordering pass.
They avoid speculative framework work.

## 2. Executor contract

### 2.1 Starting state

Before taking any work package:

1. Start from a clean worktree.
2. Confirm whether HEAD still contains the cited code. Original Emberglass line numbers
   are anchored to 91484dd315277feeafc4ee357749548a3d4c19ea; July 30 addendum line numbers
   are anchored to a00ec99a43fe6daa936d6f05eaa08e438b1bdd46; inherited R4 line numbers
   retain their original d596d751d44bc5cfec581f528929d446f286e168 anchor.
3. If the relevant code has materially changed, reproduce the finding against the new HEAD before editing. Do not mechanically apply the proposed shape to a different implementation.
4. Preserve unrelated user changes.
5. Read the domain skill named by AGENTS.md before editing that domain.

### 2.2 Simplicity rules for every repair

1. Prefer a local, explicit state transition over a new general-purpose framework.
2. Put the decisive conflict check at the publication/commit operation, not only in preflight.
3. Make ownership total: every exit either completes cleanup or transfers the resource to a clearly named owner.
4. Pass stable identities when presentation order can differ from storage order.
5. Keep hot-path cursors explicit rather than repeatedly rebuilding whole collections.
6. Extend Common only when policy, dependencies, ownership, error behavior, and performance semantics match across consumers.
7. Add a failure-first regression before changing the implementation.
8. Do not make an incidental cleanup that broadens a work package without an observed defect or a measured simplification.

### 2.3 Completion rules

A row is complete only when all of the following are true:

- the focused regression fails against the old behavior and passes after the repair;
- focused Debug builds and tests pass;
- relevant source-contract tests are updated without replacing runtime proof;
- performance-sensitive rows have deterministic measurements and a bounded archive under Specs/TestRuns;
- durable behavior is merged into the authoritative domain specification;
- routed owners are updated rather than duplicated;
- the final full suite passes;
- this plan is moved to Specs/Plans/Done and the WIP index is reconciled.

## 3. Scope and method

### 3.1 Repository coverage

The exact rolling window contains 20 commits, including 5 merges and 15 non-merge
commits. Its net diff has 236 changed paths, 66,107 additions, and 876 deletions. Most
of that volume is specifications and archived TestRun evidence. Removing plans,
specifications, generated/archive payload bodies, and vendor/generated output leaves
81 first-party code/build paths:

| Area | Changed paths |
|---|---:|
| GitHub workflows | 10 |
| Common | 12 |
| Plugins | 15 |
| RedSalamander application | 27 |
| Tests | 8 |
| Tools | 5 |
| Root build/dependency policy | 4 |
| **Total** | **81** |

Every one of those 81 paths received a diff review. Each behavior change then received
an enclosing-scope read through its callers/callees, ownership boundary, tests, and
authoritative specification. The deepest reads covered:

1. S3, Microsoft Drive, Curl, and bridge file writers, direct transfers, conflict
   handling, multipart cleanup, and object lifetime.
2. Settings read/recovery/publication, hot reload, revision stamps, backup, and
   non-cooperating external writers.
3. File Operations prompt caching, nested/top-level behavior, task lifecycle,
   accessibility state, hosted focus traversal, and theme overlay input geometry.
4. Release/Winget workflows, event trust, permissions, token exposure, resumability,
   release asset atomicity, toolchain installation, and pinned dependencies.
5. Focused selftests, source-contract tests, TestRun policy, documentation links, and
   already active/completed plans.

The pass also ran repository-wide targeted scans for prohibited catch-all handlers,
sprintf-family formatting, raw posted-message heap payloads, AddRef-plus-attach COM
ownership, detached plugin threads, unpinned dependencies, credential-bearing
workflows, and duplicated helpers. No new violation in those banned C++ pattern
families was confirmed.

This remains a risk-weighted repository review: changed code was exhaustive; surrounding
semantic reading followed every changed behavior and its trust/lifetime/publication
boundary. Unchanged, unrelated implementation bodies were inventoried and pattern-scanned
but were not all reread line by line.

### 3.2 Exclusions

- .build and other generated outputs were not reviewed as source.
- Third-party and vendored implementation bodies were excluded.
- Archived performance payload bodies were not semantically reviewed; archive manifests, policy, sizes, and cited evidence were reviewed.
- Generated design/mockup state was excluded.
- Existing findings already owned by an active plan were not copied into a new repair row unless this review found materially new contrary evidence.
- No implementation or runtime test execution was performed during the review. Commands
  below are executor gates, not claimed results.
- Web verification used primary sources only: GitHub's GITHUB_TOKEN and concurrency
  contracts, the exact pinned release action's source/defaults, and Microsoft's
  WingetCreate CI token contract. Local source, tests, and specifications are the
  primary evidence for every code finding.

### 3.3 July 15-to-HEAD continuation

The 2026-07-30 continuation selected
`98146d41de8df4846c31df7dd7425653ee81a8ab`, the final commit immediately before
July 15, as its comparison base. The range through
`a00ec99a43fe6daa936d6f05eaa08e438b1bdd46` contains 67 commits, including 57
non-merge commits. After excluding ordinary Markdown plans/docs and archived TestRun
bodies, the implementation/configuration/test diff still covers approximately 365 paths
and 58,000 changed lines.

The continuation split the review across:

1. Common, SettingsStore, the main application, SearchService, File Operations, and
   surrounding ownership/UI code.
2. S3, Microsoft Drive, Curl, MTP, Google Drive, 7-Zip, the public filesystem ABI, and
   cross-filesystem bridge tests.
3. Monitor, RedConfigure, viewer plugins, build/tooling, GitHub workflows, packaging,
   documentation, and evidence policy.

Every reported finding was reopened at its cited current source before inclusion. The
continuation also rejected stale candidates where the current code already carries the
required publication condition, mutation-time Microsoft Drive destination validation,
SearchService writer ownership, posted-payload teardown, or plugin threadpool module-pin
transfer.

At the refresh anchor:

- `master` and `origin/master` both resolve to `a00ec99a4`.
- No uncommitted source code exists. The worktree contains seven modified
  `Specs/Plans/WIP/Terminal_*.md` files only.
- The cached Keep Both repair in `a00ec99a4` is internally consistent and has focused
  sequential/parallel COPY and MOVE evidence; it remains closed as EG-FOPS-01.
- Production-source diff checks found no new prohibited `catch (...)`, non-PoC
  `sprintf_s`/`swprintf_s`, raw posted-message heap payload, or AddRef-plus-attach
  pattern.
- The review reproduced `Tools/remove-ads.ps1` failing PowerShell command discovery
  because `SupportsShouldProcess` and an explicit `WhatIf` parameter define the same
  common parameter.
- No new full-suite execution was claimed. A fresh full suite remains a mandatory
  closeout gate after remediation.

## 4. Priority and ownership matrix

Priority meanings:

- **P1:** address before the affected behavior is treated as release-ready.
- **P2:** schedule in the next hardening cycle.
- **P3:** valid but low-frequency; keep visible and define a product limit.

Effort is an implementation estimate: S is up to roughly two focused days, M is several days, L is a cross-cutting slice. Risk is the regression risk of the repair, not the severity of the defect.

| ID | Priority | Window | Finding | Impact | Effort | Repair risk | Confidence | Primary coordination |
|---|---|---|---|---|---|---|---|---|
| EG-MSD-01 | P1 closed 2026-07-22 | Introduced/widened | A cached Present destination hint can rename and later delete an object that moved elsewhere | Remote data loss/wrong-object mutation | M | High | High | Closed here; durable contract in Microsoft Drive spec; proof routed to Floodgate |
| EG-MSD-02 | P2 | Pre-existing in touched scope | Microsoft Drive accepts ranged `200`/`206` bodies without validating `Content-Range`, returned offset, or exact length | Wrong bytes or premature EOF accepted as successful reads | M | Medium | High | This plan; Microsoft Drive spec; coordinate Floodgate |
| EG-S3-01 | P1 closed 2026-07-22 | Pre-existing in touched scope | Direct directory COPY/MOVE published unconditionally after its no-overwrite refresh | Unconsented object overwrite | M | High | High | Closed here; durable contract in S3 spec; proof coordinated in Floodgate |
| EG-S3-02 | P1 closed 2026-08-04 for source-read pinning; source-removal/rollback conclusions narrowed 2026-08-08 | Pre-existing in touched scope | Direct S3 COPY/MOVE carried source key/size but no immutable ETag or VersionId through multipart copy and relay reads | Mixed/truncated destination reads | L | High | High | Exact source-read/copy pinning remains closed here. A8-S3-01/A8-S3-02 own the newly proved rollback identity and versioned source-removal defects. |
| EG-FOPS-01 | P1 closed 2026-07-30 | Introduced | Cached Keep Both bypasses top-level-only eligibility for nested conflicts | Split/duplicated COPY or MOVE results | M | Medium | High | Closed here; durable contract in File Operations spec |
| EG-CURL-01 | P1 | Pre-existing in touched scope | Returned Curl writers retain FileSystemCurl as a raw pointer | Use-after-free/plugin unload hazard | S | Medium | High | This plan; Curl spec |
| EG-CURL-02 | P3 | Pre-existing in touched scope | A delete batch folds every endpoint's concurrency into one global minimum despite an existing keyed limiter | One serial endpoint throttles unrelated connections | M | Medium | High | This plan; Curl spec; performance evidence required |
| EG-SET-02 | P1 | Introduced | Guarded bad-file backup compares a path stamp, closes the handle, then renames by path | Newer revision displaced to .bad; canonical path may later receive defaults | M | High | High | Owns reopened R4-SET-01 external-writer boundary |
| EG-SEC-01 | P1 | Pre-existing in touched scope | Winget event strings are interpolated into PowerShell before validation | Workflow-code injection before secret use | S | Low | High | This plan; Winget spec |
| EG-SEC-02 | P1 | Pre-existing in touched scope | Repository-write credentials are granted to build/test jobs and persisted by checkout | Dependency/build compromise can write repository state | S | Medium | High | This plan; plan 012 for format job |
| EG-SET-01 | P2 | Introduced | Early settings read failures discard evidence that the source exists | Defaults load silently and ordinary save cannot converge | S | Medium | High | This plan; SettingsStore spec |
| EG-SEC-03 | P2 | Pre-existing in touched scope | A mutable WingetCreate install executes at the publication-token boundary | Unreviewed tool updates run with access to upstream credentials | S | Low | High | This plan; Winget spec |
| EG-UI-01 | P2 | Introduced | File Operations hosted controls are created footer-first and queue buttons reverse visual order | Incorrect Tab/Shift+Tab navigation | S | Low | High | This plan; UI File Operations spec |
| EG-UI-02 | P2 | Introduced | File Operations accessibility status history is never pruned | Unbounded retained task state | S | Low | High | This plan |
| EG-CI-01 | P2 | Pre-existing in touched scope | Releases created with GITHUB_TOKEN cannot trigger the release.published Winget workflow | Documented automatic Winget submission never starts | S | Medium | High | This plan; Winget spec |
| EG-CI-02 | P2 | Pre-existing in touched scope | ZIP/checksum assets are published before a separately invoked MSIX upload | Public release can expose a partial requested asset set | S | Medium | High | This plan; release/MSIX specs |
| EG-BUILD-01 | P2 | Pre-existing in touched scope | Official builds install mutable Visual Studio 2026 package versions | Unreviewed compiler drift or spontaneous build failure | M | Medium | High | This plan; create Build_Toolchain spec |
| EG-DOC-01 | P2 | Introduced | Changed review/selftest documentation contains a dead plan route and contradictory Release guidance | Engineers follow stale execution contracts | S | Low | High | This plan |
| EG-ABI-01 | P2 | Pre-existing in touched scope | The public packed `FileInfo` path-list helper dereferences record fields before proving the header/name payload is in bounds | Plugin-controlled out-of-bounds host read/crash | S | Low | High | This plan; Virtual File System and shared-helper specs |
| EG-GOV-01 | P3 | Introduced in expanded window | A focused runtime repair was committed with unrelated Terminal plans and workspace configuration while later Terminal plan edits remain dirty | Broad review/rollback surface and unclear commit ownership | S | Low | High | This plan; WIP index and commit discipline |
| R4-SET-01 | P1 reopened | Cross-window interaction | Settings recovery can move a newer valid commit to the bad-file backup | Newer revision displaced; canonical configuration can become defaults | M | High | High | Remediation owned by EG-SET-02 |
| R4-S3-01 | P1 closed 2026-07-22 | Four-day code | S3 no-overwrite is only a preflight check | Unconsented object overwrite | M | Medium | High | Operation Floodgate |
| R4-S3-02 | P1 | Four-day code | Multipart abort queue does not self-converge | Leaked uploads/unload stall | M | Medium | High | Operation Floodgate |
| R4-MON-01 | P1 | Four-day code | Monitor open budgets encoded bytes but retains/duplicates decoded UTF-16 | Memory spike, slow cancel/close | L | High | High | This plan; fresh Track 6 contrary evidence |
| R4-MON-02 | P1 | Four-day code | Search cap plus eviction permanently hides retained middle matches | Incorrect search/navigation | M | Medium | High | This plan; fresh Track 6 contrary evidence |
| R4-RC-01 | P1 | Four-day code | Rectangular paste increments raw rows/cultures, not visible identities | Wrong cells edited | M | Medium | High | This plan; coordinate Rosetta Lantern |
| R4-SEARCH-01 | P1 closed 2026-07-22 | Repository-wide | Predictable permissive Global event can block SearchService startup | Low-privilege service DoS | M | High | High | Closed here; Parallax remains adjacent pipe-policy review |
| R4-MON-03 | P2 | Repository-wide | Show IDs changes display offsets without rebuilding matches | Stale navigation/copy state | S | Medium | High | This plan |
| R4-FILEACTION-01 | P2 | Repository-wide | Selected-path temporary files have incomplete ownership | Temp leak/sensitive path retention | M | Medium | High | This plan |
| R4-UI-01 | P2 | Four-day code | Animated theme overlay paints transformed but hit-tests final geometry | Invisible/misaligned click capture | S | Low | High | This plan |
| R4-CI-01 | P2 | Four-day code | Winget release rerun aborts after PR creation | Non-resumable release workflow | S | Medium | High | This plan |
| R4-TOOL-01 | P2 | Repository-wide | Duplicate ADS removal scripts have a broken WhatIf contract | Tool cannot load/incorrect summary | S | Low | High | This plan |
| R4-TEST-01 | P2 | Four-day evidence | A committed TestRun violates file and run-size policy | Review/storage contract bypass | S | Low | High | Operation PerfMeasurementContract |
| R4-TEST-02 | P2 | Four-day tests | Bridge writer decorator hides ExpectedSize | Important production path untested | S | Low | High | Operation Floodgate |
| R4-ARCH-01 | P2 | Repository-wide | WSL registry/catalog logic is duplicated and reads strings unsafely | Undefined read/behavior drift | M | Medium | High | This plan; coordinate Terminal |
| R4-ARCH-02 | P2 | Repository-wide | Three delete-on-close temp helpers drift; Curl leaks failed reservation | Temp leak/duplicate policy | S | Low | High | This plan |
| R4-MON-04 | P2 | Repository-wide | Monitor Save As performs conversion/write on UI thread under document lock | UI stall, blocked ingestion | L | High | High | This plan; semantics remain BP-B5 |
| R4-S3-03 | P3 parked | Four-day code | Fixed 64 MiB parts fail only at part 10,001 | Late failure near 625 GiB | S/M | Medium | High | Operation Floodgate |

## 5. Recommended execution order

### Wave 0 — ownership and failing proofs

1. EG-MSD-01, EG-S3-01, and EG-FOPS-01 are reproduced and closed. Reproduce EG-SET-02 and EG-SEC-01 with
   deterministic seams before changing production logic.
2. Reproduce EG-S3-02 with deterministic source-revision changes between planning,
   multipart parts or relay reads, and final MOVE deletion before changing production
   logic.
3. Keep EG-S3-02 implementation ownership in this plan and record its completed cloud
   proof in Operation Floodgate; do not create a duplicate Floodgate work package.
   The completed Microsoft Drive mutation-time proof is already recorded there. Route
   the formatter/base-SHA observations to
   012-format-autocommit-race.md instead of creating duplicate owners.
4. Treat EG-SET-02 as the exact contrary-evidence continuation of R4-SET-01. Do not
   reopen the rest of the completed settings track.
5. Preserve the inherited R4 routes to PerfMeasurementContract, Rosetta Lantern, and Terminal. R4-SEARCH-01's
   completed startup-ownership repair is coordinated with Parallax but does not resolve or reopen PAR-1 pipe
   impersonation policy.

### Wave 1 — data integrity and trust boundary

1. EG-S3-02 — closed 2026-08-04
2. EG-SET-02
3. EG-SEC-01
4. EG-SEC-02
5. EG-CURL-01

These rows can destroy or misdirect user data, violate COPY/MOVE conflict semantics,
outlive their producer, or expose protected workflow capabilities.

### Wave 2 — exact recovery, UI, and release behavior

1. EG-SET-01
2. EG-SEC-03
3. EG-UI-01
4. EG-UI-02
5. EG-CI-01
6. EG-CI-02
7. EG-BUILD-01
8. EG-DOC-01
9. EG-MSD-02
10. EG-ABI-01

### Wave 3 — inherited data integrity, convergence, and boundedness

1. R4-RC-01
2. R4-S3-02
3. R4-MON-01
4. R4-MON-02
5. R4-FILEACTION-01

### Wave 4 — inherited UI/test/workflow precision

1. R4-MON-03
2. R4-UI-01
3. R4-CI-01, including the Emberglass concurrency/adoption expansion
4. R4-TOOL-01
5. R4-TEST-01
6. R4-TEST-02

### Wave 5 — inherited consolidation and responsive export

1. R4-ARCH-01
2. R4-ARCH-02
3. R4-MON-04
4. Decide and document R4-S3-03.

### Wave 6 — measured throughput and review governance

1. EG-CURL-02
2. EG-GOV-01

These rows do not justify delaying a safety repair. Execute Curl concurrency only with
same-machine per-key measurements. Close the governance row by keeping future runtime,
plan-only, and workspace-policy changes in independently reviewable commits; do not
rewrite already published history merely to improve commit shape.

## 6. Detailed work packages

### Emberglass refresh packages

### EG-MSD-01 — Revalidate the Microsoft Drive destination identity at mutation time

**Priority:** P1 — closed 2026-07-22
**Window classification:** introduced/materially widened by 5242378b7
**Impact:** an object moved or renamed by another client can be renamed as a backup and
later deleted, even though it no longer occupies the requested destination
**Effort / repair risk / confidence:** M / High / High

#### Evidence and failure interleaving

- Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:4135-4137 performs
  the per-child destination probe before the remaining remote work.
- The full Present result is retained at :4170-4171 and passed into
  MoveOrRenameItem at :4274-4285.
- MoveOrRenameItem freshly reads the source and source parent at :4351-4367, but
  accepts a Present destination hint at :4387-4396 whenever its ID is nonempty. That
  skips the by-path destination GET at :4403-4405.
- PrepareOverwriteBackup is entered at :4444-4446. Its helper at :4010 mutates the
  hinted item by stable item ID without first proving that its parent/name/eTag still
  describe the destination path.
- The backup item ID is deleted after successful source publication at :4512-4515.
- ItemMetadata already carries id, name, parentId, and eTag at :110-117, so the
  missing information is not a representation problem; it is a validation-boundary
  problem.

Concrete interleaving:

1. RedSalamander probes destination /A/target and caches Present item D.
2. A second Drive client moves D to /B/important and optionally renames it.
3. RedSalamander enters MoveOrRenameItem, trusts the cached D ID, and skips a fresh
   path lookup.
4. The overwrite backup PATCH renames D wherever D now lives.
5. The source is installed at /A/target.
6. Cleanup deletes D by ID, destroying the object that the other client moved away.

The older implementation already had a smaller GET-to-PATCH race. Commit 5242378b7
materially widened it by allowing a substantially earlier Present probe to authorize
the later mutation. Specs/FileSystem/FileSystem_MicrosoftDrive.md:172-179 calls the
per-child probe a just-in-time freshness boundary; it is not a mutation-time boundary.

#### Simplest safe repair

Make backup preparation conditional on the exact identity observed for the destination:

- carry the expected eTag, parentId, and name with the Present observation;
- immediately before backup mutation, re-resolve the destination path and require the
  same ID, parent, name, and revision;
- issue the mutating request conditionally against the observed eTag;
- if any identity/revision check changes, return the normal conflict/revision result
  and re-enter conflict handling; never mutate the stale item ID.

A fresh GET without a conditional mutation is insufficient because it merely moves the
race. If the provider cannot make the backup PATCH conditional, fail closed rather
than treating the earlier hint as overwrite authority.

#### Required tests and closeout

- Add a deterministic hook after the outer Present probe and before backup mutation.
- Through a second client/session, move the destination item to another parent and
  change its name/content while the first operation is paused.
- Cover MoveItem and MoveItems under both pre-authorized overwrite and prompt-granted
  Overwrite, with sequential and parallel host execution. Add Retry-to-Present if that
  transition can reach the same branch.
- Assert that the moved-away item's ID, parent, name, content, and eTag are untouched;
  no backup is created for it; MOVE source deletion does not occur on conflict.
- Add the same barrier for a same-path replacement with a new item ID and for an
  in-place eTag change.
- Update FileSystem_MicrosoftDrive.md so freshness is defined at conditional mutation,
  then route the proof to the Microsoft Drive portion of Operation Floodgate.

#### Progress — DONE 2026-07-22

- The test fake now models drive-item eTags, content, external move/rewrite, and conditional PATCH rejection. The
  deterministic hook returns the original `/dst/Foo/conflict.txt` metadata to the first operation, then moves that
  exact item to `/elsewhere/important.txt` and changes its content/eTag before backup preparation continues.
- Test-enabled Release `FileSystemMicrosoftDrive` and `PluginContractTests` builds pass with 0 warnings / 0 errors in
  `.build/logs/msbuild-20260722_154325_642.log` and `.build/logs/msbuild-20260722_154415_988.log`.
- Before any production repair, `PluginContractTests` is RED: Microsoft Drive reports **167 passed / 2 failed**.
  Injection/result setup passes; the failing assertions prove the moved-away identity/content/revision and MOVE
  source are not preserved, and that the stale path performs PATCH/DELETE traffic instead of one mutation-time
  validation GET with no mutation. Compact evidence is archived at
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_1545_msdrive_mutation_red/`.
- Microsoft Graph's drive-item update contract documents `If-Match` on PATCH and a non-mutating `412 Precondition
  Failed` response for an eTag mismatch. The repair will therefore pair a fresh by-path identity/revision check with
  the conditional backup PATCH; the earlier `Present` hint alone will never authorize mutation.

#### Closeout — 2026-07-22

- Backup preparation now rejects incomplete observations, probes the rollback name, re-resolves the destination path,
  requires exact ID/parent/name/eTag equality, and sends the backup PATCH with `If-Match`. A moved, replaced, revised,
  missing, or inconclusive destination and a conditional `412` all return the normal conflict result before source
  mutation or backup deletion.
- Deterministic fake-Graph coverage proves moved-away, same-path replacement, and in-place revision races; public
  `MoveItem`/`MoveItems`; pre-authorized and prompt-granted overwrite; sequential and simultaneous host execution;
  preserved destination identity/content/eTag; retained MOVE sources; and zero cleanup DELETEs on every conflict.
- Test-enabled Release `FileSystemMicrosoftDrive` builds with 0 warnings/errors in
  `.build/logs/msbuild-20260722_155823_614.log`; the final test-enabled Debug full solution rebuild is also clean in
  `.build/logs/msbuild-20260722_161458_386.log`. Complete Release and freshly rebuilt Debug
  `PluginContractTests` pass with Microsoft Drive **202 / 0**. The tests-disabled production Release provider is
  clean in `.build/logs/msbuild-20260722_161313_843.log` and exports no debug-selftest entry point. Source contracts
  pass **149 / 0**.
- Five same-machine Release samples preserve the no-conflict merge improvement (`83/84 -> 67/68` GETs; 19.78%
  median fixed-latency reduction). The overwrite candidate adds exactly one GET per child (`83/84 -> 99/100`) and
  measures 19.60% median overhead, below the 30% advisory ceiling, with mutation counts unchanged. Evidence is under
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_1600_msdrive_overwrite_validation/`.
- The lasting behavior and metric contract is in `Specs/FileSystem/FileSystem_MicrosoftDrive.md`; the coordination
  proof is recorded in `Specs/Plans/Done/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. Emberglass
  remains WIP for its other findings, so this umbrella plan is not moved to `Done/`.

### EG-S3-01 — Make direct directory transfer publication conditional

**Priority:** P1
**Window classification:** pre-existing defect in the S3 paths changed by
52c622086 and 91484dd31
**Impact:** COPY or MOVE with no-overwrite semantics can replace an object created after
the last destination refresh
**Effort / repair risk / confidence:** M / High / High

#### Evidence and failure interleaving

- Plugins/FileSystemS3/FileSystemS3.Directory.cpp:1661-1670 refreshes destination
  metadata just before direct transfer.
- The actual direct copy at :1831-1832 is unconditional.
- CopyS3ObjectWithFallback at :973-1002 has no destinationMustNotExist parameter and its
  fallback cannot preserve that condition.
- Plugins/FileSystemS3/FileSystemS3.S3.cpp:914-934 issues the small server-side copy
  without a destination publication condition.
- Multipart copy completion at :977 explicitly passes false for the no-overwrite
  condition.
- Commit 91484dd31 correctly made writer-based publication conditional. These direct
  directory COPY/MOVE fast paths do not pass through that repaired commit point.

Concrete interleaving:

1. The directory operation refreshes key K and observes it missing.
2. Another client creates K.
3. RedSalamander performs CopyObject or completes multipart copy unconditionally.
4. K is silently replaced. For MOVE, the source can then be deleted, so retry cannot
   reconstruct the user's original choice.

This contradicts Specs/FileSystem/FileSystem_S3.md:161, which promises that a
destination created after planning is not silently overwritten. The defect predates
the window; the metadata and writer-publication changes exposed the inconsistent
direct-transfer branch. It is distinct from closed R4-S3-01.

#### Simplest safe repair

- Thread destinationMustNotExist through CopyS3ObjectWithFallback and every direct-copy
  branch.
- For multipart copy, publish with the same conditional completion used by the
  no-overwrite writer.
- If an endpoint cannot conditionally publish a small server-side CopyObject, route
  the operation through a conditional writer/relay path or fail closed. Do not retain
  the fast path by weakening conflict semantics.
- Normalize a late precondition failure to ERROR_FILE_EXISTS and re-enter the existing
  conflict decision state machine.
- Delete a MOVE source only after conditional destination publication succeeds.

#### Required tests and closeout

- Inject destination creation exactly after the just-in-time refresh and before the
  publication request.
- Cover small CopyObject, multipart copy, and any fallback/relay branch for both COPY
  and MOVE.
- Assert that the injected destination bytes and metadata remain unchanged and that a
  failed MOVE retains the source.
- Cover a standards-compliant endpoint and one that explicitly rejects/does not support
  the condition; require an explicit fail-closed or configured compatible fallback.
  Document the standards-compliance assumption because a nonconforming endpoint that
  reports success while silently ignoring a condition has no generic in-band detector.
- Archive the focused request/byte/latency comparison because the safe fallback may
  change the direct-transfer performance path.
- Update FileSystem_S3.md and own implementation/testing in Operation Floodgate.

#### Progress — 2026-07-22 failure-first checkpoint

- The test-enabled fake S3 graph now creates the competing destination inside the direct copy call, after both the
  whole-plan and just-in-time destination refreshes observed the key missing. The seam runs through the existing
  `ExecuteCopyOrMove` COPY and MOVE paths; production publication behavior is unchanged.
- Test-enabled Release `FileSystemS3` and `PluginContractTests` rebuild with 0 warnings / 0 errors in
  `.build/logs/msbuild-20260722_165809_115.log` and `.build/logs/msbuild-20260722_170054_580.log`.
- `PluginContractTests` is RED exactly at the new contract: S3 directory-transfer probes report **16 passed / 2
  failed**. COPY replaces the injected bytes; MOVE replaces them and deletes its source. Compact evidence is under
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_1705_s3_direct_publication_red/`.
- The installed AWS SDK and the official S3 conditional-write contract support destination `If-None-Match: *` on
  `CopyObject` and `CompleteMultipartUpload`. The repair can therefore retain one publication request on compliant
  endpoints; relay publication will carry the same condition, and a conflict or unsupported conditional path will
  fail closed.
- `destinationMustNotExist` now reaches small `CopyObject`, multipart completion, and relay upload. Conditional
  `409`/`412` outcomes normalize to `ERROR_FILE_EXISTS`; interactive operations refresh and re-enter the existing
  conflict decision path, and no-callback operations fail closed. MOVE source deletion remains after successful
  publication only.
- The deterministic matrix covers small, forced-multipart, and relay publication for COPY and MOVE, preserves the
  concurrently injected destination plus MOVE source, verifies multipart abort, and covers an endpoint that
  explicitly rejects conditional relay publication. Late-conflict Skip and Overwrite controls prove the existing
  decision state machine remains authoritative. Test-enabled Release `PluginContractTests` is green, with the S3
  directory-transfer slice at **37 passed / 0 failed**; source contracts pass **149 / 149**.
- Production instrumentation emits aggregate publication request/time, conditional-request, relay-fallback, and
  conflict counts. The protected 16-object same-binary Release comparison retains exactly 16 requests in both runs,
  changes the conditional subset from `0 -> 16`, and records `251,041 us -> 253,786 us` (+2,745 us, 1.09%), below
  the 30% timing-noise ceiling. Green evidence and raw JSONL are archived under
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_1740_s3_direct_publication_safe/`.
- The test-enabled full Release/x64 solution rebuild completed with 0 errors (10 existing solution-wide warnings) in
  `.build/logs/msbuild-20260722_172620_876.log`. The durable mutation, fail-closed, telemetry, and performance
  contracts are merged into `Specs/FileSystem/FileSystem_S3.md`; coordination is reconciled in Operation Floodgate.
  Emberglass remains WIP for its other findings, so the umbrella plan is not moved to `Done/`.
- Final shipped-configuration and broad gates are clean: the tests-disabled Release provider build is 0 warnings /
  0 errors; the final Debug/x64 rebuild is 0 / 0; complete Debug provider contracts pass; and Debug FileOps reports
  **111 passed / 0 failed / 20 expected environment skips** with no failure classifications. A repository-wide Full
  attempt completed Compare at 227 / 0 and recorded 371 / 0 Commands cases before its one-hour outer cap interrupted
  the next UI case, so that aggregate is classified inconclusive rather than as a product failure. The exact boundary
  and authoritative replacement gates are recorded in the green archive.

### EG-FOPS-01 — Apply cached conflict choices only where the action is eligible

**Priority:** P1
**Status:** Complete 2026-07-30
**Window classification:** introduced by da8f967b5
**Impact:** applying Keep Both to all conflicts can duplicate or split a later
directory COPY/MOVE instead of resolving only the colliding nested item
**Effort / repair risk / confidence:** M / Medium / High

#### Evidence and failure interleaving

- RedSalamander/FolderWindow.FileOperations.State.cpp:4217-4220 permits Keep Both
  only when the resolved destination equals operationDestinationPath and the
  operation/bucket/cookie eligibility conditions also hold.
- Cached bucket decisions are consumed before a new prompt at :5809; the prompt path
  also returns cached decisions at :4226-4245.
- A cached Keep Both decision is converted to provider Skip and a whole-item retry
  cookie at :5836-5843.
- The host retries the top-level sibling at :11073-11084 and :11662-11675.
- keepBothConflictDestinationPath is written at :5841 but has no reader, so it cannot
  constrain the retry to the nested path that actually collided.
- Specs/FileSystem/FileSystem_FileOperations.md:411-429 explicitly limits Keep Both
  to the top-level item. Nested provider conflicts retain provider-level semantics.
- RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp:6216-6259 covers one
  direct Keep Both case. The
  apply-to-all test uses Skip, and no test combines cached Keep Both with a nested
  merged-directory conflict.

Concrete failure:

1. A top-level collision is resolved as Keep Both plus Apply to all.
2. A later top-level directory begins merging successfully.
3. One nested child collides.
4. The cached Keep Both is consumed even though that nested conflict is ineligible.
5. The provider receives Skip; the host retries the entire partially processed root as
   folder (2).
6. COPY duplicates/splits output. MOVE can split the source tree between the original
   destination and the new sibling.

#### Simplest safe repair

Centralize one current-conflict eligibility function and evaluate it before returning
any cached decision. A cached Keep Both may be reused only when the current conflict
is the exact top-level operation item. For an ineligible nested conflict, ignore that
cached action and prompt/dispatch the provider contract normally. Remove
keepBothConflictDestinationPath if the corrected design still has no consumer.

Do not make provider Skip carry a second hidden meaning for nested Keep Both. The
whole-item retry belongs only to the eligible top-level state transition.

#### Required tests and closeout

- First choose Keep Both plus Apply to all, then force a nested collision in a later
  merged directory.
- Cover COPY and MOVE under sequential and parallel execution.
- Assert no whole-root retry occurs for the nested conflict, already processed children
  are not duplicated, and MOVE never splits the source tree.
- Cover cached Overwrite/Skip in the Exists bucket to prove only action-specific
  eligibility changes.
- Add source-contract coverage that every cache-return site calls the eligibility
  function before applying Keep Both.
- Update FileSystem_FileOperations.md with cache reuse semantics and retain the
  top-level-only invariant.

#### Closeout

- `IsCachedConflictDecisionEligible(...)` now centralizes the action-specific rule.
  Cached Keep Both is reused only for a per-item COPY/MOVE Exists conflict whose
  current destination exactly matches the item's `operationDestinationPath`.
  Every cache-return site uses the eligible loader; cached Overwrite and Skip retain
  their existing per-bucket behavior. The unused
  `keepBothConflictDestinationPath` field was removed.
- The deterministic
  `Phase9_ConflictPrompt_KeepBothNestedCacheEligibility` case first reproduced the
  defect against the unfixed production code: **2 passed / 1 failed**, with
  `copy-sequential` completing without the required nested prompt. After the repair,
  the case passes for COPY and MOVE at concurrency 1 and 2, requires two prompts,
  and proves no `tree (2)` retry, duplicate sibling, split MOVE tree, or skipped-child
  mutation.
- Debug evidence is green for the focused case (**3 / 0**), complete Phase 9
  (**11 / 0**), and complete File Operations suite
  (**112 passed / 0 failed / 20 expected environment skips**). Test-enabled Release
  focused evidence is **3 / 0**. The existing Commands conflict-prompt contract is
  also green (**1 / 0**), and source contracts pass **150 / 150**.
- The existing Apply-to-all convergence metric measured 1,014,807 us before the
  repair and 1,068,247 us, 1,123,071 us, and 1,219,868 us after it. These small,
  prompt-wait-dominated samples are directionally neutral within scheduling noise;
  they do not support a percentile claim. The new optimized Release matrix measured
  235,000-297,000 us across all four scenarios.
- Curated RED/GREEN, build, and performance evidence is archived under
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_2105_keepboth_nested_cache_eligibility/`.
  The lasting cache-reuse and regression contracts are merged into
  `Specs/FileSystem/FileSystem_FileOperations.md` and
  `Specs/Testing/Testing_TestCoverage.md`. Emberglass remains WIP for its other
  findings, so the umbrella plan is not moved to `Done/`.

### EG-CURL-01 — Give every returned Curl writer a strong COM owner

**Priority:** P1
**Window classification:** pre-existing defect in a Curl writer path exercised by the
window's short-success proof work
**Impact:** releasing the FileSystemCurl interface before Commit can dereference a
destroyed owner; plugin code/lifetime can also become inconsistent with the live writer
**Effort / repair risk / confidence:** S / Medium / High

#### Evidence

- Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:149-164 constructs
  TempFileWriter with a FileSystemCurl pointer and stores the raw owner at :326.
- TempFileWriter dereferences it during completion notification at :316-318.
- CurlStreamingWriter takes the same raw owner at :851-856, stores it at :1208, and
  dereferences it at :1058-1060.
- CreateFileWriter returns those objects with raw this at :1536 and :1559.
- The streaming writer's module pin at :873-879 keeps executable code mapped. It does
  not keep the FileSystemCurl object alive.
- FileSystemCurl.Shared.cpp:3384-3386 and :3494-3508 implement COM reference/instance
  accounting; the actual unload gate at :2081-2087 requires that instance count to be
  zero together with the pool/runtime conditions.
- There is no owner-to-writer reference, so a writer-to-owner strong reference does not
  create a cycle.

The new FTP short-success proof directly exercises CurlStreamingWriter. TempFileWriter
has the same pre-existing raw-owner defect discovered by the required surrounding-scope
review.

The returned IFileWriter is an independently held COM object. A caller may legally
release its factory/filesystem interface and then call Write/Commit. Both writer
implementations currently assume the opposite lifetime ordering.

#### Simplest safe repair

Replace the raw owner member in both returned writers with
wil::com_ptr<FileSystemCurl>, assigned from the constructor argument through normal COM
assignment. Keep the module pin where needed for callback-code lifetime; it solves a
different problem. Do not use AddRef plus attach.

This matches the already reviewed Microsoft Drive and MTP writer pattern documented in
Specs/Plans/Done/Code_PluginImprovementPlan.md.

#### Required tests and closeout

- Create each writer, release every external FileSystemCurl/IFileSystem reference, then
  complete and destroy the writer.
- Cover TempFileWriter success/failure/cancel and CurlStreamingWriter success,
  short-success, failure, and cancel.
- Assert completion notification remains valid, instance count stays nonzero until the
  writer is released, and unload cannot begin while the writer survives.
- Run the existing Curl fault-injection and plugin unload stress suites.
- Document the returned-object lifetime contract in the Curl filesystem spec.

### EG-SET-01 — Preserve source existence and identity across early read failure

**Priority:** P2
**Window classification:** introduced by the interaction with 4c71d1aa0 recovery
permission changes
**Impact:** an oversized, unreadable, or short-read settings file can load defaults
without requesting replacement, while later ordinary saves repeatedly fail revision
comparison against the still-existing source
**Effort / repair risk / confidence:** S / Medium / High

#### Evidence

- Common/Common/SettingsStore.cpp:382-385 clears sourceStamp before any file access.
- The file is opened at :387-393. Size-validation failures return at :395-405 and read
  failures return at :416-431.
- The source stamp is captured only after a fully successful read at :434-445.
- RecoverSettingsLoadFailure equates source preservation with sourceStamp.has_value()
  at :5525.
- The read-failure path enters that recovery at :5568-5571. Without a stamp it does
  not set ExplicitReplacementRequired at :5553-5558.
- LoadSettingsWithRecovery treats S_FALSE as a successful default load and prepares
  the save target at :5850-5853.
- The application accepts SUCCEEDED, including S_FALSE, as a default-settings startup
  at RedSalamander/RedSalamander.cpp:7702-7709. The replacement prompt is shown only
  for the explicit permission state at :7767-7790.
- A later normal save reaches the CAS comparison around SettingsStore.cpp:8144 with no
  expected stamp, but the original source still exists, so persistence cannot converge.

The recovery API currently collapses missing source, proven-present source with an
optional captured identity, and a presence-unknown access/metadata failure.

#### Simplest safe repair

Represent source observation explicitly:

- Missing: the file/path-not-found result proves the source did not exist at the read
  boundary.
- Present: an opened handle or reliable metadata proves existence; its stamp remains
  optional when identity capture itself fails.
- PresenceUnknown: access, sharing, or parent-path failure cannot prove either
  existence or absence.

For an opened file, capture a handle stamp before size/content validation and retain it
when those validations fail. Retain the current post-read stamp behavior and, when both
early and post-read stamps are available, compare them so a changing file is retried or
reported as a revision conflict rather than diagnosed from a mixed snapshot. The early
stamp is recovery identity, not proof of a successful stable parse. Present without a
stamp and PresenceUnknown must both stay save-blocked pending explicit
retry/replacement; neither may be treated as Missing.

Do not manufacture an empty expected stamp to make ordinary save overwrite an existing
source.

#### Required tests and closeout

- Load an oversized settings file and inject short ReadFile, access-denied/open-sharing,
  size-query, and stamp-query failures.
- Assert defaults may be returned only with an explicit source-preserved/replacement
  state; ordinary autosave must not silently overwrite or enter an unexplained
  revision-mismatch loop.
- Verify a genuinely missing file still follows first-run default/save behavior.
- Cover the same states through SettingsHotReload.
- Update Specs/Core/Core_SettingsStore.md with the three-state source observation
  contract and UI behavior for unstampable or presence-unknown sources.

### EG-SET-02 — Rename the exact diagnosed settings file through its retained handle

**Priority:** P1
**Window classification:** introduced by 4c71d1aa0; contrary evidence that reopens the
R4-SET-01 external-writer boundary
**Impact:** a non-cooperating external editor can replace the settings path after the
stamp comparison, displacing the newer revision into the bad-file backup; a later
default save can then replace the canonical configuration
**Effort / repair risk / confidence:** M / High / High

#### Evidence and failure interleaving

- Common/Common/SettingsStore.cpp:557-580 opens/stamps the source and closes the
  identity-bearing handle.
- The SettingsCommitLock acquired at :590-595 serializes RedSalamander writers only.
- The implementation compares the current path stamp at :597-605 and then calls
  pathname-based MoveFileExW at :608-617.
- Specs/Core/Core_SettingsStore.md:140-143 and :262-265 promise that the exact diagnosed
  revision is archived and that a revision mismatch modifies neither revision.

Concrete interleaving:

1. Recovery diagnoses invalid revision X.
2. Guarded backup acquires the application lock and compares path X successfully.
3. After the comparison handle closes, an external editor atomically replaces X with
   valid revision Y. That editor does not participate in SettingsCommitLock.
4. MoveFileExW resolves the path again and moves Y to the bad-file sibling.
5. The canonical path is now vacant; later default persistence can put defaults there
   while Y survives only under the bad-file name.

The 2026-07-22 R4-SET-01 repair correctly closed the cooperating-writer race but did
not make comparison and rename one filesystem transaction. Its closeout is therefore
reopened only for this exact external-writer boundary.

#### Simplest safe repair

- Open the diagnosed source once with the access needed for identity query and
  handle-based rename.
- Use sharing that excludes concurrent write/delete replacement for the lifetime of
  comparison plus rename. If an already-open writer makes that impossible, fail with
  revision/sharing conflict and leave both files untouched.
- Compare the retained handle's identity/stamp to the expected revision.
- Rename that same open handle with SetFileInformationByHandle
  FileRenameInfo/FileRenameInfoEx to the reserved bad-file sibling.
- Keep SettingsCommitLock to serialize cooperating application writers and backup-name
  reservation, but do not rely on it for external exclusion.

Do not re-open the path between comparison and rename. Do not broaden automatic
recovery into permission to overwrite an unidentifiable source.

#### Required tests and closeout

- Add a deterministic barrier after comparison and before rename.
- At that barrier, attempt to open new raw Win32 write/delete/replacement handles and
  require them to be denied until the retained handle has moved X.
- In a separate case, open a write/delete-capable handle before guarded backup begins.
  Retained-handle acquisition must then fail closed before comparison, with the
  external writer's revision Y still canonical and neither backup nor defaults written.
- In every successful backup, assert the bad-file sibling contains exactly X, never Y.
- Cover source deletion, backup collision, sharing violation, explicit replacement,
  and hot reload.
- Replace the R4-SET-01 closeout claim in Specs/Core/Core_SettingsStore.md with the
  retained-handle transaction contract; close EG-SET-02 and R4-SET-01 together.

### EG-UI-01 — Build hosted File Operations focus order from visual order

**Priority:** P2
**Window classification:** introduced by da8f967b5
**Impact:** keyboard users encounter footer actions before task controls and queued-task
buttons in the reverse of their visible order
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- RedSalamander/FolderWindow.FileOperations.Popup.cpp:5341-5490 appends footer button descriptors
  before task-row controls even though the footer is rendered below the tasks.
- Queued-task controls are appended Collapse, Move Down, Move Up at :6323-6359,
  opposite their visible left-to-right order.
- Hosted children are created in _buttons order at :4946-4963.
- Common/DxUi/DxUi.WindowHost.cpp:982-1014 collects hosted focusable children in
  that child order; Tab/Shift+Tab traversal uses the resulting order around
  :2689-2712.
- Specs/UI/UI_FileOperationsPopup.md:163-169 requires forward and reverse traversal in
  visual order.
- RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp:5823-5838 proves only
  that two hosted IDs are nonempty
  and distinct. It never asserts the actual traversal sequence.

The retained-mode tree is therefore internally consistent but is constructed from
logical descriptor insertion order, not from the visual navigation contract.

#### Simplest safe repair

After final layout, derive one explicit hosted-control navigation sequence:

- group controls by visual row from top to bottom;
- order each row by X in the active flow direction;
- retain a stable semantic tie-breaker for overlapping/equal positions;
- create/register hosted controls in that sequence, without changing paint z-order or
  command identity.

If child creation order must remain coupled to z-order, add a dedicated tab-order
surface to DxUi instead of sorting the paint collection. Do not hard-code a second list
of control IDs that will drift when task states change.

#### Required tests and closeout

- Assert the complete forward Tab sequence and exact reverse Shift+Tab sequence for
  running, queued, failed, completed, and mixed task rows plus footer actions.
- Include Collapse, Move Up, and Move Down in one queued row.
- Cover left-to-right and right-to-left flow if the host supports both.
- Prove dynamic task insertion/removal and popup resize rebuild the same logical order
  without losing current focus.
- Update UI_FileOperationsPopup.md with the ordering derivation, not merely a list of
  today's controls.

### EG-UI-02 — Prune File Operations accessibility status history

**Priority:** P2
**Window classification:** introduced by da8f967b5
**Impact:** a popup kept open across a long file-operation session retains one status
entry for every historical task ID
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- RedSalamander/FolderWindow.FileOperations.Popup.h:682 owns _announcedTaskStatuses.
- RedSalamander/FolderWindow.FileOperations.Popup.cpp:5046-5063 inserts/updates an
  entry for every rendered task.
- Cleanup at :2458-2495 builds the live task-ID set and prunes several other
  per-task maps, but never erases stale entries from _announcedTaskStatuses.

Task IDs do not recur within the popup lifetime, so this is not a stale-value
correctness bug. It is unbounded retained state on a UI object deliberately designed
to remain open while completed tasks are cleared and new work arrives.

#### Simplest safe repair

Prune _announcedTaskStatuses in the existing live-ID cleanup pass. Reuse the same live
set and erasure pattern; do not add a second timer, cap, or generational cache. Preserve
the current status value for every still-live task so accessibility announcements do
not repeat after an unrelated task is removed.

#### Required tests and closeout

- Add a deterministic stress selftest that cycles thousands of distinct tasks through
  add, status change, complete, and clear while retaining the popup.
- Expose a debug/test-only retained-entry count and assert it is bounded by live task
  count after each cleanup.
- Assert removing one task does not cause unchanged live tasks to be re-announced.
- Include popup teardown to prove all remaining announcement state is released.
- Document the bounded per-task state invariant in UI_FileOperationsPopup.md.

### EG-SEC-01 — Keep Winget event data out of generated PowerShell source

**Priority:** P1
**Window classification:** pre-existing defect in the changed Winget/release workflow
scope
**Impact:** a collaborator who can dispatch the workflow or publish a release can turn
a version/tag string into workflow code and prepare interception of the later upstream
publication token
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- .github/workflows/winget-release.yml:31-37 inserts the manual input and release tag
  directly into double-quoted PowerShell source through GitHub expression expansion.
- The semantic-version validation does not run until :40-42, after the runner has
  already parsed and executed that generated source.
- A crafted value can terminate the PowerShell string and execute commands in the
  validation step. Although that step has no Winget token, it can persist state through
  GITHUB_PATH, GITHUB_ENV, runner-temporary state, or generated executables.
- The later credential-bearing step at :209-246 supplies WINGET_TOKEN and invokes
  wingetcreate by command name at :244, giving an earlier injected path/executable a
  direct token-interception route.

This is not an anonymous-input claim: workflow dispatch and release publication require
repository capability. The boundary being crossed is between that capability and
protected workflow code/upstream-repository credentials that the actor may not be
allowed to read or modify directly.

#### Simplest safe repair

- Map every event/input string through the step env block, for example RAW_VERSION and
  RAW_TAG, and read it as data through PowerShell environment access.
- Validate and normalize before emitting any step output or using the value in a path,
  command, query, or API URL.
- Pass validated values as native argument-array elements; never splice them back into
  generated shell source.
- Keep WINGET_TOKEN scoped to the one command that needs it and resolve the expected
  executable to a verified absolute path before the secret enters the environment.

#### Required tests and closeout

- Add a workflow source-policy test rejecting direct string-valued github.event or
  inputs expression interpolation inside run blocks without banning legitimate typed
  boolean/numeric expressions.
- Exercise quotes, backticks, dollar expressions, semicolons, newlines, command
  separators, Unicode, and valid semantic versions; malicious values must remain inert
  strings and fail validation.
- Assert the untrusted corpus cannot cause PATH/environment mutation or executable
  creation before the secret-bearing step.
- Update Specs/Installer/WingetIntegration.md with the event-data boundary.

### EG-SEC-02 — Scope write authority to publishing jobs and credentials to publishing steps

**Priority:** P1
**Window classification:** pre-existing defect in changed CI/release workflow scope
**Impact:** a compromised dependency, package installer, formatter, build script, or
test can use checkout-persisted credentials to modify repository or release state
**Effort / repair risk / confidence:** S / Medium / High

#### Evidence

- .github/workflows/ci.yml:9-10 grants contents: write at workflow scope, so migration,
  build, and selftest jobs receive write authority even though only format publication
  needs it.
- .github/workflows/release.yml:22-23 likewise grants contents: write to version,
  build, package, validation, and publishing jobs.
- .github/workflows/build-reusable.yml:33-72 checks out with the default
  persist-credentials behavior before installing mutable toolchain packages and
  executing repository/dependency-controlled build logic.
- The checkout credential is stored for Git use; it need not be printed in an
  environment variable for a compromised build step to invoke an authenticated push.

The current shape violates least privilege across the largest and most dependency-rich
parts of CI. Fork pull-request token downgrades reduce some exposure but do not protect
push, manual, scheduled, or release executions.

#### Simplest safe repair

- Set workflow defaults to contents: read.
- Grant contents: write only to the exact format-publication and release-publication
  jobs.
- Use persist-credentials: false on every checkout that does not intentionally push.
- In the formatter flow, keep credentials absent during install/format/test and
  introduce explicit scoped authentication only for the final lease-checked push.
- In the release flow, keep build/package/validation jobs read-only and give the final
  release job only the permissions its API calls require.
- Audit actions for pull-requests, attestations, packages, and id-token permissions and
  declare only non-default capabilities actually used.

The format-job authorization mechanics belong in active
plans/012-format-autocommit-race.md; the general build/release narrowing belongs here.

#### Required tests and closeout

- Parse every workflow and assert a reviewed job-to-permission matrix.
- Reject contents: write at workflow scope and reject credential-persisting checkout
  in non-publishing jobs.
- Run the format and release dry-run paths with read-only defaults, then prove only the
  final intended job can authenticate a write.
- Update the CI and release specifications with the capability matrix and token
  lifetime.

### EG-SEC-03 — Pin and verify WingetCreate at the credential-bearing boundary

**Priority:** P2
**Window classification:** pre-existing defect in the changed Winget workflow
**Impact:** unreviewed WingetCreate updates execute with access to the upstream
winget-pkgs credential
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- .github/workflows/winget-release.yml:98-102 installs
  Microsoft.WingetCreate without a version constraint.
- The later step at :214-246 passes WINGET_TOKEN to that executable as a native CLI
  argument for PR creation or update.
- The workflow therefore changes privileged executable code independently of this
  repository's reviewed commit. Every upstream update reaches the token-bearing
  boundary without repository review; a compromised update would receive the same
  access.
- By contrast, the window's vcpkg baseline and GitHub actions use immutable commit
  identifiers; the privileged command-line tool is the exception.

#### Simplest safe repair

- Pin one reviewed WingetCreate version in repository policy.
- Install exactly that version from the intended source, relying on the source's
  package hash/signature validation.
- Resolve its absolute executable path and assert the reported version before any step
  receives WINGET_TOKEN.
- Supply the secret through WingetCreate's documented WINGET_CREATE_GITHUB_TOKEN CI
  environment contract and omit the token CLI argument:
  https://github.com/microsoft/winget-create
- Keep the token at the minimum upstream scope and only in the create/update command
  step; clear it from later steps.
- Make version updates explicit reviewed dependency changes, ideally with a scheduled
  notification rather than an automatic floating install.

#### Required tests and closeout

- Add a source-contract test requiring an explicit WingetCreate version and a
  pre-secret version assertion.
- Simulate a mismatched/missing executable and prove the workflow stops before the
  token-bearing step.
- Capture the final process argument vector and assert it contains no credential or
  token switch; prove only WINGET_CREATE_GITHUB_TOKEN is supplied to the verified
  process.
- Run a no-secret dry run through manifest generation and validation.
- Record the reviewed version and update process in WingetIntegration.md.

### EG-CI-01 — Invoke Winget submission directly after release publication

**Priority:** P2
**Window classification:** pre-existing defect in the changed release/Winget workflow
**Impact:** the documented automatic Winget submission does not run for releases
created by the repository's own release workflow
**Effort / repair risk / confidence:** S / Medium / High

#### Evidence

- .github/workflows/release.yml:399-416 invokes the pinned release action without an
  alternate token, so the action uses the workflow's github.token default.
- .github/workflows/winget-release.yml:3-5 waits for the release.published event.
- GitHub's official token contract states that events created with GITHUB_TOKEN do not
  normally start a new workflow; release.published is not among the documented
  dispatch/limited-PR exceptions:
  https://docs.github.com/en/actions/concepts/security/github_token
- The pinned release action documents token: github.token as its default:
  https://raw.githubusercontent.com/softprops/action-gh-release/3d0d9888cb7fd7b750713d6e236d1fcb99157228/README.md
- Specs/Installer/WingetIntegration.md:121-136 documents automatic submission after a
  published GitHub release.

The release can publish successfully while no Winget run is created. This is
deterministic platform behavior, not an intermittent event-delivery issue.

#### Simplest safe repair

Make Winget submission a reusable workflow and call it as an explicit dependent job
after release publication, passing the already validated version and only the required
secret. Keep release.published as an optional entry point for releases created by an
external human/tool if that behavior is still wanted. Do not solve recursion with a
broad personal access token merely to manufacture an event.

#### Required tests and closeout

- Add workflow-contract coverage proving release publication has a direct successful
  dependency edge to Winget submission and does not rely solely on token-generated
  events.
- Exercise manual Winget dispatch, release-workflow call, and optional external
  release.published entry points through the same validation body.
- Assert version and asset readiness are passed/checked once and secret scope remains
  in the reusable Winget job.
- Update WingetIntegration.md with the actual trigger graph.

### EG-CI-02 — Publish the requested release asset set atomically

**Priority:** P2
**Window classification:** pre-existing defect in changed release workflow scope
**Impact:** when MSIX is requested, the public release and checksum file can appear
before the MSIX upload; a later failure leaves a partial release contract
**Effort / repair risk / confidence:** S / Medium / High

#### Evidence

- .github/workflows/release.yml:399-407 first invokes the release action with portable
  ZIP/checksum assets. That invocation can create and publish the release.
- A second conditional invocation at :409-416 uploads MSIX.
- The pinned action deliberately creates a new regular release as a draft, uploads all
  files in that invocation, and finalizes afterward. Splitting the files across two
  invocations defeats that transaction:
  https://raw.githubusercontent.com/softprops/action-gh-release/3d0d9888cb7fd7b750713d6e236d1fcb99157228/src/github.ts
  https://raw.githubusercontent.com/softprops/action-gh-release/3d0d9888cb7fd7b750713d6e236d1fcb99157228/src/run.ts
- If the second action fails, the public release remains available and its checksum
  manifest may reference an MSIX that is absent.
- Specs/Installer/WingetIntegration.md:27-32 and
  Specs/Installer/Installer_Msix.md:90-99 describe the requested assets as one validated
  release set.

This is a workflow transaction split: validation is complete, but publication has two
independent externally visible commits.

#### Simplest safe repair

Use exactly one release-action invocation per mutually exclusive input branch:

- portable-only branch uploads ZIP plus checksums once;
- MSIX branch uploads ZIP, checksums, and MSIX together once;
- both branches fail on an unmatched required path.

The pinned action already supplies draft-then-finalize behavior for a new regular
release when all assets are in one invocation. Before a rerun, query the release state:

- no release or a repairable draft may enter the single publication transaction;
- an already published release with the exact expected asset/checksum set is an
  idempotent success;
- an already published release with a missing or mismatched set must fail for explicit
  repair without deleting/replacing assets one by one.

Set overwrite_files: false (or an equivalent non-destructive policy) on idempotent
published-release checks. The action otherwise deletes an existing asset before
uploading its replacement, which can create another partial published state.

#### Required tests and closeout

- Parse the workflow and assert one publication transition for each input branch.
- Simulate a missing/failing MSIX upload and require that no public release is visible.
- After success, compare the remote asset-name set and checksum manifest to the
  expected branch-specific set before Winget submission begins.
- Exercise rerun/adoption of a draft, exact published no-op, and mismatched published
  fail-without-mutation so repair does not create duplicates or remove assets.
- Update release/MSIX/Winget specs with the publication transaction.

### EG-BUILD-01 — Pin and record the compiler used for official binaries

**Priority:** P2
**Window classification:** pre-existing defect in the changed reusable-build workflow
**Impact:** the same repository commit can acquire a different compiler servicing
build, warnings, code generation, or package-install failure without a reviewed source
change
**Effort / repair risk / confidence:** M / Medium / High

#### Evidence

- .github/workflows/build-reusable.yml:38-72 installs the Visual Studio 2026 Build
  Tools/workload Chocolatey packages without version constraints.
- Official release jobs call that reusable build workflow.
- Directory.Build.props:66-67 constrains the SDK/toolset family but does not select an
  exact installed compiler servicing build.
- The vcpkg baseline in the same change set is a full immutable commit, showing the
  repository already applies stronger reproducibility policy to library dependencies.

This finding does not claim bit-for-bit reproducibility from a hosted Windows image.
It identifies an avoidable unreviewed compiler update at the official artifact
boundary.

#### Simplest safe repair

- Pin reviewed versions of the Visual Studio Build Tools and VC workload packages, or
  use an image/toolchain mechanism that exposes an equally exact compiler identity.
- Fail before build if MSBuild, cl.exe, Windows SDK, or toolset identity differs from
  policy.
- Record cl /Bv, MSBuild version, Windows SDK version, vcpkg baseline, runner image,
  architecture, and configuration in the release evidence.
- Update pins as explicit dependency changes with a focused build/performance
  comparison.

#### Required tests and closeout

- Add a workflow source-policy test requiring exact toolchain selectors.
- Inject an unexpected installed version and prove the workflow fails before compile.
- Verify x64/ARM64 Debug/Release resolve the intended identities.
- Attach the identity manifest to official artifacts.
- Create Specs/Build/Build_Toolchain.md as the durable compiler identity/update
  contract and link it from README.md and the cpp-build skill; do not leave the rule
  only in this plan.

### EG-DOC-01 — Repair changed review routing and selftest configuration guidance

**Priority:** P2
**Window classification:** introduced in the 48-hour documentation changes
**Impact:** contributors are routed to a nonexistent WIP plan and can incorrectly infer
that Release selftest builds are impossible
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- Specs/Reviews/README.md:21 routes through the backticked path
  Specs/Plans/WIP/Operation_Curl_Short2xxStagedUploadProof_2026-07-21.md, but the plan
  is complete under Specs/Plans/Done.
- README.md:176-178 titles the section Self-tests (Debug only) and calls all three
  suites debug-only.
- README.md:256-258 later states the accurate contract: Debug enables selftest plumbing
  by default, while a test-facing Release build can opt in with
  RSBuildEnableTests=true.

The two README statements cannot both be the execution contract. The later text and
build policy distinguish default configuration from capability.

#### Simplest safe repair

- Point the review routing entry to the completed Curl plan in Specs/Plans/Done and
  label it historical/complete.
- Rename the selftest section to Debug by default; opt-in Release and revise its opening
  sentence to match the build switch.
- Keep production Release exclusion explicit so the clarification does not imply
  shipping test hooks.
- Add a documentation-route checker that validates both actual Markdown links and
  backticked repository paths in Specs/Reviews and WIP/Done routing text.

#### Required tests and closeout

- Run the route checker across README/spec/plan Markdown, including a broken backticked
  path fixture.
- Assert the documented Debug command and an opt-in Release test build both match
  build.ps1/Directory.Build.props behavior.
- Assert a production Release build still omits selftest entry points.
- Update the authoritative Testing_SelfTests.md language at the same time.

### EG-S3-02 — Pin the exact S3 source revision through COPY and MOVE deletion

**Priority:** P1
**Window classification:** pre-existing defect in S3 transfer paths materially changed
and reviewed in the July 15-to-HEAD window
**Impact:** a concurrent replacement can produce a mixed or truncated destination and
MOVE can delete a newer source revision that was never faithfully copied
**Effort / repair risk / confidence:** L / High / High
**Planned at:** a00ec99a4, 2026-07-30

**Closed 2026-08-04.** `S3ObjectRevision` now carries ETag plus VersionId from
probe/list planning through small copy, each multipart part, relay GET, and conditional MOVE deletion. Versioned
transfers address and validate the captured VersionId; unversioned transfers use ETag conditions and fail with
`ERROR_REVISION_MISMATCH` without an unpinned fallback. Relay length/body counts are exact. The failure-first export
was 46/76; Debug and five test-enabled Release repeats are 131/0, including cancellation and delete-failure rollback.
The 16-object same-binary shape is `0/0 -> 16/16` source identity probes/conditional reads and median
`536,572 -> 786,180 us` (+46.52%), within the 175% gate. Durable behavior is in
`Specs/FileSystem/FileSystem_S3.md`; evidence is
`Specs/TestRuns/7d3a1247382a/FileOps/2026-08-04_1024_s3_source_revision_pinning/`; Floodgate consumes that proof without
duplicating ownership.

#### Evidence and failure interleaving

- `ResolvedS3Probe` and `PlannedTransferObject` at
  Plugins/FileSystemS3/FileSystemS3.Directory.cpp:156-171 carry kind, key, size, and
  timestamp, but no ETag or VersionId.
- Real recursive planning records only `object.GetKey()` and `object.GetSize()` at
  :990-1047.
- Small `CopyObject` at Plugins/FileSystemS3/FileSystemS3.S3.cpp:921-950 and
  multipart `UploadPartCopy` at :641-681 do not bind the source copy to an ETag or
  VersionId.
- Multipart ranges at :973-981 come from the earlier planned size. If the source
  changes between parts, one destination can contain bytes from multiple revisions.
- Relay fallback downloads the current key unconditionally at :462-510, while
  Directory.cpp:1213-1245 uploads only the earlier planned byte count.
- MOVE deletes by bucket/key at Directory.cpp:2127-2150 after destination publication.
  The delete request has no captured revision identity.
- The installed AWS SDK exposes source `If-Match`, GetObject `If-Match`/VersionId,
  and conditional DeleteObject fields. This is a missing use of the available commit
  condition, not a required new provider abstraction.

Concrete interleavings to reproduce:

1. Plan source revision A and its size.
2. Replace the key with revision B before a small copy, between multipart parts, or
   before relay GET.
3. Observe an unconditioned request reading B or mixing A/B while still reporting the
   planned size.
4. For MOVE, replace the key after destination publication and before deletion; the
   current code deletes B.

#### Simplest safe repair

Use one immutable source-revision record throughout the existing transfer plan:

- capture ETag and VersionId where available, together with key and size;
- bind small `CopyObject`, every `UploadPartCopy`, and relay `GetObject` to that exact
  source identity;
- require returned byte counts and source metadata to agree with the planned revision;
- for versioned buckets, copy and delete the captured VersionId;
- for unversioned keys, use the SDK's conditional delete against the captured ETag and
  map a failed condition to revision mismatch;
- never delete a MOVE source after any source-identity failure.

Do not add another S3 transfer framework. Extend `PlannedTransferObject` and thread the
same revision record through the existing small, multipart, relay, rollback, and delete
branches. Destination publication conditions from EG-S3-01 remain independently
mandatory.

#### Ordered implementation and verification

1. Extend the fake graph/transport with stable revision identity and deterministic
   barriers after planning, between multipart parts, before relay GET, and before MOVE
   deletion. First prove the current behavior fails.
2. Extend probe/list/plan results and request helpers to carry the revision. Reject an
   object whose exact identity cannot be established for a destructive MOVE.
3. Apply source conditions to every transfer request and conditional identity to final
   deletion. Normalize precondition failure to the existing revision/conflict result.
4. Add request-count and latency instrumentation for any extra HEAD/version lookup.
   Preserve bounded directory-operation performance evidence.

Focused gates:

    git diff --stat a00ec99a4 -- Plugins/FileSystemS3 Specs/FileSystem/FileSystem_S3.md
    .\build.ps1 -ProjectName FileSystemS3 -Configuration Debug
    .\build.ps1 -ProjectName PluginContractTests -Configuration Debug
    .\.build\x64\Debug\PluginContractTests.exe

Expected: the failure-first cases are RED before production changes; after the repair,
all source-replacement cases preserve the replacement, never mix revisions, and never
delete a MOVE source on mismatch. Existing destination-race and no-conflict cases stay
green.

#### Required tests, specification, and scope

- Cover small COPY/MOVE, multipart COPY/MOVE with replacement between parts, relay
  COPY/MOVE, same-size replacement, changed-size replacement, versioned and unversioned
  buckets, cancellation, rollback, and conditional-delete failure.
- Assert exact request identities and zero source DELETE on every failed condition.
- Archive same-machine request/latency evidence under a bounded
  `Specs/TestRuns/<commit>/FileOps/<run>/` directory.
- Update Specs/FileSystem/FileSystem_S3.md with source-revision pinning and conditional
  MOVE deletion. Reconcile the implementation proof in Operation Floodgate without
  reopening closed destination-publication rows.
- In scope: FileSystemS3 transfer plan/request/fake-test code, S3 specification,
  focused evidence, Testing_TestCoverage, this plan, and Floodgate routing.
- Out of scope: adaptive multipart sizing, a generic cloud-revision ABI, or weakening
  no-overwrite behavior to preserve a fast path.

#### Done / STOP

- Done when every successful destination is from one exact source revision and MOVE
  can delete only that same captured revision.
- STOP if the configured endpoint or installed SDK cannot apply the required source or
  delete condition. Fail closed and report the compatibility boundary; do not fall
  back to a check-then-unconditional-delete sequence.

### EG-MSD-02 — Validate every Microsoft Drive ranged response before exposing bytes

**Priority:** P2
**Window classification:** pre-existing defect in a Microsoft Drive transfer path
touched during the July 15-to-HEAD window
**Impact:** a shifted, empty, truncated, or inconsistent `206` can return wrong bytes or
premature EOF as a successful file read
**Effort / repair risk / confidence:** M / Medium / High
**Planned at:** a00ec99a4, 2026-07-30

#### Evidence

- `HttpResponse` at
  Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:100-108 has no
  `Content-Range` field.
- `SendHttpRequest` captures retry, request, location, and content-type headers at
  :2176-2195, so the transport already has the one central header-capture point needed.
- `MicrosoftDriveRangedFileReader::FetchRange` sends `Range` and optional `If-Match` at
  :4852-4861 but accepts any `200` or `206` at :4879-4889.
- Every `206` is assigned the requested offset at :4892-4907 without parsing its actual
  start/end/total or requiring the body length to match.
- `Read` treats an empty successful cache fill as EOF while `_position < _sizeBytes` at
  :4754-4787.

#### Simplest safe repair

- Capture `Content-Range` in `HttpResponse` at the existing transport boundary.
- Add one strict non-throwing parser for `bytes start-end/total`, with overflow-safe
  integer parsing and no locale dependence.
- For `206`, require start equals the requested offset, end and body length agree,
  total agrees with the known object size, and a non-final response is nonempty.
- For a legitimate `200` fallback, accept only a coherent full-object body at offset
  zero; otherwise return `ERROR_READ_FAULT`.
- Keep the current ETag `If-Match` behavior. Do not hide a malformed response by
  refreshing the download URL and retrying it as though it were authentication expiry.

#### Tests and verification

Extend the existing Microsoft Drive debug transport/provider contract tests with:

- valid exact `206`;
- shifted start, wrong end, inconsistent total, short body, long body, and empty
  non-EOF `206`;
- valid full-body `200` fallback and invalid partial `200`;
- range validation after a refreshed download URL;
- same-length shifted data, proving total-length checks alone are insufficient.

Focused gates:

    git diff --stat a00ec99a4 -- Plugins/FileSystemMicrosoftDrive Specs/FileSystem/FileSystem_MicrosoftDrive.md
    .\build.ps1 -ProjectName FileSystemMicrosoftDrive -Configuration Debug
    .\build.ps1 -ProjectName PluginContractTests -Configuration Debug
    .\.build\x64\Debug\PluginContractTests.exe

Update Specs/FileSystem/FileSystem_MicrosoftDrive.md with the exact ranged-response
contract and record the focused tests in Specs/Testing/Testing_TestCoverage.md.

#### Done / STOP

- Done when no byte is exposed until status, range metadata, body length, known total,
  and requested offset form one coherent response.
- STOP if a real supported endpoint intentionally omits `Content-Range` on `206`.
  Capture the endpoint evidence and define a documented compatibility rule before
  weakening validation.

### EG-ABI-01 — Centralize bounds validation for packed plugin FileInfo records

**Priority:** P2
**Window classification:** pre-existing defect in the public filesystem interface
reviewed as surrounding touched code
**Impact:** a buggy or untrusted plugin can cause an out-of-bounds host read while the
host converts a packed directory buffer into paths
**Effort / repair risk / confidence:** S / Low / High
**Planned at:** a00ec99a4, 2026-07-30

#### Evidence

- `BuildFileSystemPathListArenaFromFilesInformation` obtains an ABI-owned pointer,
  count, and used byte size at Common/PlugInterfaces/FileSystem.h:883-932.
- Its first traversal dereferences `entry->FileNameSize` at :963-975 before proving
  that `offsetof(FileInfo, FileName)` bytes remain.
- It checks only that `NextEntryOffset` does not exceed the remaining buffer at
  :995-1016. It does not require alignment, a minimum complete record length, or that
  the declared name fits inside that record.
- The second pass copies `entry->FileNameSize` bytes at :1088-1092, including for a
  terminal record whose `NextEntryOffset` is zero.
- `PackedFileInfoBuffer::LocateEntry` in Common/PackedFileInfoBuffer.h:182-220 and the
  File Operations helpers at
  RedSalamander/FolderWindow.FileOperations.State.cpp:2271-2298 and :2383-2421 already
  contain overlapping checked traversal logic. The fix must converge on one reviewed
  contract rather than add a fourth variant.

#### Simplest safe repair

Define one canonical inline packed-record validation/traversal helper in the lowest
dependency layer that can be used by the public path-list helper and host consumers.
For every record, before any field or payload read:

- require the fixed header through `offsetof(FileInfo, FileName)` in bounds;
- require even `FileNameSize` and the complete declared name in bounds;
- for nonterminal records, require aligned `NextEntryOffset` greater than or equal to
  the complete record payload and wholly inside the buffer;
- for terminal records, require the complete name payload inside the remaining buffer;
- require the terminal position and observed record count to agree with `GetCount()`;
- reject overlap, early termination, extra records, and arithmetic overflow.

Preserve the `FileInfo` ABI layout. Do not add virtual methods or reinterpret malformed
data to recover a partial path list. If File Operations needs different skip-and-warn
semantics for search/enumeration, keep that policy at its caller while reusing the same
structural validator.

#### Tests and verification

Extend Tests/PluginContractTests/PluginContractTests.cpp, using
`TestPackedFileInfoBuffer` as the test pattern, with:

- truncated first and later headers;
- odd and oversized names;
- offset smaller than the declared record, overlap, misalignment, overflow, and offset
  past the buffer;
- terminal name extending beyond the buffer;
- early terminal and count mismatch;
- a valid one-record terminal buffer and valid aligned multi-record buffer.

Focused gates:

    git diff --stat a00ec99a4 -- Common/PlugInterfaces/FileSystem.h Common/PackedFileInfoBuffer.h RedSalamander/FolderWindow.FileOperations.State.cpp Tests/PluginContractTests
    .\build.ps1 -ProjectName PluginContractTests -Configuration Debug
    .\.build\x64\Debug\PluginContractTests.exe

Update Specs/Plugins/Plugins_VirtualFileSystem.md with consumer validation obligations
and Specs/Core/Core_SharedHelpers.md if a shared validator is added or an existing
helper is promoted.

#### Done / STOP

- Done when every packed-record consumer in scope validates before dereference and the
  malformed corpus returns a stable failure without an out-of-bounds read.
- STOP if placing the helper in `FileSystem.h` would create a dependency or ABI problem.
  Select another Common header and update the shared-helper catalog; do not duplicate
  the validator locally.

### EG-CURL-02 — Keep Curl delete concurrency scoped to each connection

**Priority:** P3
**Window classification:** pre-existing performance defect in a Curl path touched by
the July 15-to-HEAD work
**Impact:** one endpoint capped at concurrency 1 serializes unrelated endpoints that
could safely delete concurrently
**Effort / repair risk / confidence:** M / Medium / High
**Planned at:** a00ec99a4, 2026-07-30

#### Evidence

- The delete task builder starts with the configured job cap and folds every resolved
  connection's `effectiveDeleteMaxConcurrency` into one global minimum at
  Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp:3312-3349.
- That minimum becomes the entire job worker count at :3386-3393.
- Actual delete/remove-directory calls already acquire a permit from the shared
  `ConnectionConcurrencyLimiter` by `limiterKey` and the connection's own limit at
  :571-620.
- The outer global minimum therefore repeats the limiter at the wrong scope and makes
  independent connections interfere.

#### Simplest safe repair

Size the top-level task scheduler from the configured global job cap and task count.
Leave each endpoint's safety limit to the existing keyed limiter. Preserve the existing
rule that prevents nested top-level/directory oversubscription; if recursive directory
tasks need their own budget, use the same bounded shared worker budget rather than
creating per-directory thread pools.

Do not change default connection limits or introduce a second limiter.

#### Tests, performance evidence, and verification

- Add a deterministic test with at least two limiter keys: connection A capped at 1,
  connection B capped above 1.
- Require A never exceeds 1, B never exceeds its own cap, and B overlaps A instead of
  waiting for the whole A batch.
- Cover mixed files/directories, cancellation while waiting for a permit, first
  failure, continue-on-error, and recursive delete without oversubscription.
- Instrument per-key active/high-water counts and total elapsed time.
- Capture a same-machine baseline/candidate under identical fixed-latency fake
  transport and archive bounded evidence under Specs/TestRuns.

Focused gates:

    git diff --stat a00ec99a4 -- Plugins/FileSystemCurl Specs/FileSystem
    .\build.ps1 -ProjectName FileSystemCurl -Configuration Debug
    .\build.ps1 -ProjectName PluginContractTests -Configuration Debug
    .\.build\x64\Debug\PluginContractTests.exe

Update the Curl authoritative filesystem specification with global-versus-per-key
concurrency semantics.

#### Done / STOP

- Done when independent limiter keys overlap, every key stays within its own cap, and
  nested work remains within the process budget with archived evidence.
- STOP if the current task model can create unbounded nested parallelism after removing
  the global minimum. Define and test one shared worker budget before enabling overlap.

### EG-GOV-01 — Keep runtime fixes, plan changes, and workspace policy independently reviewable

**Priority:** P3
**Window classification:** introduced in the expanded review window
**Impact:** unrelated changes share review, rollback, and blame boundaries; dirty
follow-up plan edits make the intended committed state harder to identify
**Effort / repair risk / confidence:** S / Low / High
**Planned at:** a00ec99a4, 2026-07-30

#### Evidence

- Commit a00ec99a4 is titled `Fix cached Keep Both conflict eligibility` but also
  includes seven Terminal WIP plans and `.codex/config.toml`.
- The current worktree modifies the same seven `Specs/Plans/WIP/Terminal_*.md` files
  again, with no uncommitted source code.
- The runtime Keep Both repair itself is closed and verified. Rewriting published
  history would create more risk than it removes; the required remediation is future
  commit discipline and an explicit plan-only handoff for the current Terminal edits.

#### Required action

- Preserve the current Terminal edits exactly; do not absorb, discard, or rewrite them
  while executing Emberglass runtime work.
- Have the Terminal owner review the seven-plan diff against the master Terminal plan
  and its six execution slices.
- When authorized to commit, use a dedicated plan-only commit. Keep workspace policy
  changes in a separately named commit unless they are a declared prerequisite of that
  plan update.
- For every Emberglass implementation package, keep production code, focused tests,
  authoritative spec/evidence, and the package's closeout together, but exclude
  unrelated domain plans and configuration.
- Before staging, use `git diff --name-only` and the package's explicit in-scope list.

Verification:

    git status --short
    git diff --check
    git diff --name-only --cached

Expected: no unrelated Terminal or workspace-policy file appears in a staged runtime
repair; a later Terminal plan commit contains only the reviewed Terminal WIP/index files
and any explicitly documented plan-routing update.

#### Done / STOP

- Done when the current Terminal changes have one reviewed plan-only disposition and
  every later remediation commit is scoped to one executable package.
- STOP before staging or rewriting any pre-existing Terminal change without the
  Terminal owner's scope confirmation. Do not rewrite a00ec99a4 solely for commit
  aesthetics.

### R4-SET-01 — Make settings backup participate in the same CAS transaction

**Priority:** P1
**Impact:** a recovery reader can displace a newer valid commit to the bad-file sibling
and later allow defaults at the canonical path
**Effort / repair risk / confidence:** M / High / High
**Status:** Reopened on 2026-07-22; remediation owned by EG-SET-02

#### Evidence

- Common/Common/SettingsStore.cpp:366-379 implements BackupBadSettingsFile with MoveFileExW against the path, with no commit lock and no expected source stamp.
- Common/Common/SettingsStore.cpp:507-643 contains the lock/stamp compare-and-swap path used by normal publication.
- Common/Common/SettingsStore.cpp:5443-5495 reads an invalid source, later backs up the path, and resets the expected stamp.
- Common/Common/SettingsStore.cpp:5792-5808 exposes the same unguarded behavior for explicit replacement.
- RedSalamander/SettingsHotReload.cpp:1236-1255 backs up, resets the stamp, and saves replacement settings.
- Existing coverage at Tests/SettingsSchemaTests/SettingsSchemaTests.cpp:626-704 and :801-856 proves ordinary recovery/CAS behavior, but not the interleaving between invalid-read and backup.

Interleaving:

1. Process A reads invalid revision X.
2. Process B commits valid revision Y.
3. Process A calls BackupBadSettingsFile against the path and moves Y to the bad-file name.
4. Process A resets its expected stamp; later default persistence can populate the
   canonical path.

The result violates the conflict-aware publication contract in Specs/Core/Core_SettingsStore.md:109-150 and :223-247.

#### Simplest safe repair

Add one guarded backup primitive beside the existing commit transaction:

- input: settings path and the optional exact stamp of the source that was diagnosed;
- acquire SettingsCommitLock;
- restat the source under the lock;
- compare it with the expected stamp;
- move only that exact source to a unique bad-file sibling;
- return ERROR_REVISION_MISMATCH if the path changed;
- never reset the caller stamp until the guarded operation succeeds.

Automatic recovery should reload once on revision mismatch. Explicit replacement should operate on the stamp of the snapshot presented to the user. It must not turn a stale confirmation into permission to replace a newer file.

Do not create a second lock or a backup-only retry loop. The source identity check and move belong inside the existing publication exclusion boundary.

#### Required tests

Add deterministic barriers, not timing sleeps:

- A reads invalid X and pauses immediately before guarded backup.
- B publishes valid Y.
- A resumes and must receive revision mismatch.
- Y remains the live file; no backup contains Y; defaults are not published.
- Repeat for SettingsHotReload explicit replacement.
- Cover source deletion, already-backed-up source, and backup-name collision.

Focused gates:

    .\build.ps1 -ProjectName SettingsSchemaTests -Configuration Debug
    .\.build\x64\Debug\SettingsSchemaTests.exe
    .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter settings_store_unsupported_future_schema_never_overwrites_source -TimeoutMultiplier 2.0

#### Specification and ownership

Update Specs/Core/Core_SettingsStore.md so all destructive recovery of the live path is explicitly lock-and-stamp guarded.

The Current Audit Refresh and Best Next Pick sections of Specs/Plans/WIP/README.md
preserve the Track 7 completion history while routing this exact guarded-backup
exception here. Do not reopen unrelated Track 7 work.

#### Historical closeout evidence — cooperating writers only

- Among writers participating in `SettingsCommitLock`, `BackupSettingsFileGuarded(...)` compares the exact
  optional source stamp under that lock and moves the revision that still occupies the compared path. Changed,
  created, and deleted targets
  return `ERROR_REVISION_MISMATCH` without moving either revision.
- Automatic recovery re-resolves and reloads once after a stale backup comparison. A replacement revision is
  loaded; deletion becomes a missing-target default snapshot. Explicit replacement passes the stamp presented
  to the user through `SettingsHotReload` and rejects stale confirmation.
- `SettingsSchemaTests` includes a failure-first deterministic barrier for invalid X / valid Y, source deletion,
  a forced backup-name collision, and a child-process X/Y guarded-backup race. The focused harness passes.
- The Commands case `settings_store_unsupported_future_schema_never_overwrites_source` proves changed and deleted
  explicit-replacement sources are rejected and the ordinary confirmed replacement still succeeds.
- A clean four-worker full Debug rebuild passes with 0 warnings and 0 errors. Two broader `Full -SkipBuild`
  attempts were inconclusive: the unrelated Commands run exceeded one-hour and two-hour outer caps without
  producing a failed-test result; its last scratch activity was `batch_rename_async_provider_collection`.
  After the required clean rebuild, both settings-focused gates passed again and no contamination marker or
  lingering process remained.
- The authoritative transaction and recovery contract is updated in `Specs/Core/Core_SettingsStore.md`.

Emberglass contrary evidence: the closeout statements above remain valid for
cooperating RedSalamander writers, but the implementation closes the comparison handle
before its pathname MoveFileExW. EG-SET-02 demonstrates the remaining non-cooperating
external-writer interleaving. This package must not return to Complete until the
retained-handle rename proof passes.

#### Done / STOP

- Done when stale recovery can neither move nor overwrite a newer commit in in-process and cross-process proof.
- STOP if the existing file stamp cannot uniquely identify the diagnosed source on supported filesystems; define a stronger revision token before proceeding.

### R4-S3-01 — Enforce no-overwrite at S3 publication

**Priority:** P1
**Impact:** an object created after preflight can be overwritten without user consent
**Effort / repair risk / confidence:** M / Medium / High
**Status:** Complete 2026-07-22

#### Evidence

- Plugins/FileSystemS3/FileSystemS3.IO.cpp:849-873 and :2819-2825 perform an existence preflight.
- The overwrite permission is stored at :1373-1379 and passed at :2011, but it is not consumed at the final publication operation.
- Small-object publication at :1535-1544 and multipart completion at :1574-1581 are unconditional.
- Plugins/FileSystemS3/FileSystemS3.S3.cpp:679-720 and :807-843 build those requests without a conditional header.
- RedSalamander/FolderWindow.FileOperations.State.cpp:9185-9222 writes the final key directly.
- FileSystemS3.IO.cpp:2835-2850 advertises an atomic file writer, making final publication the integrity boundary.

AWS documents If-None-Match: * for conditional PutObject and CompleteMultipartUpload. An existing object produces 412 and a conflicting concurrent operation can produce 409:

- [Amazon S3 conditional writes](https://docs.aws.amazon.com/AmazonS3/latest/userguide/conditional-writes.html)
- [PutObject API](https://docs.aws.amazon.com/AmazonS3/latest/API/API_PutObject.html)
- [CompleteMultipartUpload API](https://docs.aws.amazon.com/AmazonS3/latest/API/API_CompleteMultipartUpload.html)

#### Simplest safe repair

When overwrite was not granted:

- add If-None-Match: * to PutObject;
- add If-None-Match: * to CompleteMultipartUpload;
- map 412 and the relevant 409 conflict to the host conflict/replan result;
- preserve unconditional behavior only when overwrite was explicitly granted.

An endpoint that claims S3 compatibility but cannot honor conditional final publication must fail closed for no-overwrite or stop advertising this atomic final-writer capability. Do not add another HEAD immediately before PUT; that only moves the race.

#### Required tests

Extend RunDebugMultipartWriterContractSelfTest in FileSystemS3.IO.cpp:2035 onward and the PluginContractTests path at Tests/PluginContractTests/PluginContractTests.cpp:1092-1122:

- preflight sees no object;
- inject an object immediately before small PutObject publication;
- inject an object immediately before CompleteMultipartUpload;
- both operations preserve the injected object and return conflict;
- overwrite-authorized operations still replace it;
- exercise 409 and 412 mappings.

Focused gates:

    .\build.ps1 -ProjectName FileSystemS3 -Configuration Debug
    .\build.ps1 -ProjectName PluginContractTests -Configuration Debug
    .\.build\x64\Debug\PluginContractTests.exe

#### Specification and ownership

Update Specs/FileSystem/FileSystem_S3.md:145-147 and the atomic-final-writer publication rules in Specs/Core/Core_FileSystemBridge.md:35-50 and :88.

Historical implementation routing used Specs/Plans/Done/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md. Its completed S3 claims do not cover this final-publication race; current accepted work is routed through Cinderstar rather than a parallel cloud plan.

#### Closeout evidence

- `MultipartS3FileWriter` now passes its exact overwrite grant to both publication forms. Without a grant,
  `PutObject` and `CompleteMultipartUpload` carry `If-None-Match: *`; explicit overwrite remains unconditional.
- Conditional `412 Precondition Failed` and `409 Conflict` responses map to
  `HRESULT_FROM_WIN32(ERROR_FILE_EXISTS)`. Other failures keep the existing AWS error mapping, and endpoints that
  reject conditional publication fail closed without another `HEAD` race.
- The deterministic transport injects an existing destination only at final publication. The RED run was
  34 passed / 4 failed before the writer forwarded its no-overwrite requirement. The green Debug and test-enabled
  Release runs are 38 / 0: small and multipart publication preserve the injected object and return conflict, while
  overwrite-authorized controls replace it. Multipart conflict aborts the uncommitted upload exactly once.
- Debug `FileSystemS3` and `PluginContractTests`, test-enabled Release `FileSystemS3` and
  `PluginContractTests`, and tests-disabled production Release `FileSystemS3` build with 0 warnings / 0 errors.
  Complete `PluginContractTests` passes in Debug and test-enabled Release, including eight S3 unload-refresh cycles.
  The focused source contracts pass 149 / 0 and pin both SDK headers plus the runtime cases.
- Durable publication and bridge contracts are in `Specs/FileSystem/FileSystem_S3.md` and
  `Specs/Core/Core_FileSystemBridge.md`; dated test evidence is recorded in
  `Specs/Testing/Testing_TestCoverage.md`.

#### Done / STOP

- Done when no-overwrite is atomic for both publication forms and incompatible endpoints fail closed.
- STOP if the host has no distinct result for a late conflict; define that result before mapping the provider response.

### R4-S3-02 — Give the multipart abort queue one self-converging owner

**Priority:** P1
**Impact:** orphaned multipart uploads can remain indefinitely and prevent clean unload
**Effort / repair risk / confidence:** M / Medium / High

#### Evidence

- PendingMultipartAbortQueue::Schedule at Plugins/FileSystemS3/FileSystemS3.IO.cpp:1072-1113 is the only scheduler.
- A pass processes at most 16 entries at :1167-1173.
- Failures are requeued at :1189-1200.
- The worker clears scheduled and records nextAttempt at :1203-1208 but does not schedule its own continuation.
- WaitUntilDrained at :1122-1137 can wait until its caller deadline rather than the next retry.
- Factory shutdown/can-unload polling at Plugins/FileSystemS3/Factory.cpp:149-158 happens to call back into the queue, but ordinary background progress must not depend on unload polling.
- AbortMultipartUpload maps through FileSystemS3.S3.cpp:846-870 and FileSystemS3.Internal.h:210-218. A 404/NoSuchUpload becomes file-not-found and is requeued forever even though the desired terminal state is already true.
- Existing coverage at FileSystemS3.IO.cpp:2422-2448 fails the destructor abort once and lets the first queued retry succeed. It does not fail a queue-owned retry or exceed the 16-entry pass.

#### Simplest safe repair

Use one queue-owned continuation:

- keep the bounded batch;
- after a pass, schedule another callback when entries remain;
- delay until the earliest nextAttempt for backoff;
- make only one callback scheduled at a time;
- wake WaitUntilDrained at the earlier of its deadline or the next retry time;
- treat 404/NoSuchUpload as idempotent abort success;
- keep shutdown cancellation and module-pin ownership explicit.

Do not add a polling thread. The existing threadpool callback plus one delayed-continuation state is sufficient.

#### Required tests

- More than 16 successful aborts drain without an external Schedule call.
- A queue-owned abort fails once, then succeeds on its own retry.
- A terminal 404 drains immediately.
- Shutdown during delayed retry cancels or completes according to the documented deadline.
- CanUnloadNow converges without repeated external polling.

Focused gates:

    .\build.ps1 -ProjectName FileSystemS3 -Configuration Debug
    .\build.ps1 -ProjectName PluginContractTests -Configuration Debug
    .\.build\x64\Debug\PluginContractTests.exe

#### Specification and ownership

Update Specs/FileSystem/FileSystem_S3.md:137 and :159-161 with the exact convergence, terminal-404, retry, shutdown, and unload rules. Route through Operation Floodgate and supersede only its OBS-S3-P3-02/Track 11 assertion that this lifecycle is complete.

#### Done / STOP

- Done when the queue drains without caller polling and every terminal state is covered.
- STOP if a delayed callback cannot hold the required module pin safely; resolve callback/module lifetime before rescheduling.

### R4-MON-01 — Bound Monitor file open by retained representation

**Priority:** P1
**Impact:** accepted files can retain and transiently duplicate far more memory than the configured limit; cancellation and close can stall during CPU conversion
**Effort / repair risk / confidence:** L / High / High

#### Evidence

- RedSalamanderMonitor/RedSalamanderMonitor.cpp:2547-2552 passes the UTF-16 retention maximum as an encoded file-byte limit.
- RedSalamanderMonitor/MonitorFileReader.cpp:69-104 reads the entire encoded file into a vector.
- MonitorFileReader.cpp:106-146 decodes a second whole representation. UTF-16 validation creates a whole temporary UTF-8 representation at :119.
- Stop is not polled during conversion; it is next checked during line scanning at :148-163.
- An ASCII file at the accepted 64 MiB encoded limit can require roughly 128 MiB for retained UTF-16 before vector capacity, line metadata, search state, and temporary copies.
- UI publication at RedSalamanderMonitor.cpp:2634-2636 calls ColorTextView::SetText.
- RedSalamanderMonitor/ColorTextView.cpp:474-516 reparses/copies and rebuilds search synchronously on the UI thread.
- Document::SetText does not independently enforce the retention budget.
- CancelSynchronousIo is used at RedSalamanderMonitor.cpp:2498-2513, so blocking reads can be interrupted, but CPU conversion cannot.
- Teardown joins synchronously at :4161-4164.
- Tests/MonitorTest/MonitorTest.cpp:190-266 covers pre-cancel, not mid-decode cancellation, expansion, publication cost, or close latency.

#### Simplest safe repair

Implement one streaming strict decoder:

- retain encoding boundary state between chunks;
- count decoded UTF-16 bytes and lines, not only input bytes;
- reject before exceeding the retained representation budget;
- poll stop between bounded chunks;
- validate UTF-8 sequences and UTF-16 surrogate pairs across chunk boundaries;
- build the worker-owned line snapshot once;
- move/swap the completed snapshot into Document on the UI thread instead of passing a monolithic string through SetText;
- enforce the UINT32 coordinate ceiling explicitly before publication.

The worker should publish one immutable result or one typed failure. Avoid retaining encoded and decoded whole-file copies simultaneously. Avoid adding a second parser after the streaming decoder.

#### Required tests and performance proof

- ASCII expansion at and just over the decoded budget.
- UTF-8 multibyte sequence split across chunks.
- UTF-16 surrogate pair split across chunks.
- Strict rejection of invalid/truncated encodings.
- Cancellation in the middle of conversion.
- Close during conversion with a bounded join latency.
- No partially published document on failure/cancel.
- UI publication latency and peak retained memory at the supported maximum.

Focused gate:

    .\build.ps1 -ProjectName MonitorTest -Configuration Debug
    .\.build\x64\Debug\MonitorTest.exe --document-model-selftest

Archive before/after metrics under Specs/TestRuns within archive limits.

#### Specification and ownership

Update Specs/Core/Core_RedSalamanderMonitor.md:140 and :146-175 with encoded-input, decoded-retention, line, coordinate, cancellation, publication, and close-latency limits.

Specs/Plans/WIP/README.md:55-58 preserves the Track 6 completion history, while :86-89 now routes this exact fresh file-open/search evidence here. Reopen only this file-open contract.

#### Done / STOP

- Done when limits apply to retained decoded state, cancellation is observed during CPU work, and publication does not redo the parse.
- STOP if exact peak-memory evidence cannot be collected deterministically on the test machine; define an allocation/retained-byte instrumentation counter before implementation closeout.

### R4-MON-02 — Preserve an explicit search scan frontier across eviction

**Priority:** P1
**Impact:** search/F3 can permanently omit matches that remain in the retained document
**Effort / repair risk / confidence:** M / Medium / High

#### Evidence

- RedSalamanderMonitor/ColorTextView.cpp:733-744 removes and shifts existing match positions after document eviction.
- The full scan stops when the cap is reached at :4773-4811.
- Append processing at :5465-5475 scans only newly appended lines.

Once the cap has been reached, eviction can remove indexed matches and create free capacity. Retained middle lines that were never scanned remain behind the implicit frontier, while only new tail lines are scanned. Those middle matches are permanently invisible until a full rebuild.

Specs/Core/Core_RedSalamanderMonitor.md:152-157 promises bounded search but does not define this frontier behavior.

Tests/MonitorTest/MonitorTest.vcxproj:126-129 does not compile ColorTextView, which explains why this state machine is not directly covered.

#### Simplest safe repair

Extract only the pure match-index/cursor state needed for deterministic testing. Keep an explicit frontier:

- source line identity or retained line index;
- text offset within that line;
- whether the retained range is fully scanned.

On rebuild, reset the frontier. On eviction, translate it by the same retained-range operation. When indexed matches are evicted and capacity is freed, resume the oldest retained unindexed region before scanning new tail content. Keep match positions sorted.

Do not rebuild the whole retained document on every append or eviction; that would repair correctness by violating the hot-path performance contract.

#### Required tests and performance proof

- More matches than the cap.
- Evict only part of the indexed prefix.
- Verify newly free slots are filled from the retained, previously unindexed middle.
- Verify F3/Shift+F3 ordering and selection.
- Repeat append/evict cycles without duplicates or gaps.
- Measure append processing at the configured cap before and after.

Focused gate:

    .\build.ps1 -ProjectName MonitorTest -Configuration Debug
    .\.build\x64\Debug\MonitorTest.exe --document-model-selftest

#### Specification and ownership

Update Core_RedSalamanderMonitor.md with the frontier invariant and which retained region is preferred when capacity returns. This behavior was introduced in the recent Monitor work, including commit ce4e6c4ca; it is fresh contrary evidence for the closed Track 6 search claim.

#### Done / STOP

- Done when every retained match is eventually discoverable in document order subject to the cap.
- STOP if source line indices are not stable enough across eviction; introduce a monotonic line identity in Document rather than layering more offset arithmetic in ColorTextView.

### R4-RC-01 — Paste rectangular cells by visible row and culture identities

**Priority:** P1
**Impact:** paste can write to cells other than those visibly targeted
**Effort / repair risk / confidence:** M / Medium / High

#### Evidence

- RedConfigure/RedConfigureRoot.cpp:1967-1984 and :1990-1994 reorder/pin visible culture columns.
- Paste resolves only the selected raw culture index at RedConfigureRoot.cpp:2065-2081.
- RedConfigure/RedConfigureSession.cpp:1300-1325 increments raw target indices for rectangular paste.
- The selected row is resolved from the current view at RedConfigureRoot.cpp:356-365 and :379-385.
- The sorted/filtered view order lives in _localizationReviewViewRows at :983-1025.
- The session then increments raw rows, so a sorted or filtered non-contiguous visible selection has the same defect as reordered cultures.
- Existing coverage at Tests/RedConfigureTests/RedConfigureTests.cpp:1592-1637 proves a single-cell/raw-order case only.

The visible-order contract is stated in Specs/Core/Core_RedConfigure.md:115 and Specs/UI/UI_RedConfigure.md:47.

#### Simplest safe repair

At the UI boundary, build two ordered identity lists:

- destination row identities from the current visible review view beginning at the selected row;
- culture names/IDs from the current visible column order beginning at the selected culture.

Change the session operation to accept those identities. Validate the entire rectangular mapping before mutation, then take one undo snapshot and apply it. A missing selected identity or a shape that cannot be mapped must fail closed.

Do not teach the session about sort/filter/pinning state. Do not pass one raw starting index and assume successor order.

The product must also state whether one invalid placeholder rejects the whole rectangular paste or only the invalid cell. Prefer all-or-nothing because the operation already has one undo unit and prevalidation is simple.

#### Required tests

Use a 2x2 paste with:

- stored cultures fr, de, ja;
- visible cultures ja, fr, de;
- sorted/filtered rows that are non-contiguous in storage;
- one hidden culture;
- undo/redo;
- a missing identity and an invalid placeholder.

Verify the visible 2x2 target is the only mutated set and one undo restores all cells.

Focused gates:

    .\build.ps1 -ProjectName RedConfigureTests -Configuration Debug
    .\.build\x64\Debug\RedConfigureTests.exe

#### Specification and ownership

Update Core_RedConfigure.md and UI_RedConfigure.md with identity-based mapping and atomic validation. Coordinate with Specs/Plans/WIP/Operation_RosettaLantern_RedConfigureSatelliteLocalizationReview_2026-07-17.md for culture identity/spec changes, but do not make translation inventory work own this defect.

Specs/Plans/WIP/README.md:82-83 preserves the Track 19 completion history, while :86-89 routes this exact visible-coordinate paste case here instead of reopening all Track 19 work.

#### Done / STOP

- Done when visible rows/columns, not raw adjacency, define the destination rectangle.
- STOP until the all-or-nothing invalid-placeholder policy is explicit.

### R4-SEARCH-01 — Remove the predictable permissive Global startup guard

**Priority:** P1
**Impact:** a low-privilege interactive process can prevent the privileged SearchService from starting
**Effort / repair risk / confidence:** M / High / High
**Status:** Complete 2026-07-22

#### Evidence

- RedSalamanderSearchService/Main.cpp:277-288 derives a predictable Global event name.
- Main.cpp:291-309 grants interactive users generic-all access.
- Main.cpp:311-333 treats an existing event as another live instance.
- ServiceMain exits on that result at :3270-3283.

Microsoft documents that CreateEvent/CreateEventEx opens an existing named event and reports ERROR_ALREADY_EXISTS. Global names are shared across sessions, and event creation itself is not protected by the SeCreateGlobalPrivilege restriction that applies to certain other global object types:

- [CreateEvent documentation](https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-createeventa)
- [CreateEventEx documentation](https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-createeventexa)
- [Kernel object namespaces](https://learn.microsoft.com/en-us/windows/win32/termserv/kernel-object-namespaces)

The existing search_service_foreground_rejects_second_instance case, registered from the Compare suite, proves only legitimate same-user duplication. It does not prove resistance to precreation or a held permissive handle across service restart.

#### Simplest safe repair

Separate the modes:

- Service mode should rely on SCM service-instance serialization and the service’s real store/pipe ownership checks.
- Foreground/offline compact mode should take a resource-scoped exclusive lock in the actual store directory, where directory ACLs are already the trust boundary.
- A pre-existing closed lock file must be harmless; liveness comes from an open exclusive handle, not filename existence.

A service-SID private namespace is an acceptable alternative only if it is already required for another secured kernel object. Do not add a new private-namespace framework solely for this guard.

Coordinate default pipe/store behavior so service, foreground, and compact modes cannot mutate the same store concurrently.

#### Required tests

- Low-privilege/precreated named event does not block service startup after the event guard is removed.
- A legitimate live service prevents another writer for the same store.
- A crashed process releases the store lock and restart succeeds.
- Foreground instances with different stores do not block one another.
- ACL/owner failure fails closed with one actionable diagnostic.

Focused gate:

    .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_foreground_rejects_second_instance -TimeoutMultiplier 2.0

Add a deterministic security test or an ACL/owner proof that does not require an ambient administrator account.

#### Specification and ownership

Update the SearchService trust-boundary and mode-serialization sections in the authoritative Search specification. Coordinate with Specs/Plans/WIP/Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md because it owns adjacent SearchService security policy; keep this as a distinct startup-guard row.

#### Closeout evidence

- The predictable permissive Global event, its environment override, SDDL, and `ERROR_ALREADY_EXISTS` startup
  decision are removed. SCM still serializes the installed service identity, while service, foreground, and offline
  `--compact` retain one non-inheritable share-none WIL file lease in the resolved store directory for their full
  writer lifetime.
- Default service, foreground, and offline-compaction resolution now converges on the build-specific ProgramData
  root. A stale closed `.red-salamander-search-writer.lock` file is reusable; only an authorized process retaining
  the exclusive handle can block the same store. Different store directories remain independent and an abrupt
  JobObject-owned process exit releases the lease through handle teardown.
- Failure-first Release proof at
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-07-22_115045/` failed because a precreated legacy Global event
  blocked isolated store A. Final deterministic proof is green in Debug (**2 / 0**) and test-enabled Release
  (`search_service_foreground_rejects_second_instance` **5 / 0** and
  `search_service_store_writer_ownership_is_resource_scoped` **3 / 0**). The comprehensive case additionally proves
  offline-compaction contention, stale-file tolerance, independent stores, crash/restart, and path-specific
  fail-closed diagnostics without requiring elevation.
- `search.service.writer_ownership.acquire_us` records the startup-only ownership cost. On machine
  `7d3a1247382a`, five Release samples moved from Global-event owner/conflict medians **2,123 / 1,796 us** to
  store-file medians **4,533 / 5,015 us**, below the documented 15 ms advisory median budget. Baseline and final
  candidate archives are `2026-07-22_115034/` and `2026-07-22_121358/` in the same CompareDirectories profile;
  no steady-state throughput claim is made.
- Debug and final test-enabled Release builds complete with 0 warnings / 0 errors. The authoritative ownership,
  default-root, failure, metric, and regression contracts are in `Specs/Core/Core_Search.md`; dated validation is
  recorded in `Specs/Testing/Testing_TestCoverage.md`. Parallax records that its distinct named-pipe impersonation
  review remains deferred.
- The writer lease exposed an invalid pre-existing fixture that placed the persistent store beneath an indexed root
  and then deleted that root. The fixture now uses a sibling store directory: the isolated regression passes
  **1 / 0** and the complete Debug Compare gate passes **227 / 0 / 31 environment skips**, archived at
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-07-22_142239/` and
  `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-07-22_142625/`.
- The final mandatory test-enabled Debug full-solution rebuild passes with 0 warnings / 0 errors in
  `.build/logs/msbuild-20260722_141600_417.log`; focused source contracts pass **149 / 0 / 0**. A Full attempt was
  stopped after an unrelated order-sensitive Commands settings fixture raised an `ERROR_REVISION_MISMATCH` modal;
  its isolated case and immediate predecessor cluster both passed after recovery, so it remains separate harness
  follow-up rather than contrary SearchService evidence.

#### Done / STOP

- Done when only an authorized live owner of the same store can block startup.
- STOP if service and foreground modes currently share a store without a common exclusive resource; define that common resource first.

### R4-MON-03 — Rebuild offset-derived state when Show IDs changes display text

**Priority:** P2
**Impact:** search navigation, caret, selection, or copy can point at stale display offsets
**Effort / repair risk / confidence:** S / Medium / High

#### Evidence

- RedSalamanderMonitor/ColorTextView.cpp:298-337 toggles Document::EnableShowIds and only clamps selection/caret.
- RedSalamanderMonitor/Document.cpp:1112-1117 and :1157-1174 add/remove per-line display prefixes. Prefix length differs between metadata and ordinary lines.
- Search positions are stored as display offsets in ColorTextView.cpp:4782-4795.
- The current toggle test in RedSalamanderMonitor/RedSalamanderMonitor.cpp:3018-3025 verifies state, not search/selection coordinates.

Clamping prevents an out-of-range coordinate but does not make a formerly valid coordinate refer to the same logical payload or to the same match.

#### Simplest safe repair

Choose and document the simpler interaction contract:

- toggle Show IDs;
- rebuild all display-offset-derived search matches;
- reset active match navigation, selection, and caret to a documented safe position.

If product requirements demand preservation, add one conversion between logical source-line/payload offsets and display offsets, then use it for caret, selection endpoints, and active match. Do not preserve some state with prefix arithmetic while merely clamping the rest.

The reset-and-rebuild contract is preferred unless there is a demonstrated user requirement for preservation.

#### Required tests

- Lines with and without metadata prefixes.
- Search term before and after the prefix boundary.
- Toggle off/on with an active match, a selection, and a caret.
- F3/Shift+F3 and Copy after each toggle.
- Empty document and document eviction.

Add an ENABLE_TESTS-gated Show IDs phase to the existing Monitor application chrome selftest. That phase must drive the real ColorTextView, not a duplicate model.

Focused gates:

    .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug
    .\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest

Keep the existing chrome selftest in the unified CI/Full plan. A pure mapping helper may also be tested in MonitorTest, but it cannot replace the application-level proof because MonitorTest does not compile ColorTextView.cpp.

#### Specification and ownership

Update Specs/Core/Core_RedSalamanderMonitor.md with the chosen state-reset/preservation contract. This is pre-existing behavior and is not a reason to reopen the entire Track 6 effort.

#### Done / STOP

- Done when all display-offset-derived state follows one documented toggle rule.
- STOP if product chooses state preservation but cannot define logical selection across metadata-only prefixes.

### R4-FILEACTION-01 — Give selected-path inventory files a move-only lifecycle owner

**Priority:** P2
**Impact:** temporary files containing selected user paths can remain after launch faults, timeout, or cleanup setup failure
**Effort / repair risk / confidence:** M / Medium / High

#### Evidence

- RedSalamander/FileActionLauncher.cpp:353-435 creates and writes selected-path inventory files.
- The deferred cleanup helper at :465-488 returns success for a null process and can fail allocation or wait registration.
- The async caller at :689-695 logs cleanup scheduling failure but returns success without another owner.
- Wait mode exits without deleting or transferring cleanup ownership on a missing process handle at :697-701, timeout at :704-709, wait failure at :711-720, and GetExitCodeProcess failure at :723-728.
- Cleanup occurs only on the successful wait path at :737.
- The launch plan can outlive or fail before launch while still containing a created temporary file.
- The existing file_action_selected_paths_file_lifecycle selftest at RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp:2912-3000, registered at :15350, covers the normal success lifecycle.

#### Simplest safe repair

Represent the inventory as a move-only SelectedPathsFileLease:

- construction owns the path and deletes it on destruction;
- successful launch either keeps the lease until a process wait completes or transfers it to a deferred cleanup object;
- every return after CreateProcess/ShellExecute explicitly consumes or retains the lease;
- waiter allocation/registration should be prepared before launch when the API permits;
- timeout and wait errors transfer to deferred cleanup instead of abandoning the path.

When a launch API can succeed without returning a process handle, do not delete immediately because the child may not have opened the file yet. Use a bounded delayed owner and an age-based startup scavenger restricted to the application’s exact rsa-prefixed temp files. Scavenging must be ownership-safe and must not follow arbitrary reparse points.

Do not spread DeleteFileW calls over error branches. The lease destructor and one deferred owner should be the only deletion mechanisms.

#### Required tests

Add injectable seams for:

- inventory write failure;
- process launch failure;
- cleanup allocation failure;
- wait registration failure;
- null process handle after successful launch;
- wait timeout and WAIT_FAILED;
- GetExitCodeProcess failure;
- application restart scavenging an old owned file but preserving a young/live one.

Focused gate:

    .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter file_action_selected_paths_file_lifecycle -TimeoutMultiplier 2.0

#### Specification and ownership

Update the selected-path-file contract in Specs/UI/UI_CommandMenuKeyboard.md:692-705 and the persisted File Action contract in Specs/Core/Core_SettingsStore.md:457 and :474 with ownership transfer, timeout, privacy retention, and scavenger boundaries.

#### Done / STOP

- Done when fault injection proves every exit has one cleanup owner.
- STOP if the chosen process-launch API provides neither a handle nor a trustworthy completion signal; define and document the maximum retention fallback before implementation.

### R4-UI-01 — Use the same animation transform for theme overlay paint and hit testing

**Priority:** P2
**Impact:** the overlay can capture clicks where nothing is painted or fail to capture clicks on the painted surface
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- RedSalamander/Ui/ThemeCycleOverlayWindow.cpp:328-332 applies the animation transform to the painted surface.
- Animation state updates at :362-366 and :388-391.
- Initial state at :175-177 uses opacity 0, scale 0.94, and a positive translation.
- The window is already shown while that state is active at :979-996.
- Hit testing at :1437-1476 uses the final, untransformed surface rectangle.
- Specs/UI/UI_ThemeCycleOverlay.md:48-53 and :82-84 require only the visible surface to consume input.
- The current geometry test in RedSalamander/SelfTest/Commands/Commands.SelfTest.ThemeOverlay.cpp:414-423 checks static center/gutter geometry, then drives animation. It does not compare transformed paint and hit geometry.

#### Simplest safe repair

Create one small SurfaceTransform calculation that returns:

- transformed surface bounds;
- opacity;
- an inverse point mapping for hit testing.

Paint and hit testing must call the same calculation. At effective opacity zero, pass input through. Apply the rounded-corner containment rule after inverse mapping and use the same DPI basis as paint.

Do not maintain a second hand-written set of scale/translation equations in WM_NCHITTEST.

#### Required tests

Use state-specific geometry:

- at opacity zero, every point passes through;
- for a shrunken/translated state contained by the final rectangle, a point in the final-only region passes through and points in the painted surface are handled;
- only for a state whose transform actually extends beyond the final rectangle, a point in the transformed-only region is handled;
- reduced motion and any identity-transform recovery state use the static final geometry;
- center, gutter, rounded corners, and DPI changes always match the surface actually painted.

Focused gates:

    .\build.ps1 -ProjectName RedSalamander -Configuration Debug
    .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter theme_cycle_overlay_rendering_geometry_recovery -TimeoutMultiplier 2.0

#### Specification and ownership

Update Specs/UI/UI_ThemeCycleOverlay.md to say hit testing uses the current painted transform and opacity. Keep the existing timer-cookie, teardown, and reduced-motion implementation; they were reviewed and found sound.

#### Done / STOP

- Done when one transform function is the source of truth for paint and input.
- STOP if Direct2D and Win32 coordinates use different rounding that changes the edge by more than one physical pixel; specify the edge rule before accepting the test.

### R4-CI-01 — Make Winget release publication resumable after PR creation

**Priority:** P2
**Impact:** after a transient post-creation failure, normal reruns continue to fail
while the created PR remains open
**Effort / repair risk / confidence:** S / Medium / High

#### Evidence

- .github/workflows/winget-release.yml:231-238 aborts when a matching PR already exists.
- The workflow creates the PR at :240-257.
- Output parsing, body PATCH, and verification can fail at :259-318.
- A rerun after any of those failures sees the PR and aborts.
- The workflow has no per-package/version concurrency group. Two manual/release runs
  can both observe zero matching PRs and race to create.
- Tools/Tests/WingetValidation.Tests.ps1:127-161 validates text patterns, not the release state transition.
- Specs/Installer/WingetIntegration.md:123 and :134-138 currently specify stop-on-duplicate behavior.

#### Simplest safe repair

Extract one testable decision helper for package/version plus stable automation
ownership (expected author, head branch/repository, and an automation marker):

- zero exact PRs: create;
- exactly one exact same-release, automation-owned PR: resume it;
- one externally owned exact-version PR: stop and report it; never patch/adopt it;
- more than one exact PR: stop and list them;
- unrelated package/version PRs: never adopt.

After wingetcreate output parsing fails, use direct PR listing with bounded retries to
find the exact automation identity; do not depend on one eventually consistent search
query. Creation and resume then share the same idempotent body PATCH and verification
path and always emit the final PR URL.

Resolve and validate the normalized package/version in a predecessor job. Put the
publication job in a job-level concurrency group derived from the resolver's needs
output, with queue: max so multiple pending retries are retained:
https://docs.github.com/en/actions/how-tos/write-workflows/choose-when-workflows-run/control-workflow-concurrency
Concurrency prevents the ordinary zero-to-create race; the ownership-qualified
adoption query remains required for crash recovery and external/upstream races.

Do not match by a loose title substring. Preserve the release-tag checkout constraint in WingetIntegration.md:119; a helper added only on a newer branch cannot be assumed to exist in an older release tag.

#### Required tests

- no PR, create succeeds;
- PR created but output malformed, exact query resumes;
- rerun with exactly one matching PR patches/verifies it;
- unrelated PR is ignored;
- an external contributor's exact package/version/title PR stops without PATCH;
- two exact PRs stop with actionable output;
- PATCH retry is idempotent.
- two concurrent attempts serialize, create at most one PR, and both converge on the
  same verified URL;
- three queued retries are preserved in FIFO waiting order rather than replacing the
  older pending run;
- direct listing retries converge when the newly created PR is not visible immediately.

Focused gates:

    Invoke-Pester .\Tools\Tests\WingetValidation.Tests.ps1
    .\Tools\Run-AllTests.ps1 -Suite CI -SkipBuild

#### Specification and ownership

Update Specs/Installer/WingetIntegration.md with the resumable state machine and exact-match rule.

#### Done / STOP

- Done when every failure after successful PR creation can be rerun safely.
- STOP if the upstream tool cannot provide or query a stable package/version identity; do not resume using a fuzzy match.

### R4-TOOL-01 — Collapse ADS removal into one correct command

**Priority:** P2
**Impact:** one script cannot be loaded as written and the duplicate reports misleading WhatIf results
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- Tools/remove-ads.ps1:37-46 declares SupportsShouldProcess and also declares an explicit WhatIf parameter.
- PowerShell adds the common WhatIf parameter for SupportsShouldProcess, so Get-Command .\Tools\remove-ads.ps1 -Syntax reports that WhatIf is defined multiple times.
- Tools/remove-stream.ps1 duplicates nearly all implementation at :32-78.
- Both scripts suppress stream-enumeration errors at remove-ads.ps1:54-60 and :75 and remove-stream.ps1:52-58 and :73.
- remove-stream counts WhatIf matches as skipped, so its summary can say it would remove zero.
- remove-stream help at :25-33 references remove-ads, demonstrating drift.
- No in-repository caller or regression test was found.

#### Simplest safe repair

Keep Tools/remove-ads.ps1 as the canonical implementation:

- remove the explicit WhatIf parameter;
- rely on CmdletBinding SupportsShouldProcess;
- track matched, removed, skipped, and failed separately;
- make enumeration failures visible;
- print one summary and return nonzero after any failure.

Keep remove-stream.ps1 only as a thin compatibility wrapper if external use is plausible. It should forward parameters to the canonical script without reimplementing logic. Delete it only after an explicit compatibility decision.

#### Required tests

Use Pester and a temporary NTFS directory:

- syntax discovery succeeds;
- create/remove Zone.Identifier and a named alternate stream;
- name filter and recurse;
- WhatIf changes nothing and reports matched/would-remove counts;
- missing/unreadable path reports a failure and nonzero status;
- compatibility wrapper behavior, if retained.

Focused gates:

    Get-Command .\Tools\remove-ads.ps1 -Syntax
    Get-Command .\Tools\remove-stream.ps1 -Syntax
    Invoke-Pester .\Tools\Tests

Skip ADS runtime cases with an explicit reason when the filesystem does not support named streams.

#### Specification and ownership

Update the Tools inventory/README with the canonical command and compatibility status.

#### Done / STOP

- Done when there is one implementation and WhatIf/error summaries are truthful.
- STOP before deleting the compatibility filename unless repository release policy confirms it is not a supported external entry point.

### R4-TEST-01 — Bring the committed TestRun back inside its own archive contract

**Priority:** P2
**Impact:** oversized raw evidence bypasses the reviewability/storage guard intended for all performance work
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- Specs/TestRuns/7d3a1247382a/FileOps/2026-07-20_205550/perf/perf_metrics.jsonl is 7,009,000 bytes.
- That run is approximately 7,067,986 bytes in total.
- Tools/TestRunArchive.psm1:36-38 caps an individual file at 2,097,152 bytes and a run at 5,242,880 bytes.
- The module reports those violations at :57-64 and :103-119.
- The archive was added in commit 148f during the four-day window.
- The completed Microsoft Drive plan at Specs/Plans/Done/Operation_MicrosoftDrive_AmbiguousMutationReconciliation_2026-07-21.md:299-303 already records that the archive Pester gate failed.
- Tools/Tests/TestRunArchive.Tests.ps1:87-92 correctly includes newly changed archives. The completed plan records the failure, so this archive was committed despite the failed gate; it was not missed by changed-set discovery.
- Once an oversized archive is no longer in the changed set, that focused test does not continue reporting the historical violation. The repository therefore needs enforcement at production/staging time plus a whole-inventory closeout gate.

#### Simplest safe repair

Replace the raw event stream with bounded evidence:

- retain README, result summary, configuration/commit identity, and the key request/resource proof;
- aggregate repeated metrics;
- keep a small representative sample only when raw ordering is material;
- validate the exact archive directory before staging it.

Then fix the producer or pre-stage command so it cannot emit/accept a run above the policy caps, and make the closeout/CI inventory gate non-bypassable for unallowlisted violations. Do not merely raise the caps to fit this run.

#### Required tests

- Exact offending archive is rejected before compaction.
- Compacted archive passes and still supports the claims made by its README.
- Producer/pre-stage validation rejects a newly generated oversized archive before it is accepted.
- Whole-inventory closeout detects an unchanged historical violation.
- A failed archive gate cannot be reported as acceptable closeout evidence.
- Producer/pre-stage path returns nonzero and names the exact file/run limit.

Focused gates:

    .\Tools\Test-TestRunArchive.ps1
    Invoke-Pester .\Tools\Tests\TestRunArchive.Tests.ps1
    .\Tools\Run-AllTests.ps1 -Suite CI -SkipBuild

#### Specification and ownership

Keep Specs/TestRuns/README.md:12 and Specs/Testing/Testing_PerformanceValidation.md:172-178 authoritative. Coordinate the producer/gate fix with Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md.

#### Done / STOP

- Done when the existing run and every newly generated run pass the unchanged bounds.
- STOP if compaction removes evidence needed for the Microsoft Drive claim; replace it with a small deterministic summary, not a larger cap.

### R4-TEST-02 — Preserve ExpectedSize through the bridge writer decorator

**Priority:** P2
**Impact:** the test bridge bypasses a production upload-planning contract and can pass while production takes a different whole-file path
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- SelfTestBridgeFileWriter in RedSalamander/FolderWindow.FileOperations.State.cpp:503-604 exposes the writer and commit-proof interfaces only.
- Test destinations wrap their writers at :690-727 and :7399-7409.
- Production queries ExpectedSize at :9240-9247.
- Microsoft Drive uses expected size at Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:4877-4917 and :4949-4964 to choose upload behavior.
- Without that contract, it can fall back to whole-file spooling at :6872-6918.

The decorator changes the interface surface of the object under test, so the File Operations bridge proof is not exercising the production path it claims.

#### Simplest safe repair

Have SelfTestBridgeFileWriter query the wrapped writer once for ExpectedSize support:

- advertise/proxy the interface exactly when the inner writer supports it;
- forward SetExpectedSize unchanged;
- record that it occurred before the first Write;
- preserve current writer and commit-proof behavior.

Do not add Microsoft Drive knowledge to the bridge decorator. It should transparently preserve the inner contract.

#### Required tests

- Supporting inner writer receives the exact size before first Write.
- Non-supporting writer does not advertise the interface.
- QueryInterface identity/lifetime remain correct.
- Cross-filesystem move/copy test proves the streaming path and commit-size proof.
- A source-contract guard prevents future decorators from hiding the interface.

Focused gates:

    .\build.ps1 -ProjectName RedSalamander -Configuration Debug
    .\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -TimeoutMultiplier 2.0
    Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1

#### Specification and ownership

Route this row through Operation Floodgate because it owns cloud parity and the bridge commit/size proof. The general test-contract architecture remains with Operation Astrolabe; do not migrate this focused runtime test into a source-only assertion.

#### Done / STOP

- Done when the selftest wrapper is observationally transparent for ExpectedSize.
- STOP if interface aggregation changes COM identity; preserve the existing outer object identity and proxy only method calls.

### R4-ARCH-01 — Centralize bounded WSL registry/catalog discovery

**Priority:** P2
**Impact:** malformed or long registry strings can cause an out-of-bounds read or inconsistent distro discovery; duplicate policy will continue to drift
**Effort / repair risk / confidence:** M / Medium / High

#### Evidence

- RedSalamander/WSLDistro.cpp:146-161 reads DistributionName into an uninitialized fixed wchar_t[256] buffer, then constructs std::wstring(buffer) after RegQueryValueExW succeeds.
- Plugins/FileSystem/FileSystem.Menu.cpp:95-108 repeats the same pattern.
- Microsoft states that REG_SZ data returned by RegQueryValueEx is not guaranteed to be null-terminated and recommends RegGetValue when strings must be terminated: [RegQueryValueExW documentation](https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryvalueexw).
- Long values are silently dropped by the fixed buffer.
- The full Lxss key enumeration, GUID/filter/sort, path, and WSL-version policy is duplicated between RedSalamander/WSLDistro.cpp:91-164 and FileSystem.Menu.cpp:32-180.
- RedSalamander/FolderWindow.FileSystem.Commands.cpp:256-303 has a safer consumer-local reader, creating a third near-solution rather than one policy owner.
- HKEY handles themselves use RAII correctly.

#### Simplest safe repair

Add the lowest-layer Common helpers that are actually shared:

1. A bounded registry-string reader using RegGetValueW or a two-pass size/read operation.
2. A WSL catalog function returning stable records: distribution name, GUID, base path, and WSL version.

The reader must:

- cap allocation;
- use the returned byte count;
- reject odd byte counts and wrong types;
- normalize one or more trailing terminators;
- handle an empty string;
- retry a bounded number of times if the value changes between size/read.

The catalog should own one exclusion/filter/sort policy. UI consumers may add icons and presentation state afterward.

Do not add another WSL class hierarchy. A value record plus one enumeration function is enough.

#### Required tests

Use a volatile per-user test key:

- terminated and non-terminated REG_SZ;
- exact-fit, long, empty, wrong-type, odd-byte, and changing values;
- Docker/Rancher exclusions;
- deterministic case-insensitive sort;
- WSL1/WSL2 record values;
- source-contract proof that both consumers call the canonical catalog.

Focused gates:

    .\build.ps1 -ProjectName RedSalamander -Configuration Debug
    .\build.ps1 -ProjectName FileSystem -Configuration Debug
    .\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2

#### Specification and ownership

Update Specs/Core/Core_SharedHelpers.md and the authoritative WSL/FileSystem integration spec. Treat `Specs/Plans/Done/Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md` only as archived consumer context; current Terminal authority is `Specs/Terminal/Terminal_EmbeddedPlugin.md`, and Terminal does not own the unsafe shared readers tracked here.

#### Done / STOP

- Done when there is one bounded reader and one WSL catalog policy used by all three consumers.
- STOP if plugin dependency layering cannot reference the proposed Common location; choose the lowest existing shared library that both app and plugin already consume and document that constraint.

### R4-ARCH-02 — Share one delete-on-close temporary-file reservation helper

**Priority:** P2
**Impact:** Curl can leave a reserved temporary file when reopening it fails; three copies already differ in cleanup behavior
**Effort / repair risk / confidence:** S / Low / High

#### Evidence

- Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp:2514-2537 calls GetTempFileNameW, then returns an invalid handle if CreateFileW fails without deleting the reserved path.
- Plugins/FileSystemS3/FileSystemS3.Shared.cpp:403-426 deletes the reservation on reopen failure.
- Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:951-978 also cleans the reservation.

All three implement the same reservation, delete-on-close, sequential-access policy with small variations.

#### Simplest safe repair

Add one Common::Files helper such as CreateDeleteOnCloseTempFile:

- input: prefix and additional file flags;
- reserve a unique temp path;
- immediately own cleanup with RAII;
- reopen with FILE_ATTRIBUTE_TEMPORARY and FILE_FLAG_DELETE_ON_CLOSE plus requested flags;
- return HRESULT and an owning WIL handle;
- release the path cleanup only after the delete-on-close handle is valid.

Migrate all three consumers in the same change. Preserve provider-specific prefixes and sequential/random access flags.

#### Required tests

- Inject failure after reservation and before reopen; no file remains.
- Successful handle deletes on close.
- Prefix and requested flags are preserved.
- Concurrent creation does not collide.
- Each provider contract test still passes.

Focused gates:

    .\build.ps1 -ProjectName FileSystemCurl -Configuration Debug
    .\build.ps1 -ProjectName FileSystemS3 -Configuration Debug
    .\build.ps1 -ProjectName FileSystemMicrosoftDrive -Configuration Debug
    .\build.ps1 -ProjectName PluginContractTests -Configuration Debug
    .\.build\x64\Debug\PluginContractTests.exe

#### Specification and ownership

Add the helper and its semantics to Specs/Core/Core_SharedHelpers.md and update the relevant provider specifications only where observable behavior changes.

#### Done / STOP

- Done when all three copies are removed and the failure-injection test leaves no reservation.
- STOP if a provider intentionally needs different sharing, security, or delete semantics; keep that difference as a named policy input, not a copied function.

### R4-MON-04 — Move Monitor Save As off the UI thread without losing snapshot semantics

**Priority:** P2
**Impact:** a large save can freeze the UI, block ETW ingestion, and delay close/cancel
**Effort / repair risk / confidence:** L / High / High

#### Evidence

- RedSalamanderMonitor/RedSalamanderMonitor.cpp:2640-2657 performs Save As synchronously from the UI command.
- RedSalamanderMonitor/ColorTextView.cpp:845-847 delegates synchronously.
- RedSalamanderMonitor/Document.cpp:950-966 holds the document shared mutex across every per-line UTF-8 conversion and transaction Write. It explicitly unlocks before transaction.Commit at :966-967.
- ETW append/eviction needs the unique side of that mutex, so export blocks live ingestion through the potentially dominant conversion/write duration even though final commit happens after unlock.
- The configured retained document size makes the stall user-visible even on a healthy disk.

The previously completed Operation Blueprint row BP-B5 owns the product decision about visible text versus raw payload. That semantic question is not duplicated here. This row owns responsiveness, cancellation, snapshot consistency, and lock duration.

#### Simplest safe repair

Use one owned cancellable export job:

- keep only the file picker and job start on the UI thread;
- define snapshot-at-start semantics;
- iterate a stable snapshot in bounded chunks without holding the document lock across disk I/O;
- write transactionally and publish the target only on success;
- report progress/completion through the registered posted-payload mechanism;
- make close request cancellation and wait only to a documented bound.

Avoid a second whole-document string copy. The preferred representation is immutable line ownership or monotonic line IDs that let the job retain the exact start snapshot while the live document evicts. If the current model cannot preserve an exact snapshot cheaply, abort the target transaction on invalidation instead of silently exporting a mixture.

Do not post raw pointers, use a detached thread, or hold a shared lock for the duration of the save.

#### Required tests and performance proof

- Large export while ETW lines continue to append.
- Snapshot content is exactly start-of-export content.
- Cancel before first write, mid-conversion, and mid-write.
- Close during export has a bounded latency and drains posted payloads.
- Target file remains unchanged after cancel/failure.
- UI command returns within the responsiveness budget.
- Ingestion queue/drop counters do not regress materially during export.

Add an ENABLE_TESTS-gated application switch named --monitor-export-selftest. Use deterministic barriers in an injected export sink to pause after the first write, drive append/cancel/close, and verify the posted completion payload and target transaction. Register the switch in Tools/TestRunPlan.ps1 for CI and Full.

Focused gates:

    .\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug
    .\.build\x64\Debug\RedSalamanderMonitor.exe --monitor-export-selftest
    .\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2

Archive same-machine before/after UI return latency, ingestion impact, cancel latency, and peak retained memory.

#### Specification and ownership

First resolve or reference the BP-B5 output semantic decision in Specs/Plans/Done/Operation_Blueprint_NormativeSpecAccuracyRemediation_2026-07-03.md. Then update Specs/Core/Core_RedSalamanderMonitor.md with snapshot, progress, cancellation, teardown, target-transaction, and performance rules.

#### Done / STOP

- Done when export never holds the document mutex across disk I/O and exact snapshot/cancel behavior is proven.
- STOP until visible-versus-raw output semantics are settled; do not build an asynchronous pipeline around an undefined payload.

### R4-S3-03 — Parked: reject unsupported giant uploads before transfer

**Priority:** P3 parked
**Impact:** a transfer near 625 GiB can fail only when attempting part 10,001
**Effort / repair risk / confidence:** S for early rejection, M for adaptive sizing / Medium / High

#### Evidence

- Plugins/FileSystemS3/FileSystemS3.Internal.h:348 fixes the writer part size at 64 MiB.
- FileSystemS3.IO.cpp:1492-1494, :1786, and :1790-1794 use that fixed size.
- The writer rejects part 10,001 at :1815-1818, after approximately 625 GiB has already transferred.
- A valid part-size calculation exists at FileSystemS3.S3.cpp:129-135, :932-940, and :1008.
- The writer interface at FileSystemS3.IO.cpp:1369-1413 does not expose ExpectedSize.

#### Simplest safe next step

Implement ExpectedSize for this writer and reject an unsupported known size before upload initiation. Then choose one explicit unknown-size policy:

- require ExpectedSize for S3 writers and reject a missing size before the first Write; or
- implement adaptive legal part sizing for unknown-size streams.

The first option is the smaller, honest product limit and reuses the same contract needed by R4-TEST-02. Merely adding ExpectedSize while continuing to accept unknown-size streams is only a partial mitigation and does not close this row.

Adaptive legal part sizing is a larger follow-up. It must remain within the configured memory budget and preserve retry/resume semantics. Do not switch part size after upload begins.

#### Required tests

- Exact maximum supported size accepted.
- One byte above rejected before upload initiation/first write.
- Under the preferred simple policy, missing ExpectedSize is rejected before the first Write with an actionable error.
- If adaptive sizing is chosen later, validate part count, minimum/final part rules, memory, and retry.

#### Specification and ownership

Record the supported maximum in Specs/FileSystem/FileSystem_S3.md and route through Operation Floodgate.

#### Done / STOP

- Done when both known and unknown-size entry paths avoid a transfer that can reach part 10,001.
- STOP before implementing adaptive sizing without an expected-size contract and memory/performance proof.

## 7. Existing-plan reconciliation

This report is an evidence and execution owner for newly identified rows. It must not become a second implementation owner for work already routed elsewhere.

| Existing owner | Treatment in this review |
|---|---|
| Specs/Plans/Done/CodeReview_Last4Days_IndependentFindingsAndRemediation_2026-07-16.md | Baseline reviewed. No prior row is restated merely because it remains interesting. |
| Specs/Plans/Done/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md | Historically owned EG-S3-01, R4-S3-02, R4-TEST-02, and R4-S3-03 implementation. This Emberglass plan historically owned EG-S3-02 and EG-MSD-02 implementation; Floodgate consumed their completed cloud proof without creating duplicate work packages. Closed R4-S3-01 remains historical and is not reopened by the distinct direct-transfer paths. |
| Specs/Plans/Done/Code_PluginImprovementPlan.md | Its strong writer-to-plugin ownership contract is controlling precedent for EG-CURL-01; the completed plan is not reopened. |
| Specs/Core/Core_SettingsStore.md plus the completed Track 7 history | R4-SET-01 remains valid for cooperating writers but is reopened only at the external-writer boundary proven by EG-SET-02, which owns the remaining remediation. EG-SET-01 separately repairs early read-state classification. |
| Specs/Plans/WIP/Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md | R4-SEARCH-01's distinct startup guard is closed with store-scoped ownership. Parallax records that PAR-1 remains an adjacent deferred named-pipe impersonation review, not a duplicate owner. |
| Specs/Plans/WIP/Operation_RosettaLantern_RedConfigureSatelliteLocalizationReview_2026-07-17.md | Coordinates culture identity/spec changes for R4-RC-01; translation inventory does not own paste correctness. |
| Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md | Owns R4-TEST-01 producer/gate hardening. |
| Specs/Plans/WIP/Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md | Continues to own general source-contract architecture. R4-TEST-02 still requires runtime proof. |
| Specs/Plans/Done/Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md | Archived Terminal consumer context only; current authority is `Specs/Terminal/Terminal_EmbeddedPlugin.md`. Terminal does not own the unsafe readers. EG-GOV-01 protects the historically reviewed seven-file Terminal plan diff from accidental absorption into runtime remediation. |
| Specs/Plugins/Plugins_VirtualFileSystem.md and Specs/Core/Core_SharedHelpers.md | Own the durable packed `FileInfo` validation and helper-reuse contract for EG-ABI-01. This plan owns the implementation package; do not create a second ABI migration ledger. |
| Specs/Plans/WIP/Win32_Inventory.md | Continues to own the existing wil::unique_hwnd DestroyWindow findings in Tests/PerformanceTests2/FolderViewInternal.Access.h:769 and :779. No duplicate row was created here. |
| Specs/Plans/Done/Operation_Blueprint_NormativeSpecAccuracyRemediation_2026-07-03.md | BP-B5 remains the source for Monitor visible/raw export semantics. R4-MON-04 adds only liveness/snapshot work. |
| plans/012-format-autocommit-race.md | Continues to own the format auto-commit fork/race. Emberglass adds three pieces of contrary evidence there: github.event.before can be all-zero on a new branch and unavailable after a force-push; format-all.ps1 records an individual clang-format failure but can return the last invocation's success and publish a partial rewrite; and CI installs floating LLVM then selects the newest executable. EG-SEC-02 coordinates the format job's credential scope but does not duplicate its push/lease repair. |

The WIP README retains the broad completion history for Tracks 6, 7, 11, and 19 but
must narrow those statements with this review's exact contrary rows. Do not reopen any
unrelated completed work. The changed Docs/Reviews Curl link must point to the Done
owner; moving a plan to Done is not a reason to leave a dead route.

## 8. Reviewed areas with no new actionable finding

The following checks reduced false positives and should be preserved during remediation:

- No production catch (...) was found; the only hit was a source-contract string.
- No non-PoC sprintf_s or swprintf_s use was found.
- No raw PostMessageW new/release payload pattern was found.
- No AddRef plus attach two-step COM ownership pattern was found.
- No direct std::thread(...).detach production pattern was found.
- No repository or upstream credential value was found in source, logs, archived test
  evidence, or workflow fixtures.
- Microsoft Drive Missing-hint late creation fails closed, and ambiguous-result
  reconciliation after a mutation request appeared internally consistent. EG-MSD-01 is
  the separate stale Present identity before the request.
- S3 writer worker ownership, ordered completion, first-failure preservation, lease
  release, and join behavior appeared sound outside the inherited abort-queue
  convergence defect. EG-S3-01 is the separate direct-transfer publication path.
- The bridge IFileWriterCommitSizeProof decorator does not bypass source deletion
  safety: cleanup reopens and validates size/hash before removing the source.
- Curl's new short-success fixture correctly preserves the staged-upload failure
  contract; no second defect was found in that proof path beyond EG-CURL-01's owner
  lifetime.
- SettingsHotReload stale notifications are generation-guarded; no stale-notification
  overwrite was confirmed.
- Theme overlay timer cookies, teardown, reduced-motion handling, and source filtering appeared sound.
- File Operations queue-reorder state is lock-protected in the reviewed paths.
- WSL registry-key handles use RAII; R4-ARCH-01 is about string bounds and duplicated policy.
- No distinct RedLauncher defect was confirmed in the reviewed window.
- The ThroughputGraph branch that initially appeared redundant preserves distinct
  sample-window behavior; no simplification was justified.
- Unsigned MSIX creation remains an explicitly documented optional mode, not a newly
  introduced release defect.
- The Winget validation workflow's final uninstall command propagates its exit status;
  no masked uninstall failure was confirmed.
- All changed vcpkg references resolve to full immutable commits.
- Archived failed TestRuns identify their failed/incomplete state; they were not
  misrepresented as passing evidence.
- Monitor visible/raw export content is an existing product decision, so it was not relabeled as a new correctness bug.
- Large source files were not treated as defects by size alone.

## 9. Considered but rejected or deferred

### 9.1 Full rebuild on every Monitor append

Rejected. It would hide R4-MON-02 by violating the bounded hot-path requirement. An explicit scan frontier is simpler in operational behavior and cheaper.

### 9.2 Another HEAD immediately before S3 publication

Rejected. It remains a time-of-check/time-of-use race. The conditional belongs on PutObject/CompleteMultipartUpload.

### 9.3 Polling thread for multipart cleanup

Rejected. One delayed threadpool continuation is enough and preserves existing module-lifetime patterns.

### 9.4 Delete selected-path inventory immediately when no process handle is returned

Rejected. A successfully launched process may not have opened the file yet. Ownership needs a bounded delayed fallback, not an unsafe early delete.

### 9.5 New WSL object framework

Rejected. A bounded string helper and a value-record catalog remove the duplication without introducing an unnecessary hierarchy.

### 9.6 Raising TestRun archive limits

Rejected. The offending payload is highly repetitive raw telemetry; compact evidence is easier to review and retains the intended policy.

### 9.7 Adaptive S3 multipart sizing in the first repair

Deferred. An early, explicit supported-size rejection closes the late-failure defect with much lower risk. Adaptive sizing can follow with separate memory/retry evidence.

### 9.8 Treating a fresh Microsoft Drive GET as the commit condition

Rejected. It narrows EG-MSD-01 but still permits a change between GET and PATCH. The
mutating request must validate the observed revision/identity or fail closed.

### 9.9 Theme fallback post as a stale-window payload defect

Rejected after lifetime tracing. The overlay implementation survives window
recreation, fallback cookies are monotonic, teardown cancels and waits, and a queued old
cookie cannot match the new transition. The timer/cookie tests cover the relevant
pieces. R4-UI-01 remains valid only for painted-versus-hit-test geometry.

### 9.10 Bridge commit-size decorator as a MOVE source-deletion bypass

Rejected. The production cleanup path does not trust the hidden ExpectedSize alone; it
reopens the destination and validates size/hash before deleting the source. R4-TEST-02
still matters because the decorator fails to exercise the intended proof path, but it
is test fidelity rather than a confirmed deletion bypass.

### 9.11 S3 multipart lease retention as a permanent leak

Rejected. Lease ownership transfers through the worker/completion path and releases on
the reviewed success, failure, cancellation, and join exits. The confirmed lifecycle
issue remains R4-S3-02's abort-queue convergence.

### 9.12 Microsoft Drive Missing hint as a silent late overwrite

Rejected. A late destination creation after a cached Missing observation reaches a
conditional/fail-closed mutation and is covered by the ambiguous mutation tests. The
confirmed EG-MSD-01 case requires a stale Present identity that is mutated by item ID.

### 9.13 Triggering Winget with a broad PAT-created release event

Rejected. It works around GitHub's recursion guard by expanding credential scope and
event indirection. EG-CI-01's direct reusable-workflow dependency is simpler,
observable, and least-privileged.

### 9.14 Fixing settings backup with one more path restat

Rejected. A second path stamp check creates the same compare-to-rename gap at a later
line. EG-SET-02 requires retaining the identity-bearing handle through the rename.

### 9.15 Hard-coding a second File Operations tab-ID list

Rejected. Task controls are dynamic and direction/layout dependent. EG-UI-01 should
derive traversal from final visual geometry or use one explicit host navigation surface,
not maintain a parallel static list.

## 10. Verification matrix

Each package carries focused assertions. The minimum Emberglass matrix is:

| Packages | Focused proof | Non-regression boundary |
|---|---|---|
| EG-MSD-01 | Barrier after Present probe; second session moves/replaces item | Moved-away identity/content untouched; merge MOVE conflict converges |
| EG-MSD-02 | Valid and malformed `200`/`206` debug responses | No byte exposed from shifted, short, empty, or inconsistent ranges |
| EG-S3-01 | Inject create after refresh for small and multipart direct copy | Destination unchanged; MOVE source retained; safe fallback performance recorded |
| EG-S3-02 | Replace source after planning, between parts/relay, and before MOVE delete | One exact source revision copied; replacement retained; conditional delete only |
| EG-FOPS-01 | Cached Keep Both plus later nested collision, sequential and parallel | No whole-root retry, duplication, or split MOVE tree |
| EG-CURL-01 | Release filesystem before writer Commit/cancel/failure | Owner/unload lifetime remains valid through writer destruction |
| EG-CURL-02 | Two limiter keys with unequal caps and fixed transport latency | Per-key caps preserved; independent keys overlap; nested work stays bounded |
| EG-SET-01/02 | Oversize/read/open failures and external-writer barrier | No silent default recovery, stuck save, or wrong revision in backup |
| EG-UI-01/02 | Exact forward/reverse traversal and thousands of task lifecycles | Visual order is stable; retained map bounded by live tasks |
| EG-SEC-01 | Metacharacter event corpus through workflow parser/harness | Input remains inert; no pre-secret environment/PATH/executable mutation |
| EG-SEC-02 | Parsed permission/checkout matrix | Only final publication jobs can authenticate writes |
| EG-SEC-03 | Wrong/missing WingetCreate version plus captured final argv | Failure is pre-secret; verified argv contains no credential/token switch |
| EG-CI-01/02 | Trigger-graph and release-state harnesses | Direct Winget edge; missing MSIX cannot leave public partial release |
| EG-BUILD-01 | Unexpected compiler identity injection | Build stops before compile; successful artifact records exact identity |
| EG-DOC-01 | Repository-route and documented-command contract tests | No dead linked/backticked routes or contradictory configuration claim |
| EG-ABI-01 | Malformed packed-record corpus plus valid terminal/multi-record controls | Every field is bounds-checked before read; no overlap, count drift, or OOB copy |
| EG-GOV-01 | Staged-path review for runtime and Terminal plan commits | No unrelated Terminal/workspace-policy paths in runtime remediation |

Workflow tests must use stubs/dry runs and test repositories without production tokens;
negative tests must never publish a real release, mutate winget-pkgs, or expose a secret
in logs/process arguments.

Before closeout, run the applicable project tests followed by this repository-wide sequence from a clean worktree:

    git diff --check
    .\Tools\Test-TestRunArchive.ps1
    Invoke-Pester .\Tools\Tests
    .\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2
    git status --short

Expected results:

- git diff --check has no output.
- TestRun archive validation reports no file/run over policy limits.
- Pester has no failures; environment-dependent skips include an explicit reason.
- Full suite is green with no timeout, crash artifact, or missing result.
- git status lists only the intended implementation, tests, authoritative specs, bounded evidence, and plan/index closeout changes.

For performance-sensitive packages EG-S3-01, EG-S3-02, EG-CURL-02, EG-SET-01,
EG-UI-01, EG-UI-02, R4-MON-01, R4-MON-02, and R4-MON-04:

1. Define the scenario and counters before implementation.
2. Capture a same-machine baseline.
3. Run the focused deterministic selftest.
4. Capture the after run with the same configuration.
5. Archive only bounded summaries and representative evidence.
6. Link the archive from the authoritative domain specification.

For EG-SET-01 and EG-UI-01/02, reuse App.Startup.LoadSettings,
FileOps.Popup.HostedControls.SyncUs, and FileOps.Popup.Render.TotalUs respectively.
Archive the bounded baseline/candidate record under Specs/TestRuns and link it from the
SettingsStore or File Operations popup specification.
EG-S3-01 and EG-S3-02 link their direct-transfer baseline/candidate archives from
Specs/FileSystem/FileSystem_S3.md.

For concurrency/interleaving packages EG-MSD-01, EG-S3-01, EG-S3-02, EG-CURL-02,
EG-FOPS-01, EG-SET-01, EG-SET-02, R4-SET-01, and R4-S3-02:

- use barriers/latches and injected responses;
- do not use sleeps as the proof of the interleaving;
- repeat the focused scenario enough to establish determinism;
- include teardown and cancellation in the state matrix.

For workflow trust-boundary packages EG-SEC-01, EG-SEC-02, EG-SEC-03, and EG-CI-01:

- keep untrusted event data out of generated shell source;
- prove a failed precondition/tool assertion occurs before secret exposure;
- verify workflow logs, environment summaries, command lines, and artifacts contain no
  secrets;
- prove write authority and persisted authentication are absent from non-publishing
  jobs/steps.

For boundary-validation packages EG-MSD-02 and EG-ABI-01:

- build valid controls and malformed cases from raw response/record bytes;
- verify every length, offset, total, alignment, and count before dereference/copy;
- use overflow-edge values without allocating attacker-declared payload sizes;
- run the malformed corpus under the available ASan configuration in addition to the
  focused Debug contract test.

For EG-GOV-01:

- inspect `git diff --name-only` before staging;
- keep the seven pre-existing Terminal plan edits untouched by runtime executors;
- require the staged path set to match the active package's explicit scope;
- do not rewrite published history solely to improve commit grouping.

Closed R4-SEARCH-01 evidence preserves these review requirements:

- prove both authorized exclusion and unauthorized precreation resistance;
- avoid a test that passes only because it ran elevated;
- verify ACL/owner diagnostics contain no secrets.

For UI packages EG-FOPS-01, EG-UI-01, EG-UI-02, and R4-UI-01:

- test dynamic insertion/removal and both forward/reverse navigation;
- cover sequential and parallel task execution where state is cached;
- retain a deterministic accessibility/focus trace;
- verify teardown with queued work and no retained per-task state.

## 11. Final definition of done

This review is complete only when:

1. All eight Emberglass P1 rows and every inherited live P1 row are closed, or have an
   explicit product-approved deferral with a safe disabled/fail-closed behavior.
2. Every Emberglass and inherited P2 row has either landed or moved to one named live
   owner with its failure proof and acceptance criteria preserved.
3. R4-SET-01 and EG-SET-02 close together only after the external-writer retained-handle
   proof; the earlier cooperating-writer proof remains part of the evidence.
4. R4-S3-01 remains historically closed, EG-S3-01 closes the distinct destination
   publication path, EG-S3-02 pins the source revision through copy and conditional
   MOVE deletion, and R4-S3-03 has an explicit early maximum even if adaptive sizing
   remains parked.
5. Plan 012 records the new-branch/force-push base, partial formatter failure, floating
   LLVM, and scoped publication credential evidence without a duplicate Emberglass row.
6. No completion assertion or route in Specs/Plans/WIP/README.md,
   Specs/Reviews/README.md, or an authoritative domain spec contradicts a remaining
   live row.
7. Authoritative domain specifications contain the durable mutation, source-revision,
   ranged-response, packed-ABI validation, ownership, recovery, workflow,
   accessibility, per-connection concurrency, and build-identity contracts.
8. Required performance evidence is present and within archive limits.
9. Focused negative/interleaving proofs and the full verification suite are green.
10. EG-GOV-01 has a reviewed disposition for the current Terminal-only worktree diff
    and later runtime repairs contain no unrelated Terminal/workspace-policy paths.
11. This file is moved from Specs/Plans/WIP to Specs/Plans/Done and every routed owner
    is reconciled.
