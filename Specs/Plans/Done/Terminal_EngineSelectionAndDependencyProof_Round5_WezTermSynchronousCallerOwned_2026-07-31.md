# Terminal Engine Gate 0 Round 5 — WezTerm Synchronous Caller-Owned Boundary

> **Status — Superseded and archived 2026-08-03.** This historical WezTerm proof route was superseded by the qualified
> private Ghostty runtime recorded in `../../Terminal/TerminalEngineDecision.md`. No WezTerm implementation work
> remains. The rest of this file is preserved solely as non-normative decision history.

> **Authority:** Versioned Gate-0 round 5 for
> `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`.
> This is evidence-only engine work. It may select a winner and publish the sole
> `successfulGate0Handoff`; it never authorizes product implementation by itself.

## Status and predecessor gate

- **State:** WIP / ready — the exact Round-4 predecessor is authenticated below;
  the evidence-only proof is authorized, while product implementation remains
  blocked.
- **Round:** 5
- **Experiment:** synchronous caller-owned WezTerm terminal-state/snapshot
  boundary, with unchanged product requirements and Gate-0 scoring.
- **Proof-spike budget:** at most 15 cumulative owner-days to decide only the
  synchronous caller-owned hypothesis. This is a hard stop, not an estimate,
  an admission decision, or an exemption from the unchanged 60-owner-day
  candidate cap applied later to a complete nomination.
- **Completion record:** `TERMINAL-HANDOFF-UNRESOLVED`
- **Round-4 predecessor Done path:**
  `Specs/Plans/Done/Terminal_EngineSelectionAndDependencyProof_2026-07-22.md`
- **Round-4 predecessor closeout commit:**
  `6674119fd24284877b1b25a3a71dbe2f6e6ee9a2`
- **Round-4 predecessor Done blob:**
  `b20d498dcc16d229f1e8a68e26fcc8aae3b6ffca`
- **predecessorHandoffPublishCommit:**
  `N/A-legacy-round-no-split-closeout`
- **predecessorCloseoutSealCommit:**
  `N/A-legacy-round-no-split-closeout`
- **predecessorCloseoutSealSha256:**
  `N/A-legacy-round-no-split-closeout`
- **Round-4 decision SHA-256:**
  `e10e6c5b885fc40496fb3eb1a7db9feb82084a74a368567305374bc0787e8260`

The six predecessor identities are resolved from the immutable Round-4 Done
handoff. Reverify the commit/blob/digest before any v5 universe or candidate
work. The decision digest authenticates the exact decision blob at that
closeout commit; later current-runbook addenda in the mutable operator summary
do not rewrite or replace the historical subject. A branch, working-tree path,
“latest,” copied summary, or mutable WIP file is not evidence.
The three exact split-closeout N/A values are carried into the carry audit,
audited decision, E subject, and FD projection; substituting `none`, an inferred
commit, or a future seal for the legacy Round-4 protocol is invalid.

## Goal and non-goals

Test one precise hypothesis: a narrow WezTerm adaptation can execute all engine
mutation synchronously on the caller, return only caller-owned immutable state,
effects, and static-Kitty snapshot data, and retain no callback, borrowed engine
memory, or engine work after the call returns. If that boundary is proved inside
the proof-spike budget, permit preparation of a complete v5 nomination. Only
after that nomination's exact patch, result tree, descriptor, concern map,
owner ledger, dependencies, and approvals independently remain inside every
unchanged adaptation cap may it enter the ordinary Gate-0 lanes. A successful
spike is therefore neither admission nor a winner. Otherwise record a no-winner
Round 5 and stop product work.

This round does not:

- relax, delete, reweight, or reinterpret a master requirement, mandatory gate,
  score anchor, threshold, tie-break, resource cap, native lane, or packaging rule;
- give WezTerm incumbent or migration-cost credit;
- authorize a second VT/Kitty parser, a candidate-owned control/renderer, an
  EXE/Common terminal runtime, a floating Rust dependency, or a runtime outside
  `Plugins\Terminal.dll` and its private locked closure;
- reuse or edit an immutable Round-4 universe/review/assessment/set/decision
  record, fabricate collection/finalization evidence for a zero cohort, or
  modify the archived Round-4 Done plan; or
- start product plans 2–6 before the sole successful handoff is published and
  independently verified.

## Immutable Round-5 record set

Create the carry audit and Round-5 decision in every closeout. Create the v5
governance/proof records only after an approved `Unchanged` carry audit, and the
supersession record only on that branch when applicable. Never overwrite the
Round-4 subjects:

- `Specs/Terminal/TerminalEngineRound5CarryAudit.v1.json`
- `Specs/Terminal/TerminalEngineCandidateUniverse.v5.json`
- `Specs/Terminal/TerminalEngineRound5Proof.v1.json` after proof admission
- `Specs/Terminal/TerminalEngineCandidateReview.v5.json`
- `Specs/Terminal/TerminalEngineCandidateAssessment.v5.json`
- `Specs/Terminal/TerminalEngineCandidateSet.v5.json`
- `Specs/Terminal/TerminalEngineCandidateGovernanceFreeze.v5.json` as the
  authoritative journal binding the reserved Review/Assessment/Set bytes and
  their common logical freeze
- `Specs/Terminal/TerminalEngineRound5ExecutionFreezeAttestation.v1.json` after
  all A-bound collection/first-selection work stops whenever boundary A existed;
  it embeds the ignored execution-freeze approvals for clone-durable
  verification
- one or more cumulative, immutable
  `Specs/Terminal/TerminalEngineRound5SupersessionObservationLedger.gNNNN.v1.json`
  records after v5 work stops, including the empty/rejected/approved live
  observation history. Generation `g0001` has no predecessor. A pre-E
  observation admitted after an ordinary post-effect evidence commit creates
  exactly the next generation, which embeds and authenticates its predecessor;
  neither generation is overwritten or discovered by enumeration
- `Specs/Terminal/TerminalEngineRound5Supersession.v1.json` only as the
  post-fence/post-quiet result when an approved evidence delta invalidates this
  universe
- `Specs/Terminal/TerminalEngineRound5QualificationFailure.v1.json` only when
  Finalize names a winner but the unchanged first-selection sequence cannot
  complete
- `Specs/Terminal/TerminalEngineRound5FirstSelectionAttemptLedger.v1.json` only
  after at least one ignored live admission receipt exists and the active
  first-selection action has stopped at its admission fence; it is
  branch-neutral history and is never added to the clean repository used by
  candidate lanes
- `Specs/Terminal/TerminalEngineRound5WinnerHandoffProposal.v1.json` only after
  every non-decision first-selection qualification succeeds; it is the
  immutable approval subject, never the final Round-5 decision
- `Specs/Terminal/TerminalEngineRound5WinnerHandoffReview.1.v1.json` for the
  first proposal review, and
  `Specs/Terminal/TerminalEngineRound5WinnerHandoffReview.2.v1.json` only after
  an approving first review
- `Specs/Terminal/TerminalEngineDecision.Round5.md`

`Specs/Terminal/TerminalEngineDecision.md` remains the authoritative current
operator summary. Update it in the reviewed Round-5 result-freeze/closeout change
before this plan moves so it points to the immutable Round-5 decision path/digest
and retains the complete dependency lifecycle runbook; the later post-move
handoff publication verifies that identity and changes no Round-5 result. On the
evaluated branch, the evaluator/provider must accept the four exact versioned
governance subject paths explicitly and bind their raw SHA-256 identities; it
also receives the explicit GovernanceFreeze path/raw digest that authenticates
the atomic Review/Assessment/Set bundle. The proof record and approved carry-
audit/supersession identities are separate required inputs. It may not discover
a highest version or silently fall back to the Round-4 `.v1` files.

The v5 universe names the approved v4 universe as predecessor and changes only
the WezTerm proof hypothesis/input identities needed to test this design. It is
not permission to refresh, repin, or reconsider any other candidate inside
Round 5. Before v5 freezes, reverify every carried Ghostty, TerminalCore,
Contour, libvterm, libtsm, seed, and constraint-control pin, archive, patch,
result-tree, evidence digest, and recorded outcome byte-for-byte against Round 4;
carry them unchanged. Before any v5 universe exists, materialize the closed
`TerminalEngineRound5CarryAudit.v1.json` record. It contains the exact sorted
expected/observed path, SHA-256, size, pin, archive/tree, and recorded-outcome
identities for every carried item; a closed
`Unchanged|TransitionRejected` state; and, for every mismatch or viability
claim, a content-free reproducer/evidence digest and reason code. Two
independent non-author reviewers approve its exact bytes. `Unchanged` is the
only state that permits v5 to freeze. `TransitionRejected` closes Round 5
through the dedicated pre-universe branch below: Universe, Review, Assessment,
Set, and proof are all explicitly N/A, and no provider/evaluator action may
run.

The carry writer sorts by `{kind,id}` and derives—not accepts—its state.
`Unchanged` requires exact equality for every expected/observed
`{kind,id,path,pin,size,archiveSha256,treeSha256,evidenceSha256,
recordedOutcomeSha256}` field, no missing/extra/duplicate row, and no viability
assertion. Any mismatch or assertion derives `TransitionRejected` and one
same-key delta containing the exact expected/observed projection, closed reason,
reproducer SHA-256, and evidence SHA-256. The approved completion
`round5TransitionEvidence` is the RS-JCS identity of that exact sorted delta set;
it cannot be independently authored, omit a mismatch, or name a row absent from
the carry audit.

After v5 freezes and before the Round-5 E decision cutoff, any newly
discovered, independently reproducible non-WezTerm identity or viability change
uses a two-phase protocol. First, the coordinator pre-renders an ignored
immutable supersession proposal below
`.build/terminal-engine-spike/round5-supersession/<observation-id>/proposal.v1.json`.
It binds observation ID, proposal author, v5 universe digest, exact expected/
observed identities, content-free evidence digest/reproducer, closed reason
code, and discovery UTC—never a future milestone, cancellation, revert, result,
or decision identity. Under the coordinator lock, atomic no-replace proposal
publication is the authoritative state-transition journal and races the next
Pipeline or ControlQuiet admission CAS. Before the rename, recovery remains in the old epoch;
after it, state is derived as `ObservationPending` even if the process crashed
before updating memory. Thus exactly one wins: either a pipeline action was
already admitted and is recorded by the proposal transition, or proposal
publication prevents the next admission. In `ObservationPending`, the one
already admitted pipeline action may reach its next commit/quiet gate, but no
new ordinary Round-5 action or writer may start. The only admitted operations
are that active action's commit/quiet gate, supersession review 1 and (after
review 1 approves) review 2, rejected-observation closure/epoch resume, and the
bounded decision-generation terminalizer when the proposal interrupted an
already-published decision subject. That terminalizer may publish only the
subject's exact generation evidence/disposition and, when PrepareE already
exists, its linked `stale-before-admission` marker described below. Every output
binds the same winning proposal transition and already-published subject/E
attempt. The terminalizer requires zero Pipeline actions and cannot publish
another audit, approved decision, E preparation, candidate result, or new
generation.

Proposal publication is mutually exclusive with every ControlQuiet token
admission under the same coordinator lock. Its CAS predicate requires
`activeControlAction=N/A`. If a ControlQuiet token wins first, proposal
publication waits while that writer publishes its one reserved terminal output
and closes; it cannot interrupt the token-to-output interval. If the proposal
wins first, every new ControlQuiet token is blocked except the exact
already-published-generation terminalizer above. Therefore all observation cut
points are between completed ControlQuiet outputs, never after a token but
before its output, and no allocated-but-unpublished audit, evidence,
approved-decision, or PrepareE attempt needs an invented archive record.

Two ignored immutable ordered review records bind the proposal raw digest.
Reviewers are distinct from the proposal author and each other; review UTC is
strictly ordered. A first `Reject`, or `Approve,Reject`, closes that observation
as rejected. After the coordinator proves zero active action it increments the
epoch and may resume the unchanged v5 round; it never changes frozen inputs.
`Approve,Approve` is the sole credible-new-evidence state. The review-2 writer
first renders/validates exact immutable bytes in a private writer-owned
temporary. Review 2 binds a coordinator-allocated transition ID and exact old/
new epochs, but never the future fence digest. The coordinator pre-renders the
fence with that transition ID, epochs, reserved review-2 raw digest, active
pipeline action if any, and fence UTC. Under the lock, the single authoritative
CAS is an atomic no-replace publication of `fence.v1.json`; persistent epoch
state is derived from that journal record, not updated as a separate transaction.
Before the fence rename, recovery remains `ObservationPending`. After it,
recovery is unconditionally `Superseding`, rejects all old-epoch admissions, and
may publish only the already reserved review-2 bytes with the digest bound by
the fence. A crash between fence and review publication therefore resumes only
that writer; a crash afterward continues cancellation/quiet. Neither file is
overwritten, and review 2 cannot introduce a fence/review hash cycle.

The coordinator serializes all official local and remote pipeline work, so the
fence has at most one active pipeline action. Supersession proposal/review/fence
operations use a distinct serialized control-plane token class: they run under
the same lock, can run in `ObservationPending` while the pipeline slot remains
occupied, cannot invoke a candidate process or mutate candidate/product output,
and never count as a second pipeline action. It requests that action's normal
bounded cancellation or completion, rejects any stale result at the post-return
epoch gate, and emits a pre-cleanup quiet receipt with zero active pipeline
actions. Only cancellation/quiet, an exact
adoption cleanup, the branch-neutral attempt ledger, the closed observation
ledger, final supersession result, and no-winner decision are admitted after
`Superseding`. Every completed content-free artifact remains immutable and is
later labeled invalidated; no next lane, retry, ordinary proposal/review,
qualification writer, or paused-writer retry can publish. An unapproved observation therefore
halts at an executable admission fence rather than silently vetoing or permitting
result publication.

After pre-cleanup quiet, create the next tracked branch-neutral
`TerminalEngineRound5SupersessionObservationLedger.gNNNN.v1.json`, embedding
every proposal/review canonical object and raw digest seen in this round in
observation-ID order, including rejected observations, plus the exact prior
generation path/raw digest/canonical object or
`N/A-first-observation-ledger-generation`. The coordinator allocates the next
four-digit generation explicitly and passes it to every reader; no tool searches
for a highest generation. Then perform any required adoption cleanup and create
the tracked immutable
`TerminalEngineRound5Supersession.v1.json` result. The result binds the approved
triggering observation-ledger generation/entry, fence, pre-cleanup quiet,
cleanup-action, and post-cleanup quiet identities, derived
timing state, exact
milestones/artifacts visible at the fence, cleanup identities, and their
invalidated/absent dispositions. It never triggers its own prerequisite and
contains no later decision identity. The timing is derived, never authored, as
`BeforeProof`, `DuringProof`, `AfterProofBeforeGovernanceFreeze`,
`AfterGovernanceBeforeCollection`, `DuringCollection`,
`AfterCollectionBeforeAdoption`, `DuringAdoption`, or
`AfterAdoptionBeforeCloseout`.

If cleanup is `NotStarted`, bind
`cleanupAction=N/A-no-cleanup-action` and
`postCleanupQuiet=N/A-no-cleanup-action`; pre-cleanup quiet is the effective
final quiet. Otherwise the terminal epoch admits exactly one state-specific
`TransactionalCleanup` Pipeline action after pre-cleanup quiet. This is a
resumable deterministic repository transaction, not a candidate/effect replay
and not the generic absent-output writer retry. Its C Git-boundary specialization
and cleanup front end are the same single Pipeline token, claim, logical key,
manifest, and optional commit—not two actions. Before touching the real index,
worktree, object database, or ref, its sole writer creates an immutable
create-new transaction manifest below
`.build/terminal-engine-spike/round5-git-boundary/C/<transaction-id>/
manifest.v1.json`. The cleanup schema is a subtype of that one canonical
manifest; no second cleanup path/file is created. The manifest
binds branch/result, cleanup state, logical action key, original token/claim,
epoch/input identity, exact symbolic branch/ref, expected HEAD/ref parent,
boundary A and any B adoption commit/tree, the closed adoption path set, every
before and after path/mode/raw digest, the exact permitted governance delta,
private-index identity, intended tree, fixed commit message/author/committer/
UTC, and the deterministic cleanup commit object ID or the exact
`N/A-no-cleanup-commit`.

For `CancelledBeforeCommit`, recovery restores the exact A projection in both
the real index and worktree and proves the expected ref never moved; no cleanup
commit is created. For `CommittedThenReverted|AdoptedThenReverted`, the writer
builds the intended tree in its private index, creates at most the one manifest-
predicted commit object, and moves only the expected branch ref with an atomic
compare-and-swap from the manifest parent to that commit. It then synchronizes
the checked-out real index/worktree to the committed tree and verifies the A
projection and zero engine state. Persistent recovery derives
`Prepared|CommitObjectReady|RefAdvanced|ProjectionRestored|Verified` from the
immutable manifest, object database, exact ref, index, and worktree; it never
trusts a mutable phase flag. A ref other than the exact parent or intended
commit, a third path state, changed metadata/tree, unrelated tracked delta, or
wrong branch is STOP.

A crash after the original claim may acquire only a linked cleanup-recovery
token for the same logical action and manifest. It resumes the missing
repository-transaction boundary without rerunning an engine action, changing
inputs/metadata, allocating another commit, or reserving another original
logical key. After `Verified`, the coordinator emits a distinct post-cleanup
quiet receipt. That receipt is exactly C's generic `postTransactionQuiet`;
`postCleanupQuiet == C.postTransactionQuiet`, and a second receipt is invalid.
No Qualification/Supersession/D `ControlQuiet` writer may run until it exists.

Every four-digit pre-E observation ID, observation-ledger generation,
decision-subject generation, and per-transaction recovery-attempt ordinal is
create-new and monotonic in its own independent closed domain; the same rule
applies to PostFCleanup attempt ordinals. Four digits are an encoding bound, not
an entitlement to 9,999 archive-destined objects: the mandatory-payload table
  caps admitted observation IDs at `0008`, cumulative observation-ledger
  generations at `0009`, ordinary decision generations at `0009`, and pre-E Git
  recovery records per transaction at `0003`. Decision generations are one
  contiguous domain: a ProtocolClosure-bound decision uses the next unallocated
  ordinal, including `0001` when closure precedes any prior decision. An
  administrative SuccessorD never skips directly to `0010`. Generation `0010`
  is reserved exclusively for the final ProtocolClosure decision only when
  `0001` through `0009` already have complete non-Accepted dispositions; no
  ordinary decision/audit may allocate it. The
  next otherwise legal observation or ordinary decision generation takes the
  reserved `CapacityClosing`/`ProtocolClosed` route before allocating a token or
  path. Rejection/failure of the reserved `0010` subject/audits is explicit
  administrative STOP/WIP, never generation `0011`.
Every Git boundary reserves `0001` through `0003` at admission; needing another
while its ref state is indeterminate is an administrative STOP/WIP because a
candidate outcome cannot be inferred safely. PostFCleanup is post-F,
non-archived, and may use `0001` through `9999`; exhaustion leaves only cleanup
`UnsafeStopped` and never invalidates ClosedF. No counter wraps, widens, reuses,
gaps, or silently opens a fifth digit.

After that fixed result freezes, the coordinator enters terminal
`SupersededCarryForward`: pipeline admission and result replacement stay
permanently closed, but new pre-E observation proposal/reviews and a cumulative
ledger successor may run with control-plane tokens. A rejected observation only
extends the ledger. An approved observation also extends the ledger and is
carried to the next ordinal universe; it does not replace this result. E closes
the control plane and binds the final cumulative generation.

If adoption is active at the fence, `DuringAdoption` has exactly
`CancelledBeforeCommit` (no adoption commit; prove the adoption-path projection
still equals the pre-adoption tree) or `CommittedThenReverted` (allow the atomic
commit to finish, then revert). If adoption had already committed, use
`AdoptedThenReverted`. Every revert is non-history-rewriting, stages/commits only
the exact adoption path set, restores that path projection byte-for-byte, and
binds the adoption/revert commits/trees plus zero remaining engine-adoption
state. When no post-effect governance commit intervened, closeout files remain
untracked/unstaged and the full revert tree must equal the pre-adoption commit
tree. If a late pre-closeout supersession arrives after an immutable governance
commit, never delete that evidence: the full tree may differ only by the closed,
digest-bound governance allowlist recorded by the Supersession result, while the
adoption-path projection must still exactly equal the pre-adoption tree. A
created but never adopted lock remains invalidated evidence; it is never called
adopted.

If a qualification-failure record was already frozen, approved supersession
before Done closeout still owns the branch and binds that record as invalidated.
If receipt/ledger/proposal/review validation failed but supersession acquired its
fence first, the supersession result embeds the same canonical content-free
`invalidArtifactEvidence` object and exact absent-publication substitute plus
canonical attempted inputs; the tracked target remains absent, and the result
never loses or fabricates the pending failure.

The E decision cutoff is the final Round-5 supersession fence. Its verifier
requires no pending unapproved observation and no later approved supersession
record before the commit. Once its `ClosingPublish` CAS succeeds, observation
admission remains continuously sealed through mandatory E -> FArchive -> FDone
-> FPublish -> FSeal recovery; neither an ignored proposal nor a new coordinator epoch
may interpose. The WIP-to-Done move, completed-plan bytes, and success marker
when applicable appear only in non-publishable FDone; the master pointer and
routing then bind that known FD commit/blob only in FPublish, and the tracked
FSeal makes FP quiet clone-durable. E, FA, FD, and FP are non-publishable.
Evidence discovered after the E CAS is queued without Round-5 authority and is
admitted only after `ClosedF` by the separately approved post-publication
engine-reselection/revocation process. It never edits this round's decision or
Done bytes and never blocks deterministic closeout recovery. For every
pre-closeout supersession, create the next ordinal
universe to freeze the broader delta before any new result. Never inject that
delta into v5 after observing it. The
universe authorizes only the frozen WezTerm
hypothesis; after `ProofSucceeded`, the later Review/
Assessment/Set replace only the prior WezTerm nomination with the completed v5
design. A viability change discovered after v5 freezes requires a later
universe; it cannot be inserted after seeing results.

`ClosedF` is the permanent Gate-0 temporal cutoff. If Round 5 closes no-winner,
queued evidence is an explicit input to the already-routed next ordinal round.
If Round 5 closes winner, no later process appends another Gate-0 success marker
or rewrites this Done plan: a separately approved emergency dependency
revocation first blocks product-plan admission/ships the feature disabled, then
uses the locked engine upgrade/replacement and rollback workflow with a new
review record. The original marker remains the historical handoff proof. A
replacement that cannot satisfy that lifecycle is a new top-level migration,
not a second winner smuggled into this Gate-0 lineage.
Every retained requirement, G1–G12 gate, C7 proof, fixture, measurement protocol,
weight, 80.00 threshold, inclusive five-point cohort, lexicographic tie-break,
native x64 Release/x64 `ASan Debug`/ARM64 Release collection lane, five winner
verification lanes, and two-reviewer rule is byte-identity-bound or explicitly
reapproved unchanged before any candidate result is observed.

## Required implementation and build surface

Round 5 is executable only after the following exact surface exists and its
candidate-neutral pieces are corpus-tested. Reuse a reviewed version-generic
schema only if it can represent these states without optional-field ambiguity;
otherwise create the named v5 successor. Every reader receives exact paths and
digests. No tool may infer a highest version, enumerate `*.v5.json`, or fall back
to a v1/v4 default.

Create these Round-5 subjects:

- `Specs/Terminal/TerminalEngineRound5CarryAudit.schema.json` and
  `Specs/Terminal/TerminalEngineRound5CarryAuditCorpus.v1.json`, covering both
  `Unchanged` and `TransitionRejected`, exact expected/observed identities,
  mismatch/viability reason codes, content-free evidence, two distinct
  non-author approvals, and the prohibition on v5 subjects in the rejected
  branch;
- `Specs/Terminal/TerminalEngineRound5MandatoryPayloadBudget.schema.json`,
  `Specs/Terminal/TerminalEngineRound5MandatoryPayloadBudgetCorpus.v1.json`,
  checked-in
  `Specs/Terminal/TerminalEngineRound5MandatoryPayloadBudget.v1.json`,
  `Specs/Terminal/TerminalEngineRound5ProtocolClosure.schema.json`, and
  `Specs/Terminal/TerminalEngineRound5ProtocolClosureCorpus.v1.json`, plus
  `Specs/Terminal/TerminalEngineRound5ProtocolRepair.schema.json` and
  `Specs/Terminal/TerminalEngineRound5ProtocolRepairApproval.schema.json` plus
  `Specs/Terminal/TerminalEngineRound5ProtocolRepairCorpus.v1.json`. The
  checked-in budget table fixes the RecoveryArchive hard envelope at 256
  embedded canonical objects, 65,536 canonical bytes per object, and 1,048,576
  aggregate mandatory canonical bytes. It reserves 64 objects and 262,144 bytes
  exclusively for the worst legal terminal path from every quiet coordinator
  state; ordinary accumulated history may consume at most 192 objects and
  786,432 bytes. The table enumerates every history-growing action, its maximum
  object count/canonical-byte contribution, and the maximum remaining
  branch-specific terminal route. A pure exhaustive graph check must prove that
  the largest route—including stop/quiet, cleanup C when required,
  ProtocolClosure, the maximum ProtocolRepair subject/two approvals/wrapper,
  D/SuccessorD, both final validation receipts/quiets, the complete decision-
  generation evidence/audit chain, every still-possible PrepareE attempt and
  its maximum recovery/stale-or-accepted terminal evidence, E, and the archive
  wrapper—fits the reserve;
  a hand-entered assertion is invalid.
  RS-JCS accounting is recursive and charges the exact bytes FArchive will
  embed, including property names, arrays, commas, and wrappers once; nesting an
  object never hides it from the object count. The same canonical object/raw
  digest referenced twice is charged once only when the archive schema requires
  one shared embedded instance and both references resolve to it.
  Before any result, the same table freezes the complete exact
  `protocolRepairAllowlist` of candidate-neutral closeout tool/test paths with
  path, mode, pre-result digest, and semantic class. It excludes every
  candidate/product/runtime/adapter/lock/dependency/score/fixture/input,
  requirement/gate/expected-outcome, evidence, schema, corpus, prior decision,
  and prior result byte.
  The closed outer cardinalities are one P candidate, one A candidate, at most
  eight admitted pre-E observations, at most nine cumulative observation-ledger
  generations (an optional initial Empty plus those eight), at most nine SuccessorD boundaries
  (eight observation-driven plus one protocol-closing boundary), at most nine
  ordinary decision generations plus generation `0010` reserved exclusively
  for ProtocolClosure, at most nine ordered PrepareE attempts (at most eight
  ending `StaleBeforeAdmission` and exactly one final `Accepted` when E is
  reached), exactly one archive-embedded DecisionSubject token for each
  allocated decision generation, at most
  two linked pure-writer attempts, and three
  pre-E Git-recovery records per transaction. These maxima are safety ceilings,
  not admission authority.
  Before no-replace publication of every history-growing token, receipt,
  proposal/review/fence/ledger, boundary/recovery record, decision/audit, or
  other archive-destined object, the coordinator lock runs an exact prospective
  render and requires
  `used + candidate + maximumTerminalRoute(stateAfterCandidate)` to fit all
  three hard limits while ordinary history stays within its partition. The
  token/manifest binds the budget-table raw digest and the pre-admission budget
  snapshot; the publication CAS recomputes both. A race either publishes a
  fully budgeted object first or loses without allocating a token/path.
  When an otherwise legal next action or a ninth observation would not fit, the
  coordinator publishes no part of it and CASes to `CapacityClosing` using one
  reserved content-free `ProtocolClosure` record with reason
  `evidence-budget-exceeded`. A failed `PreProofFocused` or
  `PreResultFocused` receipt similarly CASes to `ValidationClosing` with reason
  `preproof-focused-failed` or `preresult-focused-failed`; same-round P/A repair
  and a second P/A candidate are forbidden because they would mutate or
  ambiguously reapprove the observed universe. A failed ordinary
  `FinalFocused|FinalCanonicalFull` CASes once to `ValidationClosing` with reason
  `final-validation-failed`, invalidates any winner claim, and permits exactly
  one reserved administrative SuccessorD that may repair/revert only the
  pre-frozen candidate-neutral closeout tool/test allowlist and must run both
  final suites anew. Any changed byte requires two non-author approvals over the
  canonical tracked
  `Specs/Terminal/TerminalEngineRound5ProtocolRepair.v1.json`, whose sole
  ignored subject/approval/final staging paths are respectively
  `.build/terminal-engine-spike/round5-protocol-closure/protocol-repair-subject.v1.json`,
  `.build/terminal-engine-spike/round5-protocol-closure/protocol-repair-approval-1.v1.json`,
  `.build/terminal-engine-spike/round5-protocol-closure/protocol-repair-approval-2.v1.json`,
  and
  `.build/terminal-engine-spike/round5-protocol-closure/protocol-repair.v1.json`.
  The subject lists the exact sorted before/after path, mode, SHA-256, closed defect
  reason, and old/new tool identities consumed by the rerun, and proves frozen
  requirements/gates/expected outcomes, all prior evidence bytes, and the
  excluded path set unchanged. Approval 1/2 bind the subject raw digest,
  already-final ProtocolClosure raw digest, the failed ordinary
  `FinalFocused|FinalCanonicalFull` receipt/quiet,
  budget table, allowlist, before/after tree projections, and
  intended administrative-SuccessorD parent. No ProtocolClosure, capacity
  closure, P/A validation closure, a closure before ordinary D, and any
  already-administrative D require exact
  `N/A-no-protocol-repair-permitted`; they cannot create a repair subject
  even if their closeout tooling later fails. The repair author cannot review; both
  reviewers are distinct non-authors; both verdicts are `Approve`; UTC is
  strictly subject < approval 1 < approval 2. The sole finalizer embeds the
  canonical subject and both approval objects/raw digests into
  `protocol-repair.v1.json`; no approval or wrapper approves itself. Every path
  is create-new/no-replace, and absent-output recovery may republish only the
  byte-identical reserved object under its original ControlQuiet key. The
  administrative SuccessorD adds that byte-identical approved wrapper with the
  repaired allowlisted bytes; it cannot
  alter a test expectation, schema/corpus, threshold, fixture, or observed
  result. An eligible final-validation closure needing no byte change uses exact
  `N/A-no-protocol-repair-required`; every ineligible closure uses the distinct
  exact `N/A-no-protocol-repair-permitted`. The terminal record binds the stage, attempted logical-key
  projection, exact would-be count/byte calculation or failed suite
  receipt/quiet, each applicable P/A candidate transaction plus suite
  receipt/quiet and exact zero/one/two approval canonical objects/raw digests
  with `unapproved-invalidated|approved-invalidated|approved-retained`
  disposition, stop/pre-cleanup quiet, state-specific zero-engine cleanup,
  invalidated partial artifacts, and the final budget snapshot; it deliberately
  contains no ProtocolRepair or repair-approval identity and never
  invents a candidate verdict. The canonical
  record is one RS-JCS UTF-8/no-BOM/no-trailing-newline value staged at the
  explicit ignored
  `.build/terminal-engine-spike/round5-protocol-closure/closure.v1.json` and
  added without byte change by the protocol D/SuccessorD at the fixed tracked
  `Specs/Terminal/TerminalEngineRound5ProtocolClosure.v1.json`; both exact paths
  and its raw digest are token/manifest inputs and no reader enumerates them.
  Its `ProtocolClosurePublish` ControlQuiet logical key is exactly
  `round5/control/ProtocolClosurePublish/1/Specs/Terminal/TerminalEngineRound5ProtocolClosure.v1.json`.
  BoundaryEvidence embeds the same canonical object/raw digest only when the
  closure predates its D; a later closure leaves the already-immutable
  BoundaryEvidence's exact N/A unchanged and is embedded by the successor D
  manifest, final decision, and RecoveryArchive. ProtocolClosure and
  ProtocolRepair are digest-acyclic siblings: the repair subject/approvals bind
  the already-final closure raw digest, while closure never binds repair. The
  decision and archive embed both plus the repair approvals. Before a terminal branch
  exists, or when a
  prospective Winner is invalidated, its branch is `ProtocolClosed`, outcome
  `no-winner`. If `TransitionRejected`, Superseded, Evaluated/No-winner, or
  Qualification-failed is already immutable, the record instead has
  `disposition=PreserveTerminalBranch`, seals later admission, and is an
  auxiliary input to a successor D/decision without replacing that branch or
  result.
  A second failure on that administrative
  D—including a protocol D created first for capacity closure—is an explicit
  verification STOP/WIP requiring reviewed protocol recovery,
  never another automatic history item or a false closeout. Recovery-attempt
  exhaustion while a Git transaction is indeterminate is likewise STOP/WIP;
  every boundary admission reserves all three recovery records up front.
  Closure races use one fixed precedence table. A proposal already published as
  `ObservationPending` resolves first: Approve selects/preserves Superseded and
  embeds any terminal failed suite as invalidated; Reject publishes the next
  cumulative ledger and only then retries the closure CAS against that
  generation. A `Superseding` fence always wins. If
  `CapacityClosing|ValidationClosing` wins before proposal publication, the
  proposal remains absent and its evidence is queued for the successor after
  ClosedF. An already-admitted suite may always publish its reserved
  receipt/quiet; if it fails while state is `CapacityClosing`, the state upgrades
  before ProtocolClosure publication to `ValidationClosing`, with strict reason
  precedence `preproof-focused-failed > preresult-focused-failed >
  final-validation-failed > evidence-budget-exceeded`. Once the canonical
  ProtocolClosure stages, its reason/disposition is immutable.
  `PrepareE|ClosingPublish` cannot coexist with an active/pending suite or
  closure state.
  Corpora cover every growth action, object/byte/per-object minus-one/equal/
  plus-one boundary, the statically largest route, racing observation versus
  capacity and validation CAS in both orders, branch-preserving closure, all
  three validation reasons, reason precedence, forbidden second P/A candidate,
  ProtocolRepair allowlist/path/mode/digest/approval/invariant mutations,
  repair-present on an ineligible stage and either stage-specific N/A value
  substituted for the other,
  administrative-D one-shot behavior, and recovery-reserve exhaustion;
- `Specs/Terminal/TerminalEngineRound5SupersessionProposal.schema.json`,
  `Specs/Terminal/TerminalEngineRound5SupersessionReview.schema.json`,
  `Specs/Terminal/TerminalEngineRound5SupersessionObservationLedger.schema.json`,
  `Specs/Terminal/TerminalEngineRound5Supersession.schema.json`, and
  `Specs/Terminal/TerminalEngineRound5SupersessionCorpus.v1.json`, covering the
  exact post-v5 proposal/delta evidence, digest-domain separation, strict
  proposal/review chronology, author/reviewer separation, rejection/resume and
  proposal-publication-vs-pipeline-admission CAS, recovery on both sides of that
  authoritative journal, two-approval journaled fence transitions, every crash boundary before/after
  fence and reserved-review publication, all coordinator epoch/action/quiet races, the
  eight mechanically derived pre-closeout timing states, exact milestone
  identities, and the prohibition on changing frozen subjects. The observation
  ledger copies every ignored proposal/review canonical object and raw digest,
  contains no result identity, and is empty only when no observation occurred.
  It requires an explicit four-digit generation, generation 1's exact N/A
  predecessor or the immediately prior generation's path/raw digest/embedded
  canonical object, and a strict cumulative-prefix match; it rejects a gap,
  fork, overwrite, implicit-highest lookup, omitted prior observation, or
  changed prior bytes.
  The final result requires one approved observation entry, fence/pre-cleanup-
  quiet/cleanup-action/post-cleanup-quiet identities, exact completed/aborted
  artifact snapshot, and all timing-derived
  invalidated/N/A dispositions. Whenever P existed it embeds P commit/tree,
  both canonical approval objects/raw digests and P epoch. `BeforeProof` after
  approved P binds exactly `N/A-proof-action-not-admitted`; `DuringProof` or a
  later state binds the proof action token/claim/terminal-or-partial
  disposition. Exactly
  `N/A-superseded-before-proof-freeze` is allowed only before approved P. When
  supersession interrupts attempted P or A creation/approval, it embeds that
  commit/tree and every zero/one/invalid approval identity as invalidated while
  completion uses the before-boundary N/A. For `DuringAdoption` it distinguishes
  `CancelledBeforeCommit|CommittedThenReverted`; after an adoption commit it
  requires `AdoptedThenReverted`, exact adoption/revert commits/trees and path
  projections, plus zero remaining tracked engine-adoption paths. It binds an
  already frozen qualification-failure record as invalidated, or proves none
  existed at its fence. It derives the exact attempt-ledger state plus winner-
  handoff proposal and zero/one/two review identities from actions admitted
  before the fence and completed, aborted, or quieted before the final result.
  When admission stopped after one rejected receipt, it requires
  `StoppedAfterFirstRejection`. A receipt/ledger/proposal/review validation
  failure instead carries the same canonical `invalidArtifactEvidence` object
  and exact `N/A-validation-rejected-before-publication` plus canonical attempted
  inputs while the tracked target remains absent.
  No final supersession result or observation ledger contains a later decision
  identity or is an approval subject for itself. In
  `SupersededCarryForward`, a later pre-E observation is reviewed and appended
  in the next cumulative ledger generation; Reject preserves the terminal
  Superseded branch and Approve carries the additional delta to the next ordinal
  universe. Neither case replaces the already-terminal result. The completion
  and decision bind the final cumulative ledger generation;
- `Specs/Terminal/TerminalEngineRound5CoordinatorArtifacts.schema.json` and
  `Specs/Terminal/TerminalEngineRound5CoordinatorArtifactsCorpus.v1.json`,
  covering create-new bootstrap governance state, reviewed pre-A P epoch,
  quiet P-to-A rotation, A-bound effect epoch, action token, five-runner preflight, action
  commit/rejection, observation-pending, supersession fence, bounded cancel,
  quiet receipt, epoch-resume/close, one-global-pipeline-action serialization,
  distinct serialized no-effect control-plane tokens, remote
  transfer identity, canonical logical-action key/cardinality/allowed-next
  journal, prospective mandatory-payload budget snapshot/recheck,
  `CapacityClosing|ValidationClosing|ProtocolClosed`, fresh-token duplicate
  rejection, stale/bypass rejection,
  candidate-action crash terminality,
  coordinator-synthesized terminal proof/collection/dependency-status records and
  failed-preflight closure after a claimed
  preflight crash/timeout, linked pure-writer retries, and every writer-pause
  versus supersession race, approval/cancel while a pipeline action hangs,
  `SupersededCarryForward`, observation-vs-`ClosingPublish` CAS,
  fixed-decision-absent-before-CAS enforcement, and byte-identical crash
  recovery through a before/after path manifest, intended Git tree/metadata,
  and commit-derived `ClosedE`. Epoch state and live tokens are ignored
  `.build` artifacts; tracked result records embed every canonical token object
  and raw digest needed for clone-durable verification;
- `Specs/Terminal/TerminalEngineRound5CleanupTransaction.schema.json` and
  `Specs/Terminal/TerminalEngineRound5CleanupTransactionCorpus.v1.json`.
  They specialize the generic Git-boundary transaction for the sole
  `TransactionalCleanup` Pipeline subtype, immutable
  before/after path/index/tree manifest, expected branch/ref/parent, fixed commit
  metadata/object ID, no-commit `CancelledBeforeCommit` restoration, commit-
  object/ref-CAS recovery, linked cleanup-recovery token, and mechanically
  derived `Prepared|CommitObjectReady|RefAdvanced|ProjectionRestored|Verified`
  states. The corpus rejects a second original claim/logical key, changed
  recovery inputs, non-CAS ref movement, unrelated tracked delta, third path/
  index state, a second commit, or post-cleanup governance before the distinct
  quiet receipt;
- `Specs/Terminal/TerminalEngineRound5GitBoundaryTransaction.schema.json` and
  `Specs/Terminal/TerminalEngineRound5GitBoundaryTransactionCorpus.v1.json`.
  They cover
  `P|A|B|C|D|SuccessorD|E|FArchive|FDone|FPublish|FSeal`, each as one
  resumable deterministic Git
  transaction with a caller-supplied `transaction-<32-lowercase-hex>` ID and
  explicit create-new
  `.build/terminal-engine-spike/round5-git-boundary/<boundary-kind>/
  <transaction-id>/manifest.v1.json` path. Every manifest binds the immediately
  preceding boundary-manifest path/raw digest or the legal
  `N/A-first-round5-boundary`, never an enumerated latest. The predecessor-kind
  DAG is fixed: the sole P candidate is first; the sole A candidate follows the
  approved P; B follows the approved A; C follows B or a D/SuccessorD when late approved
  supersession or a final-validation ProtocolClosure invalidating a prospective
  Winner requires reverting an earlier B; an initial D follows the latest
  existing P/A/B/C, or is first only for TransitionRejected/supersession-before-
  P; SuccessorD follows D/SuccessorD or the late C; E follows final D/SuccessorD;
  FArchive follows E, FDone follows FArchive, FPublish follows FDone, and FSeal
  follows FPublish. E/FArchive/FDone/FPublish/FSeal can never use the N/A.
  Failed P/A candidate manifests remain immutable and route only to
  `ProtocolClosed`; a repaired/second P or A boundary is never a legal
  predecessor. E and F preparation attempts also
  carry a separate same-kind `priorAttemptManifest` path/raw digest or exact
  first-attempt N/A in addition to the boundary predecessor, so a stale attempt
  never replaces lineage. It also
  binds the fixed action class/kind, canonical logical key, epoch/input identity,
  original token canonical object/raw digest, and reserved recovery-attempt/
  quiet/publication-fence paths—but no future token, verdict, receipt, or digest.
  P/A/B/C/D/SuccessorD bind their admitted Pipeline token and exact runner claim/
  create-new root. E, FArchive, FDone, FPublish, and FSeal instead bind their
  effect-limited
  `PrepareE|PrepareFArchive|PrepareFDone|PrepareFPublish|PrepareFSeal`
  `ControlQuiet` preparation token and exact
  `N/A-control-operation-has-no-runner-claim` /
  `N/A-control-operation-has-no-candidate-output-root` values. Preparation may
  write only its ignored manifest/private index and predicted Git object; it has
  no lease/ref authority. `PrepareE` is adopted only by the create-new
  ClosingPublish fence; `PrepareFArchive` is adopted only by the create-new
  `<FArchive-transaction-directory>/archive-publish-journal.v1.json` plus ref
  CAS; `PrepareFDone` is adopted only by the create-new
  `<FDone-transaction-directory>/done-publish-journal.v1.json` plus ref CAS; and
  `PrepareFPublish` is adopted only by the create-new handoff fence plus ref CAS;
  `PrepareFSeal` is adopted only by create-new
  `<FSeal-transaction-directory>/seal-publish-journal.v1.json` after the
  verified FP quiet. That journal binds the prepared manifest path/raw digest,
  the distinct `SealPublish` ControlQuiet token/key, E/FA/FD/FP and all four
  embedded post-transaction quiet objects/raw digests; only then may the FP ->
  FS ref CAS add the tracked closeout seal.
  Each journal
  binds its prepared manifest path/raw digest and distinct publication
  `ControlQuiet` token. FArchive never binds or requires the handoff fence. A
  losing preparation receives immutable
  `stale-before-admission.v1.json` and can never update a ref. It also binds the
  symbolic branch/ref and
  expected old OID, exact index/worktree before projection, path/mode/digest
  before/after set, intended tree, fixed commit metadata and predicted commit
  OID and atomic `update-ref` compare-and-swap. A recovery uses an explicit
  create-new sibling
  `recovery-attempt-<four-digit-ordinal>.v1.json` that binds the manifest raw
  digest, immediately prior recovery attempt or exact first-attempt N/A, linked
  recovery token, and unchanged inputs; only `0001` through `0003` are legal,
  and the original admission budget/reserves all three exact paths and maximum
  objects/bytes before the manifest can publish. The original manifest is never
  edited.
  After real-index/worktree synchronization and verification, the coordinator
  no-replace creates exactly
  `<transaction-directory>/post-transaction-quiet.v1.json`, binding the manifest
  raw digest, final commit/ref/tree, derived terminal state, and
  `activeActionCount=0`. C's post-cleanup quiet is this same file. Persistent
  state is derived only
  as `Prepared|CommitObjectReady|RefAdvanced|Synchronized|Verified`; recovery
  distinguishes an object created while the ref still names the old OID from a
  moved ref whose coordinator state or checked-out files were not yet updated.
  The corpus crosses every crash edge and rejects alternate refs, parents,
  trees, metadata, worktree/index states, duplicate commit objects, or unrelated
  tracked mutation. B/C additionally prove the adoption projection and zero
  engine state; D/SuccessorD bind the exact branch and cumulative observation-
  ledger generation and serialize observation admission against the boundary.
  E is this same manifest type with the additional ClosingPublish path set and
  lease predicate; its create-new `closing-publish-fence.v1.json` is the sole
  authority to seal observation admission for the rest of closeout. E publishes
  the audited Round-5 decision/cutoff but leaves the WIP plan, master handoff
  tokens, success marker, and routing unchanged. FArchive is the mandatory
  archive-only commit from E to FA and adds only the tracked RecoveryArchive.
  FDone follows FA and only fills/moves this plan WIP-to-Done, including the
  branch completion marker and the successful-handoff marker only for Winner.
  FPublish follows FD, binds accepted E/decision/outcome, known FD commit/Done
  blob, master/README/next-round routing, FA commit plus RecoveryArchive path/raw
  digest, and proves the E decision/archive and FD Done bytes unchanged. FSeal
  follows verified FP quiet and adds only the canonical tracked closeout seal
  that embeds all four E/FA/FD/FP quiets and binds their commits.
  The coordinator retains the sealed ClosingPublish lease through verified FS;
  a crash can resume only the next deterministic manifest in that
  E -> FA -> FD -> FP -> FS chain;
- `Specs/Terminal/TerminalEngineRound5ExecutionFreezeAttestation.schema.json`
  and
  `Specs/Terminal/TerminalEngineRound5ExecutionFreezeAttestationCorpus.v1.json`.
	  The post-effect tracked record binds boundary-A commit/tree, its exact tracked
	  file-set identity and pre-result test identity, then embeds both ignored
	  approval canonical objects/raw digests and revalidates their embedded A
	  transaction manifest/post-transaction quiet plus PreResultFocused
	  receipt/suite quiet. It enforces two distinct non-author
  Approve verdicts strictly after A and before the first A-bound collection or
  first-selection action,
  contains no candidate result/branch/decision identity, and is explicitly N/A
  only when the terminal timing proves A never existed;
- `Specs/Terminal/TerminalEngineRound5BoundaryEvidence.schema.json` and
  `Specs/Terminal/TerminalEngineRound5BoundaryEvidenceCorpus.v1.json`. Before D,
  the sole writer creates
  `Specs/Terminal/TerminalEngineRound5BoundaryEvidence.v1.json`, embedding every
  applicable completed/failed P/A candidate, B, and C canonical transaction
  manifest, recovery-attempt records, and post-transaction quiet objects/raw
  digests plus every P/A validation-suite receipt and suite-quiet canonical
  object/raw digest and every zero/one/two P/A approval canonical object/raw
  digest with the exact approved/invalidated disposition, with exact
  branch/timing N/As. It proves a Pass is the sole
  approval authority, a Failed receipt routes only to the canonical
  `ProtocolClosure`, B/C chronology and cleanup projection, and no
  unrepresented boundary or validation attempt. It embeds the ProtocolClosure
  only when that object predates D; otherwise it carries the exact
  `N/A-protocol-closure-after-boundary-evidence`, and the later successor-D
  manifest/decision/archive own that immutable identity. It deliberately
  excludes D/E/F
  to avoid forward/self
  identity. The final decision separately embeds final D/SuccessorD
  manifest/quiet, while E is commit-derived and FArchive archives the ignored
  chain;
- `Specs/Terminal/TerminalEngineRound5ValidationSuiteReceipt.schema.json` and
  `Specs/Terminal/TerminalEngineRound5ValidationSuiteReceiptCorpus.v1.json`.
  The four fixed suite IDs are
  `PreProofFocused|PreResultFocused|FinalFocused|FinalCanonicalFull`. Each
  content-free receipt binds its suite logical key, exact fixed command/
  argument vector and tool/script digests, P-candidate/A-candidate/final-D
  commit/tree as applicable, runner identity, Pipeline token/claim, start/end
  UTC, terminal `Pass|Failed|ClassifiedPreExisting`, exit/timeout identity, and
  result/log digests. Only `FinalCanonicalFull` may use
  `ClassifiedPreExisting`, and only with this Round-5 plan's closed prior-
  baseline/reproduction/owner/no-Gate-0-impact contract below (also repeated by
  the master). A crash/timeout receives
  one terminal failed receipt and is never replayed against the same commit.
  An administrative-SuccessorD receipt binds either the approved ProtocolRepair
  wrapper/raw digest and exactly its complete `newToolIdentities` set, or the
  stage-correct exact `N/A-no-protocol-repair-required` or
  `N/A-no-protocol-repair-permitted`,
  the frozen allowlist identity, and the unchanged prior tool set. The first N/A
  is legal only for eligible final-validation closure; the second is legal only
  for an ineligible capacity/P/A/pre-D/no-closure class that nevertheless has a
  SuccessorD. Ordinary and old-D receipts remain bound to their
  original digests. Mixed old/new, an unapproved
  new digest, or a new identity outside the repair subject is Failed before test
  invocation.
  P/A approvals require their suite Pass; E requires FinalFocused Pass and
  FinalCanonicalFull Pass or the one classified form. The final Round-5
  decision embeds the applicable canonical receipts/raw digests so clone
  verification never depends on ignored runner output. Its quiet receipt is the
  fixed sibling `post-suite-quiet.v1.json`, binds the suite receipt raw digest/
  tested commit/terminal status and `activeActionCount=0`, and is no-replace;
- `Specs/Terminal/TerminalEngineRound5RecoveryArchive.schema.json` and
  `Specs/Terminal/TerminalEngineRound5RecoveryArchiveCorpus.v1.json`. FArchive
  creates exactly
  `Specs/TestRuns/TerminalEngineRecovery/round5-<YYYYMMDDTHHMMSSZ>-<16-lowerhex>/TerminalEngineRound5RecoveryArchive.v1.json`;
  the ASCII timestamp must round-trip through invariant
  `yyyyMMdd'T'HHmmss'Z'` UTC parsing, so an impossible date/time is invalid.
  That one-segment archive run ID and forward-slash spelling are canonical, so
  `.`, `..`, backslash, ADS, rooted/UNC, case aliases, a second segment, Unicode
  decimal-digit aliases, and normalization-on-behalf-of-caller are invalid. The
  archive is one RS-JCS canonical JSON value encoded as UTF-8 without BOM and
  without a trailing newline. Its raw bytes must not exceed 2 MiB. It is built
  from the complete predecessor Git-boundary chain ending at E. For every
  applicable P/A/B/C/D/SuccessorD it embeds the canonical manifest, exact
  ordered linked `recoveryAttempts` array (`[]` is the sole no-attempt form),
  and post-transaction quiet object/raw digests. Its mandatory ordered
  `eAttempts` array is exhaustive rather than a last-attempt shortcut. Each
  entry embeds the canonical PrepareE manifest/raw digest, its explicit
  prior-attempt path/raw digest or first-attempt N/A, its exact ordered linked
  recovery-attempt canonical objects, and the associated decision-generation
  evidence/disposition. A `StaleBeforeAdmission` entry additionally embeds the
  winning proposal transition and canonical `stale-before-admission` object/raw
  digest, and requires exact N/A for E commit and E quiet; it never fabricates a
  post-transaction quiet. The one final `Accepted` entry binds the actual E
  commit, Accepted disposition, and canonical E quiet object/raw digest, and
  requires every earlier entry to be `StaleBeforeAdmission`. Missing, reordered,
  forked, unlinked, duplicate, or extra attempts are invalid. It also embeds all
  P/A and final validation
  receipt/quiet canonical objects (including failed
  receipts that selected `ProtocolClosed`), the applicable ProtocolClosure
  object plus ProtocolRepair and both approval canonical objects when
  applicable, every ordered decision-generation evidence+disposition canonical
  object through the accepted generation (including each
  `RejectedAudit|AuditedApproved|ObservationPreempted` subject prefix,
  zero/one/two audit objects, each exact approved-decision canonical object or
  closed N/A, and the final E-bound Accepted disposition),
  and the bounded cleanup
  inventory/aggregate described below. Its canonical top-level fields include
  `recordKind=terminal-engine-recovery-archive`, `roundOrdinal=5`, and the exact
  40-hex `decisionCutoffCommit=E`. It deliberately excludes the F manifest/quiet so
  the F tree/OID has no self-reference. The FArchive/FDone/FPublish/FSeal transaction
  roots, publication journals/fence, and quiet receipts remain retained ignored
  recovery state and are outside the cleanup allowlist. Before E admission, a
  mutation-free sizing pass RS-JCS-renders the exact mandatory embedded payload
  with no cleanup inventory, using the E-manifest-bound pre-rendered Accepted
  decision disposition for the one object published only after the ref CAS. It
  rejects absent, changed, oversized, or differently predicted-E bytes. It permits at most 256 embedded canonical objects,
  64 KiB per object, and 1 MiB aggregate mandatory bytes, reserving the other
  1 MiB for bounded inventory/envelope. This is the final assertion of the
  earlier per-publication mandatory-payload budget, not the first admission
  check: any legal history already reserved and rechecked its worst terminal
  route. A mismatch here is an invariant breach and hard pre-E STOP, not
  `CleanupDisabledOverCap`, `ProtocolClosed`, or permission to discard history;
  no E/FA/FD/FP/FS is created until the budget implementation is repaired
  without changing an observed candidate verdict. The archive may never
  truncate, omit, hash-only substitute, deduplicate nonidentical canonical
  objects, or silently enlarge a cap. Corpus cases cover each mandatory
  object/count/aggregate cap at minus-one/exact/plus-one, every prospective
  admission boundary, and prove cleanup-root exclusion cannot cure
  mandatory-payload overflow. Add the mutation-free
  generic `Tools/Test-TerminalEngineRecoveryArchive.ps1`; with explicit
  `-ArchivePath`, `-RoundOrdinal`, `-ExpectedRawSha256`, and
  `-ExpectedDecisionCutoffCommit`, it derives the ordinal schema/verifier
  contract without directory enumeration. In its default clone-verification
  mode it validates only canonical objects embedded in the archive, their raw
  digests, the complete predecessor DAG ending at expected E including every
  manifest/recovery-attempt/quiet state derivation, and the embedded
  cleanup inventory/aggregate. It never resolves or reopens an ignored
  `.build` path, manifest, receipt, root, or leaf; those may legitimately be
  absent after cleanup or in a fresh clone. Only the live pre-delete cleanup
  tool reopens eligible roots/leaves. The generic verifier returns no
  success-stream data. A missing/extra/reordered link, invalid terminal state,
  non-RS-JCS bytes, BOM/trailing newline, digest mismatch, self-reference,
  unknown embedded root/leaf, or mutation is failure. Round 5 also
  updates `Specs/TestRuns/README.md` with this separate clone-recovery taxonomy,
  exact path/run-ID/content/retention contract, and verifier command before an
  archive can be published;
- add the canonical LF-terminated UTF-8-without-BOM template
  `Specs/Terminal/Templates/TerminalEngineGate0SuccessorRound.template.md`,
  `Specs/Terminal/TerminalEngineSuccessorRoundRender.schema.json`, and
  `Specs/Terminal/TerminalEngineSuccessorRoundRenderCorpus.v1.json`, plus sole
  renderer `Tools/New-TerminalEngineGate0SuccessorRound.ps1` and mutation-free
  verifier `Tools/Test-TerminalEngineGate0SuccessorRound.ps1`. The explicit
  render input has no defaults and contains only next ordinal, exact absent WIP
  basename/path, predecessor plan/FD/Done blob/decision/FA/archive/outcome,
  template raw digest, renderer raw digest, and the three literal pending carry
  sentinels for FP, FS, and seal digest. It also preserves the mandatory
  RoundStart, RecoveryArchive, E -> FA -> FD -> FP -> FS, embedded-quiet seal,
  and published-closeout-verifier contracts. A no-winner decision freezes
  `successorPlanPath`, template digest, renderer digest, and canonical render-
  input digest before E; Winner freezes the exact
  `N/A-winner-routes-product-plan-2` values. FPublish may consume only those
  audited values, renders in memory, requires the target to be absent in both
  FD and the real index/worktree, and adds the one predicted blob through its
  private index—never a standalone file write. The corpus covers every
  substitution, missing/extra/duplicate token, ordinal gap, bad basename/path,
  existing target, altered template/tool, CRLF/BOM/trailing-byte variation,
  absent RoundStart/archive/seal clause, and nondeterministic re-render;
- add generic ordinal
  `Specs/Terminal/TerminalEngineRoundStart.schema.json` and
  `Specs/Terminal/TerminalEngineRoundStartCorpus.v1.json`, sole writer
  `Tools/TerminalEngine/Publish-TerminalEngineRoundStart.ps1`, and mutation-free
  `Tools/Test-TerminalEngineRoundStart.ps1`. RoundStart uses an explicit
  create-new
  `.build/terminal-engine-spike/round<ordinal>-round-start/<40hex-FS>/
  bootstrap.v1.json` plus
  `transaction-<32-lowerhex>/manifest.v1.json`, private index, predicted
  tree/commit, `round-start-publish-journal.v1.json`, `RoundStartPublish`
  ControlQuiet token, ref compare-and-swap, and fixed
  `post-transaction-quiet.v1.json`. It accepts explicit ordinal, symbolic ref,
  FS, seal path/raw digest, master path, and successor WIP path; verifies FS is
  the valid sole-parent CloseoutSeal of the immediately prior no-winner FP,
  ref/HEAD equal FS, a clean index/worktree, and no successor coordinator epoch
  or action. Under the repository-wide exclusive RoundStart file lease, the
  generic writer is the sole bootstrap issuer: `bootstrap.v1.json` binds the
  validated FS/seal, next ordinal, exact logical key
  `round<ordinal>/control/RoundStartPublish/1/<successor-WIP-path>`, token
  canonical object/raw digest, machine-independent input identity, and
  transaction path. No ordinary round coordinator may mint this token.
  The closed transition is
  `FS -> Prepared -> CommitObjectReady -> RefAdvanced -> Synchronized ->
  Verified`; only Verified plus its quiet permits creation of the successor's
  bootstrap coordinator epoch, which binds the RoundStart commit/tree,
  manifest/journal/token/quiet raw digests. Thus the next allowed action is
  carry/universe governance or the legal initial D, never a competing
  RoundStart.
  The transaction
  permits exactly six one-occurrence value substitutions: predecessor FP, FS,
  and seal digest once in the master row and once in the successor plan. Both
  files remain ordinary `100644` blobs and all other bytes/tree paths are
  identical. Recovery accepts only ref FS/RoundStart and the predicted
  index/worktree states, uses create-new four-digit recovery tokens chained to
  the prior attempt, creates no second commit, and its canonical bootstrap,
  token, manifest, journal, recovery attempts, and quiet are embedded by the successor
  RecoveryArchive. The writer rejects any action before verified RoundStart;
  the verifier works from explicit commit/seal/archive identities without
  enumerating history. Corpus cases cross every object/ref/index/worktree/quiet
  crash edge, merge/wrong parent, five or seven substitutions, changed
  non-token byte, missing/duplicate sentinel, mismatched FP/FS/digest, wrong
  ordinal/path/mode, duplicate/concurrent bootstrap, stale/missing prerequisite
  quiet, coordinator epoch before Verified, recovery-token fork, and both P and
  legal P-less initial-D successors;
- add mutation-free
  `Tools/Test-TerminalEngineRound5PublishedCloseout.ps1` plus focused Pester
  coverage. It accepts only caller-supplied E, FA, FD, FP, FS, outcome,
  WIP/Done paths, RecoveryArchive path/raw digest, fixed CloseoutSeal path/raw
  digest, exact winner-product or no-winner-successor route, and explicit
  `-Mode InitialCloseout|RetrospectiveCarry`; it never enumerates commits, plans,
  or TestRuns. `InitialCloseout`, used by the closeout orchestrator immediately
  after FS, forbids a successor archive and requires the exact
  `N/A-initial-closeout-successor-not-yet-closed` path/digest pair. It validates
  the current E -> FS chain, seal, and audited pending-sentinel successor route
  and is sufficient to derive `ClosedF`. `RetrospectiveCarry`, used by Plan 2
  and Phase 9, requires
  `-SuccessorRecoveryArchivePath/-SuccessorRecoveryArchiveSha256` for every
  earlier `no-winner`, or the exact
  `N/A-winner-routes-product-plan-2` pair for Winner. In the former case it reads
  the explicit successor archive's embedded RoundStart object and verifies the
  direct FS parent plus six-token carry before accepting the earlier FP/FS.
  It first invokes `Tools/Test-TerminalEngineRecoveryArchive.ps1`, then proves
  the single-parent E -> FA -> FD -> FP -> FS chain and ordinary `100644 blob`
  modes.
  It reconstructs FD's expected Done bytes from the E WIP blob by applying the
  checked-in Round-5 closeout projection: every allowed `ROUND5_RESULT_*`,
  state, checkbox, completion-record, FA/path/digest, and branch-marker
  replacement must occur exactly once, Winner alone adds the one success
  marker, no placeholder may remain, and every byte not named by that projection
  is identical. It reconstructs FP's expected master from FD by changing only
  the bounded Gate-0 handoff region and its exact pending Round-5 values, with
  the rest of the master byte-identical. It reconstructs FP's WIP README from
  FD with the exact branch-specific routing substitution: Winner contains
  exactly one literal
  ``- **terminalGate0Route:** `Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md```
  line and no successor-round route; no-winner contains exactly one reviewed
  successor route. On no-winner, FP also adds exactly one successor WIP plan
  rendered from the checked-in next-round template with
  `predecessorHandoffPublishCommit=pending-parent-fp` and
  `predecessorCloseoutSealCommit=pending-parent-fs` plus
  `predecessorCloseoutSealSha256=pending-parent-fs`, while the appended master
  row uses the exact labels
  `gate0Round[N].predecessorHandoffPublishCommit=pending-parent-fp` and
  `gate0Round[N].predecessorCloseoutSealCommit=pending-parent-fs` plus
  `gate0Round[N].predecessorCloseoutSealSha256=pending-parent-fs`; no other path
  or byte may change. After the tracked FSeal described below, every Round 6+
  successor
  begins with one mandatory branch-neutral `RoundStart` Git boundary as FS's
  single-parent direct child. Before observation, transition review, P, or a
  legal P-less initial D, RoundStart atomically replaces the exact pending
  predecessor-FP and predecessor-FS tokens in both the successor WIP plan and
  master row with the FP/FS OIDs and raw seal digest proven by its parent seal;
  it changes nothing else. Verified RoundStart immediately admits the
  successor's carry audit, observation, P, or legal initial D. Its later
  RecoveryArchive authenticates that already-admitted RoundStart
  retrospectively and makes the predecessor carry clone-verifiable; it is not
  an admission prerequisite. Both modes finally require `FS^=FP`, the exact
  one-path `100644` seal
  addition, canonical embedded E/FA/FD/FP quiets, and FS ancestry. This acyclic carry
  makes every earlier no-winner FP/FS pair
  clone-verifiable; the final winner FP/FS pair, which cannot self-record, is
  supplied and persisted by product Plan 2. The corpus flips every allowed and
  forbidden byte, adds
  duplicate markers/routes/placeholders, tries merge commits and wrong parents,
  changes file modes, injects impossible/non-ASCII timestamps, and rejects a
  successor action whose RoundStart did not bind its predecessor FP/FS, plus a
  P-less TransitionRejected successor whose RoundStart is valid;
- `Specs/Terminal/TerminalEngineRound5PostFCleanup.schema.json` and
  `Specs/Terminal/TerminalEngineRound5PostFCleanupCorpus.v1.json`. They define
  the fixed ignored
  `.build/terminal-engine-spike/round5-post-f-cleanup/<40hex-FS>/` root,
  create-new `attempt-NNNN.v1.json` chain, `terminal-result.v1.json`, and
  `post-cleanup-quiet.v1.json`. State derives from the archived leaf set as
  `Prepared|Recovering|Completed|UnsafeStopped`: exact full set, strict subset,
  zero leaves/removed empty roots, or any unknown/changed/reparse/path-escape
  state. Completed and UnsafeStopped are both terminal and create the one root
  `terminal-result.v1.json` plus final quiet. UnsafeStopped preserves every
  survivor, records the exact already-removed set and closed reason in both that
  attempt's immutable `unsafe-stop.v1.json` and the root result, and permanently
  forbids another cleanup attempt. Repair/deletion after UnsafeStopped is an
  out-of-scope separately reviewed administrative action, not linked Round-5
  recovery. Corpus cases crash before/after every
  leaf/directory deletion, attempt/result/quiet publication, exhaust the ordinal,
  and prove linked idempotent recovery without double deletion or root overlap.
  Add exact limit-minus-one/limit/limit-plus-one cases for every inventory cap,
  combined aggregate overflow, a huge excluded candidate root that is never
  traversed, and a perf assertion that peak inventory memory stays bounded by
  the 2-MiB archive plus one at-most-2-MiB file/hash buffer rather than total
  root size. PostFCleanup may overlap a no-winner successor's RoundStart after
  FS, but the operations are deliberately independent: cleanup consumes only
  the exact immutable pre-E inventory embedded by FA and reopens only those
  explicit old-round roots/leaves; it never enumerates the shared
  `.build/terminal-engine-spike` parent, touches `.git`, the real index/worktree,
  a successor path, or acquires the repository-wide RoundStart/Git-boundary
  lease. RoundStart uses its distinct next-ordinal root and never reads, waits
  for, resumes, or deletes predecessor cleanup state. A new successor root is
  therefore neither an unexpected cleanup child nor eligible inventory. Crash
  tests run cleanup-first, RoundStart-first, and both overlap schedules and
  require identical tracked RoundStart bytes plus the same terminal
  `Completed|UnsafeStopped` cleanup result;
- `Specs/Terminal/TerminalEngineRound5CloseoutPublication.schema.json` and
  `Specs/Terminal/TerminalEngineRound5CloseoutPublicationCorpus.v1.json`. They
  define the exact sealed E -> FA -> FD -> FP -> FS state machine. E's create-new
  `<E-transaction-directory>/closing-publish-fence.v1.json` binds the approved
  decision and permanently closes Round-5 observation admission. FArchive's
  create-new
  `<FArchive-transaction-directory>/archive-publish-journal.v1.json` binds E,
  its prepared manifest, and ArchivePublish token. FDone's create-new
  `<FDone-transaction-directory>/done-publish-journal.v1.json` binds FA, its
  prepared manifest, and DonePublish token. FPublish's create-new
  `<FPublish-transaction-directory>/handoff-publish-fence.v1.json` binds FD, its
  prepared manifest, and HandoffPublish token. E changes only the audited
  decision/cutoff path set; FA adds only the RecoveryArchive; FD alone fills and
  moves this plan WIP-to-Done and introduces the success marker on a winner; FP
  updates only the master/README/next-round routing using the now-known FD
  commit/Done blob and FA archive tuple. After FP's fixed post-transaction quiet,
  FSeal adds only
  `Specs/Terminal/TerminalEngineRound5CloseoutSeal.v1.json` as an ordinary
  `100644 blob`. No observation/proposal action is legal between the E fence and
  verified FS. FSeal's create-new
  `<FSeal-transaction-directory>/seal-publish-journal.v1.json` binds FP, its
  prepared manifest, the embedded canonical E/FA/FD/FP quiets, and SealPublish
  token.
  The corpus crosses every
  journal/object/ref/index/worktree/quiet crash edge, derives only
  `ClosingSealed|DecisionPublished|ArchivePublished|DonePublished|
  HandoffPublished|SealPublished`, and rejects a missing/direct FA/FD/FP/FS
  predecessor, reopened
  observation admission, marker or Done bytes before FD, master routing before
  FP, seal before FP quiet, changed E decision/archive/FD Done, extra tracked
  path, alternate metadata/tree, or externally published intermediate
  E/FA/FD/FP ref;
- `Specs/Terminal/TerminalEngineRound5CloseoutSeal.schema.json` and
  `Specs/Terminal/TerminalEngineRound5CloseoutSealCorpus.v1.json`. The sole
  canonical seal is RS-JCS JSON encoded as UTF-8 without BOM and without a
  trailing newline. It embeds, rather than merely points to, the canonical
  E/FA/FD/FP `post-transaction-quiet.v1.json` objects and raw SHA-256 values, and binds
  `recordKind=terminal-engine-closeout-seal`, `roundOrdinal=5`, outcome, E, FA,
  RecoveryArchive path/raw digest, FD/Done blob, FP/tree, and the exact
  branch-specific route. FS is a single-parent direct child of FP and its diff
  is exactly that one added seal file. The ClosingPublish lease admits no action
  between the embedded FP quiet and FS ref CAS. A fresh clone derives `ClosedF`
  from the valid FS commit/seal and the published-chain verifier; an ignored
  file is never required. FS itself needs no second quiet for clone admission:
  it is a pure one-file attestational commit over an already-quiet FP, and local
  index/worktree synchronization is recoverable but not part of the durable
  product gate. The corpus rejects path/digest-only quiet references, alternate
  JSON serialization, BOM/trailing newline, a merge parent, any second path,
  mismatched tree/OID, action admission between quiet and seal, and a seal whose
  embedded quiet set is missing/reordered/mismatched or has any
  `activeActionCount` other than zero;
- `Specs/Terminal/TerminalEngineRound5DecisionAudit.schema.json` and
  `Specs/Terminal/TerminalEngineRound5DecisionGenerationEvidence.schema.json`,
  `Specs/Terminal/TerminalEngineRound5DecisionGenerationDisposition.schema.json`,
  and `Specs/Terminal/TerminalEngineRound5DecisionAuditCorpus.v1.json`. They cover
  create-new four-digit subject generations under
  `.build/terminal-engine-spike/round5-decision/generation-NNNN/`, exact
  `subject.md`, `audit-1.v1.json`, `audit-2.v1.json`,
  `generation-evidence.v1.json`, `generation-disposition.v1.json`, and
  `approved-decision.md` paths, the decision
  prefix byte range/raw digest, branch and post-effect-evidence commit/tree,
  BoundaryEvidence identity and final D transaction-manifest/quiet identities,
  author/reviewer separation, distinct reviewers, ordered verdict/UTC rules,
  Reject/stale-to-new-generation, and final publication. Publication of the
  subject's coordinator token races observation proposal publication as the
  atomic generation-admission boundary described below. A proposal that wins
  earlier consumes no subject token, ordinal, or official path; a token that
  wins reserves the next ordinal and must recover the byte-identical subject
  publication before proposal admission reopens. After audit rejection, two approvals, or an
  observation proposal that interrupts the published subject before audit
  completion, and before any successor publication, the sole writer creates
  explicit ignored `generation-evidence.v1.json`
  containing the exact subject-prefix UTF-8 bytes/raw digest,
  D/ledger/budget identities, the canonical DecisionSubject ControlQuiet token/
  raw digest, zero/one/two canonical audit objects/raw digests,
  immutable `RejectedAudit|AuditedApproved|ObservationPreempted`, and the
  immediately prior
  generation-disposition canonical object/raw digest or exact generation-1 N/A.
  `ObservationPreempted` is legal only with the exact winning proposal
  transition, no audit or exactly audit 1=`Approve`, no audit 2, and absent
  approved-decision/E; proposal-first forbids every later audit for that
  generation. A proposal arriving after an audit Reject retains
  `RejectedAudit`; one arriving after two approvals retains `AuditedApproved`.
  Thus the evidence never predicts whether a completed audited generation will
  win the later observation-vs-E CAS.
  A rejected audit then gets create-new `generation-disposition.v1.json` with
  `Rejected`, even when a later proposal is pending; normal observation
  resolution follows that terminalization. `ObservationPreempted` always gets
  `Stale`. An audited-approved generation gets its disposition only after the
  authoritative CAS: `Stale` binds the winning observation/proposal transition
  and leaves E absent; `Accepted` binds the winning E commit/ref and leaves the
  competing proposal absent. Every Stale disposition records either the exact
  already-published approved-decision canonical object/raw digest or one closed
  N/A reason for observation-before-approved-decision, plus the exact
  PrepareE-attempt identity or N/A. The disposition embeds the evidence
  canonical object/raw digest and is immutable. Both objects are at most 65,536 canonical
  bytes, create-new, and archive-destined. A successor generation requires its
  predecessor disposition to be `Rejected|Stale`; `Accepted` is terminal. A
  missing/gapped/forked evidence/disposition predecessor blocks allocation.
  The reviewed prefix contains
  the final `decisionUtc`; audit 1 is strictly after both that UTC and subject
  render, and audit 2 requires and is strictly after an approving audit 1. The
  final decision embeds
  both canonical audit objects and their raw digests but contains no self digest
  or future publication identity. RecoveryArchive embeds the complete ordered
  decision-generation evidence+disposition chain, so rejected/stale/accepted generations remain
  clone-verifiable after ignored cleanup;
- `Specs/Terminal/TerminalEngineRound5FirstSelectionInputSet.schema.json`,
  `Specs/Terminal/TerminalEngineRound5FirstSelectionAttemptReceipt.schema.json`,
  `Specs/Terminal/TerminalEngineRound5FirstSelectionAttemptLedger.schema.json`,
  and
  `Specs/Terminal/TerminalEngineRound5FirstSelectionAttemptLedgerCorpus.v1.json`.
  Before the global candidate-effect fence, the input-set writer creates the
  ignored immutable
  `.build/terminal-engine-spike/round5-first-selection/input-set.v1.json`.
  Its digest-free `inputSet` object contains exactly:
  `roundOrdinal=5`, the reviewed pre-collection execution-freeze commit/tree,
  clean-status identity, ordered path/raw-digest identities for v5 Universe,
  Review, Assessment, Set, their GovernanceFreeze journal, and Finalization plus
  finalization run/winner;
  selected candidate ID/pin and the complete upstream archive/tree, patch,
  result-tree, descriptor, concern-map, owner-ledger, and shared-lock-payload-
  schema identities; every source/dependency/cache/toolchain component and
  restore/provider/harness/probe/build/verify/coordinator script or module
  identity; exact feature flags and clean/offline/network-isolation/exclusive-
  runner policy; and exactly five ordinal lane projections for native x64
  Debug/Release/`ASan Debug` and ARM64 Debug/Release with expected output,
  public-header, export, ABI-layout, capability, runtime-closure, and license/
  notice identities. Every file leaf is
  `{kind,repositoryRelativePath,sha256,sizeBytes}`; expected non-file values come
  only from the selected frozen candidate descriptor. Arrays use the closed
  order just listed, then lane ordinal and descriptor order—never filesystem or
  lexical discovery. `inputSetSha256` is lowercase SHA-256 over RS-JCS of exactly
  that digest-free `inputSet`; the object cannot contain its hash, enclosing
  file path/digest, an attempt, result, or branch.

  The coordinator transfers the exact input-set bytes to every reserved native
  runner and invokes five preflight-only checks before any candidate effect.
  Each preflight reopens and hashes every applicable leaf, proves the requested
  native machine/configuration, clean execution-freeze commit, complete offline
  cache/toolchain, exclusive disposable runner, and no candidate subprocess/
  clean/output mutation, then returns one content-free immutable preflight
  record. If a claimed preflight crashes or times out before returning that
  record, the coordinator's bounded-cancel path creates the sole terminal
  content-free failed-preflight record defined below; an absent preflight can
  never leave the receipt sequence uncloseable. The receipt writer derives—not
  accepts—the attempt disposition:
  `Admitted` only when all five fixed preflights pass; otherwise
  `RejectedPreEffect` with the first closed reason/reproducer. It closes one
  ignored, immutable JCS record at the explicit
  `.build/terminal-engine-spike/round5-first-selection/attempt-1.v1.json` or
  `attempt-2.v1.json` path. Each receipt embeds the raw input-set path/digest and
  `inputSetSha256`, all five ordered preflight path/raw-digest/verdict
  identities and `preflightSetSha256` (lowercase SHA-256 over RS-JCS of exactly
  their ordered `{ordinal,platform,configuration,path,sha256,verdict}` objects),
  winner finalization, selected candidate, coordinator epoch/action
  token, ordinal, derived `Admitted|RejectedPreEffect`, UTC, content-free reason/
  reproducer or exact admitted N/A, and
  `entryIdentitySha256`. That identity is lowercase SHA-256 over RS-JCS of the
  digest-free subobject containing exactly
  `{ordinal,inputSetRawSha256,inputSetSha256,preflightSetSha256,epoch,
  actionTokenSha256,disposition,utc,reasonCode,reproducerSha256}`; neither that
  subobject nor the hash input contains `entryIdentitySha256`, a receipt path,
  or a receipt digest. Ordinal 2 must bind the byte-identical input-set file and
  identity, revalidate every leaf, and have UTC strictly later than ordinal 1.
  A first `Admitted` receipt opens the fence; a first rejection requires exactly
  one identical-input retry unless approved supersession stops admission first;
  a second rejection forbids all candidate effects. Declining that mandatory
  retry without supersession is not a closeout branch and leaves the round WIP.

  The same coordinator serializes every official Round-5 pipeline action
  globally across the x64 and ARM64 runners and issues one immutable pipeline-
  action token per action. Only the closed control operations below use the
  separate serialized no-effect control-plane token class; a live-supersession
  control operation may run under `ObservationPending` without occupying or
  bypassing the one pipeline slot.
  A pre-A epoch binds boundary P and admits the bounded proof, all post-proof
  nomination/governance writers, A commit/approval publication, and supersession
  control operations. After all proof/nomination/governance/A-publication work
  stops, the coordinator
  rotates under lock to the A-bound epoch; that epoch admits collection and
  first-selection actions only after the two A approvals validate. Epoch
  rotation requires the prior quiet receipt, preserves the complete token/action/
  observation history, and cannot reuse a nonce or action ID.
  Every `Pipeline` token binds `actionKind` and the canonical logical key
  `round5/<stage>/<candidate-or-none>/<platform-or-none>/
  <configuration-or-none>/<ordinal>/<fixed-output-relative-path>`, plus one
  action ID, native runner identity, create-new disposable output root, input
  identity, epoch, and nonce. A `ControlLive|ControlQuiet` token instead binds
  `actionKind`, a canonical
  `round5/control/<stage>/<ordinal>/<fixed-target-relative-path>` key, control
  operation ID, epoch/input identity, nonce, and fixed target; it contains no
  runner claim or candidate output root. The coordinator's append-only action
  journal
  reserves at most one original token/claim per logical key and derives the
  allowed-next keys from this closed sequence: bootstrap carry/universe/P-
  candidate, `PreProofFocused`, then either Pass plus P approvals/quiet or Failed
  plus `ValidationClosing`; P-bound proof, nomination bundle, A-candidate,
  `PreResultFocused`, then either Pass plus A approvals/quiet or Failed plus
  `ValidationClosing`; A-bound ordered
  collection/Finalize/admission/lock-independent/materialization/lock/B/
  lock-bound/verification/dependency work; branch cleanup; D,
  `FinalFocused`, `FinalCanonicalFull` with its single fixed sibling final-suite
  quiet, then decision governance only on Pass/allowed classification. A Failed
  ordinary final suite instead permits the one
  `ValidationClosing -> terminal receipt/suite quiet -> stop/pre-cleanup quiet
  -> C/C-quiet when B exists -> ProtocolClosurePublish -> optional
  ProtocolRepairSubject -> ProtocolRepairApproval1 ->
  ProtocolRepairApproval2 -> ProtocolRepairFinalize -> administrative
  SuccessorD -> FinalFocused -> FinalCanonicalFull` route; a failure on that administrative
  D has no automatic successor. At any quiet pre-E state, a prospective-budget
  rejection permits only
  `CapacityClosing`; either closing state uses the same exact order: allow an
  already-admitted action/suite to publish only its reserved terminal
  receipt/quiet, then stop/pre-cleanup quiet, state-specific C plus C quiet,
  ProtocolClosurePublish, the ordered repair subject/approval1/approval2/
  finalizer keys only for the failed-ordinary-final-suite route,
  `N/A-no-protocol-repair-required` for that eligible route with no byte
  change, and `N/A-no-protocol-repair-permitted` for every ineligible route;
  BoundaryEvidence only when its fixed path is still
  absent, protocol D/SuccessorD, both final suites, and no-winner decision
  governance. With no byte repair the repair disposition is the applicable exact
  N/A, all four repair keys remain absent/unallocated, and no repair action is
  allocated. Decision governance then admits only
  `RejectedAudit -> DecisionGenerationEvidence -> RejectedDisposition -> next
  generation` when no proposal is pending, or after a pending proposal's
  reviews/fence/ledger and successor D when one is present;
  `ObservationPreempted -> DecisionGenerationEvidence ->
  StaleDisposition -> observation resolution/ledger -> successor D/new
  generation`, or
  `AuditedApproved -> DecisionGenerationEvidence` followed, when no proposal is
  pending, by `approved-decision -> PrepareE`; a proposal that wins after audit
  2, after evidence, or after approved-decision but before PrepareE instead
  requires `StaleDisposition -> observation resolution/ledger -> successor D/
  new generation` and forbids any still-absent approved-decision or PrepareE
  publication. The observation-vs-ClosingPublish-fence CAS then permits only
  `StaleDisposition -> stale-before-admission -> observation resolution/ledger
  -> successor D/new generation` when a proposal wins after PrepareE but before
  the ClosingPublish fence, or
  `E ref CAS -> AcceptedDisposition -> E quiet`; then post-E
  observation admission stays sealed and the only allowed keys are
  `PrepareFArchive` -> `ArchivePublish` -> FA quiet ->
  `PrepareFDone` -> DonePublish -> FD quiet -> `PrepareFPublish` ->
  HandoffPublish -> FP quiet -> `PrepareFSeal` -> SealPublish -> FS ->
  `ClosedF` -> optional `PostFCleanup` plus recovery.
  Evidence arriving during that sealed interval is not a Round-5 action and
  waits for the separate post-publication process.
  Branch
  and supersession transitions remove keys but never add an out-of-order key; a
  successor D invalidates both final-suite keys and allocates their new D-bound
  logical keys. `ProtocolClosed` invalidates every candidate verdict and
  product-handoff key, and neither P nor A has a repair/new-candidate edge. A
  fresh ID/nonce, and for Pipeline a fresh runner/output root,
  cannot repeat a logical key. Only a linked absent-output writer retry, typed
  Git-boundary recovery, or linked PostFCleanup
  crash/partial-deletion recovery before its terminal result may reuse it under
  their closed rules.
  Before Pipeline effects, that runner must
  atomically create the one runner-local claim/start receipt at the token's
  fixed path; a second/concurrent claim, copied token on another runner, reused
  action/output root, or mismatched machine rejects before effects. A claimed
  candidate/effect action that crashes is started and follows cancel/terminal
  closeout; neither it nor its token is retried. Repository mutation recovery is
  limited to the generic `GitBoundary` transaction for
  P/A/B/C/D/SuccessorD/E/FArchive/FDone/FPublish/FSeal
  and its state-specific `TransactionalCleanup` specialization. Each resumes the
  same manifest-predicted transaction under a linked recovery token and is never
  an engine/effect rerun or a second logical action. For a proof crash, the
  coordinator creates the schema-valid content-free terminal proof outcome
  `Rejected/coordinator-action-crash`, binding token, claim, timeout/cancel, and
  no fabricated proof bytes. For a collection crash, it creates that lane's
  schema-valid terminal failed manifest; all candidate result entries are
  explicitly unavailable with the same closed crash evidence, and Finalize may
  derive only `no-scored-candidate`. The remaining named collection lanes still
  run under fresh tokens so the three-manifest set closes; no winner can result
  from a set containing a synthesized failed manifest. A claimed dependency
  status/cache/notices action that crashes or is cancelled before its report
  closes receives one coordinator-synthesized schema-valid content-free
  `failed:coordinator-action-crash` report, so the closed completion grammar and
  corresponding Qualification stage remain representable. Other later
  first-selection action crashes use Qualification-failed's exact
  started/absent-artifact evidence. A claimed effect-free runner
  preflight that times out or crashes before publishing its record is not
  rerun: after bounded cancel/runner quarantine, the coordinator synthesizes
  one schema-valid content-free `RejectedPreEffect` preflight record from the
  token, claim receipt, timeout/cancel evidence, and closed runner-failure reason,
  so the five-preflight set and admission receipt can still close. A claimed
  pure writer whose atomic final output remains absent may resume only with a
  new writer-attempt token linked to the prior claim and the byte-identical
  logical action/input/UTC/verdict in the writer's same fixed action class; it
  never reruns an engine action, preflight,
  or reviewer judgment. Returned records embed every claim and linked retry
  receipt/raw digest.
  Prospective budget calculation and the
  `CapacityClosing|ValidationClosing` CAS are coordinator-internal transition
  guards, not separately invocable operations and not token-bearing action
  classes. The calculation executes under the same coordinator lock and
  publication CAS as the history-growing action it guards. Capacity rejection
  allocates no requested action token/path; suite failure derives
  `ValidationClosing` atomically with publication of that already-admitted
  suite's reserved Failed receipt. Neither transition requires a pre-existing
  quiet receipt or a second action, and the already-admitted Pipeline action
  retains its slot until its reserved terminal receipt/quiet is published.
  Only after that quiet plus the required stop/pre-cleanup and C quiet may the
  independently tokened `ProtocolClosurePublish` ControlQuiet operation start.
  Corpus tests invoke every public entry point and prove no direct callable
  closing-transition bypass exists.

  Every invocable entry point has one immutable action class:

  | Class | Exhaustive operations |
  |---|---|
  | `Pipeline` — occupies the sole global slot | carry/universe/P/A governance publication; P/A/B/D/SuccessorD `GitBoundary` transactions and linked recovery; all four typed validation suites; proof; proof result; patch/owner-ledger work; Review/Assessment/Set GovernanceFreeze bundle and approvals; provider/collection; Finalize; input-set/preflight/receipt; every lock-independent or lock-bound candidate lane; materialization and approvals; candidate lock and approvals; tracked adoption; the single C/cleanup `TransactionalCleanup` Git-boundary specialization and its linked recovery; Build/Verify; verification set; dependency status/cache/notices reports; and any linked absent-output retry of one of these writers |
  | `ControlLive` — serialized under the coordinator lock and allowed alongside the one pipeline action only in `ObservationPending` | coordinator state/token/claim journal, live supersession proposal/review/fence, cancel request, and epoch check |
  | `ControlQuiet` — requires zero pipeline actions and a quiet receipt | pre-D content-free evidence/status archiving; `ProtocolClosurePublish`; ProtocolRepair subject/approval-1/approval-2/finalize and their absent-output recovery; attempt ledger; winner-handoff proposal/reviews; observation ledger; execution-freeze attestation; qualification-failure or Supersession result; decision subject/audits; `DecisionGenerationEvidence`; Rejected/Stale/Accepted `DecisionGenerationDisposition` plus absent-output recovery; `PrepareE`; E ClosingPublish-fence/publication; sealed `PrepareFArchive`, `ArchivePublish`, `PrepareFDone`, DonePublish, `PrepareFPublish`, HandoffPublish, `PrepareFSeal`, SealPublish, FArchive/FDone/FPublish/FSeal publication/recovery and ClosedF derivation; `PostFCleanup` plus linked crash/partial-deletion recovery before a terminal result; completion verification; and `SupersededCarryForward` proposal/reviews/ledger |

  Every `Pipeline` entry point requires an explicit pipeline-action token path/
  raw digest, expected epoch/input identity, runner-local claim, and unique
  create-new output root. Every fixed `ControlLive|ControlQuiet` entry point
  instead requires the corresponding control-plane token path/raw digest and
  expected epoch/input identity; `ControlQuiet` additionally requires zero
  pipeline actions and the exact quiet receipt. Control operations never require
  or fabricate a runner claim or candidate output root. Every entry point
  reopens its required identities and rejects a stale,
  wrong-action, missing, or bypassed token before attestation, bootstrap, clean,
  subprocess, output, or tracked mutation. Each candidate lane also requires
  `-AttemptReceiptFile`, `-ExpectedAttemptReceiptSha256`, and
  `-ExpectedInputSetSha256`, consumes the same transferred admitted receipt, and
  binds those identities plus the execution-freeze commit in its build,
  verification, and lane records. The coordinator validates the epoch again
  before accepting/publishing returned evidence; a superseded in-flight result
  remains content-free invalidated evidence and cannot open the next action.

  After the current action stops, the coordinator issues an immutable
  branch-neutral quiet receipt containing epoch, last action ID/type, action-
  stopped UTC, `activeActionCount=0`, and exact admission-receipt identities.
  The tracked closed ledger requires that quiet receipt, records `frozenUtc`
  strictly later than the quiet receipt and every attempt UTC, and embeds the
  complete canonical input-set record, every attempt's five preflight records,
  action tokens, attempt receipts, and quiet receipt together with each raw
  digest. The verifier recomputes `inputSetSha256`, every preflight-set/entry/raw
  digest, derived attempt disposition, epoch/action chronology, and quiet state
  from those embedded objects without depending on ignored `.build` bytes. It
  has exactly one of four shapes:
  `AdmittedFirstAttempt`; `AdmittedAfterRetry` from
  `RejectedPreEffect,Admitted`; `SecondRejection` from two
  `RejectedPreEffect` receipts; or `StoppedAfterFirstRejection` from one rejected
  receipt when supersession closes admission before retry. No other stage,
  changed retry input, third attempt, or partial sequence is valid. It contains
  no branch, success/failure, clean/revert, proposal/review, or decision
  identity. The ledger is branch-neutral selection history, is never a
  candidate build/verification input, and therefore cannot dirty or change the
  clean repository commit bound by Finalize and the candidate lanes. Winner,
  Qualification-failed, and Superseded bind its same raw identity and add their
  own outcome/clean/revert disposition separately;
- `Specs/Terminal/TerminalEngineRound5QualificationFailure.schema.json` and
  `Specs/Terminal/TerminalEngineRound5QualificationFailureCorpus.v1.json`.
  The closed record exists only when the exact finalization has
  `selection.status="winner"` and a later first-selection stage fails. It binds
  that finalization and selected candidate, the closed failing stage
  `CandidateLane|FirstSelectionAttemptLedger|Materialization|
  MaterializationApproval|CandidateLock|
  LockApproval|TrackedAdoption|LockBoundBuildVerify|VerificationSet|
  DependencyLockedStatus|DependencyCache|NoticesLicense|
  WinnerHandoffProposal|WinnerHandoffProposalApproval`, the exact content-free
  reason/reproducer, the required attempt-ledger path/raw digest/disposition,
  and every completed, failed, or absent later artifact with its
  path/digest/state. Only `CandidateLane` may fail on the pre-effect side
  of the global fence; its ledger permits exactly the one identical-input retry
  above. Every stage other than `CandidateLane` and
  `FirstSelectionAttemptLedger` is reachable only after
  `AdmittedFirstAttempt|AdmittedAfterRetry` receipt history and has no retry.
  `FirstSelectionAttemptLedger` may be the first terminal failure while closing
  either two-rejection receipts before effects or an otherwise complete admitted
  sequence after dependency/notices closure; those two exact artifact rows are
  defined below. If another substantive pipeline stage already failed, that
  original stage remains `qualificationFailureStage` and the ledger validation
  defect is a secondary legal invalid/absent substitute; it never relabels or
  erases the original partial-artifact row. The Superseded schema separately
  owns its interrupted-retry shape and any ledger-validation failure at that
  fence. If receipt validation is terminal at `CandidateLane`, or tracked ledger
  validation/publication is terminal at closeout, the failure record binds the
  canonical content-free `invalidArtifactEvidence` object and exact
  `N/A-validation-rejected-before-publication` plus the canonical attempted
  entries/input digest; the tracked target remains absent. Those are the sole
  legal non-superseded Finalize-winner
  cases without a valid ledger. Its timing is
  exactly `PreAdoption` when no reviewed adoption commit exists or
  `AdoptedThenReverted` when one does. `PreAdoption` requires the candidate
  lock to remain below `.build`, no tracked adoption commit, and a byte-clean
  pre-adoption tree. `AdoptedThenReverted` requires the exact adoption and
  non-history-rewriting revert commits/trees, the reverted-path digest set, and
  proof that zero tracked engine adoption remains. The record always binds the
  pre-cleanup quiet receipt; `NotStarted` binds the exact cleanup/post-quiet N/A,
  while `CancelledBeforeCommit|AdoptedThenReverted` bind one cleanup Pipeline
  token/claim/terminal result plus a distinct post-cleanup quiet receipt. Both timings prohibit a
  runner-up, success marker, production lock, product-plan handoff, or reuse of
  a failed candidate result in the same round;
- `Specs/Terminal/TerminalEngineRound5WinnerHandoffProposal.schema.json` and
  `Specs/Terminal/TerminalEngineRound5WinnerHandoffProposalCorpus.v1.json`.
  The closed proposal binds the exact winner finalization, selected candidate,
  materialization, candidate lock, tracked adoption, five-lane verification
  set, dependency status/cache, notices/license closure, and proposed success-
  handoff values. It contains no Round-5 decision path/digest or later review
  identity and is never overwritten. It records exact `proposalAuthorId`,
  creation UTC, the attempt-ledger path/raw digest, and exactly the pre-closeout
  six-tuple `successfulGate0FinalizationManifest`,
  `successfulGate0FinalizationSha256`, `successfulGate0WinnerCandidateId`,
  `successfulGate0EngineLockSha256`,
  `successfulGate0VerificationSetManifest`, and
  `successfulGate0VerificationSetSha256`. It cannot name the future Done
  basename/blob, closeout commit, current/immutable Round-5 decision identity,
  or success marker/publication claim;
- `Specs/Terminal/TerminalEngineRound5WinnerHandoffReview.schema.json` and
  `Specs/Terminal/TerminalEngineRound5WinnerHandoffReviewCorpus.v1.json`.
  Each closed review record contains its fixed ordinal/path, the exact proposal
  path/raw digest/author identity, one reviewer identity, `Approve|Reject`,
  strictly later UTC, and content-free rationale. Ordinal 1 has
  `priorReview=N/A-first-review`; ordinal 2 additionally binds the exact ordinal-
  1 path/raw digest/reviewer/`Approve` verdict and must be strictly later. The
  proposal author cannot review, the two reviewers must differ, ordinal 2 is
  forbidden after an ordinal-1 rejection, and neither fixed review path may be
  overwritten. Two approving review records permit the Winner branch; either
  rejection terminates as `WinnerHandoffProposalApproval` qualification failure
  and remains exact immutable failure evidence;
- `Specs/Terminal/TerminalEngineCandidateUniverseTransition.v5.schema.json` and
  `Specs/Terminal/TerminalEngineCandidateUniverseTransition.v5.corpus.json`, or
  one reviewed version-generic successor plus an equivalently named positive/
  negative corpus, covering the Round-4 predecessor, complete seed/control
  carry-forward, one WezTerm replacement, proof-spike policy, strict chronology,
  and explicit frozen-input identities;
- `Specs/Terminal/TerminalEngineRound5Proof.schema.json`,
  `Specs/Terminal/TerminalEngineRound5ProofCorpus.v1.json`, and the immutable
  `Specs/Terminal/TerminalEngineRound5Proof.v1.json`, covering
  `PendingProof|ProofSucceeded|BudgetExceeded|Rejected`, the ordinal owner-day
  ledger, exact hypothesis/input/result identities, the seven-action ABI table
  below, stop reason, and retained-bytes accounting;
- `Specs/Terminal/TerminalEngineCandidateReview.v5.schema.json`,
  `Specs/Terminal/TerminalEngineCandidateReview.v5.corpus.json`,
  `Specs/Terminal/TerminalEngineCandidateAssessment.v5.schema.json`,
  `Specs/Terminal/TerminalEngineCandidateAssessment.v5.corpus.json`,
  `Specs/Terminal/TerminalEngineCandidateSet.v5.schema.json`, and
  `Specs/Terminal/TerminalEngineCandidateSet.v5.corpus.json`, plus
  `Specs/Terminal/TerminalEngineCandidateGovernanceFreeze.v5.schema.json` and
  `Specs/Terminal/TerminalEngineCandidateGovernanceFreeze.v5.corpus.json`.
  Bind the six Review/Assessment/Set schema/corpus leaf identities into the v5
  universe and every explicit-path reader
  before it freezes. In addition to the existing
  source/adaptation/collect branches, they define one
  `proof-rejected` branch. It requires the exact v5 WezTerm pin/hypothesis and
  proof-record digest, terminal proof state `BudgetExceeded|Rejected`, closed
  reason code, N/A patch/concern/descriptor identities, zero admission, and no
  collection. The CandidateSet contains one explicit WezTerm-v5
  `proof-rejected` ledger entry plus every carried seed/control outcome; it
  never silently reuses the prior WezTerm-v3 adaptation rejection as the v5
  result. The GovernanceFreeze is a one-way journal: it binds a coordinator-
  allocated freeze ID/common UTC, exact reserved raw digests for Review,
  Assessment, and Set, and every prerequisite approval. The three subjects bind
  only that freeze ID/common UTC, never the journal or sibling digests. Journal
  no-replace publication is the authoritative Pipeline commit; recovery then
  publishes only the three reserved bytes before that action can quiet, so a
  Supersession result sees either the complete bundle or no bundle, never a
  partial combination or hash cycle. `ProofSucceeded` may use only the ordinary complete
  adaptation-rejected or collect branches; and
- always create
  `External/TerminalEngine/CandidatePatches/wezterm-rs-v5.owner-ledger.json` and
  the proof record for every retained proof byte/day. Only after
  `ProofSucceeded`, create
  `External/TerminalEngine/CandidatePatches/wezterm-rs-v5.patch`,
  `External/TerminalEngine/CandidatePatches/wezterm-rs-v5.concerns.json`, and
  `External/TerminalEngine/CandidatePatches/wezterm-rs-v5.descriptor.json` as the
  normal nomination artifacts. `BudgetExceeded|Rejected` marks those three
  nomination paths explicitly N/A rather than fabricating a complete candidate.
  Freeze the ledger and every applicable nomination artifact only after the
  bounded proof and, on success, nomination work finishes. The
  ledger includes every candidate-specific source, generated bridge, wrapper,
  test, fuzz input, benchmark, document, descriptor, dependency/license file,
  and retained spike byte wherever it lives; moving a byte outside the upstream
  result tree never removes it from changed-file, changed-line, concern, binary,
  toolchain, runtime-leaf, or owner-day accounting.

Add or update the executable tooling:

- add `Tools/TerminalEngine/Invoke-TerminalEngineRound5Proof.ps1` as the sole
  bounded proof driver and `Tools/Test-TerminalEngineRound5Closeout.ps1` as the
  mutation-free closeout verifier;
- add
  `Tools/TerminalEngine/Test-TerminalEngineRound5MandatoryPayloadBudget.ps1` as
  the mutation-free budget-table/DAG calculator and
  `Tools/TerminalEngine/New-TerminalEngineRound5ProtocolClosure.ps1` plus
  `Tools/TerminalEngine/New-TerminalEngineRound5ProtocolRepair.ps1` and
  `Tools/TerminalEngine/New-TerminalEngineRound5ProtocolRepairApproval.ps1` as
  the sole reserved administrative-closure/repair subject/finalizer and approval
  writers. The calculator accepts only explicit
  table/current-state/candidate-object paths and never enumerates ignored
  history. The closure writer accepts only the coordinator-reserved record path,
  `CapacityClosing|ValidationClosing` token, failed-suite or prospective-budget
  canonical input, exact stop/cleanup projection, and remaining-reserve
  snapshot; it cannot publish a candidate verdict or winner field. The repair
  writer accepts only the exact `final-validation-failed` ordinary-D closure,
  its failed final-suite receipt/quiet, and explicit pre-frozen
  allowlist/current-tree/before-and-after paths; it rejects every capacity,
  P/A, pre-ordinary-D, and already-administrative-D closure, recomputes the
  exact patch and excluded invariant
  digests, stages the create-new subject, then after the approval writer has
  published exact ordinal 1/2 records verifies both and stages the fixed wrapper
  without modifying the real index/worktree, and supplies its approved bytes to
  the one administrative SuccessorD. The four fixed ControlQuiet logical keys
  are exactly
  `round5/control/ProtocolRepairSubject/1/.build/terminal-engine-spike/round5-protocol-closure/protocol-repair-subject.v1.json`,
  `round5/control/ProtocolRepairApproval/1/.build/terminal-engine-spike/round5-protocol-closure/protocol-repair-approval-1.v1.json`,
  `round5/control/ProtocolRepairApproval/2/.build/terminal-engine-spike/round5-protocol-closure/protocol-repair-approval-2.v1.json`,
  and
  `round5/control/ProtocolRepairFinalize/1/.build/terminal-engine-spike/round5-protocol-closure/protocol-repair.v1.json`;
  the coordinator permits them only in that order and budgets every object
  before the subject;
- add `Tools/TerminalEngine/TerminalEngineRound5Coordinator.psm1` and
  `Tools/TerminalEngine/Invoke-TerminalEngineRound5Action.ps1` as the sole
  official admission/commit boundary. The ignored create-new epoch state admits
  exactly one effect action at a time across every local/native runner, issues
  JCS pipeline-action and serialized no-effect control-plane tokens bound to
  epoch/action-kind/logical-key/tool/input identities, enforces the closed
  allowed-next transition and at-most-one original claim per logical key, and
  accepts a returned result only after rechecking the same epoch under lock.
  Observation proposal publication and ControlQuiet token admission use one
  mutually exclusive CAS: a token first must reach its one terminal output/
  close before proposal retry, while proposal first blocks every new
  ControlQuiet key except the bounded decision terminalizer.
  Before allocating any history-growing token/path it invokes the pure
  mandatory-payload calculator under that lock, binds the table/snapshot digest
  into the token, and reruns the calculation immediately before the
  no-replace publication. A capacity loss publishes no partial action and
  atomically selects `CapacityClosing`; a P/A/final-suite Failed receipt selects
  `ValidationClosing`. Both seal ordinary pipeline and observation admission,
  stop/quiet the admitted action, run required zero-engine cleanup, and expose
  only the reserved ProtocolClosure/D/final-validation/no-winner closeout keys,
  plus ProtocolRepairSubject/Approval1/Approval2/Finalize in that order only for
  an eligible changed-byte failed-ordinary-final-suite route. Every other route,
  including eligible no-byte-change, allocates none of those four repair keys.
  It creates the pre-A P epoch, admits proof and all post-proof nomination/
  governance writers plus A commit/approval publication there, emits its quiet
  receipt only after they stop, and atomically rotates to the reviewed A-bound
  epoch before collection. Its bounded timeout/
  quarantine path creates the single terminal failed-preflight record when a
  claimed effect-free preflight cannot publish one; that record binds the
  original token/claim and cannot be used to rerun the preflight. Its typed
  crash closeout creates only the proof
  `Rejected/coordinator-action-crash` outcome, failed collection manifest, or
  failed dependency/cache/notices report described above; it cannot synthesize
  success, candidate measurements, or a winner.
  Update proof, provider, evaluator/finalizer, materialization/lock/adoption,
  Build/Verify/verification-set, dependency/cache/notices, ledger, proposal/
  review, qualification, supersession, and decision entry points so direct or
  nested use without the matching explicit token path/raw digest rejects before
  effects. Add
	  `Tools/TerminalEngine/Test-TerminalEngineRound5RunnerPreflight.ps1` for the
	  five preflight-only remote checks; it cannot invoke a candidate subprocess or
	  create/clean a candidate output;
- add `Tools/TerminalEngine/Invoke-TerminalEngineRound5GitBoundary.ps1` as the
  sole P/A/B/C/D/SuccessorD/E/FArchive/FDone/FPublish/FSeal commit
  transaction writer and
  `Tools/TerminalEngine/Invoke-TerminalEngineRound5CleanupTransaction.ps1` as
	  its state-specific cleanup front end. Before repository mutation, the boundary
	  writer requires the caller-supplied boundary kind/transaction ID and exact
	  predecessor-manifest path/raw digest or first-boundary N/A, validates and no-
	  replace publishes the explicit ignored immutable manifest at the fixed
	  create-new path above, and materializes the after-tree in a private index from the
  expected old OID, verifies the exact intended tree/commit metadata/object ID,
  creates at most that commit object, and advances only the manifest's symbolic
  branch ref through `git update-ref <ref> <new-oid> <old-oid>` or an equivalent
  atomic compare-and-swap. It then synchronizes and verifies the checked-out real
  index/worktree and emits the post-transaction quiet receipt. Recovery uses the
  same manifest and a linked recovery token, derives object-not-ref versus ref-
  moved-not-synchronized state, and never creates a second commit or replays
  candidate work. The cleanup front end additionally implements the exact
	  no-commit `CancelledBeforeCommit` restoration or commit-backed C projection/
	  zero-engine-state contract. Add
	  `Tools/TerminalEngine/Remove-TerminalEngineRound5IgnoredEvidence.ps1` as the
	  sole `PostFCleanup` ControlQuiet action: FArchive first creates its exact
	  tracked RecoveryArchive from
	  the explicit content-free transaction-manifest chain ending at E and final
	  validation receipt identities. Automatic cleanup eligibility is closed to
	  ignored, content-free coordinator/receipt roots named by those embedded
	  objects; candidate source, checkout, build, binary, package, cache, and
	  repro roots are always `RetainedNonCleanup` and are never inventoried for or
	  touched by this action. Inventory is streaming and bounded before
	  publication: at most 64 eligible roots, 512 leaves per root, 2,048 leaves
	  total, depth 16, 512 UTF-8 bytes per normalized repository-relative path,
	  2 MiB per ordinary file, 32 MiB total eligible bytes, and 2 MiB for the
	  resulting canonical archive. A root that first breaches a per-root
	  leaf/depth/path/file/byte limit is not partially eligible: preserve the
	  entire root and record one bounded `RetainedOverCap` disposition with the
	  first breached limit and observed lower bound, without retaining a partial
	  leaf list. If the root-count, aggregate-leaf/byte, or archive-size limit
	  would be exceeded, preserve every not-yet-admitted root and record the
	  aggregate `CleanupDisabledOverCap` disposition; never allocate, hash,
	  commit, or delete beyond the limit. The archive contains, for each admitted
	  root, its normalized repository-relative path below the fixed
	  `.build/terminal-engine-spike/` root, every relative leaf's
	  `{path,type,size,sha256}`, and an RS-JCS aggregate identity. The cleanup tool
	  requires `ClosedF`, the tracked CloseoutSeal and its embedded exact FPublish
	  post-transaction quiet, and the fixed
	  logical key `round5/control/PostFCleanup/1/<RecoveryArchive-path>`. Roots are
	  unique and pairwise non-overlapping—none equals/contains another—and every
	  leaf belongs to exactly one root. The tool accepts only caller-enumerated
	  roots, resolves each against the current canonical repository root, proves
	  strict containment below the fixed ignored root, reopens every ancestor/leaf,
	  and rejects a
	  reparse, extra child, non-ordinary leaf, path escape, or identity mismatch,
	  and deletes only the exact leaf set followed by now-empty listed
	  directories. Recovery after a partial deletion accepts only an exact subset
	  of the archived leaves, revalidates every survivor, and continues
	  idempotently; an unknown child or changed survivor is STOP. It never
	  enumerates a latest/generation, uses a wildcard, follows a reparse, or removes
	  final/stale roots before accepted E and FSeal. The
	  FArchive/FDone/FPublish/FSeal transaction roots/manifests/journals/fence/
	  quiets are
	  excluded from both that self-reference-sensitive archive and this cleanup
	  allowlist and remain retained ignored recovery state. Cleanup attempts use
	  explicit create-new
	  `.build/terminal-engine-spike/round5-post-f-cleanup/<40hex-FS>/
	  attempt-NNNN.v1.json` records. Attempt 1 binds the cleanup control token,
	  archive path/raw digest, root aggregate set, and exact starting inventory;
	  each recovery attempt binds its immediate predecessor/raw digest and new
	  linked control token. State is derived from the remaining exact leaf subset,
	  not a mutable progress flag. At zero leaves/removed listed empty roots, the
	  tool no-replace writes `terminal-result.v1.json` with `Completed`. On unsafe
	  state it preserves all survivors, no-replace writes the current attempt's
	  `unsafe-stop.v1.json`, and writes the root result with `UnsafeStopped`, exact
	  removed/surviving sets, and the closed content-free reason. The coordinator
	  then writes `post-cleanup-quiet.v1.json` binding every attempt/result,
	  `activeActionCount=0`, and FS/FP/RecoveryArchive identities. Either terminal
	  result permanently closes the cleanup key; only a crash or exact strict-
	  subset state before that result permits linked recovery. This same logical
	  PostFCleanup key is never restarted as a fresh action. The cleanup-attempt
	  root is also retained and excluded from its own deletion set;
- add `Tools/TerminalEngine/Publish-TerminalEngineRound5Closeout.ps1` as the
  sole E -> FA -> FD -> FP -> FS orchestrator. It acquires ClosingPublish only after the
  approved decision generation and final quiet validate, creates E, and keeps
  Round-5 observation admission sealed until verified FS. It then invokes
  `PrepareFArchive`, creates exactly
  `<FArchive-transaction-directory>/archive-publish-journal.v1.json`, advances
  E -> FA, synchronizes/verifies/quiets, invokes `PrepareFDone`, creates exactly
  `<FDone-transaction-directory>/done-publish-journal.v1.json`, renders the
  reviewed completed-plan mapping, advances FA -> FD, moves the plan
  WIP-to-Done, and adds the success marker only for Winner. It then invokes
  `PrepareFPublish`, creates exactly
  `<FPublish-transaction-directory>/handoff-publish-fence.v1.json`, and advances
  FD -> FP. FPublish fills the master with the now-known FD closeout commit/Done
  blob and FA/path/digest handoff tokens and updates only README/next-plan
  routing. After FP's fixed quiet validates with `activeActionCount=0`, the tool
  invokes `PrepareFSeal`, renders the canonical embedded-quiet CloseoutSeal,
  creates exactly
  `<FSeal-transaction-directory>/seal-publish-journal.v1.json` with the
  SealPublish token, advances FP -> FS with an exact one-file-add tree, and
  validates the mutation-free published chain in `InitialCloseout` mode with
  the exact successor-not-yet-closed N/A pair.
  Every invocation first derives its state from the exact manifests, ref,
  index/worktree, decision, archive, and journal/fence; it resumes only the
  missing deterministic boundary and never creates a second plan/result/marker.
  E, FA, FD, and FP are local non-publishable intermediate refs. The tool rejects an
  attempted push/export/handoff from E, FA, FD, or FP, an observation action after
  the E fence, completed-plan/marker mutation before FD, master/routing mutation
  before FP, a seal before FP quiet, or any F change outside its closed path set.
  Evidence arriving while
  the lease is sealed is not
  written by this tool and is eligible only for the separately approved
  post-publication process after `ClosedF`;
- add `Tools/Invoke-TerminalEngineRound5ValidationSuite.ps1` as the sole typed
  validation-suite runner. Its explicit receipt path is
  `.build/terminal-engine-spike/round5-validation/<SuiteId>/
  <40-lowercase-hex-tested-commit>/receipt.v1.json`; callers pass that path and
  never enumerate a suite directory. The coordinator quiet path is its exact
  sibling `post-suite-quiet.v1.json`. It runs only through a `Pipeline` token/
  runner claim, reopens the P-candidate, A-candidate, or D commit/tree and exact
  command/tool digests, publishes one immutable terminal receipt, and then
  yields a distinct coordinator quiet receipt. A crash/timeout is a terminal
  `Failed` receipt for that tested commit. Failed P/A receipts route directly to
  ProtocolClosed and never authorize a second P/A commit. A failed ordinary
  final suite authorizes only the one reserved administrative SuccessorD and
  both new D-bound final suite keys; failure on that administrative D is
  STOP/WIP. Old-D receipts remain immutable ignored history and can never
  satisfy a successor D. BoundaryEvidence embeds every receipt/quiet and any
  ProtocolClosure already present before its D; a later closure is carried by
  the successor-D manifest plus decision/archive while BoundaryEvidence retains
  exact N/A. Thus legality survives ignored-root cleanup without rewriting the
  fixed BoundaryEvidence path;
  for the administrative SuccessorD only, the runner requires either the
  approved ProtocolRepair wrapper and exact complete `newToolIdentities`, whose
  raw digest it records in both suite receipts, or one stage-correct literal:
  `N/A-no-protocol-repair-required` or
  `N/A-no-protocol-repair-permitted`. Either N/A requires the unchanged original
  digest set and frozen allowlist identity. The runner validates the
  N/A against the ProtocolClosure reason/ordinary-D eligibility and rejects
  mixed/extra/unapproved identities before
  starting the command;
- add `Tools/New-TerminalEngineRound5SupersessionProposal.ps1` and
  `Tools/New-TerminalEngineRound5SupersessionReview.ps1` as the only live
  observation writers. The coordinator allocates monotonic
  `observation-<four-digit-ordinal>` directories and explicit create-new
  `proposal.v1.json`, `review-1.v1.json`, `review-2.v1.json`, and `fence.v1.json`
  paths below the ignored supersession root, only for ordinals `0001` through
  `0008` and only after prospective budget admission. A ninth request or a
  smaller budget-driven rejection wins `CapacityClosing` without allocating an
  observation path. The proposal writer recomputes v5/
  evidence identities and excludes future fields. It pre-renders privately;
  only the coordinator may no-replace publish it under lock as the authoritative
  `ObservationPending` transition, mutually exclusive with next pipeline
  admission, and recovery derives state from that file. The review writer enforces
  proposal-author exclusion, distinct reviewers, strict UTC and ordinal
  chronology; Reject closes/resumes only through the coordinator, while the
  second Approve binds the allocated transition ID/old/new epochs, is
  pre-rendered, then publishes only after the authoritative no-replace fence
  journal binds its reserved raw digest. Recovery derives `Superseding` from
  that fence and can finish only the reserved review publication. Neither tool
  discovers a latest observation or overwrites;
- add `Tools/New-TerminalEngineRound5SupersessionObservationLedger.ps1` as the
  sole tracked observation-ledger writer and
  `Tools/New-TerminalEngineRound5Supersession.ps1` as the sole final-result
  writer. After coordinator quiet, the ledger writer consumes an explicit
  generation, exact predecessor ledger path/raw digest/canonical object or the
  first-generation N/A, and ordered list of every live proposal/review path/raw
  digest. It embeds their canonical objects, writes only the explicit absent
  `...ObservationLedger.gNNNN.v1.json` path, and rejects a non-cumulative
  generation, fork, gap, overwrite, or implicit-highest lookup. It writes
  `Empty`, `RejectedOnly`, or `ApprovedSupersession` on every post-v5 closeout,
  not only the Superseded branch. On Superseded, the result writer additionally requires the
  exact approved ledger entry, fence/pre-cleanup-quiet/cleanup-action/post-
  cleanup-quiet token identities, mechanically
  snapshotted milestone set, neutral attempt ledger or validation substitute,
  optional invalidated qualification/proposal/review identities, and exact
  cleanup disposition/proofs. It rejects a live action, non-approved proposal,
  missing observation, future-bound identity, timing/artifact mismatch, retained
  adoption path, implicit discovery, or overwrite;
- add `Tools/New-TerminalEngineRound5ExecutionFreezeAttestation.ps1` as the sole
  post-effect attestation writer. It requires explicit A commit/tree, tracked
  file-set/test identities, and both ignored approval paths/raw digests; it reads
  the commit bytes through Git, embeds/recomputes both approval objects, rejects
	  author/chronology/file-set/A-bound-action-start mismatches or overwrite, and writes
	  only the fixed previously absent tracked path;
- add `Tools/New-TerminalEngineRound5BoundaryEvidence.ps1` as the sole pre-D
  boundary-evidence writer. It receives every explicit P/A/B/C transaction-
  directory manifest/recovery/quiet path and raw digest plus exact timing N/As,
  reopens and validates the predecessor/attempt chains and approval-selected P/A
  attempts, embeds the canonical objects, and no-replace writes only the fixed
  tracked BoundaryEvidence path. It rejects implicit discovery, an omitted
  failed/recovered attempt, D/E/F input, duplicate logical boundary, or a quiet/
  branch/projection mismatch;
- add `Tools/New-TerminalEngineRound5FirstSelectionInputSet.ps1` as the sole
  ignored input-set writer. It requires every explicit subject/artifact/tool/
  dependency path described by the closed schema, opens
  ordinary non-reparse files under held handles, recomputes raw size/digest and
  the digest-free RS-JCS projection, verifies the exact execution-freeze commit/
  clean status and candidate selection, and atomically writes only the explicit
  previously absent `.build` input-set path. It rejects an omitted/extra/duplicate
  leaf, ambient/global fallback, caller-supplied effective digest, changed file
  behind a claimed digest, wrong lane order, or implicit discovery;
- add `Tools/New-TerminalEngineRound5QualificationFailure.ps1` as the sole
  qualification-failure writer. It accepts the exact collection/finalization
  paths and digests, selected candidate, attempt-ledger path/digest, closed
  stage/reason, every later artifact identity, and exactly one timing-specific
  clean-tree or adoption/revert proof. It also requires pre-cleanup quiet and
  either the exact cleanup token/claim/result plus distinct post-cleanup quiet,
  or both exact no-cleanup N/As. It reopens and recomputes every identity,
  rejects a
  non-winner mismatch, stage-illegal artifact, tracked-state residue, runner-up,
  overwrite, implicit path discovery, an illegal attempt-ledger state, or any
  post-effect same-round retry, and
  writes the one closed JCS record only to the explicit previously absent
  `Specs/Terminal/TerminalEngineRound5QualificationFailure.v1.json` path;
- add `Tools/New-TerminalEngineRound5FirstSelectionAttemptReceipt.ps1` as the
  sole coordinator-only live-receipt writer. It accepts the exact input-set
  path/raw digest, five explicit preflight paths/raw digests, action token,
  ordinal `1|2`, and UTC; it reopens/recomputes every subject and derives the
  disposition/reason rather than accepting either. It atomically publishes
  without replacement to the matching explicit receipt path below `.build`.
  Ordinal 2 requires the explicit ordinal-1 path/raw digest, its
  `RejectedPreEffect` disposition, byte-identical input-set path/raw digest/
  identity, revalidation of every leaf, and `attempt2Utc > attempt1Utc`. The tool
  rejects equal/reversed time, later-stage attempt, third attempt, transition/
  winner/epoch mismatch, implicit discovery, or pre-existing output. Only an
  `Admitted` receipt permits dispatch of the first candidate lane;
- add `Tools/New-TerminalEngineRound5FirstSelectionAttemptLedger.ps1` as the
  sole tracked ledger writer. Only after the current action is stopped at its
  admission fence, it accepts the coordinator's exact quiet-receipt path/raw
  digest, one/two attempt paths/digests, and closed neutral sequence state,
  reopens and recomputes them, proves `activeActionCount=0` and strict UTC
  ordering, copies their canonical entries, and atomically publishes the JCS
  record without replacement to the
  explicit previously absent
  `Specs/Terminal/TerminalEngineRound5FirstSelectionAttemptLedger.v1.json`
  path. It rejects a receipt/state mismatch, live candidate action, changed
  retry input, third attempt, illegal transition, winner mismatch, branch/
  clean/revert/result fields, implicit discovery, or pre-existing output. It
  never runs before or during a candidate action and never decides a branch;
- add `Tools/New-TerminalEngineRound5WinnerHandoffProposal.ps1` as the sole
  proposal writer. It consumes only the exact successful non-decision
  first-selection artifacts plus the attempt ledger, requires explicit
  `-ProposalAuthorId` and `-CreatedUtc`, recomputes every identity and the exact
  pre-closeout six-tuple, and writes one JCS proposal to the explicit previously
  absent
  `Specs/Terminal/TerminalEngineRound5WinnerHandoffProposal.v1.json` path. It
  cannot write or name the Done/closeout/success-marker fields, Round-5
  decision, or either later review record. It requires `CreatedUtc` to be
  strictly later than the attempt-ledger freeze and every bound artifact
  completion/approval UTC and rejects equality, reversal, a future referenced
  artifact, or caller backdating;
- add `Tools/New-TerminalEngineRound5WinnerHandoffReview.ps1` as the sole review
  writer. It requires explicit proposal path/digest, reviewer/verdict/UTC,
  content-free rationale, and `-Ordinal 1|2`; ordinal 2 also requires the exact
  ordinal-1 path/digest. It reopens and recomputes the proposal and prior review,
  validates schema, fixed path, strict chronology, non-author/distinct-reviewer
  rules, and the prior `Approve` verdict, rejects implicit discovery or an
  existing output, and writes only the matching previously absent fixed JCS
  review path. A `Reject` record closes review admission immediately. The
  closeout verifier requires both exact approving review records for Winner, or
  the exact first rejecting record for a reviewer rejection. For a proposal/
  review validation failure before its final record exists, the qualification
  record instead requires that artifact's explicit absent disposition and
  content-free failure identity;
- add `Tools/New-TerminalEngineRound5DecisionSubject.ps1`,
  `Tools/New-TerminalEngineRound5DecisionAudit.ps1`,
  `Tools/New-TerminalEngineRound5DecisionGenerationEvidence.ps1`,
  `Tools/Publish-TerminalEngineRound5Decision.ps1` plus
  `Tools/New-TerminalEngineRound5DecisionGenerationDisposition.ps1` as the sole
  decision renderer, audit/evidence writer, final publisher, and CAS-derived
  disposition writer. The coordinator allocates an
  explicit create-new
  `.build/terminal-engine-spike/round5-decision/generation-NNNN/` directory.
  It always allocates the next contiguous ordinal: `0001` through `0009` are
  ordinary and may carry a ProtocolClosure-bound subject once closing has
  started; `0010` is accepted only after complete dispositions for all nine
  predecessors and with the exact ProtocolClosure path/raw digest and reserved
  terminal budget. An administrative SuccessorD does not alter that rule, and
  rejection of generation `0010` admits no `0011`. Every generation passes the same
  prospective budget gate before allocation;
  the two generation-state ControlQuiet logical keys are exactly
  `round5/control/DecisionGenerationEvidence/<NNNN>/.build/terminal-engine-spike/round5-decision/generation-<NNNN>/generation-evidence.v1.json`
  and
  `round5/control/DecisionGenerationDisposition/<NNNN>/.build/terminal-engine-spike/round5-decision/generation-<NNNN>/generation-disposition.v1.json`.
  A mutation-free pure renderer first produces and validates `subject.md` bytes
  in memory as UTF-8 without BOM and performs no filesystem write. Under the
  coordinator lock, it recomputes the next contiguous ordinal and prospective
  budget, then atomically no-replace publishes the normal ControlQuiet token
  whose exact logical key is
  `round5/control/DecisionSubject/<NNNN>/.build/terminal-engine-spike/round5-decision/generation-<NNNN>/subject.md`.
  That token embeds the subject raw digest, exact target/staging rules, and
  budget snapshot; its publication, not an untokened staging write, races
  proposal publication. Proposal first leaves no subject token, official
  generation directory/path reservation, or counter gap. Token first reserves
  that ordinal, temporarily blocks proposal admission, and authorizes the
  subject entry point to reopen the token, create one writer-owned same-parent
  staging directory, write only the validated bytes, and atomically no-replace
  rename it to `generation-NNNN`. A crash after token publication recovers only
  that byte-identical subject before releasing proposal admission; it cannot
  abandon the ordinal. Both evidence/disposition output paths are then reserved
  with that published subject generation and remain no-replace. Every evidence
  form embeds the canonical subject token/raw digest, making token-before-subject
  admission clone-verifiable after cleanup. The subject writer returns its exact
  prefix endpoint/digest. The audit writer writes only the explicit
  `audit-1.v1.json` or, after an approving audit 1, `audit-2.v1.json`, with
  strict author/reviewer/generation/branch/D-commit/UTC validation. Proposal
  publication seals audit admission for a published subject. With zero audits
  or exactly audit 1=`Approve`, the evidence writer publishes
  `ObservationPreempted`, embedding the winning proposal transition; with a
  terminal Reject it publishes `RejectedAudit`; with two approvals it publishes
  `AuditedApproved`. Each embeds the subject-prefix bytes, exact zero/one/two
  audits, prior disposition, and budget identity. Reject then creates disposition
  `Rejected`; only then may retry allocate a caller-specified next generation.
  `ObservationPreempted` creates `Stale`, after which the proposal reviews and
  cumulative ledger resolve before any successor D/new generation. Two
  approvals require `AuditedApproved` evidence
  before the publisher rerenders
  the prefix, embeds both canonical audit objects plus raw digests beneath the
  audit heading, preserves the prefix's already-reviewed `decisionUtc`, excludes
  any self/future digest or unreviewed publication field, and atomically creates
  only the explicit ignored `approved-decision.md` in that generation. Under the
  exclusive `ClosingPublish` lease in step 9, the coordinator byte-revalidates
  those approved bytes. If a proposal wins before approved-decision publication,
  Stale binds exact `N/A-observation-before-approved-decision`; if it wins after
  that publication, Stale embeds the approved-decision canonical bytes/raw
  digest. The authoritative observation-vs-ClosingPublish CAS then
  has exactly two terminal continuations. Observation first leaves E/decision
  absent and the disposition writer creates `Stale` bound to the winning
  proposal transition; then, when PrepareE existed, publishes its immutable
  `stale-before-admission.v1.json`; only after normal observation resolution,
  ledger, and successor D may the next generation allocate. E first
  uses PrepareE's predicted OID to pre-render and budget the exact Accepted
  disposition and bind its staging/final paths/raw digest into the E manifest/
  fence; after atomically creating the tracked decision/ref, the disposition
  writer may publish only those reserved bytes before E quiet and FArchive. Crash recovery
  derives only those states and never edits evidence/disposition bytes. The
  disposition writer accepts the explicit evidence path/raw digest plus exactly
  one of rejected-audit, winning proposal-transition/E-absent, or
  E-commit/proposal-absent inputs. It recomputes the Git ref and fixed paths; a
  both/neither/changed winner, wrong E, or pre-existing divergent disposition is
  STOP. An absent-output retry reuses the reserved logical key and byte-identical
  inputs; it never reruns an audit, proposal, or ref CAS. The
  publisher cannot create the tracked decision earlier. All five consume the
  matching coordinator token/epoch for the
  retained branch; `TransitionRejected` uses the bootstrap governance epoch and
  never fabricates P or A; `ProtocolClosed` consumes the exact reserved
  ProtocolClosure plus its final administrative D/suite identities and may
  render only `round5Outcome=no-winner`;
- all Round-5 governance writers above publish through a same-directory
  writer-owned ignored staging/quarantine path, fully validate the staged bytes,
  then use an atomic no-replace rename to the fixed tracked canonical path. An
  official writer can therefore publish only schema-valid canonical bytes; the
  tracked canonical path remains absent after validation rejection. A process,
  storage, or sharing
  failure that leaves the final path absent is an infrastructure pause, not an
  engine result: keep the plan WIP, preserve diagnostics, repair the
  environment, and rerun only that writer with byte-identical logical inputs,
  UTC, and verdict. Never rerun a candidate action, preflight, or reviewer
  judgment for this case. The writer retry must acquire a new linked
  writer-attempt coordinator token for the same action/input identity and
  bind the prior claim, then recheck the epoch; if an approved supersession
  changed it, do not publish the
  stale artifact and route its exact absent disposition to Superseded. Each writer
  returns its explicit writer-owned sibling temporary path before publication;
  retry may remove only that resolved, ordinary-file temporary after verifying
  it is in the expected directory and the final path is absent. An unowned or
  reparse temporary is STOP. Schema/path/digest-invalid staged output is a
  terminal stage validation failure—`CandidateLane`,
  `FirstSelectionAttemptLedger`, `WinnerHandoffProposal`, or
  `WinnerHandoffProposalApproval` where applicable. The ignored quarantine bytes
  remain immutable; Qualification-failed or Superseded embeds a schema-valid
  content-free `invalidArtifactEvidence` object containing the logical target,
  quarantine path/raw digest/size, token/claim, failed rule IDs, and exact
  `N/A-validation-rejected-before-publication`, while the tracked canonical path
  remains absent. A pre-existing file at a fixed tracked canonical output,
  whether valid, partial, or invalid, is an ownership collision and STOP/WIP:
  never overwrite, quarantine, delete, commit, or reinterpret it as a branch
  result. The closeout verifier rejects an invalid canonical file and accepts an
  invalid-artifact projection only when its target is absent and its canonical
  evidence is embedded in the tracked valid result.
  Supersession before the first receipt records
  `N/A-superseded-before-attempt-ledger`; supersession after one rejected
  receipt freezes the `StoppedAfterFirstRejection` ledger;
- before v5 freezes, harden the evaluator snapshot required by the current
  decision. Add every review-referenced source archive, each ancestor handle,
  and the exact authenticated
  `Tools/TerminalEngine/TerminalEngineSourceReview.psm1` bytes to the same held
  no-write/no-delete, no-reparse snapshot. Keep those handles live through
  archive hashing, member lookup/extraction, line-range verification, source
  decision, and governance consumption. A hash of a released pathname followed
  by a reopen is forbidden;
- add `Tools/New-TerminalEngineCandidateGovernanceFreeze.ps1` as the sole v5
  Review/Assessment/Set bundle writer. It pre-renders and validates all three
  fixed subject bytes, allocates their common freeze ID/UTC, and under the
  P-epoch Pipeline token atomically no-replace publishes the freeze journal
  containing their reserved raw digests. Persistent state derives from that
  journal; recovery may publish only those three reserved files, in fixed order,
  before the action quiets. Supersession can win before the journal or fence the
  already-admitted bundle action afterward, but cannot expose/accept a partial
  bundle. The tool rejects an existing output, changed reserved digest,
  sibling/journal hash cycle, implicit discovery, or overwrite;
- update `Tools/Evaluate-TerminalEngines.ps1`,
  `Tools/TerminalEngineEvaluator.psm1`,
  `Tools/TerminalEngine/TerminalEngineCandidateAssessment.psm1`, and
  `Tools/TerminalEngine/Invoke-TerminalEngineCollectionProvider.ps1` so
  `UniverseFile`, `CandidateReviewFile`, `CandidateAssessmentFile`, and
  `CandidateSetFile` are mandatory explicit leaf paths in both direct and nested
  entry points and their raw SHA-256 values flow into every descriptor,
  collection, finalization, materialization, and verification record. The
  explicit GovernanceFreeze path/raw digest is also mandatory and must
  authenticate that exact three-subject bundle;
- before any descriptor-driven bootstrap, clean, process, build, or evidence
  effect, make the provider validate the exact closed descriptor from that
  authenticated governance snapshot and then run a second neutrality proof.
  That proof binds the same Universe/Review/Assessment/Set raw digests and scans
  the complete ordinary neutrality surface plus every normalized
  descriptor-only dependency, package, artifact, runtime-leaf, tool, driver,
  and output literal. Descriptor validation is data-only; no descriptor-named
  module or command executes before the second proof succeeds. A differing
  governance digest, archive/source-review handle, descriptor digest/literal
  set, or neutrality result rejects before effects;
- update `Tools/Build-TerminalEngine.ps1`,
  `External/TerminalEngine/TerminalEngine.proj`, the collection descriptor/
  driver modules, and the Gate-0 Harness/Probe wrappers only as required for the
  frozen generic Cargo/Rust build kind and synchronous ABI actions. Candidate IDs,
  pins, package names, dependency names, and expected results remain in frozen
  subjects/descriptors, never in candidate-neutral host code; and
- add `Tools/Tests/TerminalEngineRound5Proof.Tests.ps1` and
  `Tools/Tests/TerminalEngineRound5Closeout.Tests.ps1`; extend the existing
  universe-transition, candidate-patch, assessment, evaluator, provider,
  descriptor/driver, Harness/Probe, network-isolation, evidence-schema, and
  archive tests. Add barriers that swap/reparse every review archive,
  source-review module, and ancestor before/during hash and extraction; mutate
  a descriptor or its literal set between validation and the second neutrality
  proof; inject candidate/dependency/output literals present only in the
  descriptor; and prove every case rejects before directory creation, clean,
  process start, or evidence mutation. Add
  `.github/workflows/terminal-engine-round5.yml` for the
  focused proof/closeout contracts and extend
  `.github/workflows/terminal-engine-gate0.yml` to take the four explicit v5
  subject paths plus GovernanceFreeze path in the three real native collection
  lanes. Neither workflow may
  download an unpinned input during an offline evidence phase.

The v5 universe/candidate set freezes, per platform, every logical bootstrap
input with exact leaf name, retrieval origin, version/pin, SHA-256, size, archive
tree identity, consumers, and license identity: the WezTerm source archive; the
complete Cargo vendor/dependency closure; the exact Rust compiler/Cargo/rust-std
components for `x86_64-pc-windows-msvc` and `aarch64-pc-windows-msvc`; any
candidate build helper; and the single candidate-neutral
`host-wil-headers.zip` consumed by the complete executable cohort. Materialize
them only under a digest-keyed `.build\terminal-engine-spike\inputs\` root.
Source/toolchain/package bytes and extracted trees never enter `Specs/TestRuns`;
only schema-allowed content-free identities and aggregate results do. Native
MSVC/SDK environment identities remain lane attestations, not ambient substitutes
for a missing frozen candidate input.

## Fifteen-owner-day proof-spike fence

Before proof work, freeze an ordinal work ledger with named owners and one-day
increments. Charge every design, source, generated bridge, build, test, fuzz,
benchmark, dependency/license review, documentation, and evidence task needed to
decide the synchronous hypothesis. Do not discount parallel work, generated files, abandoned
approaches, or work performed outside the repository. The experiment owner
publishes the cumulative ledger after each owner-day.

The arbitrator is monotonic and one-shot:

1. `PendingProof` becomes `ProofSucceeded` only when the exact synchronous
   boundary hypothesis and its required x64/ARM64 executable proof exist while
   cumulative proof use is at most 15 owner-days. `ProofSucceeded` authorizes
   completion of the v5 nomination, not admission, collection, or selection.
2. `PendingProof` becomes `BudgetExceeded` before starting work that would require day
   16, or immediately when a conservative lower bound proves more than 15.
3. Any hard technical stop below moves `PendingProof` to `Rejected` immediately.
4. `ProofSucceeded|BudgetExceeded|Rejected` is final for this exact proof pin and
   hypothesis. A failed proof fix, new pin, or changed boundary requires a later
   versioned Gate-0 round; the proof time box is never extended.

`BudgetExceeded` and `Rejected` are valid no-winner evidence, never permission to
finish the proof off-ledger. After `ProofSucceeded`, finish and freeze the normal
nomination without a second 15-day clock. Every retained spike path and owner-day
remains in the candidate patch/result-tree and owner ledger as applicable; there
is no prototype exclusion or duplicate credit. The proof ceiling does not replace
the frozen candidate admission caps: at most 60 three-year owner-days by the master formula,
2,500 additions-plus-deletions, 24 touched files, five closed logical concerns,
one product-new toolchain family, zero non-system runtime leaves in a clean
package, and no binary patch. Exact 60 owner-days is allowed. The complete
nomination must satisfy all caps independently before its assessment may mark it
`collect`.

## Synchronous caller-owned proof

The proof schema and corpus freeze this exact seven-action C ABI table. The
candidate may change spellings only by changing and reapproving the pre-result
v5 universe; it may not collapse actions into an unreviewed generic call.

| Action | Exact ownership and synchronous result |
|---|---|
| `Create` | Consumes copied policy plus caller allocator callbacks for the duration of the call and returns one non-refcounted opaque unique handle. The handle may own one Rust value through `Box`/equivalent unique storage; `Arc`, `Rc`, `Weak`, shared session ownership, global/TLS registration, and a retained caller pointer are forbidden. |
| `Destroy` | Consumes and nulls that unique handle synchronously. Before return all engine work, tasks, callbacks, destructors, borrows, and engine-owned temporary storage are gone; there is no deferred destroy or post-return callback. |
| `Feed` | Consumes one bounded copied byte/event batch, mutates only during the call, and returns a caller-owned deep-copied effects package or a closed failure with no partial publication. |
| `Resize` | Consumes validated cells/pixels synchronously and returns a caller-owned deep-copied resize-effects package or a closed no-publication failure. |
| `Snapshot` | Returns a complete caller-owned immutable snapshot or a closed retry/full-snapshot state. Dirty consumption is committed only with successful complete publication; no cell/string/hyperlink/image/placement borrow survives. |
| `EncodeInput` | Converts one bounded typed input event into caller-owned bytes without mutating the engine or retaining the event/buffer. Invalid enum/version/length is a closed typed failure. |
| `QueryKitty` | Returns caller-owned deep copies of bounded Kitty capabilities, image/placement state, and required response/effect bytes only after pre-allocation accounting succeeds. |

Every action has one ABI/version header, explicit byte lengths, caller-provided
output capacity, required-size reporting with no partial write, and a closed
typed status. There is no generic callback field, future/task handle, borrowed
slice in an output, action overlap, reentrancy, or implicit allocator/CRT
ownership. The sole opaque engine handle is unique storage, not an exception to
the ban on `Arc`-retained session state.

Prove all of the following with source review plus executable probes on the exact
pin:

- no borrowed cell, string, hyperlink, image, placement, allocator, panic,
  callback, task, future, thread, TLS destructor, `Arc`, `Rc`, `Weak`, or shared
  Rust session is reachable after return; only the unique non-refcounted Rust
  value owned by the opaque `Terminal.dll` handle may persist between calls;
- snapshot begin/end, dirty consumption, failure, cancellation, and full-retry
  semantics are exact; copied cells/effects/static Kitty images remain valid
  after subsequent mutation and engine destruction;
- reentrant calls and callbacks are rejected by construction; panic/unwind never
  crosses the ABI; allocation failure and every enum/length/version mismatch have
  closed typed results;
- Kitty compressed/decoded/transient/storage bytes, pixel dimensions,
  generations, placements, crop/scale/z-order/deletion, and aggregate terminal
  memory can be reserved and rejected before untrusted-input allocation;
- public headers, every crossed by-value layout, exports, compiler/Rust target,
  CRT, allocator, feature flags, dependencies, licenses, and PE/import closure are
  fingerprinted before handle/callback creation; and
- native x64 and ARM64 Windows builds plus the x64 ASan consumer boundary are
  possible offline from digest-locked inputs with no ambient Cargo registry,
  rustup state, global tool, network, or app-root runtime leaf.

Hard-stop immediately if any proof needs an asynchronous engine callback, a
borrowed/private lifetime that cannot be machine-checked, a second parser, a
candidate-owned renderer/control, post-return engine work, panic/unwind across
the ABI, unmetered pre-allocation, a non-system runtime package leaf, an
unfingerprintable layout, an unavailable native ARM64 lane, or a changed product
requirement. Also stop when the completed v5 design already contains, or the
proof makes unavoidable, a sixth independently regressible logical concern.
The measured v3 async/`Arc` baseline begins with six concerns and is charged as
retained proof work, but its transient presence is not itself a new v5 stop:
success requires demonstrably deleting that lifecycle concern from the
completed v5 design rather than relabeling it. Record the smallest exact reproducer and do not spend the remaining
experiment budget trying to route around the failed invariant.

## Execution sequence and immutable chronology

1. **Authenticate Round 4.** Resolve the exact Done path, prove its Git blob at
   the recorded closeout commit, hash the current Done bytes, and verify the
   recorded Round-4 decision digest against the exact decision blob read from
   that same closeout commit, not against the mutable current operator summary.
   Confirm its outcome is no winner and that it contains no
   `successfulGate0Handoff` marker.
2. **Prepare the candidate-neutral machinery and carry audit before any v5
   universe.** Implement and corpus-test the held-archive/source-review snapshot,
   post-descriptor neutrality proof, transition/carry/supersession/proof/v5
   governance schemas, coordinator, and explicit-path readers first. Initialize
   a create-new bootstrap governance epoch bound to the authenticated predecessor,
   current clean HEAD/tree/status, and the exact coordinator/writer/tool digests;
   it admits only carry/universe governance and transition-closeout writers, not
   proof or candidate effects. Byte-reverify the
   complete non-WezTerm seed/control inventory and outcomes into the carry
   audit. Obtain two approvals. `TransitionRejected` jumps directly to the
   dedicated closeout branch with zero v5/proof/provider effects. Only
   `Unchanged` continues.
3. **Freeze only the v5 universe before results.** Freeze the exact WezTerm pin
   and source archive, synchronous design hypothesis, proof protocol and ledger,
   unchanged requirements/caps/gates/fixtures/measurement inputs, candidate-
   neutral accounting rules, and initial bootstrap/toolchain identities in the
   v5 universe. Two non-author reviewers approve that exact universe. Do not yet
   create or approve CandidateReview, CandidateAssessment, CandidateSet, or the
   Round-5 decision. As the bootstrap epoch's final pipeline actions, create the
   P-candidate commit only through the `P` Git-boundary Pipeline transaction and
   its post-transaction quiet. Its exact commit/tree contains all candidate-
   neutral machinery, authenticated carry audit, v5 universe, proof inputs, and
   no proof result. Then admit exactly
   `PreProofFocused` against that commit/tree; its terminal receipt must be
   `Pass`, and its distinct post-suite quiet receipt plus canonical receipt/raw
   digest are inputs to both P approvals. A failed/crashed/timed-out suite makes
   that P candidate permanently unapprovable and selects
   `ValidationClosing/preproof-focused-failed`; no same-round repair, new P
   candidate, or receipt replay is legal. The reserved ProtocolClosed route
   repairs only the closeout surface and carries any candidate/protocol repair
   into the successor ordinal's newly reviewed universe. Two distinct non-
   author approvals bind P's commit/tree, v5-universe digest, exact tracked file
   set, P transaction manifest/post-transaction quiet, PreProofFocused receipt,
   and its suite quiet receipt. Only after the second
   approval publication stops, emit the bootstrap quiet receipt. Under
   the coordinator lock, rotate to a pre-A P epoch using those exact identities.
   P permits the one bounded proof action, proof-result and
   post-proof nomination/governance writers, live supersession control
   operations, and quiet; no collection or first-selection action is admitted.
4. **Run the bounded proof below `.build` through the P epoch.** Keep source, binaries, reproducers,
   and terminal data out of `Specs/TestRuns`. Archive only already-allowed
   content-free identities/results. The hard-stop arbitrator owns the exact
   result. The proof action consumes the P-bound coordinator token. The proof
   record binds P commit/tree/approvals, the pre-proof supersession-checkpoint
   UTC, and confirms no approved supersession existed at admission. After the
   action/result writer stops, emit its action-stopped receipt but keep the P
   epoch active for step 5's serialized governance writers.
   `ProofSucceeded`
   permits the next step only; it is not `collect`.
5. **Complete and freeze the post-result nomination subjects.** After
   `ProofSucceeded`, complete the exact patch/result tree, concern map,
   descriptor, dependency closure, and candidate owner ledger under the normal
   frozen caps. Every nomination/governance writer consumes a P-epoch pipeline-
   action token and rechecks the epoch before publication. For any proof
   outcome, reverify every carried seed/control and pre-render the v5
   CandidateReview, CandidateAssessment, and CandidateSet. Obtain every fresh
   source/adaptation/admission approval over those reserved raw digests strictly
   after the latest v5-universe approval. Through the sole bundle writer,
   no-replace publish the authoritative GovernanceFreeze journal and then only
   the three reserved subjects with one freeze ID/UTC strictly after the
   approvals. The bundle Pipeline action cannot quiet until all three exact
   files exist; a concurrent supersession sees the complete bundle or none.
   Do not write or approve the Round-5 decision
   in this governance-freeze step. It is written and approved only after the
   selected terminal closeout branch has complete immutable evidence.
   Equality, an anticipatory approval, a later edit, mixed record versions, or a
   digest that does not bind the raw subject bytes invalidates the round before
   collection.
   For `BudgetExceeded|Rejected`, use the closed `proof-rejected` Review,
   Assessment, and Set branches; all nomination artifacts are explicit N/A and
   the admitted count is zero. Keep the P epoch active for the A commit and
   approval actions in step 6.
6. **Freeze the executable baseline, then close or open admission.** Finish all
   candidate-neutral schemas/tools/corpora/workflows and every tracked v5 proof/
   nomination/governance subject, then create the A-candidate commit only through
   the `A` Git-boundary Pipeline transaction and its post-transaction quiet. Run
   exactly `PreResultFocused` against that commit/tree through its own Pipeline
   token/claim; require `Pass` and a distinct post-suite quiet before either A
   approval. Both approvals bind the canonical receipt/raw digest and quiet
   receipt. A failed/crashed/timed-out suite makes A permanently unapproved and
   selects `ValidationClosing/preresult-focused-failed`; it cannot mutate the
   already-observed proof/governance inputs or create a second A. Record
   A's exact
   40-lowercase-hex commit and tree; require `HEAD` to equal it and
   `git status --porcelain=v1 --untracked-files=all` to be empty (ignored
   coordinator/input/evidence roots only). Two distinct non-author ignored
   approval records bind that commit/tree, exact tracked file-set identity,
   A transaction manifest/post-transaction quiet, `PreResultFocused` receipt,
   and its suite quiet receipt. Every collection manifest,
   finalization, first-selection input
   set/receipt, and lock-independent lane binds the same commit/tree/approvals.
   After A, the only tracked write permitted before candidate effects stop is
   the exact reviewed B adoption and, when terminal cleanup requires it, its
   exact C revert. All branch-dependent evidence/governance/documents remain
   untracked and unstaged; live coordinator/evidence bytes remain ignored.

   Only after the second A approval publication stops, emit the P-epoch quiet
   receipt. From that quiet P epoch, rotate the sole coordinator under lock to a new
   A-bound epoch against that freeze; never initialize a disconnected second
   coordinator or discard the P token/quiet/observation history. If an ignored
   supersession proposal obtains both approving reviews, execute the atomic
   fence protocol—not a future final result: stop new actions, complete/cancel
   the one pre-fence action, emit pre-cleanup quiet, snapshot every action
   admitted before the fence (including any result that completes afterward),
   freeze the neutral attempt ledger when applicable, and freeze the observation
   ledger. If cleanup is required, admit its sole terminal-epoch Pipeline action
   and emit a distinct post-cleanup quiet; otherwise bind the exact cleanup/
   post-quiet N/As. Then create the immutable Supersession result and use
   the dedicated branch. Only actions admitted before the fence may appear as
   completed/aborted invalidated evidence; none admitted afterward is legal.
   Never replace frozen Assessment/Set bytes.

   Otherwise, if the proof is not `ProofSucceeded` or the completed nomination
   is not admitted, select ordinary no-winner. `Collect` and the direct provider
   reject before attestation/bootstrap/clean/subprocess/output mutation; create
   no finalization, lock, materialization, attempt receipt, or native winner
   lane. After coordinator closeout quiet, create the explicit empty/rejected-
   observation ledger. Write the Round-5 decision only in step 9 after terminal
   branch evidence is complete.
7. **If admitted, run unchanged Gate 0 through the coordinator.** Serialize the
   native x64 Release, x64 `ASan Debug`, and ARM64 Release complete-cohort
   collection actions; each consumes an exact action token, the four explicit v5
   subject paths, GovernanceFreeze path/raw digest, execution-freeze
   commit/tree/approvals,
   `-Candidate All -Bootstrap -Clean -Offline`, and host-wide outbound-network
   denial. The coordinator admits the next action only after committing or
   invalidating the prior result. Finalize only from those three named manifests
   under its own token. Every manifest/finalization binds the execution-freeze
   identities. The cohort is inclusive and candidate-neutral: no WezTerm-only
   collection, concurrent official lane, narrowed diagnostic, candidate order,
   or preference may select a winner.
8. **If and only if Finalize names a mandatory-pass winner at or above 80.00,**
   have the coordinator build/recompute the complete input set, reserve/preflight
   all five native runner tuples, and derive the first admission receipt before
   candidate effects. If rejected, it must run the one identical-input retry or
   remain WIP; only an approved supersession may stop admission after the first
   rejection. A second rejection is terminal. No other stage is pre-effect or
   retryable.

   Transfer the byte-identical admitted receipt/input set and raw digests to
   every native runner. Serialize the five lock-independent candidate lanes;
   every invocation and returned record consumes/binds the action token,
   admitted receipt, input identity, and execution-freeze commit. Then create
   immutable materialization and two approvals, non-overwriting candidate lock
   and two approvals. The reviewed tracked adoption commit is the `B` Git-
   boundary transaction and contains only the approved product/runtime/
   production-lock/notices changes and support surface
   already present in the execution freeze. It explicitly excludes attempt/
   observation ledgers, qualification/supersession records, handoff proposal/
   reviews, Round-5 decision, plan closeout, and other branch-dependent docs.
   This Round-5 rule supersedes the legacy Plan-1 sentence that placed decision/
   ledger/tests/documentation in adoption; reuse its build/materialization/lock
   commands, not that obsolete adoption-content ordering.

   From a clean adoption commit, serialize five lock-bound Build/Verify pairs,
   verification-set creation, dependency status/cache checks, and notices/
   license closure through coordinator tokens. After the last candidate action
   stops, obtain the quiet receipt and freeze the tracked branch-neutral attempt
   ledger from the exact live admission proof. It cannot stale an already
   completed candidate lane. Then create the immutable winner-handoff proposal
   and two separately immutable ordered reviews. `proposalCreatedUtc` must be
   strictly later than ledger `frozenUtc` and every bound artifact completion/
   approval UTC. The proposal binds the ledger and exact pre-closeout six-tuple,
   is distinct from the branch-specific decision in step 9, and contains no
   review identity; review records point toward it, never vice versa. Failure to
   validate/create it is `WinnerHandoffProposal`; two `Approve` verdicts permit
   Winner; reviewer `Reject` or review validation failure is
   `WinnerHandoffProposalApproval`.

   A second pre-effect rejection or any substantive post-admission failure
   requires the Round-5 qualification-failure record. Stop and emit the
   pre-cleanup quiet receipt first; create the neutral attempt ledger or its
   exact validation substitute. If cleanup is required, keep every new closeout-
   governance file untracked and unstaged, admit exactly one
   `TransactionalCleanup` Pipeline action, stage only the exact adoption path
   set when reverting, prove its
   adoption-path projection equals the pre-adoption execution-freeze tree and
   zero engine-adoption state, require full-tree equality when no governance
   commit intervened, then emit a distinct post-cleanup quiet receipt. If no
   cleanup action is required, bind both exact cleanup/post-quiet N/As. Only then
   freeze the observation ledger and qualification record unchanged for the
   post-effect evidence commit. Never
   try the runner-up or patch in place. A writer-only absent-output pause follows
   the atomic-publication rule and reacquires the coordinator epoch before
   retry; supersession wins any changed epoch and records the absent artifact.
9. **Commit evidence, audit the decision, and close.** After the chosen branch
   has all exact post-effect records, enter the coordinator's `Closing` lease:
   require zero active action and no pending observation; if A existed, create
   the tracked execution-freeze attestation from its two ignored approvals;
   create the tracked BoundaryEvidence record from every applicable P/A/B/C
   transaction/recovery/quiet identity and exact timing N/As;
   bind the latest schema-valid cumulative branch-neutral observation-ledger
   generation only when it already covers exactly every admitted observation;
   create generation `0001/Empty` only when no ledger and no observation exists,
   and never create an extra closeout-only successor to a nonempty ledger
   (`Empty|RejectedOnly|ApprovedSupersession`); and commit every branch's
   unchanged closeout evidence as the `D` Git-boundary transaction. After its
   post-transaction quiet, run `FinalFocused` and then `FinalCanonicalFull` as
   two ordered Pipeline actions from that exact clean commit/tree. Each consumes
   its D-bound token/claim, produces its one terminal receipt, and is followed by
   a distinct quiet receipt before the next allowed key. `FinalFocused` must be
   `Pass`; `FinalCanonicalFull` must be `Pass` or the sole fully evidenced
   `ClassifiedPreExisting` form. The fixed
   `FinalCanonicalFull/post-suite-quiet.v1.json` is the sole final-suite quiet;
   do not emit a second receipt/key. A Failed receipt chooses
   `ValidationClosing/final-validation-failed`, freezes or invalidates the
   prospective branch according to the ProtocolClosure disposition table, and
   admits exactly one administrative SuccessorD. That boundary preserves the
   failed receipt/quiet. When allowlisted bytes change, it consumes the exact
   pre-frozen ProtocolRepair manifest plus both approvals and changes only the
   manifest's candidate-neutral closeout-tool/test bytes. When no byte changes,
   it binds exact `N/A-no-protocol-repair-required`, allocates no repair
   subject/approval/finalizer key, and has zero repair approvals. In both cases
   it binds the exact pre-frozen repair-allowlist path/raw digest and proves all
   gates,
   expectations, schemas/corpora, prior evidence, and excluded paths unchanged,
   proves zero tracked engine state when a
   prospective Winner is invalidated, and allocates new commit-bound keys for
   both final suites. Both must then satisfy the normal Pass/classification
   rule; another failure is the explicit STOP/WIP recovery condition. No
   decision/audit/ClosingPublish operation is admissible until
   the decision input set contains all applicable P/A receipts plus these two
   final-D receipts and quiet identities. No decision or Done move is present
   yet.

   Render the branch decision subject below `.build` as exact UTF-8 without BOM.
   Its reviewed prefix ends with the final LF immediately before the literal
   `### Decision closeout audits` heading. Two ignored immutable decision-audit
   records bind the prefix SHA-256, subject generation, branch/evidence-commit
   identity, BoundaryEvidence path/raw digest, final D/SuccessorD transaction
   manifest and post-transaction quiet canonical objects/raw digests, reviewer
   ID, `Approve|Reject`, rationale digest, and UTC. The prefix
   contains its final `decisionUtc`; audit 1 is strictly later than that UTC and
   subject render, reviewers differ from the decision author and each other, and
   audit 2 is strictly later than an approving audit 1.
   Any Reject leaves WIP and requires a new create-new subject generation; it
   never overwrites an audit. After two approvals, the sole decision writer
   rerenders/recomputes the prefix, appends exactly those audit entries, and
   atomically creates only that generation's ignored `approved-decision.md`.
   The approved bytes contain no own digest; `round5DecisionAudits` binds both
   audit paths/raw digests and the common prefix digest. The fixed tracked
   decision path remains absent at this point.

   Before tracked publication, an observation-admission CAS may win, abort
   `Closing`, and enter the
   two-review protocol. `Reject` or `Approve,Reject` creates the next cumulative
   observation-ledger generation, preserves the original branch, produces a
   `SuccessorD` Git-boundary transaction, reruns both final suites through new
   commit-bound logical keys/receipts/quiet, and starts a new decision-subject
   generation. `Approve,Approve` performs the Superseding fence/quiet/cleanup,
   creates the next cumulative ledger and—unless the branch already has its
   fixed Supersession result—the result, changes to or preserves Superseded,
   then produces a successor D, reruns both suites, and starts a new decision
   audit. A proposal admitted while either final suite runs prevents the next
   suite/decision key; whether rejected or approved, the old-D receipts remain
   ignored immutable history and cannot satisfy the successor D. A later
   approval after an already-frozen Supersession result is carried to the next
   ordinal universe and never replaces that result. Every earlier D, ledger
   generation, and suite receipt remains immutable history.

   After two approving audits for the final D/generation, admit the
   `PrepareE` ControlQuiet key. It allocates that decision generation's caller-
   supplied E transaction ID and creates exactly one
   create-new E Git-boundary manifest at
   `.build/terminal-engine-spike/round5-git-boundary/E/<transaction-id>/
   manifest.v1.json`. This is the sole ClosingPublish publication manifest, not
   a second recovery authority. It binds the final D/SuccessorD boundary-
   manifest path/raw digest unconditionally, plus the prior E-attempt manifest
   path/raw digest or exact `N/A-first-E-attempt`, decision-subject generation,
   its `AuditedApproved` generation-evidence path/raw digest and reserved absent
   `generation-disposition.v1.json` path, plus the explicit E-transaction
   `accepted-generation-disposition.v1.json` staging path, exact pre-rendered
   canonical object/raw digest/size, and its mandatory-payload budget snapshot,
   final D and cumulative
   ledger generation, all four applicable validation receipt/quiet identities,
   approved-decision digest, exact symbolic branch/ref, expected old OID D,
   every fixed decision/cutoff path's expected-before raw digest/blob or absent
   state, expected-after raw digest/mode, intended tree, fixed
   commit message, author/committer identities/UTCs, and predicted E commit OID.
   Pre-render every E after-file and the Accepted disposition—which binds the
   manifest's predicted E OID—below that ignored transaction root, materialize
   and verify the intended tree in its private index, run the mutation-free
   verifier against that materialized tree, and create at most the predicted
   commit object. `PrepareE` has no ref/publication authority. The real fixed
   decision path remains absent and the checked-out ref/index/worktree remain at
   D.

   Acquire the coordinator lock and perform one CAS from `Closing` to exclusive
   `ClosingPublish` by atomically no-replace publishing
   `<E-transaction-directory>/closing-publish-fence.v1.json`. That fence binds
   the same manifest path/raw digest, accepted-disposition staging/final paths
   and raw digest, and a distinct ClosingPublish control token;
   its presence is the persistent publication authority. The predicate
   requires: symbolic `HEAD` is the manifest's exact expected branch/ref; that
   ref and `HEAD` both equal final D; the real index and worktree equal D with no
   staged/unstaged/unrelated tracked delta and no untracked/ignored entry outside
   the manifest's closed allowed ignored roots; all final-D suite receipts and
   final-suite quiet validate; zero active Pipeline action; and no pending, new,
   or unledgered observation. If an observation proposal is no-replace published
   after manifest render but before this CAS, the observation wins: leave the
   fixed decision path absent, retain the stale manifest/root immutably, produce
   the create-new `Stale` generation disposition bound to that winning
    observation, publish the prepared transaction's create-new
    `stale-before-admission.v1.json`, and only then run normal observation
    review/ledger resolution and produce the reviewed successor D/suites/
    decision generation before allocating a new E transaction path. Never
    overwrite or reuse stale bytes.

   Once `ClosingPublish` wins, observation admission is impossible. Revalidate
   the manifest and move only the expected branch ref with
   `git update-ref <expected-ref> <predicted-E-oid> <final-D-oid>` or equivalent
   atomic compare-and-swap, then synchronize the checked-out real index/worktree
   to the exact E tree. Before E quiet/ClosedE, no-replace publish only the
   pre-rendered, budgeted `Accepted` generation disposition at the reserved
   final path; reopen/recompute both staging and final bytes and require the
   predicted E OID. Crash recovery derives and completes only that exact object
   from ref E. That tree no-replace publishes
   `Specs/Terminal/TerminalEngineDecision.Round5.md` and its current
   `TerminalEngineDecision.md` cutoff/runbook addendum only. It leaves this WIP
   plan byte-identical, leaves every master handoff token/routing byte unchanged,
   creates no Done path, and contains no successful-handoff marker. Recovery
   derives exact state from manifest/object/ref/index/worktree: commit object
   present with ref D resumes the same ref CAS; ref E with stale coordinator
   state or unsynchronized files resumes synchronization/verification; ref
   outside `{D,E}`, a third path/index state, alternate object/tree/metadata, or
   unrelated delta is STOP. It cannot admit evidence, choose another branch,
   start another audit generation, or create a second commit.

   `ClosedE` is derived from the exact ref/HEAD E commit, parent/tree/metadata,
   synchronized clean index/worktree, decision raw digest, byte-identical WIP
   plan/master routing, absent Done path and success marker, and
   canonical Accepted generation disposition plus coordinator state. If the ref moved before the ignored state update, recovery
   recognizes E rather than creating another commit. Thus an observation either
   wins before the ClosingPublish fence while the fixed decision path is absent,
   or waits outside Round 5 until after ClosedF; there is no intermediate race.
   Every decision-generation
   and E-transaction root is create-new. It binds the explicit prior decision/E-
   attempt generation or exact first-attempt N/A separately from its mandatory
   final-D boundary predecessor, is never discovered as “latest,” and is
   retained through accepted FPublish. FArchive archives the content-free
   manifest chain/receipt identities before the explicit allowlisted cleanup
   tool may remove those ignored roots; cleanup failure leaves diagnostics in
   place and never changes the E decision.

   E binds only the final explicit D and final cumulative ledger generation.
   The accepted E commit is the Round-5 decision cutoff, not the Done closeout
   or successful handoff. The same sealed coordinator lease must recover through
   FA and FS; only FS is externally publishable and admits the separate
   post-publication observation process. Run the mutation-free closeout verifier
   at each clean boundary before its ref CAS.

The required Git boundaries are:

| Boundary | Required content and clean-state rule |
|---|---|
| P — pre-A proof/governance freeze | Candidate-neutral machinery, authenticated carry audit, approved v5 universe, exact proof inputs, and no proof result; exact clean commit/tree plus two approvals. The pre-A coordinator epoch binds P and serializes proof, every post-proof nomination/governance writer, and A commit/approval publication; its final quiet receipt is required before rotation to the A-bound epoch. N/A only for `TransitionRejected` or supersession before P. |
| A — execution freeze | All candidate-neutral machinery and frozen v5 proof/nomination/governance subjects; exact commit/tree plus two approvals. Required before any executable collection; may be N/A only when an earlier terminal branch prevents execution. |
| B — tracked adoption | Winner product/runtime/lock/notices path set only, with A as parent and a clean resulting tree. No branch-dependent closeout evidence. |
| C — adoption revert | Qualification/Superseded only when B exists; stage/commit only B's exact adoption path set and restore its projection to A with zero engine state. Require full-tree equality to A when no governance commit intervened; otherwise allow only the closed digest-bound immutable governance delta and never delete it. |
| D — post-effect evidence | Every applicable execution-freeze attestation, cumulative attempt/observation ledger, proposal/review, qualification or supersession result, archived content-free evidence and exact test inputs for the chosen branch. Commit only after actions/cleanup stop; run final suites clean at D. A reviewed pre-E observation creates an immutable successor D; E binds the final explicit D. |
| E — audited decision cutoff | Audited immutable Round-5 decision plus current decision-pointer/runbook addendum only. It seals observation admission but leaves this WIP plan, Done path, completion tokens/marker, master handoff tokens, README, and next-plan routing unchanged. Accept only after the clean mutation-free verifier. |
| F — closeout and handoff publication | Four acyclic mandatory subtransactions under E's still-sealed lease. FArchive adds only the tracked RecoveryArchive from E. FDone follows FA and alone fills/moves this plan WIP-to-Done plus its branch/success marker without touching master routing. FPublish follows FD and records the now-known FD closeout commit/Done blob plus FA/path/digest in master routing; it appends the exact pending-carry next WIP round on no-winner or publishes the sole winner pointer. FSeal follows verified FP quiet and adds only the tracked embedded-quiet CloseoutSeal. E, FA, FD, and FP are non-publishable intermediate refs; FS alone is the durable handoff. |

P/A/B/C/D, every successor D, E, FArchive, FDone, FPublish, and FSeal are committed only
by the generic Git-boundary transaction. E/FArchive/FDone/FPublish/FSeal use the same private-
index/tree/commit-object/ref-CAS recovery primitive under their stronger
`ClosingPublish|ArchivePublish|DonePublish|HandoffPublish|SealPublish`
rules. For each boundary, the
manifest and post-transaction quiet identity are mandatory evidence; a crash
after tree/commit-object creation, ref movement, worktree/index synchronization,
or before coordinator acceptance resumes the same deterministic transaction.
It never marks a boundary failed merely because the commit may already exist and
never allocates a second commit. A ref not equal to the manifest old/new OID, a
dirty real index/worktree outside the exact manifest transition, or a coordinator
state that disagrees with the Git-derived state is STOP. D and SuccessorD also
bind the exact cumulative observation-ledger generation; a proposal CAS that
wins before their admission changes the required D input, while a proposal
admitted after their Pipeline token may let that boundary finish but blocks the
next suite and forces the reviewed successor-D path.

These ignored manifests are recovery journals, but every pre-E boundary has a
named clone-durable consumer. P's canonical manifest/raw digest and quiet are
embedded through its approvals into the proof, execution-freeze attestation, or
Supersession result; A's are embedded in the execution-freeze attestation; B/C's
are embedded in the Winner decision or Qualification/Supersession result; and
the final D/SuccessorD manifest plus quiet are embedded in the decision subject.
Earlier D manifests remain predecessor-chain history. E cannot embed its own
manifest without a digest cycle, so its exact Git commit parent/tree/metadata
and decision digest are the durable authority for the accepted attempt. The
ignored PrepareE manifests are recovery journals, while FArchive's ordered
`eAttempts` array makes every stale attempt and the final accepted attempt
clone-durable without inventing quiet for an attempt that lost before admission.
FArchive stores that array plus the full content-free predecessor manifest chain
and receipt identities; FDone creates the Done blob/marker; FPublish records the
known FD commit/blob plus FA/path/digest in master routing before any
  ignored-root cleanup. Those E/FA/FD/FP/archive identities are authenticated
  inputs to the final chain, but FP remains non-publishable. Only valid FS plus
  its tracked canonical seal—which embeds the E/FA/FD/FP quiet objects and binds
  every preceding commit—is the clone-durable closeout authority.

P or A “exists” only after its exact commit/tree and both required approvals are
valid. Supersession during a candidate boundary's commit/approval sequence uses
the corresponding before-boundary N/A in completion, while the Supersession
result still embeds the attempted commit and every zero/one/invalid approval
identity as invalidated evidence; it never erases the partial milestone or calls
it an approved boundary.

## Ordered Gate-0 handoff chain

The master owns one explicit ordinal list of Gate-0 round basenames and one
nullable `successfulGate0Handoff`. A round never edits a predecessor Done plan.
Each versioned round records the immediately preceding round's exact Done path,
closeout commit, Done blob, decision path/digest, and outcome before it begins.
The ordinal must increase by one; forks, duplicate ordinals, gaps, a WIP
predecessor, or a predecessor whose current bytes differ from its recorded blob
are fatal.

After a round moves to Done, one reviewed handoff-publication change appends that
exact basename to the master's explicit chain, records its closeout commit/blob
and immutable decision digest, and updates WIP README routing. This publication
changes no Done bytes. If the outcome is no winner, the master pointer remains
`none`, product plans 2–6 remain blocked, and a later authorized round is a new
WIP plan with this Done round as its predecessor. Repeat without limit by
appending one ordinal plan; never reopen, rename, or rewrite a Done round.

If the outcome is winner, the Done plan contains exactly one literal success-
handoff line. The exact marker spelling appears only in the retained Winner
branch below; this explanatory section deliberately does not repeat it.

It also contains exact labeled values
for its finalization manifest/path/digest/run ID, candidate, engine-lock digest,
five-lane verification-set identity, dependency status/cache results, decision
path/digest, and predecessor. No no-winner plan contains that marker. The master
handoff publication sets its single pointer to that exact Done basename plus
closeout commit/blob and evidence identities. Source-contract/Phase-9 checks
count exact marker byte occurrences across every Done file in the explicit
ordered Gate-0 list: each no-winner has zero, the sole pointer-named winner has
exactly one, and the lineage total is exactly one. Plan 2 consumes only that
winner-bearing Done blob and
evidence; earlier no-winner rounds are authenticated lineage, not its product
predecessor. A second success marker or any later Gate-0 append after success is
forbidden unless a separately approved engine-reselection plan first revokes the
entire product handoff.

## Verification

Before the experiment, run the Round-4 closeout and dependency-lifecycle tests,
then add focused Round-5 cases proving predecessor blob mismatch, versioned-file
overwrite, implicit-highest-version discovery, approval chronology, every
15-day boundary/terminal state, each synchronous-boundary hard stop, and zero
effects for a rejected/zero cohort. Cover both carry-audit states, missing/
duplicate/stale approvals, every expected/observed identity mismatch, an
unapproved viability assertion, approved post-freeze supersession, each
`proof-rejected` governance branch, all seven ABI actions, and rejection of
`Arc|Rc|Weak`, deferred destroy, callback, borrow, or action overlap. Exercise
   all eight supersession timing states at the exact admission fence and one instant
   on either side; prove bounded cancellation of only the active isolated action,
   no next effect, no frozen-record rewrite, exact invalidated-artifact retention,
   the timing-specific completion-template combination, every
   `NotStarted|CancelledBeforeCommit|CommittedThenReverted|AdoptedThenReverted`
   cleanup shape, adoption-path projection equality, conditional full-tree
   equality, exact adoption/revert commits, pre-cleanup quiet, cleanup token/
   claim/result, distinct post-cleanup quiet, and zero tracked engine state, the final Done
   closeout fence, blocked winner-pointer publication, and reselection routing for
   evidence discovered after closeout. Crash before/after cleanup claim, commit,
   and post-cleanup quiet; prove ControlQuiet result/D publication rejects the
   stale pre-cleanup receipt and that `CancelledBeforeCommit` restores/verifies
   the exact A projection without inventing a commit.
   Exercise the stage/timing legality matrix: only `CandidateLane` may execute
   on the pre-effect side or own admission receipts; every later stage requires
   the already-open global effect fence, except the non-retryable branch-neutral
   `FirstSelectionAttemptLedger` closeout over either two rejected receipts or
   one rejection stopped by approved supersession. Ignored
   receipts capture one admitted attempt, reject/admit with identical input
   digests, or reject/reject terminal history.
   The tracked closeout ledger must reproduce those receipts without affecting
   the earlier clean-repository identity; also cover reject then supersession
   before retry as `StoppedAfterFirstRejection`. Reject a changed retry input,
   third attempt, premature tracked ledger, missing required ledger, a wrong
   digest-free projection/entry hash, or pre-effect disposition for every later
   stage other than the explicit ledger exception. Mutate each projection field,
   the excluded outer fields, and the raw receipt digest independently. For each later qualification-
   failure stage, exercise every
   reachable post-effect partial-artifact combination, including proposal absent/
   invalid and proposal present with zero/one/two review records. Prove
   `PreAdoption` clean tree, `AdoptedThenReverted` exact commit/tree restoration,
   zero tracked engine state, no runner-up, and next-ordinal routing. Exercise a
	   retry-then-Winner and retry-then-Superseded result and require both to bind the
	   same attempt-ledger path/raw digest. Prove that an atomic-writer absent-output
	   fault pauses WIP and retries only the writer, while an artifact validation
	   failure takes the exact qualification stage. For every invalid upstream
	   receipt/ledger/proposal/review, prove the bytes remain only at the ignored
	   quarantine path, the tracked canonical target is absent, and the valid
	   Qualification/Supersession result embeds canonical
	   `invalidArtifactEvidence`. A dirty/pre-existing/racing fixed canonical path
	   is STOP/WIP, never committed, overwritten, deleted, or treated as a branch
	   result. Prove that ordinary no-winner
   finalization accepts exactly `no-scored-candidate`, `below-threshold`, or
   `tied-first-no-winner`, that qualification failure is the only
   non-superseded Evaluated branch that may pair `outcome=no-winner` with
   `selection.status=winner`, and that Superseded may retain that winner
   finalization only as explicitly invalidated evidence. Every retained
   completion label—including `round5Outcome`—occurs exactly once. Assert the
   exact completion-record marker mapping against branch and `round5Outcome`,
   and reject unresolved, swapped, unknown, or duplicate markers. Proposal-
   review tests reject a missing/wrong proposal digest, proposal-authored or
   duplicate reviewer, missing/fabricated ordinal, equal/reversed UTC, ordinal 2
   without an approving ordinal 1, a review or decision identity inside the
   proposal, overwrite, or implicit path discovery. They prove
   `Approve/Approve` alone reaches Winner, while `Reject` or `Approve/Reject`
   reaches the qualification-failed branch with the exact existing review
   record identities and no final decision created before that terminal branch
   evidence. Proposal chronology tests also reject equality/reversal against the
   ledger freeze or any bound artifact/approval UTC.

Exercise coordinator admission independently: all five preflights must complete
before any candidate effect; mutating any input leaf while retaining the claimed
input-set digest rejects; direct/provider/lane bypass produces zero effect; and
every lane record binds the admitted receipt, input set, P/A identity, token,
claim, runner, and unique output root. Reject same-runner duplicate claim,
simultaneous copied-token use on a second runner, reused action/output root,
stale epoch, a crash-after-claim replay, and a fresh ID/token/root that repeats
the same logical action key. Exercise every allowed-next transition and reject
a duplicate proof, collection tuple, preflight ordinal, first-selection lane,
approval/review ordinal, or fixed-output writer before effects. Cover a claimed preflight crash
before record publication, its one coordinator-synthesized failed record, and
the resulting rejected receipt. Cover proof and collection crashes before their
normal records, the exact synthesized terminal records, no rerun, complete
three-manifest finalization, and impossibility of a winner. Cover dependency/
cache/notices crash/cancel with the exact tracked failed report and stage matrix.

Exercise all four validation SuiteIds as allowed-next Pipeline keys. Require
P-candidate -> PreProofFocused receipt/quiet -> P approvals and A-candidate ->
PreResultFocused receipt/quiet -> A approvals. Require final D -> FinalFocused
receipt/quiet -> FinalCanonicalFull receipt/final-suite quiet -> decision
subject. For P/A Failed, require
`ValidationClosing -> ProtocolClosure -> protocol D -> both final suites` and
reject P/A approval, replay, repair, or a second candidate. For an ordinary
final-suite Failed, require the sole administrative SuccessorD, new ordered
final-suite keys, branch-preserve/invalidated-Winner disposition, the exact
post-D C revert before SuccessorD when B exists, and reject a second
administrative D. Cover crash, timeout, failure, wrong commit/tree, changed command/tool
digest, administrative ProtocolRepair old/new/mixed/extra/unapproved tool
  identity sets, repair presence on a capacity/P/A/pre-D/already-administrative
  closure, either stage-specific repair N/A value used in the other class,
  missing quiet, a second final-suite quiet/key, receipt-path collision,
direct invocation, and illegal
classification. Race an observation against each running final suite; prove no
next suite/decision starts, every successor D gets new commit-keyed paths, only
the final-D receipts enter completion, and old-D receipts remain immutable
ignored history. Exercise every legal P/A timing N/A or invalidated reference and
reject all other completion grammar.

Exercise the mandatory-payload calculator independently and through every
coordinator publication. Prove recursive object counting and exact RS-JCS
wrapper bytes; 192/786432 history, 64/262144 terminal reserve, and
256/1048576/per-object-65536 minus-one/equal/plus-one cases; the generated
largest route from every quiet state; observation `0008`, ordinary decision
`0009`, reserved ProtocolClosure decision `0010`, and recovery `0003`
ceilings/roles; and rejection before token/path allocation. Race
budget closure and validation closure against proposal publication, review
fence, a running suite, each zero/one/two P/A approval publication, and each
other in both orders. Require clone verification of candidate/receipt/quiet/
approval canonical objects and their invalidated/retained dispositions, the fixed reason
precedence, `PreserveTerminalBranch` for an immutable terminal result, queued
successor evidence after a winning closure CAS, exact ProtocolClosure
stage/track/embed identity, and zero candidate verdict synthesized by capacity
or validation failure. ProtocolRepair tests also reject repair-author review,
duplicate reviewer, non-Approve/equal/reversed chronology, wrong subject/tree/
allowlist/closure/budget digest, path collision, overwrite, changed replay
inputs, and a wrapper missing/reordering either canonical approval.

For every P/A/B/C/D/SuccessorD Git-boundary transaction, crash before/after
manifest, private-index tree, commit-object creation, ref CAS, real-index/
worktree synchronization, verification, coordinator acceptance, and quiet.
Prove recovery distinguishes object-not-ref and ref-not-synchronized without a
second commit/action, checks the manifest old/new ref pair, and rejects wrong
branch/ref, external mutation, alternate metadata/tree, unrelated dirty state,
or an unchained predecessor manifest. C additionally exercises the no-commit
CancelledBeforeCommit restoration and proves its cleanup front end and Git
boundary are one token/key/action.

Exercise the pre-A/P-to-A lifecycle and control plane: TransitionRejected with
only the bootstrap epoch; supersession before P; P approval clone verification;
proof plus every post-proof governance writer serialized in P; quiet rotation to
A; and rejection of disconnected/coexisting epochs. While a pipeline action is
hung, race proposal publication against next pipeline admission, crash on both
sides of the authoritative proposal rename, derive `ObservationPending`, run
review through only `ControlLive`, publish/recover the
authoritative fence at every crash boundary, cancel/quiet the action, and prove
no old-epoch admission. Reject a control token at a pipeline entry point and the
reverse. Enumerate every ProtocolClosure/ProtocolRepair and decision-generation
evidence/disposition ControlQuiet operation; reject class swaps, direct calls,
missing quiet, wrong reason/generation/disposition, and use while a Pipeline
action remains active.

Exercise observation-ledger generation 1, a rejected pre-E observation producing
an immutable cumulative successor D without changing branch, an approved one
changing to Superseded, and late rejected/approved observations in
`SupersededCarryForward`. Reject gaps, forks, changed prior entries, overwrite,
implicit-highest lookup, stale decision subjects, or a decision binding a
non-final generation. Exercise execution-freeze attestation exact/N/A timing and
embedded approval validation on every branch.
Cover zero observations (`g0001/Empty`), one observation with and without a
prior Empty generation, and eight observations after prior Empty ending at
`g0009`; closeout must reuse the exact latest complete generation and must never
request `g0010`.

Exercise every dependency completion grammar and stage-ordered combination,
including invalid bare pass, free-form N/A, absent digest, ignored path,
later-stage output after an earlier failure, and a failure-stage pass. Exercise
decision subject generations, Reject/stale-to-next-generation,
  generation-evidence subject-token/prefix/predecessor/audit/disposition chains,
  missing/mutated/wrong-generation DecisionSubject token rejection,
fresh-clone verification after ignored cleanup, author/reviewer
separation, equal/reversed times, crash before/after evidence, rejected/stale/
  accepted disposition, approved-decision, proposal win, E ref CAS, and final
  publication. Race proposal publication against DecisionSubject token
  publication: proposal first requires no ordinal/path/token allocation; token
  first blocks proposal until byte-identical subject publication. Crash after
  token and before staging creation/write/directory rename and require recovery
  of that one subject without a gap. Race proposal CAS against audit 1, audit 2,
  generation-evidence, approved-decision, and PrepareE ControlQuiet token
  admission in both orders: token first must publish/close its reserved output
  before proposal retry; proposal first leaves that token/output absent and uses
  the preceding completed cut point. Then race proposal after subject with zero
  audits, after audit 1 Approve, after audit 1/2 Reject but before evidence,
  after Rejected evidence, after Rejected disposition, immediately after audit
  2 Approve, after AuditedApproved evidence, before approved-decision rename,
  after approved-decision but before PrepareE, and after PrepareE but before
  ClosingPublish. Require the exact `ObservationPreempted|RejectedAudit|
  AuditedApproved` evidence, approved-decision/PrepareE present-or-N/A fields,
  Stale/Rejected disposition, stale-before-admission presence only for a
  PrepareE attempt, and observation resolution/ledger/successor-D order at each
  cut. Also test pre-rendered Accepted disposition size/digest/path/predicted-E
drift or oversize, prefix mutation, unreviewed suffix fields, both/neither CAS winner,
and clone verification
of embedded canonical audits. Race observation admission against the
`ClosingPublish` CAS in both orders, including observation publication after E
manifest render but before the CAS. Prove the observation-winning path leaves
the tracked decision absent and forces a new D/suite/decision/E generation,
while the publish-winning path blocks observation until `ClosedE`. Crash before/
after E commit-object creation, coordinator CAS, `update-ref`, real-index/
worktree synchronization, and state update; recover only the same E manifest,
derive `ClosedE` from the exact committed tree/metadata/decision, unchanged WIP/
master routing, absent Done/marker, and reject any third path state, alternate
tree, second commit, or split-brain state.

Crash `PrepareFArchive` and ArchivePublish before/after the exact journal,
object, E -> FA ref CAS, synchronization, and quiet; recover only FA. Crash
`PrepareFDone` and DonePublish before/after the exact journal, object, FA -> FD
ref CAS, WIP-to-Done synchronization, and quiet; prove FD changes only the
completed-plan move/marker set and master remains unchanged. Crash
`PrepareFPublish` and HandoffPublish before/after fence, object, FD -> FP ref
CAS, synchronization, and quiet; prove FP changes only master/README/
optional-next-WIP routing, records known FD commit/Done blob and FA/path/digest,
and never changes E decision, FD Done, or archive bytes. Flip every allowed FD
projection byte and every forbidden plan/master/README byte and require the
published-closeout verifier to distinguish them. Crash `PrepareFSeal` and
SealPublish before/after the exact journal, seal render, object, FP -> FS ref
CAS, synchronization, and verification; require the one-path `100644` seal add,
embedded canonical E/FA/FD/FP quiets, and no second action between FP quiet and
FS. At every
E/FA/FD/FP crash edge, reject observation admission and external push/export;
evidence is queued only for the post-publication process after ClosedF. Exercise
the no-winner template/render inputs and mandatory RoundStart at every crash
edge, including the exact six-token carry, P and P-less initial-D successors,
and retrospective verification through the successor RecoveryArchive. Exercise ClosedF
derivation and ignored-root cleanup with reparse/path-escape/overlapping-root/
extra-child/changed-leaf rejection, crash at every leaf/directory/result/quiet
edge, linked idempotent partial-deletion recovery, and UnsafeStopped; prove the
FArchive/FDone/FPublish/FSeal and cleanup-attempt roots are never deleted. Test
mandatory archive-payload and cleanup-inventory caps at minus-one/exact/plus-one,
prove candidate/build roots are never traversed, and retain over-cap roots
without partial deletion. Exercise zero/one/three recovery attempts for every
applicable P/A/B/C/D/SuccessorD boundary and for each PrepareE attempt. Cover
zero stale E attempts plus final Accepted, and the maximum eight ordered
StaleBeforeAdmission entries plus final Accepted; delete every eligible ignored
root and require a fresh clone to reproduce each non-E
manifest -> attempts -> quiet state and each E
manifest -> attempts -> stale-with-no-quiet or accepted-with-quiet state solely
from the RecoveryArchive.
For each pre-E counter, test its lower contract ceiling (`0008` observation,
`0009` cumulative ledger, ordinary `0009` plus ProtocolClosure-only `0010` decision, `0003`
recovery), the next-value disposition, and
rejection of wrap, reuse, gap, or publication beyond that ceiling. Test
PostFCleanup separately at `0001`, `9999`, and exhaustion before `10000`.

If admitted, reuse the legacy collection/finalization and first-selection
command bodies only through
`Tools/TerminalEngine/Invoke-TerminalEngineRound5Action.ps1`; direct Round-5 use
must reject before effects. Every invocation receives the four explicit v5
subject paths/raw digests plus GovernanceFreeze path/raw digest, matching coordinator effect or control-plane token
path/raw digest, expected epoch/action/input identity, P/A commit/tree/approval
identities applicable to that action, runner-local claim path, and create-new
output root. Each candidate lane additionally receives the exact attempt-receipt
path/raw digest and `inputSetSha256`; remote lanes receive the byte-identical
input-set/receipt bytes. Tests reject an omitted/mixed/default subject, bypassed
wrapper, stale/copied/replayed token, wrong epoch/runner/output root, changed
input behind the same claimed digest, or raw-effective-object rehash.

Run the repository's focused Gate-0 suites and final required suite through the
canonical runners. Every focused proof/governance/lifecycle test and every
candidate, G1–G12, collection, finalization, first-selection, and required native
lane must pass with no skip; a fake lane, non-native ARM64 execution, dirty
repository, network-capable evidence subprocess, or unrecorded owner-day
invalidates the round. An unrelated pre-existing environmental failure in the
canonical repository full suite may be classified only under the master's exact
policy: record the command, failing test and logs, prior matching baseline/archive,
reproduction and owner classification, prove zero Gate-0 impact, and let the
wrapper finish and report every remaining lane. It never excuses a focused or
mandatory Gate-0 failure and never creates a winner artifact.

## Completion and move

The predecessor identities above are already authenticated and are not Round-5
closeout placeholders. At closeout, replace `TERMINAL-HANDOFF-UNRESOLVED` and
every `<ROUND5_RESULT_...>` token in the closed template below with exact
reviewed values. Those are the only closeout placeholders in this plan;
illustrative metavariables in commands/contracts are not closeout fields. Keep
the all-closeouts block. Then retain exactly one of `TransitionRejected`,
`Superseded`, `ProtocolClosed`, or `Evaluated`; for `Evaluated`, also retain exactly one of
`No-winner`, `Qualification-failed`, or `Winner`. Delete the unused branches.
Every retained label occurs exactly once: the all-closeouts block is the sole
owner of `round5Outcome`, and branch rules constrain its value without repeating
the label. Do not use “latest,” an unbound summary, or a bare `pass` without its
canonical result identity.

Replace the top `Completion record` value by this exact closed mapping:

- retained `TransitionRejected`, `Superseded`, `ProtocolClosed`, `No-winner`, or
  `Qualification-failed` branch plus `round5Outcome=no-winner`:
  `TERMINAL-GATE0-ROUND5-NO-WINNER`;
- retained `Winner` branch plus `round5Outcome=winner`:
  `TERMINAL-GATE0-ROUND5-WINNER`.

The closeout verifier rejects an unresolved, unknown, duplicated, or
branch/outcome-mismatched completion record. This marker reports the Round-5
result; it does not replace the separately gated successful-handoff marker.

All-closeouts block:

- `round5Outcome=<ROUND5_RESULT_OUTCOME>` — exactly `no-winner` or `winner`.
- `round5CarryAudit=Specs/Terminal/TerminalEngineRound5CarryAudit.v1.json`
- `round5CarryAuditSha256=<ROUND5_RESULT_CARRY_AUDIT_SHA256>`
- `round5CarryAuditState=<ROUND5_RESULT_CARRY_AUDIT_STATE>` — exactly
  `Unchanged` or `TransitionRejected`.
- `round5CarryAuditApprovals=<ROUND5_RESULT_CARRY_AUDIT_APPROVALS>`
- `round5MandatoryPayloadBudget=Specs/Terminal/TerminalEngineRound5MandatoryPayloadBudget.v1.json`
- `round5MandatoryPayloadBudgetSha256=<ROUND5_RESULT_MANDATORY_PAYLOAD_BUDGET_SHA256>`
- `round5ProtocolClosure=<ROUND5_RESULT_PROTOCOL_CLOSURE_OR_NA>` — exactly
  `N/A-no-protocol-closure` or
  `Specs/Terminal/TerminalEngineRound5ProtocolClosure.v1.json#<64-lowercase-hex>`.
- `round5ProtocolClosureDisposition=<ROUND5_RESULT_PROTOCOL_CLOSURE_DISPOSITION>`
  — exactly `N/A-no-protocol-closure`, `ProtocolClosed`, or
  `PreserveTerminalBranch`, consistent with the retained branch and embedded
  closure object.
- `round5ProtocolRepair=<ROUND5_RESULT_PROTOCOL_REPAIR_OR_NA>` — exact
  `Specs/Terminal/TerminalEngineRound5ProtocolRepair.v1.json#<64hex>;
  approvals=<two-ordered-path/raw-digest/reviewer-identities>` when any
  allowlisted byte changes; `N/A-no-protocol-repair-required` only for an
  eligible failed-ordinary-final-suite closure with no byte change; or
  `N/A-no-protocol-repair-permitted` when
  `round5ProtocolClosure=N/A-no-protocol-closure` and for capacity, P/A
  validation, pre-ordinary-D, and already-administrative-D closure. A real
  repair is illegal when closure is N/A; the verifier rejects either N/A value
  in the other's eligibility class.
- `round5Decision=Specs/Terminal/TerminalEngineDecision.Round5.md`
- `round5DecisionSha256=<ROUND5_RESULT_DECISION_SHA256>`
- `round5DecisionUtc=<ROUND5_RESULT_DECISION_UTC>`
- `round5DecisionAudits=<ROUND5_RESULT_DECISION_AUDITS>`
- `round5DecisionCutoffCommit=<ROUND5_RESULT_E_COMMIT>` — exact 40-hex E.
- `round5RecoveryArchiveCommit=<ROUND5_RESULT_FA_COMMIT>` — exact 40-hex FA
  whose direct parent is E.
- `round5RecoveryArchivePath=<ROUND5_RESULT_RECOVERY_ARCHIVE_PATH>` — exact
  `Specs/TestRuns/TerminalEngineRecovery/round5-<YYYYMMDDTHHMMSSZ>-<16-lowerhex>/TerminalEngineRound5RecoveryArchive.v1.json`.
- `round5RecoveryArchiveSha256=<ROUND5_RESULT_RECOVERY_ARCHIVE_SHA256>`
- `round5SuccessorPlanPath=<ROUND5_RESULT_SUCCESSOR_PLAN_PATH>`
- `round5SuccessorTemplateSha256=<ROUND5_RESULT_SUCCESSOR_TEMPLATE_SHA256>`
- `round5SuccessorRendererSha256=<ROUND5_RESULT_SUCCESSOR_RENDERER_SHA256>`
- `round5SuccessorRenderInputSha256=<ROUND5_RESULT_SUCCESSOR_RENDER_INPUT_SHA256>`
  — every no-winner branch uses the exact pre-E audited successor identities;
  Winner uses `N/A-winner-routes-product-plan-2` for all four.
- `round5BoundaryEvidence=Specs/Terminal/TerminalEngineRound5BoundaryEvidence.v1.json`
- `round5BoundaryEvidenceSha256=<ROUND5_RESULT_BOUNDARY_EVIDENCE_SHA256>`
- `round5FocusedTests=<ROUND5_RESULT_FOCUSED_TESTS>`
- `round5CanonicalFullSuite=<ROUND5_RESULT_FULL_SUITE>`
- `round5FullSuiteClassification=<ROUND5_RESULT_FULL_SUITE_CLASSIFICATION>`

The retained branch and carry state are one invariant, not independent choices.
The `Transition rejected` branch is valid if and only if
`round5CarryAuditState=TransitionRejected`. `Superseded`, `ProtocolClosed`,
`No winner`, `Qualification failed`, and `Winner` are valid if and only if
`round5CarryAuditState=Unchanged`. The verifier and corpus swap each state across
every wrong branch and reject before E preparation.

Those three values use only this closed clone-durable grammar. A suite reference
is
`decision-embedded:<SuiteId>[path=.build/terminal-engine-spike/
round5-validation/<SuiteId>/<testedCommit>/receipt.v1.json#<receipt-64hex>;
quiet=.build/terminal-engine-spike/round5-validation/<SuiteId>/
<testedCommit>/post-suite-quiet.v1.json#<quiet-64hex>;
status=<terminal-status>;testedCommit=<40hex>]` with the prose line breaks
removed. `<SuiteId>` is one of the four exact case-sensitive IDs and both path
occurrences of `<testedCommit>` must equal the final field. The decision embeds
the exact canonical receipt and quiet objects/raw digests; the paths are
evidence identities, not clone dependencies.

`round5FocusedTests` is exactly the ordered, no-space sequence
`PreProofFocused=<reference-or-N/A>;PreResultFocused=<reference-or-N/A>;
FinalFocused=<reference>`. `round5CanonicalFullSuite` is exactly
`FinalCanonicalFull=<reference>`. The two final references must name the final
retained D/SuccessorD commit, `FinalFocused` status must be `Pass`, and
`FinalCanonicalFull` status must be `Pass|ClassifiedPreExisting`. An old-D
receipt, changed command/tool digest, absent quiet, repeated SuiteId, or
reordered field is invalid. The `FinalCanonicalFull` reference's fixed sibling
`post-suite-quiet.v1.json` is the single final-suite quiet; a second final-suite
receipt/path/key is invalid.

For `TransitionRejected`, the first two values are exactly
`N/A-transition-rejected`. For Superseded, a P/A suite never admitted before its
fence is respectively `N/A-not-admitted-before-P` or
`N/A-not-admitted-before-A`; a completed suite for an unapproved boundary is the
same reference prefixed by `invalidated-` and may have `Pass|Failed`; once its
P/A boundary is approved, use the ordinary `Pass` reference. Every non-
superseded evaluated branch requires ordinary `Pass` P and A references. No
other N/A/status is legal.

For a retained `ProtocolClosed` branch, capacity closure before a boundary uses
exactly `N/A-protocol-closed-before-P` or
`N/A-protocol-closed-before-A`. `preproof-focused-failed` requires a
PreProofFocused `Failed` reference plus the latter A N/A;
`preresult-focused-failed` requires ordinary PreProofFocused `Pass` plus
PreResultFocused `Failed`; capacity closure after an approved boundary uses its
ordinary `Pass` reference. Capacity closure after a suite Pass but before both
boundary approvals uses that same reference prefixed by `invalidated-` and the
ProtocolClosure embeds the candidate manifest plus the exact zero/one/two
approval disposition; two approving records use the ordinary approved form.
The same invalidated-Pass form applies when CapacityClosing wins while the
already-reserved suite later returns Pass. For a branch-preserving ProtocolClosure, the
underlying retained branch's P/A grammar remains authoritative and the closure
record embeds any additional failed suite as invalidated evidence. In every
case the final two references name the final protocol or administrative
SuccessorD and satisfy the normal Pass/classification rule.

`round5FullSuiteClassification` is exactly `Pass` when the full receipt is
`Pass`, or
`ClassifiedPreExisting[baseline=<40hex>;reproduction=<64hex>;
owner=<closed-owner-id>;noGate0Impact=<64hex>]` when its status is
`ClassifiedPreExisting`. The embedded receipt contains and authenticates the
same baseline, reproduction, owner, and no-Gate-0-impact evidence. A failed
ordinary final suite takes the one ProtocolClosure/administrative-SuccessorD
route above. Failure on that administrative D, a bare/free-form classification,
or classification on another SuiteId leaves the round WIP.

This plan owns the exception semantics. The canonical full wrapper must finish
and report every lane. Its ordered non-pass set must match an explicit archived
pre-P baseline one-for-one by test ID, normalized failure/skip signature, native
architecture/configuration, OS/toolchain identity, and log-digest class, and
must reproduce completely on a clean checkout of that baseline commit with the
same command/environment. The canonical receipt embeds the baseline and
reproduction archive paths/raw digests, owner/UTC, complete affected-test set,
and reviewed path/component-backed `noGate0Impact` proof. A missing/new/changed
item, incomplete wrapper run, environment drift, ignored/mutable evidence,
unowned decision, focused/native/mandatory Gate-0 failure, or unsupported impact
claim is `Failed`, never classified. The schema/corpus and closeout verifier
exercise every rejection.

Dependency/cache/notices completion values use this closed grammar. A successful
artifact is `pass:<repository-relative-path>#<64-lowercase-hex>`. A completed
failure is
`failed:<closed-lowercase-kebab-reason>:<repository-relative-path>#<64-lowercase-hex>`.
Superseded uses the same two forms prefixed by `invalidated-`, or exactly
`N/A-not-started-at-supersession-fence`. Qualification-failed uses the ordinary
pass/failed forms or, respectively, exactly
`N/A-failed-before-dependency-status`,
`N/A-failed-before-dependency-cache`, and
`N/A-failed-before-notices-license`. Paths are normalized repository-relative
ordinary content-free report files committed at D under the selected
`Specs/TestRuns/<run-id>/` archive; ignored `.build` paths are invalid in
completion. The pipeline action first closes its live report below `.build`;
after quiet, the `ControlQuiet` D archiver reopens/recomputes and copies only the
schema-valid content-free canonical report to its fixed tracked archive path.
The
Winner proposal, Qualification-failure record, or Supersession result also
embeds each applicable canonical report object and raw digest. A bare `pass`,
missing digest, free-form N/A, or another prefix is invalid.

The dependency stage order is locked status, cache, then notices/license. On
Superseded, a later completed field is legal only when every earlier field is
`invalidated-pass`; after `invalidated-failed`, all later fields are the exact
not-started value. On Qualification-failed, a stage before
`DependencyLockedStatus` uses all three exact before-stage N/As;
`DependencyLockedStatus` requires locked status failed and the later two N/As;
`DependencyCache` requires locked pass, cache failed, and the
notices N/A; `NoticesLicense` requires locked/cache pass and notices failed;
`WinnerHandoffProposal|WinnerHandoffProposalApproval` requires all three pass.
The failure stage must name the first failed field, so a pass at its own failure
stage is rejected. Winner requires all three pass. The schema and closeout
verifier derive these combinations from stage, timing, action-token, and
artifact chronology.

`TransitionRejected` branch:

- `round5TransitionReason=superseded-by-expanded-universe`
- `round5TransitionEvidence=<ROUND5_RESULT_TRANSITION_EVIDENCE>` — exact sorted
  expected/observed identities, closed reason codes, and content-free evidence
  digests already bound by the approved carry audit.
- `round5ProofState=N/A-transition-rejected`
- `round5ProofReason=N/A-transition-rejected`
- `round5ProofAction=N/A-pre-universe`
- `round5ProofRecord=N/A-transition-rejected`
- `round5ProofLedgerSha256=N/A-transition-rejected`
- `round5Universe=N/A-transition-rejected`
- `round5CandidateReview=N/A-transition-rejected`
- `round5CandidateAssessment=N/A-transition-rejected`
- `round5CandidateSet=N/A-transition-rejected`
- `round5CandidateGovernanceFreeze=N/A-transition-rejected`
- `round5ProofFreezeBoundary=N/A-pre-universe`
- `round5ExecutionFreezeAttestation=N/A-pre-universe`
- `round5ExecutionFreezeAttestationSha256=N/A-pre-universe`
- `round5SupersessionObservationLedger=N/A-pre-universe`
- `round5SupersessionObservationLedgerSha256=N/A-pre-universe`
- `round5SupersessionObservationLedgerGeneration=N/A-pre-universe`
- `round5SupersessionObservationLedgerState=N/A-pre-universe`
- `round5AdmittedCandidateCount=0`
- `round5Supersession=N/A-pre-universe-transition-rejected`
- `round5Collections=N/A-no-executable-cohort`
- `round5Finalization=N/A-no-executable-cohort`
- `round5Materialization=N/A-no-winner`
- `round5EngineLock=N/A-no-winner`
- `round5VerificationSet=N/A-no-winner`
- `round5DependencyLockedStatus=N/A-no-winner`
- `round5DependencyCacheStatus=N/A-no-winner`
- `round5NoticesLicenseStatus=N/A-no-winner`
- `round5FirstSelectionAttemptLedger=N/A-transition-rejected`
- `round5QualificationFailure=N/A-transition-rejected`
- `round5WinnerHandoffProposal=N/A-transition-rejected`
- `round5WinnerHandoffReviews=N/A-transition-rejected`
- `round5SuccessfulGate0Handoff=N/A-no-winner`

The `TransitionRejected` decision authenticates the predecessor, carry-audit
bytes/approvals, tests, and zero-effect closeout. It contains no v5 universe,
proof, candidate governance, collection, finalization, or winner claim.

`ProtocolClosed` branch:

- `round5TransitionReason=protocol-closed`
- `round5ProtocolClosureReason=<ROUND5_RESULT_PROTOCOL_CLOSURE_REASON>` —
  exactly `evidence-budget-exceeded`, `preproof-focused-failed`,
  `preresult-focused-failed`, or `final-validation-failed`.
- `round5ProtocolClosureStage=<ROUND5_RESULT_PROTOCOL_CLOSURE_STAGE>` — exact
  coordinator state/logical-key projection validated against the reason and
  failed receipt or prospective-budget calculation.
- `round5ProtocolClosureFence=<ROUND5_RESULT_PROTOCOL_CLOSURE_FENCE>` — exact
  `CapacityClosing|ValidationClosing` token/raw digest, stop/pre-cleanup quiet,
  and terminal `activeActionCount=0`.
- `round5ProtocolClosureBudgetSnapshot=<ROUND5_RESULT_PROTOCOL_CLOSURE_BUDGET>`
  — exact table/raw digest, used/candidate/reserved object and canonical-byte
  values, and remaining worst terminal route.
- `round5ProtocolClosureBoundaryEvidence=<ROUND5_RESULT_PROTOCOL_BOUNDARY_EVIDENCE_OR_NA>`
  — exact tracked BoundaryEvidence path/raw digest when it predates closure, or
  `N/A-protocol-closure-before-boundary-evidence`; the closure and archive embed
  every earlier boundary/validation canonical object either way.
- `round5ProtocolClosurePartialArtifacts=<ROUND5_RESULT_PROTOCOL_PARTIAL_ARTIFACTS>`
  — the schema-closed ordered snapshot of P/A candidate transactions, suites,
  zero/one/two approval canonical objects and dispositions, then
  proof/governance/collection/
  finalization/materialization/lock/adoption/verification/dependency/proposal/
  review objects as `not-started|completed-invalidated|failed-invalidated`,
  each with its timing-specific exact N/A or path/raw digest. It uses the same
  field projections and stage ordering as Superseded but never claims a
  viability supersession or candidate verdict.
- `round5ProtocolClosureCleanup=<ROUND5_RESULT_PROTOCOL_CLEANUP>` — exact
  `NotStarted|CancelledBeforeCommit|CommittedThenReverted|
  AdoptedThenReverted` action/quiet/projection grammar, requiring zero tracked
  engine state before protocol D.
- `round5ProtocolClosureDecisionInput=<ROUND5_RESULT_PROTOCOL_DECISION_INPUT>` —
  exact final protocol D/SuccessorD manifest/quiet plus the final ordered
  validation receipts/quiets.
- `round5SuccessfulGate0Handoff=N/A-protocol-closed`

The retained ProtocolClosed decision embeds the canonical closure, every
applicable P/A failed or passed receipt/quiet, the closed partial-artifact and
cleanup projections, and the final protocol D suites. Collection/finalization
may appear only as invalidated evidence; no engine lock, success marker, or
product route is legal. If a ProtocolClosure instead has
`PreserveTerminalBranch`, do not retain this branch: retain the already-fixed
TransitionRejected/Superseded/Evaluated branch and bind the auxiliary closure
through the all-closeouts fields plus its successor-D decision input.

`Superseded` branch:

- `round5TransitionReason=superseded-by-expanded-universe`
- `round5Universe=Specs/Terminal/TerminalEngineCandidateUniverse.v5.json`
- `round5UniverseSha256=<ROUND5_RESULT_UNIVERSE_SHA256>`
- `round5UniverseFrozenUtc=<ROUND5_RESULT_UNIVERSE_FROZEN_UTC>`
- `round5UniverseApprovals=<ROUND5_RESULT_UNIVERSE_APPROVALS>`
- `round5ProofFreezeBoundary=<ROUND5_RESULT_SUPERSEDED_PROOF_FREEZE_OR_NA>` —
  exact P commit/tree, both embedded canonical approval objects/raw digests, and
  P-epoch identity, or exactly `N/A-superseded-before-proof-freeze`.
- `round5ProofAction=<ROUND5_RESULT_SUPERSEDED_PROOF_ACTION>` — exactly
  `N/A-no-proof-freeze` when P is absent,
  `N/A-proof-action-not-admitted` for `BeforeProof` after approved P, or the
  exact token/claim with
  `atFence=active|completed|cancelled|aborted|crashed` and terminal
  `atQuiet=completed|cancelled|aborted|crashed` otherwise. `active` is legal only in the
  at-fence snapshot and can never be the final Supersession disposition.
- `round5ExecutionFreezeAttestation=<ROUND5_RESULT_SUPERSEDED_EXECUTION_FREEZE_ATTESTATION_OR_NA>`
  — exact
  `Specs/Terminal/TerminalEngineRound5ExecutionFreezeAttestation.v1.json`, or
  `N/A-superseded-before-execution-freeze`.
- `round5ExecutionFreezeAttestationSha256=<ROUND5_RESULT_SUPERSEDED_EXECUTION_FREEZE_ATTESTATION_SHA256_OR_NA>`
  — exact raw digest when A existed, otherwise
  `N/A-superseded-before-execution-freeze`.
- `round5SupersessionObservationLedger=<ROUND5_RESULT_SUPERSEDED_OBSERVATION_LEDGER_PATH>`
  — exact final cumulative generation path.
- `round5SupersessionObservationLedgerSha256=<ROUND5_RESULT_SUPERSEDED_OBSERVATION_LEDGER_SHA256>`
- `round5SupersessionObservationLedgerGeneration=<ROUND5_RESULT_SUPERSEDED_OBSERVATION_LEDGER_GENERATION>`
- `round5SupersessionObservationLedgerState=ApprovedSupersession`
- `round5Supersession=Specs/Terminal/TerminalEngineRound5Supersession.v1.json`
- `round5SupersessionSha256=<ROUND5_RESULT_SUPERSESSION_SHA256>`
- `round5SupersessionTiming=<ROUND5_RESULT_SUPERSESSION_TIMING>`
- `round5SupersessionTriggerReviews=<ROUND5_RESULT_SUPERSESSION_TRIGGER_REVIEWS>`
  — exact triggering proposal and ordered review-1/review-2 path/raw-digest/
  reviewer/`Approve` identities. The final result itself is not approved; later
  carry-forward reviews appear only in the final cumulative observation ledger.
- `round5SupersessionFence=<ROUND5_RESULT_SUPERSESSION_FENCE_PATH_SHA256>`
- `round5SupersessionPreCleanupQuietReceipt=<ROUND5_RESULT_SUPERSESSION_PRE_CLEANUP_QUIET_PATH_SHA256>`
- `round5SupersessionCleanupAction=<ROUND5_RESULT_SUPERSESSION_CLEANUP_ACTION_OR_NA>`
  — exact cleanup token/claim/terminal identity unless cleanup state is
  `NotStarted`, then exactly `N/A-no-cleanup-action`.
- `round5SupersessionPostCleanupQuietReceipt=<ROUND5_RESULT_SUPERSESSION_POST_CLEANUP_QUIET_OR_NA>`
  — exact distinct receipt after a cleanup action, otherwise exactly
  `N/A-no-cleanup-action`.
- `round5ProofState=<ROUND5_RESULT_SUPERSEDED_PROOF_STATE>` — exact
  `PendingProof|ProofSucceeded|BudgetExceeded|Rejected`, or
  `N/A-before-proof`; `PendingProof` is legal only for `DuringProof`.
- `round5ProofRecord=<ROUND5_RESULT_SUPERSEDED_PROOF_RECORD>` — exact path and
  SHA-256; `N/A-before-proof`; or the exact
  `N/A-proof-action-invalidated[token=<path>#<sha>;claim=<path>#<sha>;disposition=<completed|cancelled|aborted|crashed>]`
  when P existed but no proof record froze.
- `round5ProofLedger=<ROUND5_RESULT_SUPERSEDED_PROOF_LEDGER>` — exact owner-
  ledger path/raw digest/latest published ordinal, `N/A-before-proof`, or
  `N/A-before-first-proof-ledger-entry`. `DuringProof` must use the exact
  partial ledger whenever any owner-day entry was published.
- `round5ProofLedgerSha256=<ROUND5_RESULT_SUPERSEDED_PROOF_LEDGER_SHA256>` —
  the same exact raw digest, or the same timing-specific N/A.
- `round5CandidateReview=<ROUND5_RESULT_SUPERSEDED_REVIEW>` — exact immutable
  path/SHA-256/freeze UTC, or `N/A-before-governance-freeze`.
- `round5CandidateAssessment=<ROUND5_RESULT_SUPERSEDED_ASSESSMENT>` — exact
  immutable path/SHA-256/freeze UTC, or `N/A-before-governance-freeze`.
- `round5CandidateSet=<ROUND5_RESULT_SUPERSEDED_SET>` — exact immutable
  path/SHA-256/freeze UTC, or `N/A-before-governance-freeze`.
- `round5CandidateGovernanceFreeze=<ROUND5_RESULT_SUPERSEDED_GOVERNANCE_FREEZE>`
  — exact journal path/raw digest/freeze ID/UTC with all three reserved subjects,
  or `N/A-before-governance-freeze`; no partial combination is legal.
- `round5AdmittedCandidateCount=<ROUND5_RESULT_SUPERSEDED_ADMITTED_COUNT>` —
  exact frozen count, or `N/A-before-governance-freeze`.
- `round5Collections=<ROUND5_RESULT_SUPERSEDED_COLLECTIONS>` — exact sorted
  complete/aborted content-free artifact identities labeled
  `invalidated-by-supersession`, or `N/A-before-collection`.
- `round5Finalization=<ROUND5_RESULT_SUPERSEDED_FINALIZATION>` — exact
  invalidated path/digest/run identity with any schema-valid selection status,
  including `winner`, or `N/A-not-produced`.
- `round5Materialization=<ROUND5_RESULT_SUPERSEDED_MATERIALIZATION>` — exact
  invalidated identity, or `N/A-not-produced`.
- `round5EngineLock=<ROUND5_RESULT_SUPERSEDED_ENGINE_LOCK>` — exact
  invalidated identity with `never-adopted` or `adopted-then-reverted`
  disposition, or `N/A-not-produced`.
- `round5AdoptionCleanupState=<ROUND5_RESULT_SUPERSEDED_ADOPTION_CLEANUP_STATE>`
  — exactly `NotStarted|CancelledBeforeCommit|CommittedThenReverted|
  AdoptedThenReverted`.
- `round5Adoption=<ROUND5_RESULT_SUPERSEDED_ADOPTION>` — exact adoption commit/
  tree and tracked-path digest set for
  `CommittedThenReverted|AdoptedThenReverted`; otherwise exactly
  `N/A-no-adoption-commit`.
- `round5AdoptionRevert=<ROUND5_RESULT_SUPERSEDED_ADOPTION_REVERT>` — exact
  non-history-rewriting revert commit/tree, reverted-path digest set, and
  zero-remaining-engine-adoption proof for
  `CommittedThenReverted|AdoptedThenReverted`; otherwise exactly
  `N/A-no-revert-commit`.
- `round5AdoptionProjectionProof=<ROUND5_RESULT_SUPERSEDED_ADOPTION_PROJECTION>`
  — `N/A-adoption-never-started` only for `NotStarted`; otherwise exact
  pre-adoption/adoption-path projection path/raw digest proving equality to A.
  `CancelledBeforeCommit` requires that proof with neither commit;
  `CommittedThenReverted|AdoptedThenReverted` require it with both commits.
- `round5VerificationSet=<ROUND5_RESULT_SUPERSEDED_VERIFICATION_SET>` — exact
  invalidated identity, or `N/A-not-produced`.
- `round5DependencyLockedStatus=<ROUND5_RESULT_SUPERSEDED_LOCKED_STATUS_OR_NA>`
- `round5DependencyCacheStatus=<ROUND5_RESULT_SUPERSEDED_CACHE_STATUS_OR_NA>`
- `round5NoticesLicenseStatus=<ROUND5_RESULT_SUPERSEDED_NOTICES_STATUS_OR_NA>`
  — each uses only the superseded grammar and ordered fence/quiet combination
  above.
- `round5FirstSelectionAttemptLedger=<ROUND5_RESULT_SUPERSEDED_ATTEMPT_LEDGER>` —
  exact invalidated path/raw digest with state
  `AdmittedFirstAttempt|AdmittedAfterRetry|SecondRejection|
  StoppedAfterFirstRejection` when valid; exact
  `invalid-receipt[path=<path>#<sha>;attemptedInputsSha256=<sha>]` or
  `invalid-ledger[path=<path>#<sha>;attemptedInputsSha256=<sha>]`; exact
  `N/A-validation-rejected-before-publication[artifact=<receipt|ledger>;
  attemptedInputsSha256=<sha>]`; or the
  exact timing-derived `N/A-finalization-did-not-name-winner` /
  `N/A-superseded-before-attempt-ledger`. Validation substitutes are legal only
  when the same attempted object was pending at the fence. When one valid
  rejection preceded supersession, require the exact
  `StoppedAfterFirstRejection` ledger.
- `round5QualificationFailure=<ROUND5_RESULT_SUPERSEDED_QUALIFICATION_FAILURE>`
  — exact invalidated path/digest when already produced, or
  `N/A-not-produced`.
- `round5WinnerHandoffProposal=<ROUND5_RESULT_SUPERSEDED_WINNER_PROPOSAL>` —
  exact invalidated `PV`, attempted invalid/absent `PX`, or not-reached `P0`
  canonical projection from the Supersession result.
- `round5WinnerHandoffReviews=<ROUND5_RESULT_SUPERSEDED_WINNER_REVIEWS>` —
  exact invalidated `R0|RX|R1|R2` canonical projection. `PX|RX` preserves the
  invalid raw identity or exact validation-rejected-before-publication
  substitute and attempted inputs; it cannot collapse to ordinary not-started.
- `round5SuccessfulGate0Handoff=N/A-superseded`

The supersession schema derives which exact/N/A combination is legal from its
timing state and milestone digests. It requires all three governance subjects
together because their common freeze is atomic. It admits only actions whose
tokens were issued before the fence, then records their completed, cancelled,
crashed, or aborted post-fence quiet result as invalidated; a post-fence action
admission is always illegal. It proves no success marker or product-plan handoff
was published. `DuringAdoption` and `AfterAdoptionBeforeCloseout` require the
state-specific cleanup/projection evidence above and a clean tracked-engine
projection. It never rewrites or fabricates a subject to make the combination
fit.

`Evaluated` common block:

- `round5ProofState=<ROUND5_RESULT_PROOF_STATE>` — exactly
  `ProofSucceeded`, `BudgetExceeded`, or `Rejected`.
- `round5ProofReason=<ROUND5_RESULT_PROOF_REASON>`
- `round5ProofAction=<ROUND5_RESULT_PROOF_ACTION>` — exact P-epoch token/claim/
  terminal disposition embedded by the proof record.
- `round5ProofRecord=Specs/Terminal/TerminalEngineRound5Proof.v1.json`
- `round5ProofRecordSha256=<ROUND5_RESULT_PROOF_RECORD_SHA256>`
- `round5ProofLedgerSha256=<ROUND5_RESULT_PROOF_LEDGER_SHA256>`
- `round5Universe=Specs/Terminal/TerminalEngineCandidateUniverse.v5.json`
- `round5UniverseSha256=<ROUND5_RESULT_UNIVERSE_SHA256>`
- `round5UniverseFrozenUtc=<ROUND5_RESULT_UNIVERSE_FROZEN_UTC>`
- `round5UniverseApprovals=<ROUND5_RESULT_UNIVERSE_APPROVALS>`
- `round5ProofFreezeBoundary=<ROUND5_RESULT_PROOF_FREEZE>` — exact P commit/tree,
  both embedded canonical approval objects/raw digests, P-epoch identity, and
  quiet receipt.
- `round5ExecutionFreezeAttestation=Specs/Terminal/TerminalEngineRound5ExecutionFreezeAttestation.v1.json`
- `round5ExecutionFreezeAttestationSha256=<ROUND5_RESULT_EXECUTION_FREEZE_ATTESTATION_SHA256>`
- `round5SupersessionObservationLedger=<ROUND5_RESULT_OBSERVATION_LEDGER_PATH>`
  — exact final cumulative generation path.
- `round5SupersessionObservationLedgerSha256=<ROUND5_RESULT_OBSERVATION_LEDGER_SHA256>`
- `round5SupersessionObservationLedgerGeneration=<ROUND5_RESULT_OBSERVATION_LEDGER_GENERATION>`
- `round5SupersessionObservationLedgerState=<ROUND5_RESULT_OBSERVATION_LEDGER_STATE>`
  — exactly `Empty` or `RejectedOnly`.
- `round5Supersession=N/A-not-superseded`
- `round5CandidateReview=Specs/Terminal/TerminalEngineCandidateReview.v5.json`
- `round5CandidateReviewSha256=<ROUND5_RESULT_REVIEW_SHA256>`
- `round5CandidateReviewFrozenUtc=<ROUND5_RESULT_COMMON_FREEZE_UTC>`
- `round5SourceApprovals=<ROUND5_RESULT_SOURCE_APPROVALS>`
- `round5CandidateAssessment=Specs/Terminal/TerminalEngineCandidateAssessment.v5.json`
- `round5CandidateAssessmentSha256=<ROUND5_RESULT_ASSESSMENT_SHA256>`
- `round5CandidateAssessmentFrozenUtc=<ROUND5_RESULT_COMMON_FREEZE_UTC>`
- `round5AdmittedCandidateCount=<ROUND5_RESULT_ADMITTED_COUNT>`
- `round5AdmissionApprovals=<ROUND5_RESULT_ADMISSION_APPROVALS_OR_NA>` — exact
  nonempty ordered approval path/raw-digest/reviewer identities when
  `round5AdmittedCandidateCount` is greater than zero; exactly
  `N/A-zero-admitted-candidates` when it is zero. No other N/A is legal.
- `round5CandidateSet=Specs/Terminal/TerminalEngineCandidateSet.v5.json`
- `round5CandidateSetSha256=<ROUND5_RESULT_SET_SHA256>`
- `round5CandidateSetFrozenUtc=<ROUND5_RESULT_COMMON_FREEZE_UTC>`
- `round5CandidateGovernanceFreeze=Specs/Terminal/TerminalEngineCandidateGovernanceFreeze.v5.json`
- `round5CandidateGovernanceFreezeSha256=<ROUND5_RESULT_GOVERNANCE_FREEZE_SHA256>`
- `round5AdaptationApprovals=<ROUND5_RESULT_ADAPTATION_APPROVALS>`

No-winner branch:

- `round5Collections=<ROUND5_RESULT_COLLECTIONS_OR_NA>` — either the exact
  ordered x64 Release/x64 ASan Debug/ARM64 Release path-and-digest set from a
  complete comparison, or `N/A-no-executable-cohort`.
- `round5Finalization=<ROUND5_RESULT_FINALIZATION_OR_NA>` — exact path, digest,
  run ID, and the evidence-derived non-winner `selection.status` exactly
  `no-scored-candidate`, `below-threshold`, or `tied-first-no-winner` when
  collection completed, or `N/A-no-executable-cohort`.
- `round5Materialization=N/A-no-winner`
- `round5EngineLock=N/A-no-winner`
- `round5VerificationSet=N/A-no-winner`
- `round5DependencyLockedStatus=N/A-no-winner`
- `round5DependencyCacheStatus=N/A-no-winner`
- `round5NoticesLicenseStatus=N/A-no-winner`
- `round5FirstSelectionAttemptLedger=N/A-finalization-did-not-name-winner`
- `round5QualificationFailure=N/A-finalization-did-not-name-winner`
- `round5WinnerHandoffProposal=N/A-finalization-did-not-name-winner`
- `round5WinnerHandoffReviews=N/A-finalization-did-not-name-winner`
- `round5SuccessfulGate0Handoff=N/A-no-winner`

Qualification-failed branch:

The Qualification-failure schema owns this closed pre-dependency artifact
matrix. Its canonical artifact projections use these exact state codes:

- `M0|L0|V0|P0` — stage not reached, with respectively
  `N/A-failed-before-materialization`, `N/A-failed-before-lock`,
  `N/A-not-produced`, or `N/A-proposal-not-produced`;
- `MX|LX|VX|PX` — the named stage started and ended as exactly one of
  `AbsentAfterStart|InvalidBeforePublication|FailedPublished`. It binds the
  pipeline token, claim, terminal receipt, and attempted logical target. For
  `InvalidBeforePublication`, it embeds the canonical content-free
  `invalidArtifactEvidence` object and the fixed tracked target is absent; it
  never commits/quarantines invalid bytes at that target or uses a not-reached
  N/A;
- `MV0|MV1|MV2` and `LV0|LV1|LV2` — valid materialization or lock plus exactly
  zero, one, or two ordered approving review identities. A failure at an
  approval stage additionally binds the rejecting/invalid/absent next-review
  evidence and cannot use the next higher state;
- `A0` — adoption not started; `AX` — adoption started but was cancelled/
  failed before commit and binds token/claim/terminal receipt plus the exact
  A-equal path projection; `AR` — adoption committed and the exact C revert
  restored the A path projection with zero engine state;
- `VV` — valid verification set plus all five exact ordered lane identities;
  `PV` — valid winner-handoff proposal; and
- `R0|RX|R1|R2` — respectively no review started, invalid/absent review before
  publication, one terminal review, or two ordered terminal reviews. A terminal
  review set contains a Reject or validation failure; two `Approve` reviews are
  Winner and cannot appear in Qualification-failed.

Every `X`, approval-failure, `AX`, and `R*` failure state embeds the canonical
attempted inputs and content-free reason/reproducer; every valid state embeds
the artifact canonical object/raw digest. The completion value is the exact
canonical `state=<code>;path=<path-or-exact-N/A>;sha256=<digest-or-exact-N/A>;
approvals=<ordered-set-or-N/A>;disposition=<invalidated|never-adopted|
adopted-then-reverted>` projection from that record; line breaks shown here are
only prose formatting. The stage-to-state relation is exhaustive:

| `round5QualificationFailureStage` | Materialization | Lock | Adoption/revert | Verification set | Proposal/reviews |
|---|---|---|---|---|---|
| `CandidateLane` | `M0` | `L0` | `A0` | `V0` | `P0/R0` |
| `FirstSelectionAttemptLedger` with `SecondRejection` | `M0` | `L0` | `A0` | `V0` | `P0/R0` |
| `FirstSelectionAttemptLedger` after an otherwise complete admitted sequence | `MV2` | `LV2` | `AR` | `VV` | `P0/R0`; all dependency fields are pass |
| `Materialization` | `MX` | `L0` | `A0` | `V0` | `P0/R0` |
| `MaterializationApproval` | `MV0|MV1` plus terminal next-review evidence | `L0` | `A0` | `V0` | `P0/R0` |
| `CandidateLock` | `MV2` | `LX` | `A0` | `V0` | `P0/R0` |
| `LockApproval` | `MV2` | `LV0|LV1` plus terminal next-review evidence | `A0` | `V0` | `P0/R0` |
| `TrackedAdoption` | `MV2` | `LV2` | `AX|AR` | `V0` | `P0/R0` |
| `LockBoundBuildVerify` | `MV2` | `LV2` | `AR` | `V0` | `P0/R0` |
| `VerificationSet` | `MV2` | `LV2` | `AR` | `VX` | `P0/R0` |
| `DependencyLockedStatus|DependencyCache|NoticesLicense` | `MV2` | `LV2` | `AR` | `VV` | `P0/R0` |
| `WinnerHandoffProposal` | `MV2` | `LV2` | `AR` | `VV` | `PX/R0` |
| `WinnerHandoffProposalApproval` | `MV2` | `LV2` | `AR` | `VV` | `PV/RX|R1|R2` |

The dependency columns further obey the earlier dependency-stage matrix. The
schema/writer derives this row from the first terminal stage and rejects every
cross-row state, approval count, N/A, disposition, or cleanup combination.
When ledger closeout validation is secondary to an already-terminal substantive
stage, retain that substantive row and bind the invalid/absent ledger substitute
separately; the two FirstSelectionAttemptLedger rows are legal only when ledger
closeout itself is the first failure.

- `round5Collections=<ROUND5_RESULT_COLLECTIONS>` — exact ordered three-manifest
  path/digest/run identities.
- `round5Finalization=<ROUND5_RESULT_FINALIZATION>` — exact path, digest, run ID,
  `selection.status=winner`, and selected candidate.
- `round5QualificationFailure=Specs/Terminal/TerminalEngineRound5QualificationFailure.v1.json`
- `round5QualificationFailureSha256=<ROUND5_RESULT_QUALIFICATION_FAILURE_SHA256>`
- `round5QualificationFailureStage=<ROUND5_RESULT_QUALIFICATION_FAILURE_STAGE>`
- `round5QualificationFailureTiming=<ROUND5_RESULT_QUALIFICATION_FAILURE_TIMING>`
  — exactly `PreAdoption` for `A0|AX` or `AdoptedThenReverted` for `AR`.
- `round5AdoptionCleanupState=<ROUND5_RESULT_QUALIFICATION_ADOPTION_CLEANUP_STATE>`
  — exactly `NotStarted` for `A0`, `CancelledBeforeCommit` for `AX`, or
  `AdoptedThenReverted` for `AR`.
- `round5QualificationPreCleanupQuietReceipt=<ROUND5_RESULT_QUALIFICATION_PRE_CLEANUP_QUIET_PATH_SHA256>`
- `round5QualificationCleanupAction=<ROUND5_RESULT_QUALIFICATION_CLEANUP_ACTION_OR_NA>`
  — exact cleanup token/claim/terminal identity for `AX|AR`; exactly
  `N/A-no-cleanup-action` for `A0`.
- `round5QualificationPostCleanupQuietReceipt=<ROUND5_RESULT_QUALIFICATION_POST_CLEANUP_QUIET_OR_NA>`
  — exact distinct receipt for `AX|AR`; exactly `N/A-no-cleanup-action` for
  `A0`.
- `round5QualificationFailureReason=<ROUND5_RESULT_QUALIFICATION_FAILURE_REASON>`
- `round5FirstSelectionAttemptLedger=<ROUND5_RESULT_QUALIFICATION_ATTEMPT_LEDGER>`
  — exact path/raw digest with state
  `AdmittedFirstAttempt|AdmittedAfterRetry|SecondRejection`, or the canonical
  embedded `invalidArtifactEvidence` identity plus
  `N/A-validation-rejected-before-publication` attempted-entry identity
  permitted only for a `CandidateLane` receipt-validation or
  `FirstSelectionAttemptLedger` closeout-validation failure.
- `round5QualificationArtifacts=<ROUND5_RESULT_QUALIFICATION_ARTIFACTS>` —
  exact ordered path/digest/disposition identities for every rejected
  pre-effect attempt and every started/completed/failed first-selection action.
- `round5WinnerHandoffProposal=<ROUND5_RESULT_QUALIFICATION_WINNER_PROPOSAL_OR_NA>`
  — exact `P0|PX|PV` canonical projection required by the stage row.
- `round5WinnerHandoffReviews=<ROUND5_RESULT_QUALIFICATION_WINNER_REVIEWS_OR_NA>`
  — exact `R0|RX|R1|R2` canonical projection required by the stage row.
- `round5SelectedCandidateId=<ROUND5_RESULT_WINNER_CANDIDATE_ID>`
- `round5Materialization=<ROUND5_RESULT_QUALIFICATION_MATERIALIZATION_OR_NA>` —
  exact `M0|MX|MV0|MV1|MV2` canonical projection required by the stage row.
- `round5EngineLock=<ROUND5_RESULT_QUALIFICATION_ENGINE_LOCK_OR_NA>` — exact
  `L0|LX|LV0|LV1|LV2` canonical projection required by the stage row; a valid
  lock remains below `.build` until B and has exactly
  `never-adopted|adopted-then-reverted` disposition.
- `round5Adoption=<ROUND5_RESULT_QUALIFICATION_ADOPTION_OR_NA>` — exact
  `A0|AX|AR` canonical projection required by the stage row.
- `round5AdoptionRevert=<ROUND5_RESULT_QUALIFICATION_ADOPTION_REVERT_OR_NA>` —
  exact non-history-rewriting revert commit/tree, reverted-path digest set, and
  zero-remaining-engine-adoption proof for `AR`; exactly
  `N/A-no-adoption-commit` for `A0|AX`.
- `round5VerificationSet=<ROUND5_RESULT_QUALIFICATION_VERIFICATION_SET_OR_NA>` —
  exact `V0|VX|VV` canonical projection required by the stage row.
- `round5DependencyLockedStatus=<ROUND5_RESULT_QUALIFICATION_LOCKED_STATUS_OR_NA>`
- `round5DependencyCacheStatus=<ROUND5_RESULT_QUALIFICATION_CACHE_STATUS_OR_NA>`
- `round5NoticesLicenseStatus=<ROUND5_RESULT_QUALIFICATION_NOTICES_STATUS_OR_NA>`
  — each uses only the qualification grammar and failure-stage combination
  above.
- `round5SuccessfulGate0Handoff=N/A-qualification-failed`

Winner branch:

- `round5Collections=<ROUND5_RESULT_COLLECTIONS>` — exact ordered three-manifest
  path/digest/run identities.
- `round5Finalization=<ROUND5_RESULT_FINALIZATION>` — exact path, digest, run ID,
  and `selection.status=winner`.
- `round5Materialization=<ROUND5_RESULT_MATERIALIZATION>` — exact manifest/
  digest and both approvals.
- `round5EngineLock=<ROUND5_RESULT_ENGINE_LOCK>` — exact path/digest and both
  approvals.
- `round5VerificationSet=<ROUND5_RESULT_VERIFICATION_SET>` — exact set path/
  digest plus all five ordered lock-bound lane path/digest identities.
- `round5DependencyLockedStatus=<ROUND5_RESULT_LOCKED_STATUS>` — exact `pass:`
  grammar above.
- `round5DependencyCacheStatus=<ROUND5_RESULT_CACHE_STATUS>` — exact `pass:`
  grammar above.
- `round5NoticesLicenseStatus=<ROUND5_RESULT_NOTICES_LICENSE_STATUS>` — exact
  `pass:` grammar above.
- `round5FirstSelectionAttemptLedger=Specs/Terminal/TerminalEngineRound5FirstSelectionAttemptLedger.v1.json`
- `round5FirstSelectionAttemptLedgerSha256=<ROUND5_RESULT_ATTEMPT_LEDGER_SHA256>`
- `round5FirstSelectionAttemptLedgerState=<ROUND5_RESULT_ATTEMPT_LEDGER_STATE>` —
  exactly `AdmittedFirstAttempt` or `AdmittedAfterRetry`.
- `round5QualificationFailure=N/A-first-selection-complete`
- `round5WinnerHandoffProposal=Specs/Terminal/TerminalEngineRound5WinnerHandoffProposal.v1.json`
- `round5WinnerHandoffProposalSha256=<ROUND5_RESULT_WINNER_PROPOSAL_SHA256>`
- `round5WinnerHandoffReviews=<ROUND5_RESULT_WINNER_PROPOSAL_REVIEWS>` — exact
  ordinal-1 and ordinal-2 path/digest/reviewer/`Approve` identities.
- `successfulGate0FinalizationManifest=<ROUND5_RESULT_FINALIZATION_PATH>`
- `successfulGate0FinalizationSha256=<ROUND5_RESULT_FINALIZATION_SHA256>`
- `successfulGate0WinnerCandidateId=<ROUND5_RESULT_WINNER_CANDIDATE_ID>`
- `successfulGate0EngineLockSha256=<ROUND5_RESULT_ENGINE_LOCK_SHA256>`
- `successfulGate0VerificationSetManifest=<ROUND5_RESULT_VERIFICATION_SET_PATH>`
- `successfulGate0VerificationSetSha256=<ROUND5_RESULT_VERIFICATION_SET_SHA256>`
- **successfulGate0Handoff:** `TERMINAL-GATE0-SUCCESS`

The selected evaluated branch must agree byte-for-byte with the immutable v5
subjects and current decision pointer. `winner` requires `ProofSucceeded`, a
nonzero admitted cohort, all three collection identities, winner finalization,
materialization, lock, five-lane verification set, and passing dependency/cache/
  notices closure plus the immutable winner-handoff proposal and both immutable
  approving review records.
Ordinary `no-winner` may retain exact collection/finalization
evidence only when a complete admitted cohort actually ran and its selection
status is one of the three closed non-winner shapes above; it never fabricates
later winner artifacts. `qualification-failed` is the sole non-superseded
  Evaluated no-handoff branch allowed to bind `selection.status=winner`; it
  requires the attempt ledger or its one legal validation-failure substitute,
  closed failure record, exact stage-valid partial-artifact disposition,
  `PreAdoption` clean-tree proof or the complete adoption/revert pair, zero
  tracked engine state, and no success marker. Superseded may retain a winner
  finalization and its attempt/proposal/review history only as explicitly
  invalidated evidence under its dedicated branch. Any approved supersession
  cannot be represented as an ordinary evaluated no-winner or qualification
  failure.

Before E, every branch with `outcome=no-winner` freezes and records in the
audited decision and later FD Done bytes:

- `round5SuccessorPlanPath=<exact absent Specs/Plans/WIP Round-6 path>`
- `round5SuccessorTemplateSha256=<64-lowercase-hex template digest>`
- `round5SuccessorRendererSha256=<64-lowercase-hex renderer digest>`
- `round5SuccessorRenderInputSha256=<64-lowercase-hex canonical input digest>`

Winner records the exact value `N/A-winner-routes-product-plan-2` for all four.
FP cannot select a basename/date/template after E or render from current mutable
bytes; it uses only these audited identities.

Record one of exactly six closeout branches:

- **Transition rejected:** exact `outcome=no-winner`, predecessor/carry-audit/
  decision/test identities, approved transition evidence, and every downstream
  v5/proof/candidate/collection/winner field set to the exact N/A values above.
- **Superseded:** exact `outcome=no-winner`, immutable v5 universe and approved
  supersession identities, mechanically derived timing, the exact already-
  completed proof/governance/evidence/attempt/proposal/review identities or
  timing-specific N/A values, and no success marker/product handoff.
- **Protocol closed:** exact `outcome=no-winner`, canonical ProtocolClosure
  reason/stage/fence/budget, every applicable P/A validation receipt,
  invalidated partial-artifact projection, zero-engine cleanup, final protocol
  D and passing/classified final suites, and no candidate verdict, success
  marker, or product handoff.
- **No winner:** exact `outcome=no-winner`, proof terminal state/reason,
  immutable v5 universe/review/assessment/set/decision paths and digests,
  approvals, focused/final test results, applicable exact collection/finalization
  identities or their closed N/A reasons, and explicit N/A reasons for lock,
  native winner lanes, and `successfulGate0Handoff`. Do not emit the success
  marker.
- **Qualification failed:** exact `outcome=no-winner`,
  `selection.status=winner`, selected candidate, immutable qualification-failure
  record, failing stage/timing/reason, every stage-valid partial artifact and
  the attempt-ledger identity or its legal validation-failure substitute, any
  proposal/review-record identities, pre-adoption clean-tree proof or exact
  adoption/revert pair, zero tracked engine state, no runner-up, and no success
  marker. Append the next ordinal Gate-0 round for any changed pin/design.
- **Winner:** exact `outcome=winner`, all of the above applicable identities plus
  the approved finalization/collections, winner, materialization, lock, five
  lock-bound verification identities, dependency/cache/notices results, the
  admitted attempt ledger, proposal plus both approving review records, and the
  single success marker/value contract above.

E deliberately does not close the plan. F is one logical handoff boundary with
four deterministic Git subtransactions so neither the Done plan, master, nor
seal must contain its own creating commit OID.

First, `PrepareFArchive` creates the E-parented FArchive manifest/tree/object;
its after-set contains only the exact tracked
`Specs/TestRuns/TerminalEngineRecovery/round5-<YYYYMMDDTHHMMSSZ>-<16-lowerhex>/TerminalEngineRound5RecoveryArchive.v1.json`.
Under the still-sealed coordinator lease, create-new
`<FArchive-transaction-directory>/archive-publish-journal.v1.json` binds that
manifest and the distinct ArchivePublish token. FArchive alone advances
`git update-ref <ref> <FA> <E>`, synchronizes, verifies, and quiets. Recovery
accepts only ref E/FA and the manifest-predicted object/tree/metadata; Round-5
observation admission remains closed.

Second, `PrepareFDone` creates the FA-parented FDone manifest/tree/object. Its
exact before/after path mapping fills every reviewed closeout value, sets
`State: Done`, removes
`Specs/Plans/WIP/Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md`,
and creates the byte-identical-basename
`Specs/Plans/Done/Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md`.
The Done bytes record FA/path/digest but never FD or FP; Winner alone contains
the single success marker. Under the same lease, create-new
`<FDone-transaction-directory>/done-publish-journal.v1.json` binds that manifest
and DonePublish token, then only `git update-ref <ref> <FD> <FA>` may advance.
Recovery accepts only ref FA/FD and the predicted state. There is no standalone
`git mv`, manual edit, or second plan-writing commit.

Third, from accepted FD, `PrepareFPublish` creates the routing-only FPublish
manifest/tree/object. Its after-set contains only the master chain/tokens, WIP
README routing, and—on a no-winner—the exact absent next ordinal WIP plan. It
records known FD as `closeoutCommit`, the Done blob resolved at `FD:<DonePath>`,
the decision/outcome, and known FA plus RecoveryArchive path/raw digest. It
never changes the E decision, FD Done, or FA archive. Under the sealed lease,
create-new `<FPublish-transaction-directory>/handoff-publish-fence.v1.json`
binds the exact manifest and HandoffPublish token, then only
`git update-ref <ref> <FP> <FD>` may advance. Recovery accepts only ref FD/FP
and the manifest-predicted state; an alternate ref/tree/object, unrelated delta,
duplicate next round/pointer, or changed decision/Done/archive is STOP.

On a no-winner, that exact successor plan and its new master row contain both
`predecessorHandoffPublishCommit=pending-parent-fp` and
`predecessorCloseoutSealCommit=pending-parent-fs`, plus
`predecessorCloseoutSealSha256=pending-parent-fs`; the master labels use the
same values under `gate0Round[N]`. They are deliberately unresolved in FP
because neither FP nor FS may self-record. The mandatory branch-neutral
RoundStart of that successor is FS's direct child and replaces only those six
sentinels with the FP/FS OIDs and raw seal digest before any carry review,
observation, P, or legal initial D. Verified RoundStart is the successor
admission boundary; the successor RecoveryArchive later embeds and
authenticates it only for retrospective clone verification. Round 5 is the one
legacy-predecessor exception and uses the exact legacy N/A values.

FPublish records the FD closeout commit/blob/decision/outcome and the exact
RecoveryArchive FA/path/raw digest in
`gate0Round[1].closeoutCommit|doneBlob|decisionPath|decisionSha256|
recoveryArchiveCommit|recoveryArchivePath|recoveryArchiveSha256|outcome`.
Plan 2 and clone verifiers require `FP^=FD`, `FD^=FA`, read that explicit
archive blob from FA, and never enumerate `Specs/TestRuns`. On no winner,
FPublish routes the new ordinal WIP plan and leaves all nine master successful-
handoff tokens `none`. On winner, it atomically replaces all nine values, sets
the matching bold `successfulGate0Handoff` pointer, and routes N3 to Plan 2 only
after Plan 2's fail-closed predecessor check consumes the exact lineage/archive
identities.

Fourth, after FP's fixed post-transaction quiet proves
`activeActionCount=0`, `PrepareFSeal` creates the one canonical CloseoutSeal and
the create-new
`<FSeal-transaction-directory>/seal-publish-journal.v1.json` binding the
SealPublish token. Only `git update-ref <ref> <FS> <FP>` may advance. FS has FP
as its sole parent and adds only the ordinary `100644`
`Specs/Terminal/TerminalEngineRound5CloseoutSeal.v1.json`; that RS-JCS UTF-8
no-BOM/no-trailing-newline value embeds the canonical E/FA/FD/FP quiets and binds
E/FA/FD/FP, outcome, archive, Done blob, and route. Recovery accepts only ref
FP/FS and the prepared state. E, FA, FD, and FP are never externally
publishable.

`ClosedF` is derived only from a valid FS/seal and the mutation-free published-
closeout verifier in `InitialCloseout` mode. That verifier proves `FS^=FP`,
`FP^=FD`, `FD^=FA`, `FA^=E`,
each single-parent exact semantic transform, the Done/archive/master/README
routes, and all four seal-embedded quiets with `activeActionCount=0`. The tracked
seal is sufficient in a fresh clone even after ignored roots are removed. Only
`ClosedF` admits the optional `PostFCleanup`; a missing/invalid seal or merely
prepared E/FA/FD/FP state does not.

## STOP conditions

- The Round-4 Done predecessor or any v5 frozen subject cannot be authenticated.
- A carried mismatch or viability claim lacks the closed carry-audit or
  supersession proposal, ordered reviews, coordinator fence/quiet, cumulative
  observation ledger, and final result required by its timing.
- A coordinator token/action class/epoch/claim is missing or replayed; P cannot
  rotate cleanly to A; a Git recovery reserve is exhausted with indeterminate
  ref state; PostFCleanup reaches `9999`; or ClosingPublish
  recovery sees bytes other than its exact before/after manifest.
- A history-growing publication bypasses the mandatory-payload prospective
  calculation, its budget snapshot/table digest changes, the worst terminal
  route does not fit, or a capacity loss allocates partial action bytes instead
  of taking the reserved closure CAS.
- A second P/A candidate or same-round P/A repair is attempted; ProtocolClosure
  has no exact reason/fence/cleanup/partial-artifact projection; or either final
  suite fails on the one administrative protocol D.
- E admission has a pending unapproved observation or unreverted tracked
  adoption; observation admission reopens before verified FS; or E/FA/FD/FP is
  externally published instead of deterministically recovered to FS.
- Any requirement, cap, gate, weight, threshold, tie-break, native lane, or
  product ownership rule would change after the experiment result is known.
- The proof spike reaches 15 owner-days without complete proof of the frozen
  synchronous hypothesis; stop as `BudgetExceeded` without extending the clock.
- The synchronous caller-owned boundary violates any hard stop above.
- A zero/rejected cohort would create collection/finalization/winner artifacts.
- Finalize names a winner but a later first-selection failure would be retried
  after candidate effects, fall through to a runner-up, omit the exact
  qualification-failure record/stage artifacts, or leave tracked engine state
  after closeout.
- A no-winner result would start product plan 2, or a winner result lacks the one
  exact successful handoff and full first-selection/dependency proof.
- Closing this round would edit a prior Done round, skip an ordinal, create two
  success markers, use implicit latest discovery, or leave master/README/Plan-2
  routing inconsistent.
