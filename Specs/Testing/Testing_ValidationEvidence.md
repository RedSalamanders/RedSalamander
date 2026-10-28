# Trusted Validation Evidence

Status: normative active contract
Last reviewed: 2026-10-04

## Authority and activation

This specification owns canonical validation records, affected-set decisions, promoted
entry evidence, resumable composite runs, and their trust boundaries. The implementation
history is recorded by
[`Operation Startrail`](../Plans/Done/Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md).
That Done file is non-normative implementation history. The completed
[`cross-domain remediation`](../Plans/Done/CodeReview_CrossDomain_Remediation_2026-08-10.md)
records the post-completion `L4-BUILD-01..08` and `L4-STAR-01..16` repairs;
this specification remains the authoritative behavior contract.
Build receipts and portable handoff are owned by `Specs/Build/Build_Toolchain.md`.
`Fresh`, explicit exact `Resume`, and opt-in `Affected` are active as described below;
CI, nightly, release, and final closeout remain explicit Fresh.

## Workspace and entry identity (Phase 3 active)

FileOperations child-profile isolation is runner-owned execution policy. Its unique
`LOCALAPPDATA` path is derived beneath the selected run's scratch root, independently
for every launch and classification retry. It is not a caller-controlled capability
or a reusable input directory. The existing runner dependency digest includes this
policy, so receipts produced before the change cannot authorize exact reuse afterward.
The process owns that directory for its complete lifetime; caller environment and
journal files are not modified.

`Get-RSWorkspaceSnapshot` is the canonical local source identity. It resolves only
local Git objects, uses the merge base of an explicit impact base without fetching,
and records separate committed-base, index, and worktree byte identities for changed,
staged, untracked, deleted, renamed, and type-changed paths. `.git/`, `.build/`,
`Specs/TestRuns/`, and `*.log` are evidence/generated exclusions. Other unknown paths
remain inputs. Rename sides are classified independently: an included-to-excluded rename
records the included deletion, an excluded-to-included rename records the included
addition, and only two included sides form one linked rename. Repository-relative paths
use Git's canonical spelling; Windows root
identity is case-insensitive. File mtimes are not semantic, while line-ending bytes are.
Reparse points, path escape, unresolved index stages, unsafe option-like revisions, and
an unstable double acquisition fail closed.

The Fresh runner publishes schema-valid initial, post-entry/post-repair, and final
snapshots. Any source-identity change after launch makes the run red. Before launching
entries it also publishes `entry-fingerprints.json`. Each fingerprint binds the complete
stable plan entry and mode-independent `coverage_plan_digest`, initial workspace/input-set identities, build receipt
and artifact epoch, declared artifact/runtime closure, reviewed runner/schema module
digests, classification mode, quarantine ledger, declared non-secret capabilities, and
cache-policy version. Fingerprint and snapshot computation time is reported separately
from test duration. This foundation is deliberately conservative: until the governed
impact graph activates, each declared input set binds the complete workspace snapshot.
Fingerprint publication does not permit reuse by itself.

Each entry also persists `validation-entry-components.v1`, an exact component-digest
projection. Resume compares it before the aggregate fingerprint in this precedence:
build artifact, logical input, shared contract, runner, environment, arguments,
coverage, outcome policy, capability, artifact epoch, then source input. The first
changed component selects the diagnostic reason. A version mismatch, missing component,
or unknown future component is `SCHEMA_MISMATCH`; equal components with a different
aggregate fingerprint are corrupt evidence.

Every persisted validation plan also carries a truthful invocation `plan_digest` over
the actual `Fresh`, `Resume`, or `Affected` mode. Fresh and Resume plans for the same
semantic coverage therefore have different invocation identities but the same coverage
identity. Reuse and entry fingerprints compare the coverage identity; records within one
run must still agree on that run's invocation identity and mode.

## Governed impact planning

`Tools/validation-impact.json` is the reviewed path-classification and impact graph.
Specific matching rules accumulate; rename/delete evaluation includes both path sides.
Rules distinguish product, suite-test-only, shared test, runner, build/workflow,
documentation, evidence/generated, and reviewed-no-validation inputs. A relevant path
without a specific reviewed rule selects the broad fallback and Full build/validation
surface. Every tracked solution/project/import/runtime input must be a member of the
derived graph or an exact reviewed standalone rule; adding an ungoverned project input
is a planning error rather than a silent fallback.

After shared-library adoption, current product UI test sources map to the live
`Tests/ProductUiTests/ProductUiTests.vcxproj` consumer and the existing standalone
validation surface. Retired `Common/DxUi` and `Tests/DxUiTests` projects must not remain
as live derived-closure owners. Exact DxUi pin changes and current shared UI adapter
changes must still select `ProductUiTests` and their product consumers; the conservative
Full fallback remains valid for shared adapters without a narrower reviewed mapping.

Executable specification inputs are not prose exclusions. Build and validation schemas,
settings schemas, theme runtime data, performance budgets, terminal corpora/evidence,
and normative consistency maps have explicit dependent build/test surfaces. An unknown
`Specs/**/*.json` or `Specs/**/*.json5` input widens conservatively to Full. Only reviewed
Markdown documentation may use the documentation-only/no-validation mapping.

`Run-AllTests.ps1 -ExplainPlan [-ImpactBase <local-commit-ish>]` publishes a deterministic
schema-valid shadow decision before any artifact lock, build, or validation-entry launch.
Ordinary Fresh execution publishes the same shadow decision beside its plan and
fingerprints, then still executes every entry. File Operations self-test-only inputs
select File Operations without Commands;
File Operations production/shared UI inputs select both. Documentation-only changes may
select no runtime entry, while unknown relevant inputs widen to Full.

The active mode matrix is: Fresh, CI, and release execute their complete explicit plans;
ExplainPlan is read-only and produces repository `NOT_EVALUATED`; Resume is Full-only,
requires an explicit direct origin, and permits only exact-artifact promoted reuse;
Affected is Full-only, executes the governed impacted subset without result reuse, and
produces repository `NOT_EVALUATED`. Startrail-owned exit categories remain 0
success, 1 selected-test failure, 2 repository not evaluated, 3 invalid CLI/plan, 4
blocked/missing attestation, and 5 corrupt evidence.

## Trust domains

Local receipts bind one physical checkout and output profile. Portable build
attestations bind a producer job to a declared consumer in the same workflow run and are
accepted only after source, toolchain, runtime-closure, and downloaded-artifact digests
verify against both the manifest and trusted reusable-workflow outputs. Repository or
pull-request record content alone cannot self-assert workflow provenance.

Executed test evidence is provisional until a schema-valid normalized terminal result is
`PASSED`, a post-entry source/artifact barrier succeeds, and an immutable promotion
certificate binds the terminal-result and evidence digests, entry fingerprint,
pre/post snapshot equality, artifact epoch, outcome policy, and coverage. A passing JSON
row, terminal aggregate, or interrupted run is not by itself reusable evidence.

## Canonical record contract

All Startrail schemas use JSON Schema draft-07 and reject unknown properties except at a
separately versioned extension point. Records are bounded to 16 MiB before parsing, a
maximum depth of 64, 100,000 total nodes, 1 MiB characters per string, and 1,024
characters per repository-relative path. Per-entry copied evidence is limited to 64 MiB;
one materialized run is limited to 1 GiB. Lower owning contracts may impose tighter
limits. Result artifacts are regular non-reparse files published by same-directory
create-only move after flush. Evidence stores relative path, exact size, and lowercase
SHA-256. Publication reopens both source and destination, and Resume/aggregation rehash
the destination; source drift, truncation, replacement, or quota overflow fails closed.

Canonical semantic bytes are UTF-8 without BOM, insignificant whitespace, or trailing
newline. Persisted records append one LF; digests hash the canonical semantic projection
without that LF. Object keys and strings are NFC. Keys sort ordinally after
normalization by UTF-16 code unit, and duplicate normalized keys fail. Contract-ordered arrays retain order;
set arrays are normalized first and require unique items. Legal scalars are null where
the schema permits it, lowercase Boolean, minimally escaped NFC string using lowercase
hex for control escapes, and signed 64-bit invariant decimal integer. Floating-point,
decimal, date, and implicit coercion are forbidden.
Readers enforce depth, node, and string bounds with a streaming token scan before object
materialization, then recanonicalize the parsed value and require byte-for-byte equality
with the persisted JSON payload. Pretty printing, reordered keys, alternate escapes,
decomposed Unicode, BOM, CRLF, and extra trailing LF are rejected.

Every digest names its projection and excludes itself, timestamps, durations, labels,
original path spelling, and absolute diagnostic paths unless a concrete contract says
otherwise. Repository-relative paths use forward slashes, reject absolute, traversal,
backslash, and empty-segment forms, and are checked for ordinal-ignore-case aliases on
Windows.

Startrail record creation, validation, and reuse require PowerShell 7.4 or newer.
Windows PowerShell 5.1 must reject a Startrail mode before evidence I/O. Legacy Fresh
behavior remains unchanged until the rollout gate deliberately changes the runner's
minimum host.

## Publication and retention

A writer validates and canonicalizes a complete record into a unique same-directory
temporary file, flushes it, then atomically replaces or moves the destination. A failure
preserves either the previous complete record or no record. Evidence roots must resolve
beneath the verified `X:/RedSalamander.Perf/evidence` root without reparse or
traversal escape.

`run-state.json` is the only mutable checkpoint in a run directory. Plans, normalized
terminal results, entry evidence, promotion certificates, copied result artifacts, and
decisions beneath `decisions/<decision-digest>.json` are immutable once published.
Content-addressed publication is create-only; an existing path is accepted only when its
canonical content is identical. Run state points at the active decision digest. Resume
creates a new run and references the original promoted execution
directly; provenance chains and cycles are forbidden. Pruning retains referenced origins
or first materializes an independently complete promoted copy.

### Local runtime-run retention

The transient `runs/<runId>` tree is archived on the same marked volume under
`evidence/completed-runs/<runId>` by the lifecycle in `Testing_SelfTests.md`.
Its separate create-only `evidence/run-archive-records/<runId>` intent/completion
records support crash recovery and preserve the original execution-path mapping.
It neither moves nor rewrites canonical Startrail run records or promotions.
The runner finalizes and independently copies terminal result artifacts before this
rename, then publishes its aggregate with final navigation paths. Runner, sandbox,
common owner-parser and summary changes remain in recursive runner dependency digests;
old fingerprints cannot qualify the new workflow. Retention never changes a run verdict
or confers exact Resume authority.

### External whole-run archives with reparse entries

A whole-run archive containing reparse entries MUST declare its closure representation.
`i27.v3-whole-run-no-follow-reparse.v1` uses `archive-manifest-v3.json`, an ordinary
non-reparse `whole-run` tree, a separate `no-follow-reparse-catalog.json`, and regular
raw blobs under `reparse-data`. Its complete claim is ordinary file bytes plus exact
no-follow raw reparse metadata. It is not a literal source-tree clone; it recreates no
links and reads neither resolved link targets nor their payload. Completion and scope
flags must be actual JSON booleans. Archive closure does not change the original test
verdict or authorize qualification by itself. This external transport format does not
change canonical Startrail records, promotion or Resume authority.

Admission MUST bind the exact marked archive authority, marked source authority,
original source run path and archive destination. Source and archive authorities may
differ; a D source archived under C does not confer authority over another D path.
Inventory treats supported junctions and symbolic links as leaves, never traversal
roots. All archive outputs and their ancestors remain ordinary non-reparse paths.
Unsafe paths, replacement drift, unsupported tags, malformed or truncated metadata,
duplicate mapping paths, missing or extra members, and changed ordinary bytes fail closed.

The closed controller records the exact archive manifest path, SHA-256 and successful
independent reader validation in its original invocation receipt. Later retention and
curation MUST match that original receipt, rerun archive validation and require unchanged
manifest/catalog hashes across validation. A rewritten, internally consistent manifest
is insufficient provenance.

Capture MUST compare source inventories before and after copying. Ordinary files retain
no-follow source identity, attributes, size and byte hashes before/after capture, and
matching destination size/hash. Each reparse leaf retains its source volume/file
identity, attributes, tag, exact raw buffer bytes, size and SHA-256 at the pre-copy,
post-copy and final source boundaries; those values must agree. The source reparse-path
set, manifest catalog, separate catalog and blob set MUST form a bijection. Every blob
must match its catalog size/hash and supported raw header; no missing, duplicate or
extra blob is admitted. An empty source requires an empty ordinary inventory, catalog
and blob set; an ordinary-only run has an empty catalog and blob set.

Readers MUST bound each external UTF-8 manifest/catalog to 256 MiB before parsing,
inventory/catalog sets to 200,000 entries, and each raw blob to 16 KiB. These are
independent ceilings: the effective capacity is whichever bound is reached first, and
overflow is an explicit failed/incomplete archive. These limits do not raise canonical
record, per-entry, materialized-evidence or repository-publication quotas. Capped
selected-record publication remains governed by `Testing_PerformanceValidation.md` and
`Specs/TestRuns/README.md`; its map binds actual external manifest, catalog and blob
sizes/hashes, with explicit scope and external-retention references.

An incomplete original archive and failed original run remain failed historical evidence.
Later capture establishes only the later capture boundary. It cannot retroactively
establish an omitted link's original raw buffer, resolved target or target payload,
replace the original manifest, or convert the old run into a pass. Canonical entry,
plan, terminal-result, receipt, artifact-epoch and promotion checks remain required
independently.

### Fresh checkpoint publication (Phase 5A active)

Full and CI Fresh runs create a unique `X:/RedSalamander.Perf/evidence/runs/<run-id>/`
directory, immutably publish their plan and content-addressed shadow decision, and
atomically replace only `run-state.json`. Run state is born `stable` and records the
active decision digest. An entry checkpoints `running` before launch, preserves reviewed
terminal result/log artifacts, publishes a content-addressed normalized terminal result
and provisional evidence after execution, then promotes only when the post-source barrier
matches, required artifacts exist, the outcome is a non-flaky pass, and the receipt/epoch
contract is bound. The terminal normalizer fails closed on a missing or malformed result,
wrong suite selection, no discovered cases/tests, invalid case status, count/aggregate
mismatch, incomplete coverage, all-skipped or unallowed-skipped outcomes, missing required
artifacts, nonzero exit, or blocking classification. A self-test skip is allowed only under
the entry's outcome contract (declared-capability skips; `unclassified` and `unexpected`
classes are forbidden). A Pester skip is allowed only when the runner reports named case
results and the skipped case is one of the environment-bound contracts listed by
`Get-RSPesterEnvironmentBoundSkipCases` (the Pester entries' `static_skip_cases`): archive
and history contracts that need gitignored evidence or unsquashed history. A skip count
without named cases, or any other skipped Pester case, is `UNALLOWED_SKIP`. Failed,
flaky, crashed, timed-out, partial, source-mutating, and interrupted entries do not
promote. A dead `running`/`provisional-passed` checkpoint repairs to `interrupted`.

Pruning retains the newest bounded set, caller-kept IDs, active runs, and referenced
origins. Targets must be direct validated children of the runs root, and every existing
component from the trusted evidence root through the target must be non-reparse. Planning
emits the exact canonical path, validated run ID, and Windows volume/file identity for
each target. Apply accepts only those plan entries, requires the canonical parent to be
exactly `runs`, and revalidates the complete chain plus the same directory identity
immediately before recursive removal. Nested paths, crafted ancestor-junction paths,
final reparse points, and plan/apply replacement drift fail closed. Cache publication
failure is a validation failure, never manufactured success. These records become
reusable only through explicit exact Resume validation.

Prune provenance is not trusted from state labels alone. Every referenced origin must
resolve through the same verified promotion/evidence loader used by Resume, aggregate
summaries, and duration history before it can influence retention.

### Artifact writers and epochs (Phase 5B active)

Every plan entry declares artifact reads, writes, produced outputs, and receipt
invalidation scopes. Before a non-isolated writer can delete or replace bytes, the runner
verifies the current read barrier, suspends the selected profile's receipt pointer, and
durably advances run state to a new `mutating` epoch while invalidating intersecting
promotions. Every artifact reader checks that run state is `stable` and that its epoch and
receipt match. A crash after receipt suspension therefore leaves no usable receipt; a
dead `mutating` checkpoint repairs to `interrupted` and is never reusable.

The Full `ToolsPesterBuildToolchain` deployment proof is an explicit first entry. Its
isolated environment supplies the selected platform and configuration, and the test may
delete/rebuild only `.build/<platform>/<configuration>`; it must never fall back to x64
Debug. After it runs, the runner rebuilds/re-attests that same full profile and
byte-verifies the resulting receipt. It refreshes the cached receipt binding digest
from that verified receipt before recomputing downstream entry fingerprints and the
impact decision; an entry's receipt ID and build-artifact component must describe
the same receipt. It then durably publishes a new create-only
content-addressed decision and switches decision digest, snapshot, receipt, epoch, and
`mutation_state=stable` in one run-state replacement before any artifact reader starts.
A crash before that switch leaves the old decision intact and the epoch unreadable; a
crash after it leaves one reconstructible new state. Rebuild-generated PE/PDB identity is not
required to equal the pre-entry receipt. The re-attestation build is executed under the
same suite-owned build environment as the initial build. In particular, Full preserves
`RSBuildEnableTests=true` and embedded compiler debug information across the writer
barrier regardless of `-SkipBuild`, while retaining final linker PDB output, and restores
the caller environment;
otherwise the receipt can faithfully attest artifacts
that are incapable of executing the declared Full plan. The writer is `ExactArtifact`: exact Resume may
reuse its promoted post-entry receipt/epoch evidence; when it executes, pre-existing
reuse decisions are reset conservatively to current-epoch execution. Failure to produce
or verify the compatible post-entry receipt is blocking.

Run-state records that predate the required `mutation_state` and `decision_digest`
bindings are legacy evidence and cannot be reused as current Startrail authority.

Resume MUST validate and load the origin run state before computing current entry
fingerprints and MUST initialize current planning from the origin terminal artifact
epoch. Starting every Resume at epoch zero makes exact post-writer promotions
unaddressable. Receipt, snapshot, plan, dependency, capability, outcome, and coverage
identity checks still apply independently; an epoch match never weakens those gates.

## Coverage and verdicts

A Fresh or Resume Full verdict requires complete accounting for the canonical plan.
Filters, families, repeat/shuffle probes, capability-reduced runs, and monolithic logical
candidates yield repository `NOT_EVALUATED` when they reduce required coverage.
Artifact-mutating entries declare read/write/produced sets, end the current artifact
epoch, and require re-attestation before later readers.

Aggregate v2 is the primary Full/CI result at `run-all-tests-results.json` and distinguishes
executed, reused, provisional, promoted, invalidated, and incomplete entries. The legacy
v1 projection remains beside it as `run-all-tests-results-v1.json` for history/dashboard
consumers. Reporters, archive validation, MTP live closeout, and GitHub step summaries
accept v1 and v2 during the compatibility window. A complete Fresh/Resume Full can report
repository `PASSED`; Affected/family/filtered/capability-reduced coverage reports
repository `NOT_EVALUATED` while its selected change verdict may pass.

Aggregate state labels are never sufficient evidence. One verified loader recomputes
promotion, evidence, and terminal-result digests; requires digest/filename equality;
checks origin/run, entry, invocation and coverage-plan identity, fingerprint, snapshots,
receipt, epoch, coverage/outcome, exact artifact map, and result-artifact bindings; and
requires a normalized terminal `PASSED` result. Resume, aggregate summaries, duration
history, and prune provenance use this loader. Any missing or semantically inconsistent
record makes aggregate evidence `BLOCKED` and maps to corrupt-evidence exit 5.

### Specialized Terminal VT evidence

Ghostty VT upgrade performance evidence is a governed archive format distinct
from Startrail aggregate reuse. A claim-bearing Terminal archive contains
`results.json`, `trace.txt`, and `perf/perf_metrics.jsonl`; `results.json` must
validate against `Specs/Terminal/TerminalVtUpgradeEvidence.schema.json` and bind
the archive run id, 12-hex machine-profile directory, repository/configuration,
corpus, and runtime identity. `Test-TestRunArchive.ps1` parses every nonblank
JSONL row and requires exactly the declared 200 samples, in sample order, for
each of the three VT metrics. Malformed later rows are evidence corruption even
when the first row and summary are valid.

Admission requires the explicit paired form
`Test-TestRunArchive.ps1 -BaselineRunPath <product-run> -CandidateRunPath
<candidate-run>`. It validates each archive, recomputes nearest-rank
percentiles from the raw rows, binds the common environment, requires distinct
runtime/build identities, verifies control and declared-delta digests, and
applies the 1.10 ceiling from Product p95. A candidate's own threshold field is
not independent authority and cannot make a regressed pair pass.

### Exact Resume

`Run-AllTests.ps1 -Suite Full -ValidationMode Resume -ResumeFrom <run-id-or-directory>`
creates a new direct child run. It accepts only schema-valid immutable promotions whose
entry fingerprint, pre/post workspace snapshot, exact verified build receipt/output
closure, artifact epoch, coverage/outcome policy, runner dependencies, capabilities, and
quarantine identity equal the current run. `Never`, missing, provisional, failed, or
ordinary compatible-input mismatches execute and record their reason. Schema/digest,
cross-record binding, or terminal-result corruption aborts Resume with exit 5 instead of
silently converting corrupt evidence into a fresh execution. Fresh remains the default.
The origin resolver accepts only an exact lowercase run ID or an exact fully qualified
path naming one non-reparse direct child of `runs`; the directory leaf must equal
`run-state.run_id`. Nested paths, case/path aliases, copied-renamed directories, and
junctions are rejected before evidence records are read.

### Affected execution

`Run-AllTests.ps1 -Suite Full -ValidationMode Affected [-ImpactBase <commit-ish>]`
constructs the canonical Full shadow decision, then publishes and executes a filtered
Affected plan containing only governed affected entries. Unknown relevant paths widen to
the Full set. Empty/documentation-only sets are valid selected work but never a repository
health claim. Affected does not read reusable test evidence and cannot replace final Full.

### Commands family boundary

`Run-AllTests.ps1 -Suite Commands -CommandsFamily <name>` launches one of the 14 governed
Commands source families in a new process. Native list-mode membership, the PowerShell
catalog, and the broad Commands union must remain exactly equal with no duplicates or
orphans. Family entries retain `family-process` order and `Never` cache policy. They use
the monolithic `RedSalamander.exe` runtime closure, so no family result or FileOperations-
only repair may contribute an independently reused Commands seal to Full. A future split
requires a separate reviewed payload/runtime identity before `IndependentLogicalArtifact`
can be enabled.

## Scheduling and cost diagnostics

Every entry declares resource classes and an advisory duration key. The runner executes
artifact writers first behind their re-attestation barrier, then cheap disjoint
noninteractive work, then stateful/interactive/performance/network work; those exclusive
classes remain serialized. Historical promoted durations are bounded, median-only
estimates. Missing/corrupt history means unknown cost and never affects affectedness,
fingerprints, execution authority, or verdicts. Explain JSON and console output report
reason codes and costs. Snapshot metrics separately report Git discovery, content
hashing, canonicalization, and total acquisition time.

## Exit categories

Startrail-owned exits are stable and are mapped once from the terminal state after summary
publication: `0` `PASSED`, `1` `FAILED`, `2` `NOT_EVALUATED`, `3` invalid CLI or plan,
`4` blocked or missing attestation, and `5` corrupt evidence. `ExplainPlan` and Affected
therefore return 2 even when their selected work is green; CI and release policy must not
treat 2 as a completed repository gate. Host or parameter-binding failures before the
runner owns control are not remapped.

## Scoped iteration and PR coverage

Active native test sources, helpers, fixtures and seams MUST use `Scope.Tests.Something.h/.cpp`. `Tests/native-test-files.json` owns active membership. Historical archived harnesses and sealed reproduction inputs keep their original identity; they do not define new source naming. Rename project entries, includes, current inventories and callers together. DxUi historical source/test origins retain their original identities and map to current paths separately.

`Test-Changes.ps1` is the ordinary iteration entrypoint. Its default is affected coverage, including committed changes since a local merge base and independent staged, unstaged, deletion, rename-side and untracked discovery. No ref is fetched. Unknown executable inputs widen coverage; prose alone does not require native tests. Explanations name paths, consumers and fallback reasons. Explicit selectors reject unknown scopes. Affected, filtered and environment-reduced coverage MUST NOT be reported as a full repository pass.
The entrypoint forwards explicit `-AllowAlternateVolumeTestRoot` unchanged to the canonical runner in Full and focused
FileOps work. Its default remains disabled. An enabled `path.local.alternate` resource restricts the selected root;
neither that resource nor the forwarding switch supplies authorization to initialize a new drive-root workspace.

Exact local continuation uses the existing explicit Full Resume and its promoted artifact/epoch contracts. Commands families and timing/network-dependent entries retain their declared non-cacheable policies. Affected never consumes reusable evidence.

`-Mode Full` selects the full local obligation. `-Mode PrePush` accounts for full coverage across local execution and the forthcoming PR gate. Delegation requires a clean committed candidate, the enabled GitHub workflow, matching candidate workflow bytes and their reviewed manifest digest, and a matching native profile. Staged, unstaged and untracked inputs retain obligations locally; a dirty tree or changed HEAD during the coverage lookup also rejects delegation. Nightly and weekly jobs are not PR coverage. Workflow/API uncertainty keeps obligations local. A delegated run reports `CI_PENDING`, never repository `PASSED`; the PR must pass its required checks. Main/release acceptance remains separate because its merge tree, configuration matrix or requirements can differ.

RedSalamander defers only entries whose complete portable entry contracts equal the PR plan, retaining writers, in-product cases and extra scenarios locally. Its `-NonInteractive` option omits entries that require desktop-global state and records incomplete coverage; it cannot alter an exact Resume. When every selected entry is omitted (for example FileOps admission coverage with shared foreground invariants) the run executes nothing, lists the omitted entries as `PENDING_INTERACTIVE`, writes no pass claim and ends repository `NOT_EVALUATED` with exit 2. Only `-Suite FileOps` with one exact, case-sensitive filter from `DiscoveryScope_FolderLiveProgress`, `DiscoveryScope_MultipleFoldersLiveProgress`, `FileOps_RepeatDecisionContracts`, `FileOps_ReviewRemediationContracts`, `FileOps_CrossRootLocalCopy`, `FileOps_SmallFileCopyCost`, `FileOps_SelectionShapeLocal` or `FileOps_SelectionShapeSmb`, and no flaky/order proof injection, is admitted without that desktop lease. These owned callback, storage and progress witnesses retain `--selftest-no-activate`; application captures use the canonical owned HWND WGC path; broad suites, families, prefixes, wildcards, multiple-case filters, casing near-misses and unknown filters remain leased. The plan module stays in the recursive runner dependency digest, invalidating prior entry identity after a classification change. Focused successes still report filtered coverage and repository `NOT_EVALUATED`. DxUi foreground suites remain explicit `test.ps1 -Interactive` work after agreement to the time. Noninteractive iteration and CI do not claim those manual gates; no desktop automation is introduced.

`Tests/test-scopes.json` owns the standalone scope and verified PR-profile mapping; RedSalamander's `Tools/validation-impact.json` and canonical plan continue to own its affected execution. Specific standalone test-source rules use their semantic entry tags, while common support and unknown paths retain conservative fan-out. New source/build/runtime/runner dependencies must extend the relevant manifest and focused regression coverage.

CI invokes explicit full gates rather than the affected default. The forthcoming PR executes its candidate workflow; a reviewed local workflow digest and the enabled GitHub workflow gate delegation. Workflow edits invalidate that digest until the scope contract is reviewed and updated. RedXe's independent Windows tooling/skill tests execute once in its tooling job. DxUi's tooling has one Windows host qualification and one Linux portability qualification; its MSBuild-only staging fixture executes once on Windows; native runner/watchdog checks stay profile-specific. RedSalamander source-monitor schedules skip unchanged commits only after the latest completed matching workflow attempt succeeded; manual dispatch and unreadable/failed/cancelled history execute. Scheduled source skips do not establish new environment qualification. Hosted timing investigations still use paired, fixture-matched resource evidence; cached historical observations are not new measurements.

PR delegation MUST account for conditional workflow jobs. DxUi uses the existing NativeScope contract to keep native obligations local when a documentation-only PR skips native CI; always-running Windows/Linux tooling remains covered. Native success identities exclude validator prose while tooling identities include it.

## Reviewed PR workflow execution closure (2026-10-16)

`Get-ScopedPrWorkflowDigest` binds PR delegation to repository-owned execution logic, including reusable workflows and local action contents, tooling, build/dependency configuration, executable script extensions throughout the checkout, test-scope/native inventories and Testing/Build JSON contracts. It deliberately hashes a conservative superset instead of inferring PowerShell imports or shell control flow. Every local `uses` reference must resolve; missing inputs, path escapes, unresolved expressions and reparse inputs fail closed. Local action directories include arbitrary assets such as Dockerfiles. Product C++ edits remain ordinary candidate inputs and do not by themselves change which gate is reviewed.

Text digests normalize CRLF. `Tests/test-scopes.json` participates with only its own `prWorkflowDigest` property removed, avoiding a self-reference while binding coverage declarations. Any change to these execution inputs requires recomputing and reviewing the manifest digest. A mismatch retains local obligations; an enabled forthcoming workflow is still pending acceptance, never already-passed evidence. Clean committed candidate checks before and after GitHub lookup continue to apply.

The scoped wrapper defaults its local comparison to origin/HEAD, the recorded remote default branch. A repository migration updates that symbolic ref with git remote set-head origin -a; the wrapper never fetches or guesses a branch. Tests/test-scopes.json identifies RedSalamanders/RedSalamander and main for matching hosted PR coverage. An unavailable or nonmatching hosted gate retains the local obligation.
