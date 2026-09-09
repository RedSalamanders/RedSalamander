# Repository Tooling Governance

Status: normative repository tooling contract
Last reviewed: 2026-08-28

## Scope and authority

This contract governs scripts, modules, tests, manifests, patches, and support
data under `Tools/`, plus build, CI, release, packaging, and documentation files
that consume them. It defines how the tooling surface stays discoverable,
compatible, testable, and small as files are added or changed.

The checked-in schema-v3 [`Tools/tool-inventory.json`](../../Tools/tool-inventory.json) is
the machine source of truth for file classification. The human entry point is
[`Tools/README.md`](../../Tools/README.md), and
`Tools/Get-ToolInventory.ps1 -Validate` is the canonical drift check. Prose may
summarize the manifest but must not maintain a competing file list.

## Directory contract

Native focus-taking test suites share `Tests/TestSupport/DirectedSelfTestInputWarning.h`
within the unified runner's interactive desktop lease. Keep the existing native
warning implementation and its audit exceptions centralized; do not add a second
PowerShell or test-family warning surface. Its lifetime and activation policy
are defined in `Specs/Testing/Testing_SelfTests.md`.

`Tools/` root is a stable entrypoint surface, not a general implementation
folder. A root file is allowed only when it is one of these:

- a supported human command;
- an entrypoint invoked by MSBuild, CI, packaging, release, or another stable
  repository command;
- a documented compatibility shim;
- a navigation or machine-contract file whose canonical path is the root.

Reusable implementation belongs in a domain module under `Tools/Modules/`.
PowerShell modules use `.psm1`, declare explicit exports with
`Export-ModuleMember`, avoid import-time mutation, and keep private functions
private. Runnable commands use `.ps1`. Function-library `.ps1` files must not be
added beside runnable root commands.

Pester tests live under `Tools/Tests/`. Shared test helpers belong in the
reviewed test-support module; test files must not copy a helper when identical
semantics already exist there. Terminal-specific active implementation stays
under the current Terminal runtime boundary. Sealed selection and qualification
tooling remains explicitly historical and does not masquerade as current
product tooling.

## Inventory contract

Every tracked file under `Tools/` has exactly one effective inventory
classification. Directory rules may cover homogeneous internal collections,
but every root file, public command, compatibility shim, and exceptional file
has an explicit entry. A new file must fail validation until classified.

Each effective record states at least:

- repository-relative `path`;
- `kind` such as command, module, test, data, patch, native harness, project, or
  documentation;
- `visibility` as public or internal;
- `lifecycle` as active, compatibility, historical, generated, or data;
- behavior-focused `purpose`;
- `usedBy` consumers or an explicit manual-use declaration;
- optional `consumerClosure: "derived-references"` for dependency-critical modules;
  every repository PowerShell/workflow file that names that module is then derived and
  must be covered by an exact or reviewed wildcard `usedBy` declaration;
- an explicit `prerequisites` array (including `[]` when none apply) and material
  `sideEffects`;
- persistent `outputs`;
- `ownerSpec`;
- focused `tests` when behavior is executable;
- for every historical record, non-empty `manualEntryPoints` naming exact
  existing historical commands whose manual reproduction/authentication closure
  requires the file;
- `replacement` as one exact existing repository-relative `Tools/` path or null;
  prose about migration or generation belongs in `usage`;
- non-empty `supportPolicy` for compatibility records, stating the retained
  support or exact reviewed removal condition;
- reviewed `namingPolicy` rationale for any public command whose retained name
  does not use canonical `Verb-Noun.ps1` spelling and casing.

The validator fails for unclassified tracked files, stale manifest paths,
duplicate classifications, invalid lifecycle/visibility combinations, missing
required metadata, or a public command whose help contract is incomplete.

## Public command contract

New public commands use the canonical `Verb-Noun.ps1` form. Existing noncanonical
names remain only when the inventory records them as compatibility or reviewed
legacy entrypoints. A rename migrates all known repository consumers and keeps a
thin forwarding shim whenever documented or plausible external use prevents an
atomic removal.

Every public command includes comment-based help with:

- `.SYNOPSIS` and a behavior-focused `.DESCRIPTION`;
- one `.PARAMETER` section for each public parameter;
- `.OUTPUTS`, including whether output is objects, text, files, or none;
- `.NOTES` naming prerequisites, side effects, exit-code behavior, and primary
  consumers;
- at least one safe `.EXAMPLE`.

Mutating commands use `SupportsShouldProcess` when a preview is meaningful and
must not hide source-tree writes, credential changes, network access, process
launches, or destructive cleanup. Exit codes and machine-readable output are
part of the public contract.

Environment variables and command-line paths are inputs, not authorization.
Test tooling must use exact `X:\RedSalamander.Perf` on a fixed local drive, reject
drive roots, UNC roots, arbitrary absolute paths, descendants, alternate names,
sibling-prefix tricks, alternate-data-stream syntax, and reparse-point components,
and require an exact ownership marker before mutation. First initialization and a
second-drive FileOps root require explicit public opt-in. A per-machine
`path.local.alternate` selector may restrict that opt-in to one exact marked
`<AltDrive>:\RedSalamander.Perf` root but cannot authorize arbitrary directories.
Test tooling never creates,
audits, or cleans runtime data outside this root. The maintenance cleaner accepts only
one exact `runs\<runId>` child beneath a marked root; legacy outside-root cleanup is not
part of the repository tool surface.

## Compatibility and moves

Before a tool is renamed, moved, retired, or deleted, its maintainer searches
scripts, modules, MSBuild, workflows, specs, documentation, and tests for direct
consumers. A documented path or one used by external automation is assumed to
have plausible untracked consumers.

A compatibility shim is intentionally small: it resolves the canonical command,
forwards parameters and exit behavior without reimplementing logic, names the
replacement in help and inventory, and has a focused forwarding test. The
inventory `supportPolicy` records whether support is permanent or the exact
release/policy condition for removal. Once that condition is explicitly met,
delete the shim, its inventory row, forwarding-only tests, and stale prose in
the same change rather than retaining an inert placeholder.

Moves of private implementation update all tracked callers atomically. No move
may rewrite frozen evidence bytes, silently change path-derived identities, or
weaken a build/release gate merely to achieve a cleaner directory.

## Current and historical Terminal tooling

The current product Terminal toolchain is the Ghostty private-runtime builder,
its imported identity/archive/policy helpers, current patch, notices, runtime
tests, and the build/package consumers of its product directory. Superseded
Gate-0 and Round-4 candidate selection, evidence collection, and closeout tools
are historical even when retained for reproducibility.

Historical is a retention claim, not a parking category. A historical file is
kept only when it is a manual entrypoint or belongs to the minimal transitive
closure of an exact manual entrypoint recorded in `manualEntryPoints`. The
reviewed roots are `Invoke-HistoricalTerminalTests.ps1`,
`Evaluate-TerminalEngines.ps1`, and
`Test-TerminalEngineRound4Closeout.ps1`. Delete an orphaned historical file and
its inventory/prose references; do not preserve it merely because it once
existed or appears in an old completed plan. Completed plans may retain the
former path and digest as immutable evidence without keeping the file live.

Default CI and Full suites execute
`Tools/TerminalEngine/CurrentToolingProfile.Tests.ps1`. That wrapper discovers
the current `Tools/Tests/*.Tests.ps1` set while excluding every test recorded in
`Tools/TerminalEngine/HistoricalTerminalProfiles.json`; adding a current test
must therefore require no hand-maintained CI list. The default lane retains the
small `HistoricalTerminalProfile.Tests.ps1` integrity/closure guard, which
proves that every sealed file remains at its recorded repository path with its
recorded SHA-256 and is absent from current-profile discovery.

`HistoricalTerminalProfiles.json` is the sole source of Gate0 and Round4
membership. `Invoke-HistoricalTerminalTests.ps1` accepts only an explicit
`Gate0` or `Round4` profile and is the manual archival reproduction entrypoint;
Round4 continues to assert its recorded 239-pass result. The sealed cohort must
not be mechanically reformatted, retagged, or migrated with current test
helpers. Historical exact-byte inputs remain explicit and are never generated
from a mutable current test manifest.

The Gate0 workflow remains an exact-byte historical contract: its supported
execution is manual dispatch or the sealed `codex/terminal-engine-gate0` branch
with the explicit `[run-gate0]` marker. The Round4 workflow remains manual-only,
and its reviewed source shape is frozen. Neither workflow is retrofitted to
consume the later archival launcher or manifest; the manifest and integrity
test describe the supported repository-side reproduction boundary without
rewriting historical evidence.

An active command or module that names focused tests must name at least one test
in the current profile; a sealed historical cohort alone does not establish a
current behavior contract. Dormant generic lifecycle rows remain outside the
active classification and retain their sealed reproduction tests plus explicit
support policy.

Generic lock-based Terminal lifecycle commands that are not part of the product
build remain stable-path compatibility entrypoints, descriptively classified as
internal/historical/future-upgrade tooling, until an explicit reviewed lock
reactivates that workflow or an external-use decision permits removal. This is
encoded through `visibility`, `purpose`, `prerequisites`, and `supportPolicy`, not
a separate lifecycle enum value. Their absent default lock and non-product status
must remain visible in inventory and documentation; they must not be presented as
the canonical Ghostty product lifecycle.

## Shared implementation

Repository scans, archive discovery, result formatting, suite planning,
sandbox management, quarantine policy, and other cross-command behavior have a
single reviewed implementation per semantic contract. Front ends may request
different policies, such as returning null versus throwing for a missing run,
but express that difference as an explicit option rather than copied code.

Before adding a local helper, search `Tools/Modules/`, `Tools/Tests/TestSupport.psm1`,
`Tests/TestSupport/`, and `Common/`. A local variant is allowed only when its
layer, ownership, failure, compatibility, or performance semantics differ; the
difference must be documented and tested.

`Test-TestRunArchive.ps1` is the single thin public archive-validation
entrypoint. Shared discovery, size policy, complete JSONL parsing, and
area/schema dispatch live in `Tools/Modules/Testing/TestRunArchive.psm1`.
Terminal VT admission uses the same command's paired
`-BaselineRunPath`/`-CandidateRunPath` mode; a second Terminal-only validator or
percentile implementation is prohibited. Its focused tests must cover a valid
pair, a malformed later JSONL row, schema failure, missing required files, and
a candidate p95 above the Product-derived 1.10 ceiling.

`Measure-SourceLines.ps1` owns the source-line implementation. Repository scope
owns the categorized report; Solution scope owns raw solution-backed cloc output.
`InvokeSolutionCloc.ps1` remains a supported public manual compatibility path
because developers use that command directly. It is a thin forwarder to
`Measure-SourceLines.ps1 -Scope Solution`: it preserves `SolutionPath`,
`ClocPath`, trailing raw cloc arguments, output, and failure status without
copying solution parsing or directory-selection logic. New callers should use
`Measure-SourceLines.ps1`; retaining this reviewed hand-use path does not permit
another source-line implementation or alias.

## Build, workflow, and cache closure

A cache key, build identity, release manifest, or workflow input list that
represents a tool includes its complete direct dependency closure: entry script,
imported modules, current patches, schemas, notices, version pins, and other
behavioral inputs. In particular, the Ghostty private-runtime cache key equals
the builder plus every module it directly imports, the current private-runtime
patch, and the notices file. `ReleaseWorkflowPolicy.Tests.ps1` derives the
builder imports from source and compares that set with the workflow key so a new
helper cannot silently escape invalidation. Superseded patches or modules are
removed from current keys and classified historical when retained.

Workflow test lists have one current declarative owner. Frozen historical lists
remain explicit and immutable where their exact identity is evidence; current
discovery must not replace them.

An active tool that reads or replaces compiled artifacts reuses
`Tools/Modules/Build/ArtifactOperationLock.psm1` and supplies explicit platform
and configuration evidence. Same-output operations serialize, while disjoint
artifact profiles and worktrees remain independent except for a separately
identified shared writable resource. Target, project, and suite labels are
diagnostics rather than lock partitions. A new shared resource receives the
narrowest reviewed lock that protects it; it does not justify blocking every
build and test profile.

The current explicit shared resource is the platform-scoped vcpkg install root,
`.build\vcpkg_installed\<platform>`, including its triplets and mutable metadata.
x64 and ARM64 never share that metadata root. MSBuild and
vcpkg-install phases hold its repository/platform `shared-dependency` scope,
which crosses configurations but not platforms or worktrees. Ordinary test
execution does not hold it. A multi-triplet installer acquires both scopes in a
stable order before mutation, and same-platform residual checks ignore
configuration when protecting that shared dependency.

Packaging uses the other reviewed shared scope. Build orchestration acquires
repository-wide `coordination=packaging` only after compilation and keeps it
through MSIX/MSI/ZIP/winget generation because platforms share installer inputs
and `.build\AppPackages`. Standalone packaging mutators enter the same scope;
ordinary build/test work remains outside it. Abandoned packaging state fails
closed for explicit restoration and is not cleared by a profile rebuild.
Standalone packagers that read compiled outputs acquire the artifact-profile
scope first and packaging second, then release in reverse. A dependency-capable
MSIX project acquires its platform `shared-dependency` scope after packaging and
only around contained MSBuild. Repository-scoped locks protect only reviewed
repository descendants: configurable package outputs remain beneath
`.build\AppPackages`, MSIX manifest mutation remains beneath `Installer\msix`,
and machine-global WiX extension installation is a prerequisite rather than an
ordinary packaging side effect. Package files are produced as unique same-directory
staged files, validated before publication, and published with the shared guarded
atomic-file helper only after immediate ancestor/final reparse revalidation;
force-overwriting a live package path is outside the reviewed tooling contract.

Process names are never sufficient ownership evidence. Residual build-tool
checks use boundary-aware repository plus profile evidence, and exact-output
application checks compare normalized executable paths. An independently
launched process is reported and left running; automatic termination is limited
to descendants launched and contained by the current operation. Test plan
authors mark only entries that touch desktop-global state with
`RequiresInteractiveDesktop`, and that marker must survive repair-plan
projection so the session mutex stays inside the entry execution boundary.
GUI-capable entries also declare activation independently: focus-independent
entries pass the harness's governed no-activation selector, while only reviewed
focus/caret/menu/keyboard-routing entries may omit it. Splitting a harness into
noninteractive and interactive entries must preserve one shared runtime-artifact
identity and explicit coverage-group arguments.

### Operation Startrail process impact

[`Operation Startrail`](../Plans/Done/Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md)
records the gated implementation history. Fresh remains the default and final-closeout
authority; explicit local Resume and Affected use the narrower active contracts in
`Testing_ValidationEvidence.md`.
The completed
[`cross-domain remediation`](../Plans/Done/CodeReview_CrossDomain_Remediation_2026-08-10.md)
records the post-completion `L4-BUILD-01..08` and `L4-STAR-01..16` repairs;
the Done Startrail record is not a competing authority.

Startrail schemas, modules, commands, and behavioral data are governed tooling inputs.
They require inventory classification and must participate in every applicable cache key,
fingerprint, workflow manifest, and consumer dependency closure. Consumer review is
repository-wide and includes test-run reporting, archive validation, live MTP closeout,
CI step summaries, documentation examples, and source-contract tests.
The canonical runner fingerprint derives its recursive script/module/schema closure from
`Tools/Run-AllTests.ps1`; hand-maintained parallel lists are forbidden. Missing imports
must enter the closure automatically, while unrelated files must remain outside it.

Machine-readable inputs under `Specs/` are executable contract data, not automatically
documentation. The impact manifest must classify each reviewed schema, settings/theme
input, performance budget, corpus, and governance map with its real consumers. Unknown
JSON or JSON5 beneath `Specs/` widens to Full until reviewed; the documentation-only
fallback is limited to prose such as Markdown.

Any test or tool that can replace, delete, stage, or otherwise mutate compiled/runtime
artifacts must declare that behavior in the validation plan and inventory. A hidden
artifact mutator cannot share a trusted epoch with evidence that assumed the prior
artifact set. `Affected`, `Resume`, Commands-family, baseline, and explanation surfaces
must retain complete help, stable exit categories, output-root contracts, and focused tests.

The Full runner's initial test-enabled build and post-writer full-profile re-attestation
use embedded compiler debug information while retaining final linker PDB output. This
avoids MSVC's process-global compiler-PDB service, so disjoint worktrees remain
concurrent; ordinary `build.ps1` and CI builds retain ProgramDatabase by default.
The governed MSVC 14.51 workaround disables `/MP` only for
`Common/SearchTextHelpers.cpp` in Embedded mode; it MUST NOT serialize projects,
profiles, or worktrees.

Normative specs and skills must not describe unimplemented reuse as current behavior.
Changes to an active validation surface require inventory reconciliation, consumer
migration, archive validation, Full validation proportional to risk, portable CI policy
proof, and same-machine performance evidence when planning cost is affected.

## Required change workflow

Any change that adds, modifies, moves, or removes a tool performs all applicable
steps in the same review unit:

1. search and record callers and path/identity consumers;
2. choose the correct root entrypoint, module, test, data, or historical
   location;
3. update inventory metadata and `Tools/README.md`-derived navigation;
4. update public help, compatibility forwarding, and side-effect declarations;
5. add or update focused behavior, inventory, help, caller, and policy tests;
6. update owning normative specs and contributor documentation;
7. run inventory and focused tests, then the appropriate repository suite;
8. update cache/manifest closure whenever an import or behavioral input changes.

Review must reject a new unclassified file, a public command without complete
help, a copied helper without documented semantic need, a path-breaking move
without compatibility treatment, a historical record without an exact manual
retention root, an artifact-mutating production caller without explicit scope,
process-name-only blocking or unowned process termination, a desktop-global test
without its plan marker, or an input-list change without closure proof.

## Validation

The minimum tooling-governance checks are:

```powershell
.\Tools\Get-ToolInventory.ps1 -Validate
Invoke-Pester .\Tools\Tests\ToolInventory.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\ToolingGovernance.Tests.ps1 -PassThru
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

Run focused tests for each touched module or entrypoint. Tooling closeout also
runs `.\Tools\Run-AllTests.ps1 -Suite Full`; historical Terminal behavioral
reproduction uses its explicit archival profile rather than extending every
default run.
