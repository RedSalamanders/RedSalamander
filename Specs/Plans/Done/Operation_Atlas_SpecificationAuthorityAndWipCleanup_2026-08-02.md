# Operation Atlas — Specification Authority and WIP Cleanup

> **Executor instructions:** This is primarily a documentation-information-
> architecture migration. It permits only those product/test corrections for
> which the consistency ledger conclusively proves code is the stale side of a
> current contract. Read this file in full, execute the waves in order, and run
> every gate. When a requirement is found in
> more than one document, first identify its current implementation/test owner;
> preserve the strongest current contract exactly once and replace the other
> copies with links. Never resolve uncertainty by deleting text.
>
> **Drift check (run first):**
>
> `git diff --stat fc4ba4a63..HEAD -- Specs Tools/Tests Tests/README.md`
>
> If any in-scope path changed after `fc4ba4a63`, regenerate the inventory and
> revalidate the affected disposition row before editing it.

## Status

- **Priority:** P1 documentation governance and execution safety
- **Effort:** L, expected as 6–9 reviewable commits
- **Risk:** HIGH if executed as a bulk deletion; MEDIUM with the ordered gates
  below
- **Depends on:** embedded-terminal checkpoint `fc4ba4a63`
- **Planned at:** commit `fc4ba4a63`, 2026-08-02
- **Execution state:** COMPLETE — Atlas closeout gates passed on 2026-08-04;
  this file's final location under `Specs/Plans/Done/` records the
  content-identical closeout move
- **Category:** documentation, technical debt, developer experience

> **Progress checkpoint — 2026-08-03:** Wave 3 is complete. Terminal v1 is
> implemented, its durable requirements are reconciled into
> `Specs/Terminal/Terminal_EmbeddedPlugin.md` and the current executive section
> of `Specs/Terminal/TerminalEngineDecision.md`, and all 13 former Terminal WIP
> plans are archived under `Specs/Plans/Done/`. The dispositions and Wave 3
> procedure below have been updated to the completed outcome; Atlas remains WIP
> for its non-Terminal waves.

> **Revalidation checkpoint — 2026-08-04:** The original Atlas remarks are
> useful findings, not accepted conclusions. Current review found three plan
> defects that must be corrected during execution: the frozen Done-tree rule
> predates the already-approved Terminal archive additions, the position-only
> authority wording omits repository-level contracts such as `AGENTS.md` and
> machine-consumed schemas/manifests, and the original scope incorrectly forbids
> code changes even when code is the stale side of an unambiguous contract.
> The checklist and gates below now require evidence for every disposition and
> permit narrowly scoped code/test corrections when intent is mechanically
> established. Unclear product or architecture choices are routed to a new WIP
> decision file rather than guessed.

> **Closeout checkpoint — 2026-08-04:** Atlas-owned specification, inventory,
> documentation-drift, build-contract, localization, Win32-audit and test-harness
> gates are green. The final Full aggregate is not falsely reported green:
> desktop-interactive failures ran while Windows `LockApp.exe` owned the
> foreground and the independently modified Terminal workstream has nine
> frozen-governance failures. Current Terminal authority now says verification
> is reopened and routes remediation to `J21-TERM-07`; those failures do not
> leave an Atlas consistency discrepancy or reactivate the archived Terminal
> plans.

> **Post-closeout correction — 2026-08-04:** Atlas originally overrode the
> explicit human-managed boundary in `Notes_Scratch.md`, inspected its contents
> and deleted it. That was incorrect. The file is restored and remains
> untouched. Current authority and inventory tooling classify that exact path
> as `HumanManagedNote`, outside WIP indexing, specification authority and
> content analysis. No current requirement or decision relies solely on the
> earlier invalid inspection.

## Executive outcome

Atlas starts from an explicit authority rule: current, explicitly classified
contracts under `Specs/<Domain>/` are the primary product/domain authority.
Repository-level governing documents such as `AGENTS.md` and
`Specs/Testing/*`, public ABI headers, schemas, resources, runtime manifests and
other machine-consumed contracts remain authoritative within their declared
boundaries. A path alone never makes prose correct or normative.
`Specs/Plans/WIP/` is non-normative planning material, even when a plan is
detailed or implementation-ready.
`Specs/Plans/Done/`, `Specs/Reviews/` and root `plans/` are non-normative
history/evidence. A WIP statement never overrides a normative contract and
never proves that behavior is implemented. If a normative contract is stale,
the discrepancy must be resolved against implementation, tests and intended
product behavior; copying a WIP proposal into it is not a resolution.

After this operation:

1. Every normative document has been checked against its implementation, tests,
   sibling contracts and public behavior. There are no unresolved
   contradictions, stale requirements or unverified normative files.
2. `Specs/<Domain>/` contains current, binding contracts rather than completed
   implementation diaries, proposed work, fixed-bug narratives, or duplicated
   test-run history.
3. `Specs/Plans/WIP/` contains one accurate, explicitly non-normative queue of
   unfinished work. Every
   active plan has one owner, explicit remaining work, durable-spec targets,
   verification, and STOP conditions.
4. Superseded WIP bulk is removed without breaking immutable historical links:
   a referenced old path becomes a compact in-place tombstone; an unreferenced
   file may be removed only after its unique live requirements are routed.
5. Every pre-Atlas file in `Specs/Plans/Done/` is byte-for-byte unchanged. The
   only authorized additions are the 13 already-completed Terminal plans and,
   after every Atlas gate passes, this Atlas plan itself.
6. A read-only inventory and Pester contract prevent WIP/index drift, planning
   text leaking back into normative specs, and broken local links.

This operation deliberately optimizes authority and navigability, not the
smallest possible byte count. A current requirement, security rationale,
compatibility rule, legal notice, schema, tested allowlist, or reproducibility
input is useful even when old. Closed chronology and duplicated prose are not.

## Protected historical baseline and authorized closeout additions

`Specs/Plans/Done/` is non-normative history. Existing entries are completely
out of scope:

- no deletions;
- no moves out of the folder;
- no renames;
- no content or line-ending edits;
- no link repairs inside Done.

The only additions authorized by completed closeout are:

- the 13 Terminal plans already added between `fc4ba4a63` and `6aecfde8e`;
- this Atlas plan, moved without content changes from WIP only after all final
  checks pass, as explicitly requested on 2026-08-04.

The original pre-Terminal and approved post-Terminal tree identities are:

```text
Specs/Plans/Done
pre-Terminal fc4ba4a63: 606951da84f3c5bee89e1f894df2484b70721718
post-Terminal 6aecfde8e: cfd921572d9e8b18c4bf0ec5eb044bbb8de280a5
```

Before the final Atlas move, every cleanup commit must prove the post-Terminal
tree is unchanged:

```powershell
$baseline = git rev-parse 6aecfde8e:Specs/Plans/Done
$current = git rev-parse HEAD:Specs/Plans/Done
if ($baseline -cne $current) { throw "Specs/Plans/Done changed." }
```

After the final move, tree equality is intentionally impossible. The inventory
gate instead compares every path/blob present at `6aecfde8e` and requires exact
identity, then permits exactly one additional path:
`Specs/Plans/Done/Operation_Atlas_SpecificationAuthorityAndWipCleanup_2026-08-02.md`.

This protection is an Atlas audit boundary, not a change to the repository's
ordinary closeout policy.

## Current inventory at `fc4ba4a63`

| Surface | Current shape | Classification for Atlas |
|---|---:|---|
| Normative/root/domain Markdown outside Plans, Reviews and TestRuns | 64 files, 2,318,104 bytes; includes two mockup-local Markdown files | Audit every file; mockup-local files are prototypes, not normative |
| `Specs/Plans/WIP` | 30 Markdown files, 3,003.9 KiB | Primary cleanup target |
| `Specs/Plans/Done` | 120 Markdown files, 6,415.4 KiB | Pre-Terminal immutable history; zero edits/deletions, with later closeout additions tracked separately |
| `Specs/Reviews` | 26 Markdown files | Historical snapshots; out of scope |
| root `plans/` | 31 Markdown files | Existing historical snapshot; out of scope except replacing its live WIP routing |
| `Specs/TestRuns` | 4,039 files, about 4.48 GB in the local checkout | Pinned evidence; out of scope |
| `Specs/Terminal` | 2 Markdown and 33 JSON schema/fixture/evidence files | Separate current product contract from retained Gate-0 evidence |
| `Specs/Themes` | 10 runtime themes plus legal notice | Pinned/runtime/legal; unchanged |
| `Specs/Mockups` | 32 prototype/source/image files | Prototype estate; unchanged |
| `Specs/Testing/FolderViewPerfBudgets.json5` | one live budget file | Runtime test input; unchanged |

At this baseline the WIP index was wrong: its N3 entry said no engine qualified,
directed execution to WezTerm Round 5, and marked product plans blocked. That
historical routing was corrected during Terminal closeout; it is not a claim
about the current 2026-08-04 index.

## Authority taxonomy

Every tracked specification artifact must have exactly one class:

| Class | Location and meaning | Allowed content |
|---|---|---|
| **Normative** | Explicitly classified current contracts, primarily `Specs/<Domain>/*.md`, plus bounded repository guidance and machine-consumed contracts | Current behavior, ABI, validation, security, ownership, accessibility, performance, build/package and release contracts |
| **Active plan** | `Specs/Plans/WIP/*.md` and indexed ACTIVE | Non-normative execution proposal containing only unfinished work; current facts may be summarized, not replayed |
| **Hold/decision** | WIP and indexed HOLD or DECISION | A bounded unresolved product/architecture decision with explicit promotion criteria |
| **Retired tombstone** | Original WIP or normative path retained for inbound links | At most 40 lines: retired state, reason, successor, authoritative targets, baseline commit |
| **Historical** | `Specs/Plans/Done`, `Specs/Reviews`, root `plans/`, retained Gate-0 evidence | Evidence only; never a live queue or current contract |
| **Evidence/runtime data** | TestRuns, schemas, budgets, themes, fixtures, legal notices | Machine/runtime/legal inputs; not prose authority |
| **Prototype** | `Specs/Mockups` | Design exploration; not product behavior |

### Normative document rule

A normative document may describe the implemented contract and bounded rationale.
It is authoritative only for current behavior in its declared ownership
boundary. It must be verified against implementation and tests; location alone
does not prove accuracy.
It must not contain:

- an implementation checklist;
- active TODO/Future Enhancements sections;
- a proposed phased plan;
- a closed bug diary;
- long test-run chronologies or repeated run IDs;
- an obsolete target state presented beside current behavior;
- requirement wording owned more precisely by another normative document.

Short compatibility history is allowed only when removing it would make the
current contract ambiguous. Evidence belongs in TestRuns or historical plans;
the normative spec links to the evidence once.

### Active WIP rule

WIP is always non-normative. It may propose a future contract, but that proposal
must be labelled as unimplemented until code and tests establish it. On
conflict, the current normative contract wins until an explicit reconciliation
changes implementation and/or the normative document. At feature closeout,
durable implemented behavior is merged into `Specs/<Domain>/`; only then may the
plan be treated as historical evidence.

Every ACTIVE plan must contain:

- status and priority;
- planned-at commit and drift command;
- one-sentence ownership boundary;
- exact remaining outcomes;
- authoritative specs to update;
- implementation/test files in scope;
- dependency order;
- machine-checkable verification;
- done criteria and STOP conditions.

Closed steps are summarized in no more than one checkpoint section. The plan is
not a second normative spec.

### Tombstone rule

Before removing obsolete WIP content, calculate inbound links from normative,
WIP, Done, Reviews, root `plans/`, tools, tests, and docs. Because Done cannot be
edited, any path referenced by Done remains at its original location as a
compact tombstone. A tombstone records:

- original title;
- `RETIRED` or `SUPERSEDED` status;
- retirement date and commit;
- reason;
- current active owner, if any;
- authoritative current spec;
- `git show fc4ba4a63:<old-path>` as the full historical retrieval command.

An old file may be deleted only when the inventory reports zero inbound local
links and the requirement ledger proves it owns no unique remaining work.

## Primary gate: normative correctness and contradiction audit

### Current answer

At `fc4ba4a63`, the proposition “all normative specs are correct and without
contradiction” is **UNPROVEN**. Reconnaissance already found proposal leakage,
history mixed into current contracts, duplicated ownership and one document
whose current executive decision is followed by a large, explicitly historical
no-winner decision. Atlas must not describe the normative estate as clean until
the audit below has completed.

### Required status model

Every baseline normative Markdown document receives one final consistency
status:

| Status | Meaning | Required disposition |
|---|---|---|
| `VERIFIED_MATCH` | Current normative statements agree with implementation, tests and sibling specs | Retain or editorially compact without changing behavior |
| `SPEC_STALE` | Implementation/tests establish current intended behavior but the spec describes an older contract | Correct the normative spec and add/retain a drift guard |
| `CODE_STALE` | Normative behavior is still intended and mechanically established, but implementation or tests do not meet it | Correct code/tests in a bounded change with the applicable domain skills and verification; create a separate WIP decision only when intended behavior is unclear |
| `CROSS_SPEC_CONTRADICTION` | Two normative documents claim incompatible current behavior or ownership | Identify one owner and reconcile both before cleanup |
| `PROPOSAL_LEAK` | Future/unimplemented work is presented in a normative location | Move the proposal to non-normative WIP; keep only implemented behavior |
| `UNVERIFIED` | Evidence has not yet been checked | Blocking at closeout; no normative document may retain this status |

Legal notices and prototype/evidence files are classified `NOT_BEHAVIORAL` and
are not counted as normative. Of the 64 baseline Markdown files outside Plans,
Reviews and TestRuns, two under Mockups are prototypes; every remaining
authoritative-or-legal file must be classified, and every actual normative file
must receive one of the six statuses above.

Neither prose nor source code is automatically “the truth.” For each disputed
contract, compare:

1. the normative owner and every sibling spec that restates the behavior;
2. production implementation and public ABI/schema/resources;
3. deterministic tests and source-contract checks;
4. user-visible documentation and compatibility promises;
5. recent accepted commits only as rationale, never as a substitute for the
   current contract.

If code/tests clearly and consistently establish already-shipped intended
behavior, Atlas corrects stale prose. If a current normative contract and
supporting public behavior/tests clearly establish intended behavior, Atlas may
correct stale code or tests as explicitly requested on 2026-08-04. Such a
change must follow the applicable repository skills, add regression coverage,
perform performance validation when the affected path is performance-sensitive,
and remain separately reviewable. If intent is not mechanically established,
record the exact ambiguity in a new WIP decision plan and do not guess.

### Ownership and precedence matrix

| Contract family | Primary normative owner | Implementation/evidence anchors | Contradiction check |
|---|---|---|---|
| Public plugin ABI and lifecycle | `Specs/Plugins/Plugins_PluginAPI.md` plus plugin-specific spec | `Common/PlugInterfaces/*.h`, plugin factories/manager, ABI/source-contract tests | Headers, ownership, quiet point, versioning and prose agree |
| Settings and Preferences | `Core/Core_SettingsStore.md` and `UI/UI_PreferencesDialog.md` | `SettingsStore.schema.json`, `Common/Common/SettingsStore.cpp`, `RedSalamander/Preferences*.cpp`, Preferences selftests | Defaults, validation, migration, partial failure and UI exposure agree |
| Commands, menus and shortcuts | `UI/UI_CommandMenuKeyboard.md` | `RedSalamander/CommandRegistry.cpp`, `ShortcutDefaults.cpp`, resources and command selftests | IDs, routing, default gestures, enablement and localized labels agree |
| File operations and bridge | `FileSystem/FileSystem_FileOperations.md` and `Core/Core_FileSystemBridge.md` | `FolderWindow.FileOperations*.cpp`, filesystem plugins and FileOperations selftests | Source/destination identity, conflicts, cancellation, progress and result semantics agree |
| DxUi and window behavior | `UI/UI_DxUi*.md` plus the owning window spec | `Common/DxUi/*`, `Tests/DxUiTests/*` and window selftests | Focus, UIA, DPI, theming, input, layout and teardown agree |
| Test and performance policy | `Testing/Testing_SelfTests.md`, `Testing_TestCoverage.md` and `Testing_PerformanceValidation.md` | `Tools/TestInventory.ps1`, `Get-TestInventory.ps1`, `Run-AllTests.ps1`, binaries and archived evidence | Suite/case inventory, skip policy, budgets and evidence rules agree |
| Installer/runtime dependencies | `Specs/Installer/*.md` | root `RuntimeDependencies.props`, `Tools/RuntimeDependencies.psm1`, packagers and smoke tests | Build staging, ZIP/MSI/MSIX contents, removal and clean-extraction behavior agree |
| Embedded terminal | `Terminal/Terminal_EmbeddedPlugin.md` and the current executive section of `Terminal/TerminalEngineDecision.md` | `Common/PlugInterfaces/Terminal.h`, `Plugins/Terminal/*`, `Build-GhosttyTerminalRuntime.ps1`, lifecycle/dependency tests and TestRuns | Selected engine, private boundary, current first slice and release-open gates agree |

### Normative audit execution matrix

The raw record is written to
`.build/SpecInventory/normative-consistency.json`. The executor also appends a
human-reviewable `Normative audit results` section to this plan containing
per-domain counts, each non-`VERIFIED_MATCH` file, exact conflicting statements,
the chosen owner, evidence, disposition and follow-up plan. Raw reports are
generated artifacts; this WIP plan remains the review ledger and remains
non-normative.

| Domain | Files audited | Mandatory implementation/test sampling |
|---|---|---|
| Root/Core | `Specs/README.md` and every `Specs/Core/*.md` | Shared helpers, SettingsStore, startup/bootstrap, search, monitor, connection and compare owners plus focused selftests |
| FileSystem | Every `Specs/FileSystem/*.md` | Corresponding plugin, host bridge/file-operation callers, provider selftests and failure/cancellation paths |
| Installer | Every `Specs/Installer/*.md` | Runtime manifest, ZIP/MSI/MSIX tooling, clean extraction and signed/unsigned package contracts |
| Plugins | Every `Specs/Plugins/*.md` | Public headers, factory/export implementation, host discovery/configuration and plugin-specific tests |
| Terminal | Both `Specs/Terminal/*.md` current sections | Terminal ABI/plugin/runtime scripts, exact lock/hash records, lifecycle tests and release-open gate list |
| Testing | Every `Specs/Testing/*.md` | Actual inventory scripts, suite orchestration, case enumeration, skip rules, performance evidence and CI entry points |
| UI | Every `Specs/UI/*.md` | Owning window/control implementation, command/resources/settings wiring, UIA/DPI/theme/keyboard selftests |

No normative content is deleted, split or rewritten before its audit row records
the evidence and status. Editorial cleanup follows correctness; it never
substitutes for it.

## Evidence-backed findings

| ID | Finding | Evidence at baseline | Required response |
|---|---|---|---|
| ATLAS-00 | Normative correctness is not yet proven | No complete per-file code/test/cross-spec ledger exists, and reconnaissance found proposal/history leakage | Complete the primary audit; close with zero `UNVERIFIED`, `SPEC_STALE`, `CODE_STALE`, `CROSS_SPEC_CONTRADICTION` or `PROPOSAL_LEAK` rows |
| ATLAS-01 | WIP routing was stale — **closed 2026-08-03 for Terminal** | The index now removes Terminal from the active queue and points to current domain authority | Keep the rebuilt routing rule while the remaining non-Terminal inventory is cleaned |
| ATLAS-02 | Terminal planning had multiple overlapping generations — **closed 2026-08-03** | All 13 plans are archived under Done with completion/supersession status; no Terminal WIP file remains | Preserve current requirements in the Terminal domain and treat the archived plans as non-normative history |
| ATLAS-03 | WIP mixes live plans, batons, a human-managed note and roadmaps | `Notes_Scratch.md` explicitly declares a do-not-read/edit/delete human boundary; two continuation batons duplicate master owners; WhimFiles is a roadmap | Respect the exact human-managed exclusion; triage only planning artifacts; no unindexed live owner remains |
| ATLAS-04 | Normative specs contain plans and historical diaries | `UI_CommandMenuKeyboard.md:1173` Proposed plan; `FileSystem_FileOperations.md:916` Future Enhancements and two plan appendices; `UI_NavigationView.md:62` recent improvements and `:1991` future work | Move remaining work to WIP and retain only current contracts |
| ATLAS-05 | Historical/migration documents sit beside normative contracts | `UI_PreferencesDialog_MigrationHistory.md` contains active follow-ups and phased plans; three visible-surface audit files carry long closeout chronology | Migrate live obligations, then compact to current allowlist or tombstone |
| ATLAS-06 | Test coverage prose is an oversized hand-maintained inventory | `Testing_TestCoverage.md` is 3,644 lines/303,699 bytes; `TestInventory.ps1` and test binaries already enumerate cases | Keep policy and high-risk coverage contracts; generate case inventory under `.build` |
| ATLAS-07 | Current docs lack a complete IA/link guard | `DocumentationDriftContracts.Tests.ps1` checks selected known contracts but not WIP coverage, classification, or all local links | Add one reusable inventory module and focused Pester suite |
| ATLAS-08 | A naïve Markdown-link regex is unsafe | The reconnaissance regex misread `](` inside fenced C++/PowerShell examples as links | Scanner must ignore fenced and inline code before resolving links |
| ATLAS-09 | Terminal Gate-0 data is historical but mechanically referenced | 33 JSON files are consumed by candidate evaluator tools/tests and cited by immutable history | Retain and classify them; do not delete based only on product winner status |

## 2026-08-04 remark-by-remark review

The findings above were re-read as hypotheses against the current checkout.
The verdict column distinguishes agreement with the observation from completion
of its remediation.

| Remark | Review verdict | Current evidence / correction |
|---|---|---|
| Authority is primarily under `Specs/<Domain>/` | **AGREE WITH CORRECTION** | Domain specs are the primary prose owners, but `AGENTS.md`, public headers, schemas, resources, manifests and test policy are authoritative in bounded scopes; location alone is insufficient. |
| WIP and Done are non-normative | **AGREE** | This matches `Specs/README.md` and repository closeout policy. WIP may summarize current facts but cannot establish them. |
| Done must stay at tree `606951...` | **DISAGREE; CORRECTED** | Current approved post-Terminal tree is `cfd921...`; all 120 pre-Terminal blobs and 13 Terminal additions remain protected, and the final Atlas move is explicitly authorized. |
| Atlas must never change product code | **DISAGREE; CORRECTED** | The 2026-08-04 instruction explicitly says the stale side may be code. Bounded code/test fixes are permitted when intent is mechanically established; ambiguous changes become a separate decision WIP. |
| ATLAS-00 | **AGREE; CLOSED 2026-08-04** | `Specs/NormativeConsistency.json` now records all 59 current normative candidates with implementation, test, sibling-spec, reviewed-section and reviewer evidence; strict inventory reports zero unresolved rows. |
| ATLAS-01 | **AGREE; CLOSED 2026-08-04** | The exact 20-file WIP inventory is indexed once each as 14 ACTIVE, 2 HOLD, 2 DECISION and 2 RETIRED. Terminal history is not a live owner; the unrelated July-21-to-August-4 review remains one explicit active owner. |
| ATLAS-02 | **AGREE WITH QUALIFICATION; CLOSED 2026-08-04** | The 13 Terminal plans remain protected history. Current Terminal defects are owned only by the independent July-21-to-August-4 review and do not reactivate or duplicate the archived plans. |
| ATLAS-03 | **AGREE WITH CORRECTION; CLOSED 2026-08-04** | The original deletion was wrong. `Notes_Scratch.md` is restored, untouched and explicitly excluded as `HumanManagedNote`; two batons are bounded RETIRED tombstones, roadmaps/choices are DECISION or HOLD, and every actual planning owner is indexed. |
| ATLAS-04 | **AGREE BUT ORIGINAL EXAMPLES WERE INCOMPLETE; CLOSED 2026-08-04** | The complete normative scan also found and resolved proposal/stale leakage in Monitor, SettingsStore, SharedHelpers, Batch Rename, DxUi, FolderView, FolderWindow, Preferences, VisualStyle, PluginAPI and VFS. Unprovable choices moved to the Atlas decision queue. |
| ATLAS-05 | **AGREE; CLOSED 2026-08-04** | Preferences migration history and the FTP performance comparison are classified historical tombstones; the typography audit is a compact current allowlist. Existing compact native/comctl allowlists remain normative. |
| ATLAS-06 | **AGREE; CLOSED 2026-08-04** | `Testing_TestCoverage.md` is now a 123-line risk/ownership contract, while `Get-SpecInventory.ps1` emits the exhaustive `.build/SpecInventory/test-inventory.json` from current tooling and binaries. |
| ATLAS-07 | **AGREE; CLOSED 2026-08-04** | The reusable inventory module, CLI, focused 21-fixture suite, authority catalog and documentation-drift integration are implemented. |
| ATLAS-08 | **AGREE; CLOSED 2026-08-04** | The scanner masks fenced and matching-delimiter inline code; adversarial fixtures cover C++, PowerShell, Markdown, fake links, duplicate slugs and root escape. |
| ATLAS-09 | **AGREE; CLOSED 2026-08-04** | Gate-0 schemas/fixtures are classified historical data, remain byte-identical, and retain their evaluator/test and immutable-history consumers. The current Terminal executive contract is reviewed separately from the protected historical block. |
| Ordered governance-first execution | **AGREE** | Characterization and evidence ledgers precede deletions or compaction. |
| Every disposition in the two maps is final | **DISAGREE WITH THE ORIGINAL MAP; CORRECTED AND CLOSED** | The maps were proposals. Evidence required retaining Emberglass and WarpDrive with explicit size waivers, keeping Tailwind as subordinate HOLD, and treating Rosetta as ACTIVE. The executed disposition ledger below records the final result. |

## Atlas completion checklist

This is the live execution checklist requested on 2026-08-04. A checked item
means the cited evidence exists in the current checkout; narrative confidence is
not enough.

### A. Baseline and plan correctness

- [x] Read Atlas in full and rerun its drift/tree checks against current HEAD.
- [x] Record the dirty worktree and preserve unrelated user-owned Terminal/S3
  review changes.
- [x] Correct the impossible pre-Terminal Done-tree equality rule.
- [x] Correct the authority taxonomy to include bounded repository and
  machine-consumed contracts.
- [x] Permit evidence-backed code/test correction while keeping ambiguous
  behavior behind an explicit decision plan.
- [x] Replace stale baseline counts with a generated current inventory while
  retaining the original `fc4ba4a63` snapshot as history.
- [x] Prove every pre-existing Done path/blob is unchanged and only approved
  closeout additions exist.

### B. Governance tooling and characterization

- [x] Implement `Tools/SpecInformationArchitecture.psm1` as a read-only tracked-file
  inventory and Markdown parser.
- [x] Implement `Tools/Get-SpecInventory.ps1` with JSON/Markdown output below a
  caller-supplied `.build` root and strict failure mode.
- [x] Implement all 21 focused fixtures in
  `Tools/Tests/SpecInformationArchitecture.Tests.ps1`.
- [x] Reuse the inventory from `DocumentationDriftContracts.Tests.ps1` where it
  replaces stale one-off path assumptions.
- [x] Register/document the focused suite in the supported test inventory and
  `Tests/README.md`.
- [x] Run characterization before deleting, moving or compacting any live
  requirement prose.

### C. Normative consistency ledger

- [x] Root/Core: classify and audit `Specs/README.md` plus all 12 Core specs
  against owners, callers, tests, public behavior and sibling claims.
- [x] FileSystem: audit all 7 current normative specs plus the historical FTP
  comparison tombstone, including error/cancellation, identity,
  publication and provider/bridge ownership boundaries.
- [x] Installer: audit all 4 specs against the centralized runtime manifest,
  packagers, extraction smoke and signed/unsigned behavior.
- [x] Plugins: audit all 10 specs against public ABI headers, factories,
  discovery/configuration, unload rules and plugin-specific tests.
- [x] Terminal: audit both current specs against the ABI, plugin/runtime source,
  exact dependency identities, current remediation owner and release gates.
- [x] Testing: audit all 4 specs against actual suite orchestration, case
  enumeration, skip/environment rules, performance policy and CI entry points.
- [x] UI: audit all 19 current normative specs plus the historical Preferences
  migration tombstone against owning controls/windows, commands,
  resources, settings, UIA/DPI/theme/keyboard behavior and focused tests.
- [x] Classify the two Mockup Markdown files as prototypes and the theme notice
  as legal/non-behavioral without applying normative prose rules.
- [x] Record one final status, exact source/test anchors, sibling owners and a
  reviewer note for every candidate file; leave zero `UNVERIFIED` rows.
- [x] Resolve every `SPEC_STALE`, `CODE_STALE`, `CROSS_SPEC_CONTRADICTION` and
  `PROPOSAL_LEAK`, or create one explicit WIP decision file for each genuinely
  unclear situation.

### D. Requirement preservation and cleanup

- [x] Review and evidence every WIP disposition row; preserve unique unfinished
  work with exactly one active/hold/decision owner.
- [x] Respect the explicit `Notes_Scratch.md` human-managed boundary; leave the
  file untouched and exclude it from specification authority, content analysis
  and WIP-plan indexing.
- [x] Reconcile the DxUi and FolderView continuation batons without duplicating
  their master owners.
- [x] Move the live root-plan 012 work to a self-contained WIP owner without
  rewriting historical root plans.
- [x] Add a before/after requirement-ledger row for every removed normative
  section, with code/test owner and surviving destination.
- [x] Remove or route all executable future plans/checklists from normative
  prose, including marker sites not named by the original Atlas examples.
- [x] Compact historical diaries/oversized inventories only after current
  contracts and evidence links survive elsewhere.
- [x] Rebuild `Specs/Plans/WIP/README.md` into exact ACTIVE, HOLD, DECISION and
  RETIRED queues without overwriting unrelated user edits.
- [x] Repair all live normative/WIP links and anchors; report historical broken
  links without editing Done/Reviews/root plans.

### E. Closeout verification

- [x] Strict spec inventory exits 0 with exact WIP index coverage and no
  unresolved normative rows.
- [x] Focused Spec IA Pester suite passes with zero failed tests.
- [x] Documentation drift and all Atlas-owned focused Tools suites pass; the
  complete Tools/Pester aggregate's nine independently owned Terminal failures
  are classified, reflected in current Terminal authority and routed.
- [x] Applicable domain selftests/builds pass for every Atlas code-side
  correction.
- [x] No product hot path or performance-sensitive behavior changed. The
  test-only UIA hang has deterministic regression coverage and its reproduced
  hang/passing evidence is archived under `Specs/TestRuns/`.
- [x] `Tools/Run-AllTests.ps1 -Suite Full` passes for the exact final binaries,
  or any unrelated pre-existing failure is proven and recorded without a false
  green claim.
- [x] `git diff --check` passes and the final diff contains only intended Atlas
  changes plus preserved unrelated user work.
- [x] Move Atlas from WIP to Done without modifying its content, update the WIP
  index, and rerun the protected-history/inventory gates.

## WIP disposition map

This table covers every non-index WIP file present at `fc4ba4a63` and records
the final reviewed disposition. Size waivers are explicit where exact evidence
is load-bearing; line-count reduction is not allowed to erase proof.

| File | Disposition | Remaining owner/content to preserve |
|---|---|---|
| `Core_HResultStatusFormattingCleanup_2026-06-19.md` | KEEP ACTIVE | Seven live formatting/localization work items and validation; already bounded |
| `DxUi_Uia_ContinuationBaton_2026-06-29.md` | RETIRED in Done (2026-08-25) | Do not resume; FolderView closeout is the I2 remainder |
| `FileSystem_GoogleDrivePluginPlan.md` | KEEP ACTIVE | Interactive sign-in, read/write operations and remote validation still open |
| `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md` | KEEP | Background enrichment remains a distinct user-visible feature |
| `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md` | RETIRED in Done (2026-08-25) | Do not resume; WarpDrive remainder is the remaining closeout owner |
| `Notes_Scratch.md` | RETAIN UNTOUCHED / EXCLUDED | Exact human-managed note; no specification or planning authority, no index row, and no content inspection without a direct user instruction naming the file |
| `Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md` | KEEP | Active test-contract architecture owner |
| `Operation_Emberglass_48Hour_RepositoryDeepReview_2026-07-22.md` | KEEP ACTIVE + SIZE WAIVER | Exact interleavings, refutations, inherited-row reconciliation and proof commands remain load-bearing for its live rows |
| `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md` | KEEP ACTIVE | Already reduced to a compact live checklist plus necessary closeout checkpoints |
| `Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` | ARCHIVED 2026-08-25 | Frozen in Done (SIZE WAIVER on the archive). Remaining closeout: `Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md` (I2). |
| `Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md` | KEEP as DECISION/HOLD | Twelve explicit decisions; never rank as executable remediation |
| `Operation_PerfMeasurementContract_2026-07-06.md` | KEEP ACTIVE | Retain enforcement work not already normative in `Testing_PerformanceValidation.md` |
| `Operation_RosettaLantern_RedConfigureSatelliteLocalizationReview_2026-07-17.md` | KEEP ACTIVE | Human linguistic review remains executable work with one owner |
| `Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md` | KEEP HOLD | Live exact review IDs are subordinate evidence for WarpDrive; never execute as a second master |
| `Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md` | ARCHIVED 2026-08-25 | Frozen competitive analysis in Done. Remaining DECISION routing: `Product_WhimFilesRemainingDecisions_2026-08-25.md` (D1). |
| `Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md` | MOVED TO DONE | Implemented core/runtime/model contract merged into the Terminal domain spec |
| `Terminal_CoreRendererInputAccessibility_2026-08-01.md` | MOVED TO DONE | Implemented renderer/input/accessibility contract merged into the Terminal domain spec |
| `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md` | MOVED TO DONE | Completed master plan; current authority is the Terminal domain spec and engine decision |
| `Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md` | MOVED TO DONE / SUPERSEDED | Preserved as decision history; the qualified private Ghostty route won |
| `Terminal_FirstImplementationExecutionIndex_2026-08-01.md` | MOVED TO DONE | All v1 child outcomes and closeout gates completed |
| `Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md` | MOVED TO DONE | Implemented v1 interaction/lifecycle contract merged into the Terminal domain spec |
| `Terminal_PluginRuntimeBoundary_2026-08-01.md` | MOVED TO DONE | Implemented lean ABI/private runtime boundary is normative in the Terminal domain spec |
| `Terminal_QualifyPrivateLibGhosttyWinner_2026-08-01.md` | MOVED TO DONE | Exact pin, hashes, decision, and dependency lifecycle remain normative |
| `Terminal_ReleaseUpgradeCloseout_2026-08-01.md` | MOVED TO DONE | Security, performance, package, notices, and upgrade/rollback closeout completed |
| `Terminal_RendererControlAndAccessibility_2026-07-22.md` | MOVED TO DONE | Implemented v1 renderer/input/UIA requirements merged into the Terminal domain spec |
| `Terminal_SettingsSecurityPerfPackagingCloseout_2026-07-22.md` | MOVED TO DONE | Implemented settings/security/perf/package requirements merged into current authority |
| `Terminal_ShellIntegrationHistoryAndPaneFollow_2026-07-22.md` | MOVED TO DONE | Implemented trusted PowerShell/reuse/follow contract merged into current authority |
| `Terminal_UserInteractionsAndSettings_2026-08-01.md` | MOVED TO DONE | Implemented pane commands, insertion, profiles, follow, and Preferences contract merged |
| `Win32_Inventory.md` | KEEP ACTIVE | Four explicitly live RAII/lifetime/DPI audit items; already bounded |

### Expected WIP shape

The executed queue immediately before Atlas closeout is exactly 14 ACTIVE,
2 HOLD, 2 DECISION and 2 RETIRED files. The exact count is not a deletion quota.
The following invariants are hard:

- every ACTIVE/HOLD/DECISION row has exactly one corresponding file;
- every direct planning file is indexed as ACTIVE, HOLD, DECISION or RETIRED;
- no active terminal entry mentions a no-winner or blocked product gate;
- `Notes_Scratch.md` is present but classified `HumanManagedNote` and excluded
  from planning/specification analysis;
- tombstones are at most 40 lines and 4 KiB each;
- any active file above 1,500 lines or 128 KiB has an explicit index waiver
  explaining why splitting would create duplicate ownership.

## Normative disposition map

### Root and Core

| File | Disposition |
|---|---|
| `Specs/README.md` | REWRITE as the authority/classification catalog, pinned-path list and reading map |
| `Core/Core_CompareDirectories.md` | RETAIN; remove closed rollout evidence and keep behavior/metrics |
| `Core/Core_CompilerWarnings.md` | RETAIN; link the compiler-warning skill and keep binding policy |
| `Core/Core_ConnectionManager.md` | RETAIN; remove migration chronology already represented by current implementation status |
| `Core/Core_DirectoryInfoCache.md` | RETAIN |
| `Core/Core_FileSystemBridge.md` | RETAIN; deduplicate provider/file-operation contracts by links |
| `Core/Core_Localization.md` | RETAIN |
| `Core/Core_RedConfigure.md` | RETAIN |
| `Core/Core_RedSalamanderMonitor.md` | REWRITE current contract; remove fixed-bug diary, old measured-actual tables and improvement-plan prose; route genuine open work to WIP |
| `Core/Core_Search.md` | RETAIN |
| `Core/Core_SettingsStore.md` | RETAIN as schema behavior owner; remove examples/history duplicated by `SettingsStore.schema.json` or user docs |
| `Core/Core_SharedHelpers.md` | RETAIN |
| `Core/Core_StartupBootstrap.md` | RETAIN |

### FileSystem and Installer

| File | Disposition |
|---|---|
| `FileSystem/FileSystem_FileOperations.md` | RETAIN core contract; remove `Future Enhancements` and informative implementation-plan/status appendices after routing unique open work |
| `FileSystem/FileSystem_FtpScpSftpPerformanceComparison.md` | COMPACT TO TOMBSTONE unless an explicit backend decision is still open; merge current libcurl constraints into `FileSystem_FtpSftpScp.md` |
| `FileSystem/FileSystem_FtpSftpScp.md` | RETAIN as current protocol/backend contract |
| `FileSystem/FileSystem_GoogleDrive.md` | RETAIN current implemented/unsupported matrix |
| `FileSystem/FileSystem_Imap.md` | RETAIN |
| `FileSystem/FileSystem_MicrosoftDrive.md` | RETAIN |
| `FileSystem/FileSystem_Mtp.md` | RETAIN |
| `FileSystem/FileSystem_S3.md` | RETAIN |
| `Installer/Installer_Msi.md` | RETAIN |
| `Installer/Installer_Msix.md` | RETAIN |
| `Installer/Installer_PortableZip.md` | RETAIN |
| `Installer/WingetIntegration.md` | RETAIN path for compatibility; make heading/classification normative and remove tutorial duplication |

### Plugins

| File | Disposition |
|---|---|
| `Plugins/Plugins_PluginAPI.md` | SPLIT current ABI from `New Host UI Services (Planned ABI)`; route unimplemented ABI to a WIP plan/RFC |
| `Plugins/Plugins_ViewerImgRaw.md` | RETAIN |
| `Plugins/Plugins_ViewerPE.md` | RETAIN |
| `Plugins/Plugins_ViewerPlugins.md` | RETAIN shared viewer contract; link plugin-specific details instead of copying them |
| `Plugins/Plugins_ViewerSpace.md` | RETAIN |
| `Plugins/Plugins_ViewerSqlite.md` | RETAIN |
| `Plugins/Plugins_ViewerText.md` | RETAIN |
| `Plugins/Plugins_ViewerVLC.md` | RETAIN |
| `Plugins/Plugins_ViewerWeb.md` | RETAIN |
| `Plugins/Plugins_VirtualFileSystem.md` | RETAIN as ABI owner; remove `Future Extensions` and plan/status prose, and link provider-specific policy |

### Terminal and Testing

| File | Disposition |
|---|---|
| `Terminal/Terminal_EmbeddedPlugin.md` | RETAIN as current product/runtime boundary contract |
| `Terminal/TerminalEngineDecision.md` | RETAIN current executive decision, pin, qualification and lifecycle. Preserve byte-protected Gate-0 history unless its checksum test and every immutable historical reference can remain valid without touching Done; this is an explicit size exception |
| `Testing/Testing_PerformanceValidation.md` | RETAIN and absorb only still-binding rules from the perf WIP |
| `Testing/Testing_SelfTestRemoteCredentials.md` | RETAIN |
| `Testing/Testing_SelfTests.md` | RETAIN as harness contract |
| `Testing/Testing_TestCoverage.md` | REBUILD as coverage policy/risk map plus generated-inventory instructions; move exhaustive current case listing to `.build/SpecInventory` output and update consumers |

### UI

| File | Disposition |
|---|---|
| `UI/UI_BatchRenameWindow.md` | RETAIN |
| `UI/UI_CommandMenuKeyboard.md` | RETAIN routing/command semantics; remove `Implementation Plan (Proposed)` and generate mechanical command inventory from code |
| `UI/UI_DxUiSharedGrid.md` | RETAIN current shared-control contract; remove migration/deferred-work chronology after routing genuine open work |
| `UI/UI_DxUiWinUIDesign.md` | RETAIN current design/control tokens and acceptance rules; remove superseded migration prose |
| `UI/UI_FileOperationsPopup.md` | RETAIN |
| `UI/UI_FindFilesWindow.md` | RETAIN |
| `UI/UI_FolderView.md` | RETAIN |
| `UI/UI_FolderWindow.md` | RETAIN |
| `UI/UI_KeyboardManagement.md` | KEEP as the existing minimal redirect; do not expand it |
| `UI/UI_ManagePluginsDialog.md` | RETAIN |
| `UI/UI_NavigationView.md` | RETAIN current contract; remove `Recent Architectural Improvements` chronology and `Future Enhancements` after routing live items |
| `UI/UI_PreferencesDialog_MigrationHistory.md` | MIGRATE live follow-ups to WIP/current spec, then compact to a tombstone/history pointer |
| `UI/UI_PreferencesDialog.md` | RETAIN as current contract |
| `UI/UI_RedConfigure.md` | RETAIN |
| `UI/UI_ThemeCycleOverlay.md` | RETAIN |
| `UI/UI_TopLevelToolWindows.md` | RETAIN |
| `UI/UI_VisibleComctlAudit.md` | COMPACT to current scope/conclusion and current tool command; remove closeout chronology |
| `UI/UI_VisibleNativeAudit.md` | COMPACT to current allowlist/conclusion and current tool command |
| `UI/UI_VisibleTypographyAudit.md` | Route remaining drift to a WIP owner, then compact to current allowlist/conclusion or tombstone |
| `UI/UI_VisualStyle.md` | RETAIN |

### Legal, prototype and data artifacts

| Surface | Decision |
|---|---|
| `Themes/THIRD-PARTY-NOTICES.md` | Legal evidence; unchanged |
| `Specs/Mockups/**` including local `AGENTS.md` and `design-qa.md` | Prototype estate; unchanged and excluded from normative checks |
| `SettingsStore.schema.json` | Pinned schema; unchanged except separately approved schema work |
| `Testing/FolderViewPerfBudgets.json5` | Live test budget; unchanged |
| `Terminal/*.json` and `Terminal/TerminalEngineEvidenceCorpus/**` | Retain exact bytes as historical Gate-0 schemas/fixtures until a separate tool-retirement plan proves no current consumer; classify them as historical data in `Specs/README.md` |
| `Specs/TestRuns/**` | Pinned evidence; unchanged |

## Executed disposition and preservation ledger

### Normative audit results

The tracked ledger is `Specs/NormativeConsistency.json`; the generated review
copy is `.build/SpecInventory/normative-consistency.json`. The review covered
every current normative section, its implementation/public contract, focused
tests and sibling owners. Historical/prototype/legal/evidence files were
classified separately rather than given false product authority.

| Domain | Current normative files | Final status | Resolved during Atlas |
|---|---:|---|---|
| Root/Core | 13 | 13 `VERIFIED_MATCH` | Rebuilt the authority catalog; corrected Monitor proposals, SettingsStore debounce language, the SharedHelpers Terminal queue description and one Compare anchor. |
| FileSystem | 7 | 7 `VERIFIED_MATCH` | Removed File Operations plan/history leakage; retired the separate FTP performance paper as history. |
| Installer | 4 | 4 `VERIFIED_MATCH` | Corrected Winget artifact examples and added MSI/MSIX/ZIP manifest-consumption contracts. |
| Plugins | 10 | 10 `VERIFIED_MATCH` | Removed planned host-service/VFS extensions from normative ABI prose, routed genuine choices, and removed a ViewerVLC claim against an unavailable evidence archive. |
| Terminal | 2 | 2 `VERIFIED_MATCH` | Verified the current executive/product contracts separately from the byte-protected Round-4 history. Current authority now distinguishes implemented product scope from **reopened** release verification, records the nine current frozen-governance failures, and routes the independent `J21-TERM-07` remediation owner without changing historical run IDs. |
| Testing | 4 | 4 `VERIFIED_MATCH` | Rebuilt Test Coverage as policy/risk ownership and generated the exhaustive test inventory; corrected bounded UIA worker handling, cross-worktree interactive-test serialization and generated dependency scan boundaries. |
| UI | 19 | 19 `VERIFIED_MATCH` | Corrected stale/proposed command, DxUi, FolderView, FolderWindow, Navigation, Preferences, typography, visual-style and Batch Rename prose; removed missing evidence claims and repaired the visible-native audit boundary/allowlist. |
| **Total** | **59** | **59 `VERIFIED_MATCH`; 0 unresolved** | **Atlas fixed conclusive code-side mismatches in the Win32 audit, localization audit, test orchestration and Commands UIA selftest harness. No product behavior changed. All other conclusive mismatches were documentation; uncertain product choices were routed instead of guessed.** |

The two additional domain Markdown files are historical:
`UI_PreferencesDialog_MigrationHistory.md` and
`FileSystem_FtpScpSftpPerformanceComparison.md`. The inventory also classifies
two Mockup Markdown files as prototypes, the theme notice as legal, TestRuns as
evidence, Gate-0 JSON/corpus files as historical data, and the settings/perf
budget files as machine contracts. Those classifications explain the corrected
FileSystem and UI counts above.

### Removed or rewritten normative material

Each row records where the load-bearing behavior went before the old prose was
removed. “Decision queue” means the repository could not prove that an option
was intended current behavior; it is not implementation approval.

| Source material removed or rebuilt | Implementation/test owner checked | Surviving current destination | Unfinished destination |
|---|---|---|---|
| Monitor fixed-bug chronology, old measurements and optimization/feature proposals | `RedSalamanderMonitor/*`; `Tests/MonitorTest` | `Core_RedSalamanderMonitor.md` current capture/render/search/export contract | `ATLAS-DEC-MON-01..03` |
| SettingsStore continuous/debounced placement-save proposal | `Common/Common/SettingsStore.cpp`; SettingsSchema/Commands tests | `Core_SettingsStore.md` close/shutdown persistence contract | `ATLAS-DEC-WIN-01` |
| SharedHelpers pre-Gate-0 ConPty queue description | `Plugins/Terminal/Terminal.cpp`; Terminal source contracts | `Core_SharedHelpers.md` current bounded Terminal input queue | none |
| File Operations future section and implementation/status appendices | `FolderWindow.FileOperations*.cpp`; phased FileOps selftests | `FileSystem_FileOperations.md`, `Core_FileSystemBridge.md`, `UI_FileOperationsPopup.md` | `FileOperations_RemainingStressValidation_2026-08-04.md` |
| FTP/SCP/SFTP comparative research narrative | FileSystemCurl source/tests | `FileSystem_FtpSftpScp.md`; compact historical pointer at the original path | existing Floodgate proof rows only |
| PluginAPI planned host UI services | `Common/PlugInterfaces/Host.h`; host/lifetime tests | `Plugins_PluginAPI.md` implemented services and lifecycle | `ATLAS-DEC-HOST-01` |
| VFS future paging/polling extensions and rollout prose | `Common/PlugInterfaces/FileSystem.h`; PluginContract/lifetime tests | `Plugins_VirtualFileSystem.md` complete-result ABI and notification contract | `ATLAS-DEC-VFS-01..02` |
| Hand-maintained exhaustive test-case diary | test binaries, `TestInventory.ps1`, `Get-TestInventory.ps1` and inventory tests | `Testing_TestCoverage.md` policy/risk map; generated `.build/SpecInventory/test-inventory.json` | none |
| Command/Menu proposed target plan and inaccurate parameterized-route wording | `CommandRegistry.cpp`, `ShortcutDefaults.cpp`; Commands tests | `UI_CommandMenuKeyboard.md` current IDs/routes/defaults | `ATLAS-DEC-CMD-01` |
| DxUi Grid stale Preferences status and editable-cell/future notes | `Common/DxUi/DxUi.Grid.cpp`; Grid/UIA tests | `UI_DxUiSharedGrid.md` current readonly/toggle/range contract | `ATLAS-DEC-DXUI-01` |
| DxUi proposed palette numbers, completed migration prose and broader BiDi/editing/composition proposals | DxUi theme/input/window-host source and tests | semantic roles/current mechanics in `UI_DxUiWinUIDesign.md` | `ATLAS-DEC-DXUI-01` |
| FolderView GDI fallback contradiction, future modes and filename-tooltip proposal | `FolderView*.cpp`; Commands/performance tests | `UI_FolderView.md` current DirectWrite/DxUi behavior | `ATLAS-DEC-FV-01..02` |
| FolderWindow rollout/history and future-plan narrative | `FolderWindow*.cpp`; Navigation/View/FileOps tests | `UI_FolderWindow.md` current dual-pane contract | none |
| Navigation architectural chronology and optional roadmap | `NavigationView*.cpp`; Navigation tests | `UI_NavigationView.md` current breadcrumb/edit/history/provider behavior | `ATLAS-DEC-NAV-01` |
| Preferences rollout/migration diary and unimplemented search/schema/mouse proposals | `Preferences*.cpp`, SettingsStore/schema; Preferences tests | `UI_PreferencesDialog.md`; compact historical pointer at the migration path | `ATLAS-DEC-PREF-01..03` |
| Batch Rename secondary File Operations task proposal | Batch Rename window/engines and tests | `UI_BatchRenameWindow.md` single-window task owner | `ATLAS-DEC-BR-01` |
| Visible Typography closeout chronology | DxUi typography/native input and audits/tests | compact `UI_VisibleTypographyAudit.md` current allowlist | none |
| Unavailable ViewerVLC, Terminal, FolderView, Theme Cycle and visible-native/comctl evidence-path claims | Current focused test cases, live audit commands, and repository TestRun existence | executable test/metric/audit contracts in their owning specs; no missing path presented as proof | fresh evidence is required for any new performance or closeout claim |
| Broad Win32 audit dependency-tree false positives and unclassified current residuals | live `-FailOnFindings` run plus affected source/test seams | explicit `vcpkg_installed` exclusion and current backdrop/test-only allowlist in tool and `UI_VisibleNativeAudit.md` | none |
| Item Properties broad-order Commands hangs and unsafe dialog UIA timeout detach | three broad-order reproductions (live interaction twice and stream Remove once, including two clean Full contexts), exact focused/prefix reruns, `Commands.SelfTest.Dialogs.cpp` and harness source contracts | all Item Properties owner-thread UIA reads/invokes now use the shared bounded `IUIAutomation2` timeout, COM cancellation and join path; `Testing_SelfTests.md` owns the rule | none |
| Concurrent worktree Full runs corrupted foreground, pointer, clipboard and timing evidence | runner processes/traces from two repository roots plus desktop-global test contracts | `Run-AllTests.ps1` now holds one session-wide interactive-test mutex through archival/cleanup; source contract and `Testing_SelfTests.md` updated | none |
| Localization RC audit scanned generated top-level `vcpkg_installed` content | failing Full-run Pester output and focused localization rerun | generated dependency tree excluded from the source-root allowlist audit | none |
| Targeted deployment proof hit parallel `vc145.pdb` update collisions | captured `pester-targeted-plugin-deployment` stdout and MSBuild log | focused deployment build now uses one MSBuild worker, matching the successful recovery-build discipline | none |
| Visual Style marquee future item | FolderView interaction/rendering and theme tests | current `UI_VisualStyle.md` | `ATLAS-DEC-FV-01` |

### Human-managed note correction

The earlier Atlas triage and deletion of `Notes_Scratch.md` violated the file's
explicit human-managed boundary. That analysis is withdrawn. The file is
restored and retained byte-for-byte; Atlas does not use it as documentation,
evidence, a work queue or a source of requirements. Every current decision and
work owner recorded elsewhere in Atlas was independently checked against code,
tests, current specifications or separately authorized review material.

### WIP disposition results

The original WIP map was advisory. Review produced these evidence-backed
corrections:

- Emberglass and WarpDrive remain single active masters with explicit size
  waivers in the index. Their exact interleavings, refutations and archived
  measurement/failure chronology are load-bearing; destructive compaction or a
  split would create duplicate ownership.
- Floodgate is already a compact 120-line live checklist; Core HRESULT,
  Astrolabe, Rosetta, Win32, Google Drive, thumbnail enrichment and the
  performance measurement contract remain bounded active owners.
- Tailwind remains HOLD rather than a tombstone because its still-open exact
  review IDs are subordinate evidence for WarpDrive; it cannot be executed as a
  second master. Parallax remains HOLD. WhimFiles is DECISION only.
- The two continuation batons were later moved to Done (2026-08-25) and must not
  be resumed. Historical root plan 012 is `AdvisorPlan_012-format-autocommit-race_2026-06-11.md`;
  live CI format work is `CI_FormatAutocommitRace_2026-08-25.md`.
- The new File Operations validation plan and Atlas decision queue hold the
  obligations independently established from current code, tests, normative
  contracts and authorized review material. The
  unrelated July-21-to-August-4 review remains untouched as its own active
  owner.

The final direct-WIP shape before moving Atlas is 14 ACTIVE, 2 HOLD, 2 DECISION
and 2 RETIRED files. After Atlas moves to Done it is 13 ACTIVE, 2 HOLD,
2 DECISION and 2 RETIRED; all remaining files are indexed exactly once.

## Files the executor may create or modify

### Governance tooling

- create `Tools/SpecInformationArchitecture.psm1`;
- create `Tools/Get-SpecInventory.ps1`;
- create `Tools/Tests/SpecInformationArchitecture.Tests.ps1`;
- modify `Tools/Tests/DocumentationDriftContracts.Tests.ps1` only to remove
  obsolete exact-path assumptions and reuse the new inventory;
- update `Tests/README.md` and `Tools/TestInventory.ps1` if the repository's
  tooling test inventory requires explicit registration.

### Documentation

- `Specs/README.md`;
- `Specs/Plans/WIP/README.md`;
- the WIP files and normative Markdown files named in the two disposition maps;
- user/developer documentation only when a renamed or compacted spec path is
  linked there.

### Evidence-backed implementation corrections

- production code, headers, schemas/resources, build/package manifests and
  deterministic tests may be modified only when the completed consistency row
  proves they are the stale side of an unambiguous current contract;
- each such correction must name its applicable repository skill, regression
  proof and focused verification in the ledger;
- performance-sensitive changes must satisfy the repository performance
  validation contract from the beginning, including archived evidence;
- do not bundle unrelated remediation merely because Atlas discovered it.

### Explicitly out of scope

- `Specs/Plans/Done/**`;
- `Specs/Reviews/**`;
- root `plans/**`;
- `Specs/TestRuns/**`;
- `Specs/Mockups/**`;
- theme JSON5 and legal notices;
- settings schema and performance budgets unless a consistency row proves that
  machine-consumed contract is itself stale and the intended replacement is
  unambiguous;
- engine/runtime upgrades;
- rewriting historical Done links.

If current behavior contradicts a normative requirement, determine the stale
side from implementation, tests, public compatibility and current ownership.
Correct it only when that evidence is conclusive. Otherwise stop that row and
open a separate behavior/spec decision plan; Atlas may never silently choose.

## Inventory tool contract

`Get-SpecInventory.ps1` is read-only and writes only beneath a caller-supplied
`.build` output root. It must use tracked files from `git ls-files` rather than
recursing ignored TestRuns. Its JSON result contains:

- repository/base commit and timestamp;
- relative path, bytes, line count and first H1;
- authority class and domain;
- normative consistency status, evidence anchors, reviewed sections, reviewer
  note and unresolved discrepancy IDs;
- claimed contract owner and any sibling document that restates the contract;
- parsed status for WIP;
- inbound and outbound local links;
- missing target and missing heading-anchor findings;
- WIP index membership;
- checkbox, TODO, Future Enhancements, Proposed Plan and history-marker counts;
- tombstone size/shape;
- explicit exceptions with rationale.

Markdown parsing must:

1. ignore fenced code blocks opened by three or more backticks or tildes;
2. ignore inline code spans with matching delimiter length;
3. ignore images and external `http`/`https`/`mailto` links;
4. URL-decode and resolve local paths relative to the source document;
5. reject paths escaping the repository root;
6. validate file targets for normative and WIP sources;
7. validate heading anchors using GitHub-style slugging, including duplicate
   heading suffixes;
8. report, but not fail on, links originating inside immutable Done/history.

Do not implement the scanner as one repository-wide regex. The reconnaissance
proved that code samples contain valid `](` sequences.

## Execution waves

### Wave 0 — Prove normative consistency and freeze the baseline

1. Confirm clean status and `fc4ba4a63` ancestry.
2. Record tree identities for all protected surfaces.
3. Generate an inventory of every normative/WIP file and local link.
4. Classify the 64 baseline non-plan/non-review/non-TestRun Markdown files;
   exclude the two Mockup prototypes and legal/evidence data from behavioral
   authority.
5. Audit every normative file using the domain matrix. Record:
   - consistency status;
   - exact implementation, ABI/schema/resource and test anchors;
   - sibling normative claims;
   - current-vs-proposed behavior;
   - discrepancies and their resolution owner.
6. Resolve every `SPEC_STALE`, `CROSS_SPEC_CONTRADICTION` and `PROPOSAL_LEAK`
   documentation row. Resolve `CODE_STALE` with a bounded, tested code-side
   correction when intended behavior is conclusive; otherwise route the exact
   uncertainty to a separate decision plan and stop only that row.
7. Append the human `Normative audit results` ledger to this plan and generate
   `.build/SpecInventory/normative-consistency.json`.
8. For each section scheduled for deletion, record:
   - current requirement IDs or imperative statements;
   - code owner and test owner;
   - surviving normative destination;
   - surviving active-plan owner, if unfinished;
   - inbound immutable-history links.
9. No content is removed in this wave.

**Verify:**

```powershell
git status --short
git rev-parse --is-ancestor fc4ba4a63 HEAD
git rev-parse HEAD:Specs/Plans/Done
```

Expected: unrelated dirty state is recorded rather than overwritten, ancestor
exit 0, Done tree is the approved post-Terminal
`cfd921572d9e8b18c4bf0ec5eb044bbb8de280a5` before the final Atlas move, every
normative file is classified, and zero `UNVERIFIED` rows before Wave 1.

### Wave 1 — Land governance and characterization before cleanup

Implement the inventory module, CLI and Pester tests first. Characterize the
current exceptions explicitly; do not make the first test pass by deleting
documents. Add a compact classification table to `Specs/README.md` and a
machine-readable inventory table to the WIP index.

Required tests:

- every direct WIP Markdown file is indexed exactly once;
- active/hold/decision/tombstone status values are closed;
- Done tree matches the frozen identity;
- pinned/out-of-scope trees are not modified;
- normative and WIP local links resolve;
- code samples are ignored by the link scanner;
- a temporary broken link is rejected;
- a temporary missing/duplicate WIP row is rejected;
- a normative active checkbox/future-plan marker is rejected unless allowlisted;
- tombstone size and required fields are enforced.

**Verify:** focused Pester passes before any bulk edit.

### Wave 2 — Rebuild the WIP index and remove non-terminal ambiguity

1. Replace narrative historical refreshes in `WIP/README.md` with:
   - authority rules;
   - current ranked active queue;
   - HOLD/DECISION queue;
   - retired tombstone table;
   - dependency graph;
   - last inventory commit.
2. Migrate the remaining root `plans/012-format-autocommit-race.md` work into a
   new self-contained WIP owner without modifying root `plans/`; update the
   drift contract accordingly.
3. Treat `Notes_Scratch.md` as an exact human-managed exclusion: do not inspect,
   edit, delete, summarize, index or route its contents without a direct user
   instruction naming that file.
4. Reconcile the DxUi and FolderView batons.
5. Compact Emberglass, Floodgate, WarpDrive, Perf and the partial Google Drive
   plan to open work only.
6. Keep Parallax, Rosetta and WhimFiles out of the executable queue unless
   promoted by their explicit decision criteria.

**Verify:** WIP index coverage is exact and no live owner exists only in a
historical/root plan.

### Wave 3 — Consolidate terminal planning — COMPLETE 2026-08-03

Terminal no longer has an active WIP router. The durable inputs are:

- `Specs/Terminal/Terminal_EmbeddedPlugin.md` for the complete v1 host/plugin,
  behavior, settings, security, performance, package, and release contract;
- the current executive decision at the beginning of
  `Specs/Terminal/TerminalEngineDecision.md` for the selected Ghostty engine,
  private runtime policy, exact pin/hashes, qualification, and upgrade boundary;
- `Common/PlugInterfaces/Terminal.h` for the lean unsuffixed public ABI.

Every former Terminal WIP requirement was classified as implemented v1
behavior, an explicit v1 non-goal, or superseded decision history. Implemented
behavior was merged into the Terminal domain spec. All 13 plans were moved to
`Specs/Plans/Done/`; the WezTerm Round-5 plan is explicitly superseded. The
exact Ghostty pin, x64/ARM64 identities, private DLL boundary, unsuffixed ABI,
offline reproduction, notices, upgrade/rollback procedure, and release gates
remain in current authority.

For the exact 2026-08-03 bytes, final verification passed 96/96 post-move
documentation, Terminal source, dependency lifecycle, package/reproducibility,
archive, and runner contracts. Those results remain historical evidence only.
The current edited Terminal/governance bytes reopened release verification on
2026-08-04 after nine transition/fixture/ABI/closeout Pester failures; the
current status and `J21-TERM-07` owner are recorded in
`Specs/Terminal/Terminal_EmbeddedPlugin.md` and the executive decision.

Native ARM64 whole-product compilation and high-risk ASan remain general
release-matrix lanes; the ARM64 Ghostty runtime itself is qualified and
offline-verifiable, so those environment lanes do not keep Terminal v1 in WIP.

### Wave 4 — Remove obvious planning/history leakage from normative specs

Process low-risk paths first:

1. keep `UI_KeyboardManagement.md` as its small redirect;
2. compact Preferences migration history after routing live work;
3. remove File Operations future/plan appendices after migration;
4. remove Command/Menu proposed implementation plan;
5. remove NavigationView improvement chronology/future section;
6. compact Monitor fixed-bug and old performance diary;
7. compact visible UI audits to current allowlists and commands;
8. compact the FTP/SCP/SFTP comparison to a tombstone/current decision.

Each edit must have a requirement-ledger before/after row and focused domain
tests where wording is asserted.

### Wave 5 — Rebuild oversized inventories without losing contracts

#### Test coverage

- retain suite purpose, critical coverage obligations, opt-in/environment
  policy and ownership;
- derive exhaustive case names from `Tools/TestInventory.ps1` and
  `--selftest-list-cases` into `.build/SpecInventory`;
- update `Testing_SelfTests.md`, `Tests/README.md` and drift tests to point to
  the generated inventory;
- keep only cases whose names are themselves a load-bearing contract in the
  normative file.

#### Commands

- retain routing, stable command IDs, menu behavior, shortcut policy and user
  semantics;
- derive mechanical command registration/default-shortcut inventory from
  `CommandRegistry.cpp` and `ShortcutDefaults.cpp`;
- remove target/proposed implementation prose.

#### Large coherent domain specs

For SettingsStore, File Operations, VirtualFileSystem, DxUi, FolderView and
NavigationView, remove duplicate history first. Split only when two sections
have different code/test owners and can link without duplicating invariants.
Size alone is not authority.

### Wave 6 — Classify retained Terminal Gate-0 artifacts

The 33 JSON schema/fixture/evidence files remain byte-identical in Atlas because
current tools/tests and immutable history reference them. `Specs/README.md` must
label them historical Gate-0 data rather than current product configuration.

Do not remove old evaluator tooling in this operation. If the inventory proves
it has no current release/upgrade consumer, create a separate tool-retirement
plan. That plan must rewrite `TerminalDependencyLifecycle.Tests.ps1` around
`Build-GhosttyTerminalRuntime.ps1` before removing anything.

### Wave 7 — Repair live links and documentation contracts

- update only normative/WIP/docs/tools/tests links;
- do not edit broken or old links inside Done;
- replace `DocumentationDriftContracts` assumptions about live root plans and
  stale WIP paths with the new classification/index checks;
- update `Tests/README.md` for the new Pester suite;
- verify redirect/tombstone paths continue to satisfy immutable history.

### Wave 8 — Final verification and commit discipline

Use separate reviewable commits for:

1. governance tool/tests;
2. WIP index and non-terminal triage;
3. any evidence-backed code-side corrections, grouped by domain;
4. normative low-risk cleanup;
5. generated inventory/large-document cleanup;
6. final links and documentation closeout plus the Atlas move.

Terminal consolidation is already historical at `6aecfde8e`; do not manufacture
a new Atlas commit for it.

Run focused gates after every commit and Full once at the end. No push is
authorized by this plan.

## Commands and expected results

| Purpose | Command | Expected result |
|---|---|---|
| Baseline | `git status --short --branch` | intended branch and no unrelated changes |
| Protected Done before Atlas move | `git rev-parse HEAD:Specs/Plans/Done` | exactly `cfd921572d9e8b18c4bf0ec5eb044bbb8de280a5` |
| Protected Done after Atlas move | inventory blob/path comparison against `6aecfde8e` | all baseline paths byte-identical and Atlas is the only new path |
| Inventory | `pwsh -File Tools/Get-SpecInventory.ps1 -Root Specs -OutputRoot .build/SpecInventory -PassThru` | exit 0; JSON and Markdown report below `.build` |
| Normative consistency | inspect `.build/SpecInventory/normative-consistency.json` and this plan's `Normative audit results` | every normative file has evidence; zero unresolved/non-final statuses |
| Strict IA | same command with `-FailOnFindings` | exit 0 after migration |
| Focused Pester | `Invoke-Pester -Script .\Tools\Tests\SpecInformationArchitecture.Tests.ps1 -PassThru` | 0 failed |
| Existing doc contracts | `Invoke-Pester -Script .\Tools\Tests\DocumentationDriftContracts.Tests.ps1 -PassThru` | 0 failed |
| Tool suite | repository-supported complete Tools/Pester command | Atlas-owned suites pass; every unrelated failure is classified and routed without changing its result |
| Full suite | `.\Tools\Run-AllTests.ps1 -Suite Full` | exit 0, or proven unrelated/environmental failures are recorded without a false green claim |
| Diff hygiene | `git diff --check` and `git diff --cached --check` | exit 0 |

For Pester 3, inspect the returned object's `FailedCount`; do not use unsupported
Pester 5-only output parameters.

## Test plan

Create `Tools/Tests/SpecInformationArchitecture.Tests.ps1` with temporary
fixtures obtained through `New-RSTestSandboxScratchDirectory(...)` below the
shared TestSandbox wall. This corrects the older fixed `.build` fixture proposal
to match `Tests/README.md`. Cover:

1. valid normative document;
2. valid active WIP and matching index row;
3. valid tombstone;
4. duplicate index entry;
5. unindexed WIP file;
6. stale index path;
7. normative unchecked checkbox;
8. normative `Future Enhancements`/`Implementation Plan` marker;
9. local relative link and heading anchor;
10. duplicate heading slug;
11. missing target and missing anchor;
12. fenced C++, PowerShell and Markdown containing `](`;
13. inline-code fake link;
14. repository-root escape attempt;
15. immutable Done source with a stale link reported informationally;
16. modified Done tree rejected;
17. tombstone above size/line cap;
18. protected schema/theme/TestRuns/Mockups paths excluded from prose rules.
19. every normative file has a consistency record and evidence anchors;
20. `UNVERIFIED` or any unresolved contradiction fails strict mode;
21. WIP authority claims and unlabelled proposed behavior are rejected.

Tests must not mutate tracked files or depend on ignored real TestRuns.

## Closeout verification record — 2026-08-04

Atlas closes on evidence, not on a synthetic green aggregate:

- strict inventory passed with 318 Markdown candidates, 59 normative files,
  exact indexed planning coverage, one separately classified human-managed
  note, 59/59
  `VERIFIED_MATCH`, zero unresolved discrepancies and zero blocking findings;
  after the move and index update it passed again with exact 19/19 direct-WIP
  coverage and the protected Done baseline intact;
- focused Pester passed: Spec IA 21/21, Documentation Drift 17/17,
  Test Harness Source Contracts 151/151, Build Reproducibility 13/13 and
  Resource Localization Contracts 5/5; the strict remaining-Win32 audit also
  passed with every residual classified and zero unallowed findings;
- the uncontaminated recovery rebuild passed with zero warnings/errors
  (`msbuild-20260804_163157_977.log`), and the final Full rebuilt the exact
  binaries with zero warnings/errors (`msbuild-20260804_165403_176.log`);
- all ten Item Properties cases passed in broad order after both owner-thread
  UIA call sites were moved to the bounded cancellable worker. The exact Stream
  Remove regression passed 1/1 in run
  `20260804T145335Z-50728-5cf842589c3547cfb217bbf952788edb`;
- final Full run
  `20260804T145400Z-44500-1c7f666e114f4198805c1acc02733ab7`
  completed rather than hanging. Its aggregate reported 1,149 passed, 19
  failed and 64 skipped entries. Commands reported 798 passed, 17 failed and
  13 skipped; Tools/Pester reported 621 passed, 9 failed and 5 skipped; one
  ViewerPE suite also failed its clipboard-warning contract. Compare, File
  Operations, DxUi, FileSystemCurl, ViewerSqlite, Monitor, Localization,
  RedConfigure, Plugin Contract, Settings Schema, Crash Handling, Monitor ETW,
  PerformanceTests2, deployment proof and the synthetic vcpkg contract passed;
- the desktop-interactive failures are not an Atlas regression claim:
  independent `OpenClipboard(NULL)` failed with access denied, foreground
  ownership was `LockApp.exe`, representative clipboard/menu/path-edit reruns
  failed under that locked desktop, and the isolated unfocused-pane case passed
  1/1 in run
  `20260804T160957Z-14360-76b1721895364123888c3bbae34cfb50`;
- the nine Tools failures are in the independently modified Terminal candidate
  transition (5), fixture generator (1), Gate-0 harness (1) and Round-4
  closeout (2) suites. Current Terminal authority now records verification as
  reopened and `J21-TERM-07` owns repair plus a new authoritative aggregate.
  Atlas neither rewrote those user-owned implementation changes nor relabeled
  their failures;
- the Full run directory was later removed by the runner's normal stale-run
  cleanup during focused reruns. The reproduced UIA hang and passing focused
  proof remain archived at
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-08-04_atlas_blocking_commands_item_properties_hang/`.

This satisfies Atlas's explicit unrelated/environmental-failure alternative;
it is not a claim that the whole repository or Terminal release gate is green.

## Done criteria

All must be true:

- [x] All 120 pre-Terminal `Specs/Plans/Done` files are unchanged; the only
      current additions are the 13 Terminal plans moved at completed v1
      closeout.
- [x] At final closeout Atlas is the only further Done addition and every
      previously protected blob remains byte-identical.
- [x] Every baseline normative/WIP Markdown file has a disposition and every
      removed section has a requirement-ledger destination.
- [x] Every normative file has been compared with implementation, tests and
      sibling specs; the final ledger contains zero `UNVERIFIED` rows.
- [x] There are zero unresolved `SPEC_STALE`, `CODE_STALE`,
      `CROSS_SPEC_CONTRADICTION` or `PROPOSAL_LEAK` rows.
- [x] WIP and Done are explicitly non-normative; no result relies on WIP prose
      as proof of current behavior.
- [x] Every direct WIP Markdown file is indexed exactly once.
- [x] Active, Hold, Decision and Retired are distinct queues.
- [x] WIP contains no unowned planning scratch or continuation baton; the exact
      human-managed `Notes_Scratch.md` exclusion remains untouched and outside
      plan indexing.
- [x] The WIP index removes completed Terminal v1 from the active queue and
      routes current behavior to the Terminal domain.
- [x] No Terminal WIP plan remains; deferred advanced features are explicit v1
      non-goals rather than hidden unfinished checklist items.
- [x] No current normative spec contains an executable future plan/checklist
      unless an explicit reviewed exception names why it is normative.
- [x] Normative/WIP local links and heading anchors pass.
- [x] No protected runtime/schema/evidence/prototype asset changed accidentally.
- [x] New and existing Atlas-focused Pester suites pass; independently owned
      Terminal failures are explicitly classified and routed.
- [x] Full completes on the exact final binaries; any non-green unrelated or
      environmental results are classified above without a false green claim.
- [x] `git diff --check` passes.
- [x] Final `git status --short` distinguishes intended Atlas changes from the
      recorded unrelated user work; no false clean-status claim is made.
- [x] This plan is moved to `Specs/Plans/Done/` and removed from the active WIP
      queue only after every preceding gate passes.

## STOP conditions

Stop and report; do not improvise if:

- any required cleanup would modify a pre-existing `Specs/Plans/Done` history
  file rather than adding the explicitly authorized completed Atlas plan;
- a normative statement disagrees with implementation/tests and intended
  current behavior cannot be established from repository evidence;
- resolving a `CODE_STALE` row would change C++/ABI/schema/security/package
  behavior but intended current behavior is not conclusively established;
- a candidate obsolete section owns a current security, ABI, ownership,
  localization, accessibility, performance or data-safety requirement with no
  surviving destination;
- code/tests contradict two normative documents and current behavior cannot be
  established mechanically;
- a path is consumed by build/runtime/package code rather than documentation;
- a Terminal Gate-0 JSON artifact is used by a current release/upgrade path;
- a historical byte/digest contract would be invalidated;
- the link scanner cannot distinguish examples from links without adding a
  third-party parser;
- a focused gate fails twice after a reasonable plan-local correction;
- an applicable domain skill or mandatory performance-validation gate cannot be
  satisfied for a required code-side correction.

## Considered and rejected

- **Deleting or rewriting Done:** rejected by explicit user constraint.
- **Moving completed WIP into Done during Atlas:** rejected because it changes
  the protected Done tree.
- **Deleting every old WIP file:** rejected because immutable history links to
  some original paths; compact tombstones preserve navigation.
- **Treating all large documents as bad:** rejected; File Operations, VFS and
  SettingsStore may remain large when they have one coherent normative owner.
- **Deleting Terminal Gate-0 schemas now that Ghostty won:** rejected; current
  tools/tests and immutable history reference them.
- **One regex for Markdown links:** rejected by real false positives in fenced
  C++ and PowerShell examples.
- **Cleaning Reviews, root plans, TestRuns, Mockups or Themes in the same
  operation:** rejected as scope expansion and history/evidence/runtime risk.

## Maintenance notes

- The WIP index is a live queue, not a historical narrative. Closed detail goes
  to the plan's compact checkpoint and evidence links.
- Normative specs are updated as behavior changes; plans link to them and do not
  restate their full contracts.
- WIP remains non-normative until closeout. “Implementation-ready” describes
  execution quality, not authority.
- The new inventory runs in documentation-focused CI and in Full, but remains
  read-only and writes reports only under `.build`.
- Any future tombstone removal requires a fresh inbound-link inventory and
  explicit approval; age alone is not evidence of uselessness.
- Reviewers should scrutinize deleted requirement prose, not line-count
  reduction. Every deletion must be explainable by a surviving owner or a
  recorded rejection.
