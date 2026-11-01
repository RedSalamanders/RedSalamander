# Specification Authority Catalog

Status: current normative repository policy
Last reviewed: 2026-08-13

## Purpose

This catalog defines which repository artifacts state current requirements, which record proposed work or evidence, and how a mismatch is resolved. Location alone does not make prose authoritative.

Run `Tools/Get-SpecInventory.ps1` to classify the current tracked and non-ignored untracked tree, validate links and WIP ownership, reconcile the normative consistency ledger, protect completed-plan history, and emit disposable reports under `.build/SpecInventory/`. Run `Tools/Get-ToolInventory.ps1 -Validate` for the separate machine-enforced tooling catalog governed by `Testing/Testing_ToolingGovernance.md`.

## Authority classes

| Class | Location or identity | Meaning |
|---|---|---|
| Normative domain contract | `Specs/<Domain>/*.md`, except explicitly retired records below | Current intended behavior, UI, validation, ownership, and workflow requirements. |
| Normative repository policy | `AGENTS.md`, this catalog, and the testing policies under `Specs/Testing/` | Cross-domain engineering, test, performance, and closeout rules. |
| Public ABI contract | Exported interface headers under `Common/PlugInterfaces/` | Compiled plugin ABI shape. Specs explain semantics but must not duplicate a conflicting struct definition. |
| Machine contract | Schemas, manifests, resource tables, lock data, and reviewed machine-readable policy such as `Specs/SettingsStore.schema.json`, `Specs/Testing/FolderViewPerfBudgets.json5`, and `Tools/tool-inventory.json` | Exact data/build/runtime/tooling contract consumed by code or tools. |
| Active plan | `Specs/Plans/WIP/*.md`, except the exact human-managed note below | Non-normative proposal, decision queue, or execution record for unfinished work. Every direct planning file must have one status in the WIP index. |
| Active-plan index | `Specs/Plans/WIP/README.md` | Routing and ownership only; it does not define product behavior. |
| Human-managed note | `Specs/Plans/WIP/Notes_Scratch.md` | Explicitly outside specification and planning authority. Do not read, edit, delete, move, summarize, index, or route its contents without a direct user instruction naming that file. |
| Historical | `Specs/Plans/Done/`, `Specs/Reviews/`, and the retired domain records named below | Sequencing, rationale, review, or recovery evidence; never current behavior authority. The former root `plans/` queue was archived into Done and deleted on 2026-08-25. |
| Evidence | `Specs/TestRuns/` | Immutable run evidence interpreted under the current testing/performance contracts. A green historical run is not proof for changed code. |
| Runtime data | `Specs/Themes/` and other explicitly shipped data | Runtime-consumed content; associated prose is not automatically a product contract. |
| Historical data | `Specs/Terminal/TerminalEngineEvidenceCorpus/` and decision-input JSON | Reproducible decision input, not a live requirement. |
| Prototype | `Specs/Mockups/` | Visual exploration, not shipped UI authority. |
| External documentation | `docs/` and top-level user/contributor documentation | User-facing or contributor guidance that must agree with the current contract but does not silently override it. |

The following domain-path files are explicit historical tombstones, not normative contracts:

- `Specs/UI/UI_PreferencesDialog_MigrationHistory.md`
- `Specs/FileSystem/FileSystem_FtpScpSftpPerformanceComparison.md`

## Reconciliation rule

Neither prose nor implementation wins solely because of its file type. When code, tests, a public header, machine data, user documentation, or two specs disagree:

1. identify the owning current contract and all affected consumers;
2. inspect runtime behavior, public compatibility, tests, callers, and historical rationale;
3. correct the stale side when intent is mechanically conclusive;
4. if the intended behavior is a product or compatibility decision, preserve current safe behavior and create one indexed `DECISION` WIP owner;
5. update every affected normative spec, test, schema/resource/manifest, and user-facing document in the same closeout;
6. record the completed review in `Specs/NormativeConsistency.json` with real implementation, test, and sibling-spec anchors.

Do not copy a public ABI definition, generated inventory, mutable count, dependency version, or machine schema into prose when the owning artifact can be linked instead. Prose may state semantics and compatibility rules.

## Plan lifecycle

- New scoped work lives in `Specs/Plans/WIP/` and is indexed exactly once as `ACTIVE`, `HOLD`, `DECISION`, or `RETIRED`.
- A WIP index `Rank` is unique and never reused, including an ID that an archived plan claims in its Done H1 (`# <ID> ...`). Done text is frozen, so the active plan takes a new ID on a clash.
- WIP files cannot override current contracts. An implementation plan must name the normative specs it expects to change.
- Durable requirements discovered during work move into the appropriate domain contract before closeout.
- A completed plan moves to `Specs/Plans/Done/`; it is then immutable history. Corrections to current behavior belong in current specs or a new WIP, never by rewriting Done history.
- The same closeout change adds the exact new Done path to the reviewed `Test-RSProtectedDoneHistory` additional-path allowlist and its focused admission test; the test also rejects an unreviewed sibling path. This authorizes only that new immutable file and does not permit modifying existing Done history.
- The [I55 CI implementation record](Plans/Done/CI_OrganizationStandardization_2026-10-29.md) has exact-file admission after the owner's implementation/hosted-acceptance split. [I56](Plans/WIP/CI_HostedAcceptance_2026-10-31.md) retains failed, unavailable and pending qualification; admission does not close that follow-up or qualify a release.
- The [2026-10-16 File Operations card hierarchy record](Plans/Done/FileOperations_CardInformationHierarchy_2026-10-16.md) uses this exact-file admission; adjacent dates and additional companion artifacts remain unreviewed.
- The [October 20 waiting-card record](Plans/Done/FileOperations_WaitingCards_2026-10-20.md) has exact-file admission with sibling and companion rejection; its scoped validation does not establish Full or platform qualification.
- The five 2026-10-19 File Operations UI/screenshot closeout records and the
  [five-stash review](Plans/Done/StashReview_2026-10-19.md) have individual exact-file
  admissions. Tests reject each date-shifted sibling and an additional review
  companion. Admission records reviewed history; runtime qualification remains
  subject to its owning validation contract.
- Reviewed capture companions use the same exact-path admission. List each original
  artifact and generated navigation path individually; a gallery directory does not
  authorize additional images. The focused test also rejects an unreviewed image
  inside the admitted collection. Capture/build provenance remains separate proof.
- A retired WIP tombstone is bounded, names its retirement date/reason/current owner, and gives a `git show` recovery command.
- The exact human-managed `Specs/Plans/WIP/Notes_Scratch.md` file is excluded
  from WIP indexing and content analysis. Other scratch notes and continuation
  batons require a named status and owner; unowned planning state is not a
  valid closeout condition.

## Pinned runtime and evidence paths

These locations are consumed by build/runtime/test workflows and must not be moved without updating their consumers:

- `Specs/SettingsStore.schema.json`
- `Specs/Themes/`
- `Specs/TestRuns/`
- `Specs/Testing/FolderViewPerfBudgets.json5`
- `Specs/NormativeConsistency.json`

## Naming

- Domain specs: `Domain_Title.md` without spaces.
- RFCs: `RFC_Domain_Title.md` under WIP while undecided; move them to Done after decision/implementation closeout.
- Other plans: descriptive `_YYYY-MM-DD.md` names under WIP; move, do not copy, when complete.

## Primary domain entry points

| Domain | Start with |
|---|---|
| Core | `Core/Core_SharedHelpers.md`, `Core/Core_SettingsStore.md`, `Core/Core_Search.md`, `Core/Core_ConnectionManager.md`, `Core/Core_StartupBootstrap.md` |
| UI | `UI/UI_DxUiWinUIDesign.md`, `UI/UI_FolderWindow.md`, `UI/UI_FolderView.md`, `UI/UI_DxUiSharedGrid.md`, `UI/UI_TopLevelToolWindows.md`, `UI/UI_PreferencesDialog.md` |
| File systems | `FileSystem/FileSystem_FileOperations.md`, `FileSystem/FileSystem_FtpSftpScp.md`, and the provider-specific specs |
| Plugins | `Plugins/Plugins_PluginAPI.md`, `Plugins/Plugins_VirtualFileSystem.md`, `Plugins/Plugins_ViewerPlugins.md` |
| Installer/release | `Installer/Installer_Msix.md`, `Installer/Installer_Msi.md`, `Installer/WingetIntegration.md` |
| Terminal | `Terminal/Terminal_EmbeddedPlugin.md`, `Terminal/TerminalEngineDecision.md` |
| Testing and tooling | `Testing/Testing_SelfTests.md`, `Testing/Testing_TestCoverage.md`, `Testing/Testing_PerformanceValidation.md`, `Testing/Testing_SelfTestRemoteCredentials.md`, `Testing/Testing_ToolingGovernance.md`, `Testing/Testing_AgentGuidance.md`, `Testing/Testing_ValidationEvidence.md`, `TestRuns/README.md` |

## Required validation

Before closing specification-governance work:

```powershell
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Get-ToolInventory.ps1 -Validate
Invoke-Pester .\Tools\Tests\SpecInformationArchitecture.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\DocumentationDriftContracts.Tests.ps1 -PassThru
.\Tools\Run-AllTests.ps1 -Suite Full
```

`-FailOnFindings` requires every normative file to have a `VERIFIED_MATCH` consistency row, every direct WIP file to be indexed exactly once, WIP index ranks to be unique and not to reuse an ID claimed by an archived plan title, all current local links/anchors to resolve, normative plan markers to be explicitly reviewed, and the protected Done baseline to remain byte-identical.

The canonical Markdown code mask in `Tools/Modules/Tooling/SpecInformationArchitecture.psm1`
preserves text length and line endings for link and heading diagnostics. Its supported
top-level fence opener/closer may have zero to three leading spaces; tabs or four spaces
MUST NOT change fence state or hide following unindented prose. A backtick fence
opening's info tail MUST contain no backticks; an invalid opener continues through inline-code
masking and MUST NOT hide following prose. Tilde fence info may contain backticks. A closing
fence MUST use the opening character with at least the opening run length and a whitespace-only
tail. A fence-like line with trailing nonwhitespace stays inside its current fence.
These checks do not expand the parser to general Markdown lists or indented code blocks.

The [Changes workflow](Core/Core_Changes.md) owns fragment format, validation and folding on the default branch after merge; a change keeps its fragment and never edits `CHANGELOG.md`.

The snapshot migration retains the original 133-file protected Done baseline as parentless commit 4a3a6c9dc3d094cb2269c438de7f3ae24ab5e853. Its Specs/Plans/Done tree is byte-identical to the former baseline. Both integrated master and the public source snapshot retain that commit in ancestry, so ordinary fresh clones can enforce the existing guard without importing DualTail history or changing completed documents.

Common native build policy, zero-warning acceptance, terminal output and short consumer paths: [Build_ConsoleAndPaths](Build/Build_ConsoleAndPaths.md).
