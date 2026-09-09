# Test Coverage Contract

Status: current normative policy and risk map
Last reviewed: 2026-08-09

## Purpose and authority

This document defines which risks RedSalamander test suites must cover and how coverage claims are evaluated. It deliberately does not enumerate mutable case names or totals.

The live inventory authorities are:

- `RedSalamander.exe --selftest-list-cases` for runner-native Commands, CompareDirectories, and FileOperations cases;
- `Tools/Get-TestInventory.ps1 -Format Json` for source registrations, native test projects, execution kinds, CI/Full run-plan membership, tooling tests, and source-contract classifications;
- `Tools/Modules/Testing/TestSuitePlan.psm1` (exported through the
  `TestRunPlan.psm1` facade) for CI and Full execution membership;
- suite result artifacts for what actually ran, passed, failed, crashed, or skipped.

`Tools/Get-SpecInventory.ps1` publishes the source-derived manifest as `.build/SpecInventory/test-inventory.json` beside the specification inventory. Generated `.build` output is disposable and must not be checked in. Current totals must not be copied into normative prose.

`Testing_SelfTests.md` owns result, skip, repeat, shuffle, sandbox, and runner aggregation semantics. `Testing_PerformanceValidation.md` owns metrics, sample quality, budgets, baselines, and archived performance evidence. `Testing_SelfTestRemoteCredentials.md` owns opt-in live remote prerequisites.

## Trusted validation coverage contract

The live source/runtime inventory and the canonical Full plan remain authoritative after
[`Operation Startrail`](../Plans/Done/Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md)
implementation. A test family, filter, affected plan, or partial capability environment must never be
reported as complete Full coverage.

The active validation-plan catalog makes each entry's stable ID, coverage set,
repeat/order/filter identity, outcome and skip policy, artifact read/write/production
sets, environment dependencies, and cacheability explicit. Capability-dependent skips
remain observed evidence, but are not automatically reusable after capability or policy
changes. Quarantine projections and static/runtime inventory parity must pass before an
affected-set decision may exclude a declared surface.

`ToolsPesterBuildToolchain` is the explicit artifact-mutating entry for
`RequiresBuildToolchain`; the ordinary Tools Pester entry excludes that tag. Inventory
exports each CI/Full entry ID, policy, impact tags, artifact sets, resource classes, and
metadata-completeness verdict. A new or changed runner surface is incomplete until that
inventory remains complete and the validation-plan schema/digest tests pass.

Commands exposes 14 governed source families through `--selftest-family=<name>` and
`Run-AllTests.ps1 -Suite Commands -CommandsFamily <name>`. The native family inventories
must have an exact duplicate-free union equal to broad `--commands-selftest` membership.
Family execution is focused coverage only: because the payload is still the monolithic
application executable, it is non-cacheable and cannot satisfy the broad Commands Full
entry. The broad same-process seal remains mandatory for Fresh Full and for any Resume
whose exact broad promotion is unavailable.

## Coverage standards

Every durable behavior change must have a test disposition before closeout:

1. Add or update deterministic behavioral coverage at the lowest useful boundary.
2. Use an integration or UI test when ownership, threading, message routing, lifetime, persistence, accessibility, or cross-component behavior is part of the contract.
3. Use a source-contract test only for a property that is inherently lexical or structural, such as a forbidden API, required manifest entry, or dependency graph invariant.
4. Do not treat a source-shape assertion as proof of runtime behavior. `BehavioralShadow` and behavioral portions of `MixedSourceShape` entries reported by the generated inventory remain a replacement queue and must move to runtime coverage when a deterministic seam exists.
5. When deterministic coverage is impossible on the current harness, record the missing seam and owner in an indexed WIP file. Do not hide the gap in this normative document.
6. Performance-sensitive changes require the instrumentation, deterministic scenario, before/after evidence, and archive rules in `Testing_PerformanceValidation.md` from the start.

A passing case proves only its named contract and inputs. A source search, build success, manual observation, or historical green run cannot substitute for relevant current behavioral evidence.

## Required risk coverage

### Commands and application UI

Commands coverage must exercise command registration and dispatch, enabled/disabled state, keyboard and pointer routing, focus restoration, cancellation, persistence boundaries, and visible error handling. UI contracts must cover the real owned window and control path, not a helper-only surrogate.

High-risk surfaces include:

- SettingsStore reload, save coordination, future-schema rejection, and lossless unknown-member preservation;
- Preferences, shortcuts, connection management, plugin configuration, Batch Rename, Find Files, Compare options/progress, File Operations issues and popup surfaces;
- NavigationView, FolderView, menus, accelerators, function/status bars, accessibility, DPI/theme/high-contrast/reduced-motion behavior;
- viewer launch, focus, navigation, close, and plugin lifetime paths;
- external-action selected-path manifests across private-root exclusive creation/collision exhaustion, exact UTF-16LE record bytes and invalid provider records, span-based bounded streaming, real child parsing/replacement, pre-launch faults, wait registration, asynchronous completion, timeout, missing process handles, exit-code failure, nonzero exit, same-path replacement/root-replacement/reparse identity, no-follow same-handle cleanup, and bounded-progress private-root recovery; `file_action_selected_paths_identity_cleanup` and `file_action_selected_paths_recovery_bounds` own those exact-object/recovery boundaries, and shared `%TEMP%` prefix matching is never ownership evidence;
- cross-thread posted-payload ownership, cancellation, teardown draining, and stale HWND/token handling.

### Compare, search, and indexing

CompareDirectories coverage must include matching/classification rules, display modes, asynchronous size/content work, cancellation, partial/provider failure, result stability, and settings round trips.

Search coverage must include local scan, host fallback, indexed service behavior, schema compatibility, store recovery/compaction, writer ownership, cancellation/timeouts, degraded operation, security boundaries, and isolation from stale state. Process-backed tests must capture early-exit diagnostics and must not consume the first state observation in a generic readiness probe when that observation is part of the contract.

### File Operations

FileOperations coverage must include copy, move, delete, rename, conflict resolution, cancellation, pause/resume, queue ordering, rollback/cleanup debt, metadata, reparse-point policy, alternate-volume behavior, watcher refresh, bandwidth/concurrency controls, progress/accounting, and local-to-remote/remote-to-local/remote-to-remote bridges.

Provider tests must prove exact destination/source identity and byte preservation where applicable. A simulated success response is insufficient when the contract depends on server-acknowledged offsets, revisions, IDs, parents, names, eTags, continuation tokens, or cleanup convergence.

Remaining deterministic stress validation for the File Operations popup, large listings, concurrency, delete paths, and live Queue/Parallel switching is owned by `Specs/Plans/WIP/FileOperations_RemainingStressValidation_2026-08-04.md` until closed into this domain’s runtime tests and authoritative behavior specs.

### Plugins and providers

Plugin coverage must validate public ABI size/version negotiation, capability honesty, factory/enumeration behavior, callback cookies, configuration transactions, unknown-member round trips, secret handling, unload quiet points, and module/callback lifetime.

Each shipped file-system provider must cover its distinctive identity, paging/listing, metadata, read/write, copy/move/delete, cancellation, retry, error, and configuration behavior. Live remote cases stay declared and skip with a specific prerequisite reason when credentials or environment capabilities are absent; they must not disappear from inventory.

Shared provider temporary-file coverage must prove reservation cleanup when reopen fails, delete-on-close behavior, concurrent pathname uniqueness, and preservation of provider-selected prefixes and flags. Source-contract coverage may enforce consolidation and static flag policy, but runtime handle behavior requires a native test.

MTP/WPD tests must cover COM MTA worker initialization, serialized device sessions, disconnect/quarantine/cancel teardown, and honest read-only/writable capabilities. Curl, cloud, archive, local, dummy, and S3-family providers require deterministic fake-server or fake-transport coverage for protocol edge cases where practical.

Terminal input coverage must assert exact encoded bytes, not only consumed message
status. It includes selected/unselected Ctrl+C and queued translated-character
suppression, independent application-mouse button chords and teardown release,
canonical extended/device path classification, source-generation transitions,
unsupported-provider fail-closed behavior, and a real typed plugin-backed pane
location through the staged Terminal DLL.

### DxUi, graphics, and viewers

DxUi coverage must exercise control semantics, layout, scrolling, focus/traversal, keyboard/pointer input, native text input and IME, accessibility providers, theme/DPI/high contrast/reduced motion, device loss, window lifetime, and retained rendering behavior.

Pixel or gallery evidence can supplement but cannot replace semantic and interaction assertions. Visual tests must state tolerances and avoid depending on one machine’s animation phase, subpixel rasterization, or foreground ownership unless the prerequisite is explicitly checked.

Viewer coverage must include plugin discovery, file parsing/streaming limits, navigation, shared chrome and menu models, focus/close behavior, theme changes, malformed input, and any format-specific security boundaries. Large or streaming formats require bounded-memory and cancellation proof.

### Build, packaging, localization, and tools

Tooling tests must guard:

- project selection, sanitized child environments, build locking, vcpkg identity and safe staging;
- Terminal-engine raw-byte inputs under LF and `core.autocrlf=true`-equivalent checkouts,
  localized valid tar metadata, malicious archives, empty-online/warm-offline dependency
  caches, separately tampered dependency locks/licenses, and actual x64/ARM64 PE imports;
- centralized runtime dependency staging, clean portable extraction, architecture identity, MSI/MSIX and winget publication contracts;
- CI/Full run-plan parity, result coverage, timeout/failure classification, quarantine/repair lanes, and TestSandbox containment;
- resource ownership, positional placeholder equivalence, satellite fallback, and settings/localization round trips;
- source-policy bans and shared-helper adoption where runtime observation is not the relevant property;
- specification inventory, link/anchor integrity, authority classification, WIP ownership, consistency evidence, and protected Done history.

Tests that mutate dependency or build state and cannot be made isolated must be explicitly excluded from normal CI/Full membership and documented as manual-only. Artifact-safe tooling tests must not silently depend on a local compiler installation.

### Crash, diagnostics, and monitor

Crash coverage must preserve partial results and actionable diagnostics across exceptions, access violations, hangs/timeouts, child-process termination, and dump/sidecar collection. Test code must not swallow failures behind broad exception handling.

Monitor coverage must validate ETW provider compatibility, self-originated event suppression, runtime diagnostic
gating, bounded ingestion/display behavior, and stable ColorTextView interaction. It must also cover decoded
expansion limits, split UTF-8/UTF-16 boundaries, malformed/truncated input, mid-decode cancellation, move-publication
without UI reparsing, repeated search-cap eviction/refill and F3 order, Show IDs coordinate reset/rebuild, exact
canonical read/export/read records (including empty, terminal, blank, Unicode, and line-budget boundaries),
start-snapshot asynchronous export across later append/retention/clear/replacement, zero-text-copy immutable capture,
open/save thread-admission rollback and retry, cancellation/fault target preservation, bounded worker close,
registered payload drain, three-size/five-repeat Release scaling, and selftest fixture cleanup. Optional richer clipboard formats and filter semantics are owned by
`Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md` until decided.

## Conditional and destructive scenarios

Conditional cases remain members of their suite. Missing credentials, ReFS/alternate volumes, foreground desktop ownership, optional plugins, or other declared prerequisites produce a justified `skipped` result under `Testing_SelfTests.md`; they do not remove the case.

Destructive and capability-sensitive tests must use test-owned roots, profiles, accounts, buckets, mailboxes, or devices. They must fail closed when containment cannot be proven and clean up only exact owned resources. Real external-provider tests must not run from ordinary CI without an explicit opt-in contract.

## Coverage review and closeout

A coverage review is complete only when:

- the generated inventory and runner-native listing agree with the intended suite membership;
- every changed contract has behavioral, integration, performance, or explicitly justified structural coverage;
- skips and unavailable environmental paths are visible in result artifacts;
- relevant focused tests pass, then the required CI/Full gate passes;
- performance-sensitive work has qualifying archived evidence;
- unresolved deterministic seams are in an indexed WIP owner;
- durable coverage rules are recorded here or in the narrower owning spec, without copying transient case totals or run diaries.

Use these commands for review:

```powershell
.\Tools\Get-TestInventory.ps1 -Format Json
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\.build\x64\Debug\RedSalamander.exe --selftest-list-cases
.\Tools\Run-AllTests.ps1 -Suite Full
```
