# Test Coverage Contract

Status: current normative policy and risk map
Last reviewed: 2026-08-09

## Nonactivating UIA evidence

Nonactivating tests do not own the person's desktop. Foreground witnesses attribute
events to the test process and retain transient owned activation; ambient pointer,
foreign foreground and last-input changes remain separate read-only observations.
Source contracts forbid physical input writes in the UIA observer and fixture path.
Unavailable guards/observation coverage fail closed. These checks complement native
UIA behavior and do not substitute for real-input, assistive-technology or interactive
qualification.

## Purpose and authority

Archive export coverage must exercise source paths containing each accepted separator,
including a Windows-joined nested path, through Preparing and the host Copy bridge to
both a local destination and a sandbox-only UNC destination. Assert exact output bytes,
the single selected leaf, and continued destination rejection of reserved names, trailing
dots and embedded separators. Include a multipart archive fixture. An unavailable SMB
test route remains an explicit capability skip and cannot count as network qualification.
Typed and diagnostic archive separator contracts must also agree in the standalone
plugin contract tests.

This document defines which risks RedSalamander test suites must cover and how coverage claims are evaluated. It deliberately does not enumerate mutable case names or totals.

The live inventory authorities are:

- `RedSalamander.exe --selftest-list-cases` for runner-native Commands, CompareDirectories, and FileOperations cases;
- `Tools/Get-TestInventory.ps1 -Format Json` for source registrations, native test projects, execution kinds, CI/Full run-plan membership, tooling tests, and source-contract classifications;
- `Tools/Modules/Testing/TestSuitePlan.psm1` (exported through the
  `TestRunPlan.psm1` facade) for CI and Full execution membership;
- suite result artifacts for what actually ran, passed, failed, crashed, or skipped.

`Tools/Get-SpecInventory.ps1` publishes the source-derived manifest as `.build/SpecInventory/test-inventory.json` beside the specification inventory. Generated `.build` output is disposable and must not be checked in. Current totals must not be copied into normative prose.

`Testing_SelfTests.md` owns result, skip, repeat, shuffle, sandbox, and runner aggregation semantics. `Testing_PerformanceValidation.md` owns metrics, sample quality, budgets, baselines, and archived performance evidence. `Testing_SelfTestRemoteCredentials.md` owns opt-in live remote prerequisites.

Named documentation galleries such as `FileOps_VisualGallery` are explicit opt-in rendering fixtures.
They are excluded from broad behavioral phase totals and phase-order integrity requirements, while
remaining selectable by their exact case name. Their screenshot count is not a behavioral test count.
The paired `FileOps_CommonPresentationPerf` and `FileOps_IntegratedPresentationPerf` fixtures follow the same exact-selection rule: their
measurements remain outside broad behavioral phase totals and do not establish full-suite coverage.

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

Structural source guards must tolerate formatter whitespace without weakening required arguments,
call boundaries, ordering or forbidden APIs. The File Operations durable-store guard matches the
complete publication call and fail-if-exists/replace policy; native durable-store cases remain the
authority for atomic publication, rejected paths and recovery behavior.

Application production foreground requests route through `Common::RequestForegroundWindow`.
The source contract rejects direct bypasses, and physical-input/activation-write bans recognize
the wrapper too. Real-input predecessor plus documentation-gallery runs provide the behavioral
regression evidence; structural routing alone does not qualify no-activation behavior.

Commands clipboard setup must account for desktop listeners notified by fixture clears and seeds.
The shared fixture quiet-point helper waits for an unchanged sequence and no open owner for 250 ms,
with a final availability probe and a three-second scaled deadline. Product copy/cut/paste commands
still execute once, and their exact text/format assertions remain mandatory. A setup timeout remains
visible; neither command retries nor a cached fallback establish a native clipboard success.
Retain failed attempts and compare repeated same-source runs before attributing external contention.

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

Find header-input coverage must use DPI-correct client coordinates, verify the intended logical
column and prove intermediate sort-order changes. Restored sorting must preserve the exact selected
path, reordered columns and resized widths; an unchanged preexisting order cannot prove a click.

### File Operations

FileOperations coverage must include copy, move, delete, rename, conflict resolution, cancellation, pause/resume, queue ordering, rollback/cleanup debt, metadata, reparse-point policy, alternate-volume behavior, watcher refresh, bandwidth/concurrency controls, progress/accounting, and local-to-remote/remote-to-local/remote-to-remote bridges.

Consent coverage must preserve stale bindings across replacement prompts on the same task, including
UIA, overflow menus, scope toggles and Escape, and prove one-shot acceptance under concurrent
submitters. Cancellation coverage must distinguish aggregate status from storage evidence: missing
receipts after possible mutation cannot prove source retention or destination publication. Known,
unknown, malformed, null and absent receipts, explicit/partial Skip, and stronger host-managed facts
require distinct assertions. Known no-commit receipts with an absent original or Published owned stage
must be rejected through both the pure reducer and the real bulk callback/dense finalizer, across
Copy/Move/Delete, success/Skip/error/cancel statuses and subsequent task cancellation. Valid no-commit
controls must remain precise. UI lifecycle coverage must exercise real hidden/minimized windows,
unchanged-model work suppression, retained path-detail state without native focus theft, alert-provider
presentation lifetime, and rollback/retry after every partial native failure-surface allocation.
Parallel recursive-copy receipts must include child mutation facts without mixing selected roots.
Bulk Local Copy/Move must report per-index receipts in both sequential and parallel branches;
a successful first root cannot supply committed truth to a later failed root. Unknown cleanup
remains independent of known final non-publication, and callback order is not item identity.
One-shot and apply-to-all collision controls must validate final typed storage axes as well as bytes.
UI selection coverage must include programmatic selection, real UIA clear/reselect, append, and
removal of the selected stable row with no stale detail text. Fixtures that admit real tasks must
drain and dismiss their own terminal UI through the production path before returning to the suite.
Successful no-op Copy merges must preserve `Completed` independently of `NotPublished`: cover an
empty source directory, an existing destination with sentinel data, and an entirely existing empty
directory tree at serial and parallel concurrency. A new destination is a committed control.
Known no-commit Move/Delete success, Copy `S_FALSE` without Skip, and unresolved stage debt must
retain their failure/indeterminate classifications.

Actual Local storage witnesses must cross their named production boundary: cancellation after
an observed child removal, cross-volume Managed Move failure after an observed publication and
later positive stage write, locked replacement-authority refusal through normal consent, native
stage-write failure after positive I/O, ordinary-child DACL denial, exact case-only rename, owned
application shutdown during actual stage I/O, and two intentional queued non-clipboard commands.
Prove exact before/after bytes and canonical BLAKE3, identity where required, bounded nofollow
inventories, real stage disposition and separate producer/callback/finalizer axes after quiet.
Cancellation is not an injected write failure; a retained selected root is not proof that no child
changed, and leafwise Managed Move is not an atomic whole-tree transaction. A Dummy Copy-only
source-retention control cannot stand in for actual cross-volume partial publication.

Failure/lifetime coverage must include empty synchronous admission with no created task,
same-path/hard-link storage preservation, stale/concurrent normal consent, valid and contradictory
receipts, visible Stopping plus hidden Preparing cancellation, all-post-failure owner recovery,
worker-return gating and actual completion-callback owner destruction. Private shutdown-child
coverage must prove natural application close, owned settings persistence, quiet-before-destroy
and external storage after real process exit. Bounded job containment or a killed child is failure
evidence. Child-only environment overrides require native inheritance/override controls without
changing the parent's environment.

Admission-performance coverage measures actual UI-thread admission, not provider helper calls
alone. Retain cold/prime/warm-requested labels, actual cache hits/misses and prime age, real volume
API and allocation counters, immutable plan partitions, cancellation-before-consumption proof,
exact closed inventory and foreground/pane disposition. Mounted routes use only the test-owned
VHDX volume that an elevated harness provisions, proves and mounts inside the current run; the test
validates that hand-off and never mounts, unmounts or adopts a volume, and mounting a real user
volume is forbidden (see the sandbox rules in `Testing_SelfTests.md`). Capability skips stay explicit and cannot be counted as measured route coverage or
performance non-regression. Focused families, test-enabled Release, native alternate-volume and
mounted qualification, repeat/order runs, sanitizers and architecture gates remain separate claims.

Failed evidence must survive worker retirement and safe settings/security restoration. Preserve
actual failed axes/bytes/inventories/diagnostics and owned scratch; stop same-process continuation
and refuse automatic reuse after retained failure. Never remove an ambiguous owned mount or tree
to make a later run green. Registration/catalog parity proves declared reachability only; runtime
receipts and the canonical Full plan remain required for executed coverage.

Provider tests must prove exact destination/source identity and byte preservation where applicable. A simulated success response is insufficient when the contract depends on server-acknowledged offsets, revisions, IDs, parents, names, eTags, continuation tokens, or cleanup convergence.

Remaining deterministic stress validation for the File Operations popup, large listings, concurrency, delete paths, and live Queue/Parallel switching is owned by `Specs/Plans/WIP/FileOperations_RemainingStressValidation_2026-08-04.md` until closed into this domain’s runtime tests and authoritative behavior specs.

### Plugins and providers

Plugin coverage must validate public ABI size/version negotiation, capability honesty, factory/enumeration behavior, callback cookies, configuration transactions, unknown-member round trips, secret handling, unload quiet points, and module/callback lifetime.

When one DLL exposes several logical plugin IDs, mixed unload outcomes, repeated refresh and
partial deferred sweeps must preserve every alias without retaining an extra executable owner.
After the final unload settles, discovery must restore all aliases. Cover this shared-module
boundary directly as well as through broad-process plugin discovery.

Every plugin exposing `IInformations` must prove it accepts the defaults its own schema declares. `Tests/PluginContractTests` drives `GetConfigurationSchema` through `Common::PluginConfiguration::ParseSchema` + `MakeDefaultValue` + `SerializeConfiguration` — the exact path the Preferences page uses — requires `SetConfiguration` to return `S_OK`, and then requires the resulting `GetConfiguration` output to still satisfy the schema and be re-acceptable. This closes the gap where a plugin declares an integer-serialized `value` field and rejects the integer the shared codec emits for it, making a transactional `SetConfiguration` discard every other member and silently revert to compiled defaults. The Terminal is covered by the same transactional proofs as the file-system providers (unknown-member preservation, malformed and non-object roots reporting `ERROR_INVALID_DATA`, `{}` restoring defaults). Host-side plumbing is covered by the `settings_terminal_plugin_roundtrip` Commands case, which stores an integer `fontSizeDip` alongside a `fontFamily` and asserts the family survives `CreateTerminalInstance` → `ApplyConfigurationFromSettings`.

Each shipped file-system provider must cover its distinctive identity, paging/listing, metadata, read/write, copy/move/delete, cancellation, retry, error, and configuration behavior. Live remote cases stay declared and skip with a specific prerequisite reason when credentials or environment capabilities are absent; they must not disappear from inventory.

Shared provider temporary-file coverage must prove reservation cleanup when reopen fails, delete-on-close behavior, concurrent pathname uniqueness, and preservation of provider-selected prefixes and flags. Source-contract coverage may enforce consolidation and static flag policy, but runtime handle behavior requires a native test.

MTP/WPD tests must cover COM MTA worker initialization, serialized device sessions, disconnect/quarantine/cancel teardown, and honest read-only/writable capabilities. Curl, cloud, archive, local, dummy, and S3-family providers require deterministic fake-server or fake-transport coverage for protocol edge cases where practical.

Terminal input coverage must assert exact encoded bytes, not only consumed message
status. It includes selected/unselected Ctrl+C and queued translated-character
suppression, independent application-mouse button chords and teardown release,
canonical extended/device path classification, source-generation transitions,
unsupported-provider fail-closed behavior, and a real typed plugin-backed pane
location through the staged Terminal DLL.

Terminal font coverage must prove that a configured family which is not installed, or
which lacks private-use (Nerd Font) glyph coverage, reaches the user as the non-blocking
in-grid notice rather than degrading silently: the monospace-only fallback chain, the
one-shot glyph probe that fires only once such a codepoint is painted, banner dismissal
and re-arming when the configured family changes, and the accessibility snapshot text.
DxUi's standalone Foundation suite covers the shared typography side: availability memoization, the
`checkForUpdates` rescan, and `InvalidateFontFamilyAvailability` dropping cached answers
for one factory.

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

RedMonitor client integration coverage validates the immutable local header/license, direct producer build
verification, legacy ETW GUID/field compatibility, normalized application/severity/Unicode/PID/TID/time,
absent-helper behavior, installed production Release discovery and explicit/optional Debug startup routing.
Preferences tests retain live filter/mask application and the shared Release settings-file link. Consumer
builds, test plans and packages must contain no viewer target, executable or satellite dependency.

The standalone viewer's detailed model, ingestion, file I/O, interaction, lifetime and Release scaling
obligations transferred to [RedMonitor validation](https://github.com/RedSalamanders/RedMonitor/blob/main/Specs/Testing/Testing_Validation.md).
Historical viewer receipts remain historical; consumer component passes cannot qualify those viewer gates.
The impact manifest maps the vendored client and its lock to the full consumer build/test closure and
does not select retired viewer projects or suites.

## Conditional and destructive scenarios

The File Operations safety contract exercises repeated Enter, Space and Escape through the owned
popup's native message route after prompt A's controls are replaced by generation B. Its synthetic
fixture observes actual hosted callbacks, requires exact B default/cancel identities, retains fresh-press
positive controls and repeated-navigation coverage, and preserves native foreground/focus ownership.
The Commands stale-presentation case separately checks that held Escape leaves a real B prompt
unresolved after its cancel binding refreshes. These message-dispatch checks do not replace physical
keyboard qualification under the interactive desktop lease.

The legacy FileOps archive read/seek and mounted-pane backing-archive change cases remain explicit
coverage gaps. Compare's 7z directory-size and name-only-search tests do not qualify those scenarios.

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

## I19 product UI regressions

`ProductUiTests` is the active CI/Full consumer target for all 23 product-owned cases
in the I19 gap analysis, plus viewer adapter mapping/polarity regressions. It links
the pinned archive through public headers. The legacy library source/suite is retired;
its old results are not candidate coverage. Shared control/embedded tests belong to DxUi. Commands,
viewer tests and RedConfigure continue to establish product behavior. RedMonitor client and Preferences
integration is consumer coverage; standalone viewer behavior belongs to its independent project.

Compare's low-hardening smoke keeps the Find checkbox repaint contract as an
interaction with the pinned public library. Left on an unchecked mixed checkbox
and Right on a checked mixed checkbox must clear the mixed glyph and request
repaint while leaving the Boolean value unchanged. Consumer qualification must
not depend on reading a retired shared-library source file for that behavior.

The Preferences category-churn regression obtains current item rectangles through the
product debug adapter and clips them to the visible client area. It must not
assume content-space row coordinates remain client coordinates after scrolling.
The adapter only reads geometry; actual input still selects the category, and
the existing focus, retained-host and invalidation-budget assertions remain.

The extracted control target does not own product source policies. `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
retains the NavigationView delivered-input generation prohibition and the credential prompt modal-quit, raw-child
teardown and joined-UIA-worker policies. These four checks remain in the product tooling suite alongside runtime
Commands coverage when the old DxUiTests executable is retired.

ViewerSpace's authoritative-progress source guard follows its monotonic observed-byte
helper across the scanner and model translation units. Native gated summary coverage
must additionally bind the early and settled totals to the same named, presented real
folder tile; source-pattern checks alone do not qualify accounting or presentation.

## Folder Copy live progress

The two exact no-activation cases are `DiscoveryScope_FolderLiveProgress` and
`DiscoveryScope_MultipleFoldersLiveProgress`, registered in the existing DiscoveryClosure source
and runner. The first uses real serial/overwrite/parallel Native, parallel bridge and Verify bridge
copies, independent 4-file/3-folder/96-MiB totals/content, rendered found-count ink and same-file
advancing track/fill pixels. Its slow-finalization and exact small-file Publishing pauses check
real meter lifetime and retained size. A separate sequential bridge traversal uses 32 nonempty
sibling folders beyond the four-directory lookahead (32 files, 33 folders, 64 MiB at 2 MiB/s).
It checks independently defined positive open count/size prefixes and visible text ink;
empty roots and actual publication followed by committed-error/proof failure check count truth
and typed root failure. Sampling can use at most three actual child attempts with retained misses;
it never weakens fractions/pixels or changes the five paired routes' speed/delay.
Only these five meter routes and the wide open probe cap each harness message batch at eight,
preventing ordinary queued paint traffic from starving their 100-ms sampling. The original
first named file must remain in actual nonterminal Finalizing for the 650-ms meter witness;
phase/full-byte/terminal transitions remain diagnostic evidence. Other closure cases keep the
existing 1,024-message default.

Multiple selected roots cover monotonic counters, bounded rows, simultaneous regular workers,
file-shaped preserved links with zero payload, literal link targets and independent file bytes.
The direct Native `CopyItems` control uses real I/O plus the production reducer on a non-admitted
task, separately from product PerItem popup coverage. Synthetic reducer/presentation controls
cover all-scope support, zero work, Skip/cancel, publication+retry, stale terminal, producer/worker
coexistence, replay uncertainty, known-size summation, partial-total presentation, Verify On zero
bytes and malformed enum/BOOL/immutable-kind withdrawal while independent certified size survives.
Only those two exact cases have the reviewed no-activation runner classification; broad FileOps
families retain their foreground prerequisites. Focused results never assert a Full verdict.

One ordinary large route samples automatic WM_PAINT records without forcing redraw or repeatedly
showing the popup. It requires distinct advancing generations for the same file, excludes every
capture-induced generation, and retains the same pixel probes afterward. Other routes explicitly
use synchronous redraw for their renderer/phase witnesses. Semantic-link I/O requires Windows
file-link creation privilege (or Developer Mode already enabled). ERROR_PRIVILEGE_NOT_HELD(1314)
records named skipped link subcases and their reason; regular-file PerItem and Native Bulk I/O
remain required and cannot be skipped. Synthetic producer controls remain distinct from actual
Native semantic-link qualification. No test changes the machine's privilege or Developer Mode.

## Supplied-review safety regressions (2026-10-16)

ProductUiTests covers FULL, ID_BOTH and notification record truncation, record-local name bounds, alignment, nonprogress/overflow, unsafe leaf/path names and valid layouts. It also covers payload destination/message/type mismatch preservation, rightful adoption, distinct trivial types whose deleters may fold in Release, duplicate/stale delivery, teardown and unclaimed synchronous reclamation, plus deterministic entropy failure, zero and live-collision exhaustion. Existing coalescing/order and teardown stress remain required.

`FileOps_ReviewRemediationContracts` owns descendant parent pin/race and cancellation coverage, plus XP-09 typed skip/refusal outcomes. `FileOps_SmallFileCopyCost` validates actual names and content on both routes and rejects missing/corrupt destination controls. ScopedTesting Pester fixtures change reusable workflows, composite actions, assets and imported scripts, and verify that a committed emptied reusable gate cannot delegate obligations. These focused checks do not qualify native ARM64, real hostile-server input delivery or all SMB/NAS implementations.
