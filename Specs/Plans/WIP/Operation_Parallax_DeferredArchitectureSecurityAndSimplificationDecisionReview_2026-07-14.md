# Operation Parallax — Deferred Architecture, Security, and Simplification Decision Review

> **Reviewer instructions:** This is a decision ledger, not an implementation plan. Review one scenario at a
> time, refresh its evidence against the live tree, record an explicit decision, and route any approved work to
> a separate executable WIP plan. Do not change production code from this operation. Do not turn a previous
> rejection into a defect merely because a matching name or pattern still exists; first prove the behavior,
> policy, or risk is actually equivalent.
>
> **Drift check (run first):**
>
> ```powershell
> git diff --stat 0bee16269..HEAD -- Common Plugins RedConfigure RedSalamander `
>   RedSalamanderMonitor Specs/Plugins Specs/Plans/WIP/Win32_Inventory.md
> ```
>
> If cited code, an authoritative spec, or an owner plan changed, inspect the live source and update the
> scenario's evidence before deciding. Line numbers are evidence anchors, not edit instructions.

## Status

- **Status:** HOLD — twelve Lighthouse dispositions preserved for an explicit later decision
- **Priority:** P3 decision review; individual routed work may receive a higher priority
- **Effort:** M for the complete review; S per scenario
- **Risk:** LOW for the review, potentially HIGH for resulting security/lifetime/policy changes
- **Depends on:** none for evidence review; resulting work must respect active owner plans
- **Category:** architecture, security policy, maintainability, tests, product direction
- **Planned at:** commit `0bee16269`, 2026-07-14
- **Source ledger:**
  `Operation_Lighthouse_WholeRepositoryAuditFindingsAndRemediationRouting_2026-07-10.md`
- **Current owner:** Operation Parallax owns future decisions for PAR-1 through PAR-12. Lighthouse retains the
  audit-time snapshot but must not become a parallel decision queue.

## Why this operation exists

Lighthouse correctly rejected several broad claims because the live code already contained the required
protection, the behavior was intentional, or similarly named implementations had different contracts. Those
rejections should not disappear: requirements, platform behavior, product policy, and abstraction pressure can
change. Parallax preserves the scenarios with enough context to revisit them deliberately without repeating the
original audit or accidentally converting a nuanced decision into a mechanical cleanup.

The goal is not to approve all consolidation. A valid outcome is to reaffirm a rejection with current evidence.
When a policy change or remediation is approved, Parallax records the decision and creates or names exactly one
implementation owner; it never serves as permission to edit the whole repository opportunistically.

## Allowed decision outcomes

Use exactly one outcome per scenario:

- `KEEP REJECTED` — current behavior and ownership remain correct; record fresh evidence and the next reopen
  trigger, if any.
- `REOPEN INVESTIGATION` — evidence is incomplete or contradictory; define a bounded spike with questions and
  no production behavior change.
- `ACCEPT POLICY CHANGE` — the maintainer intentionally changes a product or architecture rule; update the
  authoritative domain spec and create an implementation plan if code must change.
- `ROUTE REMEDIATION` — a concrete defect or worthwhile consolidation is now proven; name one WIP owner,
  priority, scope, tests, and dependencies.
- `SUPERSEDED` — another completed decision or authoritative spec fully resolves the scenario; link it and
  explain why no Parallax work remains.

For every decision, record the date, reviewed commit, reviewer, outcome, evidence, rationale, resulting owner,
authoritative spec impact, and explicit reopen trigger. A blank owner is allowed only for `KEEP REJECTED` or
`SUPERSEDED`.

## Decision queue

| ID | Scenario | Lighthouse disposition | Parallax status | Primary decision |
|----|----------|------------------------|-----------------|------------------|
| PAR-1 | Search-service named-pipe impersonation | Rejected as stale | PENDING REVIEW | Is existing per-request impersonation complete for every authorization path? |
| PAR-2 | ViewerWeb external navigation | Intentional, guarded product behavior | PENDING REVIEW | Retain opt-in system-browser launch or adopt a stricter product policy? |
| PAR-3 | `DestroyWindow` on `wil::unique_hwnd` owners | Not automatically a double-destroy; cleanup still owned | PENDING REVIEW | Keep scoped CHK-1 cleanup or strengthen the repository rule/tooling? |
| PAR-4 | Exception, CRT-formatting, and posted-payload hygiene | No general defect found | PENDING REVIEW | Is periodic scanning sufficient or should CI enforce the bans? |
| PAR-5 | Committed production credentials | None found | PENDING REVIEW | Is current review sufficient or should automated secret scanning be adopted? |
| PAR-6 | Stable visual hash algorithms | Algorithms are not interchangeable | PENDING REVIEW | Keep versioned variants or approve an output migration? |
| PAR-7 | Windows absolute-path predicates | One universal predicate rejected | PENDING REVIEW | Approve a policy taxonomy and staged caller migration? |
| PAR-8 | Strict versus replacement UTF conversion | One universal converter rejected | PENDING REVIEW | Standardize named policies while preserving caller intent? |
| PAR-9 | Contrast thresholds | Shared threshold rejected | PENDING REVIEW | Share mathematics only or adopt product-wide accessibility policy? |
| PAR-10 | Generic plugin threadpool scheduler | Broad scheduler rejected | PENDING REVIEW | Keep narrow lifetime envelopes or revisit after stronger commonality evidence? |
| PAR-11 | Common viewer base/rendering framework | Broad framework rejected | PENDING REVIEW | Continue pure-helper extraction or approve a design spike? |
| PAR-12 | JSON accessor coercion | Universal coercion rejected | PENDING REVIEW | Standardize explicit strict/flexible policies and error reporting? |

## Detailed decision scenarios

### PAR-1 — Search-service named-pipe impersonation

**Scenario:** A client asks the search broker to authorize or access a filesystem location over a named pipe.
Without impersonation, the broker could perform the access check as its own process identity rather than as the
requesting client, turning the service into a confused deputy.

**Why Lighthouse rejected the claim:** the current broker calls `ImpersonateNamedPipeClient` and installs a
WIL scope guard that calls `RevertToSelf` in the authorization paths at
`Common/SearchServiceBroker.cpp:2443-2448`, `:2649-2654`, and `:2666-2671`. Therefore “the search service lacks
client impersonation” is stale as a repository-wide statement.

**Decision to take later:**

- Keep rejected if every client-controlled root and transient-directory authorization reaches one of these
  guarded paths and no impersonated identity escapes the synchronous authorization scope.
- Reopen investigation if a new pipe command performs filesystem access before authorization, authorization
  moves across an async/thread boundary, or a failure path can skip reversion.
- Route remediation only with a named bypass or lifetime defect and a focused client/server authorization test.

**Evidence required:** enumerate pipe commands to authorization functions; test allowed/denied clients; verify
reversion on success and every early failure; review new pipe endpoints added since the planned commit.

### PAR-2 — ViewerWeb external navigation policy

**Scenario:** displayed HTML, Markdown, or generated content contains a user-activated HTTP(S) link. The product
must decide whether ViewerWeb cancels it, prompts, or opens it in the system browser. This is a product/security
boundary, not automatically a vulnerability.

**Why Lighthouse did not record a defect:** `Specs/Plugins/Plugins_ViewerWeb.md:49-58` defines an opt-in
`allowExternalNavigation` setting, default `false`. Only explicit user-initiated HTTP(S) requests may launch;
the in-view/new-window request must first be canceled or handled, otherwise the controller closes and no
external duplicate launches. External content never replaces the private viewer document. This is distinct
from the already-fixed inline-script injection issue.

**Decision to take later:**

- Retain the current opt-in system-browser behavior.
- Require a per-navigation confirmation showing the destination origin.
- Replace the boolean with an allowlist/managed policy.
- Disable external navigation entirely for selected viewer modes or trust classes.

**Evidence required:** threat model local versus provider-supplied documents; enterprise/managed-setting needs;
usability frequency; default and migration behavior; tests for user gesture, redirects, frames, new windows,
cancel-before-launch, and failure to suppress the WebView request.

### PAR-3 — Explicit `DestroyWindow` on a `wil::unique_hwnd` owner

**Scenario:** code calls `DestroyWindow(_hWnd.get())` while a `wil::unique_hwnd` still appears to own the same
window. Synchronous `WM_NCDESTROY` handling may reset or release the wrapper, but correctness depends on the
window procedure, reentrancy, and every exit path. A later refactor can turn a currently safe instance into a
double-destroy or stale-owner defect.

**Why Lighthouse rejected the broad defect claim:** the remaining instances are not proven double-destroys
solely by their syntax. However, the repository rule deliberately bans the pattern, and
`Specs/Plans/WIP/Win32_Inventory.md:36` already owns CHK-1. Current anchors include
`RedSalamander/RedSalamander.cpp:822` and `:1214`.

**Decision to take later:**

- Keep the work scoped to CHK-1 and close Parallax as `SUPERSEDED` when that owner proves every instance.
- Strengthen CI/static checks so the banned owner pattern cannot return.
- Permit a narrowly documented exception only when ownership is explicitly transferred before destruction.

**Evidence required:** trace `WM_NCDESTROY`, wrapper reset/release, nested close, failed creation, and reentrant
owner callbacks for each instance. A syntax-only count is discovery evidence, not proof of a crash.

### PAR-4 — Exception, CRT-formatting, and posted-payload hygiene

**Scenario:** repository rules ban `catch (...)`, non-PoC `sprintf_s`/`swprintf_s`, and raw cross-thread
`PostMessageW` ownership transfer through `new` or `.release()`. A broad violation would create swallowed
failures, localization/formatting drift, or leak/use-after-free risks during window teardown.

**Why Lighthouse rejected a general defect:** the production sweep found no matching broad violation. That is
a point-in-time result, not a permanent guarantee.

**Decision to take later:**

- Keep periodic audit-only checks if review and existing tests prevent recurrence.
- Add a source-contract CI check for the unambiguously banned forms.
- Route a focused remediation if a live match is found; do not create a repository-wide cleanup from comments,
  PoCs, generated resources, or approved named exception boundaries.

**Evidence required:** rerun scoped production searches; classify each match manually; verify posted-payload
window initialization and `WM_NCDESTROY` drain in touched hosts; record allowlists in the authoritative rule.

### PAR-5 — Committed production credentials

**Scenario:** a token, password, private key, or other production credential is committed in source, settings,
test data, build files, or history. Removal alone is insufficient because a committed credential must be treated
as exposed and rotated.

**Why Lighthouse recorded no defect:** its defensive filename/pattern sweep found no committed production
credential. No secret value was copied into the audit or this operation.

**Decision to take later:**

- Keep the existing review practice if repository exposure and contributor volume remain low.
- Adopt an approved pre-commit/CI secret scanner with a reviewed false-positive allowlist.
- Route incident remediation only by credential type and location: revoke/rotate, remove from active files,
  assess history exposure, and move configuration to the approved secure mechanism. Never paste the value into
  a plan, issue, log, or test artifact.

**Evidence required:** scanner/tooling cost, false-positive rate, public/private repository exposure, history
coverage, and the repository's credential provisioning model.

### PAR-6 — Versioned stable visual hashing

**Scenario:** several viewers derive deterministic colors from labels or identifiers. Changing the hash changes
visible colors and may invalidate snapshots or user expectations even when the new implementation looks cleaner.

**Why Lighthouse rejected interchangeability:** the byte hash used by `Common/DxUi/DxUi.Theme.cpp` and several
viewers is not bit-compatible with `RedSalamander/AppTheme.cpp`, which hashes both bytes of each UTF-16 code
unit. Identical function names therefore do not establish identical output contracts.

**Decision to take later:**

- Keep both algorithms with explicit versioned/policy names and share only bit-identical implementations.
- Approve a visible-output migration to one algorithm, with before/after examples and release-note impact.
- Introduce a new version only for newly created state while preserving legacy output where stability matters.

**Evidence required:** golden ASCII and non-ASCII hashes, every persisted/snapshotted consumer, user-visible
color impact, compatibility requirements, and a migration/rollback strategy.

### PAR-7 — Windows path classification taxonomy

**Scenario:** callers ask whether a path is drive-qualified, drive-relative, rooted, UNC, extended, device, or
fully absolute. Treating all these questions as one `IsAbsolutePath` operation can weaken validation or reject
valid navigation, especially for `C:`, `\foo`, UNC, and extended/device forms.

**Why Lighthouse rejected one universal predicate:** current classifiers in `NavigationLocation.h`,
`FolderView.Enumeration.cpp`, `NavigationViewInternal.h`, and `FolderWindow.FileSystem.cpp` encode different
questions. In particular, callers do not agree that drive-qualified `C:` is fully absolute.

**Decision to take later:**

- Approve a shared taxonomy with separately named predicates and migrate callers only after assigning intent.
- Keep domain-local classifiers where navigation, filesystem, or security policy intentionally differs.
- Route a security remediation immediately if a classifier can bypass traversal, ADS, leaf-name, or destination
  containment rules; do not wait for architectural consolidation.

**Evidence required:** caller matrix and expected results for drive-relative/rooted, UNC, extended/device, slash
variants, dot segments, reserved names, trailing spaces/dots, long paths, and normalization boundaries.

### PAR-8 — Strict and replacement UTF conversion

**Scenario:** malformed UTF enters settings, provider protocols, filenames, or display text. Strict parsers must
reject malformed data, while selected display/convenience paths may intentionally use replacement characters so
the UI remains usable.

**Why Lighthouse rejected one converter:** `Common/Helpers.h` exposes replacement-oriented conversion while
SettingsStore, ThemeDefinitionIo, Curl, VLC, and other parsers retain strict variants. Combining them behind an
ambiguous name would silently change validation behavior.

**Decision to take later:**

- Standardize explicit `Strict` and `ReplaceInvalid` APIs and migrate only policy-matched callers.
- Keep provider-local adapters when dependency or protocol error reporting differs.
- Approve a policy change only with a caller-by-caller malformed-input decision and migration tests.

**Evidence required:** invalid/truncated sequences, embedded NUL, supplementary characters, byte limits,
provider error contracts, display fallback behavior, and proof that strict parsers never silently replace.

### PAR-9 — Shared color mathematics versus one contrast threshold

**Scenario:** theme expressions, application chrome, and viewers choose foreground/background colors using
luminance and contrast calculations. The mathematical conversion may be common, but the threshold represents a
product and accessibility policy for a particular surface.

**Why Lighthouse rejected a universal threshold:** luminance math recurs in ThemeExpression, RedConfigure, and
ViewerText, while application/viewer contrast heuristics use different thresholds. The differences are not
proven defects.

**Decision to take later:**

- Share only color-space/luminance mathematics and retain named surface policies.
- Adopt a product-wide accessibility target, then update authoritative theme/viewer specs and visual tests.
- Route a focused fix if a surface fails an agreed contrast requirement; do not change unrelated generated
  palette output in the same patch.

**Evidence required:** selected accessibility target, high-contrast behavior, representative themes, threshold
boundary tests, generated-color snapshots, and design approval for visible changes.

### PAR-10 — Generic plugin threadpool scheduler

**Scenario:** many plugins submit asynchronous work and therefore appear to duplicate scheduling. Their actual
contracts differ in cancellation, latest-wins behavior, pending capacity, module pins, COM apartment, terminal
delivery, resource budgets, and unload quiet points.

**Why Lighthouse rejected the generic scheduler:** the Farsight closeout at
`Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md:359-365`
records that a generic `std::function` dispatcher would erase safety distinctions. Lighthouse therefore permits
only narrow callback-state, callback-return module-pin, and identical fail-fast MTA entry envelopes.

**Decision to take later:**

- Keep narrow lifetime primitives and domain-specific schedulers.
- Reopen a design spike only after at least three implementations share the same complete state machine—not
  merely the same submission API.
- Accept a generic scheduler only with typed cancellation/result semantics, explicit module lifetime, resource
  budgets, and per-plugin teardown proofs.

**Evidence required:** state-machine comparison, unload/cancel stress, submission failure, callback-return pin
transfer, COM initialization, terminal delivery, pending-work bounds, and proof no domain guarantee disappears.

### PAR-11 — Common viewer base class or rendering framework

**Scenario:** viewers repeat title-bar theming, file-combo popup geometry, clipboard operations, and some window
plumbing. A broad shared base could reduce lines but also couple DLLs and erase different cancellation,
rendering, identity, and unload contracts.

**Why Lighthouse rejected the broad framework:** Farsight retired the shared window/surface stage as
maintainer-gated design-spike work. `Specs/Plugins/Plugins_ViewerPlugins.md:341` explicitly keeps viewer async
dispatch, window thunks/surfaces, and navigation plugin-specific where contracts differ. Lighthouse instead
routes pure chrome, popup, clipboard, and geometry helpers, plus the separately proven FunctionBar/StatusBar
render-target lifecycle.

**Decision to take later:**

- Continue extracting dependency-light pure helpers with no shared mutable window state.
- Reopen a viewer-host design spike if several viewers require the same behavioral change in lockstep and the
  abstraction provides independent value beyond line reduction.
- Reject a base class that changes plugin ABI, owns plugin-specific state, or imposes one render/unload model.

**Evidence required:** caller/state-machine matrix, DLL dependency impact, ABI boundary, testability, per-viewer
exceptions, code removed versus framework code added, and teardown/device-loss/mixed-DPI tests.

### PAR-12 — Strict and flexible JSON access policies

**Scenario:** settings and cloud/provider protocols read missing, null, boolean, signed/unsigned, numeric-string,
and wrong-type JSON values. Some protocols legitimately encode a number as a string; other schemas must reject
that coercion.

**Why Lighthouse rejected universal coercion:** `Common/FileSystemPathIdentity.cpp`, FileSystem7z,
GoogleDrive, Curl, S3, MicrosoftDrive, viewer settings, and Preferences contain accessors with different
required/optional and strict/flexible behavior. GoogleDrive, for example, has a deliberately flexible unsigned
integer reader in addition to strict typed readers.

**Decision to take later:**

- Standardize dependency-light accessors whose names/options explicitly state required/optional and
  strict/flexible behavior.
- Keep protocol-local validation and error mapping even when low-level typed access is shared.
- Approve coercion changes only against provider/schema documentation and malformed-input tests.

**Evidence required:** per-caller matrix for missing/null/wrong type/overflow/numeric string, unknown-field and
future-schema behavior, UTF-8 errors, provider compatibility, and RAII cleanup on every failure.

## Review sequence

1. Review PAR-1, PAR-3, and PAR-5 first because they touch security or ownership claims.
2. Review PAR-2 and PAR-9 with an explicit product/accessibility decision-maker.
3. Review PAR-7, PAR-8, and PAR-12 before the corresponding Lighthouse common-primitives work, so shared APIs
   encode approved policies.
4. Review PAR-6, PAR-10, and PAR-11 before any consolidation that changes visible output or lifetime models.
5. Review PAR-4 last as a tooling-policy question after current remediation branches settle.
6. For every non-rejected result, create or name one WIP owner and add it to `Specs/Plans/WIP/README.md`; never
   execute an accepted decision directly from Parallax.

## Evidence refresh commands

These commands are discovery gates. Every match still requires source review.

```powershell
rg -n "ImpersonateNamedPipeClient|RevertToSelf" Common/SearchServiceBroker.cpp
rg -n "allowExternalNavigation|put_Cancel|put_Handled" Specs/Plugins/Plugins_ViewerWeb.md Plugins/ViewerWeb
rg -n "DestroyWindow\([^\r\n]*\.get\(\)\)" Common Plugins RedConfigure RedSalamander RedSalamanderMonitor
rg -n "catch \(\.\.\.\)|sprintf_s|swprintf_s" Common Plugins RedConfigure RedSalamander RedSalamanderMonitor
rg -n "StableHash32|ColorFromHSV|RelativeLuminance|Contrast" Common Plugins RedConfigure RedSalamander
rg -n "Is.*AbsolutePath|LooksLikeWindowsAbsolutePath|Utf16FromUtf8|Utf8FromUtf16" Common Plugins RedConfigure RedSalamander
rg -n "TryGetJson|yyjson_get_(str|int|uint|bool)" Common Plugins RedConfigure RedSalamander
```

For credential review, use only an approved defensive scanner or filename/pattern audit. Findings must name the
credential type and location without reproducing its value.

## Done criteria

- [ ] PAR-1 through PAR-12 each have a dated reviewed commit and one allowed decision outcome.
- [ ] Every `KEEP REJECTED` decision contains current evidence and a reopen trigger or explicitly says no
  scheduled revisit is justified.
- [ ] Every `REOPEN INVESTIGATION`, `ACCEPT POLICY CHANGE`, or `ROUTE REMEDIATION` result names exactly one WIP
  owner with scope, priority, dependencies, verification, and authoritative-spec impact.
- [ ] Accepted product/architecture policies are merged into the relevant authoritative `Specs/<Domain>/`
  document before implementation closeout.
- [ ] Lighthouse points to Parallax as the decision owner and contains no conflicting decision status.
- [ ] The WIP index records every resulting executable owner and no duplicate queue item.
- [ ] No production source was modified as part of the Parallax decision review itself.
- [ ] This operation moves to `Specs/Plans/Done/` when all twelve decisions and resulting routes are recorded;
  implementation plans may remain WIP independently.

## STOP conditions

Stop and report instead of deciding or implementing when:

- Live evidence contradicts the scenario and the discrepancy cannot be resolved by reading current code and
  authoritative specs.
- A security-sensitive finding would require recording a credential value, operational exploit sequence, or
  other sensitive material. Record only defensive evidence and location.
- The proposed decision conflicts with an active WIP owner, completed authoritative decision, plugin ABI, or
  documented product contract without an explicit maintainer policy change.
- A consolidation changes malformed-input handling, visible hash/color output, path acceptance, JSON coercion,
  cancellation, COM apartment, module lifetime, queue ordering, or teardown behavior without characterization
  tests and an approved migration.
- Review reveals a concrete active vulnerability, data-loss route, or lifetime defect. Create a narrowly scoped
  high-priority remediation owner rather than leaving it as a Parallax discussion item.
- The reviewer is about to edit production code from this operation. Record the decision and route the work
  instead.

## Maintenance notes

- Keep Lighthouse's original rejection text as historical audit evidence; update decision status only here.
- When a scenario is superseded, link the authoritative spec or completed operation rather than deleting the
  rationale. This prevents the same broad claim from being re-audited without new evidence.
- Re-run only the affected decision after material platform, threat-model, plugin ABI, viewer policy, or shared-
  primitive changes; a full twelve-item review is not required for every code change.
- Reviewers should favor explicit policy names and narrow primitives over abstractions justified solely by a
  repository-wide match count.
