# Last-Four-Days Independent Code Review and Remediation Routing

| Field | Value |
| --- | --- |
| Status | DONE |
| Reviewed | 2026-07-16 |
| Git range | `eaf640798eb0fa21f2b5b6e3883891ac2442a255..f4e0c8c3bed86ed7ae8b1505b72d338b59e1ffd6` |
| Calendar window | 2026-07-12 through 2026-07-16 |
| Working tree | Clean at review start and closeout |
| Scope | All executable/configuration changes in the range, directly affected callers, tests, build/CI, specifications, and committed validation evidence |
| Remediation progress | All 18 findings are implemented, closed by their authoritative completed owner, or explicitly routed to Astrolabe/Rosetta Lantern |
| Merged implementation | `177600a8d` on `master` (2026-07-17) |
| Primary rule | Fix newly introduced durability/data-integrity defects first; route inherited defects to their existing owner instead of creating competing implementations |

## Executive verdict

The range contains substantial good work: the cross-filesystem bridge closes real trust-boundary and
integrity holes, viewer plugins now have much stronger async/lifetime contracts, settings recovery is
section-scoped and forward-safe, several duplicated policies moved behind shared helpers, and the current-head
RedConfigure/Common/DxUi and SettingsSchema slices build and test cleanly.

At review time this was not a safe closeout. The review confirmed **18 actionable findings**:

- **1 P0 verification blocker**;
- **4 P1 durability/data-integrity defects**;
- **10 P2 correctness, accessibility, localization, and architecture defects**;
- **3 P3 robustness/DX simplifications**.

The most serious newly introduced defect is the Windows session-end writer bypassing the serialized settings
coordinator. A delayed older snapshot can overwrite the newer session-end snapshot, and simultaneous writers
share the same fixed `.tmp` path. RedConfigure has two separate stale-preview commit problems and can classify
unreadable localization inputs as warnings, allowing a second-click export to replace a target satellite with
English fallback text. The shared posted-payload teardown contract also remains unsafe under recycled HWNDs:
it frees payloads but leaves their messages queued.

The targeted repair and routing described below is complete. Causeway/Farsight/Firebreak/Lighthouse changes were
retained, behavioral tests now cover the failing boundaries, and every finding has one explicit disposition.

## Scope and method

- Confirmed a clean range from `eaf640798` to `f4e0c8c`: 43 dated commits and 797 merge-inclusive changed
  paths. The executable/configuration review surface was 399 files with approximately 59,896 additions and
  22,401 deletions; historical test-run archive churn was excluded from source-quality line review.
- Reviewed changed production code under `Common`, `RedConfigure`, `RedLauncher`, `RedSalamander`,
  `RedSalamanderMonitor`, `RedSalamanderSearchService`, and every changed plugin family. Reviewed the changed
  test/tool/build/CI surfaces and directly affected callers.
- Used commit-level diffs and blame to distinguish **INTRODUCED** behavior from **PRE-EXISTING IN TOUCHED
  CODE**. A finding is not attributed to this range merely because a changed helper made it easier to see.
- Reconciled against the completed Farsight, Causeway, Firebreak, expressive-theme, and Lighthouse ledgers and
  the current Observatory WIP plan. Existing owners remain authoritative where named below.
- One independent settings/configuration reviewer completed before subagent credits were exhausted. Its
  candidates were treated as untrusted leads and re-read directly; seven other planned independent passes
  could not run because the Copilot subagent credit limit was reached. The remaining review was performed
  directly, including explicit rejection of unsupported candidates.

Generated/minified assets, vendored dependencies, and historical TestRuns payload bodies were not audited line
by line. Their manifests, ownership, closeout claims, and build/test effects were reviewed. Live cloud, MTP,
SMB, and installed-VLC endpoints were unavailable; provider findings therefore rely on deterministic fake
transport tests and source contracts. Linguistic translation quality was not evaluated beyond proving that all
77 new RedConfigure strings are byte-identical to English in each affected non-English satellite.

## Validation baseline

Current-head validation performed during this review:

- `./build.ps1 -Configuration Debug -Platform x64 -ProjectName RedConfigure`: **PASS**, 0 warnings / 0
  errors, log `.build/logs/msbuild-20260716_221005_200.log`.
- `./build.ps1 -Configuration Debug -Platform x64 -ProjectName RedConfigureTests`: **PASS**, 0 warnings / 0
  errors, log `.build/logs/msbuild-20260716_221137_139.log`.
- `.build/x64/Debug/RedConfigureTests.exe`: **PASS**. Repository-sized scenario: 6 owners, 1,500 rows,
  10 themes, 191 ms scan, 3 ms validation, 264 us preview.
- `./build.ps1 -Configuration Debug -Platform x64 -ProjectName SettingsSchemaTests`: **PASS**, 0 warnings /
  0 errors, log `.build/logs/msbuild-20260716_221249_703.log`.
- `.build/x64/Debug/SettingsSchemaTests.exe`: **PASS**, including the shared plugin-configuration model,
  coercion, unknown-member preservation, and real-schema cases.
- The complete solution build was attempted after dependencies were warmed. It compiled/link-produced many
  changed targets (including plugin and standalone test binaries) but the execution host canceled it while
  the main application was still compiling. No C++ compiler diagnostic was emitted; the canceled attempt is
  **not** a green build and **not** evidence of a source failure.
- Pester could not be run in this environment because Pester 5 is not installed. No package/module install was
  performed during this read-only audit.
- The committed Lighthouse closeout records a clean current-head x64 Debug rebuild at
  `.build/logs/msbuild-20260716_150845_700.log`. Its Full run reported 1,151 passed / 3 failed / 54 skipped;
  inventory bookkeeping was repaired afterward, while Monitor ETW and ViewerSpace remained accepted external
  blockers. That evidence does not make the canonical Full gate green.

## Remediation progress (2026-07-17)

- **F4-SET-01 CLOSED:** `2888e34a3` routes confirmed Windows session-end persistence through the existing
  serialized settings coordinator, admits one bounded settings-only final snapshot, fences later submissions,
  and preserves schema/modal teardown boundaries. The pre-existing `Common::Files::LocalFileTransaction`
  supplies unique attempt-owned sibling staging, so no second settings writer or fixed `.tmp` path was added.
- The focused real-coordinator race case `settings_save_queue_serializes_coalesces_and_flushes` passed 1/1 in
  local archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-17_171010`; the final session-end write completed
  successfully in 1,493,760 us and the test covers an older active write, an older queued write, final-snapshot
  dominance, schema preservation, and the later-submission fence. The durable contract is in
  `Specs/Core/Core_StartupBootstrap.md` and `Specs/Core/Core_SettingsStore.md`.
- **F4-MSG-01 CLOSED by this operation:** `lParam` is now an opaque process token rather than a payload
  address. Teardown invalidates tokens before deleting storage, while stale queued tokens return null even after
  simulated HWND reuse. The RED run at
  `Specs/TestRuns/4cb089111a23/DxUi/2026-07-17_1747_posted_payload_red` failed on the stale queue boundary. The final
  WindowHost candidate at
  `Specs/TestRuns/4cb089111a23/DxUi/2026-07-17_1853_posted_payload_token_candidate` passed and measured 4,395 us to
  post 1,024 payloads and 1,280 us to invalidate/drain their registry storage; all 1,024 stale messages were pumped
  and rejected. Debug and ASan WindowHost, Compare/Find focused cases, viewer close/unload cases, the all-consumer
  Debug build, and the posted-payload source contracts are green.
- **F4-RC-01 through F4-RC-06 CLOSED by reconciliation:** Observatory Tracks 16/19 already landed typed fail-closed
  validation, complete preview identity and before-state checks, parsed theme transforms, transactional one-step
  Undo, and typed/localized duplicate results in `ce4e6c4ca`. The focused RedConfigure proof and authoritative
  UI/testing contracts are recorded by that Done owner; this plan does not duplicate them.
- **F4-THEME-01, F4-THEME-02, and F4-SET-02 CLOSED by this operation:** formatter output uses the schema's
  camelCase function spellings, `ensureContrast` measures the rendered translucent foreground over a required
  opaque background, and duplicate/unusable inline theme JSON retains its authored value and array position. The
  focused Debug RedConfigure build has 0 warnings/errors and the candidate run at
  `Specs/TestRuns/4cb089111a23/RedConfigure/2026-07-17_1850_four_day_review_theme_closeout` passed with 204 ms scan,
  3 ms validation, 410 us single preview, and 46,460 us for the 512-token mass preview, all within contract.
- **F4-LOC-01 ROUTED:** linguistic correctness for the 77 reviewed Czech/Japanese/Slovak strings is owned only by
  `Operation_RosettaLantern_RedConfigureSatelliteLocalizationReview_2026-07-17.md`. It requires target-language
  human review and must not be falsely closed by mechanically copying or generating satellite text.
- **F4-GATE-01, F4-SET-03, F4-TEST-02, F4-DX-01, and F4-REPO-01 CLOSED by Observatory; F4-TEST-01 ROUTED to
  Astrolabe:** their completed or active owner evidence is authoritative. Track 0 used the user's bounded-convergence
  decision: CI and the sandbox audit are green, one Full pass classified every unrelated failure, and no redundant
  broad-suite rerun is required for these focused delta fixes.

## Closeout validation (2026-07-17)

- Final all-consumer x64 Debug build: **PASS**, 0 warnings / 0 errors,
  `.build/logs/msbuild-20260717_185645_246.log`.
- Focused Debug RedConfigure build: **PASS**, 0 warnings / 0 errors,
  `.build/logs/msbuild-20260717_184603_706.log`; `RedConfigureTests.exe`: **PASS** with the formatter/schema,
  alpha-composited contrast, positioned duplicate-theme round trip, locale, Undo, and existing repository-sized
  cases.
- Focused Debug SettingsSchema build: **PASS**, 0 warnings / 0 errors,
  `.build/logs/msbuild-20260717_185611_162.log`; `SettingsSchemaTests.exe`: **PASS**.
- Debug WindowHost: **PASS** from rebuilt binaries with 1,024/1,024 stale tokens rejected; ASan Debug WindowHost:
  **PASS** with no sanitizer finding. Builds have 0 warnings / 0 errors at
  `.build/logs/msbuild-20260717_185152_548.log` and `.build/logs/msbuild-20260717_185353_604.log`.
- ViewerSqlite's intentional zero-token fallback versus nonzero-stale-token boundary was audited explicitly;
  focused plugin build **PASS**, 0 warnings / 0 errors at `.build/logs/msbuild-20260717_190251_760.log`, and the
  complete `ViewerSqliteTests.exe` run passed.
- Documentation-drift and posted-payload source contracts: **19 passed / 0 failed** under the installed Pester
  3.4 host. Earlier Compare/Find and viewer teardown/close focused runs remained green; the final public helper API
  did not change after those runs.
- Broad gate ownership remains Observatory Track 0's bounded-convergence evidence: CI and its TestSandbox audit
  are green; one Full pass classified every unrelated failure under explicit owners. Per the user's convergence
  decision, this focused delta operation did not run redundant CI/Full cycles.

## Finding index

| ID | Priority | Status | Finding | Effort | Fix risk | Confidence | Owner |
| --- | --- | --- | --- | --- | --- | --- | --- |
| F4-GATE-01 | P0 | CLOSED by Observatory Track 0 bounded convergence | Canonical Full remains red while the operation is marked DONE | M | LOW | HIGH | Observatory Track 0 |
| F4-SET-01 | P1 | CLOSED (`2888e34a3`) | Session-end persistence bypasses serialization and can lose the newest settings | M | MED | HIGH | This plan Track 1; coordinated with completed Observatory Tracks 7/14 |
| F4-RC-01 | P1 | CLOSED by Observatory Track 19 (`ce4e6c4ca`) | Localization validation downgrades read/parse failures and can export fallback over translations | M | MED | HIGH | Observatory Track 19 |
| F4-MSG-01 | P1 | CLOSED by this operation | Payload teardown frees queued `lParam` storage without removing the message | M | HIGH | HIGH | This plan Track 3 |
| F4-RC-02 | P1 | CLOSED by Observatory Track 19 (`ce4e6c4ca`) | Localization batch approval can apply stale text to the wrong/currently changed row | M | MED | HIGH | Observatory Track 19 |
| F4-RC-03 | P2 | CLOSED by Observatory Track 19 (`ce4e6c4ca`) | Theme batch approval ignores changed arguments and authored model state | M | MED | HIGH | Observatory Track 19 |
| F4-RC-04 | P2 | CLOSED by Observatory Track 19 (`ce4e6c4ca`) | Theme recipes mutate sources outside their advertised semantics | M | MED | HIGH | Observatory Track 19 |
| F4-RC-05 | P2 | CLOSED by Observatory Tracks 16/19 (`ce4e6c4ca`) | Workbench theme commands bypass session undo/redo ownership | M | MED | HIGH | Observatory Tracks 16/19 |
| F4-THEME-01 | P2 | CLOSED by this operation | Theme formatter emits spellings rejected by the published schema | S | LOW | HIGH | This plan Track 5 |
| F4-THEME-02 | P2 | CLOSED by this operation | `ensureContrast` certifies translucent colors using RGB-only contrast | M | MED | HIGH | This plan Track 5 |
| F4-RC-06 | P2 | CLOSED by Observatory Track 19 (`ce4e6c4ca`) | Duplicate Theme treats every validation failure as an ID collision and fails silently | S-M | LOW | HIGH | Observatory Track 19 |
| F4-LOC-01 | P2 | ROUTED to Operation Rosetta Lantern | All 77 new RedConfigure strings remain English in every affected satellite | L | LOW | HIGH | Rosetta Lantern |
| F4-SET-02 | P2 | CLOSED by this operation | Duplicate inline themes are discarded on the next canonical save | S-M | LOW | HIGH | This plan Track 5 |
| F4-SET-03 | P2 | CLOSED by Observatory Track 7 (`ce4e6c4ca`) | Shared uint32 settings reads wrap values above `UINT32_MAX` | S-M | LOW | HIGH | Observatory Track 7 |
| F4-TEST-01 | P2 | ROUTED to Astrolabe | Regex source-contract tests are a brittle shadow compiler | L | MED | HIGH | Astrolabe |
| F4-TEST-02 | P3 | CLOSED by Observatory Track 0 (`ce4e6c4ca`) | TestSandbox contamination is measured but non-blocking | M | LOW | HIGH | Observatory Track 0 |
| F4-DX-01 | P3 | CLOSED by Observatory Track 19 (`ce4e6c4ca`) | Theme numeric grammar depends on mutable process locale | S-M | LOW | MED | Observatory Track 19 |
| F4-REPO-01 | P3 | CLOSED by Observatory Track 18 (`ce4e6c4ca`) | Theme mockup commits generated `dist` without an artifact contract | S | LOW | HIGH | Observatory Track 18 |

## Detailed findings

### F4-GATE-01 - Canonical Full remains red while the operation is marked DONE

- **Evidence:** `Specs/Plans/Done/Operation_Lighthouse_WholeRepositoryAuditFindingsAndRemediationRouting_2026-07-10.md`
  records the current-head Full run `20260716T131155Z-7948-bdf657ae03db41d0b7f87ad6a08317cd` as
  1,151 passed / 3 failed / 54 skipped. Inventory bookkeeping was changed after that run; Monitor ETW and a
  ViewerSpace accounting race remain accepted blockers rather than passing cases.
- **Evidence:** `AGENTS.md` defines `Tools/Run-AllTests.ps1 -Suite Full` as the canonical green closeout gate.
- **Impact:** The repository has no all-green current-head executable baseline. A subsequent regression can be
  misclassified as one of the accepted failures, and “DONE” no longer means the canonical command exits zero.
- **Effort:** M.
- **Risk:** LOW; this is test classification/repair, not product behavior.
- **Confidence:** HIGH.
- **Routing:** Do not create a competing implementation here. Execute Observatory Track 0, rerun Full after
  inventory repair, and close or explicitly quarantine each remaining failure under the repository's blocking
  quarantine policy. A maintainer acceptance note is not a green gate.

### F4-SET-01 - Session-end persistence bypasses serialization and can lose the newest settings

- **Evidence:** `RedSalamander/RedSalamander.cpp:2594-2624` calls
  `Common::Settings::SaveSettings(...)` directly from `WriteSessionEndSettings`.
- **Evidence:** normal synchronous, asynchronous, and process-shutdown saves are serialized by
  `SerializedSettingsSaveCoordinator` in `RedSalamander/SettingsHotReload.cpp:390-725`.
- **Evidence:** `Common/Common/SettingsStore.cpp:480-555` uses one fixed `<settings>.tmp` path and replaces the
  destination unconditionally.
- **Evidence:** `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp:121-143` injects a replacement
  session-end writer, so it proves snapshot/idempotence behavior but never races the real coordinator.
- **Impact:** An older debounced snapshot queued before `WM_ENDSESSION` can run after the direct session-end
  write and replace it. If a coordinator write is already active, both writers contend for the same `.tmp`;
  the session-end save can fail or consume/replace the other writer's staging file. Pane state, history,
  layout, and menu state captured for logoff/restart can be lost.
- **Effort:** M.
- **Risk:** MED; shutdown ordering and bounded Windows session-end behavior are sensitive.
- **Confidence:** HIGH.
- **Fix sketch:** Add a settings-only, bounded final-save request to the serialized coordinator. Under one
  submission transition, supersede older queued snapshots for the same app, reject later normal submissions,
  and wait only to the session-end deadline. Give every low-level attempt a unique sibling temp path as defense
  in depth. Keep schema generation and modal/plugin teardown out of `WM_ENDSESSION`.
- **Required tests:** Delay an older queued save, send confirmed `WM_ENDSESSION`, release the delay, and prove
  the disk contains the session-end snapshot; repeat with a write already active; prove no `.tmp` collision,
  one final generation, future-schema save blocking, and no modal teardown.

### F4-RC-01 - Localization validation can export fallback over translations

- **Evidence:** `RedConfigure/RedConfigureWorkflow.cpp:294-301` decides whether a workspace diagnostic is an
  error by searching English text for only `Parse localization` or `Duplicate resource id`; everything else
  is a warning.
- **Evidence:** `RedConfigure/RedConfigureSession.cpp:458-489` emits read/parse failures as strings such as
  `Read localization target failed ...` and `Parse localization target failed ...`; workspace discovery and
  theme loading may emit only a path.
- **Evidence:** `RedConfigure/RedConfigureRoot.cpp:1883-1903` permits export after a second confirmation when
  there are warnings but no classified errors.
- **Evidence:** missing/unreadable target cells receive English source fallback at
  `RedConfigure/RedConfigureSession.cpp:572-573`, and
  `BuildLocalizationReviewExportPreviews` at `:1446-1490` writes a complete owner/culture table whenever one
  cell is dirty.
- **Impact:** If an existing satellite cannot be read or parsed, editing one cell and confirming warnings can
  emit a complete replacement in which previously translated cells are English fallback. Exporting to the
  satellite destination destroys translations. A malformed embedded source can also be treated as warning-only
  even though the review inventory is incomplete.
- **Effort:** M.
- **Risk:** MED; validation policy affects export availability.
- **Confidence:** HIGH.
- **Fix sketch:** Replace presentation strings with structured diagnostic code, severity, owner, culture, path,
  and HRESULT. Read/parse failures for an embedded source or an existing target are export-blocking for that
  owner/culture. Preserve the original target bytes or refuse to generate its replacement until it parses.
- **Required tests:** unreadable target, malformed target, unreadable source, malformed source, and target-only
  warning. Dirty one unrelated cell and prove the four hard failures write nothing while the target-only warning
  follows the documented confirmation path.

### F4-MSG-01 - Payload teardown frees storage while its message remains queued

- **Evidence:** `Common/Helpers.h:2510-2555` removes tracked entries and deletes their payloads during
  `DrainPostedPayloadsForWindow`, but does not remove the corresponding posted messages from the thread queue.
- **Evidence:** `Common/Helpers.h:2428-2437` erases the closed-HWND tombstone when a numeric HWND value is
  reused, and `TakeMessagePayload` at `:2614-2618` adopts any raw `lParam` even when it is no longer registered.
- **Evidence:** `Tests/DxUiTests/DxUiTests.WindowHost.cpp:940-961` verifies 64 payload destructors run during
  teardown, but never pumps the stale queue entries or forces HWND reuse after deletion.
- **Impact:** If Windows reuses the HWND before a stale queued message is removed, the new window can receive a
  freed pointer. `TakeMessagePayload` then dereferences/double-deletes it. The result is use-after-free, heap
  corruption, or a crash on any of the many async surfaces migrated to this helper.
- **Effort:** M.
- **Risk:** HIGH; this is a high-fan-in ownership primitive and queue ordering must remain correct.
- **Confidence:** HIGH from ownership analysis; the exact HWND reuse timing is nondeterministic in production.
- **Fix sketch:** Do not delete a tracked payload while its message remains queued. During teardown, remove and
  match queued payload messages for the dying HWND before deletion, or redesign registry takeover so a receiver
  adopts only a still-registered pointer tied to the current window generation. Unknown/stale pointers must
  return null and must never be wrapped blindly.
- **Required tests:** queue 64 payloads, destroy, churn windows until HWND reuse, pump every stale message, and
  prove exactly-once destruction/no dispatch. Include unrelated thread messages to prove ordering is not
  disturbed and run under ASan.

### F4-RC-02 - Localization batch approval can target stale or wrong rows

- **Evidence:** `LocalizationBatchRequest` and `LocalizationBatchChange` in
  `RedConfigure/RedConfigureWorkflow.h:75-98` carry source/target cultures, find/replace text, row index,
  owner/resource identity, and before/after text.
- **Evidence:** `RedConfigure/RedConfigureRoot.cpp:1998-2009` approves a pending preview by comparing only batch
  kind and target culture.
- **Evidence:** `RedConfigure/RedConfigureSession.cpp:1220-1245` applies by mutable row index and culture, without
  verifying owner/resource identity or the captured `before` text; missing rows are skipped and the operation
  still returns success.
- **Impact:** Changing search text, source culture, selection, reloading/reordering rows, or editing a cell
  between preview and approval can apply old text to a different resource or overwrite a newer edit.
- **Effort:** M.
- **Risk:** MED; batch mutation and undo behavior must remain atomic.
- **Confidence:** HIGH.
- **Fix sketch:** Give rows stable owner/resource IDs and a model generation. Compare the complete request on
  approval; at apply, resolve each stable ID and require current text to equal `before`. Any mismatch rejects the
  whole batch as stale with no mutation.
- **Required tests:** changed find/replace argument, changed source culture, row reorder/reload, deleted row, and
  intervening target edit. Every case must reject atomically and preserve one-step undo for a valid batch.

### F4-RC-03 - Theme batch approval ignores arguments and model changes

- **Evidence:** `ThemeMassRequest`/`ThemeMassChange` in `RedConfigure/RedConfigureWorkflow.h:127-148` carry
  recipe, keys, argument, alpha, and before/after text.
- **Evidence:** `RedConfigure/RedConfigureRoot.cpp:2065-2074` approves by comparing only recipe and key list,
  ignoring argument and alpha.
- **Evidence:** `RedConfigure/RedConfigureWorkflow.cpp:541-560` applies captured changes without checking that
  the current authored source still equals `before`.
- **Impact:** A second click after changing the replace expression, palette name, alpha slider, or a token edit
  applies the old preview. The UI says preview/apply, but approval is not approval of the visible inputs or
  current model.
- **Effort:** M.
- **Risk:** MED.
- **Confidence:** HIGH.
- **Fix sketch:** Add a monotonic theme model revision and compare the complete request plus revision. Revalidate
  every `before` value at apply and reject the whole batch on any drift.
- **Required tests:** change argument, alpha, selected group, authored source, and palette contents between
  preview/apply; all stale approvals must be rejected without partial palette creation.

### F4-RC-04 - Theme recipe names do not match their mutation boundaries

- **Evidence:** `RedConfigure/RedConfigureWorkflow.cpp:155-159` implements Replace Reference as unrestricted
  substring replacement over the authored expression.
- **Evidence:** the same block implements Convert Solids to References by replacing every selected source with
  `ref(palette.<argument>)`, without checking that the source kind is direct/solid.
- **Impact:** Replace Reference can alter substrings inside a different token name or function argument.
  Convert Solids can overwrite existing references, dynamic expressions, and inherited intent despite its UI
  name. Preview makes the damage visible only if every row is inspected; the summary shows one representative.
- **Effort:** M.
- **Risk:** MED; parser-aware transformation must preserve canonical formatting and dependency checks.
- **Confidence:** HIGH.
- **Fix sketch:** Transform parsed `ThemeColorSource` nodes. Replace only exact reference operands; convert only
  `Direct` sources and report skipped/non-applicable keys in the preview. Reject malformed `find=replace` and
  invalid palette names before creating a preview.
- **Required tests:** prefix/suffix-colliding key names, nested functions, dynamic sources, existing references,
  direct solids, malformed arguments, and mixed applicable/non-applicable groups.

### F4-RC-05 - Theme commands bypass session undo/redo ownership

- **Evidence:** `RedConfigureSession::UpdateThemeColor` at `RedConfigure/RedConfigureSession.cpp:1141-1156`
  records an undo snapshot.
- **Evidence:** `RedConfigure/RedConfigureRoot.cpp:1661-1797` directly calls mutable model methods for palette
  create/rename, group transforms, and reset; `:2067` applies a mass recipe directly to the model.
- **Evidence:** `RedConfigure/RedConfigureSession.h:129-130` publicly exposes both const and mutable model references, making
  bypass the easiest integration path.
- **Impact:** The visible Undo button cannot undo these user-visible edits, redo history is not cleared, and
  one batch can require manual reconstruction. Session state and model state no longer have one mutation owner.
- **Effort:** M.
- **Risk:** MED; command consolidation touches all workbench theme edits.
- **Confidence:** HIGH.
- **Fix sketch:** Expose only a const preview model publicly. Add session-owned palette, transform, reset, and
  mass-apply commands that record exactly one pre-mutation snapshot, mutate atomically, and clear redo only on
  success.
- **Required tests:** each command supports one-step undo/redo; a failed batch creates no history entry; a
  multi-key transform is one undo unit.

### F4-THEME-01 - Theme formatter and schema disagree on canonical function spelling

- **Evidence:** `Common/Common/ThemeExpression.cpp:756-770` emits `perceptualtone`, `ensurecontrast`,
  `systemcolor`, `seededrainbow`, and `seededchoice` in lowercase.
- **Evidence:** `Specs/SettingsStore.schema.json:705-716` accepts only camelCase `perceptualTone`,
  `ensureContrast`, `systemColor`, `seededRainbow`, and `seededChoice`.
- **Evidence:** the runtime parser lowercases input before matching, so runtime round-trip tests pass while an
  editor/schema validator rejects the generated output.
- **Impact:** RedConfigure and SettingsStore can write theme expressions that the repository's own published
  schema marks invalid. Users see false diagnostics and schema consumers cannot trust generated files.
- **Effort:** S.
- **Risk:** LOW.
- **Confidence:** HIGH.
- **Fix sketch:** Choose the documented camelCase spellings as canonical formatter output; retain
  case-insensitive input compatibility if desired. Add a generated-expression-to-schema parity test covering
  every function.

### F4-THEME-02 - `ensureContrast` ignores alpha when certifying WCAG contrast

- **Evidence:** `Common/Helpers.h:224-229` computes ARGB luminance from RGB channels only.
- **Evidence:** `Common/Common/ThemeExpression.cpp:310-346` uses that luminance for `EnsureContrast`, while
  `ToOklch`/`FromOklch` preserve the foreground alpha.
- **Evidence:** the expressive-theme plan requires candidate-aware WCAG contrast and alpha preservation, but
  `Tests/RedConfigureTests/RedConfigureTests.cpp:593-618` covers only opaque candidates.
- **Impact:** A 50%-alpha black foreground on white is treated as 21:1 even though the rendered composite is
  gray at roughly 4:1. Themes can pass a requested 7:1 accessibility contract while visibly failing it.
- **Effort:** M.
- **Risk:** MED; the correct policy needs an explicit backdrop contract for translucent backgrounds.
- **Confidence:** HIGH.
- **Fix sketch:** Compute contrast from rendered/composited colors. Preserve alpha only when a candidate with
  that alpha can meet the target; otherwise return unattainable. Require an opaque background or supply the
  known surface backdrop explicitly.
- **Required tests:** opaque vectors at 3.0/4.5/7.0, 50%-alpha black/white, translucent background policy,
  unattainable target, alpha preservation, and RedConfigure's displayed ratio.

### F4-RC-06 - Duplicate Theme hides validation failures and can never duplicate long names

- **Evidence:** `RedConfigureSession::DuplicateActiveTheme` at
  `RedConfigure/RedConfigureSession.cpp:1183-1202` returns one `false` for invalid ID, name length, empty name,
  and actual collision.
- **Evidence:** `RedConfigure/RedConfigureRoot.cpp:2124-2153` treats every failure as a collision, retries up to
  100 suffixes, then returns without user feedback.
- **Impact:** A valid 64-character source name becomes permanently non-duplicable when the `Copy` suffix is appended.
  Other validation failures trigger 100 pointless attempts and silently do nothing.
- **Effort:** S-M.
- **Risk:** LOW.
- **Confidence:** HIGH.
- **Fix sketch:** Return a typed result (`Success`, `DuplicateId`, `InvalidId`, `NameTooLong`, etc.). Retry only
  `DuplicateId`; truncate the source name to leave room for the localized suffix, and show a localized error for
  non-collision failures.

### F4-LOC-01 - New RedConfigure UI is untranslated in every affected satellite

- **Evidence:** The range adds 77 `IDS_REDCONFIGURE_*` strings. All 77 are byte-identical to embedded English in
  each of `RedConfigure/Lang/cs-CZ/RedConfigure-cs-CZ.rc`,
  `RedConfigure/Lang/ja-JP/RedConfigure-ja-JP.rc`, and
  `RedConfigure/Lang/sk-SK/RedConfigure-sk-SK.rc`.
- **Evidence:** `docs/dev/Localization.md:174-185` requires ordinary UI strings to be translated in each
  satellite; only documented language-neutral tokens may intentionally duplicate English.
- **Impact:** The complete new palette, workflow, batch, validation, and review/export UI appears in English for
  Czech, Japanese, and Slovak users. Structural parity tests pass while actual localization coverage regresses.
- **Effort:** L (translation/review, not code generation).
- **Risk:** LOW.
- **Confidence:** HIGH.
- **Fix sketch:** Translate the 77 strings with native review, preserving placeholder tokens exactly. Add a
  report for suspicious source-identical ordinary UI strings; do not make exact equality globally blocking
  because some real translations legitimately match brands/tokens.

### F4-SET-02 - Duplicate inline themes are discarded on canonical save

- **Evidence:** `Common/Common/SettingsStore.cpp:1367-1375` keeps the first exact theme ID and drops later
  duplicate entries after incrementing a diagnostic count.
- **Evidence:** structurally unparseable entries are retained in `opaqueThemeEntries`, but duplicate entries are
  not. The serializer at `:5772-5783` can therefore never re-emit them.
- **Impact:** Loading and then saving settings permanently deletes authored theme data. The user gets only a
  debug warning, not a repair UI or preserved raw entry.
- **Effort:** S-M.
- **Risk:** LOW.
- **Confidence:** HIGH.
- **Fix sketch:** Define theme-ID uniqueness (including case policy) explicitly. Keep the deterministic winner
  active, but retain every rejected duplicate as opaque JSON with original array position so a canonical save is
  non-destructive. Add exact and case-variant duplicate round-trip tests.

### F4-SET-03 - Shared uint32 reads wrap out-of-range JSON

- **Evidence:** `Common/Common/SettingsStore.cpp:572-581` accepts any yyjson unsigned integer and casts it to
  `uint32_t` without a maximum check.
- **Evidence:** callers include diagnostics memory/flush limits at `:3415-3451`, cache watcher limits at
  `:4249-4258`, history, masks, workers, DPI, and legacy shortcut modifiers.
- **Impact:** A hand-edited/imported value of `4294967297` becomes `1`; `4294967296` becomes `0`. This can turn
  diagnostics flush into near-continuous I/O, alter retention limits, or silently convert invalid settings into
  valid but wrong behavior.
- **Effort:** S-M.
- **Risk:** LOW.
- **Confidence:** HIGH.
- **Routing:** Observatory Track 7 owns the root fix. Make the typed read reject values above
  `UINT32_MAX`, then apply per-field range validation before assignment. Test boundary-1, boundary, boundary+1,
  and a very large unsigned value for every operational-limit family.

### F4-TEST-01 - Source-contract tests are a brittle shadow compiler

- **Evidence:** `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` is now approximately 2,578 lines and the
  range adds 907 lines while deleting 446.
- **Evidence:** seven additional `*SourceContracts.Tests.ps1` files assert exact symbol names, call shapes,
  regex counts, and implementation text for render resources, modal shells, packed buffers, plugin lifetime,
  plugin configuration, posted payloads, and viewer chrome.
- **Impact:** Semantics-preserving refactors break tests, while semantic defects pass when the expected token is
  present. This review found exactly that pattern: session-end had a passing writer-seam test but no coordinator
  race, and payload teardown had exactly-once destructor coverage but no queued-message/HWND-reuse coverage.
- **Effort:** L.
- **Risk:** MED; deleting guards before behavioral replacements land would reduce coverage.
- **Confidence:** HIGH.
- **Routing:** Observatory Track 17 owns this migration. Keep source checks only for build graph, forbidden APIs,
  exports, and contracts that cannot be observed at runtime. Move ownership, ordering, bounds, and failure
  behavior into C++/integration tests; split the monolith by subsystem.

### F4-TEST-02 - TestSandbox contamination is non-blocking

- **Evidence:** committed CI pass 2 reports exit 0 with `test_sandbox_audit.is_clean=false` and 238 issues in
  `Specs/TestRuns/SINON/Continuation/2026-07-15_lighthouse_track_d_ci_gate/ci_pass_2/run-all-tests-results.json`.
- **Evidence:** `Tools/Run-AllTests.ps1:1332` prints issue counts in yellow, while the aggregate contract treats
  the audit as informational.
- **Impact:** Hundreds of stale per-test directories can retain locks, consume disk, affect path discovery, and
  contaminate later tests without failing CI. The audit signal exists but has no enforcement value.
- **Effort:** M.
- **Risk:** LOW once each suite owns its cleanup.
- **Confidence:** HIGH.
- **Routing:** Observatory Track 0 and Lighthouse D18 own this. First eliminate current leaks and classify
  intentional retained artifacts; then make unexpected current-run leftovers blocking while preserving bounded
  diagnostics.

### F4-DX-01 - Theme number parsing depends on mutable process locale

- **Evidence:** `Common/Common/ThemeExpression.cpp:72-87` parses the durable expression grammar with
  `std::wcstod`, which uses mutable C `LC_NUMERIC` state.
- **Evidence:** canonical formatting emits `.` decimals. No current production call changing `LC_NUMERIC` was
  found, so a non-English Windows locale alone is not a trigger; an in-process library/caller changing the C
  locale is required.
- **Impact:** The same saved theme can become invalid or interpret decimal separators differently after a
  process-global locale change. Grammar behavior is not isolated from unrelated components.
- **Effort:** S-M.
- **Risk:** LOW.
- **Confidence:** MED because no current production locale mutation was found.
- **Fix sketch:** Parse with a locale-independent ASCII decimal routine or an explicit C-locale API. Add tests
  under a comma-decimal locale and concurrently with unrelated locale-sensitive work.

### F4-REPO-01 - Theme mockup has no generated-artifact contract

- **Evidence:** `Specs/Mockups/ThemeGalleryWorkbench/.gitignore` ignores only `.vite/`; tracked `dist/` contains
  generated hashed assets.
- **Evidence:** `Specs/Mockups/ThemeGalleryWorkbench/package.json` has a normal `vite build` script but no README, deployment consumer,
  reproducibility check, or regeneration rule explaining why `dist` is source-controlled.
- **Impact:** Dependency bumps produce opaque minified churn, reviewers cannot tell whether source and `dist`
  agree, and stale generated output can be committed independently of source.
- **Effort:** S.
- **Risk:** LOW.
- **Confidence:** HIGH.
- **Routing:** Observatory Track 18 owns the policy decision. Either ignore/remove `dist` and build it in the
  preview/deployment workflow, or document it as a release artifact and add deterministic regeneration plus a
  clean-tree comparison gate. Move Vite/build-only packages to `devDependencies` in either case.

## Execution tracks

### Track 1 - Serialize session-end persistence

**Owns:** F4-SET-01. **Priority:** P1. **Dependencies:** coordinate with Observatory Tracks 7 and 14; do not
implement a second settings commit coordinator.

1. [x] Add the RED delayed-queued-save versus real `WM_ENDSESSION` test.
2. [x] Introduce one coordinator transition that admits a final settings-only snapshot, supersedes older
   snapshots for the same app, fences later submissions, and preserves the bounded no-schema/no-modal contract.
3. [x] Keep low-level staging unique per attempt through the existing `Common::Files::LocalFileTransaction` and
   clean only the attempt-owned sibling.
4. [x] Update `Specs/Core/Core_StartupBootstrap.md` and `Specs/Core/Core_SettingsStore.md`.
5. [x] Verify the focused real session-end/settings family and archive the passing run. Per the bounded
   convergence decision used for recent completed operations, do not add a redundant Full rerun when the
   focused behavioral proof, affected build, and current CI baseline are sufficient for this isolated fix.

### Track 2 - Make RedConfigure export fail closed

**Disposition:** CLOSED by Observatory Track 19 in `ce4e6c4ca`; do not reimplement.

1. [x] Replace string-inspected diagnostics with structured codes/severities.
2. [x] Block affected owner/culture export on source/target read or parse failure.
3. [x] Preserve original target content and prohibit fallback-generated replacement after failed target ingestion.
4. [x] Add fault tests for source/target read and parse failures plus a warning-only target-extra case.
5. [x] Update `Specs/Core/Core_RedConfigure.md` and `Specs/UI/UI_RedConfigure.md`.

### Track 3 - Make posted-payload drain queue-safe

**Owns:** F4-MSG-01. **Priority:** P1. **Dependencies:** none; high fan-in requires one shared fix.

1. [x] Add the RED destroy, HWND-reuse, and stale-queue dispatch test under ASan.
2. [x] Replace payload-address `lParam` values with opaque registered tokens. Drain invalidates tokens before
   deletion; stale queued/unregistered/type-mismatched tokens return null and are never adopted.
3. [x] Re-run all payload consumers' teardown tests, Compare/Find coalescing tests, and viewer close/unload stress.
4. [x] Update `Common/Helpers.h` guidance, `AGENTS.md`, the Win32 skill, and
   `Specs/Core/Core_SharedHelpers.md`.

### Track 4 - Put RedConfigure mutations behind one transactional session API

**Disposition:** CLOSED by Observatory Tracks 16/19 in `ce4e6c4ca`; do not reimplement.

1. [x] Remove the public mutable preview-model accessor and add typed session commands/results.
2. [x] Add stable identity/model revision and full-request comparison to localization/theme previews.
3. [x] Make apply validate all captured `before` states and fail atomically on drift.
4. [x] Replace textual recipes with parsed source transformations and explicit applicable/skipped results.
5. [x] Return typed Duplicate Theme failures and localized feedback.
6. [x] Add one-step undo/redo and stale-approval tests for every command.

### Track 5 - Align theme grammar, schema, persistence, and accessibility

**Owns:** F4-THEME-01, F4-THEME-02, F4-SET-02, F4-DX-01. **Priority:** P2. **Dependencies:** settle canonical
function spelling and alpha-compositing policy before updating goldens.

1. [x] Emit documented camelCase function names and add formatter/schema parity coverage.
2. [x] Define and implement composited alpha-aware contrast semantics with an explicit opaque-background contract.
3. [x] Preserve duplicate/rejected inline theme entries as positioned opaque values.
4. [x] Make numeric parsing locale-independent (closed by Observatory Track 19).
5. [x] Update theme/settings specs and run the focused RedConfigure, SettingsSchema, and theme validation owned by
   the changed boundaries. Use Observatory Track 0's current broad-gate evidence rather than another redundant Full.

### Track 6 - Complete localization and repository hygiene

**Owns:** F4-LOC-01. **Routes:** F4-TEST-01/02, F4-SET-03, and F4-REPO-01 to their existing Observatory
owners.

1. [>] Translate and review the 77 RedConfigure strings — routed to Operation Rosetta Lantern because native
   linguistic review is required.
2. [>] Add a non-blocking source-identical translation report with a reviewed neutral-token allowlist — owned by
   Operation Rosetta Lantern.
3. [x] Execute/reconcile the routed Observatory tracks rather than duplicating their implementation here.

## Recommended order

1. ~~F4-SET-01 session-end serialization.~~ **CLOSED by `2888e34a3`.**
2. ~~F4-MSG-01 payload queue/storage ownership.~~ **CLOSED by this operation with focused Debug/ASan/perf proof.**
3. ~~F4-RC-01 export fail-closed behavior.~~ **CLOSED by Observatory Track 19.**
4. ~~RedConfigure session mutation and stale-preview tracks (F4-RC-02 through F4-RC-06).~~ **CLOSED by
   Observatory Tracks 16/19.**
5. ~~Theme grammar/accessibility/preservation (F4-THEME-01/02, F4-SET-02, F4-DX-01).~~ **CLOSED by this operation
   plus Observatory Track 19.**
6. ~~Observatory Track 0 gate/sandbox ownership.~~ **CLOSED by bounded convergence.**
7. Translation is routed to **Operation Rosetta Lantern**; source-contract architecture is routed to **Astrolabe**.

## Considered and rejected

These candidates were directly checked and must not be re-filed without new evidence:

- **Plugin configuration values out-of-bounds:** rejected. `ParseConfiguration` allocates one default value
  per schema field before overlay, so both UI loops have matching `fields`/`values` sizes.
- **Canceled `WM_ENDSESSION` permanently claims save ownership:** rejected. The owner is claimed only inside
  the `sessionEnding == true` branch. F4-SET-01 is the real defect: coordinator bypass and temp collision.
- **Unknown/theme JSON conversion failure should be preserved after failure:** rejected as stated.
  `ConvertYyjsonToJsonValue` supports every JSON type; practical failure is allocation failure, where a second
  preservation allocation is not a reliable recovery path.
- **Shared HWND render-target helper lost resize/DPI behavior:** rejected. FunctionBar and StatusBar both reset
  targets on size/DPI changes and on `D2DERR_RECREATE_TARGET`; extraction retained behavior.
- **ViewerSpace detached worker unloads its DLL:** rejected. The detached worker owns the viewer reference and
  deliberately retains a permanent unload gate after an uncooperative provider forces quarantine.
- **ViewerWeb raw post-failure message corrupts a new request:** rejected at reportable severity. The raw message
  carries no payload and the handler rechecks per-object request IDs. Identity-bound payload completion remains
  the primary path.
- **Plugin lifecycle helper changed shutdown/unregister semantics:** rejected. Pre/post manager behavior is
  equivalent, and shared-module localization registration now has reference-count coverage.
- **Packed `FileInfo` helper under-allocates or misaligns entries:** rejected. Aggregate size is bounded to
  `ULONG_MAX`, entries are aligned to `alignof(FileInfo)`, names include terminators, and traversal validates
  offsets and UTF-16 byte counts.
- **New yyjson builders borrow temporary strings:** rejected. Dynamic keys/values in the reviewed changes use
  copying APIs or document-owned mutable strings.
- **Custom window-message collisions:** rejected. The current registry scan found no duplicate numeric
  `WM_APP`/`WM_USER` values.

## Closeout rule — satisfied

- [x] Every finding is implemented with named behavioral proof, closed by its authoritative completed owner, or
  explicitly routed to one active owner.
- [x] The final current-head x64 Debug all-consumer build completes with zero warnings/errors.
- [x] Focused rebuilt Debug/ASan/runtime/performance and source-contract proof covers every implementation owned
  by this plan.
- [x] Observatory Track 0 supplies the current bounded CI/Full/TestSandbox classification; no unowned broad-gate
  failure is hidden and no redundant broad rerun is required for this delta.
- [x] Durable payload, theme grammar/contrast, settings preservation, localization-review, and test contracts are
  present in authoritative specs/repository guidance.
- [x] F4-LOC-01 and F4-TEST-01 have unique active WIP owners rather than stranded requirements in this Done plan.
- [x] This file is moved to `Specs/Plans/Done/` with exact evidence.
