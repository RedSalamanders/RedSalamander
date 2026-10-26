# Scoped testing and validation cost

Status: ACTIVE: implementation and qualification in progress.
Date: 2026-10-05
Owner: [current validation contract](../../Testing/Testing_ValidationEvidence.md)

## Accepted request

All active native test sources and helpers use `Scope.Tests.Something.h/.cpp`; Scope identifies the component and Something the scenario, runner, fixture, or helper. Migrate project entries, includes, inventories, tooling and current documentation together. Immutable historical evidence and imported/external source are not active test files.

Everyday iteration selects affected scopes from committed changes since the merge base plus staged, unstaged, deleted, renamed and untracked paths. Report paths and reasons. Unknown executable inputs, shared infrastructure and dependency/build changes conservatively widen coverage. A focused pass is never a full repository verdict.

Reuse only successful local evidence for identical source/dependency closure, executable/runtime bytes, architecture, configuration, runner arguments and environment. Invalidate before execution; publish only after success and an unchanged post-run identity. Corrupt/missing evidence executes again. Explicit force runs, performance investigations and environment-dependent/manual acceptance remain available.

Before a PR push, full coverage is required across local and CI execution. Delegate only the identical scopes/configuration that the enabled PR workflow actually runs. A nightly gate is not a PR gate. A base/merge-tree change, configuration difference or modified runner invalidates equivalence. Keep main/release acceptance distinct from PR feedback.

## Review evidence

- DxUi: test.ps1 supports suite and named-test filters, but unconditionally runs tooling and a benchmark; CI runs six native profiles plus a separate validation job and paired benchmark. The initially running main validation subsequently passed; an earlier main ASan menu-popup fixture had failed even though its PR was green.
- RedXe: test.ps1 is an unconditional sequence of standalone suites, tooling, host WARP and crash checks. PR CI runs all tests on x64 Release; main runs all six profiles. Latest main CI passed.
- RedSalamander: governed Affected, ExplainPlan and exact Resume already exist. Reuse these contracts. PR CI runs Suite PR on x64/ARM64 Debug; long in-product CI suites run nightly and sanitizer suites weekly. Latest CI failed in format, migration and native jobs.
- RedPrism: test.ps1 orchestrates document, IO, host, compositor and in-product checks on the current UI branch. No checked-in or enabled GitHub workflow was found. Plan only; do not rename files, edit runners, build, or change CI during this work.

## Implementation sequence

- [x] Inventory active source files, test scopes, dependencies, desktop/network requirements and CI coverage.
- [x] Rename active native tests/helpers and reconcile all current callers; enforce naming.
- [x] Provide explainable affected/default iteration and explicit full/pre-push entrypoints.
- [x] Prove conservative fallback, rename/delete/untracked discovery, dependency fan-out and invalid selectors.
- [x] Prove unchanged reuse, source/binary/configuration/environment invalidation, failures and post-run mutations.
- [x] Align CI commands and pre-push delegation without weakening existing main/release gates.
- [ ] Run focused tooling/scenario checks, applicable builds and noninteractive native tests; record unavailable/manual gates separately.
- [ ] Persist durable requirements in the owning contract and agent guidance; close only after required qualification.

## Qualification constraints

No test may seize the person's foreground or pointer without agreement to the time. Use existing harness desktop leases. Native ARM64 claims require ARM64 execution; cross-build evidence is separate. No consumer pin changes are implied by library changes. Retain baseline/candidate fixture and source provenance for resource comparisons.

## Implementation and reviewed qualification

All 100 active native tests/helpers/self-test translation units and fake plugin seams use contextual names. Current project/include/registration/scanner references are reconciled. Immutable Done/TestRuns/Terminal evidence remains historical. The root command preserves canonical Affected, ExplainPlan, source/build attestation and exact explicit Resume; it introduces no independent runtime-evidence cache. Precise standalone tags select their owners. Unknown/shared mappings cover the complete canonical Full plan, including SQLite viewer entries; PR deferral removes only complete equal portable entry contracts and retains unmatched arguments, writers and in-product scenarios.

Qualification completed: the new scoped/PR policy tests, impact and inventory regressions, canonical runner planning tests, tooling/module/governance, protected-history/spec consistency, documentation drift and release workflow policy. The 196-case combined check had one documentation inventory failure; the omitted original Tests README was restored and its entire seven-case owning suite passes after repair. The final 231-case source-contract/naming check passes after the renamed include assertion and historical inventory were corrected. The later scheduler policy/governance check separately covers latest-success/failed/cancelled/missing history, dispatch forcing and workflow wiring. All 123 affected native files pass the pinned CI formatter; tool and spec inventories report zero blockers. PrePush Explain defers exact PR matches and preserves unmatched local obligations; filtered plans remain NOT_EVALUATED.

Native compilation is blocked before renamed product/tests compile by the pinned vcpkg libjpeg-turbo 3.2.0 pkg-config check: libturbojpeg requires spng.pc, which the restored libspng package does not supply. Two canonical Debug build attempts reproduce that dependency failure. Preserve their logs; do not bypass receipts or claim a native pass. Current GitHub [37350902259](https://github.com/DualTail/RedSalamander/actions/runs/37350902259) failed format/migration/native jobs on d93dea1c64fe; individual failed-job logs were unavailable through the API. Do not infer their cause from the local package failure.

Nightly/weekly source monitors skip unchanged code only after the latest completed attempt succeeded; manual dispatch, failures/cancellations and unreadable history run. A scheduled skip is not new environment/runtime qualification. Pending closeout: resolve the dependency provisioning failure through the owning build contract, qualify the renamed native targets/profiles and remaining in-product/performance/foreground scenarios, then verify the changed GitHub policies. Foreground scenarios require agreement to the time. No dependency pin or consumer/library implementation is changed here. Keep this plan active; no Full repository acceptance is claimed.

Final audit restored 57 historical source-snapshot blobs exactly from HEAD after the broad rename filter had incorrectly included them. The active inventory excludes archived/external trees, and naming rejects archival inventory entries. The public-runner fixture preserves an old-named historical source. Git confirms zero changes under Specs/TestRuns and protected Done history.

Final receipt review includes documentation, skill Markdown and source-origin mappings in validator identities. Uncommon inline/generator extensions invalidate build attestation. Documentation edits select only independent validators; native scopes remain omitted. The two standalone runner fixtures pass all eighteen cases, including those invalidation boundaries.

The final owning failure-report/build contract repair passes 41 assertions, zero failures and one explicitly unbuilt runtime skip. These repairs reconcile the renamed routing header and PowerShell schedule policy assertions; they do not substitute for the blocked native build.

Final scoped/impact/tooling policy verification passes 29 assertions with no failures or skips after the receipt-identity update.

Remaining support contracts pass 30 assertions. The standalone relocated-sandbox native fixture separately compiles and executes on x64 Debug with the required profile and existing marked sandbox, proving the renamed support header outside a checkout. This independent fixture does not qualify the dependency-blocked product/native solution.

DxUi preserves one Windows tooling qualification and one Linux portability qualification, with the staging fixture only once. Conditional native PR coverage reuses its existing NativeScope contract; skipped native jobs are never delegated. Native receipts retain identical-code reuse across validator prose edits. The standalone fixtures cover these boundaries in eighteen cases.

The final basename audit adds six initially missed active sources: Common sandbox/activation headers, the owned-shutdown self-test pair, MTP fake backend and Terminal VT upgrade tests. All 100 active sources now use contextual names; the separately inventoried historical Terminal adapter remains sealed.

The expanded basename migration passes 250 source/documentation/scoping/native-sandbox assertions, zero failures or skips. The pinned 22.1.3 formatter passes all 123 affected native files after include-order/line-ending repairs. Tool and spec inventories still report zero blockers.

Pre-merge review also restored 21 historical capture/provenance and paused payload files to their exact original blobs. Their recorded paths describe the historical source identity; only current harness links in the owning README follow the active naming migration.

Post-merge policy review adds a clean committed candidate barrier around PR coverage discovery. Staged, unstaged, untracked and concurrent changes, including a concurrent clean commit, retain local obligations. All seven owning scoped-policy regressions pass. Original PR #114 was merged after the user confirmed local checks were done; GitHub refused to start its jobs because of account billing/spending limits. This is an external execution block, not a hosted pass. Keep historical dependency failures and remaining runtime gates distinct from that user-confirmed local check.
