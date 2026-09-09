# Tools Inventory, Simplification, and Governance

> **FILE OPERATIONS REVIEWED (2026-08-21).** Phase 0 adds no public Tools command. Any later File Operations script or Pester file follows this plan's inventory, help, caller, and source-contract rules.


> **NON-NORMATIVE ACTIVE PLAN.** This plan owns the bounded reorganization of
> repository tooling. Current command behavior remains governed by the existing
> build, testing, release, Terminal, and contributor contracts until each change
> is implemented and reconciled into those authoritative documents.

## Status

- **State:** ACTIVE — requested simplification and build-safety hardening are
  complete; Full-suite closeout remains blocked by unrelated Commands self-test
  failures
- **Priority:** P1 developer-experience and CI maintainability
- **Planned at:** `f9cee9d47`
- **Ownership boundary:** `Tools/`, direct build/workflow/documentation callers,
  tooling tests, tooling guidance, and the authoritative tooling-governance
  contract
- **Drift check:** `git diff f9cee9d47..HEAD -- Tools .github/workflows Directory.Build.targets RuntimeDependencies.props build.ps1 vcpkg-install.ps1 README.md docs Specs/Testing AGENTS.md .github/skills`
- **Parallelism:** inventory/docs, shared helper extraction, and Terminal/cache
  cleanup may proceed in parallel only while their file ownership is disjoint;
  integration and final verification remain serialized.

## Baseline

At the planned-at commit, `Tools/` contains 142 tracked files. Fifty files are
at the directory root, including runnable commands, dot-sourced libraries,
modules, data, compatibility aliases, and historical Terminal selection tools.
There is no `Tools/README.md` or machine-validated general tool inventory.

The main confirmed maintenance problems are:

1. purpose, consumers, side effects, ownership, and lifecycle are not visible
   without repository-wide searches;
2. runnable commands and implementation libraries share one namespace;
3. `TestRunPlan.ps1` combines sandbox, quarantine, reporting, and suite planning;
4. several source audits and test-run report commands duplicate infrastructure;
5. two source-line measurement commands overlap;
6. common Pester helpers are copied between test files;
7. superseded Gate-0/Round-4 executable tooling remains mixed with current
   Ghostty runtime tooling;
8. the Ghostty runtime CI cache does not hash the builder's full dependency
   closure;
9. public command help and naming requirements are incomplete;
10. no normative rule or project skill prevents the same drift from returning.

## Required outcomes and checklist

### 1. Inventory and discoverability

- [x] Add `Tools/README.md` as the human entry point.
- [x] Add a checked-in machine inventory that classifies every tracked file by
      kind, visibility, lifecycle, purpose, owner, and consumer/usage model.
- [x] Provide a command that validates and renders the inventory.
- [x] Add deterministic tests for unclassified, missing, and stale entries.
- [x] Link the catalog from root contributor documentation.

### 2. Directory and command surface

- [x] Reserve the `Tools/` root for stable public/automation entrypoints and
      navigation files.
- [x] Move private implementation behind explicitly exported modules grouped by
      Build, Testing, Reporting, Auditing, Packaging, and Terminal ownership.
- [x] Preserve reviewed paths used by MSBuild, CI, documentation, or external
      operators as thin shims where an atomic caller migration is unsafe.
- [x] Mark every compatibility shim with its canonical replacement and support
      policy; delete the unused `remove-stream.ps1` shim after the explicit
      follow-up removal decision.
- [x] Require canonical `Verb-Noun.ps1` names for new public commands.

### 3. Shared implementation simplification

- [x] Split sandbox, quarantine, reporting, and suite construction out of the
      monolithic test-plan implementation with explicit exports.
- [x] Consolidate repository source scanning for visible/native/UI audits while
      preserving their current command paths and output contracts.
- [x] Consolidate test-run archive discovery, formatting, and summary helpers;
      retain explicit throw-versus-null policies at each front end.
- [x] Make `Measure-SourceLines.ps1` the sole source-line implementation, with
      categorized repository and raw solution-backed modes; retain
      `InvokeSolutionCloc.ps1` as the confirmed hand-use compatibility path and
      forward it to Solution scope.
- [x] Add shared Pester support for repeated repository-text and assertion
      helpers, then migrate the duplicate consumers.

### 4. Terminal current-versus-history boundary

- [x] Identify the exact active Ghostty runtime builder/helper/test closure.
- [x] Move superseded executable Gate-0/Round-4 tooling into an explicitly
      historical namespace without rewriting immutable evidence semantics.
- [x] Replace default-suite execution of the frozen behavioral cohort with a
      small integrity/closure contract and an explicit archival verification
      profile triggered only for historical-material changes or manual review.
- [x] Retire or clearly internalize the generic lock lifecycle facade whose
      default lock is absent; do not remove an externally plausible path without
      a compatibility shim and documented replacement.
- [x] Hash the current patch and every direct runtime-builder helper in the CI
      cache key, with a test proving dependency closure.

### 5. Durable governance

- [x] Add complete comment-based help to every public command, including
      prerequisites, outputs, side effects, exit behavior, and an example.
- [x] Add an authoritative `Specs/Testing/Testing_ToolingGovernance.md` contract.
- [x] Add a project-local `tooling-governance` skill and register it in
      `AGENTS.md`.
- [x] Update build/testing skills and contributor guidance to point at the
      inventory and normative contract.
- [x] Add enforcement that every added or renamed tool updates inventory, help,
      tests, callers, and the owning spec in the same change.

### 6. Verification and closeout

- [x] Run inventory/help/source-contract Pester tests.
- [x] Run affected audit, reporting, build-policy, Terminal, and test-plan tests.
- [x] Run `Tools/Get-SpecInventory.ps1 -FailOnFindings`.
- [ ] Run `Tools/Run-AllTests.ps1 -Suite Full`. The canonical run reached the
      Commands lane at 405 passed, 18 failed, and 0 skipped; the failures are
      unrelated application selftests, so the STOP condition applies.
- [x] Reconcile durable requirements into normative specs and contributor docs.
- [x] Record the final file-count/root-surface delta and any retained shims.
- [ ] Mark this plan complete and move it to `Specs/Plans/Done/`.

## Implementation and verification evidence

- The schema-v3 inventory classifies 163 of 163 discovered `Tools/` files with
  zero blocking findings. Every record has prerequisites; all 9 compatibility
  records have an explicit support policy, every non-null replacement names an
  existing exact path, and all 52 retained historical records name one or more
  of the three reviewed manual reproduction roots.
- The root surface changed from 50 files at `f9cee9d47` to 37 files (net -13).
  Total inventory grew from 142 to 163 because the cleanup adds maintained
  modules, tests, profiles, documentation, and governance metadata. Fourteen
  private root implementations moved behind grouped modules.
- Retained compatibility paths are inventory-governed. `TestRunPlan.ps1` and
  `SanitizedEnvironment.ps1` are logic-free loaders required by sealed manual
  profiles; dormant generic Terminal commands are consumed by the explicit-lock
  manual lifecycle and `External/TerminalEngine/TerminalEngine.proj`.
  `remove-stream.ps1` had no live consumer or manual support contract and was
  deleted after explicit removal authorization. `InvokeSolutionCloc.ps1` was
  restored after confirmed developer hand use; it is a supported thin wrapper
  over `Measure-SourceLines.ps1 -Scope Solution`, not a second implementation.
- Current tooling validation passed 521 of 521 tests. The final integrated
  tooling, coordination, packaging, release, inventory, and governance set
  passed 166 of 166 tests, including 10 of 10 focused ProcessStreaming tests.
  All 153 relevant PowerShell files beneath `Tools/`, `Installer/`, and the root
  build entrypoints parse without errors.
- `Get-SpecInventory.ps1 -FailOnFindings` reports zero blocking findings.
  `git diff --check` is clean. The final serial Debug x64 RedSalamander build
  completed with zero warnings and zero errors, and the focused native
  FileOperations smoke passed 3 of 3. Disposable ownership/contamination
  markers are absent; a foreign-checkout RedSalamander process remained alive
  throughout verification.
- All 20 sealed Terminal test hashes match
  `HistoricalTerminalProfiles.json`. Frozen workflow hashes remain
  `e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c`
  for Gate0 and
  `c714fcfe1a327f21b406a366283289cfb15f757cfe86c27d2c4904975e3b8748`
  for Round4.
- The Full run's 18 Commands failures were:
  `settings_store_unsupported_schema_backup_recovery`,
  `file_action_selected_paths_file_lifecycle`,
  `settings_hot_reload_creates_missing_directory`,
  `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation`,
  `pane_view_options_preview_uses_builtin_embedded_viewer_with_empty_associations`,
  `pane_view_options_preview_falls_back_to_item_properties_when_no_embedded_preview_matches`,
  `pane_view_options_preview_properties_card_scrolls_and_uses_rainbow_theme`,
  `alternate_view_launches_external_viewer_action`,
  `view_with_parameterized_command_launches_external_viewer_action`,
  `edit_with_parameterized_command_launches_external_editor_action`,
  `cmd_pane_batchRename_window_success_callback_parent_child_execution_order`,
  `cmd_pane_batchRename_cancel_mid_batch_tracks_completed_rows`,
  `cmd_pane_batchRename_execution_engine_direct`,
  `cmd_preferences_dialog_keyboard_header_drag_reorders_columns_without_sort`,
  `cmd_preferences_dialog_keyboard_header_resize_changes_visible_width`,
  `cmd_preferences_dialog_keyboard_reordered_resized_columns_stay_stable`,
  `cmd_preferences_dialog_keyboard_export_live_dx_interaction`, and
  `cmd_preferences_dialog_themes_save_theme_live_dx_interaction`.
  No application-source fix was attempted as part of this Tools-only plan.
- Closeout remains: resolve or independently establish the Commands baseline,
  rerun the canonical Full suite, then check the final two items and move this
  plan to `Specs/Plans/Done/`.

## Follow-up simplification and build-safety hardening

Requested after the initial inventory review:

- [x] Delete the empty optional quarantine ledger; absence is the canonical
      no-quarantine state while an explicitly supplied ledger retains the
      reviewed blocking repair semantics.
- [x] Remove compatibility paths with no live tracked consumer, useful manual
      support contract, or protected reproduction dependency.
- [x] Remove historical artifacts that do not participate in a concrete manual
      reproduction/authentication entrypoint; retain only the minimal transitive
      closure needed to launch valuable archival checks by hand.
- [x] Replace the repository-wide artifact-operation collision domain with a
      canonical repository/output scope so disjoint platform/configuration
      artifacts do not block each other.
- [x] Limit interactive-desktop serialization to operations that actually use
      the shared desktop instead of holding it across unrelated build and
      noninteractive test work.
- [x] Require exact executable path, repository/output scope, PID start identity,
      and owned-process evidence before any cleanup may terminate a process.
      Missing evidence must fail closed, and a process from another checkout
      must never be stopped.
- [x] Add focused concurrency, stale-owner/PID-reuse, path-mismatch, and
      foreign-checkout survival tests; update the normative spec, inventory,
      contributor documentation, and build/tooling skills in the same change.
- [x] Re-run focused Pester, inventory/spec validation, the current tooling
      profile, and the appropriate canonical build/test gates.

## Compatibility and safety rules

- Do not change a command's exit code, output schema, or default side effects
  merely to move its implementation.
- Do not edit frozen Terminal evidence bytes or derive historical identities
  from mutable current manifests.
- Do not delete a path with plausible external use until it has a documented
  replacement and reviewed removal authorization. Once authorized, remove the
  shim and its stale inventory, tests, and prose together.
- Do not weaken CI, Full, packaging, release, or runtime validation. Historical
  checks may leave the default lane only after equivalent integrity coverage and
  an explicit archival lane exist.
- Do not introduce a second inventory source of truth. Human documentation is
  rendered or reviewed against the checked-in machine inventory.

## Verification commands

Run focused commands as their components land, then finish with:

```powershell
Invoke-Pester .\Tools\Tests\ToolInventory.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\ToolingGovernance.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\ReleaseWorkflowPolicy.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\DocumentationDriftContracts.Tests.ps1 -PassThru
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Run-AllTests.ps1 -Suite Full
git diff --check
```

## Done and STOP conditions

Move this plan to Done only when the inventory covers the entire tracked Tools
tree, stable entrypoints remain callable, current and historical Terminal paths
are unambiguous, focused tests and the Full suite pass, and all lasting rules are
present in the authoritative spec and skill.

Stop and report rather than guessing if a move would invalidate protected
historical identity, an externally consumed command cannot be forwarded without
behavior change, a workflow policy requires a maintainer product decision, or
the full suite reveals an unrelated pre-existing failure that cannot be safely
separated from this work.
