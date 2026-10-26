# Validation record — current checkpoint October 1, 2026

The [original V55 gallery](implementation/2026-10-01-v55-final/README.md) retains 562 static and ten
protected captures, 512 reviewed image hashes and native 96/144-DPI provenance. Broad Debug FileOps
remains V55 197/0/20. V65 x64 build status: passed, with zero warnings or errors; Release Full/Fresh `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e`:
2211 passed, 0 failed, 46 skipped (2257 total). The independent A78/B19 resource
review records acceptance for I26 with stated limits: the final matched Embedded presentation and logical-resource workload, with all eight processes and outliers retained; the earlier ProgramDatabase pair stays separate, without whole-product or driver-memory claims. See the [V65 packet](../../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/README.md)
and [I26 checkpoint](../../Done/UI_FileOperationsReview_2026-09-16.md) for evidence identities, historical failures and limits.
The gallery was not recaptured for V65. Native 192-DPI, external AT, consumer ARM64/ASan, I19 and H4
remain separate; dated entries preserve their original results.

## Historical validation records

September 27 current evidence: the [focused Debug inspection journey](implementation/2026-09-27-conflict-inspection/README.md) passes all twelve pane/role/invocation combinations, with 28 byte-exact originals. [Seeded legacy migration](implementation/2026-09-27-legacy-layout/README.md) passes on native 96/144-DPI displays in Release V10 and Debug V13; four Debug originals are reviewed. The final Debug gallery reports 1/0/0 and all sixty same-window terminal transitions pass. V9's intermittent footer failure remains unexplained; its later passing runs are not a fix. The current library pin's multiline technical-grid clipping and the final configuration/DPI/AT/resource gates remain open. [Build/run/source receipts and retained failures](../../../TestRuns/4cb089111a23/Validation/2026-09-27_migration-transition/README.md).

Latest September 23 update: [readable, scrollable confirmations](implementation/2026-09-23-confirmation-layout-release/README.md). The nonactivating Release gallery passes with 376 captures, including 76 visually reviewed confirmations at 96/144 DPI. Integrated acceptance remains open; earlier results below retain their original scope.

Earlier September 20 native update: [one common filename and one waiting state](implementation/2026-09-20-common-name/README.md),
with a regenerated 300-frame native 144/96-DPI catalog and focused before/after captures. Focused checks
pass 3/3; broader validation is 34/36 because PNG clipboard snapshot retention still fails. Narrow
buttons and remaining inspection/layout/qualification work are explicitly open.

This record covers the original review. The [September 20 interaction audit](interaction-review-2026-09-20/README.md)
adds separate current receipts, 309 French native-144/96-DPI captures, a 48-frame direct-review list,
and explicit remaining journey/UX gates. Its passing capture cases do not accept the redesign.

The pack preserves the original English Debug evidence and adds French static and interaction Debug evidence. Captures from different receipts are identified individually; this is not a cross-configuration pixel-parity claim.

| Check | Result | Evidence / limit |
|---|---|---|
| x64 Debug / english-baseline | Passed native capture; 128 PNGs | [Receipt](evidence/build-receipt.json), [native results](evidence/fileops-results.json), original ignored build log absent from this repository |
| x64 Debug / french-static | Passed native capture; 300 PNGs | [Receipt](evidence/french-static/build-receipt.json), [native results](evidence/french-static/native-results.json), original ignored build log absent from this repository |
| x64 Debug / french-interaction | Passed native capture; 9 PNGs | [Receipt](evidence/french-interaction/build-receipt.json), [native results](evidence/french-interaction/native-results.json), original ignored build log absent from this repository |
| Final Debug compilation | 0 warnings, 0 errors | Final build logs retained; tests enabled; earlier English Release run also passed |
| Desktop input | Tab/Shift+Tab, actual target hover and owned mouse capture asserted at 144 and 96 DPI | [Interaction record](evidence/french-interaction/interaction.tsv); warning/lease and foreground/pointer ownership checks precede input |
| Monitor transition | Same HWND moves 144 → 96 → 144 DPI | Captures 109–113; [actual work areas](evidence/french-static/displays.tsv) |
| Runner plan tests | 56 source checks + 1 native inventory check passed | [Source tests](evidence/revision-plan-source-tests.json), [native inventory](evidence/revision-plan-runtime-tests.json); exact interaction case alone requests the lease |
| Test inventory / tooling governance | 7 + 12 passed | [Focused results](evidence/revision-inventory-tests.json); both gallery cases excluded from broad FileOps totals |
| Specification inventory | 0 blocking findings | historical result; original ignored log absent from this repository |
| Tool inventory | 0 blocking findings | historical result; original ignored log absent from this repository |
| Artifact integrity | All PNG hashes/dimensions and local links pass; gallery JavaScript syntax passes | [Integrity result](evidence/revision-artifact-check.json), [manifest](capture-manifest.json) |
| Visual inspection | All published application captures and all three revised concepts inspected | Contact sheets plus full-size risk, dialog and interaction screens; findings remain visible |
| Formatting / whitespace | Changed native sources formatted; diff whitespace checked; [one post-capture line-wrap record](evidence/post-capture-formatting.json) | No baseline reformat required |
| DxUi scope | Guidance only; no library implementation or consumer pin changed | Pinned `b129956d01b725cca131d1369580f61ba3def76b` |

The initial input driver failed foreground acquisition, then exposed coordinate scaling and pointer-release sequencing defects. Those attempts were excluded from the published manifest; the final lane requires successful native assertions before images are accepted. Input cleanup is drained before moving the HWND to the next monitor.

The final revised Release rerun could not build because an independently launched Release application owned the exact output executable (historical build refusal; original ignored log absent from this repository). It was left running. Final French capture coverage uses the validated Debug binary; the earlier English Release capture passed, but it does not qualify the revised source in Release.

The previously run DxUi dependency validator rejected the pre-existing `src/DxUi.vcxproj.user`. This unrelated local file was not modified or deleted. DxUi skills/spec validation passed in the initial review. No library behavior changed here; this is not a capture blocker.

The native runner retains its existing sandbox audit findings (224 old run directories in the retained results). This work does not claim a globally clean sandbox or delete unrelated evidence.

## Reproduction and provenance

Use the two governed commands in the [README](README.md). The final source patch is [retained here](evidence/capture-source-revision.patch); the manifest records source-file SHA-256 values and each receipt's source snapshot. French resource loading is asserted by an actual translated Cancel label. Per-file DPI comes from `GetDpiForWindow`, not a scaled image.

- **english-baseline**: run `20260916T085559Z-81200-4b9ac2f15139453cb017b0aac343aee7`, receipt `cc2fb03b1470797b043fd577b2a3205b83e24c6d2f6ed441848b43100e47055f`.
- **french-static**: run `20260916T132245Z-112164-60aea098d9e2492eaded02eb7511e4de`, receipt `316ddd3f631a0ad6f3cf144697c3f89cc271d7be430bdee69e0207551ae0fec8`.
- **french-interaction**: run `20260916T131714Z-98580-340fc0efb8c049edb9c75d061d2097e5`, receipt `798f4c1f4567f455acaee810355027f9132d14e3ed35f2b4429839c4c242ef3a`.

This is focused harness and UI-review evidence. It does not establish Fresh Full, native ARM64/ASan, provider behavior, comprehensive tooltip/UIA/assistive-technology coverage or performance acceptance. Existing platform qualification remains with its current owners. 200% and Japanese/RTL require future native execution. Capture-driver completion is distinct from the product's open no-clipping requirements.

Native source hashes and the source patch describe the tested capture sources. One later clang-format string-literal line wrap is recorded separately; concatenated literal content and every non-string token were verified unchanged.

## September 19 decision and implementation-plan update

All six product choices are now recorded in the [proposal](proposal.md), authoritative popup/settings/
conflict contracts and [ACTIVE I26 plan](../../Done/UI_FileOperationsReview_2026-09-16.md). The library has a
separate indexed generic-layout plan. This update changes documentation only: no product/library
code, dependency pin, schema, native PNG or prior capture receipt changed. Mockup 06 and its gallery
caption now explicitly describe the user-collapsed state; the accepted initial state is expanded.

- RedSalamander `Get-SpecInventory.ps1 -FailOnFindings`: **0 blocking findings**.
- DxUi `validate-specs.ps1`: **66 Markdown files, 20 authority files, one active plan validated**.
- DxUi `validate-skills.ps1`: **all nine repository skills validated**.
- Changed-document local references: **973 checked, none missing**; all **437 native gallery entries
  are byte-identical** to the previous gallery entries. Application images were not regenerated.
- Diff whitespace checks passed in both repositories.
- DxUi dependency validation still fails on pre-existing `src/DxUi.vcxproj.user` (last modified
  September 13). Formatting check fails on unchanged `src/Controls/DxUi.cpp:33`. Neither file was
  changed by this update; these are repository check failures, not unresolved product decisions.

No new native build, runtime, performance or accessibility pass is claimed by a documentation-only
update. P0 retains fresh baseline/witness work; the September 16 evidence remains historical.
