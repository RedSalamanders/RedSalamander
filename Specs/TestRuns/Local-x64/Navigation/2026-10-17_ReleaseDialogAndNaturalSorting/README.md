# Release acceptance: dialog keyboard focus and breadcrumb natural sorting

Date: 2026-10-17. Scope: test-enabled x64 Release, focused canonical runs on this Windows host.
The user closed the independently running Release instance and explicitly authorized building and testing both fixes.
No independently launched application was terminated. The Commands run used the existing visible warning and desktop lease.

| Scope | Result | Native run | Build receipt |
| --- | --- | --- | --- |
| FileOpsFamily_DecisionMirror: seven mirror cases plus setup/cleanup | 9 passed, 0 failed, 0 skipped | 20261017T162800Z-80208-c502e7dfb5484d5d89ea79eb3fcd6d7b | 4022a40b1a3b7a4443c1dd1ea29afd9bf3df56449260acdadd3dd78e10a40da4 |
| cmd_pane_fileops_decision_mirror_real_focus and cmd_navigation_sibling_natural_sorting | 2 passed, 0 failed, 0 skipped | 20261017T163444Z-107504-ef07845e7c224789a110c4039566a6fa | 479a733e2a272305e7b9ef5a166e51a37c90f9b9b14599f8eaebd9d1ed14c2a1 |

Both Release builds completed with zero warnings/errors (4m34s and 3m07s respectively).
Both runs have zero flaky/regression/isolation-suspect classifications and zero repair attempts.
Each native scope has its own build receipt; the receipt source snapshot equals its original final workspace snapshot.
The common source snapshot is 6668705216482cbe5208ba6228dcee626da9dc5140873addcaf102d0c2feaecc,
on HEAD a48e7d7880cb37ea6898d462ae188a6213a5f3f8 plus the bound dirty worktree.
The dependency remains pinned DxUi e0260dde04695f0794eed79708030798b7575e0c.

The dialog group checks immediate Tab/Shift+Tab/arrows, safe Enter/Escape behavior, foreground focus from either
pane, restoration, repeat suppression and F6 opt-out across overlap/ingress/compare/consecutive/queued prompts.
The real-input case checks actual foreground and keyboard ownership before every SendInput batch, then immediate
Tab/arrows/F6/Escape behavior and preservation of the cancelled delete target. No typeahead delay is used.

The sorting regression calls the real warmed sibling query used by both ordinary and full-path dropdown builders.
It covers case, digit runs, leading zeros, Unicode/dot names, excluded files, empty/exact-99/1000-folder sets,
current-folder retention and cache-only enumeration behavior. All 64 Release timing samples remain in the raw
Commands trace. This is functional Release evidence; no paired Release performance conclusion is claimed.
The separate [Debug baseline/candidate packet](../2026-10-17_NaturalSiblingSorting/README.md) retains the old-order
failure and paired same-fixture timings.

Original native and aggregate results, traces, workspace snapshots, build logs and complete build receipts are
retained separately per run. Phase bindings preserve original diagnostic paths and SHA256 hashes; root
artifact-bindings.json verifies the byte-preserving archive copies. The exact archive has -text attributes.
All six changed C++ files match the passing Debug hashes in tested-code.json. Subsequent edits only close the two
plans, add exact Done-path admission/tests and inventory metadata, and preserve this evidence.

Commands used, with RSBuildEnableTests=true:

```powershell
.\Test-Changes.ps1 -Suite FileOps -CaseFilter FileOpsFamily_DecisionMirror -Interactive -BaseRef HEAD -Configuration Release
.\Test-Changes.ps1 -Suite Commands -CaseFilter 'cmd_pane_fileops_decision_mirror_real_focus,cmd_navigation_sibling_natural_sorting' -Interactive -BaseRef HEAD -Configuration Release
```

This closes the bounded x64 Debug/Release fix plans. Full, ASan, native ARM64, assistive-technology qualification,
broader provider/desktop/platform coverage and the separate File Operations post-merge program remain separate.
Each run reported 41 existing surrounding-sandbox disk-audit issues; neither run is sandbox-audit closeout.

Final closeout checks: 302 Pester tests passed, 0 failed/skipped/pending across exact Done admission,
tooling inventory/governance, documentation drift and native harness source contracts. Both new admission
tests also reject an unreviewed sibling date. The Tools inventory classifies all 197 discovered files
with zero blockers; the specification inventory has zero blocking findings. git diff --check passes.
The final Release executable still matches the Commands receipt, and all six C++ hashes match both
Debug and Release tests. closeout-checks.json records the tested policy/source-contract inputs.

Archive validation passes for all 22 packet files in explicit pre-stage mode and all 13844 checked-in files in whole-inventory mode.
