> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** Reviewed repository commit `eb779e0bd836` (31 October 2026, 22:38 Europe/Paris), `main` after the changelog fold in pull request #198. This file authorizes no implementation, no plan move, and no index edit. Live work stays indexed by [Specs/Plans/WIP/README.md](../Plans/WIP/README.md). Domain specs under `Specs/<Domain>/` remain the behavior authority.

# WIP folder quality — 1 November 2026

## Outcome

The indexed queue is still the place unfinished work lives, and several product defects named there are still in source. The September 2026 priority table is the wrong execution order for a shippable build. A later index rewrite should put the Local Delete context alias first, then hosted qualification, then the compiler and terminal gates that the index headline already names.

This review does not perform that rewrite. It records the evidence so the next owner can edit the index, the stale banners, and the evidence tree without treating this file as a second plan.

## Scope

- Indexed plans: 35 markdown files directly under `Specs/Plans/WIP/`, excluding `README.md`. Status at this commit: 29 ACTIVE, 4 HOLD (H0, H1, H3, H4), 2 DECISION (D0, D1).
- Nested directories that the plan inventory does not index: `FileOperationsUiReview_2026-09-16`, `ViewerSpace_ResponsiveTreemapAndAnalysis_2026-10-08`, `ViewerSpace_D2_Slice0_Microbenchmarks`, and `CodeReview_MasterMerges_2026-10-05`.
- `Notes_Scratch.md` is the human-managed note excluded by [Specs/README.md](../README.md). It was not opened.
- Method: read the index and each plan's status banner; mechanically resolve path citations under `Specs/Plans/WIP/` except that note; re-read the product claims that the mechanical pass marked absent, out of range, or contradictory. No Windows build, no Full, no hosted-log refetch.

## Recommended ship order

This order is a review recommendation. It does not replace the index until someone edits `Specs/Plans/WIP/README.md`.

| Order | Owner | Why it is here |
|---:|---|---|
| 1 | I50 H5, Local Delete context alias | `kDeleteStreamContextSlots` is still 16, and a stream id of 16 or more still uses slot 0. On a machine with 16 or more logical CPUs, two Delete workers can share one `OperationContext`. The plan already says the wrong file is not deleted; the hazard is garbled progress, mis-attributed results, or a crash. |
| 2 | I52, then I56 | Hosted checkout, nightly, ASan, and the six-project acceptance record are still open on their own pages. I56 is the later owner. I52 should not grow a second acceptance list. |
| 3 | I51, with I15 behind it | Ghostty 1.4 admission and the failed local Commands qualification are the live terminal gate. I15 keeps the older paired-evidence obligation. |
| 4 | I48 and I53 | The index headline already names native build policy and the MSVC 14.52 / Release optimization study. Both still have qualification gates. |
| 5 | I18, characterization only, after H5 | Confirmation preferences and typed Copy/Move outcomes are unimplemented. Characterize producers before building the settings page. Do this after the Delete alias, because Copy/Move/Delete trust starts with a Delete that does not share a context. |
| 6 | I0, after a banner rewrite | Writable Drive behavior is already in the tree. First-run PKCE sign-in is the remaining product gap. The July status block still says the plugin is read-only. |
| 7 | I33, selected qualification only | ViewerSpace remains the single owner for its selected enhancements. Growth comparisons and rejected extras stay out. |

Park until a named owner rewrites the banner or the blocking gate above lands: I10 Done admission (implementation is delivered; the Full that covers Tools identity `bae5f9546` was still open at the index's 17 October 2026 check, and this review did not rerun Full), I9 operator evidence, I57 (schema 2 lives on another worktree), I46 (cross-repository theme extraction, not started), I1, H3's deferred features, D0, D1, I4, I7, and H0.

## What is still in the tree

| Claim | Where it still is |
|---|---|
| I50 H5. Delete stream ids of 16 or more alias to slot 0. | `Plugins/FileSystem/FileSystem.FileOps.cpp` lines 13728–13733. The slot test is `streamId < deleteStreamContexts.size() ? streamId : 0`. |
| I50 L33. The MTP fake backend is compiled into the product DLL. | `Plugins/FileSystemMtp/FileSystemMtp.Tests.FakeBackend.cpp` has no `ENABLE_TESTS` guard. `FileSystemMtp.vcxproj` line 146 compiles it with no condition. |
| Find keyboard settle omits `isForegroundWindow`. | `FocusFindRootNavigationForKeyboard` in `RedSalamander/SelfTest/Commands/Commands.Tests.Search.cpp` lines 1250–1269 requires `hasWin32Focus` and `rootNavigationHasWin32Focus`. `FindFilesWindow.h` line 160 still has `isForegroundWindow`. The plan cited line 1199. |
| `FormatHResult` is still a one-line alias. | `RedSalamander/FolderViewInternal.h` lines 171–174 return `FormatHResultMessage`. The plan cited line 307. |
| Google Drive can upload. Cross-root Copy remains refused. First-run PKCE is absent. | `UploadResumable` and `kDriveUploadChunkBytes` are in `Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.cpp`. `oauth2Pkce` is an `authMode` string check. No `code_verifier` or `LaunchInteractiveAuth` in that plugin. `Specs/FileSystem/FileSystem_FileOperations.md` still refuses Google Drive cross-connection and shared-drive `CopyItem` pairs with `ERROR_NOT_SAME_DEVICE`. |
| I18 confirmation page is unimplemented. | The plan's checklist is open. No `ConfirmationPreference` symbol was found. |
| I56 hosted acceptance is open on the plan's own record. | Prepared from `738745a8` on 31 October 2026. The plan records a failed RedSalamander regression run and billing blocks for other repositories. This review did not refetch those logs. |
| Publication may clear staging-only Hidden and Temporary, then reapply source attributes. | `ClearTemporaryAttributes` is `Plugins/FileSystem/FileSystem.cpp` line 4408. `Specs/FileSystem/FileSystem_FileOperations.md` lines 967–970 require the identity-checked reopen. I50 M4 cites `FileSystem.FileOps.cpp` lines 4594–4635, which is `CopyProgressRoutine`. This review did not reproduce an "ends visible" failure. |
| Product `DestroyWindow` on a WIL owner is gone. The performance-test site remains. | `Tests/PerformanceTests2/Performance.Tests.FolderViewInternal.Access.h` lines 772 and 782 still call `DestroyWindow(_hWnd.get())`. |
| Move breadcrumbs on this commit are schema 1. | `RedSalamander/FileOperationMoveBreadcrumb.h` line 16 is `kSchemaVersion = 1u`. |

## What the plans say that this commit does not match

| Plan text | This commit |
|---|---|
| I0 status, last updated 17 July 2026: read-only listing; mutations stubbed; capabilities all false; entry points return `ERROR_NOT_SUPPORTED`. | Resumable upload exists. The closeout checklist still has open download, upload, shortcut, and remote-validation boxes, and the PKCE caveat is the part that still matches. `LaunchInteractiveAuth` is not at `FileSystemMicrosoftDrive.cpp:2454`. |
| I7, verified 2 July 2026: `FormatStatusText` exists at `FolderWindow.FileOperations.IssuesPane.cpp:51`, and both aliases are untouched. | `FormatStatusText` has no match in `*.h` or `*.cpp`. Line 51 of the Issues pane is `IssuesDetailFontRole`. `FormatHResult` remains. |
| H0 and the index evidence section: off-screen selected rows expose SelectionItem and then reject selection queries, in `Common/DxUi/DxUi.Accessibility.cpp`. | That path is not in this repository. In the sibling DxUi checkout read for this review, `SnapshotSupportsGridSelectionItem` returns true when the row is in the snapshot or is selected, and the comment says GetSelection can return rows beyond the visible cache. `get_IsSelected` and `get_SelectionContainer` use that predicate. This review did not diff that checkout against RedSalamander's pinned DxUi commit, so the pin's binary is a separate check. The cited in-repo path and the "visible-only rejection" sentence are stale. |
| I57 W2: current behavior is `kSchemaVersion = 2`. The same file's banner says this main writes schema 1 and the hide exists only in the dirty worktree. | This checkout is schema 1. The banner is the sentence that matches HEAD. The W2 "current behavior" section describes another tree. |
| I5 status: enforcement tasks have not started, and the worker must use `superpowers:subagent-driven-development`. | `CheckFolderViewPerfBudgets` lives in `RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp`. `RequirePerfBudgets` exists on `Tools/Run-AllTests.ps1`. `superpowers` is not a RedSalamander skill. The plan also links a Done plan as if it were still under WIP. |
| H0 "verified correct": `DeleteCopiedSourceEntryForMove` and `sourceDirectoryStillMatches` at `FolderWindow.FileOperations.State.cpp:7302-7498`. | Those names have no production C++ matches. `Commands.Tests.PluginConfig.cpp` asserts the old name is absent from the executor and local File Operations source. |
| I50 H4, revision 2: `core.longpaths` is still missing in `release.yml` and `winget-release.yml`. | Neither workflow file exists under `.github/workflows/`. The row cannot be executed as written. `publish.yml` delegates to the shared RedSalamanders workflow and has no local checkout. |
| File Operations UI review README: I26 contains the ACTIVE implementation plan. | The link target is `Specs/Plans/Done/UI_FileOperationsReview_2026-09-16.md`. The same README records a passing V65 Release Full (2211/0/46) and leaves Final Full open. |
| Index recommended-priority table, 8 September 2026 against `3dcf7ab1`. Next action is I10 Full, then I18. | The table omits I25, I29, I31, I32, I33, I34, I36, I46, I47, I48, I50, I51, I52, I53, I56, I57, and H4. The file's first lines name I53 and I48. The machine index below the table does include the October plans. |

## Adversarial readings

**Release engineer.** The first defect that can corrupt a Local Delete or Recycle on a large machine is H5. Hosted acceptance (I52, then I56) is the gate that tells you the repair survived the runners. Compiler and terminal qualification (I48, I53, I51) are the gates the index headline already admits. Done-path paperwork for I10 does not change the binary a user runs.

**Product owner.** Trustworthy Copy, Move, and Delete come before a new confirmation page and before ViewerSpace exploration. Drive is usable for upload once a refresh token exists; the missing product step is first-run sign-in, which the I0 banner hides by still calling the plugin read-only. Localization (I6) and thumbnail enrichment (I1) are real and can wait behind Delete safety.

**New agent.** Three "do this next" sentences disagree. The opening bullets name I53 and I48. The priority table names I10, then I18. I57 says implementation has not started and priority is high, while its own banner says current main is fine. An agent that obeys the first unread banner will edit the wrong tree or reopen delivered Drive mutations.

**Spec governor.** `Tools/Get-SpecInventory.ps1 -FailOnFindings` checks that each direct `Specs/Plans/WIP/*.md` plan is indexed once, that ranks are unique, and that links resolve. It does not see nested directories, status freshness, or whether the prose priority table covers every ACTIVE rank. A green inventory at this commit is compatible with a 229 MiB unindexed evidence tree and a priority table frozen in September.

**Skeptic.** ACTIVE here means the file has not moved to `Specs/Plans/Done/`. I10, I2, I29, and I32 describe delivered implementation plus an evidence or closeout remainder. Counting 29 ACTIVE plans as 29 unimplemented features overstates the product gap and hides the five defects that are still in source.

## Citation scan

The mechanical pass extracted 1,976 path citations from the WIP markdown except `Notes_Scratch.md`.

| Result | Count |
|---|---:|
| Resolved at the cited path | 1,246 |
| Resolved by a unique basename | 335 |
| Resolved by a unique suffix | 11 |
| Ambiguous basename | 43 |
| Unresolved string | 341 |

A second pass classified line citations that named an identifier: 88 still within a small window of the cited line, 77 present in the same file and moved farther, 211 line citations with no identifier to check, and 23 marked absent inside the cited window. The absent list is a window check. Several rows attach a symbol from a neighboring citation on the same line, so absence in that list is a lead, not a verdict. The semantic table above is the set that was re-read.

Unresolved strings cluster by citer: 51 from I46 (other-repository or proposed future files such as a DxTheme repository), 41 from the nested File Operations UI review, 22 from I50, 22 from I56 (many are local evidence paths under a performance root), and 7 from H0 (DxUi sources that used to live in this repo). One indexed spec anchor fails: `Operation_FileOperations_LaterFeatures_2026-09-05.md` links a Done heading id that does not match the destination heading. The destination file exists.

Screenshots were not opened. The 77 moved symbols were confirmed present in the cited file and were not each re-read at the new line. I3, I6, I25, I34, and the I50 rows other than H5, H4's missing workflow files, L33, and M4 were not re-executed. I46's cross-repository names were not opened in the RedXe or RedPrism checkouts.

## Folder quality

These are recommendations for a later change. This review does not apply them.

1. Rewrite the 8 September 2026 priority section so every ACTIVE rank appears, or delete the prose order and leave the machine table as the only order.
2. Move `FileOperationsUiReview_2026-09-16` out of WIP. It is 4,993 tracked files and about 229 MiB of the folder's about 230.5 MiB. Its README still calls a Done plan ACTIVE. Evidence can live under `Specs/TestRuns/` or `Specs/Reviews/` with a pointer from the owning plan.
3. Give each direct WIP directory an index row, or make the inventory fail on an unowned directory. Today those directories are invisible to the gate.
4. Keep one live owner per overlap: I56 over I9 and over I52's acceptance list, I51 over I15, H4 over I31's ARM64 and ASan deferral, I50 H5 over I36's re-closed rows.
5. Mark I57 HOLD until the schema-2 worktree is the tree being merged. Leave the schema-1 sentence in the banner and remove the W2 section's claim that schema 2 is current on this main.
6. Refresh the status banners of I0, I5, and H0 before anyone executes them. Remove the imported `superpowers` recipe from I5.
7. Allow a plan to stay ACTIVE only with a recheck date. Delivered-but-unarchived work should say what evidence is missing, which I10 already does in the machine table and which the priority table then buries at rank 1.

## Rejected readings

- The queue is empty. H5, the MTP fake, the Find settle predicate, `FormatHResult`, and first-run Drive PKCE are still open.
- Google Drive is still a read-only stub. Resumable upload is in the plugin. Cross-root Copy refusal is specified.
- I57 should be implemented on this main. This main writes schema 1 and shows those breadcrumbs. The hide the plan describes is the other worktree's schema 2.
- H0 can be implemented from the cited `Common/DxUi` path. That tree is gone from this repository, and current DxUi source accepts a selected row that is outside the visible cache.
- A green spec inventory means the folder is current. The inventory does not grade prose freshness or nested directories.
- I50 M4 is a reproduced "hidden file ends visible" bug. The strip function exists in another file at another line, and the normative spec already requires source attributes to be applied again after publication.
