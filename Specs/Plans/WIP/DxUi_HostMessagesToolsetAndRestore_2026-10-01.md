# DxUi follow-ups: registered host messages, vcpkg toolset and pin restore

Status: **ACTIVE** (I31; formerly I27, renumbered 2026-10-09 because the archived File Operations safety plan owns I27). Owner: the RedSalamander build/test maintainer for the remaining qualification and tooling fixes.
Small and independent of the I26 product work.
Implemented on `dxui/update-main-2026-09-30`, which adopts DxUi main `49a9988` (API revision 3) on 4 October after the
3 October master merge had selected `094ca1c`; the qualification tasks below stay open for the adopted pin.
Source: the DxUi plan `MergeAndHostHygiene_2026-10-01` (RedSalamanders/DxUi). Current requirements stay in
`docs/DxUi.md`, `Specs/Build/Build_Toolchain.md` and `Specs/UI/UI_DxUiSharedGrid.md`.

## Why

- **Host messages.** DxUi's private window messages were `WM_APP` offsets on the host window.
  - `RedSalamander.cpp`'s test message `kCommandsSelfTestStartMessage = WM_APP + 0x6A` is DxUi's accessibility action.
  - The unmerged pin-bump branch `dxui/update-main-2026-09-30` reserves DxUi's `WM_APP + 0x06D` in `Common/WindowMessages.h`.
  - DxUi now registers every private message by name (`RedSalamanders.DxUi.<Component>.<Purpose>.v1`) and uses no `WM_APP` value.
- **vcpkg toolset.** vcpkg builds dependencies with the newest MSVC toolset of the Visual Studio installation, while MSBuild links with its default toolset. This repository links compiled vcpkg libraries (curl, 7-Zip, libjpeg-turbo) into its binaries, where a library from a newer toolset than the linker is a known hazard. DxUi now pins its own vcpkg restore to the discovered installation's default toolset.
- **Pin restore.** The governed restore (`Tools/Modules/Build/DxUiDependency.psm1`) clones without Git long paths.
  - DxUi `a0b4934` could not be restored from `D:\RedSalamander` until DxUi kept every tracked path within 150 characters.
  - That budget leaves only 35 characters for a consumer root, and deeper worktrees still fail.

## Tasks

- [x] **Required** when adopting a DxUi pin with registered messages (RedSalamanders/DxUi#41):
  - replace the `PostMessageW(target, WndMsg::kDxUiContextMenuRootHoverChanged, …)` in `RedSalamander/RedSalamander.cpp`
    (about line 5348) with `DxUi::ContextMenu::PostMenuBarHover(target, hoverIndex.value(), pendingSequence)`, and keep its
    failure branch. Otherwise hovering another menu-bar item silently stops switching the open menu.
  - Done: the call posts the hover through DxUi and its failure branch still logs the failed post.
- [x] When adopting that pin:
  - remove the mirrored DxUi values from `Common/WindowMessages.h`: `kDxUiAccessibilityAction`,
    `kDxUiAccessibilityCreateProvider`, `kDxUiWindowHostProcessExitDetach`, `kDxUiContextMenuRootHoverChanged`, and the
    pin-bump branch's `WM_APP + 0x06D` reservation with its `UI_DxUiSharedGrid.md` requirement;
  - state in `docs/DxUi.md` that host windows forward every message to `HandleMessage`, registered ones (0xC000–0xFFFF) included;
  - confirm that no window procedure filters registered messages before forwarding.
  - Done: the four mirrored values are gone; the pin-bump branch's reservation was never merged, so `UI_DxUiSharedGrid.md`
    states the forwarding requirement for registered messages instead. The audit covered all 60 `HandleMessage` call sites
    in 31 files. Their early returns handle `WM_NCCREATE`, `WM_NCDESTROY`, `WM_CREATE`, shortcuts and Tab/Escape routing, and
    three of them skip a fixed list of standard messages (the file-operations popup, the preferences page host and the
    theme-cycle overlay). None filters by message range or drops a registered message.
- [ ] Optionally register this product's own private messages by name as well (`Common/WindowMessages.h`, the self-test
  messages), so that no plugin or library can share their values. Not done: it is optional and independent of the pin.
- [x] Make this repository's vcpkg restore (`vcpkg-install.ps1`, `Tools/Modules/Build/VcpkgInstallSafety.psm1`) use the
  Visual Studio installation the build discovers and its default toolset, as DxUi's `vcpkg-install.ps1` now does.
  Keep the governed custom-triplet rules.
  - Done with an own implementation, because the installer runs before any DxUi restore. A test compares it with DxUi's
    `Update-DxUiVcpkgOverlayTriplet` once the pin is restored. Manifest installs that MSBuild drives, and the inline install of
    the CI workflow, keep vcpkg's own toolset choice.
- [x] Restore the DxUi pin with `core.longpaths`, and consider a sparse checkout that leaves out `Measurements/`,
  `docs/gallery` and `Specs/`. Add a restore test from a deep root.
  - Done in `DxUiDependency.psm1`, with a restore test from a 160-character root, where a kept library file's restored
    path passes `MAX_PATH`. Git cannot enter a checkout directory longer than 258 characters, so a root of up to 186
    characters restores. The restore also compares the API revision the pinned source reports with the one
    `DxUi.lock.json` records.
- [x] Added with this adoption, because the Debug and ASan Debug suites must run unattended: route a failed Debug-CRT or STL
  check away from the CRT's modal dialog in every test process (`Tests/TestSupport/TestSupport.Tests.FailureReports.h`, as in DxUi's
  `Tests/Support/FailureReports.h`), and apply `SEM_NOGPFAULTERRORBOX` to every process of the runner's job.
  `Specs/Testing/Testing_SelfTests.md` states the requirement; `Tools/Tests/FailureReportRouting.Tests.ps1` checks it.
- [x] Added with this adoption, because the ARM64 Debug PR gate failed ViewerSpace's two click-navigation checks once
  dependencies installed again (runs 36935752820, 36951251308 and 36976285319; x64 Debug passed). ViewerSpace created
  its own D3D11 device with the SDK debug layer for both its hardware and WARP attempts, and the `windows-11-vs2026-arm`
  image lacks that optional component, so Debug ViewerSpace had no renderer there and its tiles never grew from their
  zero-size start. Diagnostic run 37067992283 reported renderer stage 1 with `DXGI_ERROR_SDK_COMPONENT_MISSING`.
  ViewerSpace now uses `DxUi::CreateD3D11DeviceWithWarpFallback`, as FolderView does, and run 37072930429 passes the
  scenario on ARM64 Debug. The scenario checks for a renderer before it probes and reports what its probes hit;
  `Specs/Plugins/Plugins_ViewerSpace.md` states the requirement. RedSalamanderMonitor's `ColorTextView` created its device
  the same way, which no CI suite caught because none paints the log view; it uses the helper too, and `--chrome-selftest`
  checks that the log view gets a renderer (`Specs/Core/Core_RedSalamanderMonitor.md`).
- [x] Merged master on 3 October. Its File Operations work had pinned DxUi `19ca44d`, a `codex/multiline-caret-viewport`
  commit at API revision 2, and this branch `271bd54`. DxUi main `094ca1c` contains both `271bd54` and the caret clip
  ported to main (`cc6edae`), with #54 to #57 at revision 3, so `Update-DxUi.ps1 -UpdateOnly` pins it. Master's restore
  change (Git long paths only) yields to this branch's sparse, revision-gated restore, and its cold long-path restore
  test, which expected a restored `Measurements/` file, gives way to the deep-root restore test. The merged tree builds
  in x64 Debug with no warning.
- [ ] Qualify each change with the product suites named in `docs/DxUi.md`, record the pin and its rollback, then move
  this plan to Done. The explicit remaining work below stays open after PR #110 merges.

## Validation already available

- The adopted lock is DxUi main `49a99881cecd37014dc1a92075fd1a610fdfded2`, API revision 3. It replaces `094ca1c` with
  the latest main requested by the developer on 4 October. The original PR title/description named the still older
  `271bd54`; the current lock is authoritative. The latest pin includes the slider touch feedback and subsequent
  control-lifetime, menu-lifetime and interactive-validation fixes; it does not change the consumer API revision.
- [Hosted PR CI run 37127867050](https://github.com/DualTail/RedSalamander/actions/runs/37127867050), on product commit
  `fd1e26373336fab7f963a7badb41a4e89a2f75da`, passed the x64 Debug and native ARM64 Debug build/Suite PR jobs on
  the previous pin `094ca1c`, plus migration guards and formatting. Suite PR is narrower than Fresh Full; these passes
  do not close Full or release gates.
- Current-pin consumer build/Suite PR results are retained in
  [PR #110's checks](https://github.com/DualTail/RedSalamander/pull/110/checks), with exact product and library identities.
  Only successful checks on the updated pin count as current-pin evidence; the Full tasks below remain open.
- [Current-pin consumer CI run 37222498754](https://github.com/DualTail/RedSalamander/actions/runs/37222498754), on
  `3480b08852a45d8023f541b9d5e59f05b5cc3a04`, could not start any job: GitHub reported failed account payments or a
  spending limit. The reported failures are scheduling failures, not executed build/test results. The developer's
  requested merge retains the missing current-pin native evidence as work below.
- Current-pin local migration guards (`Verify-NoSubclassManager`, `Audit-ComctlReportSurfaces`,
  `Audit-VisibleNativeSurfaces`) and pinned clang-format 22.1.3 pass. The exact pin's PowerShell build-matrix validator
  reports 121 native project/solution files and zero six-profile mapping errors; the lock matches API revision 3.
  These are static checks, not a native consumer build or runtime pass.
- [DxUi PR validation run 37218201708](https://github.com/RedSalamanders/DxUi/actions/runs/37218201708) passed all six
  native configurations on `9ebc5b1`. A Git comparison confirms the latest pin has identical `src`, `include`, `Tests`,
  `Tools`, `Build` and PowerShell inputs. The exact latest-main
  [run 37220881885](https://github.com/RedSalamanders/DxUi/actions/runs/37220881885) nevertheless reported an x64 ASan
  Menu Mica-popup appearance timeout; its other jobs were still running during this closeout. A failed-job retry was
  refused while the run remained active. Do not describe that exact-commit run as green without a successful retry.
- The 4 October plan/pin closeout passes tool inventory and `git diff --check`. Spec inventory reports seven existing
  missing evidence links in unchanged `FileOperationsUiReview_2026-09-16` records inherited from master; it reports no
  blocking findings in the updated I31 plan (then numbered I27) or WIP index. This is not a clean repository-wide spec-inventory claim.
- Historical local Suite PR runs at `28f0841bd` recorded Debug and ASan Debug 855/858 passing, and test-enabled Release
  857/859 passing. The PR recorded unavailable symbolic-link privilege and loopback/S3 cancellation failures, with
  baseline reproductions. These earlier runs do not qualify the current `49a9988` pin or the merged product source.

## Remaining work after merge

The developer requested merging the implemented adoption while retaining this work under WIP on 4 October 2026.
Merging PR #110 does not declare the consumer fully qualified. Current requirements remain in the domain contracts
linked above and [Testing_SelfTests](../../Testing/Testing_SelfTests.md).
The closed [release recovery I28](../Done/Build_ReleaseDependencyRecovery_2026-10-02.md) left its October 2 Full failures unfinished; those cases are owned by the Fresh Full consumer qualification below, together with I29, and that recovery record stays closed.

- [ ] **Restore current-pin hosted qualification.** Once the GitHub account can start jobs again, run consumer CI
  against the merged product revision, with the exact `49a9988` pin, and retain x64/native ARM64 Debug build/Suite PR
  receipts. Also retry the latest DxUi run's failed x64 ASan job after the active run ends; investigate a repeated
  material-popup timeout upstream rather than presenting it as a pass or silently changing the pin. Recheck all
  exact-commit conclusions before promoting the pin or declaring release readiness.
- [ ] **Fresh Full consumer qualification.** At the merged product revision and the exact lock above, run
  `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Platform x64 -Configuration Debug`, then the corresponding
  test-enabled Release and ASan Debug profiles. Agree the desktop time before the interactive entries; use the existing
  runner warning/lease and restoration. Archive product/DxUi commit identities, build receipts, per-entry results,
  skips and restoration evidence beneath the governed test root, and retain reviewed receipts under `Specs/TestRuns/`.
  Include ProductUiTests, RedConfigureTests and product integration/focus/clipboard/IME/UIA/DPI/teardown coverage required
  by [docs/DxUi.md](../../../docs/DxUi.md). A PR-suite pass or DxUi-library pass is not a replacement.
- [ ] **Resolve the environment-dependent failures.** Repeat the symlink/reparse tests on a host with the necessary
  privilege and the S3 cancellation scenarios with a working loopback fixture. Preserve any remaining failure with
  candidate/baseline evidence; do not count an unavailable fixture or a skip as a pass.
- [ ] **Complete the separate platform and release handoff.** Retain the existing native ARM64 Debug PR evidence;
  obtain missing ARM64 Release/ASan Debug builds and native runtime qualification for the applicable product suites.
  Keep unavailable native hardware/manual capabilities explicit. Debug CI and cross-builds alone do not establish these
  gates. Record required paired resource/performance observations for the affected product scenarios.
- [ ] **Serialize cold DxUi source restore across platforms.** Address
  [review 4161248136](https://github.com/DualTail/RedSalamander/pull/110#discussion_r4161248136).
  `Restore-RSDxUiDependency` currently holds a platform-specific dependency lease while both platforms share
  `source/<commit>`. `Install-RSDxUiPinnedSource` can remove that shared path after its own failed clone without proving
  that it created the directory. Use a checkout-specific lease or invocation-owned staging/publication, delete only
  owned storage, and add simultaneous x64/ARM64 cold-restore and failed-clone regression cases. Until repaired, serialize
  cold restores within one consumer checkout.
- [ ] **Confine generated overlay writes against reparse points.** Address
  [review 4161248214](https://github.com/DualTail/RedSalamander/pull/110#discussion_r4161248214).
  The new vcpkg overlay path uses lexical confinement before `WriteAllBytes`. Use the existing
  `Resolve-RSGuardedRepositoryPath` and `Publish-RSGuardedRepositoryFile` helpers for the directory and final file; add
  directory-junction and final-file-symlink cases proving external targets remain unchanged. Follow
  [Core_SharedHelpers](../../Core/Core_SharedHelpers.md) rather than adding another path utility.
- [ ] **Keep failure-report probes inside the selected profile lease.** Address
  [review 4161248251](https://github.com/DualTail/RedSalamander/pull/110#discussion_r4161248251).
  `FailureReportRouting.Tests.ps1` currently enumerates and launches other built configurations while the runner locks
  only its selected profile. Pass the selected receipt-verified profile, probe only it, declare executable/runtime
  artifact reads for `tooling.pester`, and test a permitted concurrent build of a different profile. Qualify other
  profiles through separate coordinated runs; meanwhile avoid parallel cross-profile rebuilds during these probes.
- [ ] **Record rollback and closeout.** The prior consumer was `fbe113a0f747f89af32d6fb337e4e67b9f5bdd6b`, pinning
  `19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1` at API revision 2. A rollback must restore the compatible consumer code,
  restore tooling and lock together, then rebuild; reverting the lock alone leaves revision-3 API calls. After the
  required fixes and qualification pass, reconcile their normative specs and evidence, move this plan to Done, and
  update both indexes. Preserve or explicitly route the optional product-message registration task above.
