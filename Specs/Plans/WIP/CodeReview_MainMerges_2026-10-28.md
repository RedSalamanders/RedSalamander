# I50 Review of main merges 2026-10-18..2026-10-30

Status: ACTIVE. Review record and remediation routing. No product code has been changed by this plan.
Owner: next engineering owner chosen by the product owner.

- Revision 1 (2026-10-28) reviewed `0f6cd8ddb..8ea956a7e`.
- Revision 2 (2026-10-30) reviewed the follow-on merges `8ea956a7e..8f0563355`. It re-checked every
  revision 1 finding at `8f0563355`, added the revision 2 findings (H5, H6, M35 onward, L50 onward, S13
  onward, Q15 onward) and updated the verdict, the CI diagnosis, the slices and the questions.

Drift check (run first):

```powershell
git diff --stat 8f0563355..HEAD -- RedSalamander Plugins Common Tools Tests .github build.ps1 Build Directory.Build.props Directory.Build.targets
```

Any in-scope file that changed after `8f0563355` must be re-read against the cited excerpt before its
slice is executed. Line numbers:

- Revision 1 findings (H1-H4, M1-M34, L1-L49) cite `8ea956a7e`. Their files did not change before
  `8f0563355`, except where the revision 2 status table gives a new location.
- Revision 2 findings cite `8f0563355`.

## Scope

| Merge | Content | Size |
|---|---|---|
| PR #142 (`5c9292498`, 2026-10-26) | Source snapshot `397e15adf` of DualTail up to `582d644be` (about 2026-09-19..10-26: File Operations UX/reliability I35, Issues pane, decision mirror, leftover cleanup, ViewerSpace workspace/analysis, RedMonitor extraction, MSVC 14.52 + Release LTCG, DxUi API 4, tooling) plus 10 follow-ups (hosted MSVC, capture WinRT, ViewerSpace scan queue, ViewerWeb profile/lifetime) | 25,590 files; production code about 304 files, +65k/-12.5k lines (renames detected, CR ignored); tests about +110k; `Specs/TestRuns` 19,500 evidence files |
| PR #144 (`e101eac66`) | Agent/skill audit record (docs) | 6 files |
| PR #143 (`fbcc9fc0d`) | Hosted capture admission, S3 fixture latency (test-only) | 8 files |
| PR #145 (`8ea956a7e`) | PowerShell formatter tooling, agent guidance/skill trims, spec IA checks | 45 files |
| PR #147 (`b4ff85bea`, 2026-10-29) | Revision 1 of this plan | 3 files |
| PR #148 (`0ed4d6ff1`) | Ghostty 1.4 runtime with locked Zig 0.17, candidate review/overlay v2, native qualification workflow (I51) | 62 files, +6.7k |
| PR #146 (`28e70e694`) | Hosted CI repair (I52): deferred plugin/viewer lifetimes, directory-cache invalidation, File Operations pool sizing, journal-less index rebuild, service pipe contention, F24 foreground handoff, viewer restoration observation, TestRuns archive migration | 690 files (about 600 evidence) |
| PR #149 (`632dc13b2`) | Shared native build policy (/Wall without /WX, /ZI, Release ICF), short DxUi dependency paths, FTP passive fallback fixtures (I53) | 63 files |
| PR #151 (`2ec73448a`) | Native ARM64 Zig build of the Ghostty runtime with bounded children | 8 files |
| PR #138 (`9c99afd3b`) | Dependency batch (curl 8.23, sqlite 3.54.1, 7zip 26.5, libjpeg-turbo 3.2.0#1, vcpkg 2026.10.16) and changelog fold-on-main tooling (I54) | 32 files |
| PRs #152-#155 (`8f0563355`, 2026-10-30) | CI contract repair, I55 CI organization plan, I54 closeout, changelog fold | 25 files |

Revision 2 delta, excluding `Specs/TestRuns` and plan prose: about 150 files, +11.6k/-1.3k lines.

- About 1,000 lines are production C++.
- About 800 are build and CI.
- About 3,500 are Tools and Pester.
- About 2,500 are test harness and self-test cases.

Reviewed (revision 1): all production C++ (RedSalamander, Common, Plugins, RedConfigure), build/CI/installer,
Tools modules and entry points, test-harness code that gates qualification. Not reviewed in depth:
test-case bodies (about 110k lines), evidence under `Specs/TestRuns`, plan prose under `Specs/Plans`,
the external pinned DxUi package, translation quality.

Reviewed (revision 2):

- The whole non-evidence delta. That includes the Ghostty patch compiled into the shipped DLL, the
  self-test harness and case diffs (checked for weakened assertions), and the dependency bumps (checked
  against the APIs the code uses).
- The scheduled-run logs at `632dc13b2`, `28e70e694` and `8ea956a7e`, and the ARM64 Debug validation
  artifact of run 41030069121.

Method (revision 1): 12 area reviewers (Sonnet, read-only) with a shared brief, then lead
verification. Method (revision 2): 7 area reviewers (Sonnet, read-only), each also re-checking the
revision 1 findings in its area. The lead re-read every High and Medium finding and every FIXED
claim, and read the CI logs directly. Every finding below carries one verification label:

- **CONFIRMED**: the lead re-read the cited code and the trigger chain.
- **AGENT**: the reviewer confirmed it by reading; the lead did not re-trace it.
- **PLAUSIBLE**: the mechanism is confirmed but the user-visible effect needs a runtime check.

"New" means introduced in the window; "pre-existing" means present at `0f6cd8ddb` (still a production
defect, listed because the review covered the surrounding code).

## Verdict (revision 2, `8f0563355`)

`main` at `8f0563355` is **still not production-ready**.

The follow-on merges fixed four revision 1 findings: H2, M29, M33 and M34. Six more are partly fixed:
H4, M27, M30, M31, M32 and L49 (see the status table). Hosted x64 failures dropped from 17-18 cases to
1-2. The merges did not touch H1, H3 or any revision 1 Medium in product code other than M29 and M33.

They also introduced two new High defects and moved qualification backwards on ARM64:

1. **New data-path race (H5).** The Local provider's pool grew to 17 workers. On machines with 16 or
   more logical CPUs, worker 16 and worker 0 share one Delete operation context, so a multi-item Local
   Delete runs two workers on the same non-thread-safe state. The fix is one constant.
2. **No qualification at any merged identity (X1, H6).**
   - Nightly Native 41030069121 at `632dc13b2`: x64 Debug fails 1 case, x64 Release 2.
   - ARM64 Debug and Release fail 970 each. 968 of those are the whole Commands suite: the new F24
     foreground handoff never succeeds on the hosted ARM64 image (H6).
   - The weekly ASan x64 suite still runs into the 6-hour job limit (M40).
   - The Nightly Flake lane cannot finish within its own timeout (M43).
   - An intermittent access violation after the combo-host viewer cases is now visible on both
     architectures (M36).
3. **Release path.** H3 (MSIX) and the H4 checkouts in `release.yml` and `winget-release.yml` are still
   open, and release builds run on an unqualified runner image (M42).
   - The owner's I55 decisions change the target: ZIP-only publication, and automatic test-enabled
     Release on native x64 and ARM64 before production Release.
   - That makes H3 and M42 a choice between fixing and retiring (Q15, Q22).
4. **Still open, as at revision 1:** H1 (unrouted host services) and 27 revision 1 Medium defects, 4 more
   partly fixed. Many have direct user impact (M1, M3, M4, M9, M12, M13 and others).

No data-loss path was found in the revision 2 delta. The ViewerPluginManager deferred-load change does
not create two module instances or a use-after-unload. The directory-cache invalidation retention does
not loop.

### Verdict (revision 1, `8ea956a7e`, kept for the record)

`main` at `8ea956a7e` is **not production-ready yet**. No data-loss path was found in the Local
provider's delete, overwrite, Move or Recycle code, in the leftover-cleanup Delete, or in the File
Operations result reducer and conflict broker. What blocks a production claim:

1. **No green qualification at the merged identity** (X1). The scheduled Nightly Native Suite (all
   four profiles), weekly ASan, Nightly Flake Convergence and Ghostty status workflows are red on
   `main`. The merged I47 (RedMonitor) and I48 (MSVC 14.52) plans still record "final six builds,
   three Fresh Full profiles, measurements and closeout remain pending". The last recorded local
   Debug Fresh Full (2026-10-24) failed (Commands 951/2/6, FileOps 278/4/40, ViewerPE interactive).
   Most of the nightly red comes from the hosted runner (1024x768 display, Windows Terminal owns the
   foreground) and harness defects (H2, M28, M31, M32), not the product. Three product failures
   reproduce there and also failed at the baseline (M29, M30, M33). An in-flight branch,
   `ej/ci-runtime-repair-20261028` (plan `HostedCiRepair_2026-10-28.md`, not on `main`), already owns
   the hosted CI repair; route CI items there instead of duplicating it.
2. **Three High defects**: plugin host services that return success without doing anything when
   called off the UI thread (H1); MSIX packaging that always fails (H3); a repository that no longer
   checks out on default Windows git, which breaks the release and several scheduled workflows (H4).
3. **34 Medium defects**, several with direct user impact: ViewerSpace Delete of a folder,
   silent loss of saved popup/Issues settings, Hidden attribute lost on Overwrite, sticky MTP
   fallback that drops atomic creates, lost-wakeup hangs, a running Move shown as interrupted.

## Status of revision 1 findings at `8f0563355`

Only changed items are listed. Every other revision 1 finding is **still open**: its cited files did
not change in the window, or the area reviewer re-read it at `8f0563355` and found it unchanged. That
covers H1, M1-M28 except M27, and L1-L48 except L40, L44 and L49.

| ID | Status | Evidence at `8f0563355` |
|---|---|---|
| H2 | FIXED | `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:1017-1022` admits a focus child on another thread under the saved foreground root. `:1246-1293` evaluates the read-only "already restored" observation before any write predicate. Writes still require the same pid and thread (`:1261`). Hosted ARM64 proof: run 41030069121 logs `parent observes exact saved foreground and keyboard focus without restoration writes` for the case that failed at revision 1. |
| H3 | OPEN; severity depends on Q15 | `Directory.Build.targets:324` still conditions `ValidateMsvcToolsVersion` only on `RSIsFirstPartyProject`. MSIX is opt-in in `release.yml:212-215`, and I55 §9 (owner-selected, not implemented) publishes no MSI/MSIX. |
| H4 | PARTIAL | `core.longpaths` is now set before checkout in `build-reusable.yml`, `ci.yml`, `nightly-flake.yml` and the three Ghostty workflows; Ghostty runtime status is green since 2026-10-29. Still missing in `release.yml` (5 checkouts), `winget-release.yml` (2), `terminal-engine-gate0.yml` (2, sealed by hash) and `terminal-engine-round4-closeout.yml`. No policy test. Long tracked paths are unchanged: 611 exceed 260 characters under a 34-character runner prefix. |
| M27 | PARTIAL | An empty `-NonInteractive` plan now finishes NOT_EVALUATED (`Tools/Run-AllTests.ps1:1496-1503`, `:2371-2373`). A non-empty legacy suite with dropped desktop entries still reports PASSED (`:2374-2376`). |
| M29 | FIXED (code; hosted journal cases pass) | The journal-unavailable rebuild seeds by tracked-root traversal (`Common/LocalSearchIndexCore.cpp:3399`). New residuals are L59 and L60. |
| M30 | PARTIAL | `Query` now retries a busy pipe for up to 30 s and can be cancelled (`Common/SearchServiceBroker.cpp:1550`, `:3513`). `GetStatus`, `RequestRebuild` and `RequestCompact` keep the 750 ms window (L58). The hosted x64 Release nightly still fails `search_service_multi_client_and_rebuild_control` through its new fixture (M44). |
| M31 | CHANGED (likely fixed, unproven) | `nightly-flake.yml:43-46` and `:64-66` prepare MSVC and the pinned DxUi source before attestation. `Test-RSBuildReceiptForArtifactUse` still returns a bare `$false` (`Tools/Modules/Build/BuildEvidence.psm1:993-1018`). No flake run has reached a verdict (M43). |
| M32 | PARTIAL | The upload is guarded and uses `if-no-files-found: error` (`nightly-flake.yml:108-115`). It still uploads the whole sandbox without the reusable workflow's exclusions. |
| M33 | FIXED (code) | `RedSalamander/FolderWindow.FileOperations.State.cpp:6641` sizes the pool at 4-16 transfer workers plus one reserve. Spec: `FileSystem_FileOperations.md:802-805`. The same change in the Local provider introduced H5. |
| M34 | FIXED | The inline install passes the overlay triplet, plus `--overlay-ports` on ARM64 (`build-reusable.yml:382-401`). The cache key covers `Build/vcpkg-ports/**` and the installer modules (`:235`). `VcpkgInstallSafety.Tests.ps1` executes the step. |
| L40 | By design (withdrawn) | Duplicate-text fragments are folded once; regression case added in `6d4c81893`. |
| L44 | PARTIAL | The `vcpkg.json` homepage typo is fixed. The two `...Tests.OwnedShutdown.cpp` files still share one name. |
| L49 | CHANGED (test-side cause fixed, no repeat evidence) | ImgRaw completion now waits for an applied final result (`ViewerPE.Tests.Runner.cpp:5111`, used at `:5466`, `:5740` and `:6791`), and all four hosted profiles passed once. The ARM64 `/GL` hypothesis was not tested. |
| X1 | OPEN | See the revision 2 verdict and the CI diagnosis below. |

## Revision 2 findings (new in `8ea956a7e..8f0563355`, or newly exposed by it)

Line numbers are on `8f0563355`. "New" means introduced by the revision 2 merges. "Exposed" means
present before, but first visible or first reachable through a revision 2 change.

### High (revision 2)

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| H5 | Local provider Delete | new | CONFIRMED (mechanism); PLAUSIBLE (exact symptom) | The pool now has 4-16 transfer workers plus a reserve, so up to 17 workers with stream ids 0-16, but the Delete per-stream context array kept 16 slots and maps any id of 16 or more to slot 0. On a machine with 16 or more logical CPUs, workers 0 and 16 run Delete steps on the same `OperationContext` at the same time. That context is not thread-safe: item and progress arenas whose paths are handed to the host callback, the file-work call state, traversal counters, `objectBinding` and `completedItems`. A multi-item Local Delete or Recycle can report garbled or freed paths, mis-attribute item results, or crash. The deletion targets come from each worker's own `PathInfo`, so a wrong file is not deleted. The commit updated the Copy slots (`:7946`) and the throttle slots (`:1008`) but missed this array. | `Plugins/FileSystem/FileSystem.FileOps.cpp:13728-13733` (16 slots, alias to 0), `:429` (pool size), `:669-673` (stream id = worker index); duplicate literal `kMaxWorkers = 16u` at `:6826` |
| H6 | Qualification (test harness) | new | CONFIRMED (CI log) | The F24-hotkey foreground handoff never succeeds on the hosted ARM64 image, and Commands isolation now hard-requires it. Every ARM64 Commands case fails before it starts: Debug and Release both fail 968 of 968, against 892 passed / 58 failed at the base. ViewerSpace "activates only its owned window" also fails 14 times per ARM64 profile. The foreground belongs to a system UWP host (`WWAHost`, `Windows.UI.Core.CoreWindow`). The failure text omits the existing `ForegroundAdmissionDiagnostics` (stage, error, held key, sent), so the failing stage cannot be told from the log; the 968 cases take 77.5 s, which rules out the 500 ms dispatch timeout. This is fail-closed tightening, not a weakened assertion, but ARM64 has no Commands coverage at any merged identity. | `Tests/TestSupport/TestSupport.Tests.DirectedSelfTestInputWarning.h:126-238`; `RedSalamander/SelfTest/Commands/Commands.Tests.Settings.cpp:9415-9440` (no diagnostics at `:9436`); `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:15733`; run 41030069121 (ARM64 Debug/Release), 40941503630 |

### Medium (revision 2)

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| M35 | Batch Rename | new | CONFIRMED (call chain); PLAUSIBLE (runtime) | `UpdateTaskUi` disables the whole Rules/Manual `RadioButtons` group, not its items, whenever a preview or target rebuild is pending or running. `RequestPreviewRebuild` calls it on every rule keystroke, and `SwitchMode` reaches it through `RebuildPreview`. Selecting Manual with Space disables the group under the focused radio, and DxUi prunes focus under a disabled ancestor and does not restore it. Mouse clicks during the 150 ms debounce plus the preview are ignored with no visual cue, because the radios keep painting as enabled. UIA reports enabled children inside a disabled group. | `RedSalamander/BatchRenameWindow.cpp:4291-4298`, `:3455-3460`, `:3739` |
| M36 | Viewer plugins / harness | exposed | CONFIRMED (CI log); NOT DIAGNOSED (cause) | Viewer harness children exit with `0xC0000005` after all of their checks passed. ViewerPE combo host, ARM64 Debug: 4 of 12 churn cycles on 2026-10-30. ViewerWeb combo host, x64 Debug at `8ea956a7e`: 2 times. The child last prints `ViewerPETests passed.` or `window close succeeds`, so the crash is in teardown or module unload after UIA was used on the DxUi combo host. Only the ViewerWeb fixture now runs the plugin quiet point (Shutdown, then CanUnloadNow) before `FreeLibrary` (`:3791-3850`). ViewerPE, ViewerText, ViewerSqlite and ViewerVLC export neither, so the product host also frees them at once. No dump is retained, so fixture-side versus product-side is undecided (Q17). | `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:3358-3470`, `:3735`; `RedSalamander/PluginModuleLifecycle.cpp:17-80`; runs 41030069121 (ARM64 Debug artifact), 40845520699 |
| M37 | Terminal runtime (shipped patch) | carried into 1.4 | CONFIRMED (order); PLAUSIBLE (reach) | The patched Kitty `addPlacement` checks the count and byte limits against a placement count that still includes garbage (pruned) placements, and reaps them only after admission. Upstream reaps first. Once a screen holds 4096 retained placements, every new image placement fails with `OutOfMemory` until a delete or erase: images silently stop working in long sessions. | `Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch:454-471` |
| M38 | Terminal runtime build | new | CONFIRMED | The innermost build `finally` removes the reproducible cache without `-PreservePrimaryError`. When Zig's `maker.exe` stays mapped beyond the 20 x 250 ms retry window after a build failure or timeout, the cleanup exception replaces the real failure. This is the failure mode already fixed for the fetch cache at `:1069`. | `Tools/TerminalEngine/GhosttyRuntimeLifecycle.psm1:1418-1429` |
| M39 | Terminal runtime review | new | CONFIRMED (mechanism) | The candidate "security review" findings are hard-coded constants for earlier upgrades' commits, filtered to those whose commits are all in range. The schema now accepts an empty list (`minItems` 0). A review of a new upgrade can therefore pass with no disposition for commits it categorises as `clipboard-paste`, `kitty-image` or `snapshot-serialization`. | `Tools/TerminalEngine/GhosttyCandidateReview.psm1:311-317`; `Specs/Terminal/GhosttyRuntimeCandidateReview.schema.json:179` |
| M40 | CI | pre-existing | CONFIRMED | Long build and suite steps have no step timeout, and only the job's 360-minute cap applies. ASan x64 (run 40941634645) was cancelled at 6 h 01 min again, and the ARM64 Debug dependency batch build hit the same limit. A hung step holds a scarce ARM64 runner for six hours and leaves no diagnosis. The flake job restores and bootstraps the DxUi source with no cache on every run. | `.github/workflows/build-reusable.yml:52`, `:454`, `:614`; `nightly-flake.yml:64-66` |
| M41 | CI | pre-existing | CONFIRMED | `ci.yml` cancels in progress per `github.ref`, so a quick second merge to `main` cancels the gate of the first. Six of the ten `main` push runs since 2026-10-29 ended `cancelled`, so those merged commits (including `632dc13b2`, `2ec73448a` and `9c99afd3b`) never got a complete result. | `.github/workflows/ci.yml:19-21` |
| M42 | Release | pre-existing | CONFIRMED (config) | `release.yml` passes no `native_runner`, so release builds run on `windows-latest` with a Chocolatey Visual Studio install. The ARM64 release is cross-built on an x64 host. CI qualifies only `windows-2025-vs2026` and `windows-11-vs2026-arm`. I55 §8 makes this the target design (Q22). | `.github/workflows/release.yml:35,72,166,205,398`; `build-reusable.yml:14,84` |
| M43 | CI | pre-existing | CONFIRMED (CI log) | The Nightly Flake lane can never finish. It runs `-Suite All -SelfTestRepeat 5`, while one Suite All pass takes 3-4 h on hosted runners. Run 40941643057 "timed out after 210 minutes". Scheduled run 41037399793 was cancelled after 50 min (cause not recorded). | `.github/workflows/nightly-flake.yml:106` and its run step |
| M44 | Search self-test | new | PLAUSIBLE | The new busy-pipe fixture waits with `WaitNamedPipeW(pipeName, 5000)`. That call fails at once when no instance is listening in the gap after the previous client was served, so the fixture is racy. It failed the hosted x64 Release nightly with "Busy-pipe fixture did not find the isolated service listener", and the message omits `GetLastError`. | `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.Tests.Cases.SearchAndIndex.cpp:12577` |

### Low (revision 2)

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| L50 | Plugin lifetime | new reach | CONFIRMED (path); PLAUSIBLE (ViewerSpace effect) | Creating a viewer or Terminal on a deferred entry now sweeps, and the sweep calls `RedSalamanderPluginShutdown` before `CanUnloadNow`, on every user retry while an old instance lives. Before, only Discover and the configuration APIs did this. ViewerSpace's Shutdown latches `g_scanSchedulerShutdown` permanently. If the same mapping is ever reused (another module reference kept it mapped), every later scan ends Canceled. | `RedSalamander/ViewerPluginManager.cpp:237`, `:287`, `:1044-1053`; `PluginModuleLifecycle.cpp:28-38`; `Plugins/ViewerSpace/ViewerSpace.cpp:1808-1814`, `:9163-9168` |
| L51 | Plugin lifetime | new | CONFIRMED | `ViewerPluginManager` has no UI-thread assertion, although Create can now load and free DLLs. `FileSystemPluginManager` asserts. All current callers are on the UI thread. | `RedSalamander/FileSystemPluginManager.cpp:354` |
| L52 | Plugin lifetime | pre-existing | CONFIRMED | The file-system manager did not get the viewer repair: deferred placeholders keep their stale `disabled` flag, so the Plugins UI can show the wrong state. `DisablePlugin` then appends a duplicate id. Its configuration APIs return BUSY without sweeping. | `RedSalamander/FileSystemPluginManager.cpp:1188-1190`, `:572-575` |
| L53 | Directory cache | new | PLAUSIBLE | The retained invalidation reposts a follow-up load with no minimum interval. A visible folder with a continuously growing file re-enumerates back to back, throttled only by FolderView's 200 ms debounce. | `RedSalamander/DirectoryInfoCache.cpp:1754-1760` |
| L54 | Batch Rename | new | CONFIRMED | `CancelPendingPreviewRebuild` and the synchronous `ExecuteRename` failure returns skip `UpdateTaskUi`, so the mode selector stays disabled after the error until the next edit. | `RedSalamander/BatchRenameWindow.cpp:3895`, `:3917`, `:3934`, `:3942`, `:3987` |
| L55 | Batch Rename | pre-existing | PLAUSIBLE | Rename is not gated on a pending rebuild. A click within 150 ms of a keystroke renames per the previous preview and silently drops the last edit. | `RedSalamander/BatchRenameWindow.cpp:3881-3896`, `:4306` |
| L56 | Popup | pre-existing | PLAUSIBLE | Two child-slot moves (retire by `std::move`, then `stable_sort`) bypass `TakeChild` and `NotifyRetainedTreeMutation`, so the UIA child order can go stale after a reorder-only reconcile. | `RedSalamander/FolderWindow.FileOperations.Popup.cpp:8283`, `:8425` |
| L57 | Popup | pre-existing | PLAUSIBLE | The menu-driven Custom speed path does not re-check host lifetime or `IsWindow` after its nested pump; the sibling path does. The popup is not disabled under `ShowModal`, so a second prompt can open. | `Popup.cpp:13180-13204`, `:2530-2595` |
| L58 | Search service | pre-existing | CONFIRMED | Status, rebuild and compact requests keep the 750 ms busy-pipe window, so they fail with `0x800700E7` while a query holds the pipe. A user rebuild shows an error. | `Common/SearchServiceBroker.cpp:36`, `:3255-3259`, `:3819-3822`, `:3885-3888` |
| L59 | Search index | exposed | CONFIRMED | Traversal rebuild aborts on the first directory it cannot list (only not-found is skipped). `EnsureReady` then fails and the in-process index falls back to a live scan. | `Common/LocalSearchIndexCore.cpp:2745-2755` |
| L60 | Search index | new | PLAUSIBLE | A journal-less volume re-walks its whole tracked root on every `EnsureReady` (every query), where it previously ran one MFT enumeration. | `Common/LocalSearchIndexCore.cpp:3382-3399`, `:5181`, `:5784` |
| L61 | Search service | new | CONFIRMED (latent) | `operationDeadline == 0` means "none" in the `std::min` guard but "expired" in the retry loop. Every current caller sets a deadline. | `Common/SearchServiceBroker.cpp:1550-1562` |
| L62 | Terminal runtime tooling | new | CONFIRMED unless noted | (a) The diagnostics watchdog (20 min) is shorter than the per-child bound (30 min), so the stall snapshot never appears. (b) The runtime identity embeds the host-specific Zig archive hash, so x64- and ARM64-host builds of one DLL never match. (c) The `aro` archive is an on-demand Codeberg tarball with no mirror (PLAUSIBLE). (d) The notice gate does not require the new `aro`, `pixels` or `wuffs` notices. (e) `jsonschema` installs unpinned transitive packages. (f) The HKCU WER `LocalDumps` parent key is created and never removed; a partial Zig extraction is accepted when only `zig.exe` exists; a mismatched download never self-heals. (g) The diagnostics artifact uploads the whole scratch folder, including TEMP. (h) The `codeFiveRetry*` evidence fields are constant. | `ghostty-runtime-native-qualification.yml:207`, `:242`, `:274`, `:332-343`, `:424`; `GhosttyRuntimeLifecycle.psm1:211`, `:896`, `:1171`, `:1244`; `GhosttyRuntimeLock.v1.json:345` |
| L63 | Changelog tooling | new | AGENT unless noted | (a) No workflow runs `Submit-ChangelogFold.ps1`; when a fold PR is already open the run exits 0, indistinguishable from success. (b) Every fold checks out all 27,840 files (792 longer than 205 characters). (c) The fold runs the fetched tools but validates with the caller's `Test-Changelog.ps1`. (d) The exact-fold check rejects valid histories spanning two or more folds, because it expects one-shot order (reproduced by the reviewer). (e) A cancelled push run skips that push's changelog check (M41). (f) An existing fold branch is reused without a content check. (g) Archive inventories drop non-ASCII names because they do not use `-z`. (h) `-ZigToolchainTransition` is self-attested. (i) `HostedTestDesktop.psm1` gates only on `GITHUB_ACTIONS`, so a self-hosted runner or `act` would change a developer's display, and its restore exception masks the action's own error. | `Tools/Submit-ChangelogFold.ps1:64-90`; `Tools/Modules/Documentation/Changelog.psm1:207`; `Tools/Modules/Testing/TestRunArchive.psm1:513-594`; `Tools/Modules/Testing/HostedTestDesktop.psm1:93`, `:120`, `:147-149` |
| L64 | Build and CI | new unless noted | CONFIRMED | (a) `DxUiPaths.psm1` is missing from the Terminal cache key and from the identity/ProjectGraph closures. (b) The aws-c-common overlay README still names 1.0.0. (c) Zero-warning enforcement lives only in `build.ps1`, not in direct MSBuild callers, and silently overrides the Gate0 projects' `TreatWarningAsError`. (d) In CI the `ForceImportAfterCppTargets` global property suppresses the environment value for `DxUi.vcxproj` (PLAUSIBLE). (e) The format job checks out the head ref, which breaks on fork PRs (pre-existing). (f) `fetch-depth: 0` everywhere, `GH_TOKEN` exported to every step, and free text interpolated into PowerShell in the Ghostty status workflow (pre-existing hygiene). (g) The build receipt still records `programdatabase` for `/ZI`. | `BuildEvidence.psm1:703-760`; `Build/vcpkg-ports/aws-c-common/README.md:3`; `Build/NativeBuildPolicy.targets:26,38`; `build-reusable.yml:54,74,482`; `ci.yml:68`; `ghostty-runtime-status.yml:41-52`; `build.ps1:841-927` |
| L65 | Test harness | new | PLAUSIBLE | A crash or job kill between the F24 key-down and key-up leaves F24 latched. Every later run then refuses with `held-input`, and nothing sends the key-up. | `TestSupport.Tests.DirectedSelfTestInputWarning.h:147`, `:193`, `:225` |
| L66 | Test harness | mixed | CONFIRMED | (a) Four legacy raisers still use `AttachThreadInput`, against the new contract. (b) The ViewerVLC keyboard fixture forces the HUD visible, so the hosted "VLC missing" state is no longer exercised. (c) Six `FlushQueuedSettingsSaves(INFINITE)` calls. (d) R4A19's independent-volume control fails on hosted ARM64 because the VHDX shares the physical disk; the gate is honest, but the premise does not hold there (Q16). | `Commands.Tests.Search.cpp:1230`, `ViewCommands.cpp:148`, `Preferences.cpp:152`, `Settings.cpp:9519`; `ViewerPE.Tests.Runner.cpp:16555-16630`; `Commands.Tests.Settings.cpp:4939-5519`; `FolderWindow.FileOperations.Tests.Runner.cpp:6905-6934` |

### Checked and found sound (revision 2)

- **ViewerPluginManager deferred load.** `EnsureLoaded` loads only after the sweep has released every old
  module for that path, so there is never a second instance or a factory from a freed module.
  `AddDeferredPlaceholder` always appends, and the disabled reconcile uses the current settings.
- **DirectoryInfoCache `dirtyVersion`.**
  - Every read and write is under `_mutex`, and `refreshPosted` is reset before the follow-up post.
  - Waiters perform one reload, the failure path keeps the entry dirty, and eviction skips loading and
    borrowed entries.
- **Host per-item pool (5-17 workers).** No per-worker array exists in the host: the in-flight and stream
  arrays evict, and `fileWorkReceipts` is indexed by item. The 1-2 core change (4 concurrent items) is
  the documented intent.
- **Popup.** About 90% of the delta is `ENABLE_TESTS`. The speed-prompt debug token protocol is
  type-checked and drained, and public `CompactChildren` is a no-op when nothing was removed.
- **Batch Rename `HoldForDispatch`.** It is balanced and defers `delete this`.
- **Common and services.**
  - The `ThemeExpression` rewrite is equivalent; the only change is that a color key `palette.X` no
    longer collides with palette entry `X`.
  - `LocalFileTransaction::Abort` marks deletion through the owned handle.
  - The search service changes are test-only, the MTP path fast path is equivalent to the slow path,
    and the FTP fixture is `ENABLE_TESTS`-only.
- **Dependency bumps.**
  - The curl, sqlite and 7zip API surface used by the code is unchanged, and the staged DLL names
    match the produced ones.
  - The aws-c-common overlay is blob-identical to the registry tree, plus one include line.
  - The new `.gitattributes` keeps evidence `-text` and pins the Ghostty lock and patches to LF.
- **Ghostty runtime.**
  - Every download is HTTPS and SHA-256-pinned, and is verified before extraction or execution.
  - Children run in kill-on-close job objects and nothing kills by name.
  - Cache deletes are confined to `.build` and reject reparse points.
  - Workflow actions are SHA-pinned with `contents: read`.
  - The new patch enum values sit after the public range with compile-time ABI assertions.
- **Changelog fold.** No `--force`. It deletes only folded fragments, and `Test-Changelog -Base` passes
  on `main`. `CHANGELOG.md` holds 196 Unreleased entries with no duplicates.
- **Test harness.**
  - The F24 handoff refuses before sending input in every unsafe state and releases the key on every
    normal, timeout and failure path.
  - The restoration change is not vacuous.
  - No assertion, timeout or budget was weakened; several were tightened (ImgRaw completion, Shortcuts
    copy, entry-lookup count, Commands focus wait).

## High

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| H1 | Host services | pre-existing | CONFIRMED (routing); PLAUSIBLE (MS Drive effect) | `IHostConnections::SetConnectionSecret`, `DeleteConnectionSecret` and `IHostViewers::OpenViewer` called off the UI thread do `SendMessageW` to the FolderWindow, whose `WndProc` has no case for `kHostSetConnectionSecret`/`kHostDeleteConnectionSecret`/`kHostOpenViewer`; `DefWindowProcW` returns 0, the sender returns it as `S_OK`. Microsoft Drive stores rotated/new OAuth refresh tokens this way from provider operations, so a token can silently never be saved or deleted (repeated sign-in). | `RedSalamander/FolderWindow.cpp:1404-1413` (only dispatcher), `RedSalamander/HostServices.cpp:907,934,1216` (senders), `:1332,1346,1419` (unreachable handlers), caller `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:7909,7936,7956`, reached from `MicrosoftDriveFileWriter::UploadSimple` (`:7516-7519`) on a copy worker |
| H2 | Test harness | new | CONFIRMED (code + CI log); **rev 2: FIXED** | ViewerPE desktop-restoration identity requires the saved focus HWND to belong to the foreground window's thread. A focus child created by another thread of the same process (same pid, same GA_ROOT) fails `savedIdentityValid`, the no-write shortcut is also skipped, and the Viewer suite aborts. This is the ARM64 nightly failure (`savedFocusTid=1448`, `savedTid=7896`, same pid 3196, input unchanged). | `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:1021-1044` |
| H3 | Packaging | new | CONFIRMED (target condition; reviewer ran the target) | `ValidateMsvcToolsVersion` runs for every first-party project, including the MSIX `.wapproj`, which imports no VC props, so `VCToolsVersion` is empty and the build stops with "requires MSVC build tools 14.52.x; selected ." No MSIX can be produced (`build.ps1 -Msix`, `release.yml` package-msix). `RSVerifyRedMonitorClient` beside it already filters on `.vcxproj`. | `Directory.Build.targets:324-326` |
| H4 | Repository / CI | new | CONFIRMED; **rev 2: PARTIAL** | 607 tracked paths (up to 288 characters) under `Specs/Plans/Done/.../visual-gallery/` and `Specs/TestRuns` exceed MAX_PATH on a default Windows checkout; none did at the baseline. Workflows without `core.longpaths` fail at Checkout: Ghostty runtime status and Nightly Flake (logs), and by the same mechanism every `release.yml` job and `winget-release.yml` (not yet run). Contributors on stock Windows git cannot clone. | `.github/workflows/ghostty-runtime-status.yml:26-31`, `nightly-flake.yml:19-24`, `release.yml:42,75,174,219,404`, `winget-release.yml:77,83` |
| X1 | Qualification | process | CONFIRMED | See verdict item 1. Runs: Nightly Native 40660649657 (build passes; suite fails 17-18 cases on x64 and 59-66 on ARM64), ASan 40484804050 (ARM64 59 failures; x64 FileOps suite silent for 4 h 10 min until the 6 h job limit, cause not diagnosed), Flake 40691121033, Ghostty 40665991855. Earlier "successful" nightlies were 8-second skips; the only real baseline run (35498140593 at `0f6cd8ddb`) failed the same cluster. The "Manifest validation failed" and "Merge of test-triplet" log lines come from passing Pester tests, not failures. | `.github/workflows/*`, `Specs/Plans/WIP/Msvc1452ReleaseOptimization_2026-10-23.md`, `RedMonitor_Extraction_2026-10-21.md` |

## Medium

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| M1 | Viewer actions | new | CONFIRMED (code); PLAUSIBLE (runtime) | ViewerSpace "Delete" (and recycle escalation) is admitted with `FILESYSTEM_FLAG_NONE`, so `deletion.recursive=false` and a marked non-empty folder fails with `ERROR_DIR_NOT_EMPTY` after the user confirmed. Pane and Find Files delete pass `FILESYSTEM_FLAG_RECURSIVE`. | `RedSalamander/FolderWindow.FileOperations.cpp:3722-3728`, `:2629` |
| M2 | FileOps engine | pre-existing | CONFIRMED (mechanism) | Lost-wakeup class: `PerItemTaskScheduler::NotifyWorkAvailable` and `Shutdown`, `NotifyQueueChanged` (used by `RequestCancel`), and the bridge pipeline `readerFinished`/`pipelineStop` stores notify without taking the waiter's mutex while the predicates read state written elsewhere. Waits are untimed, so Resume/Cancel can be lost and a job, queued task or `readerThread.join()` hangs. Same class as the earlier `SetPaused` bug; `WakePauseWaiters` already shows the fix. | `RedSalamander/FolderWindow.FileOperations.State.cpp:6479-6482`, `:6497-6525`, `:6899`, `:12660-12692`, `:18394-18395`, `:18490-18493`; `State.Runtime.cpp:2560-2567`; `State.Queue.cpp:1850`, `:2041` |
| M3 | Preferences | pre-existing (+1 key new) | CONFIRMED | `MaybeResetWorkingFileOperationsSettingsIfEmpty` hand-copies the non-default test and omits `popupFooterOnly`, `popupCompactDensity`, `popupThroughputExpanded` and the three `issuesPane*` keys. Turning a Preferences option back to default resets the whole `fileOperations` object and Apply saves it: popup density/disclosure and the Issues pane sort/columns are silently lost. | `RedSalamander/Preferences.Internal.cpp:1469-1491` vs `Common/Common/SettingsStore.cpp:5517-5535` |
| M4 | Local provider | pre-existing | CONFIRMED | Overwrite (staged) copies and preserved links lose Hidden: `ClearTemporaryAttributes` strips `HIDDEN|TEMPORARY` from the owned stage after the source attributes were applied. A hidden file copied over an existing one ends visible; a new-name copy keeps Hidden. | `Plugins/FileSystem/FileSystem.cpp:4408-4431`, `:4607`; `FileSystem.FileOps.cpp:4594-4635`, `:5719-5750` |
| M5 | Local provider | pre-existing | CONFIRMED | Stage and backup names are `<leaf>.rs_copy_tmp_<pid>_<tid>_<counter>` (+47 chars) and `<leaf>.rs_bak_{GUID}` (+46). Leaves of 209+ characters fail Overwrite/new-directory/link copy with an invalid-name error; the names are predictable. The spec forbids PID/TID/counter fallbacks; `Common::Paths::BuildUniqueSiblingName` already does this right. | `Plugins/FileSystem/FileSystem.FileOps.cpp:4184-4192`, `FileSystem.cpp:4389-4406`; `Specs/FileSystem/FileSystem_FileOperations.md:1074-1075` |
| M6 | Local provider | pre-existing | CONFIRMED | Link Preserve rebuilds the target from the advisory PrintName instead of the SubstituteName, and `Win32PathToNtPath` maps `\\?\C:\x` to `\??\UNC\?\C:\x`, so such links are copied dangling or pointing elsewhere. | `Plugins/FileSystem/FileSystem.FileOps.cpp:11280`, `:2398-2411` |
| M7 | Local provider | pre-existing | PLAUSIBLE | Direct and staged copy fail with `ERROR_FILE_INVALID` when the path-query size differs from the bound reader's size (file being appended, or changed while waiting for a transfer slot); Retry fails the same way. | `Plugins/FileSystem/FileSystem.FileOps.cpp:5473-5509`, `:5221-5224`, `:4429-4432` |
| M8 | MTP | new | CONFIRMED | The "create temp refused" memo is set before the direct fallback, includes `ERROR_ACCESS_DENIED`, is never reset, and `SupportsAtomicWriterCommit` keeps answering TRUE. One denied folder turns every later create on the device into a non-atomic direct upload for the rest of the session. | `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:268-273`, `:3524-3541`, `:4532-4543`, `:3309-3326` |
| M9 | Move breadcrumb | pre-existing | CONFIRMED | Every `FileOperationState` ctor loads all non-terminal breadcrumbs, and records carry no writer identity. Opening Compare Directories (own embedded FolderWindow) or a second process during a long Move shows a false "Move was interrupted" card; dismissing it deletes the live record, so a crash later produces no notice. | `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp:567-571`, `:578-665`; `FileOperationMoveBreadcrumb.cpp:486-508`; `FileOperationMoveBreadcrumb.h:50-61` |
| M10 | Move breadcrumb | new | CONFIRMED | The interrupted-Move card header uses `sourceRootCount` (unique provider roots, usually 1) as the item count and a full path as the single name, so a 500-file interrupted Move reads "Move C:\...\first.txt". | `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp:629-631`, `:650-652` |
| M11 | FileOps engine | pre-existing | CONFIRMED (early exits); AGENT (classification) | Inline Rename refusals before any mutation (file in use, vanished, access denied, case-only unsupported) return before `recordStatus` exists, so the task ends Indeterminate / `ERROR_IO_INCOMPLETE` ("outcome unknown") instead of a clean "not renamed". Batch Rename handles this case. | `RedSalamander/FolderWindow.FileOperations.State.cpp:13505-13548` |
| M12 | Popup keyboard | new | CONFIRMED | With no card control focused, Escape is bound to the sole actionable decision in the whole snapshot, even when that card is scrolled out of view, and submits its Escape action instead of hiding the popup. A user pressing Escape to close the popup can cancel a task they cannot see. Spec allows the sole-decision fallback but does not require visibility (Q4). | `RedSalamander/FolderWindow.FileOperations.Popup.cpp:8393-8460`, `:8479-8482`, `:13931-13938` |
| M13 | Popup a11y | pre-existing | CONFIRMED | A plain Running task's progress bar has an empty UIA Name and the failure surface prints "Copy: " with nothing after it: `StatusTextForTask` returns `{}` for Running and the fallback only covers a missing task. | `RedSalamander/FolderWindow.FileOperations.Popup.cpp:8711-8720`, `:1765-1810`, `:13807-13810` |
| M14 | Plugin buffers | new (leftover) + pre-existing (pane, sibling menu) | CONFIRMED | Three walks of plugin `IFilesInformation` buffers do not bound `FileNameSize` (and the sibling menu not even `NextEntryOffset`); the leftover walk also joins unvalidated child names with `/=`, so a rooted or `..\` name escapes the folder and is offered pre-checked for deletion. `Common::Plugins::PackedFileInfoCursor` exists for exactly this. | `RedSalamander/FolderWindow.LeftoverCleanup.cpp:683-709`; `FolderView.Enumeration.cpp:424-516`; `NavigationView.Menus.cpp:2394-2452` |
| M15 | Leftover cleanup | new | CONFIRMED (mechanism) | `~LeftoverCleanupDialogState` calls `StopScan(true)`, which loops `CancelSynchronousIo` + `Sleep(20)` on the UI thread with no bound. `CancelSynchronousIo` does not interrupt Curl/HTTP socket waits, so closing the window or the app after a subfolder walk on a stalled FTP/SFTP/S3/Drive location hangs until the provider times out. Non-final closes also sleep up to 2 x 200 ms. | `RedSalamander/FolderWindow.LeftoverCleanup.cpp:77-91`, `:126-129`, `:800-817`, `:1104-1137`; `FolderWindow.cpp:2286-2291` |
| M16 | Leftover cleanup | new | CONFIRMED | Unreadable folders in the subfolder walk are skipped silently; if the root fails the dialog shows "No leftovers ... found here" after it listed rows a moment before. | `RedSalamander/FolderWindow.LeftoverCleanup.cpp:684-692`, `:767-794`, `:544-585` |
| M17 | Leftover cleanup | new | PLAUSIBLE | The "row in use" filter is textual and window-local, a live scope without `pathIdentity` fails open, and pre-check ignores age, so another window's or process's live stage can be offered pre-checked. No user data is lost (a stage is a copy) but another transfer can fail. | `RedSalamander/FolderWindow.LeftoverCleanup.cpp:336-351`; `State.Queue.cpp:1590-1636` |
| M18 | FolderView | pre-existing | PLAUSIBLE | Cancel on a busy enumeration only bumps the generation; the provider call keeps the single enumeration worker, so the next navigation waits until the redirector times out. Spec says Cancel "aborts enumeration" (Q6). | `RedSalamander/FolderView.Enumeration.cpp:329`, `:928`; `FolderView.Interaction.cpp:593`; `DirectoryInfoCache.cpp:23-48` |
| M19 | ViewerSpace | new | CONFIRMED (code) | `_modeCombo` is never assigned, so `GetFocusControl() == _modeCombo` is true whenever no DxUi control has focus; Up/Down in the navigation row then silently switch Explorer/Analyst and persist the configuration. | `Plugins/ViewerSpace/ViewerSpace.NavigationHost.cpp:403`, `ViewerSpace.NavigationHost.h:121` |
| M20 | ViewerSpace | new | CONFIRMED | `WorkspaceOwnedBytes()` is counted inside `AnalysisOwnedBytes()` and again in `ViewerOwnedBytes()` (and in deferred-disposal charges); a process-global 8 MiB export charge inflates every viewer, so refresh/cache/export are refused near the budget. | `ViewerSpace.Analysis.cpp:235`, `ViewerSpace.Layout.cpp:247,271`, `ViewerSpace.Model.cpp:88`, `ViewerSpace.Workspace.cpp:1134-1146` |
| M21 | ViewerSpace | pre-existing | CONFIRMED | Up at a UNC share root is enabled and rescans `\\server\` (MSVC `root_path` of `\\nas\media` is `\\nas\`), ending in Error with the good model discarded; a trailing separator rescans the same folder. Spec says Up is disabled at a share root. | `Plugins/ViewerSpace/ViewerSpace.cpp:1411-1438`, `:13715-13724`, `:13852-13861` |
| M22 | ViewerSpace | pre-existing | CONFIRMED | Closing during a scan tests `done` immediately after `request_stop`; a cooperative worker is almost always still running, so it is detached and the unload gate is kept, quarantining plugin unload for the rest of the process. The "uncooperative provider" exception is the common close path. | `Plugins/ViewerSpace/ViewerSpace.cpp:8436-8503`, `:2238` |
| M23 | FileOps queue | pre-existing | PLAUSIBLE | The same-host overlap advisory runs the pairwise O(N x M) scope predicate while holding `_queueMutex`, blocking workers and the UI (Start now, queue moves) for large selections. The gate already has an indexed path. | `RedSalamander/FolderWindow.FileOperations.State.Queue.cpp:1168-1219`, `:1432-1526` |
| M24 | Host services | pre-existing | CONFIRMED | Eight synchronous host messages still cast `lParam` to a raw pointer (any same-desktop process can crash the UI thread with one `SendMessage`); only two were moved to a token registry whose tokens are a sequential counter. Contradicts `Plugins_PluginAPI.md:188` and AGENTS.md; `docs/dev/PluginHostModel.md:107-108` documents the old pattern. | `RedSalamander/HostServices.cpp:116-218`, `:1295-1421` |
| M25 | Tooling | new | CONFIRMED | `format-all.ps1` formats every `*.h` outside build folders, including the hash-locked LF-only `ThirdParty/RedMonitor/include/RedMonitor/Client.h`; `.clang-format` forces CRLF and `ThirdParty/RedMonitor/.gitattributes` is `* -text`, so the lock hash breaks and every first-party project fails to build. It also uses the newest local clang-format, not the pinned 22.1.3. | `format-all.ps1:48-137`, `.clang-format:7`, `Dependencies/RedMonitor.lock.json` |
| M26 | Tooling | new | CONFIRMED | An exception in the end-of-run evidence archive (leftover junction, locked file, unstable inventory) is uncaught after the evidence run was already completed as passed: no summary, no verdict, exit 1 instead of the documented 5. | `Tools/Run-AllTests.ps1:2291`, `:2306-2307`, `:2588-2611` |
| M27 | Tooling | new | CONFIRMED; **rev 2: PARTIAL** | `-NonInteractive` / `Test-Changes.ps1 -Suite <family>` report PASSED and exit 0 for legacy suites although desktop entries were dropped; help promises NOT_EVALUATED / exit 2. | `Tools/Run-AllTests.ps1:1491-1499`, `:1746`, `:2356-2371`; `Test-Changes.ps1:38-62` |
| M28 | Test harness | new | PLAUSIBLE | x64 nightly restoration of a foreign foreground window is one unsettled `SetForegroundWindow` attempt, judged 1.4 ms later with no diagnostics, and its failure aborts the Viewer suite. | `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:1088-1112` |
| M29 | Search index | pre-existing | AGENT (reproduced by the nightly and the baseline run); **rev 2: FIXED** | When the USN journal is unavailable, the rebuild calls `BuildIndex`, a whole-volume MFT enumeration, instead of the tracked-root traversal; `local_index_core_journal_replay_rename_delete_create` fails. | `Common/LocalSearchIndexCore.cpp:3375`, `:3479` |
| M30 | Search service | pre-existing | AGENT (reproduced); **rev 2: PARTIAL** | Multi-client service access fails with `0x800700E7` (pipe busy) under the 750 ms connect limit; `search_service_multi_client_and_rebuild_control` fails on hosted runners. | `Common/SearchServiceBroker.cpp` |
| M31 | CI | pre-existing | PLAUSIBLE; **rev 2: CHANGED** | The Nightly Flake `-SkipBuild` consumer can never use the receipt: `Test-RSBuildReceiptForArtifactUse` needs the pinned DxUi source tree, which the consumer job never restores, and returns a bare `$false` that does not name the failed predicate (100 of 100 runs failed since 2026-07-22). | `Tools/Modules/Build/BuildEvidence.psm1:993-1018`, `Tools/Run-AllTests.ps1:1432-1441`, `.github/workflows/nightly-flake.yml` |
| M32 | CI | pre-existing | AGENT; **rev 2: PARTIAL** | When the test-root step did not run, the Nightly Flake upload uses `path: \` and scans the drive root (`EPERM ... System Volume Information` masks the real error); on success it uploads the whole sandbox, including VHDX evidence, without the reusable workflow's exclusions. | `.github/workflows/nightly-flake.yml:84-90` |
| M33 | FileOps throughput | pre-existing | AGENT (reproduced); **rev 2: FIXED** | `FileOps_SelectionShapeLocal` fails on 4-core runners: selected-file copy is capped at 3 transfers against 4 for folders because the per-item pool sizing counts the starvation reserve. | `RedSalamander/FolderWindow.FileOperations.State.cpp:6630-6637` |
| M34 | CI / build | new | PLAUSIBLE; **rev 2: FIXED** | The inline CI `vcpkg install` passes neither the overlay triplets nor the ARM64 `aws-c-common` overlay port, and the cache key omits `Build/vcpkg-ports`; a cold ARM64 cache would compile stock `aws-c-common` with the C1001 the overlay avoids. Every observed run restored from cache, so this path is untested. `build.ps1` comments claim CI uses the governed installer. | `.github/workflows/build-reusable.yml:381-388`, `Build/vcpkg-ports/aws-c-common/README.md` |

## Low

| ID | Area | New? | Verified | Finding | Where |
|---|---|---|---|---|---|
| L1 | ViewerSpace | new | CONFIRMED | Retained refresh has no visible Cancel button, but a hidden Cancel hit target in the header discards it on click. | `ViewerSpace.cpp:4072-4088` vs `:5622-5626` |
| L2 | ViewerSpace | new | PLAUSIBLE | Clicking the treemap never takes keyboard focus, so Backspace/arrows/Enter go to the filter or grid. | `ViewerSpace.cpp:5608-5652` |
| L3 | ViewerSpace | pre-existing | CONFIRMED | Arrow-key selection walks culled (invisible) draw items; Up/Down equal Left/Right. | `ViewerSpace.cpp:5442-5460` |
| L4 | ViewerSpace | new | PLAUSIBLE | Alt+Left/Right/Up/Insert arrive as `WM_SYSKEYDOWN` and are dropped when the canvas has focus. | `ViewerSpace.cpp:3318-3335` |
| L5 | ViewerSpace | new | CONFIRMED | `IDS_VIEWERSPACE_HEADER_PROBLEMS` is English in all four satellites ("Hotovo · 3 unreadable"); the translated `..._DONE_WITH_UNREADABLE_FORMAT` is unused. | `ViewerSpace.cpp:7366-7375`, `ViewerSpaceResources.rc:85,252` |
| L6 | ViewerSpace | new | CONFIRMED | `_workspaceScanStartedSeconds` is not reset by `StartScan`, so a restarted scan's items/s uses the previous start time. | `ViewerSpace.cpp:7417-7433`, `:8095-8109` |
| L7 | ViewerSpace | pre-existing | AGENT | A canceled scan stuck in a provider call keeps the volume permit; the next scan shows "Queued..." without explanation. | `ViewerSpace.cpp:1684-1729`, `:9196-9220` |
| L8 | ViewerSpace | pre-existing (heavier now) | AGENT | Natural sort, summary encoding and per-child record building run under `frontierMutex`. | `ViewerSpace.cpp:10220-10422` |
| L9 | ViewerSpace | new | CONFIRMED | File-count weighting evicts/sorts by bytes while area uses file count. | `ViewerSpace.Layout.cpp:356-357,470,536` |
| L10 | ViewerSpace | new | CONFIRMED | Breadcrumb folder menu takes the first 512 children in arena order, then sorts. | `ViewerSpace.cpp:3711-3735` |
| L11 | ViewerSpace | new | CONFIRMED | `_exportProjectionRequested` is not cleared on the WAIT/LIMIT early returns, so the alert re-shows each tick and the timer never idles. | `ViewerSpace.Workspace.cpp:78-112`, `:259-263` |
| L12 | ViewerSpace | new | AGENT | Deferred refresh adoption retry drops the "view folder selected" resolution. | `ViewerSpace.Model.cpp:526-530`, `:615` |
| L13 | ViewerSpace | new | AGENT | One unpaired-surrogate name fails the whole export with a generic I/O error. | `ViewerSpace.Workspace.cpp:296-303` |
| L14 | ViewerSpace | new | AGENT | Escape inside the analysis panel never reaches its own Escape handler (dead `panelClose` path or spec mismatch, Q7). | `ViewerSpace.AnalysisHost.cpp:195-202`, `:722-733` |
| L15 | FileOps engine | new | AGENT | 64-slot write-only consent-receipt table turns the 65th per-item consent into `ERROR_NOT_ENOUGH_MEMORY`. | `State.cpp:6145-6182`, `:6235-6246` |
| L16 | FileOps engine | pre-existing | AGENT | `Task::_stopToken` assigned on the worker while UI threads read it (formal data race). | `State.cpp:11824` |
| L17 | FileOps runtime | pre-existing | CONFIRMED | Interrupted-Move notices are evicted by the 24-summary cap although Clear keeps them (they return next launch). | `State.Diagnostics.cpp:2327-2335` |
| L18 | FileOps runtime | pre-existing | AGENT | Rejected/corrupt breadcrumbs are never logged (dead warning branch); failed `Advance(Executing)` logs at Debug. | `State.Runtime.cpp:582-592`, `State.cpp:12006-12018` |
| L19 | FileOps queue | pre-existing | AGENT | Indexed overlap scan sorts keys by raw string order (`DIR-OLD` sorts between `DIR` and `DIR\X`), so an ancestor can be popped early. Same weakness in Local `DeleteItems` (not reachable: host rejects nested selections). | `State.Queue.cpp:1043-1083`; `FileSystem.FileOps.cpp:13608-13647` |
| L20 | FileOps completion | pre-existing | CONFIRMED | Directory-cache notifications use the live pane's file system instead of the task's, leaving stale entries after navigation. | `FolderWindow.FileOperations.cpp:5255-5330` |
| L21 | Localization | pre-existing | AGENT | Hard-coded English plural `s` feeds `{1:s}` in cs/sk/ja confirmation text ("3 položkas"). | `FileOperationConfirmation.cpp:63-87`, `State.Runtime.cpp:1312-1337` |
| L22 | Diagnostics | pre-existing | AGENT | Issues export leaves a truncated file on write failure; `OpenDiagnosticsLogForTask` scans the Logs folder under `_mutex`. | `State.Diagnostics.cpp:645-649`, `:803-891` |
| L23 | Popup | mixed | AGENT | User-visible strings stitched with hard-coded separators (`" · "`, `": "`, `"{0}. {1}"`). | `Popup.cpp:11871`, `:13794-13810`; `PromptText.cpp:59-122` |
| L24 | Popup | pre-existing | AGENT | Arrow keys stop at a disabled control instead of skipping it. | `Popup.cpp:7646-7668` |
| L25 | Popup | new | AGENT | Focus on a Completed-group header control jumps to Options when the header scrolls out. | `Popup.cpp:9700-9703`, `:8839-8847` |
| L26 | Popup | unknown | PLAUSIBLE | Speed-limit prompt destroys `_dxHost` before `_hWnd`; WM_NCDESTROY then dispatches into the destroyed host at shutdown. | `Popup.cpp:3286-3287`, `:2582-2613` |
| L27 | Popup | pre-existing | PLAUSIBLE | "Cancel all" consent defaults to OK with no held-Enter guard (the failure-surface variant has one). | `Popup.cpp:12810-12859` |
| L28 | Popup perf | pre-existing | AGENT | Hidden popup rebuilds a full snapshot and calls `ITaskbarList3` at 10 Hz without change detection. | `Popup.cpp:14305-14309`, `:12496-12539` |
| L29 | Popup perf | new | PLAUSIBLE | Middle truncation re-runs a DirectWrite binary search for every path row every frame. | `Popup.cpp:3399-3464` |
| L30 | Overlay | new | PLAUSIBLE | `AlertOverlayWindow::WndProc` forwards only four messages to its DxUi host, against the DxUi "forward every message" contract. | `Ui/AlertOverlayWindow.cpp:1223-1310` |
| L31 | Specs | new | CONFIRMED | `Core_StartupBootstrap.md:105-117` still specifies the removed RedSalamanderMonitor auto-launch. | `Specs/Core/Core_StartupBootstrap.md` |
| L32 | MTP | new | AGENT | Create temps (`.rs-mtp-create-*.tmp`) have no recovery journal; a crash leaves a full-size hidden object on the phone. | `FileSystemMtp.Core.cpp:307-367` |
| L33 | MTP | new | CONFIRMED | `FileSystemMtp.Tests.FakeBackend.cpp` (~1,300 lines) has no `ENABLE_TESTS` guard and ships in the Release DLL; every sibling fake is guarded. | `Plugins/FileSystemMtp/FileSystemMtp.Tests.FakeBackend.cpp:1`, `FileSystemMtp.vcxproj:146` |
| L34 | ViewerVLC | pre-existing | AGENT | Rapid selections during the initial libVLC load spawn parallel loaders. | `ViewerVLC.cpp:5333-5400` |
| L35 | ViewerWeb | new (test builds) | CONFIRMED | Test-profile admission failure is silent (no alert, no log). | `ViewerWeb.cpp:4202-4214` |
| L36 | Local provider | new | AGENT | `\\?\unc\` (lower case) is UNC for routing but not for the prefix helpers. | `FileSystem.cpp:961-964`, `FileSystem.Path.cpp:29-32,266-369` |
| L37 | Local provider | pre-existing | PLAUSIBLE | An abandoned `Win32FileWriter` deletes by path after closing and swallows failure. | `FileSystem.cpp:4931-4943` |
| L38 | Diagnostics | new | CONFIRMED | ETW messages over 16,000 code units are now dropped (RedMonitor `ERROR_BUFFER_OVERFLOW`) instead of truncated. | `Common/Helpers.h:1715-1721`, `ThirdParty/RedMonitor/include/RedMonitor/Client.h:106-110` |
| L39 | Search | new | AGENT | Dead `state == Ready` branch and always-true `storedVolumeReady` probe after the new early return. | `Common/LocalSearchIndexCore.cpp:4012-4036` |
| L40 | Tooling | new | CONFIRMED | Changelog fold drops a distinct fragment whose text equals any Unreleased entry, then deletes its file. | `Tools/Modules/Documentation/Changelog.psm1:94-101` |
| L41 | Tooling | new | AGENT | `New-RSTestSandboxLink` validates a relative `-Path` against the process directory but creates it against the PowerShell location. | `Tools/Modules/Testing/TestSandbox.psm1:341-396` |
| L42 | Build | new | AGENT | `RSVerifyRedMonitorClient` spawns `pwsh` (~1.1 s) for each of 114 projects on every build. | `Directory.Build.targets:327-329` |
| L43 | Repo hygiene | new | CONFIRMED | Public repo now tracks 19,500 evidence files (pack 502 MB); the private NAS hostname appears in 3 Done plans and 27 evidence files; `advisor-plans/` (agent executor plans) and `.codex/config.toml` (`workspace-write`, `network_access = true`) were published. | `Specs/TestRuns`, `Specs/Plans/Done`, `advisor-plans/`, `.codex/config.toml` |
| L44 | Repo hygiene | pre-existing | CONFIRMED | Two different files share the name `FolderWindow.FileOperations.Tests.OwnedShutdown.cpp` (compiled hooks in `RedSalamander/`, included fixture in `SelfTest/FileOperations/`); `vcpkg.json` homepage typo `Redsalamanderse`. | `RedSalamander/RedSalamander.vcxproj:413,530`; `vcpkg.json:4` |
| L45 | Build | new | CONFIRMED (reading) | `Verify-RedMonitorClient.ps1` rejects any reparse point up to the drive root, so a checkout below a junction or OneDrive placeholder fails every build; it needs `pwsh` on PATH even for IDE builds (with L42). | `Tools/Verify-RedMonitorClient.ps1`, `Directory.Build.targets:327-329` |
| L46 | Monitor removal | new | AGENT | Dangling RedSalamanderMonitor remnants (none breaks the build): dead `g_redSalamanderMonitor*` and `kColorTextView*` constants, stale tool/test lists, ETW scripts that still say "launch RedSalamanderMonitor", docs. `.sln`, projects, installers and workflows are clean. | `Common/Helpers.h:1392-1393`, `Common/WindowMessages.h:1051-1054`, `init-etw-trace.ps1`, `start-etw-trace.ps1`, `stop-etw-trace.ps1`, `docs/SettingsFile.md`, `docs/Preferences.md`, `docs/dev/Diagnostics.md` |
| L47 | Specs | new | AGENT | `Core_RedSalamanderMonitor.md` keeps about 40 KB of the removed viewer's requirements as live-sounding prose with nonexistent paths. | `Specs/Core/Core_RedSalamanderMonitor.md` |
| L48 | Packaging | pre-existing | PLAUSIBLE | MSI and MSIX harvest the whole build output folder, so build receipts, toolchain identity and test `*.DxUi.json` sidecars can ship; the ZIP uses an allow-list. | `Installer/msix/RedSalamanderInstaller.wapproj:76`, `Installer/msi/build-msi.ps1:168` |
| L49 | ViewerImgRaw | unknown | NOT DIAGNOSED; **rev 2: CHANGED** | ARM64 Release only, flaky (retry passed): an oriented WIC TIFF reports a completed decode but no image (`source=0x0`). Candidates: decode-complete published before the image is installed, or ARM64 `/GL` codegen. Needs an isolated repeat. | `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:5054`, `Plugins/ViewerImgRaw/ViewerImgRaw.cpp:884` |

## Checked and found sound

So nobody re-flags them: the plugin ABI (`FileSystem.h` vtables unchanged, only a new interface, tail-extended options with exact-size check, V1/V2/V3 mutation-result prefixes); `SettingsStore` migrations and schema; `DirectoryRecordCursor`; `SynchronousIoCancelWatch`; `PerfJsonl`; `NaturalSort` (strict weak order, tested empirically); `WindowMessages.h` id uniqueness; every satellite `.rc` placeholder token; `TakeMessagePayload` call sites and payload drains in all viewers; File Operations reducer, conflict broker and arbiter protocol; DecisionMirror (no cross-task decision leak); Issues pane/projection consistency; leftover cleanup cannot delete by itself (ordinary confirmed Delete, Cancel default); Local delete/overwrite/Move/Recycle; S3 test-latency code fully inside `ENABLE_TESTS`; Microsoft/Google Drive diffs are formatting only; ViewerWeb's three merge-day lifetime commits are refcount-balanced and test-only; ViewerVLC worker-owned media switch; ViewerMenuHost ownership; sandbox root authorization, no-follow deletion, owned-VHDX formatting guards, no kill-by-name; clang-format wheel hash pinning; skill trims keep every AGENTS.md rule and all links resolve; FileOps self-test family registration (every scheduled step is in a family); Fold-Changelog `-Check`, Test-Changelog and Get-SpecInventory are green on `main`; the solution, all 114 project include lists, installers and release workflows carry no dangling RedSalamanderMonitor reference.

## Red scheduled workflows: diagnosis

| Workflow (run) | Root cause | Class | Owner finding |
|---|---|---|---|
| Ghostty runtime status (40665991855) | Checkout of 98+ over-long paths without `core.longpaths`; the upload failure is secondary. | repository/CI | H4 |
| Nightly Flake Convergence (40665816743, 40477285880 on `main`; 40691121033 on the repair branch) | On `main`: Checkout (H4), then the upload scans `D:\` (M32). On the repair branch, after those fixes: `BLOCKED ATTESTATION` from the `-SkipBuild` receipt check (M31). | CI/harness | H4, M31, M32 |
| Nightly Native Suite (40660649657) | Build passes in all four profiles; the suite fails. Dominant cause is the hosted desktop: 1024x768 display (work-area fixtures cannot fit) and Windows Terminal owning the foreground (about 45 ARM64 menu/popup/foreground cases). Harness: H2 (ARM64 restoration identity), M28 (x64 foreign-foreground restoration), `theme_cycle_overlay_*` (animations off on hosted OS), `R4A19_DiscoveryIndependentVolumes` (timing gate / no alternate volume). Product, also failing at the baseline: M29, M30, M33. Flaky ARM64 Release: L49. | runner + harness + 3 product | H2, M28, M29, M30, M33, L49 |
| AddressSanitizer weekly (40484804050) | ARM64: the same 59-case set as the nightly plus `Phase9_ConflictPrompt_ApplyToAllUiCache` (0 us convergence wait, unclassified). x64: the FileOps suite printed nothing for 4 h 10 min until the 6 h job limit; the native suite step has no step-level timeout. Hang versus budget exhaustion is not diagnosed (artifacts not downloaded). | runner + harness + undiagnosed | X1 |

The in-flight `ej/ci-runtime-repair-20261028` branch already patches the checkout and upload steps
of `ghostty-runtime-status.yml` and `nightly-flake.yml`; it does not cover `release.yml` and
`winget-release.yml` (H4), MSIX (H3), the harness predicates (H2, M28) or the receipt check (M31).
Whether interactive cases should run on hosted runners at all is Q11.

### Revision 2 runs (merged I52 repair, `28e70e694`..`2ec73448a`)

| Workflow (run) | Result | Root cause | Owner finding |
|---|---|---|---|
| Nightly Native Suite (41030069121 at `632dc13b2`) | x64 Debug 2631 passed / 1 failed; x64 Release 2625 / 2 | Viewer foreign-foreground restoration (PE and ImgRaw combo cases) aborts the Viewer suite; Release also fails `search_service_multi_client_and_rebuild_control` through the new busy-pipe fixture | M28, M44 |
| same run, ARM64 Debug and Release | 970 failed each | 968 Commands cases fail the F24 foreground handoff (foreground is `WWAHost` CoreWindow). Also: ViewerSpace owned-activation failures, R4A19 independent-volume timing (concurrent 2.44 s against a 1.69 s gate) and the ViewerPE combo-host `0xC0000005`. The runner labels all three "REGRESSION". | H6, L66(d), M36 |
| Nightly Native Suite (40941503630 at `28e70e694`, dispatch) | x64 Debug 1 failed; ARM64 Release 970 | Same classes | H6, M28 |
| AddressSanitizer (40941634645, dispatch) | x64 cancelled at 6 h 01 min; ARM64 failed | x64 hang versus budget still not diagnosed; no step timeout | M40, X1 |
| Nightly Flake (40941643057; 41037399793) | timed out at 210 min; cancelled after 50 min | `-Suite All -SelfTestRepeat 5` cannot fit the timeout | M43, M31 |
| Ghostty runtime status (41037608178 and 40941651181) | green | H4 fixed for this workflow | H4 |
| CI on `main` (10 push runs since 2026-10-29) | 3 success, 6 cancelled, 1 failure (`b4ff85bea`) | Cancel-in-progress per ref | M41 |

I55 (owner-selected, implementation pending) will remove the scheduled native, ASan and flake
workflows and replace them with batched test-enabled Release on native x64 and ARM64 plus manual
qualification. Until it lands, these workflows are the only hosted evidence. Route fixes through I52
(hosted CI repair) and I55, not as parallel edits.

## Simplification and architecture opportunities

Ordered by value. Each item is behavior-neutral unless it names the finding it also fixes.

| ID | What goes | Size | Risk | Also fixes |
|---|---|---|---|---|
| S1 | Replace the ten per-message host marshals (`Pending*` structs, `SendMessageW` blocks, the counter-token registry, three `kHost*` ids) with one UI-thread call marshal over `SendMessagePayload`/`TakeMessagePayload`, the pattern `AlertOverlayWindow` already uses for UIA (`AlertOverlayWindow.cpp:195-223`). A new service can then not be added unrouted. | about -300 lines in `HostServices.cpp`, `WindowMessages.h`, `FolderWindow.cpp` | medium | H1, M24 |
| S2 | One validated `IFilesInformation` walk: use `Common::Plugins::PackedFileInfoCursor` in `FolderWindow.LeftoverCleanup.cpp`, `FolderView.Enumeration.cpp`, `NavigationView.Menus.cpp` (and check `State.cpp:2896-3040`), plus a shared child-name validator. | about -120 lines | low-medium (FolderView loop is hot; keep it non-allocating) | M14 |
| S3 | Preferences: replace the hand-copied predicate with `Common::Settings::HasNonDefaultFileOperationsSettings`. | -12 lines | low | M3 |
| S4 | Local provider: build stage/backup names with `Common::Paths::BuildUniqueSiblingName`; delete the dead pre-staging overwrite-removal cluster (`CollectDirectoryChildrenVerified`, `RemoveDirectoryRecursiveNoFollow`, `RemovePathForOverwrite`, `FileSystem.FileOps.cpp:4977-5085`); one owned-stage create/abort helper instead of three; one bound-source pump instead of two; fold `StripWin32ExtendedPrefix`/`NtPathToWin32Path`/`Win32PathToNtPath`/`IsExtendedUncPath` into `Common::Paths::ClassifyWindowsPath`. | about -350 lines | medium (data path; run the FileOps families) | M5, M6, M7, L36 |
| S5 | Popup: delete the dead footer/task-button chrome (`DrawFooterAutoDismissControl`, `DrawCheckboxBox`, four unreachable `OnActivatedHit` branches, `_footerNewTasksText`, duplicate AUTODISMISS strings); move `HitTest` under `ENABLE_TESTS`; one status-icon and one ETA helper; carry rate fields in `TaskSnapshot` so a tick builds one snapshot, not two. | about -350 lines | low (S4 of the popup report: medium) | L28 |
| S6 | ViewerSpace: delete never-written layout state (`_layoutCandidateCache`, `_layoutWorkspaceByDepth`, `_autoExpandedOtherByNode`, `_lastRectsByNode` and their types), the dead `_modeCombo` branch and captions, the dead "State" column, 30 unreferenced `IDS_` strings x 5 languages; one `GuardedInvoke` + base class for the two child-host skeletons; move the 575 lines of `#ifdef ENABLE_TESTS` `WndProc` cases to `ViewerSpace.Debug.cpp`; one `ScanProgressState` struct for capture/restore/reset; one paint epilogue; one join-and-release-gate helper. | about -1,100 lines | low | M19, L6 |
| S7 | FileOps engine/runtime: remove the write-only consent-receipt table; one `IsCancellationStatus`; one consent-fact copy helper; one selection-count/phrase helper for the three confirmation builders; one `TakePopupIfIdleLocked()`; drop dead `_pauseCv` notifies in discovery. | about -200 lines | low | L15, L21 |
| S8 | Tooling: delete `Remove-RSTestSandboxStaleRunDirectories` (no caller, weakest guard) and the unused second receipt mechanism in `ScopedTesting.psm1` (about 135 lines); verify the RedMonitor client with MSBuild `GetFileHash` instead of a `pwsh` per project; restrict `format-all.ps1` to the CI path set and the pinned formatter (removes a 90-line resolver). | about -300 lines | low | M25, L42 |
| S9 | Plugins: guard the MTP fake backend with `ENABLE_TESTS`; one `ViewerMenuHost` attach-and-log helper for Web/Text/PE/ImgRaw (only ImgRaw logs a non-hosted outcome today); Dummy uses `PackedFileInfoBuffer::Build`; drop unused `IDS_VIEWERVLC_DETAILS_PATH_NOT_LOCAL_FILE`. | about -1,450 lines from the Release MTP DLL, -150 elsewhere | low | L33 |
| S10 | Split the four god files along their existing seams: `FolderWindow.FileOperations.State.cpp` (25,543 lines; e.g. the 6,000-line cross-file-system bridge, per-item scheduler, batch rename), `FolderWindow.FileOperations.Popup.cpp` (17,666; `Render` alone is 3,400 lines: split the card painter by card kind), `Plugins/FileSystem/FileSystem.FileOps.cpp` (14,507), `Plugins/ViewerSpace/ViewerSpace.cpp` (14,092). | mechanical | medium (large diffs; do it with families green, one file per change) | review-ability, compile time |
| S12 | CI and RedMonitor packaging: one composite "Windows checkout" action instead of eight copies of the long-path step; ship one license file instead of the byte-identical `ThirdParty/RedMonitor/LICENSE.txt` copy; keep one of the two `-text` rules for `ThirdParty/RedMonitor`; retire the `Restore-RedMonitor.ps1` / `RedMonitorDependency.psm1` shims if the inventory allows; delete the dead Monitor constants and stale list entries (L46). | about -100 lines | low | H4, L46 |
| S11 | Search: `MatchNamePattern`/`MatchFileContent` exist twice (`Plugins/FileSystem/FileSystem.Search.cpp:1301`, `RedSalamander/SearchFallbackEngine.cpp:509`, pre-existing); a header-only shared matcher in `Common/` would keep plugin and fallback results identical. | about -80 lines | low | none |
| S13 | One transfer-pool contract for the host and the Local provider, which now each implement "clamp to 4-16, plus one reserve": `State.cpp:6634-6641` reuses `Task::kMaxInFlightFiles`, an array-size constant, as its cap, and the provider uses `kFileOpsMaxTransferWorkers`. Use one named constant plus a helper with `static_assert(cap >= 4)`, and derive every per-stream slot array from `kFileOpsSchedulerWorkerSlots` (including `kDeleteStreamContextSlots` and the local `kMaxWorkers = 16u` at `FileSystem.FileOps.cpp:6826`) with a `static_assert` beside each. | about -20 lines | low | H5 |
| S14 | ViewerPluginManager: `CreateViewerInstance` and `CreateTerminalInstance` are about 45 near-identical lines, and four configuration APIs repeat the factory and QI sequence. Extract `AcquireCreatable(entry, type)` and `CreateInstanceFor(entry, iid)`. Have `AddDeferredPlaceholder(entry, disabled)` return the placeholder instead of relying on `_plugins.back()`. Apply the same helper to the file-system manager. | about -120 lines | low | L52 |
| S15 | Ghostty tooling. Delete the unreachable uucode legacy-recovery paths (`GhosttyRuntimeLifecycle.psm1:603`, `:952-985`) and the likely-dead content-archive copy (`:1327-1331`, confirm with one build). Merge the two drifted candidate git wrappers (`GhosttyCandidateOverlay.psm1:11-24`, `GhosttyCandidateReview.psm1:31-46`); the overlay copy lacks `core.longpaths`. Name the P4 snapshot from the frozen-binding digest instead of an operator path. | about -150 lines | low-medium | M38, L62 |
| S16 | Changelog fold. Use a `--no-checkout` sparse worktree (`CHANGELOG.md`, `Changes`, `Tools`) instead of 27,840 files. Keep one git wrapper instead of three. Replace the about 80 lines of named heuristics in `Get-ChangelogFoldProblems` with one diff of expected against actual text, computed fold by fold. | about -120 lines | low | L63(b), L63(d) |
| S17 | CI. Replace the 155-line inline vcpkg bootstrap and install (`build-reusable.yml:246-401`) with `vcpkg-install.ps1` if the install lock allows. Use one `changes` job for `asan.yml` and `nightly-ci.yml`, and one "select test root" step instead of four. Keep the diagnostic regex, source-path rule, policy name and `RequiredMsvcToolsVersion` parsing in one module each. Drop the plugin/test `/wd` lists that the native policy targets now cover. Fold into S12's composite checkout action, which also closes the H4 gaps. | about -250 lines | low-medium | H4, L64 |
| S18 | Search service client: one "context, deadline, connect" helper for `GetStatus`/`RequestRebuild`/`RequestCompact`/`RequestShutdown`, which also fixes L58 by construction. Delete the dead `ERROR_FILE_NOT_FOUND` tests (`SearchServiceBroker.cpp:1589`, `:1617`) and L39's dead branch. | about -60 lines | low | L58, L39 |
| S19 | Test code: one `DebugRequest` helper for the five speed-prompt blocks (`Popup.cpp:17677-17760`); move `#include "Helpers.h"` under `ENABLE_TESTS`; collapse the duplicated restoration counterfactuals (`ViewerPE.Tests.Runner.cpp:891-967`, `:1479-1522`); one foreground-raise route (the F24 handoff) instead of four `AttachThreadInput` raisers (L66(a)). | about -200 lines | low | L66 |

## Remediation slices

Execute in order; each slice is one reviewable change with its own focused validation. Respect the
scoped-testing policy (`Test-Changes.ps1 -Explain`, affected execution, `-Mode PrePush`) and record
each change as a `Changes/` fragment.

0. **Revision 2 data-path regression (H5, S13). Do this first; it is small.**
   - Size `kDeleteStreamContextSlots` from `kFileOpsSchedulerWorkerSlots`, never fall back to slot 0
     (assert or fail the step instead), and replace the remaining `16` literals.
   - Add a self-test that forces a 17-worker pool through a test seam and runs a multi-item Local
     Delete and Recycle at concurrency 16. The test must assert that no two concurrent steps share a
     context.
   - Run the Local delete/recycle FileOps families and the Local plugin contract tests on a machine
     with 16 or more logical CPUs.
   - Exit: the new case fails before the fix and passes after it, and the families are green.

Revision 2 additions to the slices below:

- **Slice 1:**
  - H6: print the `ForegroundAdmissionDiagnostics` on every failure first, then fix or reclassify per
    Q16.
  - M36: retain WER LocalDumps for harness children, decide fixture versus product, then Q17.
  - M40: step timeouts. M41: no cancel on `main` pushes until I55 batching. M43: per Q20.
  - M44: the fixture waits for its own listener-ready event; report `GetLastError`.
  - L66(d): R4A19 on a shared physical disk, per Q16.
  - M42 and H3: route to I55 per Q15 and Q22.
- **Slice 3:** add M35 (disable the radio items, keep focus and a visible state, per Q18), L54 and L55.
- **New slice 8a (plugin lifetime: L50-L52, S14, M36 product side):** move Shutdown after a successful
  `CanUnloadNow` probe, or make ViewerSpace's Shutdown re-armable. Add the UI-thread assertion and the
  file-system manager disabled reconcile.
- **Route to I51 (Ghostty owner):** M37 (reap before admission, with a 4096-garbage regression case),
  M38, M39 (Q21), L62 and S15.
- **Slice 9 (search):** L58-L61 and S18.
- **Slice 10 (tooling):** L63, L64 and S16-S17. Coordinate L64(a) and L64(c) with I53.
- **Slice 11:** add L53, L56, L57, L65, L66(a-c) and S19.

1. **Release and qualification blockers (H2, H3, H4, M28, M31, M32, M34, X1).** Coordinate with the
   owner of `ej/ci-runtime-repair-20261028` first so the work is not done twice. Add
   `'$(MSBuildProjectExtension)'=='.vcxproj'` to `ValidateMsvcToolsVersion` and a Pester check that
   evaluates the `.wapproj` (H3). Put the long-path checkout in one composite action used by every
   Windows job, with a policy test, or shorten the evidence paths (Q12) (H4). Fix the restoration
   predicate (drop the same-thread requirement, evaluate the no-write shortcut first, print GA_ROOT
   and focus facts) and add a bounded settle plus diagnostics to foreign-foreground restoration
   (H2, M28). Make the receipt check name its failed predicate and restore or exempt the DxUi source
   for the portable consumer (M31). Guard the Flake upload (M32). Pass the overlays in the inline
   vcpkg step and add `Build/vcpkg-ports` to the cache key (M34). Add a step timeout to the native
   suite and diagnose the silent x64 ASan FileOps run. Exit: scheduled workflows green on one `main`
   commit; Fresh Full passes for x64 Debug, Release and ASan Debug on that identity; `build.ps1 -Msix`
   produces a package.
2. **Host services (H1, M24, S1).** Add worker-thread self-tests for every `IHost*` entry point first
   (they must fail today for the three unrouted services), then collapse the marshaling. Exit: a
   worker-thread `SetConnectionSecret` round-trips through `GetConnectionSecret`; forged `lParam`
   values are ignored; `docs/dev/PluginHostModel.md` corrected.
3. **User-visible correctness quick wins.** M1 (pass `FILESYSTEM_FLAG_RECURSIVE` for viewer
   Delete/Recycle, with a host-side non-empty-folder test), M3/S3 (predicate), M13 (Running name),
   M10 (admitted item total in the interrupted card), M19 (delete the `_modeCombo` branch), M20
   (single count of workspace bytes), M21 (share root has no parent), L5 (translate or reuse the
   existing string).
4. **File Operations concurrency (M2).** Apply the `WakePauseWaiters` idiom (take the waiter's
   mutex after the flag store, before `notify_all`) at every listed site; add a stress self-test that
   hammers Pause/Resume/Cancel against a scheduler wait. Run the full FileOps families.
5. **Local provider (M4, M5, M6, M7, L36, L37, S4).** Data path: run every FileOps family and the
   Local plugin contract tests. Add cases: Overwrite of a Hidden file keeps Hidden; 230-character
   leaf Overwrite succeeds; `\\?\`-target symlink Preserve round-trips; copying a file being appended
   succeeds; the stage name contains no PID/TID.
6. **Leftover cleanup and plugin buffers (M14, M15, M16, M17, S2).** Shared validated walker, child
   name validation, `shared_ptr` scan context with a bounded destroy wait, an "unreadable folders"
   status line, conservative pre-check (Q5).
7. **File Operations UX (M9, M11, M12, L17, L18, L21).** Breadcrumb writer identity and once-per-
   process load (Q3), inline-rename pre-bind terminal results (Q8), Escape-binding visibility rule
   (Q4), notices exempt from the summary cap, plural strings.
8. **MTP (M8, L32, L33).** Memo only after a successful direct fallback, no ACCESS_DENIED, reset on
   reconnect, `SupportsAtomicWriterCommit` honest while set (Q9); guard the fake backend.
9. **FolderView and ViewerSpace resilience (M18, M22, L1-L14).** Per-generation stop source for
   enumeration (Q6); ViewerSpace reaper for cooperative workers instead of immediate detach.
   Search and throughput failures that the nightly reproduces (M29, M30, M33) belong here too:
   tracked-root traversal for the journal-unavailable rebuild, the pipe-busy connect policy, and
   per-item pool sizing.
10. **Tooling (M25, M26, M27, L40-L42, S8).**
11. **Remaining Low items and simplifications S5-S11**, one area per change.
12. **Repo hygiene (L43, L44)** after the owner answers Q12.

## Questions for the owner

Answers change what the slices do; until answered, the slice keeps the current behavior.

1. **Release gate.** Is "production-ready" defined as green scheduled workflows plus a Fresh Full
   in x64 Debug, Release and ASan Debug at one `main` identity (slice 1 exit)? Do native ARM64
   execution and interactive suites belong to that gate or stay separate handoffs (as I48 records)?
2. **Host services (H1).** Should a host service ever report success for an unrouted call, or should
   the marshal fail closed (`E_NOTIMPL`) so a future gap is visible?
3. **Move breadcrumbs (M9, L17).** Bind each record to its writer (process id + start time, or a
   per-process session id) and load notices once per process? Keep notices until dismissed,
   exempt from the 24-summary cap? Should `Finalize` try the removal before the Terminal write so a
   failed write cannot produce a false "interrupted" warning?
4. **Popup Escape (M12).** Should the sole-decision Escape fallback require the decision card to be
   visible (or scroll it into view when the decision is published)? Should a Partial result show
   a full bar and "100%" (popup report)? Is a modeless-in-practice Custom speed-limit prompt intended?
5. **Leftover cleanup (M15-M17).** Pre-check by name shape alone, or leave unchecked anything modified
   in the last N minutes (or require a cross-window lease)? When the walk is incomplete, keep Delete
   Selected enabled (report only) or disable it? Is an unbounded UI wait at window destroy for a
   wedged remote walk acceptable?
6. **FolderView Cancel (M18).** Must Cancel abort the provider enumeration (spec `UI_FolderView.md:514`),
   or is "abandon the result, keep the worker busy" intended (then the spec changes)?
7. **ViewerSpace.** Escape in the analysis panel: hide the panel or cancel/close the viewer? File-count
   weighting: should eligibility and squarify order follow the weight? Up at `\\server\share`: disabled
   (assumed)? Should one malformed provider directory (control character, case-only duplicate in a
   case-sensitive folder) fail the whole scan or only that folder? Is permanent unload quarantine after
   a mid-scan close acceptable, or adopt a reaper? Export of an unrepresentable name: fail the export or
   substitute U+FFFD with a flag?
8. **File Operations results.** Inline Rename pre-bind refusals: a terminal "not renamed" result like
   Batch Rename (assumed)? Should skipped and partially failed roots count in the live "N of M items"
   meter? Is the consent-receipt table meant to become an audit record, or delete it?
9. **MTP (M8, L32).** Should `ERROR_ACCESS_DENIED` on the hidden temp count as "device cannot rename"?
   Session-wide memo, or per folder / re-learned after reconnect? Must creates get the same journal and
   replay guarantee as overwrites?
10. **Local provider.** Does `GetFileInformationByHandleEx(FileIdInfo)` succeed on FAT32/exFAT/UDF
    media? If not, native Copy from and Overwrite onto USB sticks fail like the NAS case; this needs a
    manual check with a real stick. Should one malformed directory record fail the whole listing
    (`DirectoryRecordCursor`) or be skipped with a Warning? Should the spec's "no predictable stage
    name" rule govern the Local provider (assumed yes, M5)?
11. **Tooling and hosted runners.** Retention and reclamation for
    `X:\RedSalamander.Perf\evidence\completed-runs` (kept forever today, including detached VHDX
    files)? Should interactive Commands/Viewer cases run on hosted runners (1024x768, Windows Terminal
    in the foreground on ARM64 Windows 11), or move to a self-hosted desktop? On a hosted runner, is an
    unrestorable foreign foreground with unchanged input a failure that aborts the Viewer suite, or
    inconclusive evidence? Should `-NonInteractive` on legacy suites exit 2 as documented? Should MSI and
    MSIX carry build receipts and test sidecars, or use the ZIP allow-list (L48)?
12. **Public repository content (H4, L43).** Must the repository stay cloneable on stock Windows git
    (`core.longpaths=false`)? If yes, the 607 long evidence paths must be shortened or moved. Should the
    public repo carry the 19,500 evidence files under `Specs/TestRuns`, `advisor-plans/` and
    `.codex/config.toml` (which grants agents network access)? Should the private NAS hostname be
    replaced by a fixture name in published plans and evidence?
13. **Shell.** Should direct-mutation commands (Change Attributes, Create Directory, Edit/New File,
    Pack/Unpack, Make File List) also be refused while an accepted exit is pending? Can a Compare pane
    host an embedded Terminal (its exit notification goes to the Compare window, which has no handler)?
    Should the Preferences "Monitor" page keep the "RedSalamanderMonitor" name? Should the RedMonitor
    settings path follow the viewer's version instead of the fixed `-7.0` file?
14. **Diagnostics (L38).** Truncate ETW messages over 16,000 code units with a marker, or drop and count?

Revision 2 notes on earlier questions:

- **Q1, partly answered by I55 (owner-selected 2026-10-29).** The automatic gate becomes test-enabled
  Release on native x64 and ARM64, then production Release. Debug and ASan become manual qualification.
  Still open for this review: does "production-ready" also need one manual Debug and ASan Fresh Full at
  the same `main` identity, and the ARM64 interactive suites (now blocked by H6)?
- **Q11, partly answered by I55.** The scheduled native, ASan and flake workflows are to be removed.
  The hosted-runner question remains for the batch and manual qualification lanes that replace them.

Revision 2 questions:

15. **MSIX and MSI (H3, L48).** I55 §9 publishes ZIP only. Should H3 be fixed (filter
    `ValidateMsvcToolsVersion` to `.vcxproj`), or should `build.ps1 -Msix`, the MSI script and the
    `release.yml` MSIX job be retired? Until one of those happens, `-Msix` cannot build.
16. **Hosted ARM64 foreground (H6, L66(d)).**
    - The hosted ARM64 image keeps a system UWP CoreWindow (`WWAHost`) in the foreground. Should
      interactive Commands and ViewerSpace cases run there at all, or move to a self-hosted desktop?
    - May the F24 handoff displace that window?
    - Should a refused handoff be classified as environment/inconclusive instead of REGRESSION?
    - Should R4A19's independent-volume control skip, with an explicit reason, when its VHDX shares
      the physical disk with the primary sandbox?
17. **Viewer module unload (M36, L50).**
    - Must every viewer that hosts DxUi UIA providers export the quiet point
      (`RedSalamanderPluginShutdown`/`CanUnloadNow`), or should the host retain viewer modules until
      process exit?
    - Should the host call Shutdown only after `CanUnloadNow` reports quiet (today the order is
      reversed), so a live instance never sees its module state shut down?
18. **Batch Rename (M35).** While a preview rebuild is pending, should Rules/Manual stay enabled and
    queue the switch, or be visibly disabled at the item level with focus kept on the selector?
19. **Main CI cancellation (M41).** Until I55's five-minute batching lands, should a push to `main`
    keep a complete gate per commit (no cancel-in-progress), at the cost of runner time?
20. **Nightly Flake (M43).** Retire it now (I55 removes scheduled stress), cut it to a scope that fits
    210 minutes, or raise the timeout?
21. **Ghostty upgrade review (M39).** Must every category the review assigns (`clipboard-paste`,
    `kitty-image`, `snapshot-serialization`) carry at least one written disposition? That would
    restore a non-empty findings list whenever such commits exist.
22. **Release runner (M42).** Fix `release.yml` now to use the qualified `windows-2025-vs2026` and native
    ARM64 images, or wait for I55's production-Release-after-test design?

## Out of scope and stop conditions

- This plan authorizes no push, PR, release, workflow re-run or change in another repository
  (DxUi, RedMonitor). Each slice is executed only after the owner picks it.
- Stop if the drift check shows an in-scope file changed and the cited excerpt no longer matches;
  re-verify the finding first.
- Stop a data-path slice (4, 5, 6, 8) if any FileOps family regresses; preserve the failed run.
- Do not relax timeouts, delete failing tests or weaken focus/sandbox guards to make slice 1 green.

## Done criteria

- Every High and Medium finding (revisions 1 and 2) is fixed with a failing-first regression case, or
  explicitly accepted by the owner with the answer recorded here. CI, harness and tooling findings
  routed to I51, I52, I53 or I55 count as closed here only when that owner records the fix.
- Slice 1 exit holds at one `main` identity, with the receipts archived under `Specs/TestRuns`.
- Durable rules found here are merged into their owning specs (`Plugins_PluginAPI.md` marshal rule,
  `FileSystem_FileOperations.md` breadcrumb ownership and stage naming, `UI_FileOperationsPopup.md`
  Escape rule, `UI_FolderView.md` Cancel, `Plugins_ViewerSpace.md` answers, `Core_StartupBootstrap.md`).
- Then move this plan to `Specs/Plans/Done/` through the reviewed Done-admission steps and remove its
  index row.

## Evidence

Per-area reviewer reports and the lead verification ledger were working files for this review and
are not retained in the repository; every finding above carries its own file and line references.
Revision 2 cites GitHub Actions runs by id. Their logs and artifacts expire under the repository
retention policy, so the quoted counts and messages above are the durable record.
