# Operation Granite — 5-Day Master+Branch+Worktree Review Remediation (2026-07-02)

Tracking plan for findings from the max-effort multi-agent review of the last five days of work:
range `acfdc1a4e` (master@2026-06-27) → **current working tree** on `codex/folderview-warpdrive`
(HEAD `275c04034`, master `4778dd4f9`, plus uncommitted changes). Scope covered: the Bedrock DxUi
merge, the FS deep-audit/safety fixes, the MTP/PTP plugin merge, the FolderView WarpDrive perf
work, the Iron Ledger data-safety work, and all uncommitted edits (~103k diff lines;
`Specs/TestRuns/` excluded).

Method: 10 finder angles (5 correctness + reuse/simplification/efficiency/altitude/conventions) →
69 candidates → dedup → 1-vote 3-state verification (13 verifier batches, 46 verdicts) → fresh
gap sweep (4 more candidates, all confirmed). Every item below is **CONFIRMED** (survived an
independent verifier who was told to refute it) unless marked PLAUSIBLE. Line anchors are
verifier-corrected against the working tree at plan time.

> **Planned-at anchors.** Working tree of 2026-07-02, branch `codex/folderview-warpdrive`
> HEAD `275c04034`, master `4778dd4f9`, review base `acfdc1a4e`. GR-1 lives in **uncommitted**
> changes; GR-2/GR-3 were introduced by branch commit `b9af84d72`; the CW-7 revert entered via
> merge `d4a064dce`. Before executing any item, re-check drift with
> `git diff <anchor-file>` — this ledger does not auto-track the tree.

> **Addendum 2026-07-04.** A max-effort local re-review of the **merged**
> `codex/folderview-warpdrive` (merge `6e6ff863b`, review base `45ae2a9a`) was run after
> this pass. Most of its findings were already tracked here (they map to GR-2, GR-3,
> TW-2/TW-3, GR-18/TW-4, GR-A1, GR-S4b/d — see the reconciliation table in
> **"Granite addendum 2026-07-04"** below). Seven genuine deltas were added as
> **GR-20, GR-21, GR-22, GR-23, GR-24, GR-A6, GR-S4(f)** with tests **GR-T20–GR-T24, GR-TA6**.
> All are CONFIRMED at `6e6ff863b` with verifier-exact anchors.

> **Audit correction 2026-07-04 (Codex static re-check).** Re-checked this ledger against
> current `master` at `6e6ff863b` with the local 2026-07-04 addendum still dirty. Corrections:
> (1) GR-1 is stale in the current tree because the one-shot reset is present at
> `NavigationView.Edit.cpp:765` and `NavigationView.Interaction.cpp:808`; do not execute it
> without a fresh repro. (2) The UIA addendum count is **34 direct**
> `RefreshWindowHostAccessibilitySnapshot` call sites plus **32 wrapper**
> `RefreshAccessibilitySnapshot` call sites (37 textual wrapper occurrences including
> declarations/definitions), not 43. (3) GR-4's backend mismatch is confirmed by static
> backend code, but the cited public selftest anchors are weaker evidence because they call
> public `FileSystemMtp::MoveItem`, which may route through core temp-swap. (4) Several anchors
> are shorthand paths; resolve them through the current tree before patching, e.g.
> `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
> `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp`,
> `RedSalamander/IconCache.cpp`, and `Plugins/FileSystemMtp/Factory.cpp`. (5) This re-check was
> static only; RED statuses in the new-test table remain proposed until implemented and run.

## Execution rules (inherited from Clearwater, still binding)

- This file is the **triage ledger**. Before dispatching an item, expand it into a self-contained
  sub-plan: planned-at SHA, drift check, in/out-of-scope file list, exact verification command,
  expected output, STOP conditions.
- Green check = `.\Tools\Run-AllTests.ps1 -Suite Full` / `-SkipBuild` (`AGENTS.md:134`,
  `README.md:204-214`). Every fix needs a **RED-on-bug** test first (GR-T table below).
- Perf-sensitive items (GR-P*) additionally require scenario definition, instrumentation,
  deterministic selftest coverage, and archived evidence under `Specs/TestRuns/`
  (`AGENTS.md:48-55`, `Specs/Testing/Testing_PerformanceValidation.md:55-64,156-166`).
- Closeout: move to `Specs/Plans/Done/` and merge durable rules into `Specs/<Domain>/` per
  `AGENTS.md:57-58`.

## Final closeout (2026-07-06)

Granite is complete and ready for archive.

- Current closeout audit found no live Granite rows with:
  `rg --pcre2 -n "^\| GR-[^|]+ \| (?!.*(DONE|STALE|SIGNED OFF|DECIDED|ROUTED))|^\| GR-T[0-9A-Za-z-]+ \| (?!.*(DONE|STALE))|^\| GR-TA[0-9A-Za-z-]+ \| (?!.*DONE)" Specs\Plans\WIP\Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`.
- Latest Full skip-build gate after the Granite fixes:
  `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run\run-all-tests-results.json`
  reports suite `Full`, exit code `0`, `1117` passed / `0` failed / `52` skipped /
  `1169` total, duration `2612504` ms.
- Durable behavior changes were merged into the authoritative specs named by the
  per-item closeout notes (`Specs/Core/*`, `Specs/FileSystem/*`, `Specs/UI/*`,
  and `Specs/Testing/*` as applicable).
- The remaining long-term GR-A5 fault-injection ratchet is intentionally owned by
  `Specs/Plans/WIP/Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`.
- With this note in place, move this plan to `Specs/Plans/Done/` and remove the
  live Granite row from `Specs/Plans/WIP/README.md`.

## Routed to existing plans — do NOT duplicate, execute there (with Granite addenda)

| Owner | Item | Granite addendum |
|-------|------|------------------|
| **CW-7** (Clearwater) | **DONE 2026-07-04.** TabControl closes a tab when the press did not start on its close X (`Common/DxUi/DxUi.Controls.cpp:5766-5771`). | Fixed per CW-7 by deleting the loose `\|\| (! closePressed.has_value() && hit.part == CloseButton)` term; CW-T7 added RED/GREEN coverage in `DxUiTests.Controls.cpp`; durable close-button press/release contract merged into `Specs/UI/UI_DxUiWinUIDesign.md`. GR-T19 merge audit completed below. |
| **CW-4** (Clearwater) | 7z archive unopenable on normalized-key collision (`Plugins/FileSystem7z/FileSystem7z.cpp:4461-4464`, `:4396-4400`). | Re-confirmed live on the working tree; `tar -uf` / appended-replacement zips are the everyday triggers. No addendum — execute CW-4/CW-T4 as written ("degrade, don't fail"). |
| **CW-10** (Clearwater) | **DONE 2026-07-05.** Search authorization cached `false` on transient `CreateFileW` errors (`Common/SearchServiceBroker.cpp`). | Fixed in Clearwater: service candidate directory authorization now caches `false` only for durable negative results, treats transient open failures as incomplete candidate skips with `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`, and re-evaluates the directory on a later query. RED archive `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_111008/`; GREEN archive `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_111440/`; adjacent SearchService checks passed in `2026-07-05_111529/`, `2026-07-05_111546/`, `2026-07-05_111603/`, and `2026-07-05_111619/`. |
| **CW-P2** (Clearwater) | `copiedEntries` manifest grows unbounded (`RedSalamander/FolderWindow.FileOperations.State.cpp:8377`, `RecordCopiedFile:7082`). | Two addenda: (a) `RecordCopiedFile` runs for plain **COPY** too, where the manifest is never consumed — gate recording on `task._operation == FILESYSTEM_MOVE` (the adjacent `#ifdef ENABLE_TESTS` block already checks exactly that); (b) `RecordCopiedFile`/`RecordCopiedDirectory` are `noexcept` around throwing `wstring` copies + `unordered_map` insert → `std::terminate` under memory pressure mid-transfer. The extract-and-erase consumption fix in CW-P2 stands. |
| **CW-P1** (Clearwater) | Cross-FS MOVE verify re-reads both sides in full (`State.cpp:7541`, `ReadersHaveEqualContent:7192-7247`). | Re-confirmed unconditional (no verify-level gate anywhere in the chain). Addendum: GR-P2 removed MTP reader-open whole-file materialization, but CW-P1's hash-during-copy fix is still required to avoid cleanup verify re-reading both sides after cross-FS MOVE. Execute CW-P1 in Clearwater after Granite's MTP viability slice. |
| **TW-2 / TW-3 / TW-D1** (Tailwind) | Provider-thumbnail machinery dead in Release; deadline race in `ExtractProviderAllowedThumbnailWithDeadline`; cached-only visible-path product decision. | Re-confirmed all three. **TW-D1 must be decided before merge to master**: Release users get generic icons for cold-cache .mp4/.mkv/.pdf/.svg where the pre-WarpDrive build produced thumbnails. Either sign off cached-only + background enrichment follow-up, or wire `ExtractProviderAllowedThumbnailWithDeadline` into the production cold-miss path (which also resolves TW-2's dead code and makes TW-3's race fix mandatory). |
| **TW-4** (Tailwind) | **DONE 2026-07-05 with GR-18.** Thumbnail pending-counter underflow via late-post double decrement. | Fixed in the GR-18 pass: late provider/unaccounted thumbnail bitmap payloads now carry `countsPending=false`, while normal worker posts still own exactly one pending increment/decrement. The shared regression case covers both stale-message and late-delivery mechanisms. |
| **DxUi_Uia_ContinuationBaton_2026-06-29** | UIA snapshot refresh plumbing / coalescing. | Granite confirmed the architecture problem quantitatively: 32 direct `RefreshWindowHostAccessibilitySnapshot` call sites + ~30 indirect via control wrappers, full-tree rebuild under a global recursive mutex on every keystroke/click/selection change, and `UiaClientsAreListening` appears **nowhere**. The baton's dirty-flag/debounce proposal is correct — add the listener gate to it (GR-A1) and note commit `275c04034` added 2 more manual sites (OnSize, ExecuteToggleOnWindowThread), proving the missed-site failure mode. **Addendum 2026-07-04, corrected by static re-check:** at merge `6e6ff863b` the count is now **34 direct `RefreshWindowHostAccessibilitySnapshot` call sites + 32 indirect wrapper `RefreshAccessibilitySnapshot` call sites** (37 textual wrapper occurrences including declarations/definitions; was previously misstated as 43) — the drift continued through the merge. `RegisterWindowHostAccessibilityTarget` is called unconditionally at window creation (`DxUi.WindowHost.cpp:1215`), so the eager full-tree `PublishWindowHostAccessibilitySnapshot` (`DxUi.Accessibility.cpp:394`) runs for **every** user regardless of `WM_GETOBJECT` — the `UiaClientsAreListening`/first-query gate is the load-bearing fix, not just the debounce. |

## Track P0 — fix before the next commit/merge (all in the last 5 days' work)

| ID | Sev | Item (anchor) | Fix direction | Proof |
|----|-----|---------------|---------------|-------|
| GR-1 | **STALE / CLOSED 2026-07-05** | **Path-edit focus trap** (historical uncommitted finding; current-tree evidence refutes the missing-reset premise). Static recheck confirms `_pathEditBlurSuppressActive = false` is present at both reset points (`NavigationView.Edit.cpp:765`, `NavigationView.Interaction.cpp:808`). Historical finding: *(Single owner — supersedes Tailwind TW-12 and TW-D2: same blur-suppression mechanism, escalated from UX-decision to confirmed focus trap.)* `NavigationView.Edit.cpp:760-765` was reported to remove the one-shot reset (`_pathEditBlurSuppressActive = false` inside the suppression window — present at base `:761`; same reset also reported removed from the WM_KILLFOCUS handler `NavigationView.Interaction.cpp:799-808`), while the 2000ms arming surface grew from one site (WM_SETREDRAW) to virtually every paint/layout: `UpdatePathEditHostLayout` (`NavigationView.cpp:1464`), `RefreshActiveEditHostAfterParentPaint` (`Edit.cpp:2431-2435`, runs via `wil::scope_exit` on every path-section paint per `NavigationView.Breadcrumb.cpp:34`), `SetPaneFocused` (`NavigationView.cpp:1420`). | Closed as historical drift. Only re-open with a fresh current-tree repro; if reproduced, restore one-shot semantics and shrink the arming surface as originally described. | GR-T1 STALE |
| GR-2 | **HIGH — DONE 2026-07-04** | **Async paste-shortcut silently overwrites a freshly created .lnk.** `FolderView.FileOps.cpp:877` (`TrySubmitThreadpoolCallback`) had no single-flight guard; `CreateShellShortcut` probed then saved with replace semantics, so racing same-stem paste-shortcuts could collapse to one `.lnk`. | Done: FolderView now serializes Paste Shortcut work with `_pasteShortcutInFlight` plus a pending-request queue, and `CreateShellShortcut` retries deterministic slots on collision. `SaveShellShortcutExactPath` now uses the same temp-file plus final no-replace move path for short and long paths. | GR-T2 DONE; see GR-2/GR-3 closeout note below. |
| GR-3 | **HIGH — DONE 2026-07-04** | **Paste-shortcut completion dropped when the pane navigated away.** *(Single owner — absorbs Tailwind TW-7-corrected: wire `result.generation` into the staleness check in the same pass; GR-3 makes invalidation/error-reporting unconditional, TW-7 gates only the view-refresh part.)* The old handler returned on folder mismatch before cache invalidation and failure UI. | Done: completion always emits worker perf, invalidates `result.targetFolder` when links were created, and reports failure via the Paste Shortcut pane alert. Only current-view refresh/focus is gated by target path and path-visit generation, so stale completions cannot navigate/focus an old folder while revisits see the created link. | GR-T3 DONE; see GR-2/GR-3 closeout note below. |
| GR-4 | **HIGH — DONE 2026-07-04** | **MTP real/fake backend parity break — tests certify behavior hardware rejects.** (a) `WpdMtpBackend` discards `allowOverwrite` in all four mutators (`FileSystemMtp.Device.cpp:1551/1604/1635/1684` — literal `static_cast<void>(allowOverwrite)`) and always enforces `EnsureDestinationDoesNotExist` (`:1626/1653/1702`); `FakeMtpBackend` honors it (`FakeBackend.cpp:518-520, 868-892, 818-821`). Core pre-handled overwrite via temp-swap for Write/Copy/Move (`Core.cpp:2460-2487, 2703-2727, 2798-2822`) **but not RenameItem** (`Core.cpp:2896-2898` forwarded raw) → rename+ALLOW_OVERWRITE replaced on fake, `ERROR_ALREADY_EXISTS` on device. (b) `FakeMtpBackend::MoveItem` = `RenameItemLocked` (`FakeBackend.cpp:642`): succeeded for directory, cross-device, and leaf-renaming moves that `WpdMtpBackend::MoveItem` rejects (`Device.cpp:1709-1725`: native move only when `sameDevice && preservesLeaf`, else `ERROR_NOT_SUPPORTED` for directories). Static backend evidence confirmed the mismatch. **Audit caveat 2026-07-04:** the cited public selftest anchors (`CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp:3288/5212/5369`) were not sufficient proof of every fake-only direct-move semantic because they call public `FileSystemMtp::MoveItem`, which may route through core temp-swap. | Done: public `RenameItem` now handles existing-destination overwrite through the same temp-sibling PUID swap used by device-sourced copy/move and calls backend rename with overwrite disabled when no destination exists. `FakeMtpBackend::MoveItem` now mirrors the WPD direct-move contract for destination-exists handling, directory no-transfer-fallback behavior, and file transfer fallback. Durable contract persisted in `Specs/FileSystem/FileSystem_Mtp.md`. | GR-T4a/b DONE; see GR-4 closeout note below. |
| GR-5 | **HIGH — DONE 2026-07-04** | **Revealed password left in freed heap.** `TextField::InvalidateSingleLineLayoutCache` (`Common/DxUi/DxUi.TextInput.cpp:3152-3155`) called `ClearSingleLineTextLayoutCache(_singleLineLayoutCache)` with default `secureText=false`, while the single-line DirectWrite layout cache can hold revealed plaintext. | Done: TextField layout-cache invalidation now always uses the secure-clear path, and the shared `GetOrCreateSingleLineTextLayout` helper carries a `secureCacheText` flag so TextField failure paths wipe cached text while non-secret callers such as editable ComboBox keep the default plain clear. Other `ClearSingleLineTextLayoutCache` sites were audited. | GR-T5 DONE; see GR-5 closeout note below. |
| GR-8 | **HIGH — DONE 2026-07-04** (tooling) | **The documented FolderView perf gate exited 2 on every run.** *(Single owner — Tailwind TW-28 is the same Show-PerfRuns.ps1 sample-quality latch and routes here; GR-8 additionally covers the compare-mode baseline latch and the gauge-vs-distribution taxonomy. TW-28's per-metric `minimumSamples` lookup from FolderViewPerfBudgets.json5 is part of this fix.)* `Tools/Show-PerfRuns.ps1:448-452` `Register-SampleQuality` latched `$script:QualityFailure` for ANY metric group with < `$MinimumSamplesForP95 = 200` samples (`:24`), called unconditionally per group (`:476`). The `-FolderViewPreset` list (`:36-84`) includes one-shot gauges (`folder.cold_first_visit.*`, `folder.scale.working_set_bytes/private_bytes`) emitted exactly once per run (`Commands.SelfTest.ViewCommands.cpp:20376-20393`) → `-FolderViewPreset -FailOnQuality` (mandated by `Testing_PerformanceValidation.md`) always exited 2. In `-CompareRun` mode the **baseline** run also latched (`:565-566` before `:787`). Added in-range by `781af3463`. | Done: `Show-PerfRuns.ps1` now loads p95 `minimumSamples` from `Specs/Testing/FolderViewPerfBudgets.json5`, treats unbudgeted FolderViewPreset metrics and one-shot gauge/event rows as informational for preset quality gating, still fails low-sample p95-budgeted distributions, and registers compare-mode quality only for the candidate run. The updated semantics are persisted in `Specs/Testing/Testing_PerformanceValidation.md`. | GR-T8 DONE; see GR-8 closeout note below. |

## GR-1 stale closeout note (2026-07-05)

Scope: the historical path-edit focus-trap finding and its proposed GR-T1 test.

- Static recheck confirmed the old missing-reset premise is false in the
  current tree: `NavigationView.Edit.cpp:765` and
  `NavigationView.Interaction.cpp:808` both set
  `_pathEditBlurSuppressActive = false` after the one-shot focus-suppression
  window expires.
- No production change or RED test was added because there is no current-tree
  bug premise to reproduce. GR-T1 is retired as stale with this row; re-open
  only from a fresh current-tree repro.

## Track M — MTP plugin (correctness + viability)

| ID | Sev | Item (anchor) | Fix direction | Proof |
|----|-----|---------------|---------------|-------|
| GR-7 | **MEDIUM — DONE 2026-07-05** | **Stale overwrite journal replayed before every backend command, forever.** `ReplayOverwriteJournal` (`FileSystemMtp.Core.cpp:1314-1336`) retained destination-present/no-tempPUID journals when the orphan sweep found zero candidates, which is the expected completed-swap-then-crash state after the temp was renamed into the final path but before the host journal was cleared. Replay then ran before every backend command, causing repeated local journal probes, backend existence checks, and parent enumeration. | Done: replay now infers a completed swap when the final destination exists, the exact temp is gone, the no-tempPUID orphan sweep retained zero candidates, and the destination size matches `declaredSizeBytes`; it emits `mtp.overwrite.journal_replay_completed_swap_inferred` and clears the journal. Unresolved retained no-tempPUID replays persist `replayAttemptCount`, emit retained metrics, and quarantine exhausted stale journals to `.stale`. Pure read/metadata commands cache journal absence per device identity, while mutating commands opt out and journal writes invalidate the cache. Durable contract lives in `Specs/FileSystem/FileSystem_Mtp.md`; evidence is archived under `Specs/TestRuns/4cb089111a23/Mtp/2026-07-05_094500_gr_7_completed_swap_journal/`. | GR-T7 DONE |
| GR-P1 | **HIGH — DONE 2026-07-04** (perf) | **Stateless backend: every filesystem call re-discovers the device, opens a new session, re-enumerates every ancestor, and spawns a thread.** `WpdMtpBackend` originally only owned `_cancelState`; each call rediscovered devices, opened a WPD session, walked every path ancestor, and pre-fix `RunBackendCommand` spawned a fresh thread. Copying N files at depth D was approximately N×(2 CoCreateInstance + Open + D enumerations) + N threads. | Done: `RunBackendCommand` now uses one long-lived backend command queue worker per `FileSystemMtp` instance. `WpdMtpBackend` now caches WPD sessions/content by PnP id, upgrades cached access for mutating calls, memoizes normalized path→objectId resolution, invalidates affected path-cache subtrees on writes, and clears device/session/path caches on unexpected WPD resolution/session failures. The durable contract lives in `Specs/FileSystem/FileSystem_Mtp.md`; perf/test evidence is archived under `Specs/TestRuns/4cb089111a23/Mtp/2026-07-04_214700_gr_p1_worker_queue/` and `Specs/TestRuns/4cb089111a23/Mtp/2026-07-04_221500_gr_p1_wpd_session_path_cache/`. | GR-T-P1 DONE |
| GR-P2 | **HIGH — DONE 2026-07-05** (perf/memory) | **`CreateFileReader` buffers entire files in RAM.** `Core.cpp:2343-2385`: `backend.ReadFile(commandPath, *bytesResult)` pulled the whole file (`ReadPortableDeviceStream` chunk loop `Device.cpp:1115-1157`, no `reserve`, ~2× transient during growth) then wrapped it in `MtpMemoryReader` — the plugin's only IFileReader. First byte was gated on full USB transfer; 2 GB video = 2 GB+ resident; it also had to finish within the watchdog timeout. | Done: `CreateFileReader` is streaming-first. Backend readers keep WPD stream/content/session ownership alive, forward `Read`/`Seek`/`GetSize` through the serialized backend command queue, and emit read-byte accounting only when bytes are actually read. Non-seekable WPD streams fall back to a memory-backed reader to preserve the public seek contract. Durable contract lives in `Specs/FileSystem/FileSystem_Mtp.md`; evidence is archived under `Specs/TestRuns/4cb089111a23/Mtp/2026-07-05_090500_gr_p2_streaming_reader/`. | GR-T-P2 DONE |

## Track S — search service (execute together with CW-10)

| ID | Sev | Item (anchor) | Fix direction | Proof |
|----|-----|---------------|---------------|-------|
| GR-11 | MEDIUM (latent) — **DONE 2026-07-05** | **Rebuild/invalidate rejected for deleted roots — stale index entries unpurgeable.** `HandleRebuildRequest` used the exact-root query authorization path, so deleted indexed roots failed `RequestRebuild(root)` with `ERROR_FILE_NOT_FOUND` before `InvalidateRoot` could purge stale entries. | Done: query authorization remains exact-root, while rebuild/invalidate now normalizes the requested root, rejects malformed/device/UNC roots, and on `ERROR_FILE_NOT_FOUND` / `ERROR_PATH_NOT_FOUND` impersonates the named-pipe client to ACL-check the deepest existing ancestor before invalidating the original normalized root. Durable contract lives in `Specs/Core/Core_Search.md`; GREEN evidence includes `.build\logs\search_service_rebuild_deleted_root_gr11_green_attempt1.txt`, adjacent SearchService logs, archived CompareDirectories runs `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_095517/`, `_095657/`, `_095746/`, `_095836/`, `_095853/`, and Debug x64 build `.build\logs\msbuild-20260705_095307_267.log` (`0 warning(s)`, `0 error(s)`). | GR-T11 DONE |
| GR-13 | LOW-MED — **DONE 2026-07-05** | **New server-side root-validation errors are not in the client's fallback list.** `IsServiceFallbackCandidate` lacked `ERROR_BAD_PATHNAME` / `ERROR_PATH_NOT_FOUND` / `E_INVALIDARG`, so service root-validation refusals could hard-fail a plugin search instead of degrading to local INDEX/SCAN fallback. | Done: `IsServiceFallbackCandidate` now treats `ERROR_BAD_PATHNAME`, `ERROR_PATH_NOT_FOUND`, and `E_INVALIDARG` as fallback candidates. Durable contract lives in `Specs/Core/Core_Search.md`; RED evidence `.build\logs\plugin_contract_gr13_red.txt` failed FileSystem debug selftests (`passed=39, failed=3`), GREEN evidence `.build\logs\plugin_contract_gr13_green.txt` passed (`passed=42, failed=0`). | GR-T13 DONE |
| GR-14 | LOW-MED — **DONE 2026-07-05** | **Whole query aborts on one candidate's impersonation failure.** `CheckClientCanListDirectory` returned failed impersonation HRESULTs through `AuthorizeCandidateForClient`, and `StreamServerCandidate` treated them as fatal enumeration failures. | Done: candidate authorization failures now skip only that candidate, OR `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED` into query stats/runtime warnings, and complete the query. Added debug foreground-service fault injection `--test-fail-client-auth-impersonation-once` for deterministic coverage. Durable contract lives in `Specs/Core/Core_Search.md`; RED evidence `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_104801/` failed with `hr=0x80070558`, GREEN evidence `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_105520/` passed, adjacent SearchService checks passed in `2026-07-05_105549/`, `_105550/`, `_105551/`, `_105551_001/`, and `_105552/`; final Debug x64 build `.build\logs\msbuild-20260705_105305_813.log` had `0 warning(s)`, `0 error(s)`. | GR-T14 DONE |
| GR-S3 | LOW (architecture) — **DONE 2026-07-05** | **Stale store info masked by a one-shot query retry.** `LocalSearchIndexCore.cpp:4910-4932`: on sqlite failure with reason ∈ {StoreMissing, StoreInvalid, CutoverBlocked, StoreStale} AND zero emitted rows, refresh cache + re-run the whole query (`search.backend.sqlite.retry_query_ms`); a query that already emitted rows falls to slow soft-fallback; worst-case latency doubles exactly at rotation. `InvalidateCachedPersistentStoreInfo` exists (`:4473`) but only 2 in-process call sites — nothing covers external rotation. | Done: SQLite `meta.store_generation` is initialized during bootstrap/v2 migration, copied into `PersistentStoreInfo`, incremented by committed `ReplaceVolume`, `ApplyJournalDelta`, and existing-volume `DeleteVolume` transactions, and validated before direct SQLite query paths use cached store info. Generation mismatch/read failure refreshes the cache before query and emits `search.backend.sqlite.store_generation_refreshes`; stale/missing/invalid/cutover retry is now a belt-and-suspenders path that can run after emitted rows, with path-key dedupe across retry/live fallback. Durable contract lives in `Specs/Core/Core_Search.md`; RED evidence `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113058/` failed GR-T15 with post-rotation `retry_query_ms`, GREEN evidence `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113650/` passed without a post-rotation retry. Adjacent focused checks passed in `2026-07-05_114114/`, `_114131/`, and `_114218/`; live-journal-currentness guarded checks skipped in `_114043/`, `_114147/`, and `_114203/` because the live journal cursor was unavailable. | GR-T15 DONE |

## Track U — UIA/accessibility (snapshot model)

| ID | Sev | Item (anchor) | Fix direction | Proof |
|----|-----|---------------|---------------|-------|
| GR-9 | MEDIUM — **DONE 2026-07-05** | **Grid row names/cells dropped horizontally scrolled-out columns.** Snapshot builder (`DxUi.Accessibility.cpp:1072-1076, 1106-1120, 1162-1175`) used `record.gridVisibleColumns` from `Grid::GetVisibleColumnAt` → `ComputeVisibleColumnSpan` (`DxUi.Grid.cpp:1479-1501, 4140-4184`) which is viewport-clipped by `_horizontalScrollDip`. Deleted `BuildGridRowAccessibleName` iterated ALL model columns. Served verbatim at `:4519`. Narrator read incomplete rows; per-cell navigation returned nothing for off-view columns. | Done: accessibility snapshots now keep viewport-clipped `gridVisibleColumns` for headers/point-hit behavior and a full-model `gridAccessibleColumns` list for row names, row-owned `GridCell` records, `FindSnapshotGridCellRecord`, and row cell sibling navigation. Off-view cells remain offscreen for UIA while visible headers/hit-testing stay viewport-clipped. Durable contract lives in `Specs/UI/UI_DxUiSharedGrid.md`; RED evidence `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` failed at `scrolled grid row name includes all model columns, including horizontally off-view cells` after build log `.build\logs\msbuild-20260705_115326_983.log`; GREEN evidence after fix passed `Accessibility` after build log `.build\logs\msbuild-20260705_115509_020.log`. | GR-T9 DONE |
| GR-10 | MEDIUM — **DONE 2026-07-05** | **Point hit-testing uses unclipped content-space rects.** `ElementProviderFromPoint` (`DxUi.Accessibility.cpp:4980`) → `FindSnapshotPointHit` (`:878-889`) first-match over flat records appended with raw `GetHitBounds()` and no ancestor intersection (`:829-876`). Old path required `PointInRect` at every level (`FindSemanticControlAtPoint:683-719`, now dead — CW-S1 wants it deleted). ScrollPanel children are recorded at content-space rects (`DxUi.Controls.cpp:6673-6703` paints translated by `-_scrollOffsetDip`) → scrolled-out controls hit-testable in empty dialog area, and scrolled panels resolve hits at displaced coordinates. | Done: snapshot point-hit building now carries an accumulated window-space translation and ancestor viewport clip, applies `DxUi::ScrollPanel` `-GetScrollOffset()` to descendants, clips tree/grid/control/password-reveal hit rectangles before appending them, skips empty clipped records, and appends a root fallback after child records. Empty clipped viewport areas now resolve the root provider while visible scrolled children resolve at their viewport-translated positions. Durable contract lives in `Specs/UI/UI_DxUiSharedGrid.md`; RED evidence `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` failed at `scrolled panel empty viewport provider is the root provider hr=0x80004002` after build log `.build\logs\msbuild-20260705_120301_879.log`; GREEN evidence after fix passed `Accessibility` after build log `.build\logs\msbuild-20260705_120534_990.log`. | GR-T10 DONE |
| GR-6 | MEDIUM — **DONE 2026-07-05** | **Cross-thread UIA dispatch writes through popped stack on timeout.** `SendMessageTimeoutW(..., SMTO_ABORTIFHUNG \| SMTO_BLOCK, 5000, ...)` with stack request structs + raw provider pointers (`DxUi.Accessibility.cpp:4076-4077, 6314-6315`; handler writes at `:7057-7099`). Win32 nuance (verifier-corrected): an **undelivered** timed-out message is revoked — the unguarded window is a message the window thread has *dequeued* but not finished when the sender gives up (handler stalls >5s on `GetAccessibilityTargetMutex` or DirectWrite). Then the RPC frame pops, UIA may Release the provider, and the handler completes into freed stack. Same pattern in the test-only `DebugGetContextMenuPopupState` (`DxUi.Menu.cpp:5294-5298`, 1s timeout — GR-17). | Done: UIA action dispatch now posts a heap-owned shared dispatch block through `PostMessagePayload(...)`, keeps the element/text-range provider alive with `wil::com_ptr_nothrow`, owns rectangle output storage inside the request, waits on a completion event, and returns `ERROR_TIMEOUT` without reading late output when the sender times out. `WindowHost` initializes/drains the posted-payload registry for this message path. The test-only context-menu debug state query now uses synchronous `SendMessageW(...)`, folding GR-17. Durable contract lives in `Specs/UI/UI_DxUiSharedGrid.md` and `Specs/UI/UI_DxUiWinUIDesign.md`; RED evidence: Accessibility source guard failed at `accessibility UIA action shared dispatch helper is found` after build log `.build\logs\msbuild-20260705_121657_689.log`, and Menu source guard failed at `context menu debug state cross-thread query does not time out while using caller stack output storage` after build log `.build\logs\msbuild-20260705_122109_640.log`; GREEN evidence: build log `.build\logs\msbuild-20260705_122720_807.log` passed with 0 warnings / 0 errors and `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` passed with the late-handler timeout regression; focused `Menu` and adjacent `WindowHost` also exited 0, `git diff --check` exited 0 with only LF-to-CRLF warnings, `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5/0, and final Debug x64 app build `.build\logs\msbuild-20260705_123200_423.log` passed with 0 warnings / 0 errors. | GR-T6 DONE |
| GR-16 | LOW — **DONE 2026-07-05** | **Root-collapse criterion diverges between snapshot and retained paths.** Snapshot collapsed on `semanticControlOrder.size()==1` of ANY type incl. Label (`:1247`, `IsSemanticAccessibilityControl:555`); retained path gated on `ShouldExposeSingleSemanticRootControl` (`:650-657`) which excludes Label. A Label-only window reported a different UIA tree shape depending on which path served. | Done: published accessibility snapshots now carry `hasCollapsedSemanticRoot`, computed through `TryResolveSingleSemanticRootControlPath(...)` and therefore the same `ShouldExposeSingleSemanticRootControl(...)` predicate as the retained path. Snapshot navigation/properties consult that flag instead of collapsing any one semantic control, so Label-only roots expose the label as a child. Durable contract lives in `Specs/UI/UI_DxUiSharedGrid.md` and `Specs/UI/UI_DxUiWinUIDesign.md`; RED evidence `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` failed at `label-only root exposes the label as a child instead of collapsing it into the root` after build log `.build\logs\msbuild-20260705_123529_940.log`; GREEN evidence `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` passed after build log `.build\logs\msbuild-20260705_123625_707.log`. | GR-T16 DONE |
| GR-18 | LOW — **DONE 2026-07-05** | **Thumbnail pending-counter decremented by stale-batch messages.** `OnCreateThumbnailBitmap` registered the decrement `scope_exit` before the batchId/generation checks, so stale payloads could drive the current batch's `pendingBitmapCreates` to 0 while applies were queued. TW-4 was the sibling mechanism: late current provider-delivery payloads posted after deadline did not own a matching pending increment but still decremented. Impact: ENABLE_TESTS thumbnail diagnostics only. | Done: `OnCreateThumbnailBitmap` now returns on stale batch/generation before registering the decrement, and `ThumbnailBitmapRequest::countsPending` makes pending ownership explicit. Normal worker posts keep the default counted path; abandoned late provider delivery and the selftest's unaccounted current delivery use `countsPending=false`. Durable contract lives in `Specs/UI/UI_FolderView.md`; RED evidence: stale half failed with `pending=0 expected=2 staleDrops=1` after build `.build\logs\msbuild-20260705_124514_955.log`, and the TW-4 extension failed with `pending=1 expected=2 completed=1` after build `.build\logs\msbuild-20260705_125655_971.log`; GREEN evidence: focused case passed after build `.build\logs\msbuild-20260705_125952_183.log`, then rebuild `.build\logs\msbuild-20260705_130500_471.log` passed with 0 warnings / 0 errors and focused/adjacent thumbnail guards passed from `.build\logs\gr18_tw4_*.out.txt`. | GR-T17 DONE |

## Track A — architecture / duplication (the MTP identity cluster + misc)

| ID | Sev | Item (anchor) | Fix direction |
|----|-----|---------------|---------------|
| GR-A2 | **MEDIUM — DONE 2026-07-05** | **ConnectionManagerWindow hosted a second full WPD stack + test fixtures in production source.** `ConnectionManagerWindow.cpp` had host-side `EnumerateMtpPickerDevices`, `EnumerateMtpPickerStorages`, `ReadMtpPickerDevicePersistentId`, COM init/client-info/disambiguation, and `#ifdef ENABLE_TESTS` picker fixture globals. The plugin was never consulted by `MtpPickerWorkerCallback`, so two WPD stacks could drift. | Done: picker browse now goes through the optional plugin factory export `RedSalamanderBrowseConnectionTargets(...)` keyed by `pluginId`; `FileSystemPluginManager` dispatches and parses device/storage JSON; the MTP plugin owns live WPD browse plus fake-backend selftest browse; `ConnectionManagerWindow` no longer includes `PortableDevice`/WPD picker helpers or host picker fixture APIs. Durable contract lives in `Specs/FileSystem/FileSystem_Mtp.md`; focused evidence is archived at `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_133530/`, and adjacent MTP Compare evidence is `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_133656/`. |
| GR-A3 | **MEDIUM — DONE 2026-07-05** | **MTP identity scheme exists as 4 bit-identical-required copies.** Case-folded FNV hash: `Core.cpp:572` (`StableDeviceHash`), `Device.cpp:190`, `FakeBackend.cpp:173`, host `ConnectionManagerWindow.cpp:941` (`StableMtpPickerHash`). Sanitizer: `Core.cpp:586`, `Device.cpp:254`, host `:955`. Suffix formats `' [puid:{:016X}]'`/`' [oid:{:016X}]'`/`'[devid:{:016X}]'`: `Device.cpp:858/861/207`, `FakeBackend.cpp:191/194`, host `:1243-1244`, `Core.cpp:605-612`. `JsonEscape` ×3 (`Core.cpp:673`, `Device.cpp:266+299`, `FakeBackend.cpp:137`). Host-composed paths are re-derived by the plugin at connect time — any drift = unresolvable storages, no compile error. | Done: GR-A2 removed host-side identity derivation; the remaining plugin identity hash, path sanitizer, device/object/duplicate suffixes, overwrite-journal hash formatting, and UTF-8 JSON escaping now flow through shared helpers declared in `FileSystemMtp.Internal.h` and implemented in `FileSystemMtp.Shared.cpp`. Core, live WPD, and fake backend local clones were removed. The source-contract selftest `mtp_identity_helpers_are_shared` guards the shared declarations/implementations and rejects the old clone symbols. Durable contract lives in `Specs/FileSystem/FileSystem_Mtp.md`; focused evidence is archived at `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_135203/`, and adjacent MTP Compare evidence is `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_135647/`. |
| GR-A4 | **LOW - DONE 2026-07-05** | **Locale-divergent hand-rolled string predicates.** `EqualsNoCaseLocal` (`ConnectionManagerWindow.cpp:967`) and `EqualsOrdinalIgnoreCase` (`Device.cpp:210`) used CRT `towlower` loops that disagreed with `OrdinalString::EqualsNoCase` (`Common/Helpers.h:121`, CompareStringOrdinal) for non-ASCII device names; both files already included Helpers.h. `StartsWithNoCase` was re-implemented at `FileSystemMtp.Shared.cpp:12` and `SearchServiceBroker.cpp:2167` despite `OrdinalString::StartsWithNoCase` (`Helpers.h:126-144`) in scope. | Done: the targeted local predicates are gone or stayed gone, `FileSystemMtp.Device.cpp` routes equality and `CaseFoldKey` through `OrdinalString::EqualsNoCase` / `OrdinalString::FoldCaseInvariant`, and `SearchServiceBroker.cpp` routes prefix checks through `OrdinalString::StartsWithNoCase`. `TestHarnessSourceContracts.Tests.ps1` now guards all four named sites plus the `CaseFoldKey` helper. Durable MTP identity comparison contract lives in `Specs/FileSystem/FileSystem_Mtp.md`; existing search root validation contract remains in `Specs/Core/Core_Search.md`. RED/GREEN Pester, Debug build, and focused Compare archives are recorded in the GR-A4 closeout note below. |
| GR-A5 | **MEDIUM (long-term) - DONE / ROUTED 2026-07-05** | **Fault-injection seams hand-woven through the production copy/move pipeline.** `FolderWindow.FileOperations.State.cpp`: 27 `ForSelfTest` occurrences, 30 `ENABLE_TESTS` blocks in one file (705 in the app project, 1009 repo-wide); `Sleep(1)` spin pause points with 5s bailouts (`:246-286`), env-var destination mutation (`:321-378`), fail-next counters (`:216-244`), invoked inline at `:7022/:7862/:8381/:8395/:9342/:9936`. Every new data-safety test adds a global+env-var+call-site; injection points silently rot on refactor. | Done for Granite: the interface-boundary seam now exists as a test-only `IFileSystemIO` / `IFileReader` decorator injected where the bridge acquires its IO objects, and the first concrete source-size hook was migrated. Remaining opportunistic hook migration is routed to `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`. | Ratchet slice: the cross-FS bridge source `GetSize` fail-next hook now lives in test-only `IFileSystemIO`/`IFileReader` decorators injected at bridge IO acquisition; inline `CopyFileWithBuffer(...)` consumption was removed. RED/GREEN Pester, Debug build, and focused FileOps archives are recorded in the GR-A5 ratchet note below. The remaining long-term count ratchet is owned by `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`. |
| GR-S1 | **LOW — DONE 2026-07-05** | **Selftest pause-point machinery cloned (~50 lines per quiet-point).** `MaybePauseAfterTaskFinishedBeforeSummaryForSelfTest` (`State.cpp:246-265`) vs `MaybePauseBeforeBridgeMoveSourceCleanupForSelfTest` (`:267-286`) byte-identical bodies; globals `:54-56` vs `:57-59`; Set/Has/Release triads `Queue.Part.cpp:144-167` vs `:169-192`; header decls ×2. | Done: one `SelfTestPausePoint` struct owns the enabled/entered/release atomics plus `Set/HasEntered/Release/Pause(bailoutMs)`, and the two existing quiet points are now two instances. This closes the concrete GR-S1 duplication slice under GR-A5's ratchet; the broader GR-A5 interface-boundary decorator seam is routed to `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`. |
| GR-S2 | **LOW - DONE 2026-07-05** | **32 `#ifdef ENABLE_TESTS` trace sites frozen into one modal loop; 14 hand-rolled `GetMessageW` modal loops repo-wide.** *(Anchor corrected 2026-07-02: the symbol `TraceArchivePackPrompt` does not exist — the real machinery is the `g_archivePackPromptWindow`/`ArchivePackPromptDebugCommand` ENABLE_TESTS trace blocks.)* The archive-pack prompt's loop is step-traced at every statement (`FolderWindow.FileSystem.Commands.Part.cpp:4578-4885+`, helpers `:3817-3826`); identical hand-rolled loops in ConnectionCredentialPromptDialog `:414`, FileOperations.Popup `:922`, ConnectionManagerWindow `:4869`, FolderWindow.FileSystem ×5, RedSalamander.cpp ×2, Commands.Part ×3 more. | Done for the concrete archive prompt slice: `Common/DxUi` now exposes `RunDxUiModalLoop(hwnd, options)` with one shared `GetMessageW` failure diagnostic and `WM_QUIT` repost propagation; `ArchivePackPromptWindow::ShowModal()` and `ArchiveUnpackPromptWindow::ShowModal()` route through it and no longer own local `GetMessageW(&msg, ...)` loops. `TestHarnessSourceContracts.Tests.ps1` guards the helper shape, archive prompt migration, and quit propagation. Other hand-rolled modal loops remain opportunistic migrations. Durable contract lives in `Specs/UI/UI_DxUiSharedGrid.md`; RED/GREEN Pester, Debug build, focused runtime archives, and preserved hung-run evidence are recorded in the GR-S2 closeout note below. |
| GR-S4 | **LOW - DONE 2026-07-05** (batch of quick wins, from finder pass — re-verify inline while fixing) | (a) **DONE 2026-07-05:** the duplicated issues-pane focus-restore block now routes through `RestoreActivePaneFolderViewFocusIfWindowHadFocusBeforeHide(...)`. (b) **DONE 2026-07-05:** `IsTruthySelfTestEnvironmentVariable`/`ShouldForceFolderViewWarpDevice` duplication is gone; FolderView WARP/perf env gates now route through shared `EnvironmentVariables::IsTruthyFlagSet(...)`. (c) **DONE 2026-07-05:** `MaybeInjectBridgeCreateDirectoryRaceForSelfTest` now uses the existing `TryReadEnvironmentVariableForSelfTest(...)` helper instead of hand-rolling the env read. (d) **DONE 2026-07-05:** `PasteShortcutWork` and `ProviderAllowedWork` now execute through shared `SubmitOwnedThreadpoolCallback(...)` instead of duplicating the owned-threadpool-submit dance. (e) **DONE 2026-07-05:** `IsViewerSpace20kLayoutEnabled` and `IsViewerSpaceLargePerfEnabled` now share `IsViewerSpaceEnvFlagEnabled(name)`. (f) **DONE 2026-07-05:** WarpDrive maintenance sprawl is reduced by shared shell-thumbnail stat increments, the earlier GR-A6 shared D3D11 creation helper, and one generic pending-to-paint metric shape. | All slices (a), (b), (c), (d), (e), and (f) are closed with RED/GREEN Pester, Debug builds, focused runtime evidence, and durable spec updates in the GR-S4 notes below. |

## GR-A2 MTP picker ownership closeout note (2026-07-05)

Scope: the Connection Manager MTP picker host path, the plugin-manager browse dispatch,
the optional MTP factory browse export, fake-backend picker seeding, deterministic
Commands coverage, adjacent MTP Compare coverage, and the durable MTP/test coverage
specs.

- RED `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_connection_manager_window_mtp_picker_populates_profile -FailFast -TimeoutMultiplier 3` failed with `Failed to seed the MTP plugin fake picker browse backend. hr=0x8007007F` before the plugin debug fake-browse setter/export existed.
- GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors after the GR-A2 changes; build log `.build\logs\msbuild-20260705_133324_146.log`.
- GREEN focused Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_133530/` passed `cmd_connection_manager_window_mtp_picker_populates_profile` with 1 passed / 0 failed / 0 skipped and recorded `mtp.connection_browse.devices`, `mtp.connection_browse.devices_us`, `mtp.connection_browse.storages`, and `mtp.connection_browse.storages_us`.
- GREEN plugin contract check `.\.build\x64\Debug\PluginContractTests.exe` passed.
- GREEN adjacent MTP Compare archive `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_133656/` passed with 50 passed / 0 failed / 1 skipped; the skipped case was the explicit no-device `mtp_live_device_smoke`.
- Static cleanup guard found no remaining host-side MTP picker WPD clone symbols in `RedSalamander\ConnectionManagerWindow.cpp`, `RedSalamander\ConnectionManagerWindow.h`, or `RedSalamander\SelfTest`.
- `Specs/FileSystem/FileSystem_Mtp.md`, `Specs/Testing/Testing_TestCoverage.md`, and `Tests/README.md` were updated so the durable contract says the Connection Manager picker browses through the MTP plugin, not a host-owned WPD stack.

## GR-A3 MTP identity helper closeout note (2026-07-05)

Scope: the MTP plugin identity/hash/suffix/sanitizer/JSON helper clones in Core,
live WPD Device, FakeBackend, shared declarations/implementation, source-contract
coverage, adjacent MTP Compare coverage, and durable MTP/test coverage specs.

- RED `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_identity_helpers_are_shared -FailFast -TimeoutMultiplier 3` failed before shared identity declarations existed with `MTP identity source guard: shared stable-hash declaration is missing`; archive `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_134733/`.
- GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors after the GR-A3 code changes; build log `.build\logs\msbuild-20260705_135001_254.log`.
- GREEN focused Compare archive `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_135203/` passed `mtp_identity_helpers_are_shared` with 1 passed / 0 failed / 0 skipped.
- GREEN adjacent MTP Compare archive `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_135647/` passed with 51 passed / 0 failed / 1 skipped; the skipped case was the explicit no-device `mtp_live_device_smoke`.
- Static cleanup guard found no remaining `StableDeviceHash`, local `StableHash`, `SanitizePathComponent`, local `JsonEscape`, or `DuplicateSuffix` clone symbols in `FileSystemMtp.Core.cpp`, `FileSystemMtp.Device.cpp`, or `FileSystemMtp.FakeBackend.cpp`.
- `Specs/FileSystem/FileSystem_Mtp.md`, `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and `Tools\Tests\TestInventory.Tests.ps1` were updated for the shared-helper contract and the new Compare/MTP test count.

## GR-S1 FileOperations self-test pause-point closeout note (2026-07-05)

Scope: duplicated FileOperations self-test pause-point state in
`FolderWindow.FileOperations.State.cpp`, the existing public pause helper APIs in
`FolderWindow.FileOperations.State.Queue.Part.cpp`, source-contract coverage,
focused FileOps runtime cases for both pause points, and test-inventory docs.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  first reported 35 passed / 1 failed / 0 skipped because
  `struct SelfTestPausePoint final` did not exist.
- GREEN Pester after implementation:
  `TestHarnessSourceContracts.Tests.ps1` passed 36 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors after explicitly deleting the
  helper's copy/move special members; build log
  `.build\logs\msbuild-20260705_161428_240.log`.
- GREEN focused FileOps archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_161924/` passed
  `Riptide_LiveFinishedSnapshotCarriesDiagnostics` with 3 passed / 0 failed / 0
  skipped, exercising the post-finished-completion pause.
- GREEN focused FileOps archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_161955/` passed
  `Floodgate_CrossFsMoveCleanupDetectsDestinationCorruption` with 3 passed / 0
  failed / 0 skipped, exercising the bridge move-source cleanup pause.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`,
  `Tools\Tests\TestInventory.Tests.ps1`, and `Specs/Plans/WIP/README.md` were
  updated for the added source-contract case and the GR-S1 closeout. GR-A5's
  broader interface-boundary decorator ratchet is now owned by
  `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`.

## GR-A5 FileOperations bridge IO decorator ratchet note (2026-07-05)

Scope: the cross-filesystem bridge source `GetSize` fault-injection hook,
`FolderWindow.FileOperations.State.cpp` bridge IO acquisition, source-contract
coverage, focused copy/move FileOps runtime coverage, and test-inventory docs.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  first reported 36 passed / 1 failed / 0 skipped because
  `enum class SelfTestBridgeIoRole` did not exist.
- GREEN Pester after implementation:
  `TestHarnessSourceContracts.Tests.ps1` passed 37 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_162741_277.log`.
- GREEN focused FileOps archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_162948/` passed
  `Floodgate_CrossFsCopyGetSizeFailureRefusesCommit` with 3 passed / 0 failed /
  0 skipped.
- GREEN focused FileOps archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_163017/` passed
  `Floodgate_CrossFsMoveGetSizeFailurePreservesSource` with 3 passed / 0 failed /
  0 skipped.
- The migrated hook now routes through `SelfTestBridgeIoDecorator` and
  `SelfTestBridgeFileReader`, injected with
  `DecorateBridgeIoForSelfTest(fileSystemIo, SelfTestBridgeIoRole::Source)`.
  `CopyFileWithBuffer(...)` now reads size through `reader->GetSize(...)`
  directly, so future bridge data-safety tests have an interface-boundary seam to
  extend instead of adding another inline global/call-site hook.
- Remaining GR-A5 work is long-term ratchet work and has been routed to
  `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`: migrate
  additional existing fault hooks when their pipeline sites are touched, and
  require new bridge fault-injection tests to use the decorator seam rather than
  adding new inline production-pipeline hooks.

## GR-A4 shared ordinal string helper closeout note (2026-07-05)

Scope: the targeted local case-folding predicates named in the GR-A4 row, the
remaining local MTP `CaseFoldKey` helper, source-contract coverage, focused MTP
and Search Compare runtime checks, and durable MTP/test coverage documentation.

- RED #1 `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru` reported 37 passed / 1 failed / 0 skipped because
  `EqualsOrdinalIgnoreCase` still existed in `FileSystemMtp.Device.cpp`.
- GREEN #1 after removing the named duplicate predicates passed 38 / 0 / 0.
- RED #2, with the old `CaseFoldKey` manual `::towlower` loop restored while the
  new source-contract assertion was present, reported 37 passed / 1 failed / 0
  skipped because `CaseFoldKey` did not use
  `OrdinalString::FoldCaseInvariant`.
- GREEN #2 after changing `CaseFoldKey` passed 38 / 0 / 0.
- GREEN app build `.\build.ps1 -Configuration Debug -Platform x64` passed with
  0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_164220_979.log`.
- GREEN focused Compare archive
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_164425/` passed
  `mtp_wpd_session_and_path_cache_reuse` with 1 passed / 0 failed / 0 skipped.
- GREEN focused Compare archive
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_164455/` passed
  `search_service_rejects_device_root_and_continues` with 1 passed / 0 failed /
  0 skipped.
- `ConnectionManagerWindow.cpp` and `FileSystemMtp.Shared.cpp` were already clean
  in the current tree; the new source-contract case keeps those previously
  cleaned sites from regressing. `FileSystemMtp.Device.cpp` now uses
  `OrdinalString::EqualsNoCase` / `OrdinalString::FoldCaseInvariant`, and
  `SearchServiceBroker.cpp` now uses `OrdinalString::StartsWithNoCase`.
- `Specs/FileSystem/FileSystem_Mtp.md`, `Specs/Testing/Testing_TestCoverage.md`,
  `Tests/README.md`, `Tools\Tests\TestInventory.Tests.ps1`, and
  `Specs/Plans/WIP/README.md` were updated for the shared-helper contract and
  the added Pester source-contract case.

## GR-S2 shared DxUi modal loop closeout note (2026-07-05)

Scope: the archive Pack/Unpack prompt modal loops in
`FolderWindow.FileSystem.Commands.Part.cpp`, the shared DxUi modal pump API,
source-contract coverage, focused archive prompt Commands runtime checks, and
durable DxUi/test coverage documentation.

- RED #1 `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru` reported 38 passed / 1 failed / 0 skipped because
  `struct DxUiModalLoopOptions final` did not exist.
- GREEN #1 after adding `RunDxUiModalLoop(...)` and routing the archive
  Pack/Unpack prompt loops through it passed 39 / 0 / 0.
- A focused pack prompt runtime then exposed a quit-propagation hazard: the
  Debug app process stayed alive and held self-test state after a wrapper run
  failed to collect `commands\results.json`. Evidence was preserved at
  `Specs/TestRuns/4cb089111a23/Continuation/2026-07-05_1700_gr-s2_archive_prompt_hang/`.
- RED #2 after adding the source-contract assertion for quit propagation reported
  38 passed / 1 failed / 0 skipped because `RunDxUiModalLoop(...)` did not
  repost `WM_QUIT`.
- GREEN #2 after adding `PostQuitMessage(static_cast<int>(msg.wParam))` passed
  39 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_170405_259.log`. A previous build attempt
  `.build\logs\msbuild-20260705_170252_113.log` failed only because a stale
  Debug `RedSalamanderSearchService.exe` locked `Common.dll`; that process was
  stopped before the green rebuild.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_170612/` passed
  `cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions` with
  1 passed / 0 failed / 0 skipped.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_170618/` passed
  `cmd_pane_unpack_prompt_uses_dxui_destination_unpacker_and_mask` with
  1 passed / 0 failed / 0 skipped.
- `Specs/UI/UI_DxUiSharedGrid.md`, `Specs/Testing/Testing_TestCoverage.md`,
  `Tests/README.md`, `Tools\Tests\TestInventory.Tests.ps1`, and
  `Specs/Plans/WIP/README.md` were updated for the shared modal-loop contract
  and the added Pester source-contract case. The broader repo-wide modal loop
  cleanup remains opportunistic after this archive prompt slice.

## GR-S4(a) FileOperations issues-pane focus helper ratchet note (2026-07-05)

Scope: the issues-pane hide/close focus-restore paths in
`FolderWindow.FileOperations.State.Diagnostics.Part.cpp` and
`FolderWindow.FileOperations.IssuesPane.cpp`, the shared private helper in
`FolderWindow.FileOperationsInternal.h`, source-contract coverage, focused
Commands runtime coverage, and test-inventory documentation.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` reported 42 passed / 1 failed / 0 skipped because
  `RestoreActivePaneFolderViewFocusIfWindowHadFocusBeforeHide(...)` did not
  exist and both call sites still carried local folder-view focus fallback
  chains.
- GREEN Pester after extracting the shared helper and routing both call sites
  through it passed 43 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_173537_689.log`.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_173744/` passed
  `cmd_pane_fileops_issues_pane_hide_restores_folder_focus` with
  1 passed / 0 failed / 0 skipped.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools\Tests\TestInventory.Tests.ps1` were updated for the added Pester
  source-contract case. GR-S4 remains open for subitem (f).

## GR-S4(b) shared truthy env-flag helper ratchet note (2026-07-05)

Scope: the FolderView WARP-forcing env flag reader in
`FolderView.Rendering.cpp`, the Commands FolderView perf/force-WARP env flag
readers in `Commands.SelfTest.ViewCommands.cpp`, the shared helper in
`Common\Helpers.h`, source-contract coverage, focused WARP runtime coverage, and
test-inventory documentation.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru` reported 39 passed / 1 failed / 0 skipped because
  `EnvironmentVariables::IsTruthyFlagSet(...)` did not exist and the local
  truthy-env helper clones were still present.
- GREEN Pester after adding the shared helper and routing all three FolderView
  perf/WARP gates through it passed 40 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_171253_615.log`.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_171549/` passed
  `folderView_perf_huge_folder_scale` with
  `REDSALAMANDER_FOLDERVIEW_FORCE_WARP=1` at 1 passed / 0 failed / 0 skipped.
  The archived perf artifact recorded `warpRunExecuted: true` and
  `warpRunStatus` as covered, proving the shared self-test env flag reader and
  the production FolderView WARP-forcing gate agreed.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools\Tests\TestInventory.Tests.ps1` were updated for the added Pester
  source-contract case. GR-S4 remains open for subitem (f).

## GR-S4(c) FileOperations env-helper reuse ratchet note (2026-07-05)

Scope: the create-directory race self-test hook in
`FolderWindow.FileOperations.State.cpp`, source-contract coverage, focused
FileOps runtime coverage, and test-inventory documentation.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru` reported 40 passed / 1 failed / 0 skipped because
  `MaybeInjectBridgeCreateDirectoryRaceForSelfTest(...)` still called
  `GetEnvironmentVariableW(...)` directly instead of
  `TryReadEnvironmentVariableForSelfTest(kRacePathEnv)`.
- GREEN Pester after routing the race hook through the shared self-test env
  helper passed 41 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_172008_860.log`.
- GREEN focused FileOps archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_172218/` passed
  `Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial` with
  3 passed / 0 failed / 0 skipped.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools\Tests\TestInventory.Tests.ps1` were updated for the added Pester
  source-contract case. GR-S4 remains open for subitem (f).

## GR-S4(d) FolderView owned threadpool submit helper ratchet note (2026-07-05)

Scope: the async paste-shortcut worker in `FolderView.FileOps.cpp`, the
provider-allowed thumbnail worker in `FolderView.Icons.cpp`, the shared helper
in `Common\Helpers.h`, source-contract coverage, focused Commands runtime
coverage, and test-inventory documentation.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` reported 43 passed / 1 failed / 0 skipped because
  `SubmitOwnedThreadpoolCallback(...)` did not exist and both workers still
  owned local `TrySubmitThreadpoolCallback` / `work.release()` submit
  scaffolding.
- GREEN Pester after extracting `SubmitOwnedThreadpoolCallback(...)`, giving
  both work payloads `Execute() noexcept`, and correcting the provider block
  source-contract anchor passed 44 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_174449_942.log`.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_174925/` passed
  `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete` with
  1 passed / 0 failed / 0 skipped.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_174931/` passed
  `folderView_perf_slow_virtual_provider` with
  1 passed / 0 failed / 0 skipped.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools\Tests\TestInventory.Tests.ps1` were updated for the added Pester
  source-contract case. GR-S4 remains open for subitem (f).

## GR-S4(e) ViewerSpace env-flag helper ratchet note (2026-07-05)

Scope: the ViewerSpace opt-in env flag readers in
`Commands.SelfTest.ViewCommands.cpp`, source-contract coverage, focused Commands
skip-path runtime coverage, and test-inventory documentation.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru` reported 41 passed / 1 failed / 0 skipped because
  `IsViewerSpaceEnvFlagEnabled(...)` did not exist and the two opt-in helpers
  still carried local `GetEnvironmentVariableW(...)` readers.
- GREEN Pester after extracting the shared parser and routing
  `IsViewerSpaceLargePerfEnabled()` / `IsViewerSpace20kLayoutEnabled()` through
  it passed 42 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_172622_129.log`.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_172847/` exercised
  `cmd_viewer_space_layout_20k_visible_optin` without the opt-in env var and
  passed with 0 passed / 0 failed / 1 skipped.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_172852/` exercised
  `viewer_space_perf_large_optin` without the opt-in env var and passed with
  0 passed / 0 failed / 1 skipped.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools\Tests\TestInventory.Tests.ps1` were updated for the added Pester
  source-contract case. GR-S4 remains open for subitem (f).

## GR-S4(f) FolderView WarpDrive maintenance cleanup ratchet note (2026-07-05)

Scope: FolderView shell-thumbnail stat increment/perf emission duplication in
`FolderView.Icons.cpp`, pending input/refresh-to-paint metric shape duplication
in `FolderView.h` and `FolderView.cpp`, the already-shared D3D11 creation path
from GR-A6, source-contract coverage, focused runtime/perf coverage, and
durable FolderView documentation.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` reported 44 passed / 1 failed / 0 skipped because
  `IncrementThumbnailStat(...)` did not exist and shell thumbnail stat updates
  still paired direct `fetch_add(...)` calls with duplicate `PerfEmitCounter(...)`
  calls.
- GREEN Pester after adding `IncrementThumbnailStat(...)` and routing shell
  provider/cache/success counters through it passed 45 / 0 / 0.
- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` then reported 45 passed / 1 failed / 0 skipped because
  `PendingToPaintMetric` did not exist and FolderView still carried separate
  `PendingInputToPaintMetric` / `PendingRefreshToPaintMetric` shapes.
- GREEN Pester after replacing the two shapes with one `PendingToPaintMetric`
  while keeping separate pending input and refresh slots passed 46 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_175607_052.log`.
- GREEN focused Commands archives
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_175832/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_175839/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_175845/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_175904/` passed
  `folderView_thumbnail_cached_only_no_close_stall`,
  `folderView_perf_slow_virtual_provider`,
  `folderView_perf_refresh_preservation`, and
  `folderView_perf_scroll_render_stress`, respectively, each at
  1 passed / 0 failed / 0 skipped.
- The D3D11 creation duplication portion was already closed by GR-A6:
  `FolderView.Rendering.cpp` now uses
  `RedSalamander::DxUi::CreateD3D11DeviceWithWarpFallback(...)` instead of
  spelling the local hardware/WARP `D3D11CreateDevice` sequence.
- `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_TestCoverage.md`,
  `Tests/README.md`, and `Tools\Tests\TestInventory.Tests.ps1` were updated for
  the durable telemetry/helper contracts and added Pester source-contract cases.
  GR-S4 is now fully closed.

## Track T — perf gating (with GR-8 above)

| ID | Sev | Item (anchor) | Fix direction |
|----|-----|---------------|---------------|
| GR-P3 | **MEDIUM (tooling) - DONE 2026-07-05** | *(Execute with GR-8 and Tailwind TW-18 — the runner-wiring gap: `--selftest-perf-budget=` is passed by no automated runner — as one perf-gate tooling cluster.)* **Perf budget gate is a silent no-op everywhere but one machine.** `Specs/Testing/FolderViewPerfBudgets.json5:3` hard-codes `machineHash: "4cb089111a23"`; `CheckFolderViewPerfBudgets` (`Commands.SelfTest.ViewCommands.cpp:17140-17146`) trace-and-`return true` on mismatch; single-machineHash format enforced (`:16995-16997`); also silent-true on empty path (`:17126-17129`). Any hardware/driver change on the one machine changes the hash and disarms the gate with only a trace line. | Done: `Run-AllTests.ps1` and `TestRunPlan.ps1` now forward optional `-PerfBudgetPath` / `-RequirePerfBudgets` only to native self-test entries; the native harness parses `--selftest-require-perf-budgets`, reads `FolderViewPerfBudgets.json5` as `machines[]`, emits a visible scaffold warning for unknown machines, and strict-fails missing path/current-machine/no-hard-build matches. Durable contract lives in `Specs/Testing/Testing_PerformanceValidation.md`. RED Pester proved the missing runner/source contracts; GREEN Pester, app build `.build\logs\msbuild-20260705_155942_502.log`, and runtime archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_160750/` prove the implemented gate path. |

## GR-P3 FolderView perf-budget tooling gate closeout note (2026-07-05)

Scope: FolderView perf-budget runner wiring, strict native selftest gate behavior,
multi-machine budget-file shape, Pester source/plan guards, runtime parser exercise,
and durable testing/performance documentation.

- RED `Invoke-Pester .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru` first
  reported 8 passed / 1 failed because `Get-RSTestRunPlan` had no
  `PerfBudgetPath` parameter.
- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  first reported 34 passed / 1 failed because the native harness did not expose
  `--selftest-require-perf-budgets`.
- GREEN Pester after implementation: `RunAllTestsPlan.Tests.ps1` passed 9 / 0 / 0
  and `TestHarnessSourceContracts.Tests.ps1` passed 35 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_155942_502.log`.
- GREEN runtime exercise
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_thumbnail_cached_only_no_close_stall -FailFast -TimeoutMultiplier 3 -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5`
  passed 1 / 0 / 0; archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_160750/`. The trace confirms
  the `machines[]` budget file was parsed and the current machine matched before
  the Debug run explicitly skipped the Release-only hard budget.
- `Specs/Testing/FolderViewPerfBudgets.json5`,
  `Specs/Testing/Testing_PerformanceValidation.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`,
  `Tools\Tests\TestInventory.Tests.ps1`, and `Specs/Plans/WIP/README.md` were
  updated for the multi-machine/strict-mode contract and the two new Pester cases.

## Track C — conventions (AGENTS.md, all rule-quoted and verified)

| ID | Anchor | Violation → fix |
|----|--------|-----------------|
| GR-C1 | **DONE 2026-07-05** | `MtpBufferedWriter` and the adjacent `MtpBackendReader` now keep their owning `FileSystemMtp` alive through `wil::com_ptr<FileSystemMtp>` instead of owner-local manual `AddRef`/`Release`. |
| GR-C2 | **DONE 2026-07-05** | `FileSystemMtp::ReadDirectoryInfo(...)` now keeps `FilesInformationMtp` in `std::unique_ptr` ownership until successful `IFilesInformation**` handoff, with no manual `Release()` on failure. |
| GR-C3 | **DONE 2026-07-05** | `OpenSnapshotTempFile(...)` now keeps `std::bad_alloc` fatal and documents/logs the `noexcept` boundary before translating non-allocation `std::exception` temp-path construction failures to `E_FAIL`. |
| GR-C4 | **DONE 2026-07-05** | The two SearchAndIndex `EnumerateVolume(...)` noexcept callback lambdas now keep `std::bad_alloc` fatal and document/log before translating non-allocation `std::exception` vector-append failures to `E_FAIL`. |
| GR-C5 | **DONE 2026-07-05** | `TraceNativeTextInputTsfStep(...)` now uses `std::array<wchar_t, 768>` + `std::format_to_n(...)` bounded formatting for diagnostics, with no local `wchar_t line[768]` / `StringCchPrintfW` block. |
| GR-C6 | **DONE 2026-07-05** | The FileSystemMtp `#pragma warning(disable: ...)` sites now carry local rationale comments matching the sibling plugin WIL/yyjson suppression patterns. |

## GR-C1/GR-C2 MTP RAII ownership closeout note (2026-07-05)

Scope: public MTP helper COM object owner lifetime in
`FileSystemMtp.Core.cpp`, `FilesInformationMtp` construction/handoff in
`ReadDirectoryInfo(...)`, source-contract coverage, focused MTP runtime
coverage, and durable MTP spec text.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` reported 46 passed / 1 failed / 0 skipped because
  `MtpBufferedWriter` still stored an owning raw `FileSystemMtp*` and manually
  called `_owner->AddRef()` / `_owner->Release()`.
- GREEN Pester after converting `MtpBufferedWriter` and adjacent
  `MtpBackendReader` owners to `wil::com_ptr<FileSystemMtp>` and converting
  `ReadDirectoryInfo(...)` to `std::unique_ptr<FilesInformationMtp>` handoff
  passed 47 / 0 / 0.
- GREEN plugin build
  `.\build.ps1 -ProjectName FileSystemMtp -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_180344_864.log`.
- GREEN focused Compare archives
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_180425/` and
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_180431/` passed
  `mtp_fake_backend_enumerate_read_and_capabilities` and
  `mtp_public_writer_stages_until_commit`, respectively, each at
  1 passed / 0 failed / 0 skipped.
- `Specs/FileSystem/FileSystem_Mtp.md`, `Specs/Testing/Testing_TestCoverage.md`,
  `Tests/README.md`, and `Tools\Tests\TestInventory.Tests.ps1` were updated for
  the durable ownership contract and added source-contract case. GR-C3 is now
  closed; continue Track C with GR-C4.

## GR-C3 LocalSearch snapshot exception-handling closeout note (2026-07-05)

Scope: `OpenSnapshotTempFile(...)` in `Common/LocalSearchIndexCore.cpp`,
source-contract coverage, durable search spec text, and test inventory docs.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 47 / 1 / 0 against the unfixed catch block because
  `OpenSnapshotTempFile(...)` returned `E_FAIL` from
  `catch (const std::exception&)` without the mandatory comment and
  `Debug::Error(...)` log.
- GREEN Pester after adding the comment/log passed 48 / 0 / 0.
- GREEN app build
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_181218_425.log`.
- GREEN focused Compare archives
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_181423/` and
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_181429/` passed
  `search_source_allocation_and_folding_guard` and
  `search_low_hardening_smoke`, respectively, each at
  1 passed / 0 failed / 0 skipped.
- `Specs/Core/Core_Search.md`, `Specs/Testing/Testing_TestCoverage.md`,
  `Tests/README.md`, and `Tools\Tests\TestInventory.Tests.ps1` were updated for
  the snapshot `noexcept` exception contract and added source-contract case.
  Continue Track C with GR-C4.

## GR-C4 SearchAndIndex callback exception-handling closeout note (2026-07-05)

Scope: the two `EnumerateVolume(...)` callback lambdas in
`RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`,
source-contract coverage, focused Compare runtime coverage, and test inventory
docs.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 48 / 1 / 0 against the unfixed callback blocks
  because both `catch (const std::exception&)` branches returned `E_FAIL`
  without the mandatory comment and `Debug::Error(...)` log.
- GREEN Pester after adding the comments/logs passed 49 / 0 / 0.
- GREEN app build
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_181849_764.log`.
- GREEN focused Compare archives
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_182054/` and
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_182059/` passed
  `sqlite_index_store_load_and_apply_journal_delta` and
  `sqlite_index_store_root_lookup_case_insensitive`, respectively, each at
  1 passed / 0 failed / 0 skipped.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools\Tests\TestInventory.Tests.ps1` were updated for the added
  source-contract case. GR-C5 is now closed; continue Track C with GR-C6.

## GR-C5 DxUi NativeTextInput formatting closeout note (2026-07-05)

Scope: `TraceNativeTextInputTsfStep(...)` in
`Common/DxUi/DxUi.NativeTextInput.cpp`, source-contract coverage, focused
DxUi runtime coverage, and test inventory docs.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 49 / 1 / 0 against the old formatting block because
  it still used `wchar_t line[768]` and `StringCchPrintfW`.
- GREEN Pester after converting to `std::array<wchar_t, 768>` and
  `std::format_to_n(...)` passed 50 / 0 / 0.
- GREEN app build
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_182416_209.log`.
- GREEN `DxUiTests` build
  `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_182634_951.log`.
- GREEN focused native test run
  `.\.build\x64\Debug\DxUiTests.exe --suite=NativeTextInput
  --perf-jsonl=Specs\TestRuns\local_scratch\dxui_native_textinput_gr_c5_format_to_n_20260705_1828.jsonl`
  exited 0 with "All DxUi tests passed."
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools\Tests\TestInventory.Tests.ps1` were updated for the added
  source-contract case. Continue Track C with GR-C6.

## GR-C6 FileSystemMtp warning-suppression rationale closeout note (2026-07-05)

Scope: uncommented `#pragma warning(disable: ...)` suppressions in
`Plugins/FileSystemMtp/FileSystemMtp.h`, `Factory.cpp`,
`FileSystemMtp.Core.cpp`, `FileSystemMtp.Device.cpp`, and
`FileSystemMtp.FakeBackend.cpp`, source-contract coverage, and test inventory
docs.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 50 / 1 / 0 against the uncommented MTP suppression
  sites because the new source contract required local rationale comments.
- GREEN Pester after adding the WIL and yyjson rationale comments passed
  51 / 0 / 0.
- GREEN plugin build
  `.\build.ps1 -ProjectName FileSystemMtp -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_183102_114.log`.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools\Tests\TestInventory.Tests.ps1` were updated for the added
  source-contract case. Track C is now closed; the later final audit closed the
  stale GR-1/GR-T1 cleanup and routed the remaining GR-A5 ratchet.

## Decisions needing explicit sign-off (blocking merge to master)

- **GR-D1 = TW-D1 (thumbnails) — SIGNED OFF 2026-07-04.** Ship cached-only visible thumbnails as spec-mandated in `Specs/UI/UI_FolderView.md`. Background provider-allowed enrichment is a real follow-up, not deferred cleanup, and is tracked in `Specs/Plans/WIP/FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md`. TW-2/TW-3 cleanup routes through that follow-up; TW-4's pending-accounting sibling was fixed with GR-18 on 2026-07-05.
- **GR-D2 (ComboBox click-through light dismiss) — SIGNED OFF 2026-07-04.** Preserve standard Windows combobox outside-click behavior. `DxUi.WindowHost.cpp:2448` may dismiss the dropdown and continue hit-testing, including activation of an underlying destructive command; destructive surfaces must protect themselves via confirmation/enablement, not by changing ComboBox light-dismiss semantics. Requirement persisted in `Specs/UI/UI_DxUiWinUIDesign.md`.
- **GR-D3 plugin-manager retry (V35, PLAUSIBLE) — DECIDED 2026-07-04.** *(Renamed 2026-07-02 — was mislabeled "GR-16", colliding with Track U's GR-16 root-collapse item.)* Add automatic unload-deferred retry: sweep on `GetConfigurationSchema`/schema-load `ERROR_BUSY` or use a bounded retry timer. Apply/Refresh is not an acceptable sole recovery path. Requirement persisted in `Specs/Plugins/Plugins_PluginAPI.md` and `Specs/UI/UI_ManagePluginsDialog.md`.

## Closed plausible low items

- **GR-12 — DONE 2026-07-05.** `IconCache.cpp` path failure stores now use insert-only `emplace(...)` semantics and emit `iconcache.duplicate_path_query_race` when another thread wins the store race, so a slow failed shell query cannot downgrade a concurrent successful icon-index result to a 5s negative entry. This is independent from GR-23's live-lookup transient-failure fix, which remains done.
- **GR-17 — DONE 2026-07-05 via GR-6.** `DxUi.Menu.cpp:5294-5298` (ENABLE_TESTS only): `SendMessageTimeoutW(1000ms)` + stack `outState` used the same dangling shape as GR-6. Folded into GR-6 by reverting `DebugGetContextMenuPopupState` to synchronous `SendMessageW(...)` and adding `TestContextMenuDebugStateCrossThreadQueryDoesNotUseTimedOutStackStorage`.

## GR-12 IconCache path failure-store race closeout note (2026-07-05)

Scope: `IconCache::QuerySysIconIndexForPath(...)` path failure-store behavior in
`RedSalamander/IconCache.cpp`, source-contract coverage, durable FolderView spec
text, and test inventory docs.

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 51 / 1 / 0 against the old failure-store block
  because it still used `_pathToIconIndex.insert_or_assign(...)`.
- GREEN Pester after changing the failure store to insert-only
  `_pathToIconIndex.emplace(...)` passed 52 / 0 / 0.
- GREEN app build
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_183746_054.log`.
- `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_TestCoverage.md`,
  `Tests/README.md`, and `Tools\Tests\TestInventory.Tests.ps1` were updated
  for the IconCache failure-store duplicate-race contract and added
  source-contract case.

## Tests to add (each must go RED on the un-fixed code)

| ID | For | Required proof |
|----|-----|----------------|
| GR-T1 | GR-1 | **STALE / NOT ADDED 2026-07-05.** Retired with GR-1 after static recheck confirmed the current tree already has the one-shot `_pathEditBlurSuppressActive = false` resets at `NavigationView.Edit.cpp:765` and `NavigationView.Interaction.cpp:808`; the historical RED premise no longer reproduces from the cited anchors. |
| GR-T2 | GR-2 | **DONE 2026-07-04.** Added `cmd_pane_clipboardPasteShortcut_concurrent_invocations_create_distinct_links`: two same-stem Paste Shortcut commands are widened at `PasteShortcutAfterSlotProbe`; RED produced one link, GREEN produces `alpha - Shortcut.lnk` and `alpha - Shortcut (2).lnk`. |
| GR-T3 | GR-3 | **DONE 2026-07-04.** Extended `cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker` to revisit stale folder A through cache and see the created link, and added `cmd_pane_clipboardPasteShortcut_failure_after_navigate_shows_alert` with a forced `CreateShellShortcut` HRESULT after navigating away. Both subcases went RED then GREEN. |
| GR-T4a | GR-4a | **DONE 2026-07-04.** Added `mtp_rename_overwrite_uses_temp_puid_swap`: RED showed rename-overwrite kept the source PUID (`source=puid-14 final=puid-14`), proving it used fake backend overwrite directly; GREEN requires final PUID to differ from both old destination and source plus fake-backend `"copyItemCalls":1` trace evidence for the temp-copy path. |
| GR-T4b | GR-4b | **DONE 2026-07-04.** Added `mtp_fake_backend_move_rejects_directory_transfer_fallback`: RED showed a non-leaf-preserving cross-parent directory move returned `S_OK`; GREEN requires `ERROR_NOT_SUPPORTED`, source directory/child preservation, and destination absence. Full `mtp_` filter remained green. |
| GR-T5 | GR-5 | **DONE 2026-07-04.** `DxUiTests.TextField` source-contract coverage now requires `TextField::InvalidateSingleLineLayoutCache()` to call `ClearSingleLineTextLayoutCache(_singleLineLayoutCache, true)` and requires the TextField single-line accessor to pass secure clearing through shared helper failure paths. Both expectations went RED before the implementation and GREEN afterward. |
| GR-T6 | GR-6 | **DONE 2026-07-05.** Added `TestAccessibilityUiActionDispatchOwnsTimedOutRequestStorage` as a source-contract RED/GREEN guard for heap-owned dispatch storage, provider COM keepalive, `PostMessagePayload(...)`/`TakeMessagePayload<T>(...)`, event timeout handling, no stack request sends, and no caller-owned rectangle output pointer. Added `TestAccessibilityTextRangeBoundingRectanglesTimeoutKeepsLateHandlerStorageAlive`, which stalls a dequeued UIA action handler until the `GetBoundingRectangles` sender returns `ERROR_TIMEOUT`, then releases the handler and proves the caller `SAFEARRAY*` output stayed untouched. Added folded GR-17 guard `TestContextMenuDebugStateCrossThreadQueryDoesNotUseTimedOutStackStorage`; RED failed while the menu debug path used `SendMessageTimeoutW(...)` with stack output, GREEN passes after `SendMessageW(...)`. |
| GR-T7 | GR-7 | **DONE 2026-07-05.** Added `mtp_overwrite_journal_clears_completed_swap_without_temp`: RED retained the stale completed-swap journal; GREEN clears the final-present/temp-missing no-tempPUID journal on declared-size match and proves the following backend command does not repeat the stale replay sweep work. |
| GR-T8 | GR-8 | **DONE 2026-07-04.** Added `Tools\Tests\ShowPerfRuns.Tests.ps1` coverage: one-shot gauge plus sufficient p95-budgeted distribution exits 0; p95-budgeted `folder.frame.input_to_paint_us` uses its 40-sample budget; unbudgeted preset metrics stay informational; low-sample `folder.frame.total_us` exits 2; compare mode ignores sparse baseline quality and gates the candidate. |
| GR-T9 | GR-9 | **DONE 2026-07-05.** Added `TestAccessibilityProviderExposesHorizontallyScrolledGridRowStructure`: RED failed because the scrolled grid row UIA `Name` omitted the horizontally off-view `Alpha` cell; GREEN passes with row name `Alpha | Ready | Archived`, first row child `GridCell` `Alpha`, `UIA_IsOffscreenPropertyId=true`, and next structural cell `Ready`. |
| GR-T10 | GR-10 | **DONE 2026-07-05.** Added `TestAccessibilityProviderPointHitsClipAndTranslateScrollPanelChildren`: RED failed because the raw content-space point for the scrolled-out button returned a child provider instead of the root provider (`hr=0x80004002` when queried as root); GREEN passes with the empty viewport point resolving the root provider and the visible button resolving at its viewport-translated point. |
| GR-T11 | GR-11 | **DONE 2026-07-05.** Added `search_service_rebuild_deleted_root_purges_index`: RED failed at `RequestRebuild(root)` with `hr=0x80070002`; GREEN succeeds after deleted-root rebuild authorization checks the deepest existing ancestor and purges stale indexed candidates before the recreated root is queried. |
| GR-T13 | GR-13 | **DONE 2026-07-05.** Added FileSystem debug selftest coverage for `IsServiceFallbackCandidate` returning true for `ERROR_BAD_PATHNAME`, `ERROR_PATH_NOT_FOUND`, and `E_INVALIDARG`; RED PluginContract failed three selftest assertions, GREEN PluginContract passed. |
| GR-T14 | GR-14 | **DONE 2026-07-05.** Added `search_service_candidate_impersonation_failure_is_incomplete_warning`: RED failed the whole query with `hr=0x80070558`; GREEN completes with the unaffected candidate and `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`. |
| GR-T15 | GR-S3 | **DONE 2026-07-05.** Added `search_service_sqlite_external_rotation_refreshes_without_retry`: RED archive `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113058/` failed because the post-rotation query used `search.backend.sqlite.retry_query_ms`; GREEN archive `2026-07-05_113650/` returns the externally rotated SQLite contents with `QueryExecutionMode::Sqlite`, `FallbackReason::None`, and no post-rotation retry metric after the selftest's perf-file offset. |
| GR-T16 | GR-16 | **DONE 2026-07-05.** Added `TestAccessibilityLabelOnlyRootDoesNotUseDirectSemanticRootCollapse`: RED failed because the snapshot path collapsed a Label-only root and `Navigate(FirstChild)` returned null; GREEN passes after snapshot collapse eligibility uses the retained-path single-semantic-root predicate, with the label exposed as a `UIA_TextControlTypeId` child and no duplicate nested label. |
| GR-T17 | GR-18 | **DONE 2026-07-05.** Added `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count`: RED first failed stale batch/generation payload handling with `pending=0 expected=2 staleDrops=1`; after the stale-return ordering fix, the TW-4 extension failed late current unaccounted delivery with `pending=1 expected=2 completed=1`; GREEN passes after `countsPending=false` routes abandoned late provider delivery off the pending-decrement path while normal posts remain counted. |
| GR-T19 | CW-7 addendum | **DONE 2026-07-04.** Merge audit completed; see audit note below. |

## GR-2/GR-3 Paste Shortcut closeout note (2026-07-04)

Scope: Paste Shortcut async worker correctness in `FolderView.FileOps.cpp`, collision-safe shell-link saving in `FolderViewInternal.h`, path-visit generation in `FolderView.cpp` / `FolderView.h`, and focused command selftests.

Commands/evidence:
- RED `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_clipboardPasteShortcut_concurrent_invocations_create_distinct_links -FailFast` failed with only `alpha - Shortcut.lnk`.
- RED `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker -FailFast` failed because revisiting stale folder A through cache did not show the new shortcut.
- RED `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_clipboardPasteShortcut_failure_after_navigate_shows_alert -FailFast` failed before the forced failure hook was consumed because the stale destination still got a shortcut.
- GREEN `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors after the Paste Shortcut changes.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_clipboardPasteShortcut_failure_after_navigate_shows_alert -FailFast` passed after the worker consumed the forced failure hook.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_clipboardPasteShortcut_ -FailFast` passed all 6 Paste Shortcut cases.

Implementation notes:
- `PasteShortcutFromClipboard` captures clipboard paths on the UI thread, serializes same-FolderView requests with `_pasteShortcutInFlight`, and queues later requests with the captured payload.
- `CreateShellShortcut` retries the next deterministic shortcut slot when a race reports `ERROR_FILE_EXISTS` / `ERROR_ALREADY_EXISTS`; `SaveShellShortcutExactPath` no longer uses short-path replace semantics.
- `OnPasteShortcutComplete` always emits perf and reports failures; created-link cache invalidation runs for the result target folder regardless of the current pane path. Refresh/focus is guarded by current target path plus `_folderPathGeneration`, so ordinary same-folder enumeration refreshes do not stale the completion.

## GR-4 MTP backend parity closeout note (2026-07-04)

Scope: public MTP rename-overwrite routing in `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp`, fake backend direct-move parity in `Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp`, MTP compare selftests, and the durable MTP filesystem contract.

Commands/evidence:
- RED `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_rename_overwrite_uses_temp_puid_swap -FailFast -TimeoutMultiplier 3` failed with `destination kept the source PUID instead of taking a temp PUID. source=puid-14 final=puid-14`.
- RED `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_fake_backend_move_rejects_directory_transfer_fallback -FailFast -TimeoutMultiplier 3` failed with `non-leaf-preserving directory move should be unsupported, got hr=0x00000000`.
- GREEN `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors after the MTP changes.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_rename_overwrite_uses_temp_puid_swap -FailFast -TimeoutMultiplier 3` passed.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_fake_backend_move_rejects_directory_transfer_fallback -FailFast -TimeoutMultiplier 3` passed.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3` passed 46 cases, failed 0, skipped `mtp_live_device_smoke` because no approved live device was configured.

Implementation notes:
- `FileSystemMtp::RenameItem` now detects existing destinations for `FILESYSTEM_FLAG_ALLOW_OVERWRITE`, blocks directory destinations and created-object-PUID-unsupported devices consistently with copy/move, and commits replacement through `CommitDeviceSourceOverwriteWithTempSwap(..., moveSource=true, ...)`.
- Missing-destination rename now calls `backend.RenameItem(..., false)`, so public rename no longer depends on fake-only backend overwrite behavior.
- `FakeMtpBackend::MoveItem` now mirrors WPD direct move: destination must not exist; non-leaf-preserving directory moves return `ERROR_NOT_SUPPORTED`; leaf-preserving directory/file moves use the native rename path; non-leaf-preserving file moves use transfer-copy plus source delete, preserving the delete-source failure injection for file fallback tests.
- `Specs/FileSystem/FileSystem_Mtp.md` now records both the rename-overwrite temp-swap contract and the direct backend move parity contract.

## GR-5 password layout-cache closeout note (2026-07-04)

Scope: DxUi single-line TextField layout-cache clearing in `Common/DxUi/DxUi.TextInput.cpp`, shared single-line helper failure-path cache clearing in `Common/DxUi/DxUi.SingleLineTextEditing.cpp`, and focused `DxUiTests.TextField` source-contract coverage.

Commands/evidence:
- RED `.\.build\x64\Debug\DxUiTests.exe --suite=TextField` failed with `single-line layout invalidation securely clears cached text retained for DirectWrite layout reuse` before the invalidator passed `secureText=true`.
- RED `.\.build\x64\Debug\DxUiTests.exe --suite=TextField` then failed with `TextField single-line accessor asks the shared helper to securely clear cached text on layout failure` before the shared helper accepted and consumed the secure-cache flag.
- GREEN `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors after the final header/source signature alignment.
- GREEN `.\.build\x64\Debug\DxUiTests.exe --suite=TextField` passed after the TextField invalidation and helper failure-path fixes.

Implementation notes:
- `TextField::InvalidateSingleLineLayoutCache()` now calls `ClearSingleLineTextLayoutCache(_singleLineLayoutCache, true)` because reveal-visible password states can leave plaintext in the retained single-line layout cache.
- `GetOrCreateSingleLineTextLayout(...)` now carries a defaulted `secureCacheText` flag and uses it for helper-owned cache clears on layout creation, resize, and geometry failure paths.
- TextField passes `secureCacheText=true`; editable ComboBox and other non-secret shared-helper callers keep the default false path.
- `Specs/UI/UI_DxUiWinUIDesign.md` now records that masking is not full secure memory while the field is live, but TextField-owned transient/native/layout caches that can retain revealed plaintext must secure-clear their cached strings on remask, invalidation, teardown, and destruction.

## GR-8 FolderView perf-gate closeout note (2026-07-04)

Scope: `Tools/Show-PerfRuns.ps1` sample-quality gating, focused Pester coverage in `Tools/Tests/ShowPerfRuns.Tests.ps1`, tool-test inventory counts, and `Specs/Testing/Testing_PerformanceValidation.md`.

Commands/evidence:
- RED `powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\ShowPerfRuns.Tests.ps1 -EnableExit"` initially failed three GR-T8 paths: one-shot gauge count 1 latched `-FailOnQuality`, `folder.frame.input_to_paint_us` ignored its 40-sample p95 budget, and compare mode latched the sparse baseline.
- RED after the first implementation pass added the unbudgeted-preset case and failed because `icons.extract_us` count 1 still latched `-FolderViewPreset -FailOnQuality`.
- GREEN `powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\ShowPerfRuns.Tests.ps1 -EnableExit"` passed 5/5 after the final preset quality classification.
- GREEN real gate evidence is archived at `Specs/TestRuns/4cb089111a23/Tooling/2026-07-04_213000_show_perfruns_gr8/`: `pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tools\Show-PerfRuns.ps1 -Run "Z:\src\RedSalamander\Specs\TestRuns\4cb089111a23\Commands\2026-06-28_174423" -FolderViewPreset -FailOnQuality` exited 0.

Implementation notes:
- `Show-PerfRuns.ps1` now loads p95 `minimumSamples` from `Specs/Testing/FolderViewPerfBudgets.json5`; only p95 budget rows participate in FolderViewPreset quality latching.
- FolderViewPreset metrics without p95 budget rows are still summarized but report `p95q=n/a` and do not fail `-FailOnQuality` unless the user requests the metric explicitly with `-Metric`.
- One-shot gauge/event metrics are not treated as p95 distribution gates by the default analyzer.
- `Compare-PerfRuns` calls `Measure-Metric` with quality registration disabled for the baseline and enabled for the candidate, so sparse historical baselines remain comparable without failing the command.
- Tool-test inventory/docs were updated for 120 Pester-style cases, 678 Commands `RunCase` registrations, and 232 CompareDirectories `RunCase` registrations.

## GR-P1 MTP worker-queue closeout note (2026-07-04)

Scope: `RunBackendCommand` thread churn in `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp`, fake-backend backend-thread instrumentation, MTP compare selftests, and the durable MTP filesystem/test coverage contracts.

Commands/evidence:
- RED `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_backend_command_worker_is_reused -FailFast -TimeoutMultiplier 3` failed with `MTP backend worker reuse: expected one long-lived backend worker thread, observed 5.`
- GREEN `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors after the worker-queue changes.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_backend_command_worker_is_reused -FailFast -TimeoutMultiplier 3` passed.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3` passed 47 cases, failed 0, skipped `mtp_live_device_smoke` because no approved live device was configured.
- GREEN runner inventory listed `mtp_backend_command_worker_is_reused`; source inventory now reports 233 CompareDirectories `RunCase` registrations.
- Perf/test evidence is archived at `Specs/TestRuns/4cb089111a23/Mtp/2026-07-04_214700_gr_p1_worker_queue/`.

Implementation notes:
- `FileSystemMtp` now owns a long-lived backend command queue worker. Concurrent host calls enqueue backend work and wait on per-command completion instead of creating one `std::jthread` per call.
- The queue still executes one command at a time under the existing device I/O mutex and replays retained overwrite journals before command execution.
- On watchdog timeout or Disconnect cleanup, the active worker is detached into module-level quarantine, pending queued commands complete with `ERROR_DEVICE_NOT_CONNECTED`, backend cancellation is requested, and unload remains deferred until the worker exits.
- GR-P2 streaming reader is closed below; the next Granite-owned MTP item is GR-7 stale overwrite journal replay.

## GR-P1 MTP WPD session/path cache closeout note (2026-07-04)

Scope: `WpdMtpBackend` session reuse, normalized path-to-objectId cache, cache invalidation, debug-only deterministic WPD-cache selftest backend, MTP compare selftests, and durable MTP/test coverage specs.

Commands/evidence:
- RED `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_wpd_session_and_path_cache_reuse -FailFast -TimeoutMultiplier 3` failed with `MTP WPD cache: create selftest instance failed. hr=0x8007007F` before the debug WPD-cache backend export existed.
- GREEN `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors after the final cache/failure-invalidation changes.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_wpd_session_and_path_cache_reuse -FailFast -TimeoutMultiplier 3` passed.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3` passed 48 cases, failed 0, skipped `mtp_live_device_smoke` because no approved live device was configured.
- GREEN runner inventory listed 242 CompareDirectories cases; source inventory now reports 234 CompareDirectories `RunCase` registrations.
- Perf/test evidence is archived at `Specs/TestRuns/4cb089111a23/Mtp/2026-07-04_221500_gr_p1_wpd_session_path_cache/`.

Implementation notes:
- `WpdMtpBackend` caches `IPortableDevice`/`IPortableDeviceContent` by PnP id and reuses a cached session when its access mask covers the requested read/write access.
- The backend memoizes normalized path-to-objectId resolution; repeated metadata calls for the same deep path hit the cache instead of re-enumerating each ancestor.
- Successful writes invalidate the affected cached path subtree, and unexpected WPD session/path-resolution failures clear device descriptors, sessions, and path-cache entries.
- The debug-only `mtp_wpd_session_and_path_cache_reuse` guard proves one device enumeration, one session open, four child-resolution enumerations for the first ancestor walk, and repeated path-cache hits.

## GR-P2 MTP streaming reader closeout note (2026-07-05)

Scope: MTP `CreateFileReader` no longer materializes a full file at open. Reader open and `GetSize` stay metadata/stream setup operations; bytes move on `Read`. Streamed reader operations still run through `RunBackendCommand`, so watchdog timeout, backend cancellation, worker quarantine, and plugin unload deferral continue to apply to reads.

Commands/evidence:
- RED `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_reader_streams_on_read_not_open -FailFast -TimeoutMultiplier 3` failed before the streaming reader implementation with `opening the reader materialized the file; readFileCalls=1`.
- RED `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_hung_device_times_out -FailFast -TimeoutMultiplier 3` failed after the watchdog test was updated for streamed reads with `delayed Read expected device-gone/0 bytes, got hr=0x00000000 bytes=1`.
- GREEN `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors (`.build\logs\msbuild-20260705_085022_370.log`).
- GREEN focused cases passed: `mtp_reader_streams_on_read_not_open`, `mtp_reader_seek_contract`, `mtp_hung_device_times_out`, `mtp_watchdog_requests_backend_cancel`, and `mtp_runtime_refresh_defers_when_worker_quarantined`.
- GREEN `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3` passed 49 cases, failed 0, skipped `mtp_live_device_smoke` because no approved live device was configured.
- GREEN `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5, failed 0.
- Perf/test evidence is archived at `Specs/TestRuns/4cb089111a23/Mtp/2026-07-05_090500_gr_p2_streaming_reader/`.

Implementation notes:
- `IMtpBackendFileReader` is the backend-owned streaming reader seam; `FileSystemMtp::CreateFileReader` now returns an `MtpBackendReader` adapter instead of reading the whole file into a core memory reader.
- WPD-backed readers keep `IPortableDeviceContent`, `IStream`, size metadata, and shared cancellation state alive. `Read` and `Seek` enter `ScopedActiveWpdContent`/`ScopedActiveWpdStream` and emit `mtp.transfer.read_bytes` per successful read chunk.
- The adapter copies backend reads through a scratch vector before copying into the caller buffer, so a timed-out backend read cannot write through an expired caller stack pointer.
- The fake backend now distinguishes open/size from actual reads: `readFileCalls` remains zero through `CreateFileReader` and `GetSize`, increments on `Read`, and `readFileDelayMs` delays streamed `Read` for watchdog/quarantine coverage.
- `Specs/FileSystem/FileSystem_Mtp.md`, `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and `Tools\Tests\TestInventory.Tests.ps1` were updated for the streaming reader contract and the new Compare/MTP test count.

## GR-T19 merge audit note (2026-07-04)

Scope: merge `d4a064dce` (`744aae30c` + `3c0ad3d75`), with merge-base `a72512919`.

Commands/evidence:
- `git show --cc --unified=0 --format= d4a064dce` listed true combined-resolution hunks only in `Common/DxUi/DxUi.Internal.h`, `Common/DxUi/DxUi.TextStoreACP.cpp` (3 hunks), `Common/DxUi/DxUi.WindowHost.cpp`, and `Tests/DxUiTests/DxUiTests.NativeTextInput.cpp`.
- `git range-diff a72512919 d4a064dce^1 d4a064dce` showed all first-parent commits through `744aae30c` retained, with only `3c0ad3d75 fix(dxui): complete Bedrock remediation` added on top.
- `git diff 744aae30c d4a064dce -- Common/DxUi/DxUi.Controls.cpp` identified the actionable first-parent revert at `TabControl::OnMouseUp`: the merge reintroduced the loose close-button release clause that `ec403afac` had removed.

Manual dispositions:
- `DxUi.Internal.h`: combined declaration merge only; retained both native text-store teardown entry points and backdrop capture declaration. No action.
- `DxUi.TextStoreACP.cpp`: reconciled `DetachHost` with `Disconnect`, preserving host/control pointer severing, lock/sink reset, and secure wipe of observed text. No action.
- `DxUi.WindowHost.cpp`: reconciled key-up native IME handling with reentrant key-target validation, and preserved native text-input shutdown during process-exit host shutdown. No action.
- `DxUiTests.NativeTextInput.cpp`: reconciled teardown source-contract tests with the newer split between clearing TSF focus and thread-manager shutdown. No action.
- `DxUi.Controls.cpp` / `TabControl::OnMouseUp`: actionable 744-side revert. Fixed by restoring close-only-if-pressed-on-same-close-button semantics and adding CW-T7 regression coverage.

## Granite addendum 2026-07-04 — folderview-warpdrive merge re-review (base `45ae2a9a` → merge `6e6ff863b`)

Max-effort local re-review of the **merged** `codex/folderview-warpdrive` (merge `6e6ff863b`), run
after Granite's 2026-07-02 pass to catch what the merge introduced or the first pass missed.
Method: 10 finder angles → 8 verifier agents → an adversarial reconciliation critic run against
*this* ledger. The range overlaps Granite's, so most findings were **already tracked** (below);
seven are genuine deltas. Every new item is **CONFIRMED** at `6e6ff863b` with verifier-exact anchors
(the tree is the merged branch, so these anchors are newer than the 275c04034 anchors elsewhere in
this plan — re-check drift before executing).

**Already covered — execute the named item, do NOT duplicate:**

| Re-review finding | Owned by |
|-------------------|----------|
| Async paste-shortcut concurrent `.lnk` overwrite | **GR-2** |
| Paste-shortcut `result.generation` staleness guard inert | **GR-3** (already absorbs TW-7) |
| Provider-thumbnail machinery dead in Release + `ExtractProviderAllowedThumbnailWithDeadline` deadline race — incl. the abandoned-worker **double-post** of `kFolderViewCreateThumbnailBitmap` (late real bitmap + main-loop fallback → one `pendingBitmapCreates` increment, two decrements) and the opposite-order **dropped bitmap** (valid shell thumbnail discarded to fallback) | **TW-2 / TW-3** remain routed; the double-post pending-counter part was fixed with **GR-18 / TW-4** on 2026-07-05, confined to the `ForceProviderAllowedProbe` ENABLE_TESTS path |
| UIA snapshot full-tree rebuild per keystroke/selection with no `UiaClientsAreListening` gate | **GR-A1** (routed to the UIA baton) — site count updated in the routing table above |
| `IsTruthySelfTestEnvironmentVariable` / owned-threadpool-submit scaffold duplication | **GR-S4(b) / (d)** |

**New / escalated items** (Track placement noted; anchors at `6e6ff863b`):

| ID | Sev | Item (anchor) | Fix direction | Proof |
|----|-----|---------------|---------------|-------|
| GR-20 | **MED — DONE 2026-07-05** | **Hot-reload startup blocks the UI thread up to 2000ms and silently disables itself on a slow/locked settings directory.** `SettingsHotReload::Start()` spawned the watcher then synchronously blocked the caller on `WaitForSingleObject(readyHandle, 2000)`, while the watcher only signaled readiness after a successful `FindFirstChangeNotificationW(...)`; transient arm failures therefore stalled window creation, returned `HRESULT_FROM_WIN32(WAIT_TIMEOUT)`, and stopped the watcher instead of letting it self-heal. | Done: `SettingsHotReload::Start(...)` now launches the watcher and returns `S_OK` without waiting on `readyEvent`; transient `FindFirstChangeNotificationW(...)` failures remain in the worker retry loop instead of causing `Stop()` + `WAIT_TIMEOUT`. An `ENABLE_TESTS` hook and `settings_hot_reload_transient_arm_failure_is_async` cover one forced arm failure, the sub-200ms start budget, no hard-fail/teardown, and a later settings-change message post after self-arm. Durable contract lives in `Specs/Core/Core_SettingsStore.md`. RED compile log `.build\logs\msbuild-20260705_141406_457.log`; RED behavioral archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_141831/`; GREEN build `.build\logs\msbuild-20260705_142433_667.log`; GREEN focused archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_142640/`; GREEN adjacent archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_142715/`. | GR-T20 DONE |
| GR-21 | **MED — DONE 2026-07-05** | **Compare "Invert selection" is sticky and re-applies on every pane refresh/navigation.** `IDM_COMPARE_INVERT_DIFFERENCES_SELECTION` persisted `_selectionMode = Inverted`, so later compare-window pane refresh/navigation could carry or re-apply inverted selection instead of returning to the default decision-model selection. | Done: Invert now applies the inverted decision predicate only to the current panes and immediately resets `_selectionMode` to `Default`; compare-window left/right refresh commands also re-apply the default compare selection before forcing a pane refresh so `FolderView` refresh-preservation cannot carry an inverted selection forward. Durable contract lives in `Specs/Core/Core_CompareDirectories.md`. Extended `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state` to cover default selection, invert, and left-refresh default restoration. RED archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_143810/` failed with `leftSelected=0`; GREEN build `.build\logs\msbuild-20260705_144528_275.log` passed with 0 warnings / 0 errors; GREEN focused archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_144738/`; GREEN adjacent archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_145147/` passed 15 / 0 / 0. | GR-T21 DONE |
| GR-22 | **MED - DONE 2026-07-05** | **PendingRefreshToPaint telemetry leaks past its paint, corrupting the refresh-latency metric.** `UpdatePendingRefreshToPaintResult` sets `resultReady=true` (`FolderView.Enumeration.cpp:1695`), but `EmitPendingRefreshToPaintMetricAfterPresent` (`FolderView.cpp:164`) is only called from the two successful-Present branches (`FolderView.Rendering.cpp:2025/:2064`). Every `Render` early-return bypasses it without clearing the metric: no render target (`:1018`), EndDraw/device-loss (`:1971`), Present1 fail (`:2010`), legacy Present fail (`:2049`). The stranded metric is then either overwritten by the next `RecordPendingRefreshToPaintStart` (`:126`, lost sample) or emitted by an *unrelated* later Present on the same folder - the emit gate (`:166`) checks only `has_value() && resultReady`, never re-validating generation - reporting `ElapsedUs(startedAt)` to that much-later paint and grossly inflating `folder.refresh.request_to_paint_us`, the headline metric this branch exists to measure. `CancelPendingRefreshToPaint` (`:146`) is generation-gated and can't clean a leaked prior-generation metric. | Done: `RecordPendingRefreshToPaintStart(...)` resets before recording a new slot; successful emit revalidates the current enumeration generation; ready pending refresh-to-paint metrics are discarded on no-present/failure render paths (`OnPaint` fallback, no-target render exit, `EndDraw`, `Present1`, and legacy `Present` failure). Durable contract lives in `Specs/UI/UI_FolderView.md`, and `folderView_refresh_to_paint_metric_clears_after_failed_render` covers the failed-paint plus later-unrelated-Present stale-metric leak. RED archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150226/`; GREEN build `.build\logs\msbuild-20260705_150338_519.log`; GREEN focused archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150541/`; adjacent archives `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150654/`, `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150739/`, `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151107/`, and `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151143/`. | GR-T22 DONE |
| GR-23 | **MED - DONE 2026-07-05** | **Transient live icon-lookup failures cached 5s with no F5 recovery** *(escalates GR-12 — distinct defect, independent fix).* `QuerySysIconIndexForPath` now caches every `SHGetFileInfoW` failure for `kPathIconFailureTtl`=5s (`IconCache.cpp:40`) keyed by {path,attrs,useFileAttributes}, including live lookups (`useFileAttributes==false`; store `:1235`, branch `:1226`); commit `104f4b5fa` deleted the guarding comment *"Don't cache failures — they may be transient (network drives, shell extensions loading)."* The read path short-circuits a cached failure to `nullopt` (`failureCacheHit`, `:1169-1194`) before re-calling the shell. No path-scoped invalidation exists: `Clear()` (`:1415`) is a global settings/theme reset (`RedSalamander.cpp:8515`); ForceRefresh/F5 (`FolderView.cpp:331`) re-runs the per-file icon worker (`FolderView.Enumeration.cpp:848`) but it hits the cached `nullopt`. So a transient failure on a network/offline drive or still-loading shell extension shows a generic icon that stays wrong for the full 5s across repeated F5. | Done: live path lookup failures (`useFileAttributes=false`) now return `nullopt` without storing a negative cache entry and emit `iconcache.path_live_lookup_failed_uncached`; the same path can immediately re-enter `SHGetFileInfoW` and recover, then positive-cache the successful index. Attribute-mode `SHGFI_USEFILEATTRIBUTES` failures retain the bounded negative cache. Durable contract lives in `Specs/UI/UI_FolderView.md`; focused and adjacent Commands coverage is `folderView_iconcache_live_path_failure_retries_without_negative_cache`, `folderView_perf_slow_virtual_provider`, and `folderView_perf_icon_pipeline_cold_slow`. RED archives `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151809/` and `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_152050/`; GREEN build `.build\logs\msbuild-20260705_152904_817.log`; GREEN focused archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153136/`; adjacent archives `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153107/` and `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153209/`. | GR-T23 DONE |
| GR-24 | **LOW - DONE 2026-07-05** | **Dead regression guard: the per-item transient-brush counter has no producer** *(restates Tailwind TW-1 with fresh anchors; ENABLE_TESTS-only, no Release impact).* `_debugDrawItemTransientBrushCreateCount` (`FolderView.h:1479`) is surfaced in `RenderingDebugSnapshot` (`FolderView.cpp:1108`), zeroed by the reset (`:1159`), and asserted `== 0u` by two self-tests (`Commands.SelfTest.ViewCommands.cpp:20319` and `:23303`) — but grep confirms the **only** write anywhere is the `.store(0)` reset; there is no `fetch_add`/`++` producer. `DrawItem` (`FolderView.Rendering.cpp:2245`) now uses cached member brushes, so the guard was dead. | Done: `DrawItem` still uses cached member brushes for normal selected/hovered rendering, while an `ENABLE_TESTS` forced transient-brush path creates a RAII-owned `ID2D1SolidColorBrush` and increments `_debugDrawItemTransientBrushCreateCount`. `folderView_draw_item_brush_reuse_guard` now proves both sides of the invariant: normal selected/hovered warm render keeps the counter at zero, and the positive-control seam drives it nonzero. `folderView_perf_huge_folder_scale` remains the adjacent scale guard for the zero-counter path during select-all scrolling. Durable contract lives in `Specs/UI/UI_FolderView.md`. RED build `.build\logs\msbuild-20260705_153748_205.log` failed before the pane debug seam existed; GREEN build `.build\logs\msbuild-20260705_153949_588.log` produced the app; GREEN focused archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154145/`; GREEN adjacent archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154806/`. | GR-T24 DONE |
| GR-A6 | **MED — DONE 2026-07-05** | *(→ Track A)* **FolderView re-implements device-loss detection + hardware→WARP fallback that shared `DxUi::WindowHost` already provides, and the copies diverge.** `IsFolderViewDeviceLoss` (`FolderView.Rendering.cpp:33`) treats {`D2DERR_RECREATE_TARGET`, `DXGI_ERROR_DEVICE_REMOVED`, `DXGI_ERROR_DEVICE_RESET`, `DXGI_ERROR_DEVICE_HUNG`} as loss; the shared WindowHost predicate is inlined twice (`DxUi.WindowHost.cpp:3430` and `:3649`) and **omits `DXGI_ERROR_DEVICE_HUNG`**, so a GPU hang recovers a WindowHost surface but not a FolderView pane. Both layers also carry their own hardware-then-WARP `D3D11CreateDevice` dance (`WindowHost.cpp:573-597` vs `FolderView.Rendering.cpp:220-246`), and FolderView owns a private `_d3dDevice`/`_swapChain` + `recoverFromDeviceLoss` lambda (`:1063`) instead of consuming the shared device. Not GR-S4(b)/(d) (different code). | Done: DxUi now exposes `IsDeviceLossHResult(...)`, including `DXGI_ERROR_DEVICE_HUNG`, and `CreateD3D11DeviceWithWarpFallback(...)`; both WindowHost render overloads classify EndDraw/Present failures through the shared predicate, WindowHost shared graphics creation uses the shared D3D helper, and FolderView no longer carries a private device-loss predicate or direct D3D11 fallback dance. Durable contracts live in `Specs/UI/UI_DxUiWinUIDesign.md` and `Specs/UI/UI_FolderView.md`. RED build `.build\logs\msbuild-20260705_140035_875.log` failed before the helper existed; GREEN builds `.build\logs\msbuild-20260705_140302_574.log` and `.build\logs\msbuild-20260705_140513_048.log` passed with 0 warnings / 0 errors. Focused WindowHost and FolderView device-loss coverage passed, including `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_140722/`. | GR-TA6 DONE |
| GR-S4(f) | **Cleanup — DONE 2026-07-05** | *(→ extends GR-S4)* **WarpDrive parallel-maintenance sprawl.** (a) Four new `shell*` thumbnail counters each threaded through 6 tiers — atomics (`FolderView.h:1576-1579`), `store(0)` resets (`FolderView.Icons.cpp:752-755`), `fetch_add` each paired with a **redundant `PerfEmitCounter`** of the same `_count` (`:1195/1206/1236/1241`), `load()`s (`:1643-1646`), snapshot fields (`FolderView.h:442-445`), pane-snapshot copies (`FolderWindow.FileSystem.Commands.Part.cpp:10594-10597`); (b) the 9-arg `D3D11CreateDevice` spelled 3× in `EnsureDeviceResources` (`FolderView.Rendering.cpp:199/220/232`); (c) `PendingRefreshToPaintMetric` is a near-verbatim clone of `PendingInputToPaintMetric` (`FolderView.h:968-976` vs `:959-967`; methods `FolderView.cpp:124-175`; emit duplicated at `Rendering.cpp:2024-2025/2063-2064`). | Done: shell thumbnail stat counter/perf pairs now route through `IncrementThumbnailStat(...)`; the D3D11 creation duplication is already gone through the shared GR-A6 `CreateD3D11DeviceWithWarpFallback(...)` path; and pending input/refresh-to-paint telemetry now uses one `PendingToPaintMetric` shape while preserving separate pending slots. Durable contracts live in `Specs/UI/UI_FolderView.md`; focused evidence is in the GR-S4(f) ratchet note above. | GR-S4(f) DONE |

**New tests (each RED on the un-fixed code):**

> **Audit caveat 2026-07-04:** These RED labels remain proposed until implemented and run; the
> static re-check did not execute builds or selftests. GR-T4 needs direct backend parity coverage,
> GR-T6 needs a deterministic stalled-handler/AppVerifier-style harness, GR-T20/GR-T23 need
> explicit fault hooks, and GR-T24 is a mutation/source-contract guard rather than a normal
> behavioral RED test.

| ID | For | Required proof |
|----|-----|----------------|
| GR-T20 | GR-20 | **DONE 2026-07-05.** Added `settings_hot_reload_transient_arm_failure_is_async`: forces one transient `FindFirstChangeNotificationW(...)` arming failure, asserts `SettingsHotReload::Start(...)` returns within a 200ms budget, asserts no `WAIT_TIMEOUT`/watcher teardown, then saves a later external settings change and observes the settings-changed message after the worker self-arms. |
| GR-T21 | GR-21 | **DONE 2026-07-05.** Extended `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state`: active non-file plugin compare starts with a known left-only default selection, invokes `IDM_COMPARE_INVERT_DIFFERENCES_SELECTION` and observes the inverted selection, then sends `IDM_LEFT_REFRESH` and requires the default compare selection to return. RED failed with `leftSelected=0`; GREEN passes after Invert resets to Default and compare-window refresh re-applies the default decision selection. |
| GR-T22 | GR-22 | **DONE 2026-07-05.** Added `folderView_refresh_to_paint_metric_clears_after_failed_render`: drives one same-folder refresh until the result is ready while withholding target pane paint, forces a synthetic non-device-loss `EndDraw` failure, asserts no `folder.refresh.request_to_paint_us` row fires for that failed paint, then drives an unrelated later Present with no new refresh and asserts no stale metric fires. RED failed with `Stale refresh-to-paint metric emitted on an unrelated later Present; before=0 after=1`; GREEN passes after ready pending metrics are cleared on no-present/failure paths and emit revalidates the current generation. |
| GR-T23 | GR-23 | **DONE 2026-07-05.** Added `folderView_iconcache_live_path_failure_retries_without_negative_cache`: an `ENABLE_TESTS` hook forces the first live `QuerySysIconIndexForPath(path, attrs, useFileAttributes=false)` to fail, the first call returns `nullopt` without negative-caching, the second call re-invokes `SHGetFileInfoW` and returns the real icon, and the third call reuses the positive cache without consuming another live lookup hook. RED first failed before the hook was wired, then failed on the actual negative-cache bug with `Second live icon lookup should re-query the shell and recover after a transient failure.` |
| GR-T24 | GR-24 | **DONE 2026-07-05.** `folderView_draw_item_brush_reuse_guard` keeps the normal selected/hovered warm-render assertion that `drawItemTransientBrushCreateCount == 0u`, then enables an `ENABLE_TESTS` forced transient-brush path and requires the same counter to become nonzero. RED build failed before `DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest(...)` existed; GREEN focused and adjacent runs passed with the live counter. |
| GR-TA6 | GR-A6 | **DONE 2026-07-05.** `TestDxUiDeviceLossAndD3dCreationAreShared` feeds `DXGI_ERROR_DEVICE_HUNG` to the shared `IsDeviceLossHResult` predicate, asserts the WindowHost render paths and FolderView route through that same predicate, and verifies WindowHost/FolderView use the shared D3D11 hardware-to-WARP creation helper. RED build failed before the shared predicate existed; GREEN DxUiTests WindowHost coverage, app build, and focused FolderView Commands coverage are recorded in the GR-A6 closeout note. GR-S4(f) remains structural cleanup for thumbnail counters and pending-metric duplication. |

## GR-24 draw-item transient-brush guard closeout note (2026-07-05)

Scope: FolderView draw-item cached brush reuse, the `ENABLE_TESTS`
transient-brush diagnostic counter, focused Commands coverage, adjacent
large-folder scale coverage, Tailwind TW-1/TW-T1 routing, and durable FolderView
test/spec ledgers.

- RED app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` failed before the debug seam existed with `DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest` missing from `FolderWindow`; build log `.build\logs\msbuild-20260705_153748_205.log`.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` produced `.build\x64\Debug\RedSalamander.exe`; build log `.build\logs\msbuild-20260705_153949_588.log`.
- GREEN focused Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154145/` passed `folderView_draw_item_brush_reuse_guard` with 1 passed / 0 failed / 0 skipped.
- GREEN adjacent Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154806/` passed `folderView_perf_huge_folder_scale` with 1 passed / 0 failed / 0 skipped.
- `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, `Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`, and `Specs/Plans/WIP/README.md` were updated for the live brush-reuse guard and duplicate Tailwind closeout.

## GR-A6 shared DxUi/FolderView device-loss closeout note (2026-07-05)

Scope: shared DxUi device-loss classification, shared D3D11 hardware-to-WARP
fallback creation, WindowHost render recovery, FolderView render/device creation,
focused native DxUi coverage, focused FolderView Commands coverage, and durable
DxUi/FolderView/test coverage specs.

- RED `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64` failed before the shared predicate existed with `error C3861: 'IsDeviceLossHResult': identifier not found`; build log `.build\logs\msbuild-20260705_140035_875.log`.
- GREEN `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors; build log `.build\logs\msbuild-20260705_140302_574.log`.
- GREEN focused `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` passed, including `TestDxUiDeviceLossAndD3dCreationAreShared`.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors; build log `.build\logs\msbuild-20260705_140513_048.log`.
- GREEN focused Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_140722/` passed `folderView_render_device_loss_recovers` with 1 passed / 0 failed / 0 skipped.
- Static source scan shows direct `D3D11CreateDevice(` call sites remain only in the shared DxUi helper and `FolderView.Rendering.cpp` no longer contains `IsFolderViewDeviceLoss`.
- `Specs/UI/UI_DxUiWinUIDesign.md`, `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_TestCoverage.md`, and `Tests/README.md` were updated for the shared device-loss/fallback contract and the new DxUi WindowHost test count.

## GR-20 settings hot-reload async arm closeout note (2026-07-05)

Scope: `SettingsHotReload::Start(...)`, watcher arming/retry behavior,
`ENABLE_TESTS` watcher-open fault injection, Commands hot-reload selftest coverage,
the SettingsStore watcher contract, and test inventory ledgers.

- RED compile seam `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` failed before `SettingsHotReload::DebugSetChangeNotificationOpenFailuresForSelfTest(...)` existed; build log `.build\logs\msbuild-20260705_141406_457.log`.
- RED behavioral proof `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter settings_hot_reload_transient_arm_failure_is_async -FailFast -TimeoutMultiplier 3` failed before the production fix because the forced transient watcher-arm failure blocked `Start(...)`/returned the old timeout behavior; archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_141831/`.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors; build log `.build\logs\msbuild-20260705_142433_667.log`.
- GREEN focused Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_142640/` passed `settings_hot_reload_transient_arm_failure_is_async` with 1 passed / 0 failed / 0 skipped.
- GREEN adjacent Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_142715/` passed the `settings_hot_reload_` family with 5 passed / 0 failed / 0 skipped.
- `Specs/Core/Core_SettingsStore.md`, `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and `Tools/Tests/TestInventory.Tests.ps1` were updated for the asynchronous hot-reload arming contract and the new Commands case.

## GR-21 Compare Directories one-shot invert-selection closeout note (2026-07-05)

Scope: Compare Directories Restore/Invert selection commands, compare-window pane
refresh behavior, non-file plugin compare path-form command coverage, the durable
Compare Directories selection contract, and the Tailwind TW-8 duplicate row.

- RED focused `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state -FailFast -TimeoutMultiplier 3` failed after the test extension with `leftSelected=0`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_143810/`; captured summary `.build\logs\gr21_red.out.txt`.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors; build log `.build\logs\msbuild-20260705_144528_275.log`.
- GREEN focused Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_144738/` passed `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state` with 1 passed / 0 failed / 0 skipped; captured summary `.build\logs\gr21_focused_green.out.txt`.
- GREEN adjacent Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_145147/` passed `cmd_compare_directories_` with 15 passed / 0 failed / 0 skipped; captured summary `.build\logs\gr21_adjacent_compare_directories_green.out.txt`.
- `Specs/Core/Core_CompareDirectories.md`, `Specs/Testing/Testing_TestCoverage.md`, and the Tailwind TW-8 ledger were updated for the one-shot invert-selection contract.

## GR-22 FolderView refresh-to-paint stale metric closeout note (2026-07-05)

Scope: FolderView same-folder refresh-to-paint telemetry, failed/no-present render
paths, enumeration-generation validation, focused Commands coverage, adjacent
rendering/refresh metric coverage, the FolderView UI spec, and test inventory ledgers.

- RED focused `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_refresh_to_paint_metric_clears_after_failed_render -FailFast -TimeoutMultiplier 3` failed before the production fix with `Stale refresh-to-paint metric emitted on an unrelated later Present; before=0 after=1`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150226/`.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors; build log `.build\logs\msbuild-20260705_150338_519.log`.
- GREEN focused Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150541/` passed `folderView_refresh_to_paint_metric_clears_after_failed_render` with 1 passed / 0 failed / 0 skipped.
- GREEN adjacent Commands archives `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150654/`, `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150739/`, `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151107/`, and `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151143/` passed `folderView_rendering_error_overlay_requires_persistence`, `folderView_render_device_loss_recovers`, `folderView_perf_refresh_preservation`, and `folderView_perf_directory_change_storm`.
- `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and `Tools/Tests/TestInventory.Tests.ps1` were updated for the refresh-to-paint stale-metric contract and the new Commands case.

## GR-23 IconCache live lookup transient failure closeout note (2026-07-05)

Scope: IconCache live path lookup failures, attribute-mode failure caching, focused
Commands coverage, adjacent FolderView icon-pipeline coverage, the FolderView UI spec,
and test inventory ledgers.

- RED focused `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_iconcache_live_path_failure_retries_without_negative_cache -FailFast -TimeoutMultiplier 3` first failed before the hook was wired with `Forced transient live icon lookup failure should report no icon index.`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151809/`.
- RED focused rerun after wiring the hook failed on the actual negative-cache bug with `Second live icon lookup should re-query the shell and recover after a transient failure.`; archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_152050/`.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` passed with 0 warnings / 0 errors; build log `.build\logs\msbuild-20260705_152904_817.log`.
- GREEN focused Commands archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153136/` passed `folderView_iconcache_live_path_failure_retries_without_negative_cache` with 1 passed / 0 failed / 0 skipped.
- GREEN adjacent Commands archives `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153107/` and `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153209/` passed `folderView_perf_slow_virtual_provider` and `folderView_perf_icon_pipeline_cold_slow`. The pre-update adjacent archive `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_152816/` failed only because the slow-provider guard still expected the removed live negative-cache behavior.
- `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and `Tools/Tests/TestInventory.Tests.ps1` were updated for the uncached live-failure contract, telemetry names, and the new Commands case.

## Verified-REFUTED — do not re-chase (evidence in review transcript 2026-07-02)

- Move-cleanup hardcoded `ALLOW_REPLACE_READONLY` (`State.cpp:7314`): deliberate parity with the disk plugin's cross-volume move semantics (`FileSystem.FileOps.cpp:5676-5678`), gated by byte verification; attributes propagate to destination.
- `WindowHost::~WindowHost()` calling `Detach()`: fully gated on actual attachment (`WindowHost.cpp:1232-1254`); deliberate Bedrock B-S0-1 fix with its own test.
- `Grid::Paint` animation latch before `!dc` return: unreachable in production (host clears+paints with a live context) and self-healing via device-loss recovery repaint.
- `Tree` `SetFocus` on key handling: still not a confirmed bug, but the old rationale was too strong. Production key dispatch does exist (`DxUi.WindowHost.cpp:2621+` forwards `WM_KEYDOWN` to focused controls and `DxUi.Tree.cpp:861` handles `Tree::OnKeyDown`); during real `WM_KEYDOWN` the host HWND should already have focus, so the `SetFocus(hwnd)` branch is likely a no-op in production. Keep refuted only as "not currently proven harmful," not as "no production path."
- `MtpBufferedWriter::Commit` `_committed=true` before attempt: no caller retries Commit on the same writer; the pre-set deliberately guards double-commit.
