# Operation Clearwater — 2-Day Master Review Remediation (2026-06-28)

Tracking plan for findings from the max-effort multi-agent review of the last two days of work on
`master` (diff `acfdc1a4e..4778dd4f9`: the Bedrock DxUi remediation — accessibility snapshotting,
controls/window-host, text input; the FS deep-audit tracks — cross-FS move source-cleanup safety,
reparse/canonical traversal, search-index data safety and the search-service trust boundary; the
FileOps observability/staged-promotion work; and S3/7z provider transfer contracts).

Method: 10 finder angles (5 correctness + reuse/simplification/efficiency/altitude/conventions) over 8
subsystem partitions → 1-vote 3-state verification → two independent gap sweeps → a per-finding
fix-spec pass that re-confirmed each mechanism on `master`, drafted the surgical fix, defined a RED-on-bug
proof, and checked whether the active branch already carries the fix. Every item below is **confirmed**
(survived an independent skeptic) unless marked PLAUSIBLE/LATENT. Severity is the verifier-adjusted
severity.

> **Branch topology note.** The review target is `master` (tip `4778dd4f9`). The active working branch
> `codex/folderview-warpdrive` (HEAD) is **23 commits ahead** of `master`. Only **CW-1** is already fixed
> on HEAD; for every other item `git diff master..HEAD -- <file>` is empty, so the finding is live on both
> `master` and the active branch. Several findings are **regressions** introduced inside this 2-day window
> (CW-5, CW-7, CW-11, CW-13, CW-15) — base `acfdc1a4e` did not have them.
>
> **Status re-verification (2026-07-02 folder review; updated 2026-07-05 for CW-7/CW-10 closeout).**
> All 19 non-CW-1 items were independently re-verified open against HEAD `275c04034`; since then CW-7
> and CW-10 are closed in this working tree. The remaining prescribed test seams (SaveGenericCredential
> fault seam, `isNonFatalChildError`) still exist nowhere; CW-10 uses the SearchService warningFlags path
> instead of adding a separate `clientAuthorizationIncomplete` field. Known anchor drift from the 2026-06-28 auto-format
> (`45ae2a9a9`) and later commits — re-locate by text: CW-1 `:7346`→`:7376`, CW-7 `:5768`→`:5766-5771`,
> CW-10 `:2221`→`:2238-2246`, CW-14 `:3089`→`:3157`, CW-S1 `:685`→`:683`, CW-S3 `:5040`→`:5044`.
> **Operation_Granite (2026-07-02) routes five of its findings back here — Clearwater stays the single
> owner of all CW fixes/tests; Granite contributes only the addenda now inlined in the rows below.
> Headline: CW-7 is escalated to P0** (root-caused as merge `d4a064dce` silently reverting fix `ec403afac`;
> execute with Granite's GR-T19 merge audit).

## Execution corrections from repo-backed plan review

- Treat this file as the **triage ledger**, not as a complete one-shot executor handoff. Before dispatching
  any item to an implementation agent, expand it into a self-contained sub-plan with: planned-at SHA,
  drift check over the in-scope files, explicit in-scope/out-of-scope file list, exact verification command,
  expected output, and STOP conditions. Repo verification guidance names `.\Tools\Run-AllTests.ps1 -Suite
  Full` / `.\Tools\Run-AllTests.ps1 -SkipBuild` as the green-check commands (`AGENTS.md:134`,
  `README.md:204-214`).
- Perf-sensitive items are not complete with a RED correctness test alone. `AGENTS.md:48-55` requires
  scenario definition, instrumentation, deterministic selftest coverage, and archived perf evidence under
  `Specs/TestRuns/`; `Specs/Testing/Testing_PerformanceValidation.md:55-64,156-166` requires measurements
  or a documented blocked reason, plus archive-path planning. This applies at minimum to **CW-P1/CW-P2**
  and to any FileOps/search/UI hot-path change discovered while fixing the correctness items.
- Do **not** rely on the earlier shorthand that every fix is one-function. Some are deliberately
  cross-cutting: **CW-P1** must touch the `CopiedEntry` shape (`State.cpp:6569`), `RecordCopiedFile`
  (`:7052`), `CopiedFileStillMatchesDestination` (`:7224`), and the serial/pipeline copy loops
  (`:7940`, `:8135`); **CW-10** spans `CheckClientCanListDirectory`, batch state, progress/completion
  warning propagation, and cache semantics. Keep each implementation scoped, but plan these as multi-site
  changes.
- **CW-6 requires a design decision before execution.** The current `LocalSearchIndexCore::Entry` stores one
  `parentId`, one `name`, and one `fullPath` per `NodeId` (`Common/LocalSearchIndexCore.cpp:174-181`), and
  `RebuildDerivedState` clears/recomputes that single `fullPath` from the parent graph (`:2228-2305`).
  Therefore a duplicate physical directory reached through both a canonical path and a junction alias cannot
  produce two visible result paths without either an alias-edge/path model or a reparse-aware identity
  strategy that gives the alias directory its own node. The current `OpenPathHandle` follows reparse targets
  because it uses `FILE_FLAG_BACKUP_SEMANTICS` without `FILE_FLAG_OPEN_REPARSE_POINT` (`:1761-1769`).
- **CW-S2 is an investigation, not an automatic remove-`noexcept` cleanup.** The code evidence is real:
  `MenuBar::EnsureMenuBarLayoutCache` and `TabControl::EnsureTabHeaderLayoutCache` are `const noexcept`
  (`DxUi.Controls.cpp:4641`, `:5061`; declarations in `DxUi.h:1878`, `:2002`) and populate vectors. But
  repo policy explicitly says `std::bad_alloc` is fatal (`AGENTS.md:73`). Do not make `bad_alloc` unwind
  through UI/Win32 boundaries unless that policy is intentionally changed; prefer a no-throw/cache-fail
  design if recoverability is required.
- Closeout is part of the remediation. When this WIP is finished, move it to `Specs/Plans/Done/` and merge
  any durable behavior, validation rule, or workflow requirement into the authoritative spec under
  `Specs/<Domain>/` or repo-level guidance, per `AGENTS.md:57-58`.

## Already fixed on the active branch (must land on master + needs a regression test)

- **[HIGH] CW-1 — Cross-FS MOVE silently leaves a full duplicate subtree on the source.**
  `DeleteCopiedSourceEntryForMove` (`RedSalamander/FolderWindow.FileOperations.State.cpp:7346`) returned
  `S_OK` on `! BasicFileInfoStillMatchesSource(entry)` **before** the per-child recursion (7443), so any
  drift of the source directory's `creationTime`/`lastWriteTime`/`attributes` (indexer, AV, attribute
  toggle, or a failed re-probe) left the whole copied subtree on the source — the MOVE degraded to a COPY
  and reported `ERROR_PARTIAL_COPY`. **Fixed on HEAD by `3cb869a0d` "Stabilize filesystem safety
  selftests"** (captures `sourceDirectoryStillMatches`, early-skips only for a drifted *reparse* link,
  recurses to delete matching children for a plain directory, and guards only the final shell removal with
  `if (! sourceDirectoryStillMatches) return S_OK;`). **Two follow-ups:** (a) ensure that commit reaches
  `master`; (b) add the targeted regression test **CW-T1** — the fix commit's selftest edits are general
  stabilization and **no test injects source-*directory* basicInfo drift between copy and cleanup**.

## Routed — correctness / data-safety

| ID | Sev | Item (master anchor) | Fix direction | Proof |
|----|-----|----------------------|---------------|-------|
| CW-2 | **HIGH** | **Rotated OAuth refresh token discarded on a transient credential-store failure → forced re-login.** `HostServices::SetConnectionSecretOnUiThread` (`RedSalamander/HostServices.cpp:~2327`) now saves to CredMan *first* and `return saveHr` on failure **before** the in-memory `SessionSecretEntry` update (~2353). `FileSystemMicrosoftDrive.cpp:6484` stores the freshly-rotated token `persist=TRUE` and only `Debug::Warning`s on failure; a transient `CredWriteW` failure (e.g. `ERROR_NO_SUCH_LOGON_SESSION`) leaves the now server-invalidated old token live in the session cache (which `GetConnectionSecretOnUiThread` reads first). Same inversion hits PASSWORD/SSH-passphrase rotation and the empty-secret delete sub-branch. | Restore base ordering: update the in-memory `SessionSecretEntry` **first** (`SecureClear` old, assign new / clear-on-empty), **then** attempt persistence. A `SaveGenericCredential`/`DeleteGenericCredential` failure must keep the newly-cached secret and surface via `Debug::ErrorWithLastError` — never return `saveHr`/`deleteHr` before the cache holds the new value. Keep WIL + `SecureClear` idioms. | CW-T2 below. |
| CW-3 | MEDIUM | **S3 reads fail entirely on GET-allowed / HEAD-denied buckets.** `S3RangedFileReader::Read` (`Plugins/FileSystemS3/FileSystemS3.IO.cpp:438`) now calls `EnsureSizeKnown()` (a `HeadObject`) unconditionally and returns its HRESULT; base learned size lazily from the ranged-GET short body. A bucket granting `s3:GetObject` but denying `s3:HeadObject` (403 → `ERROR_ACCESS_DENIED`) now fails every Read/GetSize/`FILE_END` Seek. | Recover size from the GET: in `FillBufferFrom` parse `GetObjectResult::GetContentRange()` (`"bytes a-b/total"`) to set `_sizeBytes`/`_sizeKnown` (fallback: validated short read → `_bufferStart + total`). Make `EnsureSizeKnown` non-fatal on HEAD-denied — `Debug::Warning` and return `S_OK` leaving `_sizeKnown=false`; have Read/GetSize tolerate that by establishing size via the discovery GET. Validate the discovery GET against the parsed total, not `kChunkBytes`. `GetContentRange()` exists in the vendored aws-sdk-cpp 1.11.790. | CW-T3 below. |
| CW-4 | MEDIUM | *(Re-confirmed live by Granite 2026-07-02; `tar -uf` / appended-replacement zips are everyday triggers — execute as written.)* **7z archive becomes un-browsable on a normalized-key collision.** `BuildIndex` (`Plugins/FileSystem7z/FileSystem7z.cpp`) has two abort points: `ensureDir` returns `ERROR_INVALID_DATA` when a file shadows a needed parent dir (`:4399`), and the duplicate-key guard returns it for any repeated `raw.key` (`:4461`). Either fails `EnsureIndex` (`:907`) and the whole archive is unopenable. Base used last-wins (`outEntries[key]=entry`). Triggers on common layouts: file `foo` + path `foo/bar`, or `foo` + `foo/` (slash stripped by `NormalizeArchiveEntryKey`). | Degrade, don't fail. `ensureDir`: when an existing entry is a non-directory, skip dir creation (`return S_OK`, the file wins), `Debug::Warning` the key. Duplicate-key guard: `continue` (keep the already-indexed entry) instead of returning. Keep `BuildIndex`→`S_OK`; only push a key into its parent's child list when actually inserted (the tail dedup/sort already tolerates it). **Both** sites must be patched. | CW-T4 below. |
| CW-5 | MEDIUM | **`GetDirectorySize` aborts the whole traversal on one transient child error** (regression). `pushDirectory` (`Plugins/FileSystem/FileSystem.DirectoryOps.cpp:734`) calls `setFatalStatus`+`return false` for any `FindFirstFileExW` error outside `{FILE_NOT_FOUND, ACCESS_DENIED, PATH_NOT_FOUND}`; the call site (`:872`) then `return result->status`, abandoning unvisited siblings and truncating the total. `ERROR_SHARING_VIOLATION`/`LOCK_VIOLATION`/`NETWORK_BUSY` on one subdir under-reports the whole tree. Base continued past it. | Factor an `isNonFatalChildError(DWORD)` predicate (add `SHARING_VIOLATION`, `LOCK_VIOLATION`, `NETWORK_BUSY`, `NETNAME_DELETED`, `BAD_NETPATH`, `DEV_NOT_EXIST`, `NOT_READY`); route those through `markPartial()`+`return true`, reserving `setFatalStatus` for genuinely global errors (e.g. `OUTOFMEMORY`). Apply the same predicate in `advanceFrame` (`:767`). Log skipped subtrees via `Debug::Warning`. | CW-T5 below. |
| CW-6 | MEDIUM | **Index drops an entire directory subtree on a junction/FRN alias.** `HydrateDirectorySubtree` (`Common/LocalSearchIndexCore.cpp:2636`) `continue`s for *any* pre-existing `NodeId`; `OpenPathHandle` (`:1761`) omits `FILE_FLAG_OPEN_REPARSE_POINT`, so junctions/mountpoints are followed and the same physical dir reached via two paths shares a FileId — descendants reachable only via the alias path are dropped with merely `hardlinkAliasCoverageIncomplete=true`. | **Do not assign as a simple one-line gate fix.** First choose the intended contract: (A) *coverage-only* — index objects reachable only through the alias path, but canonical-path results are acceptable when the physical object already has a canonical node; or (B) *alias-path results* — searches must be able to emit both canonical and junction paths. Current `Entry` has only one `parentId/name/fullPath` per `NodeId` (`Common/LocalSearchIndexCore.cpp:174-181`) and `RebuildDerivedState` rebuilds one path per node (`:2228-2305`), so contract B requires an alias-edge/path representation or a reparse-aware directory identity strategy. If choosing A, gate the existing short-circuit on file aliases and recurse directory aliases with a visited set, but document that duplicate alias paths are not emitted. If choosing B, prefer a reparse-aware `OpenPathHandle` variant or explicit alias path table, then audit sibling USN/MFT seed paths (`~2820/2990`) separately. | CW-T6 below. |
| CW-7 | ~~MEDIUM~~ **DONE 2026-07-04** *(was P0, escalated by Granite 2026-07-02)* | **TabControl closes the wrong tab when a drag/reorder ends over a close button** (regression). Fixed in `Common/DxUi/DxUi.Controls.cpp`: `TabControl::OnMouseUp` now requires `closePressed.has_value()` and the same close-button index before calling `CloseTab`. Durable contract added to `Specs/UI/UI_DxUiWinUIDesign.md`. | Completed with RED/GREEN coverage in `Tests/DxUiTests/DxUiTests.Controls.cpp` (`TestTabControlBodyDragReleaseOverCloseButtonDoesNotCloseTab`). Verification: `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64` and `.\.build\x64\Debug\DxUiTests.exe --suite=Control` passed on 2026-07-04. Granite GR-T19 merge audit completed; see Granite plan audit note. | CW-T7 DONE. |
| CW-8 | MEDIUM | **Animations freeze permanently after restore-from-minimize.** `OnAnimationTick` (`Common/DxUi/DxUi.WindowHost.cpp:4036`) latches `_animationSuspendedWhileHidden=true` + zeroes the subscription when `IsIconic`; the flag is cleared **only** in `WM_SHOWWINDOW(TRUE)` (`:2341`). Restore-from-minimize sends `WM_SIZE/SIZE_RESTORED` + `WM_ACTIVATE` (no `WM_SHOWWINDOW`), so caret-blink/expander/scrollbar-fade stay dead. | In the `WM_SIZE` handler (`~:2327`), after `OnSize()`, when `wp == SIZE_RESTORED || SIZE_MAXIMIZED` and `IsHostWindowEffectivelyVisible(hwnd)` and `_animationSuspendedWhileHidden`: clear the flag and call `RequestAnimation()` (mirrors the `WM_SHOWWINDOW` re-arm; idempotent since `RequestAnimation` early-returns when already subscribed). No `WM_WINDOWPOSCHANGED` handler exists, so `WM_SIZE` is the correct anchor. | CW-T8 below. |
| CW-13 | LOW-MED | **Staged-promote swallows a final attribute failure and returns `S_OK`** (regression). `PromoteStagedTempIntoFinalPath` (`Plugins/FileSystem/FileSystem.Path.cpp:639`) logs `Debug::Warning` and falls through to `S_OK` when the post-commit `SetFileAttributesW` fails, leaving residual `FILE_ATTRIBUTE_TEMPORARY`/`HIDDEN`/`READONLY` while callers treat the promotion as fully successful. Worst caller: `Win32FileWriter::PromoteTempIntoFinalPath` with `stripTemporaryAttributes=true` (`FileSystem.cpp:1847`) — the committed file may stay TEMPORARY. | Restore the base failure return: capture the error and `return HRESULT_FROM_WIN32(error ? error : ERROR_ACCESS_DENIED)` (keep the new `attributesNeedUpdate` guard so only a genuine mismatch+failure surfaces; optionally log via `Debug::ErrorWithLastError`). Both callers already propagate `FAILED(promoteHr)`. | CW-T13 below. |
| CW-10 | LOW-MED — **DONE 2026-07-05** | *(Granite addendum: fixed in the same pass as sibling findings GR-11/GR-13/GR-14.)* **Search authorization cached `false` on a transient error → silently incomplete results.** `CheckClientCanListDirectory` cached `outAllowed = (bool)handle` for any `CreateFileW` failure (`SHARING_VIOLATION`, `TOO_MANY_OPEN_FILES`, …), poisoning `clientDirectoryAccessCache` for the session and dropping every candidate under that dir with no warning flag. | Done: `CheckClientCanListDirectory` now caches `false` only for durable negative results (`ERROR_ACCESS_DENIED`, `ERROR_FILE_NOT_FOUND`, `ERROR_PATH_NOT_FOUND`). Transient directory authorization open failures are logged, are not cached, and flow through the per-candidate skip path with `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`. Durable contract lives in `Specs/Core/Core_Search.md`; RED evidence `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_111008/` failed with `warnings=0x00000000`, GREEN evidence `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_111440/` passed, adjacent SearchService checks passed in `2026-07-05_111529/`, `2026-07-05_111546/`, `2026-07-05_111603/`, and `2026-07-05_111619/`; final Debug x64 build `.build\logs\msbuild-20260705_111223_980.log` had 0 warnings / 0 errors. | CW-T10 DONE. |
| CW-11 | LOW | **Error remap mislabels staged-overwrite failures** (regression). `CopyReparsePointInternal` (`Plugins/FileSystem/FileSystem.FileOps.cpp:3961`) dropped the `! allowOverwriteEffective &&` guard, so any `ERROR_INVALID_PARAMETER` — including ones from `CopyFileExWithStagedOverwrite`→`ReplaceFileW`/`MoveFileExW` — is remapped to `ERROR_NOT_SUPPORTED`, masking the real cause. Error-quality only. | Restore the guard: `if (! allowOverwriteEffective && copyHr == HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER))` (`allowOverwriteEffective` is in scope at `:3915`). Scopes the remap to the symlink-copy `CopyFileExW(COPY_FILE_COPY_SYMLINK)` call only. | CW-T11 below. |
| CW-12 | LOW | **Checkpoint timestamp recorded before a checkpoint that can fail.** `RunAutomaticMaintenance` (`Common/SqliteIndexStore.cpp:2895`) upserts `kMetaLastCheckpointUtc` **before** the final `RunWalCheckpoint`; a `SQLITE_BUSY`→`S_FALSE` deferral leaves `lastCheckpointUtc` recording a checkpoint that never truncated the WAL. Observability only. | Move the upsert to *after* a successful checkpoint inside the `shouldFinalCheckpoint` block (after `result.ranCheckpoint = true`), keeping the busy→`S_FALSE` early-return from persisting a bogus timestamp. | CW-T12 below. |
| CW-14 | LOW | **Enumeration counts authorization-skipped candidates as `emittedRows`.** `EnumerateVolume` (`Common/SqliteIndexStore.cpp:3089`) increments `emittedRows` after the `FAILED`/`S_FALSE` checks, but `kSkipCandidateHr` is a `SEVERITY_SUCCESS` code, so deduped/skipped candidates inflate `emittedRows` and trip `injectedFailAfterEmittedRows` early. | Add the codebase's existing guard before the increment: `if (hr == LocalSearchIndexCore::kSkipCandidateHr) { continue; }` (same idiom as `EnumerateLiveFileSystem` `:4072/4141` and `emitEntry` `:3601`). `scannedRows` stays correct. | CW-T14 below. |
| CW-9 | LOW-MED | **Theme `accentPressed` derived from `darkMode` instead of `darkBase`.** `RefreshAccentVariants` (`Common/DxUi/DxUi.Theme.cpp:904`) picks the `+16`/`-12` pressed shift from `palette.dark` (==`darkMode`), but `MakeThemePaletteFromViewerTheme` builds all surrounding chrome from the independent `darkBase`. When `darkMode != darkBase` the pressed/checked accent shifts the wrong way (menu hover fills `Menu.cpp:750/754`, `popupActiveFill` `Theme.cpp:825`, rainbow check-glyph). Cosmetic. | Thread the base flag in: `RefreshAccentVariants(ThemePalette&, bool darkBase) noexcept`, use `darkBase ? +16 : -12` for `accentPressed` (hover stays `+8`); pass `dark` from `MakeDefaultThemePalette` and the local `darkBase` from the two viewer-theme call sites (`:1028`,`:1042`). | CW-T9 below. |
| CW-15 | LOW | **Masked-password dot count switched from grapheme to UTF-16 code-unit count** (regression). `GetSecretVisibleDotCount` (`Common/DxUi/DxUi.TextInput.cpp:1168`) uses `_text.size()`, so an astral-plane char (1 navigable grapheme) renders/announces 2 dots and desyncs from caret-navigable elements (UIA `NativeTextInput.cpp:119`). | Restore grapheme counting: `exactCount = CountTextElements(_text);` and re-add the small `CountTextElements(wstring_view) noexcept` grapheme stepper (existed at base). **Confirm with author** whether the switch was an intentional `secret_render_us` perf simplification; if so → WONTFIX. | CW-T15 below. |

## Routed — performance / scaling

| ID | Sev | Item (master anchor) | Fix direction |
|----|-----|----------------------|---------------|
| CW-P2 | MEDIUM | *(Granite addenda 2026-07-02: (a) `RecordCopiedFile` runs for plain COPY too, where the manifest is never consumed — gate recording on `task._operation == FILESYSTEM_MOVE`; (b) `RecordCopiedFile`/`RecordCopiedDirectory` are `noexcept` around throwing `wstring` copies + map insert → `std::terminate` under memory pressure mid-transfer. The extract-and-erase fix below stands.)* **Cross-FS MOVE manifest grows unbounded.** `copiedEntries` (`RedSalamander/FolderWindow.FileOperations.State.cpp:6589`) accumulates one `CopiedEntry` (two `wstring` paths + `FileSystemBasicInformation`) per file **and** per directory for the whole op, never evicted — peak residency scales with total tree size (allocator pressure / OOM on multi-million-entry moves the old flat recursive delete avoided). | Make the manifest shrink as consumed: in `DeleteCopiedSourceEntryForMove` (`~:7302`) convert `TryGetCopiedEntry` into a take-and-erase via `copiedEntries.extract(sourcePath)` under `copiedEntriesMutex`, freeing each entry the moment its cleanup is attempted. (Structural note: copy and cleanup are still separate passes so end-of-copy peak is unchanged; a fuller fix interleaves per-directory cleanup into the copy traversal while preserving the per-entry equality checks.) **Perf validation required:** define the protected scenario as large cross-FS bridge MOVE cleanup memory retention; add/reuse a `FileOps.*` metric for manifest size/peak entries or retained bytes; archive baseline and candidate selftest runs under `Specs/TestRuns/` or document a blocker. |
| CW-P1 | LOW-MED — **DONE 2026-07-13 via Causeway CW-9** | *(Single owner — absorbs former Keystone KS-P5 and Floodgate FG-P2-7 rows. Granite addendum: on MTP sources the re-read is a full USB re-download into RAM, so this fix is a prerequisite for shipping MTP move at all — execute with GR-P2.)* **Cross-FS MOVE re-reads both source and destination in full before deleting.** `CopiedFileStillMatchesDestination` (`State.cpp:7224`) short-circuits on size mismatch but otherwise re-reads **both** sides via `ReadersHaveEqualContent` (~2× read I/O); a source that can't be re-read (evicted cloud placeholder / transient fault) returns a hard error, so cleanup logs `bridge.move.cleanup.skip` and the MOVE leaves a duplicate. **Premise correction:** there is *no* copy-time hash today; the fix must add one. | Done in Operation Causeway: `CopiedEntry` carries the copy-time FNV-1a, both bridge pumps hash the written bytes, and bridge MOVE cleanup reopens/hashes only the destination. Native local cross-volume MOVE takes the permitted strong-consistency branch: a successful copy plus stable no-follow source/destination snapshots permits deletion with zero post-copy content rereads. The dual-reader cleanup compare is gone. Durable contract: `Specs/Core/Core_FileSystemBridge.md`; final evidence is linked from the Causeway Done plan. |

## Routed — simplicity / robustness

| ID | Sev | Item (master anchor) | Fix direction |
|----|-----|----------------------|---------------|
| CW-S1 | LOW | **Dead code: `FindSemanticControlAtPoint`** (`Common/DxUi/DxUi.Accessibility.cpp:685`) is now only self-recursive after `ElementProviderFromPoint` migrated to snapshot point-hit lookup — emits MSVC C4505 (warning-hygiene; the tree has **no** warnings-as-errors, so not a build break — the "0 warnings" bar is convention). | Delete the definition (`:685-722`). Its helpers (`PointInRect`, `TryAppendPathIndex`, `IsSemanticAccessibilityControl`) all retain live callers — nothing cascades. The source-text guard test at `DxUiTests.Accessibility.cpp:~366` should also assert the definition is absent. |
| CW-S2 | LOW · INVESTIGATE | **`noexcept` layout-cache builders grow vectors, but repo policy treats `bad_alloc` as fatal.** `MenuBar::EnsureMenuBarLayoutCache` (`DxUi.Controls.cpp:4641`) and `TabControl::EnsureTabHeaderLayoutCache` (`:5061`) are `const noexcept` and populate `std::vector` members; declarations are also `noexcept` (`DxUi.h:1878`,`:2002`). | Do **not** blindly drop `noexcept`: `AGENTS.md:73` says `std::bad_alloc` is fatal. First decide whether these layout-cache paths must remain fatal-on-OOM, or whether the UI needs an explicit non-throwing fallback. If recoverability is required, prefer a no-throw construction pattern (reserve/build temporary state, fail closed/log once, keep previous invalid cache) rather than allowing `bad_alloc` to unwind through UI/Win32 boundaries. If fatal-on-OOM is intentional, mark WONTFIX and optionally add a short comment/test documenting the policy. |
| CW-S3 | LOW · LATENT | **TabItem title-width cache validated by raw `IDWriteTextFormat*` pointer identity (ABA-unsafe).** `MeasureTabTitleWidthDip` (`DxUi.Controls.cpp:5040`) compares `tab.measuredTitleTextFormat == format`; a freed-then-reallocated format at the same address would return a stale width. **Latent today** — there is no runtime font/format-reconfigure path that recreates formats while host+TabControl persist (Detach/DPI/density all self-invalidate); this is hardening against a future regression. | Change the member to `wil::com_ptr<IDWriteTextFormat>` (`DxUi.h:1903`, matching `_configuredTextFormats`), compare `.get() == format`, store by copy (AddRef keeps the address from being ABA-reused), `.reset()` at the two reset sites (`:4850`,`:4973`). |

## Tests to add (lock the fixes — each must go RED on the un-fixed code)

| ID | For | Required proof |
|----|-----|----------------|
| CW-T1 | CW-1 | Cross-FS bridge MOVE of a 2-level tree; mutate the **source root directory's** basicInfo (e.g. `SetFileBasicInformation` bumps `creationTime`) before cleanup so `BasicFileInfoStillMatchesSource` is false for the dir only. Assert matching children are deleted, the source **shell** remains, `hr == ERROR_PARTIAL_COPY` with the preserve-shell skip note. (Add to `SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Fairstream.cpp`.) |
| CW-T2 | CW-2 | `Commands.SelfTest.Connections.cpp`: seed an OAuth session token, re-set a rotated token `persist=TRUE` under an **injected `SaveGenericCredential` failure** (add a minimal fault seam), then `GetConnectionSecretOnUiThread` must return the **new** token. RED today (returns old). |
| CW-T3 | CW-3 | Add `FsS3::TryParseS3ContentRangeTotal` + a HEAD-denied size-discovery decision; assert in `RRunDebugRangeReadContractSelfTest` (`FileSystemS3.IO.cpp:91`). Plus an `IFileReader` harness: `HeadObject=403` + `GetObject` success must still yield `GetSize()==S_OK` and a full Read. |
| CW-T4 | CW-4 | 7z fixtures: (a) file `foo` + path `foo/bar`; (b) file `foo` + dir entry `foo/`. Browse via the enumerate path that reaches `BuildIndex`/`EnsureIndex`; assert the archive opens and lists its root. RED today (open fails). |
| CW-T5 | CW-5 | `_DEBUG` fault hook (mirror `directorySizeDelayMs`) forcing `FindFirstFileExW`→`ERROR_SHARING_VIOLATION` on one named child. Tree `dirA{100B}, dirB(fault), dirC{200B}`: assert `status==ERROR_PARTIAL_COPY` **and** total includes dirC's 200B (siblings still visited). RED today. |
| CW-T6 | CW-6 | `CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp` (mirror the symlink-loop-guard cases): choose the test assertion to match the CW-6 design decision. For **coverage-only**, use a junction to a target reachable only through the alias path and assert the unique `needle_alias.txt` is indexed; do not require duplicate alias-path emission. For **alias-path results**, use a target also reachable canonically and assert both result paths or the selected alias path is emitted. Skip on `ERROR_PRIVILEGE_NOT_HELD`. RED today for the coverage-loss case. |
| CW-T7 | CW-7 | **DONE 2026-07-04.** `DxUiTests.Controls.cpp`: TabControl >=3 closable tabs; `OnMouseDown` at tab0 body -> `OnMouseMove` past the drag threshold -> `OnMouseUp` at tab2's close-button center; asserts tab count unchanged and no close callbacks fired. RED before the fix (`FAILED: TabControl body-started drag release over a close button is not a close action`), GREEN after the guarded-close fix via `.\.build\x64\Debug\DxUiTests.exe --suite=Control`. |
| CW-T8 | CW-8 | `DxUiTests.Animation.cpp`: attach to an iconic HWND, `OnAnimationTick` latches suspend (subscription 0); dispatch `WM_SIZE/SIZE_RESTORED` with the window visible; assert the flag clears and a new subscription is requested. RED today. |
| CW-T9 | CW-9 | `DxUiTests.Theme.cpp`: divergent `ViewerTheme` `darkMode=TRUE, darkBase=FALSE`, mid-luminance accent; assert `Luminance(accentPressed) < Luminance(accent)` (light-base contract). RED today (brightens). |
| CW-T10 | CW-10 | **DONE 2026-07-05.** Added `search_service_transient_authorization_failure_is_incomplete_not_cached`: RED forced a transient directory authorization `CreateFileW` failure with an exclusive directory handle and failed because `warningFlags` was `0x00000000`; GREEN completes the first query with `ACCESS_DENIED_SKIPPED`, releases the handle, and proves the same service session re-evaluates the directory instead of caching `false`. Existing `search_service_filters_cached_descendants_denied_to_client` remains the durable `ACCESS_DENIED` contrast and passed in the adjacent sweep. |
| CW-T11 | CW-11 | Extend `ShouldFailStagedCopyPromoteForSelfTest` to honor an injected HRESULT; symlink overwrite copy with injected staged-promote `ERROR_INVALID_PARAMETER`; assert `CopyItem` returns `ERROR_INVALID_PARAMETER` (not `ERROR_NOT_SUPPORTED`). RED today. |
| CW-T12 | CW-12 | Sibling of `sqlite_index_store_automatic_checkpoint`: hold a second connection's lock so `CHECKPOINT_TRUNCATE`→`SQLITE_BUSY`; `RunAutomaticMaintenance` returns `S_FALSE` and `lastCheckpointUtc` stays empty. RED today. |
| CW-T13 | CW-13 | Debug hook forcing the final `SetFileAttributesW` in `PromoteStagedTempIntoFinalPath` to fail while attributes genuinely differ (`stripTemporaryAttributes=true`); assert `FAILED(hr)`. RED today (`S_OK`). |
| CW-T14 | CW-14 | Direct `EnumerateVolume` call whose callback returns `kSkipCandidateHr` for the first candidate then emits a real row; assert `emittedRows==1` (not 2) and the injected gate doesn't fire on the skip. RED today. |
| CW-T15 | CW-15 | `DxUiTests`: TextField text = `"A"` + one astral codepoint, `SetMasked(true)`, `PasswordMaskLengthPolicy::Exact`; assert `GetSecretVisibleDotCount()==2`. RED today (returns 3). |
| CW-T-P1 | CW-P1 — **DONE 2026-07-13 via Causeway** | Causeway removed the cleanup source-reader dependency entirely, retained the destination-corruption regression, and archived the candidate FileOps evidence linked from its Done plan. |
| CW-T-P2 | CW-P2 | `ENABLE_TESTS` accessor for `copiedEntries.size()`; multi-file/dir cross-FS move with the existing pre-cleanup pause hook; assert the manifest size strictly decreases during cleanup and reaches 0. RED today (pinned at full count). Also emit/reuse a `FileOps.*` manifest-size/retained-entry metric and archive baseline/candidate runs under `Specs/TestRuns/`. |

## Refuted / out of scope (do NOT re-flag)

- **DxUi gap sweep came back clean.** A dedicated second-pass reviewer over the accessibility, controls,
  window-host and text-input diffs found its candidate defects (inverted selection rect, MenuBar hit-rect
  mismatch, TabControl bounds re-layout, surrogate caret math) were **all explicitly guarded** in the
  source. The DxUi findings that survived are only CW-7/CW-8/CW-9/CW-15/CW-S1/CW-S2/CW-S3 above.
- **CW-P1 review premise was wrong in part:** there is no copy-time content hash in the bridge today, and
  size is already short-circuited; the fix must *add* a hash, not reuse one. Severity lowered HIGH→LOW-MED
  because master already verifies-before-delete and fails **safe** (skips the delete, reports PARTIAL)
  rather than deleting unverified data.
- **CW-S1 is not a build break.** The tree has no `TreatWarningAsError`/`<WX>`; C4505 is warning hygiene
  against the project's stated "0 warnings" bar, not a hard failure.
- **CW-S3 is latent.** No current runtime path recreates text formats while the host + TabControl persist;
  the fix is forward-looking hardening, not a live bug.
- **CW-S2 is not yet a confirmed bug.** The code does grow vectors inside `noexcept` layout-cache builders,
  but repo policy explicitly treats `std::bad_alloc` as fatal. Keep it as an investigation or policy
  clarification, not a mechanical remove-`noexcept` cleanup.

## Suggested sequencing

1. **P0 — auth/data:** land CW-1 on `master` + add CW-T1; fix **CW-2** (forced re-login is the highest
   user-visible regression with a security flavor).
2. **P1a — FileOps/data movement:** CW-P2, CW-P1, CW-13, CW-11. These share staged/cross-FS verification
   and must include FileOps perf evidence where applicable.
3. **P1b — provider/search availability:** CW-3, CW-4, CW-5, CW-10, then CW-6 only after the alias-path
   contract decision above.
4. **P1c — DxUi user-visible regressions:** CW-8, CW-9, CW-15 (CW-15 pending the author's intent call). CW-7 is done.
5. **P2 — cleanup/hardening:** CW-S1, CW-S2 investigation, CW-S3.

The recurring shape across the real findings is **over-strict hardening turning a recoverable condition into
a hard failure** (a metadata drift aborts a MOVE, a denied HEAD kills a read, one bad archive entry sinks the
archive, a transient error truncates a walk or poisons a cache) — bias the fixes toward *degrade-and-flag*
over *abort*. For each batch, close with a focused selftest run, a full green check when practical
(`.\Tools\Run-AllTests.ps1 -Suite Full`), archived `Specs/TestRuns/` evidence for perf-sensitive work, and
the spec closeout required by `AGENTS.md:57-58`.
