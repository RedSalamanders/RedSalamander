# Core Search Plugin Architecture Plan

Last updated: 2026-03-10

Status: WIP

References:
- `Specs/Core/Core_Search.md`
- `Specs/Testing/Testing_SelfTests.md`
- `Specs/Plans/WIP/RFC_Core_InstantFileSearchUsnMftIndex.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Common/PlugInterfaces/FileSystem.h`
- https://github.com/dfs-minded/indexer-plus-plus as inspiration but not current implementation

## Living Implementation Checklist

This checklist is the execution control list for the feature. It is intentionally a living section:
- mark completed work with `[x]`
- keep pending or active work as `[ ]`
- mark blockers as `[blocked]` and capture the missing dependency or decision
- when implementation reveals new required work, add it immediately under the relevant phase before continuing
- do not mark a phase complete until its validation items have passed

### Phase 0 - Contract and design freeze

- [x] Re-read `RFC_Core_InstantFileSearchUsnMftIndex.md` and reconcile it with this plan, especially the NTFS, ReFS, and unsupported-root boundaries.
- [x] Inventory all host and plugin references to `IFileSystemSearch`, `FileSystemSearchQuery`, and related callback types, and confirm the current payload layout is not relied on by shipped code.
- [x] Update `Common/PlugInterfaces/FileSystem.h` to widen the search enums and structs in place and add comments for all new fields.
- [x] Update `Specs/Plugins/Plugins_VirtualFileSystem.md` so the documentation, ABI contract, and capability JSON all describe the same behavior.
- [x] Freeze the `sizeBytes` compatibility rules for query, match, and progress payloads.
- [x] Freeze the canonical result identity format used by append/intersect/subtract operations.
- [x] Freeze default limits for `maxResults`, `maxContentBytesPerFile`, and `maxSnippetCharacters`.
- [x] Record explicit fallback behavior for plugins that do not expose `IFileSystemSearch` and plugins that expose it but cannot satisfy index or content requests.
- [x] Normalize callback-returned `E_ABORT` / `HRESULT_FROM_WIN32(ERROR_CANCELLED)` to `HRESULT_FROM_WIN32(ERROR_CANCELLED)` in the built-in file plugin search path.
- [x] Clarify that search query, match, and progress strings follow call/callback lifetime rules and are not `FileSystemArena` inputs.
- [x] Validation: header text, plugin spec text, and this plan say the same thing about fields, defaults, and fallback behavior.
- [x] Validation: ABI review confirms the unchanged COM method shape is sufficient and `IFileSystemSearch2` is unnecessary.

#### Phase 0 frozen decisions

- `sizeBytes` is an exact-match contract for `FileSystemSearchQuery`, `FileSystemSearchMatch`, and `FileSystemSearchProgress`; mismatches are rejected with `E_INVALIDARG`.
- Result-set identity is frozen as `pluginId + normalized instanceContext + normalized fullPath`.
- Default limits are frozen as `maxResults == 0` meaning unlimited, `maxContentBytesPerFile == 0` meaning plugin default `64 MiB`, and `maxSnippetCharacters == 0` meaning plugin default `160`.
- Search query strings are caller-owned and call-scoped; search match/progress strings are plugin-owned and callback-scoped; search payload strings are not `FileSystemArena`-backed.
- Fallback behavior is frozen as: native plugin search first when available, host scan fallback otherwise, and warning-flag degradation when a plugin cannot honor index or content preferences.
- Repository inventory on 2026-03-10 found search ABI usage only in the public header, the built-in file plugin, self-tests, and documentation; no shipped host UI path currently depends on the old payload layout.

#### Phase 0 exit evidence

- `Common/PlugInterfaces/FileSystem.h`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, and this plan were reconciled on 2026-03-10 around exact `sizeBytes` validation, fallback rules, default limits, and search string lifetime rules.
- Verification build passed for `Plugins/FileSystem/FileSystem.vcxproj` and `RedSalamander/RedSalamander.vcxproj` in `Debug|x64`.
- `RedSalamander.exe --compare-selftest --selftest-fail-fast` passed on 2026-03-10 and covered `local_search_qi_and_capabilities`, `local_search_callback_contract`, `local_search_name_wildcard_recursive`, `local_search_content_literal`, `local_search_name_and_content_and_semantics`, and `local_search_invalid_query_rejected`.

### Phase 1 - Shared host search foundation

- [x] Extract reusable text-search helpers from the text viewer into a shared helper module.
- [x] Compile the shared helper module into both `RedSalamander` and the built-in `FileSystem` plugin so fallback and native scan backends use the same decoding and content-match behavior.
- [x] Implement BOM detection, UTF-8 validation, ANSI fallback, binary-file detection, and chunk-overlap matching as reusable primitives.
- [x] Build a generic recursive scan engine that uses `IFileSystem::ReadDirectoryInfo` for traversal and `IFileSystemIO::CreateFileReader` for content scanning.
- [x] Add cancellation polling at directory traversal boundaries and during chunked content reads.
- [x] Add directory loop protection for reparse points or plugin-specific link behavior.
- [x] Add progress reporting for directory count, file count, candidate count, and match count.
- [x] Define how the scan engine reports degraded capabilities when content search is requested but no readable stream API exists.
- [x] Validation: unit tests cover decoding heuristics, chunk boundaries, regex/literal matching, and binary skip rules.
- [x] Validation: integration tests confirm the scan engine works for local, dummy, archive, and remote-style plugins where supported.
- [x] Validation: cancellation remains responsive and the engine never blocks the UI thread.

#### Phase 1 exit evidence

- Shared helper extraction landed in `Common/SearchTextHelpers.h` and `Common/SearchTextHelpers.cpp`, and the built-in file plugin search path now consumes those helpers instead of maintaining a second private text-decoding implementation.
- Host fallback execution landed in `RedSalamander/SearchFallbackEngine.h` and `RedSalamander/SearchFallbackEngine.cpp` with scan traversal, degraded capability reporting, callback-driven cancellation, and chunked content search.
- Verification build passed on 2026-03-10 for `.\build.ps1 -ProjectName RedSalamander`, which rebuilt both `RedSalamander.exe` and `Plugins\FileSystem.dll`. A direct `MSBuild.exe Plugins\FileSystem\FileSystem.vcxproj /p:Configuration=Debug,Platform=x64` verification pass also succeeded because the current `build.ps1 -ProjectName FileSystem` path still misroutes to an MSBuild target name instead of a project selection.
- `RedSalamander.exe --compare-selftest --selftest-fail-fast` passed on 2026-03-10 and the trace recorded `search_text_helpers_decoding_and_binary`, `search_text_helpers_chunk_overlap_literal_and_regex`, `host_fallback_search_local_plugin_path_root`, `host_fallback_search_content_degraded_without_io`, `host_fallback_search_short_read_and_cancel`, `host_fallback_search_dummy_name_only`, `host_fallback_search_7z_name_only`, and `host_fallback_search_remote_ftp_name_only`.
- The remote FTP fallback smoke case is intentionally gateable: it runs when the configured remote self-test profile and sandbox checks pass, and otherwise self-records a skipped result without failing the compare suite.
- The shared fallback engine remains a synchronous utility with no HWND, message-pump, or UI-thread dependency; Phase 2 host integration must continue to dispatch it on a worker thread, and the Phase 1 cancellation self-test covers responsive cancellation during chunked reads.

### Phase 2 - Find dialog and session orchestration

- [x] Audit existing themed host controls and identify which search surfaces can reuse them directly versus which require extension or new controls.
- [x] Add a dedicated modeless Find dialog instead of reusing `FolderView` as the primary results surface.
- [x] Implement the Find UI on the existing host DirectX-rendered, theme-aware custom control stack rather than introducing an unrelated stock-dialog UI path.
- [x] Implement `SearchSessionController` to own query execution, worker-thread lifetime, cancellation, result batching, and final summary state.
- [x] Implement backend selection rules: native plugin search first, host fallback second, with explicit degradation reporting.
- [x] Implement result-set algebra for `Find`, `Append`, `Intersect`, and `Subtract` using the canonical result identity.
- [x] Add result actions for open, navigate to directory, and focus-in-parent behavior.
- [x] Extend existing controls or add new themed controls where search needs behavior not already present, especially for the result list, status/progress surface, and high-density filter/options layout.
- [x] Persist dialog placement, recent roots, recent patterns, and last-used options.
- [x] Preserve editable combo text when refreshing search history lists so repeated searches do not silently lose criteria.
- [x] Make debug/self-test idle waits observe queued result/completion delivery instead of only worker-thread exit.
- [x] Validation: `cmd/pane/find` and `Alt+F7` open the dialog and no longer fall through to not-implemented behavior.
- [x] Validation: start, cancel, restart, and repeated searches behave correctly without stale worker state.
- [x] Validation: append/intersect/subtract produce correct results across repeated searches and mixed backends.
- [x] Validation: dialog updates are incremental and remain responsive on large result sets.
- [x] Validation: the dialog follows the active theme, responds to runtime theme changes, and preserves expected keyboard/focus behavior on the custom control stack.

#### Phase 2 exit evidence

- The host dialog/session slice landed in `RedSalamander/FindFilesWindow.h`, `RedSalamander/FindFilesWindow.cpp`, `RedSalamander/RedSalamander.cpp`, `RedSalamander/RedSalamander.rc`, `RedSalamander/resource.h`, `Common/WindowMessages.h`, `Common/SettingsStore.h`, `Common/Common/SettingsStore.cpp`, `RedSalamander/SettingsSave.h`, `RedSalamander/FolderWindow.h`, and `RedSalamander/FolderWindow.FileSystem.cpp`.
- The dialog is host-owned and modeless, uses themed host controls and owner-drawn actions, streams batched results on posted messages, persists placement/history/options, and routes open/parent navigation back through the active pane execution flow.
- Phase 2 implementation surfaced two required follow-up tasks that were completed in the same phase: preserving editable combo text across history refreshes, and making debug idle waits account for queued completion/results so repeated-search tests observe settled UI state.
- Verification build passed on 2026-03-10 for `.\build.ps1 -ProjectName RedSalamander`.
- Commands self-tests passed individually on 2026-03-10 with `Start-Process -Wait`: `settings_store_search_roundtrip`, `modeless_window_ownership`, and `cmd_pane_find_dialog_search_ops`.
- `cmd_pane_find_dialog_search_ops` now verifies default `Alt+F7` mapping for `cmd/pane/find`, shortcut-based dialog opening, repeated `Find`/`Append`/`Intersect`/`Subtract` execution, content search, large-result cancellation, and settled post-search state.

### Phase 3 - Built-in file plugin native search adoption

- [x] Implement the updated `IFileSystemSearch` in the built-in local file plugin.
- [x] Start with scan-backed native search so the plugin contract is exercised before index work begins.
- [x] Add backend selection policy and configuration override values (`auto`, `service`, `local-index`, `scan`).
- [x] Normalize local paths and metadata collection so scan, index, and service backends return compatible results.
- [x] Keep host fallback relative-path separator style aligned with the queried root so native and fallback result payloads stay parity-compatible for local plugins.
- [x] Ensure progress callbacks identify the active backend and surface degradation warnings consistently for the current scan backend.
- [x] Validation: plugin-native scan results match host scan fallback for the same root and query.
- [x] Validation: capability JSON accurately reflects what the plugin can do before indexed backends are enabled.
- [x] Validation: unsupported or unavailable indexed backend paths fall back cleanly to scan behavior without user-visible breakage.

#### Phase 3 exit evidence

- The built-in file plugin now snapshots `searchBackendPreference` from configuration, selects the active backend through a dedicated policy helper, reports that backend through progress callbacks, and degrades `service`, `local-index`, and `PREFER_INDEX` requests back to scan with `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX` until indexed backends exist.
- Local scan result normalization now uses shared path-normalization helpers in `Plugins/FileSystem/FileSystem.Search.cpp`, and single-file metadata population was consolidated so root-file searches and recursive directory searches return compatible payload fields.
- Phase 3 implementation discovered one additional compatibility gap in the host fallback engine: nested relative paths could switch from `\` to `/` for backslash-rooted plugins. That was fixed in `RedSalamander/SearchFallbackEngine.cpp` during the same phase so native and host-fallback search results compare cleanly.
- Verification build passed on 2026-03-10 for `.\build.ps1 -ProjectName RedSalamander`.
- Focused compare self-tests passed on 2026-03-10 with `Start-Process -Wait`: `local_search_qi_and_capabilities`, `local_search_backend_preferences_roundtrip`, `local_search_native_matches_host_fallback`, `local_search_name_and_content_and_semantics`, and `host_fallback_search_local_plugin_path_root`.
- `local_search_backend_preferences_roundtrip` now verifies schema exposure, configuration round-trip, `SomethingToSave` behavior, truthful capability JSON, and scan fallback when `service`, `local-index`, or `PREFER_INDEX` is requested against the current scan-only backend set.
- `local_search_native_matches_host_fallback` now verifies native built-in search and host fallback produce matching final result payloads and completed progress summaries for the same local root and query.

### Phase 4 - Shared local index core

- [x] Create a shared local index core library used by both the built-in file plugin and the Windows service.
- [x] Implement snapshot persistence, path normalization, and query planning in the shared core.
- [x] Implement NTFS initial seeding via `FSCTL_ENUM_USN_DATA`.
- [x] Implement NTFS live updates via `FSCTL_READ_USN_JOURNAL`.
- [x] Implement ReFS initial traversal via `NtQueryDirectoryFile`.
- [x] Implement ReFS live updates via `FSCTL_READ_USN_JOURNAL`.
- [x] Add a user-mode traversal-seed fallback when USN/MFT or journal access is unavailable to the non-service process so indexed search still works before Phase 5.
- [x] Define rebuild behavior for journal ID mismatch, journal wrap, snapshot corruption, and removable-volume churn.
- [x] Keep v1 content search out of the index core; the core returns candidates and metadata only.
- [x] Validation: indexed name search and scan fallback produce equivalent final result sets for supported local roots.
- [x] Validation: snapshot reload, restart, journal replay, rename storms, and delete/create bursts preserve consistency.
- [x] Validation: warm indexed search returns the first batch within the target budget and cancellation still works.

#### Phase 4 exit evidence

- Shared index-core code landed in `Common/LocalSearchIndexCore.h` and `Common/LocalSearchIndexCore.cpp`, and both `Plugins/FileSystem/FileSystem.vcxproj` and `RedSalamander/RedSalamander.vcxproj` now compile it directly.
- The built-in local file plugin now owns a shared `LocalSearchIndexCore::Repository`, advertises `indexed=true` / `preferredBackend=local-index`, and routes `auto` / `local-index` searches through the indexed backend in `Plugins/FileSystem/FileSystem.Search.cpp`.
- NTFS seeding and live-update code paths now exist in the shared core via `FSCTL_ENUM_USN_DATA` and `FSCTL_READ_USN_JOURNAL`; ReFS initial traversal and journal-replay code paths now exist via `NtQueryDirectoryFile` plus the same journal replay machinery.
- Phase 4 implementation discovered an important user-mode constraint on this machine: `FSCTL_ENUM_USN_DATA` and journal access are not always available to the regular app process even when the volume supports USN. The shared core therefore now falls back to tracked-root traversal seeding and traversal-based rebuilds when those privileged paths are unavailable, while keeping the real USN/MFT code paths for service/elevated scenarios in Phase 5.
- Indexed native search now preserves root-relative `relativePath` payloads so native built-in results stay parity-compatible with the host fallback engine for nested matches.
- Verification build passed on 2026-03-10 for `.\build.ps1 -ProjectName RedSalamander`.
- Focused compare self-tests passed on 2026-03-10 with `Start-Process -Wait`: `local_search_qi_and_capabilities`, `local_search_backend_preferences_roundtrip`, `local_search_callback_contract`, `local_index_core_snapshot_reload`, `local_index_core_journal_replay_rename_delete_create`, `local_index_core_snapshot_corruption_rebuild`, and `local_search_native_matches_host_fallback`.
- `local_index_core_snapshot_reload` now verifies supported-root probing, UNC rejection, cold snapshot creation, warm snapshot reload, and warm-query budget on the shared core.
- `local_index_core_journal_replay_rename_delete_create` now verifies that indexed results stay correct across rename/delete/create changes using journal replay when available and traversal rebuild fallback when it is not.
- `local_index_core_snapshot_corruption_rebuild` now verifies rebuild behavior after invalid snapshot magic and after invalid `NextUsn` state.

### Phase 5 - Windows service backend

- [x] Implement `RedSalamanderSearchService` as the elevated preferred backend for supported local roots.
- [x] Define a versioned named-pipe protocol for query, progress, result batches, status, and rebuild control messages.
- [x] Apply named-pipe ACLs that allow local interactive users while preventing broader access.
- [x] Persist service-owned index data under `%ProgramData%\\RedSalamander\\SearchIndex\\`.
- [x] Implement service restart, corruption recovery, and client reconnection behavior.
- [x] Keep file-content scanning on the client side in v1; the service must not return raw content.
- [x] Add installer integration for service install, uninstall, upgrade, and opt-out/failure behavior.
- [x] Add a Release-build regression guard for `FindFilesWindow` debug-only helpers so packaging validation does not fail on `_DEBUG` type leakage from Phase 2.
- [x] Exclude runtime `*.WebView2` state directories from MSIX payload harvesting so package generation only ships intended binaries and assets.
- [x] Validation: service unavailability falls back automatically to in-process index or scan behavior.
- [x] Validation: IPC version mismatch, service restart during query, and multi-client use all fail safely.
- [x] Validation: security review confirms the privilege model, ACLs, and storage location are acceptable.

#### Phase 5 exit evidence

- Service backend code landed in `Common/SearchServiceBroker.h`, `Common/SearchServiceBroker.cpp`, `RedSalamanderSearchService/Main.cpp`, and `RedSalamanderSearchService/RedSalamanderSearchService.vcxproj`, with the solution and dependent project files updated so the service builds as part of the normal solution.
- The shared local index core now supports redirectable snapshot storage through repository options, and the service uses `%ProgramData%\\RedSalamander\\SearchIndex\\` while the user-mode app/plugin keep their existing per-user defaults.
- The named-pipe protocol is versioned at `1`, covers status/query/progress/result-batch/complete/rebuild flows, and applies a restricted pipe security descriptor for `SYSTEM`, built-in administrators, and local interactive users together with `PIPE_REJECT_REMOTE_CLIENTS`.
- The built-in local file plugin now prefers `service -> local-index -> scan`, advertises `serviceBacked=true` / `preferredBackend=service`, and safely falls back when the service is absent, mismatched, disconnected, or otherwise unavailable.
- V1 content search remains client-side: the service returns candidate paths and metadata only, while the client/plugin reuses the existing streaming content scan path for literal/regex text matching and snippets.
- Installer integration landed in `Installer/msi/build-msi.ps1`, `Installer/msi/Product.wxs`, and `Installer/msix/RedSalamanderInstaller.wapproj`. MSI now stages and installs `RedSalamanderSearchService.exe` as a Windows service with opt-out via `INSTALLSEARCHSERVICE=0`, and MSIX now ships the service executable as package content without attempting service registration.
- Phase 5 validation surfaced two packaging regressions that were fixed in the same phase: `RedSalamander/FindFilesWindow.cpp` now keeps debug-only helper declarations and definitions behind `_DEBUG`, and the MSIX content harvest now excludes `*.WebView2` runtime cache trees that are not shippable artifacts.
- Verification build passed on 2026-03-10 for `.\build.ps1 -Configuration Release`.
- Packaging validation passed on 2026-03-10 for `.\Installer\msi\build-msi.ps1 -Configuration Release -Platform x64`, producing `.build\\AppPackages\\RedSalamander-7.0.183-x64.msi`.
- Packaging validation passed on 2026-03-10 for `MSBuild.exe Installer\\msix\\RedSalamanderInstaller.wapproj /t:Build /p:Configuration=Release;Platform=x64`, producing `.build\\AppPackages\\RedSalamanderInstaller_7.0.0.183_x64_Test\\RedSalamanderInstaller_7.0.0.183_x64.msix`.
- Focused compare self-tests passed on 2026-03-10 with `Start-Process -Wait`: `local_search_qi_and_capabilities`, `local_search_backend_preferences_roundtrip`, `local_search_native_matches_host_fallback`, `search_service_status_and_query_roundtrip`, `local_search_service_matches_host_fallback`, `local_search_service_protocol_mismatch_falls_back_local_index`, `local_search_service_disconnect_falls_back_local_index`, and `search_service_multi_client_and_rebuild_control`.
- `search_service_status_and_query_roundtrip` now verifies the direct broker handshake, protocol version, service-reported pipe name, `%ProgramData%` storage root, and candidate query flow against a foreground test instance.
- `local_search_service_matches_host_fallback`, `local_search_service_protocol_mismatch_falls_back_local_index`, and `local_search_service_disconnect_falls_back_local_index` together verify the most important safety contracts: service-backed local search parity, mismatch fallback, and mid-query disconnect fallback.
- `search_service_multi_client_and_rebuild_control` now verifies concurrent clients and explicit rebuild control against the same foreground service instance.

### Phase 6 - Hardening, parity, and rollout

- [x] Run the full fallback, indexed, service, content, and performance test matrix.
- [x] Add instrumentation or debug logging needed to diagnose backend selection, rebuilds, and degraded execution paths.
- [x] Extend shared query statistics with snapshot bytes, estimated memory, and ready/query timing data and validate broker transport of those fields.
- [x] Add a host-fallback regression for `ACCESS_DENIED` traversal handling and completed-warning propagation.
- [x] Add a native-vs-fallback regression for Unicode and deep/long local paths.
- [x] Review memory usage, snapshot size, and cold-start behavior on current NTFS and ReFS developer volumes and capture the evidence in self-tests and diagnostics.
- [x] Review accessibility, keyboard flow, and localization readiness of the Find dialog.
- [x] Update user-facing help or command descriptions if Find behavior changes visibly.
- [x] Re-open this checklist during implementation and append newly discovered work instead of keeping it in ad hoc notes.
- [x] Validation: all exit criteria from earlier phases are green or explicitly marked `[blocked]` with an owner and follow-up plan.
- [x] Validation: acceptance evidence covers NTFS, ReFS, unsupported roots, service down, access denied, Unicode paths, long paths, and cancel-on-large-search scenarios.
- [x] Validation: rollout decision is based on parity and stability evidence, not just feature completeness.

#### Phase 6 exit evidence

- Phase 6 diagnostics landed in `Common/LocalSearchIndexCore.cpp`, `Common/SearchServiceBroker.cpp`, `Plugins/FileSystem/FileSystem.Search.cpp`, and `RedSalamander/SearchFallbackEngine.cpp`. The search stack now logs backend selection, rebuild paths, degraded execution, snapshot bytes, estimated memory, and ready/query timings.
- `LocalSearchIndexCore::QueryStats` now carries `snapshotFileBytes`, `estimatedMemoryBytes`, `ensureReadyDurationMs`, and `executeQueryDurationMs`, and the service broker transports those fields end-to-end.
- The compare suite gained `local_index_core_refs_probe_and_query_if_available`, `local_search_native_unicode_long_path_matches_host_fallback`, and `host_fallback_search_access_denied_warning`, while `local_index_core_snapshot_reload` and `search_service_status_and_query_roundtrip` were tightened to assert the new telemetry fields. The ReFS probe always participates in the suite and self-records `skipped` with a reason when no fixed ReFS volume is available.
- Current-machine volume inventory on 2026-03-10 confirmed both NTFS and ReFS coverage is available for targeted validation on this machine: `C:` is `NTFS` and `Z:` is `ReFS`.
- Verification build passed on 2026-03-10 for `.\build.ps1 -ProjectName RedSalamander`.
- Full compare self-tests passed on 2026-03-10 with `.\.build\x64\Debug\RedSalamander.exe --compare-selftest --selftest-fail-fast`, with the ReFS-specific case cleanly skipped on machines that do not expose a fixed ReFS volume.
- Focused compare self-tests also passed on 2026-03-10 for `local_index_core_snapshot_reload`, `local_search_native_unicode_long_path_matches_host_fallback`, `search_service_status_and_query_roundtrip`, `host_fallback_search_access_denied_warning`, and `local_index_core_refs_probe_and_query_if_available` on a machine that exposed `Z:` as `ReFS`.
- Commands self-tests passed on 2026-03-10 for `settings_store_search_roundtrip`, `modeless_window_ownership`, and `cmd_pane_find_dialog_search_ops`, and `RedSalamander/RedSalamander.rc` now describes `cmd/pane/find` as opening the Find Files and Directories dialog.
- Release verification build passed on 2026-03-10 for `.\build.ps1 -Configuration Release`.
- The Phase 6 rollout recommendation is positive for the current backend order (`service -> local-index -> scan`) because parity, cancellation, degraded-mode signaling, and ReFS/NTFS probe coverage are now backed by passing tests and runtime diagnostics rather than inferred behavior.

## Summary

This document defines the implementation plan for RedSalamander search across the existing plugin architecture.

Key decisions:
- Reuse and evolve `IFileSystemSearch` in place.
- Keep the callback-based synchronous `Search(...)` method.
- Widen the query, match, and progress payloads using `sizeBytes`.
- Make `cmd/pane/find` a real modeless Find dialog with streaming results.
- Support three execution paths:
  - generic scan fallback for any file system plugin,
  - built-in local indexed search for `builtin/file-system`,
  - Windows service-backed local indexed search as the preferred elevated backend.

V1 includes:
- name/path search,
- content search for text files,
- append/intersect/subtract result-set operations,
- cancellation and progress,
- plugin fallback behavior when native search is unavailable.

V1 does not include:
- binary or hex content search in the Find dialog,
- a synthetic `search:` virtual file system,
- duplicate-file search,
- saved named searches.

## Goals

- Provide a single search UX for all file system plugins.
- Make search capability discoverable through the existing plugin contracts.
- Allow plugins to opt into fast native search without breaking generic correctness.
- Keep the host responsive by running all searches off the UI thread.
- Prefer instant indexed local search when available, but always preserve a correct fallback path.

## Non-Goals

- Introduce a separate search-plugin type.
- Replace `FolderView` with a search-specific virtual folder in v1.
- Perform privileged content indexing inside the Windows service in v1.
- Support FAT/exFAT/UNC/WSL with MFT- or journal-based instant indexing.

## User Experience

`cmd/pane/find` opens a modeless host-owned `Find Files and Directories` dialog scoped to the focused pane.

The UI is not intended to be a stock Win32 common dialog. It should be implemented on the existing RedSalamander DirectX-rendered, theme-aware custom control framework so it looks, behaves, and repaints like the rest of the application. Existing custom controls should be reused where they already fit, and the plan explicitly allows improving them or adding new themed controls where search-specific behavior requires it.

Default scope:
- current plugin,
- current instance context,
- current pane path as root.

V1 dialog behavior:
- Enter starts a new `Find`.
- Results stream into a report-style list as matches arrive.
- `Cancel` stops the current search.
- `Append` unions the new result set with the current result set.
- `Intersect` keeps only items present in both sets.
- `Subtract` removes newly found items from the current set.
- Double-click on a result:
  - file: open using the existing open/view flow,
  - directory: navigate the pane to that directory,
  - file secondary action: navigate to parent and focus the item.

Result columns:
- Name
- Path
- Size
- Modified
- Attr
- Snippet (shown only when content search is active and a snippet is available)

Status area:
- active backend,
- progress counts,
- capability degradation warnings,
- final summary.

## Public Interface

The COM method surface remains:

```cpp
interface __declspec(uuid("00417f3e-f0f5-4add-8dea-4407d5169ef6"))
         __declspec(novtable) IFileSystemSearch : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Search(
        const FileSystemSearchQuery* query,
        IFileSystemSearchCallback* callback,
        void* cookie) noexcept = 0;
};
```

`IID_IFileSystemSearch` remains unchanged because the method surface does not change.

The data contract is widened in place.

### Updated search types

```cpp
enum FileSystemSearchFlags : uint32_t
{
    FILESYSTEM_SEARCH_NONE                = 0,
    FILESYSTEM_SEARCH_RECURSIVE           = 0x1,
    FILESYSTEM_SEARCH_INCLUDE_FILES       = 0x2,
    FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES = 0x4,
    FILESYSTEM_SEARCH_FOLLOW_SYMLINKS     = 0x8,
    FILESYSTEM_SEARCH_MATCH_CASE_NAME     = 0x10,
    FILESYSTEM_SEARCH_MATCH_CASE_CONTENT  = 0x20,
    FILESYSTEM_SEARCH_WANT_SNIPPETS       = 0x40,
    FILESYSTEM_SEARCH_PREFER_INDEX        = 0x80,
};

enum FileSystemSearchNameMode : uint32_t
{
    FILESYSTEM_SEARCH_NAME_DISABLED = 0,
    FILESYSTEM_SEARCH_NAME_WILDCARD = 1,
    FILESYSTEM_SEARCH_NAME_LITERAL  = 2,
    FILESYSTEM_SEARCH_NAME_REGEX    = 3,
};

enum FileSystemSearchContentMode : uint32_t
{
    FILESYSTEM_SEARCH_CONTENT_DISABLED     = 0,
    FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL = 1,
    FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX   = 2,
};

enum FileSystemSearchBackend : uint32_t
{
    FILESYSTEM_SEARCH_BACKEND_UNKNOWN = 0,
    FILESYSTEM_SEARCH_BACKEND_SCAN    = 1,
    FILESYSTEM_SEARCH_BACKEND_INDEX   = 2,
    FILESYSTEM_SEARCH_BACKEND_SERVICE = 3,
};

enum FileSystemSearchMatchSource : uint32_t
{
    FILESYSTEM_SEARCH_MATCH_SOURCE_NONE    = 0,
    FILESYSTEM_SEARCH_MATCH_SOURCE_NAME    = 0x1,
    FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT = 0x2,
};

enum FileSystemSearchWarningFlags : uint32_t
{
    FILESYSTEM_SEARCH_WARNING_NONE                   = 0,
    FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX     = 0x1,
    FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT   = 0x2,
    FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED = 0x4,
    FILESYSTEM_SEARCH_WARNING_OVERFLOW              = 0x8,
};

enum FileSystemSearchPhase : uint32_t
{
    FILESYSTEM_SEARCH_PHASE_INITIALIZING = 0,
    FILESYSTEM_SEARCH_PHASE_ENUMERATING  = 1,
    FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP = 2,
    FILESYSTEM_SEARCH_PHASE_CONTENT_SCAN = 3,
    FILESYSTEM_SEARCH_PHASE_COMPLETED    = 4,
};

struct FileSystemSearchQuery
{
    uint32_t sizeBytes;
    const wchar_t* rootPath;
    const wchar_t* namePattern;
    const wchar_t* contentPattern;
    FileSystemSearchFlags flags;
    FileSystemSearchNameMode nameMode;
    FileSystemSearchContentMode contentMode;
    uint64_t maxResults;
    uint64_t maxContentBytesPerFile;
    uint32_t maxSnippetCharacters;
    uint32_t reserved;
};

struct FileSystemSearchMatch
{
    uint32_t sizeBytes;
    const wchar_t* fullPath;
    unsigned long fullPathSize;
    const wchar_t* relativePath;
    unsigned long relativePathSize;
    const wchar_t* displayName;
    unsigned long displayNameSize;
    const wchar_t* previewText;
    unsigned long previewTextSize;
    unsigned long fileAttributes;
    __int64 creationTime;
    __int64 lastAccessTime;
    __int64 lastWriteTime;
    __int64 changeTime;
    __int64 endOfFile;
    __int64 allocationSize;
    uint32_t matchedBy;
    uint64_t contentMatchByteOffset;
    uint32_t contentMatchByteLength;
    uint32_t reserved;
};

struct FileSystemSearchProgress
{
    uint32_t sizeBytes;
    FileSystemSearchPhase phase;
    FileSystemSearchBackend backend;
    uint32_t warningFlags;
    HRESULT statusHint;
    uint64_t scannedDirectories;
    uint64_t scannedFiles;
    uint64_t candidateFiles;
    uint64_t matchedEntries;
    const wchar_t* currentPath;
    unsigned long currentPathSize;
};

interface __declspec(novtable) IFileSystemSearchCallback
{
    virtual HRESULT STDMETHODCALLTYPE FileSystemSearchMatch(
        const FileSystemSearchMatch* match,
        void* cookie) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE FileSystemSearchProgress(
        const FileSystemSearchProgress* progress,
        void* cookie) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE FileSystemSearchShouldCancel(
        BOOL* pCancel,
        void* cookie) noexcept = 0;
};
```

### Contract rules

- The caller must set `sizeBytes` on every in/out search struct.
- The callee must validate `sizeBytes` before reading optional trailing fields.
- At least one of `nameMode` or `contentMode` must be enabled.
- If both `nameMode` and `contentMode` are enabled, a match must satisfy both.
- `WANT_SNIPPETS` is a hint, not a guarantee.
- `PREFER_INDEX` is a hint; the implementation may fall back to scan and report that via `warningFlags`.
- `Search(...)` remains synchronous; the host is responsible for worker-thread execution.
- Callback payload pointers are valid only for the duration of the callback.
- A plugin must stop promptly and return `HRESULT_FROM_WIN32(ERROR_CANCELLED)` if:
  - `FileSystemSearchShouldCancel` sets `*pCancel = TRUE`, or
  - a callback returns `E_ABORT` or `HRESULT_FROM_WIN32(ERROR_CANCELLED)`.

## Capabilities JSON

Extend `IFileSystem::GetCapabilities()` with an optional `search` section:

```json
{
  "search": {
    "version": 1,
    "name": true,
    "content": true,
    "indexed": true,
    "serviceBacked": true,
    "supportsRegex": true,
    "supportsSnippets": true,
    "preferredBackend": "service"
  }
}
```

Semantics:
- `name`: supports native name/path search.
- `content`: supports native content search.
- `indexed`: can answer name/path queries from an index instead of full scan.
- `serviceBacked`: can use an out-of-process service backend.
- `supportsRegex`: supports regex for any enabled query mode.
- `supportsSnippets`: may populate `previewText`.
- `preferredBackend`: one of `scan`, `index`, `service`.

If `search` is absent:
- the host assumes no native search support,
- the host may still provide scan fallback when the plugin exposes the required base interfaces.

## Architecture

### 1. Host search stack

The host owns:
- Find dialog lifecycle,
- search session threading,
- result-set algebra,
- plugin capability discovery,
- host-side scan fallback,
- final result actions.

New host components:
- `FindDialog`
- `SearchSessionController`
- `SearchResultSet`
- `SearchFallbackEngine`
- `SearchTextHelpers`

### 2. Plugin-native search

If the active plugin exposes `IFileSystemSearch`, the host uses it first.

Expected plugin-native cases:
- `builtin/file-system`: yes
- `builtin/file-system-dummy`: optional later
- `builtin/file-system-7z`: no in v1
- remote plugins: optional later

### 3. Host scan fallback

If a plugin does not expose `IFileSystemSearch`, the host falls back to:
- recursive enumeration using `IFileSystem::ReadDirectoryInfo`,
- optional content streaming using `IFileSystemIO::CreateFileReader`.

Fallback eligibility:
- name-only search requires `IFileSystem`.
- content search requires both `IFileSystem` and `IFileSystemIO`.

Fallback guarantees:
- correctness over speed,
- identical dialog UX,
- degradation surfaced in progress/status.

### 4. Local indexed search

The built-in `file` plugin implements `IFileSystemSearch` with backend precedence:
1. Windows service
2. in-process local index core
3. scan fallback

The local index core is shared logic used by both the plugin and the service.

## Built-in File Plugin Plan

### Supported roots

Indexed:
- NTFS local volumes
- ReFS local volumes

Scan-only:
- FAT/exFAT
- UNC shares
- WSL paths
- unsupported or unknown file systems

### NTFS plan

Initial seed:
- `FSCTL_ENUM_USN_DATA`

Live updates:
- `FSCTL_READ_USN_JOURNAL`

Recovery:
- rebuild on journal ID mismatch,
- rebuild on detected wrap/overflow,
- rebuild on snapshot corruption.

### ReFS plan

Initial seed:
- recursive `NtQueryDirectoryFile` traversal

Live updates:
- `FSCTL_READ_USN_JOURNAL`

### Scan fallback plan

Use the same path normalization and traversal rules already used by the file plugin for directory enumeration.

For content search:
- open each candidate via `IFileReader`,
- apply text heuristics,
- stream through chunked matcher with overlap,
- respect `maxContentBytesPerFile`,
- report `DEGRADED_NO_INDEX` when index was preferred but unavailable.

### Content search semantics

V1 content search is text-only.

Detection and decoding rules:
- detect BOMs for UTF-8/UTF-16/UTF-32,
- validate UTF-8 when no BOM is present,
- treat obvious binary files as skipped,
- default to ANSI/ACP only for text-oriented fallback,
- do not expose hex/binary matching in the Find dialog.

Snippet generation:
- optional,
- best effort,
- centered around the first content hit when cheaply available,
- clipped to `maxSnippetCharacters`.

## Search Service Plan

### Service role

`RedSalamanderSearchService` is the preferred local elevated backend for name/path/metadata search.

The service owns:
- local volume discovery,
- NTFS/ReFS index build and persistence,
- live journal monitoring,
- query candidate generation for local roots,
- backend status and rebuild requests.

The service does not own in v1:
- file-content scanning,
- snippet generation from file bodies,
- non-local file system plugins.

### IPC

Use a versioned named-pipe protocol.

Requirements:
- local-machine only,
- interactive-user clients,
- bounded message sizes,
- explicit protocol versioning,
- batched result frames,
- cancellation support,
- status and rebuild commands.

Initial commands:
- `Hello`
- `Query`
- `CancelQuery`
- `GetBackendStatus`
- `RequestRebuild`

### Storage

Persist service-owned index data under:

```text
%ProgramData%\RedSalamander\SearchIndex\
```

Storage contents:
- per-volume metadata,
- persistent file-entry snapshot,
- journal checkpoints,
- recovery markers.

### Security

- Service must run with sufficient privilege to access volume and journal APIs.
- Pipe ACLs must restrict access to local interactive users.
- Service query surface must remain read-only.
- Service must never return raw file contents in v1.

### Installer

Installer work items:
- install/uninstall service,
- upgrade-safe restart behavior,
- preserve or clean index store based on upgrade vs uninstall,
- tolerate environments where service registration is unavailable.

## Host Implementation Details

### Find dialog

The dialog is modeless and owned by the main window.

The UI implementation should follow the existing host rendering and theming architecture:
- render through the current DirectX-based UI path used by themed host surfaces
- use the existing custom control framework as the default building block set
- prefer extending current controls over introducing parallel one-off widgets
- add new controls only where search needs capabilities not already available
- respond to runtime theme changes, DPI changes, and font updates the same way as other host windows

It stores:
- current pane scope,
- last query,
- last options,
- current result set,
- current worker/session,
- current capability summary.

Suggested controls:
- root path
- name pattern
- content pattern
- recurse
- include files
- include directories
- name mode
- content mode
- case-sensitive name
- case-sensitive content
- prefer index
- want snippets
- max results
- cancel button
- result-set operation buttons
- result list
- status text

Control strategy:
- text entry, toggles, combo/dropdown behavior, and buttons should reuse existing themed controls if they already satisfy search needs
- the result list may require extending an existing owner-drawn or DirectX-rendered list/grid control for multi-column search results, incremental insertion, sorting hooks, and high row counts
- the status surface may require a themed progress/banner control if current controls do not already cover backend, degradation, and summary display cleanly
- any new control created for search should be evaluated for reuse elsewhere instead of being treated as Find-specific throwaway UI

### SearchSessionController

Responsibilities:
- launch search on a worker thread,
- choose native vs fallback engine,
- stream results to the dialog,
- coalesce progress UI updates,
- own cancellation token/state,
- maintain operation generation IDs to discard stale callbacks.

### Result-set algebra

Set identity key:
- plugin ID
- normalized instance context
- normalized full path

Operations:
- `Find`: replace current result set
- `Append`: union
- `Intersect`: keep only items already present and newly found
- `Subtract`: remove newly found items from the current set

### Result actions

Double-click behavior:
- directory result: navigate focused pane to the directory
- file result: invoke existing open/view behavior
- alternate action: navigate to parent and focus file if the direct open path is not desired

## Shared Helper Extraction

Move reusable text-search helpers from `ViewerText` into a shared module used by both:
- host fallback search,
- file-plugin content search.

Helpers to extract:
- BOM detection
- UTF-8 validation
- binary-data heuristic
- chunked forward text matching with overlap
- chunked backward matching if needed later

New helper module responsibilities:
- remain non-UI,
- avoid viewer-specific state,
- work from `IFileReader`,
- support literal and regex content matching.

## Plugin Matrix

### V1 expectations

`builtin/file-system`
- native search: yes
- indexed backend: yes
- service backend: yes
- content search: yes

`builtin/file-system-dummy`
- native search: optional
- fallback usable: yes

`builtin/file-system-7z`
- native search: no
- fallback usable:
  - name search: yes
  - content search: only if reader path is acceptable for archive entries in host fallback

Remote file system plugins
- native search: optional
- fallback usable:
  - name search: yes if enumeration exists
  - content search: yes if `IFileSystemIO` exists

## Testing Scenarios

### ABI and contract tests

- reject null `query`
- reject null `callback`
- reject invalid `sizeBytes`
- reject query with both `nameMode` and `contentMode` disabled
- ensure callback payloads always set `sizeBytes`
- ensure cancellation returns `HRESULT_FROM_WIN32(ERROR_CANCELLED)`
- ensure no callbacks occur after `Search(...)` returns

### Host UX tests

- `Alt+F7` opens the real Find dialog
- `cmd/pane/find` no longer routes to not-implemented UI
- dialog remains modeless
- `Find` replaces results
- `Append` unions results
- `Intersect` narrows correctly
- `Subtract` removes correctly
- `Cancel` is responsive during scan and content read
- double-click directory navigates
- double-click file opens correctly

### Fallback correctness tests

- local file plugin without native search forced off
- dummy plugin
- `7z` plugin
- one remote plugin
- name-only query
- content-only query
- name + content query
- Unicode paths
- long paths
- access denied items
- empty roots
- reparse point loops

### Indexed backend tests

- NTFS initial seed
- ReFS initial traversal
- journal replay after create/delete/rename
- rename storms
- delete/create bursts
- journal wrap recovery
- snapshot reload
- backend fallback when service is down
- parity between indexed and scan results for the same query/root

### Content search tests

- UTF-8 with BOM
- UTF-8 without BOM
- UTF-16 LE/BE
- UTF-32 LE/BE
- invalid UTF-8
- ANSI text fallback
- binary skip heuristic
- large-file chunk boundary match
- literal content search
- regex content search
- case-sensitive content search
- snippet generation around first hit

### Performance targets

- warm indexed name search returns first batch in under 100 ms
- cancel responsiveness stays under 250 ms
- UI thread never blocks on active search
- indexed and scan backends produce equivalent final result sets for supported local roots

## Review Checklist

Review this spec in the following order:

1. ABI review
- updated `IFileSystemSearch` payloads
- `sizeBytes` compatibility rules
- capability JSON shape

2. UX review
- dialog scope and operation model
- result actions
- degradation messaging
- use of the existing DirectX/themed custom control stack
- runtime theme-change and DPI behavior

3. Backend review
- scan fallback correctness
- local index core split
- NTFS/ReFS support boundaries
- service fallback rules

4. Security review
- named-pipe ACLs
- service privilege level
- `%ProgramData%` storage
- no content indexing in service v1

5. Test review
- coverage of fallback vs indexed parity
- coverage of content decoding heuristics
- performance and cancellation validation

## Implementation Phases

This section expands the rollout order into implementation phases with explicit validation gates. A phase is only complete when its exit criteria are satisfied and the corresponding checklist items above are checked.

### Cross-phase validity rules

The implementation is valid only if all of the following remain true through every phase:
- the public ABI and the plugin documentation stay synchronized
- unsupported plugins or roots degrade to correct scan behavior rather than failing the command
- search work runs off the UI thread and cancellation remains responsive
- indexed execution and scan execution converge to the same final result set for equivalent supported queries
- service use does not change the security boundary for file content in v1

### Phase 0 - Contract and design freeze

Objective:
- lock the public contract before behavior is distributed across the host, the built-in file plugin, and the service

Implementation approach:
- update `Common/PlugInterfaces/FileSystem.h` first so the source of truth for the ABI is concrete
- immediately mirror those changes into `Specs/Plugins/Plugins_VirtualFileSystem.md`
- define compatibility rules for `sizeBytes`, including how callers and implementations must behave when they know different struct sizes
- define canonical identity for result-set algebra before any UI or backend code stores results

Validity checks:
- the header, plugin documentation, and this plan agree on every enum, field, and default
- review confirms the unchanged `Search(...)` signature is enough and the added payload fields cover v1 scope
- every known search caller can be upgraded without inventing a second interface generation

### Phase 1 - Shared host search foundation

Objective:
- establish the always-correct fallback path and the reusable text-search primitives before adding optimized backends

Implementation approach:
- move the content-decoding and text-matching primitives into a shared helper module so both the host fallback and the local file plugin can reuse the same behavior
- implement the generic scan engine as the behavioral baseline, using plugin directory enumeration plus file-reader streams
- report progress in a backend-agnostic way so the UI does not need different code paths for scan versus indexed backends
- treat missing I/O capability as a degraded mode instead of a hard failure when content search is requested

Validity checks:
- unit tests prove the helpers handle BOMs, invalid UTF-8, ANSI fallback, binary detection, and chunk overlap correctly
- integration tests prove the scan engine works for plugins with and without native search
- cancellation and progress reporting work on deep trees and large files without UI stalls

### Phase 2 - Find dialog and session orchestration

Objective:
- deliver the user-facing search workflow once the fallback engine already works

Implementation approach:
- introduce a dedicated dialog and `SearchSessionController` so search state does not leak into folder browsing code
- build the UI on the existing themed DirectX/custom-control stack instead of creating a separate stock-dialog path, so search inherits the current visual language and runtime theme behavior
- reuse existing controls first, then extend them where necessary for search-specific density, streaming, and results presentation
- run query execution on a worker thread and post batched results back to the dialog
- keep set operations in the host using canonical identities so backend implementations stay simple
- surface backend choice, degraded capabilities, and final summary consistently in the dialog status area

Validity checks:
- the old not-implemented Find path is removed
- repeated start/cancel/start flows do not leak state or keep stale results
- append/intersect/subtract are correct for identical searches, disjoint searches, and mixed-backend searches
- double-click behavior is consistent for both files and directories
- the dialog fits the active theme, survives runtime theme changes, and does not regress keyboard, focus, DPI, or scrolling behavior compared with other host windows

### Phase 3 - Built-in file plugin native search adoption

Objective:
- make the local file plugin participate in the public search contract before index features are layered in

Implementation approach:
- implement `IFileSystemSearch` in the built-in file plugin with scan-backed behavior first
- centralize path normalization and metadata filling so later indexed and service backends can reuse the same result shaping
- add backend selection plumbing early, even if only scan is active at first, so later phases only plug in new executors

Validity checks:
- host fallback and plugin-native scan mode return the same results for the same queries
- local plugin capabilities advertise only the features that are actually enabled
- unsupported local roots always degrade to scan mode instead of partial or incorrect indexed behavior

### Phase 4 - Shared local index core

Objective:
- add the local instant-search engine in a reusable form without making the service mandatory

Implementation approach:
- create a shared core library responsible for snapshots, journal replay, path normalization, and candidate lookup
- use NTFS USN/MFT APIs where available and a ReFS traversal-plus-journal model where USN exists but MFT-style seeding does not
- keep the index limited to names, paths, and metadata in v1; let content search remain a streamed client-side filter on candidate files
- define rebuild paths up front so corruption and journal churn do not leave the index in a half-valid state

Validity checks:
- indexed candidate lookup is measurably faster while still converging to the same final matches as scan mode
- snapshot reload and journal replay survive restart, rename storms, and delete/create bursts
- unsupported roots never enter the indexed path

### Phase 5 - Windows service backend

Objective:
- enable elevated indexing and monitoring without forcing the main application to hold those privileges

Implementation approach:
- host the shared index core in `RedSalamanderSearchService`
- communicate over a private versioned named-pipe protocol that sends query requests, progress frames, and result batches
- keep content scanning in the client process so the service never becomes a privileged file-content broker in v1
- expose only read-oriented operational controls such as status and rebuild

Validity checks:
- pipe ACLs restrict access appropriately
- client fallback to in-process index or scan works when the service is missing, stopped, or mid-restart
- multi-client queries, service upgrades, and corrupted on-disk state all recover safely

### Phase 6 - Hardening, parity, and rollout

Objective:
- prove the feature is correct, secure, and supportable before enabling it broadly

Implementation approach:
- run the full test matrix against the host fallback, plugin-native scan, in-process index, and service backends
- add instrumentation and diagnostics for backend selection, rebuild causes, degraded execution, and cancellation timing
- review UI polish, themed rendering behavior, keyboard flow, localization, and settings persistence
- decide rollout defaults only after parity and stability data are available

Validity checks:
- performance targets are met or explicit mitigations are documented
- acceptance evidence covers the important filesystem and failure-path scenarios
- a final release-candidate human pass should still cover theme switching, high DPI, and custom-control interaction regressions on the custom Find dialog
- unresolved issues are tracked explicitly rather than hidden behind a nominally complete phase

## Assumptions

- V1 content search is text-oriented only.
- The existing RFC for USN/MFT indexing remains the low-level design reference.
- Changing `IFileSystemSearch` directly is lower-risk than introducing `IFileSystemSearch2` because the interface is defined but not yet adopted by the host or shipped plugins.
