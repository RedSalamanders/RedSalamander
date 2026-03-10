# Search Specification

## Overview

RedSalamander search is a host-owned feature that spans the command system, the Find dialog, optional plugin-native search, a generic host scan fallback, the built-in local index core, and the Windows search service backend.

This document is the authoritative architecture and behavior spec for that system.

Related normative documents:
- `Specs/Plugins/Plugins_VirtualFileSystem.md` for the `IFileSystemSearch` ABI and capability JSON.
- `Specs/UI/UI_CommandMenuKeyboard.md` for `cmd/pane/find` command routing and shortcut defaults.
- `Specs/UI/UI_FolderWindow.md` for pane ownership, focus routing, and modeless window integration.
- `Specs/Core/Core_SettingsStore.md` for persisted Find dialog state.
- `Specs/Testing/Testing_SelfTests.md` for self-test result and skip semantics.

Historical implementation detail and rollout evidence remain in:
- `Specs/Plans/WIP/Core_Search_PluginArchitecturePlan.md`
- `Specs/Plans/WIP/RFC_Core_InstantFileSearchUsnMftIndex.md`

## Goals

- Provide one Find experience for all filesystem plugins.
- Prefer fast local indexed search when the active pane is browsing a supported local root.
- Preserve correctness when native or indexed search is unavailable by falling back to a generic scan path.
- Keep the UI responsive by running all searches off the UI thread.
- Make degradation visible instead of silently changing behavior.

## Non-Goals

- No separate search-plugin type in v1.
- No synthetic `search:` virtual filesystem in v1.
- No binary or hex content search in the Find dialog in v1.
- No service-side file-content indexing or content transport in v1.

## Command and Window Behavior

### `cmd/pane/find`

- `cmd/pane/find` opens or activates the modeless `Find Files and Directories` window.
- The command targets the focused pane when focus is inside a pane; otherwise it targets the active pane.
- The default shortcut is `Alt+F7`.
- If the window already exists, the host must reuse it and refresh its pane context instead of opening duplicates.

### Default scope

The initial search scope is derived from the target pane:
- active filesystem plugin,
- active instance context,
- current pane path as `rootPath`.

### Find window responsibilities

The Find window is host-owned and theme-aware. It must:
- run on the existing DirectX-rendered themed host control stack,
- stream results incrementally,
- show backend, progress, warning, and completion state,
- preserve keyboard/focus behavior consistent with the rest of the app,
- persist placement and recent inputs through the shared settings store.

## Query Semantics

- At least one of name matching or content matching must be enabled.
- If both name and content matching are enabled, a result must satisfy both.
- Name matching supports wildcard, literal, and regex modes.
- Content matching supports text literal and text regex modes.
- v1 content search is text-oriented only.
- `FILESYSTEM_SEARCH_WANT_SNIPPETS` is a hint, not a guarantee.
- `FILESYSTEM_SEARCH_PREFER_INDEX` is a hint, not a guarantee.

## Result Model

### Set operations

The Find window supports four result-set operations:
- `Find`: replace the current result set.
- `Append`: union the new result set into the current result set.
- `Intersect`: keep only items present in both sets.
- `Subtract`: remove newly found items from the current result set.

### Canonical result identity

Result identity is frozen as:
- `pluginId + normalized instanceContext + normalized fullPath`

This identity is used by the Find window for de-duplication and set algebra across native, indexed, service, and fallback backends.

### Result actions

- Double-clicking a file opens it through the existing host open/view flow.
- Double-clicking a directory navigates the pane to that directory.
- The secondary parent action navigates to the parent directory and focuses the matched item when possible.

## Execution Model

### SearchSessionController

The host runs searches through a dedicated session controller that owns:
- worker-thread lifetime,
- cancellation,
- backend selection,
- result batching,
- progress aggregation,
- completion delivery.

`IFileSystemSearch::Search(...)` remains synchronous, so the host must always invoke it on a worker thread.

### UI-thread contract

- The UI thread must never block on directory traversal, content reads, index warm-up, or service calls.
- Worker results are marshaled back to the Find window through posted messages.
- Cancellation must be checked at traversal boundaries and during chunked content reads.

## Backend Architecture

### Backend precedence

For the built-in local `file` plugin, backend selection order is:
- `service`
- `local-index`
- `scan`

The configuration override values are:
- `auto`
- `service`
- `local-index`
- `scan`

`auto` uses the precedence above.

### Generic host scan fallback

The host scan fallback is the correctness baseline for all plugins. It uses:
- `IFileSystem::ReadDirectoryInfo` for traversal,
- `IFileSystemIO::CreateFileReader` for content scanning when available.

Fallback behavior:
- If a plugin does not expose `IFileSystemSearch`, the host may still provide name search through scan fallback.
- Content fallback additionally requires `IFileSystemIO`.
- If no readable stream API exists, content search must degrade cleanly and surface `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT`.

### Built-in local indexed backend

The shared local index core is used by both the built-in `file` plugin and the Windows service backend.

Supported indexed roots:
- NTFS
- ReFS

Unsupported roots must fall back to scan:
- FAT / exFAT
- UNC
- WSL-backed paths
- other roots where required journal or enumeration support is unavailable

The local index core is responsible for:
- path normalization,
- snapshot persistence,
- query planning,
- journal replay,
- tracked-root traversal rebuilds when direct journal access is unavailable.

### Windows service backend

`RedSalamanderSearchService` is the preferred local backend when available.

Service rules:
- communicate through a versioned named-pipe protocol,
- store service-owned snapshots under `%ProgramData%\\RedSalamander\\SearchIndex\\`,
- keep file content out of the service protocol in v1,
- fall back safely to `local-index` or `scan` when the service is absent, mismatched, disconnected, or unavailable.

### Content search in indexed flows

In v1, indexed backends only narrow the candidate set. File contents remain client-side:
- the service or index core returns candidate paths and metadata,
- the client/plugin opens files through normal reader APIs,
- literal/regex text matching and snippet generation happen in the client process.

## Warning and Degradation Semantics

Search backends must report visible degradation through `FileSystemSearchProgress::warningFlags`, including:
- `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX`
- `FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT`
- `FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`
- `FILESYSTEM_SEARCH_WARNING_OVERFLOW`

The Find window must surface these warnings in status text rather than silently hiding them.

## Persistence and Configuration

### Settings Store

Persisted Find dialog state lives under:
- `search`

This includes:
- recent roots,
- recent name patterns,
- recent content patterns,
- last-used root and pattern values,
- name/content modes,
- recursive/include/match-case/follow-symlinks toggles,
- `preferIndex`,
- `wantSnippets`,
- `maxResults`.

Window placement is persisted through:
- `windows["FindFilesWindow"]`

### Built-in file plugin configuration

The built-in local file plugin persists backend preference under its plugin configuration payload:
- `searchBackendPreference`

Allowed values:
- `auto`
- `service`
- `local-index`
- `scan`

## Diagnostics and Verification

Search changes must be validated through:
- Commands self-tests for the Find window, command routing, persistence, and modeless ownership,
- Compare self-tests for native search, fallback parity, indexed parity, service parity, cancellation, Unicode paths, long paths, access denied, NTFS, and ReFS,
- packaged Release validation for MSI and MSIX when service packaging changes.

Machine-dependent coverage remains part of the suite and must skip with a reason when preconditions are absent. Example:
- the ReFS indexed probe remains declared and records `skipped` when no fixed ReFS volume exists.

## Normative Summary

- `cmd/pane/find` is implemented and opens the modeless Find window.
- The Find window is host-owned, themed, and worker-threaded.
- `IFileSystemSearch` is the native plugin contract; the host must provide scan fallback when native search is unavailable.
- The built-in local plugin prefers `service -> local-index -> scan`.
- Indexed backends return names, paths, and metadata; content matching remains client-side in v1.
- Every visible degradation must be surfaced through progress warnings and reflected in the UI.
