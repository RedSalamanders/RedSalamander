# Compare Directories — Implementation tracker

Last updated: 2026-03-04

This is the living engineering tracker for **Compare Directories**:

- Cross‑plugin compare (different filesystem plugins and/or different instance contexts)
- Responsiveness & shutdown safety (no UI-thread deep traversal; close/app-exit remains responsive)

## Status / Checklist

### Cross‑plugin enablement (done)

- [x] Spec: update `Specs/Core/Core_CompareDirectories.md` for cross‑plugin compare (dual FS/IO + path rules + scope rules)
- [x] Host launch: remove compare guard + plumb per‑pane contexts (`RedSalamander/RedSalamander.cpp`)
- [x] Compare window: accept per‑pane contexts; create per‑pane FS instances; inject wrappers with correct metadata (`RedSalamander/CompareDirectoriesWindow.h/.cpp`)
- [x] FolderWindow safety: don’t overwrite injected FS when plugin id matches (`RedSalamander/FolderWindow.FileSystem.cpp`)
- [x] Engine: dual‑filesystem session + per‑style path math + dual‑IO content compare (`RedSalamander/CompareDirectoriesEngine.h/.cpp`)
- [x] Wrapper QI: forward unknown `QueryInterface` to base FS (so optional interfaces work)
- [x] UI: disable `compareContent` toggle when unsupported; explain non‑fatally (localized)
- [x] Selftests: update existing tests for new session ctor + add dual‑IO regression + plugin‑path regression (`RedSalamander/CompareDirectoriesEngine.SelfTest.cpp`)
- [x] Selftests (optional): add remote smoke cases for file↔S3 and file↔FTP (skipped unless connection profiles + secrets exist) (`RedSalamander/CompareDirectoriesEngine.SelfTest.cpp`)
- [x] Build gates: `.\build.ps1 -ProjectName RedSalamander` + `.\build.ps1`
- [x] Test gate: `--compare-selftest`
- [ ] Manual matrix: file↔file, file↔S3, S3↔S3

### Startup freeze + shutdown hang fixes (in progress)

- [x] Spec: update `Specs/Core/Core_CompareDirectories.md` to require background subtree scan + UI cache-only access
- [x] Engine: remove synchronous subtree traversal from `CompareDirectoriesSession::GetOrComputeDecision`
- [x] Engine: add background folder-scan workers + progressive `SubdirPending` semantics
- [x] Engine: keep memory bounded in differences-only mode (elide per-file `ContentPending` placeholders; track folder-level pending count)
- [x] Engine: skip content compare enqueue when enabled non-content criteria already differ (avoid useless remote reads)
- [x] UI: replace UI-thread `GetOrComputeDecision` calls with cache-only (`TryGetCachedDecision`)
- [x] UI: extend decision refresh timer to flush both content + subdir updates (budgeted)
- [x] UI: add `keepIdenticalItems` retention toggle + gate **Compare → Show Identical Items** view toggle (no rescan for view toggle) (`RedSalamander/CompareDirectoriesWindow.cpp`, `Specs/Core/Core_CompareDirectories.md`)
- [x] Selftests: update `compareSubdirectories` cases for progressive scan completion
- [x] Selftests: add regression proving `GetOrComputeDecision` does not enumerate descendants
- [x] Selftests: add regression for differences-only content-compare elision (`content_pending_elided`)
- [x] S3: content compare reader uses ranged `GetObject` (no full-object download to temp file before first byte compare)
- [ ] Shutdown: avoid UI-thread joins when closing compare window / app mid-compare
  - [x] Cleanup scheduling is always off-UI (threadpool; scheduling failure logs once and abandons deferred cleanup rather than blocking UI teardown or detaching a thread).
  - [ ] Verify with S3↔FS “close app mid-compare” repro (ensure exit is prompt; no stuck process).

### Remote perf + memory + exit (in progress)

- [x] Instrumentation: compare perf stats + `DirectoryInfoCache` stats snapshot (`RedSalamander/CompareDirectoriesEngine.h/.cpp`, `RedSalamander/CompareDirectoriesWindow.cpp`)
- [x] Engine: compare scans must not populate `DirectoryInfoCache` (`RedSalamander/CompareDirectoriesEngine.cpp`)
- [x] Engine: visible-first scan/content scheduling (hi/lo) + upgrade semantics (`RedSalamander/CompareDirectoriesEngine.h/.cpp`)
- [x] Engine: bounded content-compare queues + low-priority backpressure (`RedSalamander/CompareDirectoriesEngine.h/.cpp`)
- [x] Engine: content-compare cache partial eviction (no nuclear clear) (`RedSalamander/CompareDirectoriesEngine.cpp`)
- [x] Engine: decision-cache eviction budget (~300MB) + pinned visible folders/ancestors (`RedSalamander/CompareDirectoriesEngine.h/.cpp`, `RedSalamander/CompareDirectoriesWindow.cpp`)
- [x] S3: set `requestTimeoutMs=30'000` (no infinite hangs) (`Plugins/FileSystemS3/FileSystemS3.Shared.cpp`)
- [x] S3: reuse S3 clients per ctx; ranged reader uses shared client; chunk=8MiB (`Plugins/FileSystemS3/*`)
- [x] Selftests: add bounded-memory + cache-pollution regressions (`RedSalamander/CompareDirectoriesEngine.SelfTest.cpp`)
- [ ] Manual: S3↔FS large tree remains responsive; cancel/close exits promptly

## Process

- Every PR/commit that completes a checkbox MUST tick it in this file in the same change.

## Resolved findings (keep for history)

These items were originally flagged during implementation, and are now already addressed in the codebase:

- Hash/equality mismatch in case-insensitive name matching: resolved by switching to ordered maps using `WStringViewNoCaseLess` (no `unordered_map` hash/eq mismatch).
- Recursion eliminated: directory traversal uses an iterative worklist (no recursive call-stack growth).
- Content compare responsiveness: content compare runs asynchronously on background workers; the UI coalesces decision-update refreshes via a timer, applies pending updates in bounded batches, and borrows base enumerations via `DirectoryInfoCache` to avoid re-enumerating remote plugins on every refresh tick.
- Content compare Win32 direct I/O: uses `IFileSystemIO::CreateFileReader` + `IFileReader::Read` (no `CreateFileW`/`ReadFile` in compare logic).
- Root cache key sentinel: root cache key uses `L"."` (not `""`).
- Banner/UI strings: localized via `.rc` strings (no hardcoded “Compare Folder / Options… / Rescan”).
- Window ownership: `ShowCompareDirectoriesWindow` uses `std::unique_ptr` + `release()` ownership transfer (no manual `delete` on failure).
- Options dialog proc: extracted as `OptionsDlgProc` (no giant inline lambda).

## Deferred / follow-ups (not required for cross‑plugin compare)

- `g_compareDirectoriesWindows` global vector: UI-thread-only assumption should be documented or replaced with a safer registration strategy.
- Splitter drag invalidation: avoid invalidating the full window on every drag tick.
- Extra selftests (nice-to-have): reparse-point skipping, concurrent invalidation stress, very deep trees, “size unknown” behavior.
