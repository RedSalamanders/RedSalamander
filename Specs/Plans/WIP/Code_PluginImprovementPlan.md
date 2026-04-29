# Plugin Improvement Plan — FileSystem7z, FileSystemCurl, Cloud Plugins

**Status:** WIP - one deferred FileSystem7z cancellation gap plus closeout verification remains
**Branch:** `squad/dxui-filesystem-improvements`
**Worktree:** `Z:\src\RedSalamander-improvements`
**Based on:** Squad deep review of 5 filesystem plugins (~37K LOC)

---

## Closeout audit (`2026-04-25`)

The checked safety items still match the current code at a targeted symbol level: FileSystem7z has module pinning for extract threads, and the cloud/Curl plugin hardening work is represented in the plugin implementations. This plan stays in WIP because `FileSystem7z::EnsureIndex()` still builds synchronously and waits behind `_indexBuildInProgress`; it is not yet non-blocking or cancellable for large archives.

Remaining closeout checklist:

- [ ] Decide whether `FileSystem7z::EnsureIndex()` remains a documented synchronous limitation or gets a cancellable/non-blocking index-build path.
- [ ] Refresh full solution and targeted plugin verification after that decision.
- [ ] Move this plan to Done after verification, or split the 7z indexing limitation into a separate WIP plan and close the completed safety slice.

## Summary

Deep reviews by Sysadm (FileSystem7z, FileSystemCurl, FileSystemS3) and Ripley (FileSystemGoogleDrive, FileSystemMicrosoftDrive) identified critical module-pinning gaps, callback safety issues, and several medium-priority improvements. FileSystemS3 is clean — no changes needed.

---

## Phase 1 — Critical: Module Pinning (P0)

Same crash-on-FreeLibrary pattern already fixed in core FileSystem plugin. Copy-paste fix: `AcquireModuleReferenceFromAddress(...)` at each thread launch site.

- [x] **1.1** FileSystemCurl: Pin `SharedCopyMoveJobScheduler` singleton (static jthread pool survives DLL unload)
- [x] **1.2** FileSystemCurl: Pin `CurlStreamingReader` jthread (COM object returned to host, thread captures `this`)
- [x] **1.3** FileSystemCurl: Pin `CurlStreamingWriter` jthread (same pattern as reader)
- [x] **1.4** FileSystem7z: Pin `SevenZipItemFileReader` extract jthread (COM object returned to host)

## Phase 2 — Critical: Callback Safety (P0)

- [x] **2.1** FileSystemMicrosoftDrive: Add drain guard to `SetCallback` — generation tracking + condition variable drain, matching the pattern already used in FileSystemGoogleDrive
- [x] **2.2** FileSystemGoogleDrive + FileSystemMicrosoftDrive: Fix `GetConfiguration` returning `.c_str()` under lock — caller uses pointer after lock release. Return `std::wstring` copy or use a shared lock that outlives the return.

## Phase 3 — High: COM Ownership (P1)

- [x] **3.1** FileSystemMicrosoftDrive: Replace raw `FileSystemMicrosoftDrive*` with `wil::com_ptr<IFileSystem>` in FileReader/FileWriter classes (manual AddRef/Release violates AGENTS.md rule)

## Phase 4 — Medium: Reliability & Performance (P2)

- [x] **4.1** FileSystemCurl: Add `curl_global_cleanup` coordination for static `CurlEasyPool` — prevent crash during `DLL_PROCESS_DETACH` when pool is still active
- [x] **4.2** FileSystemCurl: Add retry policy for transient network errors (single-attempt curl operations currently fail permanently on transient errors)
- [ ] **4.3** FileSystem7z: Make `EnsureIndex()` non-blocking or cancellable — currently blocks UI thread with no way to cancel on large archives (DEFERRED: requires threading cancellation through 7zip COM interface — complex)

## Phase 5 — Verification

- [ ] **5.1** Full solution build — zero new warnings
- [ ] **5.2** Commit all changes with descriptive messages
- [ ] **5.3** Update this plan status to ✅ Complete

---

## Scope Exclusions

- **FileSystemS3** — Gold standard, zero issues found. No changes.
- **FileSystemDummy** — Test stub, not reviewed.
- **Google Drive feature parity** — Out of scope. Google Drive is an early alpha; feature additions are a separate initiative.
- **Shared utility extraction** — ~400 lines duplicated across cloud plugins. Worth doing but separate from safety fixes.

---

## Review Sources

- Sysadm: FileSystem7z (~4.7K LOC), FileSystemCurl (~15.3K LOC), FileSystemS3 (~8.6K LOC)
- Ripley: FileSystemGoogleDrive (~2.8K LOC), FileSystemMicrosoftDrive (~6.4K LOC)
- Review criteria: threading safety, RAII compliance, error handling, module pinning, performance, API compliance
