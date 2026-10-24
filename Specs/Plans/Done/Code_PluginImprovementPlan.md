# Plugin Improvement Plan — FileSystem7z, FileSystemCurl, Cloud Plugins

Squad safety review of 5 filesystem plugins (~37K LOC). All Phase 1-3 and 4.2 safety fixes verified landed at HEAD `275c04034` (2026-07-02 folder review). Observatory Track 12 closed the two remaining items on 2026-07-17.

Scope exclusions (unchanged): FileSystemS3 (clean, no changes), FileSystemDummy (test stub), Google Drive feature parity, shared-utility extraction across cloud plugins (separate initiative).

---

## Completed & verified (2026-07-02)

- [x] **1.1-1.4** Module pinning via `AcquireModuleReferenceFromAddress` at each thread launch site (FileSystemCurl scheduler/reader/writer, FileSystem7z extract thread).
- [x] **2.1** FileSystemMicrosoftDrive `SetCallback` generation tracking + condition-variable drain guard (FileSystemMicrosoftDrive.cpp:4970-5020).
- [x] **2.2** `GetConfiguration` double-buffer fix — no more `.c_str()` returned after lock release (FileSystemMicrosoftDrive.cpp:4909).
- [x] **3.1** FileReader/FileWriter own the plugin via `wil::com_ptr<IFileSystem>` instead of manual AddRef/Release (FileSystemMicrosoftDrive.cpp:4561 / :4678).
- [x] **4.2** FileSystemCurl retry infrastructure for transient network errors (landed in commit `e5bbfb023`).

---

## Completed 4.1 — coordinated libcurl shutdown

- [x] **4.1** Coordinate `CurlEasyPool` drain and process-global libcurl cleanup outside `DLL_PROCESS_DETACH`.

The box was previously checked on the strength of commit `e5bbfb023` (2026-03-21), whose MESSAGE claims `DrainPool`/`ShutdownCurlEasyPool`/`curl_global_cleanup` were added — but its diff (1 file, 56 insertions) contains only the 4.2 retry helpers.

Verified 2026-07-02 folder review:

- No `ShutdownCurlEasyPool`/`curl_global_cleanup` exists anywhere in the tree; the only mention is the original defect record (`.squad/decisions.md:987`).
- The instance=0 quiet point (FileSystemCurl.Shared.cpp:3635-3638) shuts down only the copy/move job scheduler.
- The static `CurlEasyPool` (FileSystemCurl.Shared.cpp:2211) still destructs during DLL-unload CRT teardown under loader lock — the exact `DLL_PROCESS_DETACH` crash hazard 4.1 was filed to prevent.

**Implemented 2026-07-17:** `Common::CurlRuntime::ProcessLease` coordinates independently unloadable Curl and
Google Drive DLLs. FileSystemCurl stops new instances, waits for active easy-pool borrows, drains idle handles,
releases its share handle, and only then releases its process participant. Google Drive follows the same explicit
instance/handle quiet point. Only the final process participant calls `curl_global_cleanup`, outside loader lock.
The contract test performs eight alternating physical-unload cycles and creates a real easy handle from the
survivor after each peer unload.

---

## Completed 4.3 — cancellable FileSystem7z indexing

- [x] **4.3** Keep index construction on FolderView's existing background enumeration worker and make the full synchronous call cancellable, progress-aware, generation-gated, and bounded.

FileSystem7z.cpp:852-948 still builds the index inline on the calling thread (`BuildIndex` at :907, `_indexBuildCv.wait` at :878-882, no cancellation token) — the UI blocks uncancellably on large archives. Deferred because cancellation must be threaded through the 7zip COM interface. (Distinct from Clearwater CW-4, which is the normalized-key collision bug in the same file.)

**Decision and implementation 2026-07-17:** a second plugin-owned worker would duplicate the host worker and
complicate module unload. The optional `IFileSystemCancellableDirectoryEnumeration` instead propagates the host
worker stop token synchronously through archive open, concurrent-builder waits, entry scan, index construction,
sorting, and injected delay. Configuration/mount changes invalidate builders by generation. Cancellation returns
`ERROR_CANCELLED`, publishes no partial result, is not cached as a permanent failure, and a subsequent retry
succeeds. The index is capped at 1,000,000 entries and 128 MiB of retained key text; password copies are securely
cleared with their session owners.

---

## Closeout

Complete. Debug builds for FileSystem7z, FileSystemCurl, FileSystemGoogleDrive, RedSalamander, and
PluginContractTests produced zero warnings/errors. `PluginContractTests` passed, including 26/26 7z Debug checks,
transactional configuration/legacy-secret proofs, and eight cross-plugin libcurl refresh cycles. The complete
source-contract suite passed 145/145. Durable behavior is recorded in
`Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Core/Core_SharedHelpers.md`, and the provider/testing specs.
