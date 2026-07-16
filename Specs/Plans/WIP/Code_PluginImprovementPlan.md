# Plugin Improvement Plan — FileSystem7z, FileSystemCurl, Cloud Plugins

Squad safety review of 5 filesystem plugins (~37K LOC). All Phase 1-3 and 4.2 safety fixes verified landed at HEAD `275c04034` (2026-07-02 folder review). Two items remain open: 4.1 (below — was wrongly checked done) and 4.3.

Scope exclusions (unchanged): FileSystemS3 (clean, no changes), FileSystemDummy (test stub), Google Drive feature parity, shared-utility extraction across cloud plugins (separate initiative).

---

## Completed & verified (2026-07-02)

- [x] **1.1-1.4** Module pinning via `AcquireModuleReferenceFromAddress` at each thread launch site (FileSystemCurl scheduler/reader/writer, FileSystem7z extract thread).
- [x] **2.1** FileSystemMicrosoftDrive `SetCallback` generation tracking + condition-variable drain guard (FileSystemMicrosoftDrive.cpp:4970-5020).
- [x] **2.2** `GetConfiguration` double-buffer fix — no more `.c_str()` returned after lock release (FileSystemMicrosoftDrive.cpp:4909).
- [x] **3.1** FileReader/FileWriter own the plugin via `wil::com_ptr<IFileSystem>` instead of manual AddRef/Release (FileSystemMicrosoftDrive.cpp:4561 / :4678).
- [x] **4.2** FileSystemCurl retry infrastructure for transient network errors (landed in commit `e5bbfb023`).

---

## OPEN 4.1 — FileSystemCurl shutdown coordination (was WRONGLY checked done)

- [ ] **4.1** Add `curl_global_cleanup` coordination for the static `CurlEasyPool` — prevent crash during `DLL_PROCESS_DETACH` when the pool is still active.

The box was previously checked on the strength of commit `e5bbfb023` (2026-03-21), whose MESSAGE claims `DrainPool`/`ShutdownCurlEasyPool`/`curl_global_cleanup` were added — but its diff (1 file, 56 insertions) contains only the 4.2 retry helpers.

Verified 2026-07-02 folder review:

- No `ShutdownCurlEasyPool`/`curl_global_cleanup` exists anywhere in the tree; the only mention is the original defect record (`.squad/decisions.md:987`).
- The instance=0 quiet point (FileSystemCurl.Shared.cpp:3635-3638) shuts down only the copy/move job scheduler.
- The static `CurlEasyPool` (FileSystemCurl.Shared.cpp:2211) still destructs during DLL-unload CRT teardown under loader lock — the exact `DLL_PROCESS_DETACH` crash hazard 4.1 was filed to prevent.

**Fix direction:** add explicit pool drain + `curl_global_cleanup` coordination at the plugin quiet point (before module release), mirroring the ordering contract in `Specs/Plugins/Plugins_VirtualFileSystem.md`.

---

## OPEN 4.3 — FileSystem7z::EnsureIndex non-blocking/cancellable (deferred since 2026-03-21, idle)

- [ ] **4.3** Decide: documented synchronous limitation vs cancellable index build.

FileSystem7z.cpp:852-948 still builds the index inline on the calling thread (`BuildIndex` at :907, `_indexBuildCv.wait` at :878-882, no cancellation token) — the UI blocks uncancellably on large archives. Deferred because cancellation must be threaded through the 7zip COM interface. (Distinct from Clearwater CW-4, which is the normalized-key collision bug in the same file.)

---

## Next action

Implement 4.1 — it is a real crash-hazard item that was believed done.

## Closeout

After 4.1 lands and the 4.3 decision is recorded, move this plan to Done.
