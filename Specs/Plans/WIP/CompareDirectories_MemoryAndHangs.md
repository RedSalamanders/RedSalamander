# Compare Directories (S3↔FS): Memory Explosion + Hangs (root cause and fixes)

## Context

When comparing an S3 bucket vs local filesystem with **Compare Content** enabled (often with **Compare Subdirectories** enabled), the process can:

- Consume **multiple GiB of RAM** (often ~2 GiB on typical dev machines).
- Feel “stuck” (UI sluggish), especially at startup and during large completion bursts.
- Become slow to exit when closing the app mid-compare (remote calls can hang indefinitely).

This document is the engineering plan to make S3↔FS compares **bounded-memory** and **smooth**.

## Root causes (what actually drives the GiB spike)

### 1) Compare scan pollutes `DirectoryInfoCache` (GiB-scale by design)

The compare engine’s scan path enumerates each folder via:

- `DirectoryInfoCache::BorrowDirectoryInfo(..., BorrowMode::AllowEnumerate)`

`DirectoryInfoCache` is a global cache sized to a **large fraction of physical RAM** by default (roughly `RAM/16`, clamped up to **4 GiB**). When a subtree scan touches many folders, the cache can legitimately grow toward that limit, which looks like a compare-session “memory leak” but is actually global cache retention.

**Fix:** Compare scans MUST enumerate via direct `IFileSystem::ReadDirectoryInfo` instead of using the global directory cache.

### 2) Unbounded backlog (producer outpaces consumer)

For large trees, background scan workers can discover content-compare candidates much faster than the content compare workers can drain them (remote I/O is slow). Without bounds and prioritization, queue/inflight structures can grow large, and UI update notifications can become noisy.

**Fix:** Add a **visible-first** scheduler:
- High priority: folders the user navigates to (or the panes are currently showing).
- Low priority: background subtree scan.

Bound the total queued work with **backpressure on low priority work** (never on UI).

### 3) Infinite remote request timeouts → cancel/exit can hang indefinitely

The S3 plugin currently uses an infinite request timeout (`requestTimeoutMs = 0`). If the network hangs (or DNS stalls), worker threads can remain inside AWS calls indefinitely, making “Cancel” or app exit feel stuck.

**Fix:** Set a finite default request timeout (30s) and reuse client resources efficiently.

## Implementation plan (files + decisions)

### A) Engine: stop compare scans from using `DirectoryInfoCache`

**File:** `RedSalamander/CompareDirectoriesEngine.cpp`

- In `TryReadDirectoryEntries(...)`, replace `DirectoryInfoCache::BorrowDirectoryInfo(...AllowEnumerate)` with direct `baseFs->ReadDirectoryInfo(...)`.
- Keep the existing `FileInfo` parsing and ignore-pattern logic unchanged.

**Why:** The compare session is already caching decisions; the global directory cache should not be used as a “scan history” for remote subtree scans.

### B) Engine: visible-first queues + bounded backlog

**Files:** `RedSalamander/CompareDirectoriesEngine.h/.cpp`

- Split scan queue: high (visible/requested) vs low (background subtree).
- Split content-compare queue: high vs low, with a reserved high-capacity.
- Apply **backpressure only to low-priority producers** (scan workers), never UI:
  - Low scan waits on a condition variable when the low content queue is full.
  - Wakes on queue space, cancel token/version change, stop request, or background-work disable.

### C) Engine: bounded caches (avoid nuclear clears, enforce a budget)

**Files:** `RedSalamander/CompareDirectoriesEngine.h/.cpp`

- Content-compare cache: replace “clear everything at 16K” with partial eviction.
- Decision cache: add an approximate byte-estimate and LRU eviction to a ~300MB budget.
- Pin the currently visible folder decisions (and ancestors) so the UI remains stable.

### D) S3 plugin: timeout-bounded, reusable client resources

**Files:** `Plugins/FileSystemS3/FileSystemS3.Shared.cpp`, `Plugins/FileSystemS3/FileSystemS3.*`

- Set `requestTimeoutMs = 30'000`.
- Reuse `S3CrtClient` per instance/context (avoid per-call client churn).
- Important: when caching/reusing `Aws::S3Crt::S3CrtClient`, construct it directly in a `std::shared_ptr` (do not copy/move a client object into the cache). Copying can leave the internal CRT client in an invalid state and trigger `aws_fatal_assert` on first request.
- `S3RangedFileReader`: use a shared client and increase range chunk size (8MiB) to reduce request count.

### E) Instrumentation + regressions

**Files:** `RedSalamander/CompareDirectoriesEngine.h/.cpp`, `RedSalamander/CompareDirectoriesEngine.SelfTest.cpp`, `RedSalamander/CompareDirectoriesWindow.cpp`

- Add a perf stats struct (`GetPerfStats`) including `DirectoryInfoCache::Stats` snapshot and high-water marks.
- Add selftests proving:
  - compare scan does not materially grow `DirectoryInfoCache`,
  - hi/lo queues stay within caps,
  - decision cache stays within budget and pinned folders survive eviction,
  - cancel completes without deadlock.

## Verification gates

1. Build: `.\build.ps1 -ProjectName RedSalamander`
2. Build: `.\build.ps1`
3. Test: `RedSalamander.exe --compare-selftest`
4. Manual: S3↔FS subtree compare remains interactive, memory is bounded, cancel/close exits promptly.
