# Compare Directories Code Review

This document captures correctness/performance risks spotted while implementing the Compare Directories feature (`CompareDirectories*`).
It focuses on:
- Correctness of name matching (case-folding) and cache behavior
- Responsiveness (instant-ish UI feedback; avoid blocking enumeration/UI thread)
- Cross-filesystem support (non-Win32 filesystem plugins)
- Robustness on large/deep directory trees

Note: file/line references are approximate and may drift as the code changes.

## CompareDirectoriesEngine.h / .cpp

### 1. Hash/equality mismatch violates `std::unordered_map` contract (correctness bug)

**Location:** `CompareDirectoriesEngine.h:46`

`WStringViewNoCaseHash` hashes by applying `std::towlower`, while `WStringViewNoCaseEq` compares with `CompareStringOrdinal(..., TRUE)` (ordinal case-insensitive compare).
These two definitions can disagree for some characters, which breaks the requirement that “if `eq(a,b)` then `hash(a)==hash(b)`”.
That can lead to wrong lookups/insertions, not just extra collisions.

Recommendations:
- Prefer switching to an ordered container (e.g. `std::map`) with a comparator using `CompareStringOrdinal(..., TRUE)`.
- Or normalize keys with a Windows ordinal-compatible fold and use a case-sensitive hash/eq on the normalized key.

### 2. Recursive `GetOrComputeDecision` can stack overflow on deep hierarchies

**Location:** `CompareDirectoriesEngine.cpp:1043`

When `compareSubdirectories` is enabled, `GetOrComputeDecision` recursively calls itself for each subdirectory. A deeply nested folder tree (e.g., `node_modules`) can blow the stack.

Recommendations:
- Use an iterative worklist (queue/stack) rather than recursion.
- If recursion remains, add a maximum depth guard and treat “too deep” as “unknown / needs user action” rather than crashing.

### 3. Content comparison needs a two-phase/asynchronous strategy for responsiveness

**Location:** `CompareDirectoriesEngine.cpp:583-670`

`AreFilesEqualContent` can be expensive (large files, slow I/O, remote/S3). Doing this synchronously blocks enumeration and makes the UI feel “hung” when “show identical/different” is enabled.

Recommendations:
- Two-phase compare:
  - Phase 1 (fast): enumerate + compare metadata (existence/type/size/date/attrs) and show results immediately.
  - Phase 2 (slow): compute content equality in background for candidates that need it; update UI incrementally as results arrive.
- Cache content-compare results per (path, size, lastWriteTime, attributes) and discard work when `_version` changes.
- Add an explicit cancellation check (e.g., compare a captured `_version` against current `_version`) inside the read loop.
- Ensure the UI gets a final “decision updated” notification when the last background compare completes (throttling must not drop the final update), otherwise items can remain stuck showing “Comparing...”.

### 4. Content compare must be filesystem-agnostic (don’t use Win32 file APIs directly)

**Location:** `CompareDirectoriesEngine.cpp:AreFilesEqualContent`

Compare needs to work for non-Win32 filesystem plugins. If `AreFilesEqualContent` uses `CreateFileW`/`ReadFile` on a `std::filesystem::path`, it will fail for virtual paths (e.g., `s3:`) and other plugin-backed filesystems.

Recommendations:
- Always go through `IFileSystemIO::CreateFileReader` for both sides.
- Keep Win32 direct file I/O as an optional fast-path only when the filesystem explicitly indicates “local Win32 path” support.

### 5. Content compare doesn’t handle “size unknown” robustly

**Location:** `CompareDirectoriesEngine.cpp:1070-1081`

When `sizeDifferent` is false, the code passes `item.leftSizeBytes` to `AreFilesEqualContent`. If both sides enumerate as size 0 (sparse files, or plugin doesn’t report sizes), `AreFilesEqualContent` can return “equal” without reading meaningful data.

Recommendations:
- Treat size as “unknown” when the plugin doesn’t provide it; make the size parameter optional.
- If size is unknown, either:
  - fall back to streaming until EOF, or
  - classify content as “unknown” and schedule background compare (see #3).

### 6. No cancellation mechanism for stale scans (wasted I/O)

**Location:** `CompareDirectoriesEngine.cpp:884-1161`

When the user changes settings or roots, `_version` is bumped and the cache is cleared, but any in-progress scan continues doing work to completion. The stale result is discarded at the end, but the wasted I/O can be significant.

Recommendations:
- Capture the starting `_version` at the beginning of a scan and check it periodically:
  - before recursing into a directory
  - before/after expensive content compares
  - within long I/O loops (read loop)
- Bail out early when the version changes.

### 7. `InvalidateForRelativePathLocked` subtree invalidation is O(n)

**Location:** `CompareDirectoriesEngine.cpp:313-325`

The subtree invalidation iterates all cache entries to find prefix matches. For large scanned trees, this is O(n).

Recommendations:
- Consider using an ordered structure + prefix range (`lower_bound`) for O(log n + k) invalidation.
- Or store per-directory child relationships so invalidation can walk only the impacted subtree.

### 8. Root cache key sentinel (resolved)

**Location:** `CompareDirectoriesEngine.cpp:MakeCacheKey`

The root folder mapping should use a non-empty, explicit sentinel (e.g. `L"."`) rather than `""` to avoid ambiguity.

Update:
- Root cache key now uses `L"."`.

---

## CompareDirectoriesWindow.cpp

### 9. Global mutable vector `g_compareDirectoriesWindows`

**Location:** `CompareDirectoriesWindow.cpp:47`

This is a global `std::vector<CompareDirectoriesWindow*>` for theme updates. It's not thread-safe and relies on all windows being created/destroyed on the UI thread. Consider using a more robust registration pattern (e.g., a message broadcast, or an `std::set` guarded by a comment about UI-thread-only access).

### 10. Hardcoded strings "Compare Folder", "Options...", "Rescan"

**Location:** `CompareDirectoriesWindow.cpp:1985, 1998, 2012`

The CLAUDE.md explicitly says "No Hardcoded Strings - UI strings go in `.rc` resources." These banner strings should use `LoadStringResource` like the rest of the UI.

### 11. `OnNcDestroy` calls `delete this`

**Location:** `CompareDirectoriesWindow.cpp:836`

The window deletes itself in `OnNcDestroy`. This pattern is valid but fragile - if any code holds a raw pointer to the window after destruction (e.g., a callback fires late), it's a use-after-free. The scan progress callback captures `HWND` (not `this`), which is good, but the `FolderWindow` callbacks at lines 2076-2102 capture `this` directly. If a callback fires during or after destruction, it's UB. Consider `weak_ptr`-based guard or ensure callbacks are explicitly unregistered before destruction.

### 11b. Posted-message payload registry must not race with message delivery

Compare uses `PostMessagePayload(...)` for scan/content progress. If a payload pointer is posted to the UI thread and the UI thread handles the message before the payload is registered for draining/unregister, the pointer can be deleted once by the receiver and then deleted again during `WM_NCDESTROY` draining (stale pointer / double-delete).

Recommendations:
- Ensure payload registration happens before the message can be processed (e.g., register under a lock before `PostMessageW`, and roll back registration on `PostMessageW` failure).
- Keep payloads strictly owning (no `string_view`/`path_view` fields) at the post boundary.

### 12. Options dialog lambda is ~250 lines of deeply nested code

**Location:** `CompareDirectoriesWindow.cpp:2107-2358`

The `CreateDialogParamW` callback is an enormous inline lambda. This is hard to read, test, and maintain. Extract it into a named static function (e.g., `OptionsDlgProc`).

### 13. `ShowCompareDirectoriesWindow` raw `new` without smart pointer

**Location:** `CompareDirectoriesWindow.cpp:3886-3891`

The window is allocated with `new` and if `Create` fails, manually `delete`d. On success, the window self-destructs in `OnNcDestroy`. This pattern works but violates the project's "no raw new/delete" rule. Consider `std::unique_ptr` with a release on success:
```cpp
auto window = std::make_unique<CompareDirectoriesWindow>(...);
if (!window->Create(owner)) return false;
window.release(); // ownership transferred to HWND lifecycle
return true;
```

### 14. `SetSplitRatio` invalidates entire window

**Location:** `CompareDirectoriesWindow.cpp:1773`

`InvalidateRect(_hWnd.get(), nullptr, TRUE)` invalidates the whole window on every splitter drag movement. This causes unnecessary repainting of both FolderViews. Only invalidate the splitter region and the areas that actually change size.

### 15. Verbose RAII handle declarations

**Location:** `CompareDirectoriesWindow.cpp:548-560`

The project has 13+ handles declared with the full `wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject>` pattern. Consider a type alias:
```cpp
using wil::unique_hbrush = wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject>;
using wil::unique_hfont  = wil::unique_any<HFONT,  decltype(&::DeleteObject), ::DeleteObject>;
```

### UX follow-up (Feb 2026)

- Options dialog **OK** should apply settings and trigger a rescan (same intent as banner **Rescan**).
- **Compare scope is explicit**: the left/right compare roots are only established/changed by **OK** (options) or banner **Rescan**; navigation must not implicitly change roots.
- While compare scope is active:
  - Navigating within either pane keeps both panes synchronized by **relative path** under their roots.
  - If the corresponding folder does not exist on one side, the other side still navigates; the missing side should behave like an empty folder (no hard error).
- When navigation leaves the compare scope (pane path is outside its root):
  - Cancel compare mode: stop synchronization and treat both panes as independent navigation.
  - Clear/hide scan status/progress and stop scheduling background compares for the old scope.
  - A subsequent **OK** or **Rescan** re-establishes scope using the panes' current folders as the new roots.
- Show a scan-in-progress indicator in the top banner that disappears when enumeration and background content compares complete.
- Surface progress similar to the File Operations dialog: total entries explored + per-file content-compare read progress.

---

## CompareDirectoriesEngine.SelfTest.cpp

### 16. No test for reparse point (symlink/junction) skipping

The engine has logic to skip reparse points during subdirectory comparison (`CompareDirectoriesEngine.cpp:1038-1040`), but there's no self-test covering this. On Windows, creating junctions in tests is feasible with `CreateSymbolicLinkW`.

### 17. No test for concurrent access or version invalidation

The engine is designed for multi-threaded use (mutex, atomics), but the self-tests are all single-threaded. Consider a test that calls `Invalidate()` while `GetOrComputeDecision` is running to verify the version check at line 1154 discards stale results.

### 18. Missing “acceptance-style” tests for responsiveness and cross-filesystem behavior

Consider adding self-tests that exercise the UX-critical scenarios:
- Deep trees (thousands of directories; very deep nesting) to ensure no stack overflow and acceptable performance.
- “Show identical/different” enabled on:
  - many small files (stress scheduler / cache)
  - a few large files (stress I/O and cancellation)
- Version invalidation mid-scan (change options or roots while content compare is running) to ensure work cancels quickly.
- Unknown-size items (plugin returns size=0/unknown) to ensure the engine doesn’t incorrectly claim “identical”.
- Non-Win32 filesystem paths (e.g. Dummy FS) to ensure content compare goes through `CreateFileReader` and doesn’t accidentally call Win32 APIs.

---

## Summary of priorities

| Priority | Issue | Location |
|----------|-------|----------|
| **High** | Hash/equality mismatch violates `unordered_map` contract | Engine.h:46 |
| **High** | Content compare should be two-phase + async (UI responsiveness) | Engine.cpp:AreFilesEqualContent |
| **High** | Content compare must use filesystem `CreateFileReader` (cross-FS) | Engine.cpp:AreFilesEqualContent |
| **High** | Unbounded recursion depth on deep trees | Engine.cpp:1043 |
| **High** | Hardcoded UI strings in banner | Window.cpp:1985-2012 |
| **Medium** | No cancellation for stale scans (wasted I/O) | Engine.cpp:884+ |
| **Medium** | “Size unknown” can lead to false “identical” | Engine.cpp:1070 |
| **Medium** | `this`-capturing callbacks risk use-after-free | Window.cpp:2076+ |
| **Medium** | 250-line inline lambda for dialog proc | Window.cpp:2107+ |
| **Medium** | Raw `new`/`delete` window ownership pattern | Window.cpp:3886 |
| **Low** | O(n) subtree invalidation | Engine.cpp:313 |
| **Low** | Missing self-tests (reparse/cancel/deep tree/cross-FS) | SelfTest.cpp |
| **Low** | `InvalidateRect` for entire window during splitter drag | Window.cpp:1773 |
