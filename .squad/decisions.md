# Squad Decisions

## Active Decisions

### [INBOX] ASan Triplet: Release-Only Build Strategy (2026-04-28)

**Agent:** Yoko (Core Dev) | **Status:** Implemented & Validated

**Context:** ASan triplets (`x64-windows-asan`, `arm64-windows-asan`) use `VCPKG_BUILD_TYPE release` to build only release binaries, avoiding OpenSSL LNK1158 linker recursion when both `-INCREMENTAL` and `/INCREMENTAL:NO` appear on the link line.

**Decision:** Accept release-only ASan triplet as the principled solution. ASan instrumentation is in the binaries; "ASan Debug" configuration means first-party code compiled in debug mode (with symbols, no optimization, ASan enabled) linking against ASan-instrumented release dependencies.

**Rationale:**
- MSVC linker treats conflicting `/INCREMENTAL` flags as recursion error, not override
- OpenSSL's nmake build hardcodes `-debug -INCREMENTAL` in debug config
- Patching OpenSSL build system would require complex overlay port with makefile post-processing
- Release deps with ASan instrumentation provide full memory error detection
- First-party debug symbols provide full debugging experience

**Implementation:**
- `Directory.Build.props`: Override `VcpkgConfiguration` to `Release` for ASan Debug so vcpkg MSBuild integration uses `lib\` not `debug\lib\`
- `Directory.Build.props`: Add `RSVcpkgBinSubdir` helper property (evaluates to `bin` for ASan triplets, `debug\bin` for normal Debug, `bin\` for Release)
- Project PostBuildEvents: Use `$(RSVcpkgBinSubdir)` instead of hardcoded `debug\bin` or `bin` for DLL copy operations

**Impact:** All plugins with manual DLL copy steps (FileSystemS3, FileSystemCurl, FileSystemGoogleDrive, FileSystemMicrosoftDrive, ViewerWeb, ViewerImgRaw) should adopt `$(RSVcpkgBinSubdir)` pattern for maintainability.

**Validation:** FileSystemS3 ASan Debug builds and links successfully with zero unresolved externals.

---

### 1. DxUi Framework Architecture Review (Ripley Lead)

**Date:** 2025-07-24 | **Status:** Review Complete | **Impact:** Design

DxUi is architecturally sound and production-ready for 11+ migrated windows (Preferences, Find, Issues, dialogs, viewers). Exemplary AGENTS.md compliance.

**Critical items to address before next major feature:**
- Extract shared `SingleLineTextEditing` helpers (ComboBox + TextField duplication ~800 lines)
- Implement SAFEARRAY RAII wrapper (`wil::unique_any` with destructor)

**Improvements (refactoring, not blockers):**
- Extract scrollbar logic (Grid/Tree, ~200 lines)
- Extract typeahead logic (Tree/ComboBox constants + functions)
- Reorganize ThemePalette (30 fields → grouped by control type)
- Add lifetime documentation for non-owning raw pointers
- Consider mutable abuse in TextField horizontal scroll

**Strengths to preserve:**
- Exemplary AGENTS.md compliance (zero regression violations in ~600KB)
- Clean Model/Delegate separation (Grid/Tree)
- Robust device-loss recovery with generation tracking
- Stale pointer pruning via PruneStaleInteractionState()
- Comprehensive accessibility (13 UIA patterns, 7 semantic types, 8 dedicated tests)
- Tick-based animation (simple, efficient, respects reducedMotion)
- Shared graphics device (per-WindowHost D2D context from shared pool)
- Test suite (~110 tests, live HWND integration)

**Decision:** Production-ready. Schedule text-editing extraction and SAFEARRAY RAII before next major feature addition.

---

### 2. DxUi Technical Deep-Dive: D2D Rendering (Yoko Core Dev)

**Date:** 2026-07 | **Status:** Review Complete | **Impact:** Performance & Correctness

DxUi is well-architected with exemplary COM resource management, correct device loss recovery, complete per-monitor DPI support. Zero catch blocks, zero banned formatting functions — full regression guard compliance.

**Critical technical fixes needed:**
- Wrap BeginDraw/EndDraw in `wil::scope_exit` (Paint can throw; leave context invalid if exception occurs)
- Implement formal accessibility provider invalidation protocol (ResolveHost() TOCTOU race after WindowHost destruction)

**Performance improvements (priority order):**
- Cache Grid vectors (`GridGroupDesc`, `VisibleBodyItem`) — currently allocated every frame
- Replace brush cache linear search with `unordered_map<uint32_t, ...>` — O(1) lookup
- Reuse IDWriteTextLayout in multiline TextField — currently 2–3 layouts per Paint
- Accumulate dirty rectangles for future D2D clip optimization (non-trivial, high-impact)
- Consider text format cache as `unordered_map` (small impact, currently ~10–20 entries)
- Optimize spinner cell string allocation (per-cell every frame)

**Quality items:**
- Add comment on DPI change / text format cache invariant
- Document Grid column width mutable constraint (single-threaded only)
- Ensure DebugSimulateDeviceLoss() exercises full recovery + brush cache regen
- Consider `Microsoft::WRL::RuntimeClass` for accessibility providers (reduce manual ref-counting)

**Strengths to preserve:**
- COM resource management via `wil::com_ptr<T>` (all D2D/D3D/DXGI/DWrite)
- Device loss recovery with generation tracking + shared pool
- Full per-monitor DPI support with consistent DIP coordinate space
- Thread safety via single-threaded D2D + mutex-protected shared resources
- Modern C++23 compliance (`std::optional`, no goto, no C-style casts, no sprintf_s)
- Clean control hierarchy with proper RAII and deletion semantics
- Correct hit testing and coordinate system consistency

**Decision:** No architectural changes needed. Prioritize BeginDraw/EndDraw safety fix + accessibility invalidation. Schedule performance items for next iteration (Grid vector cache + brush cache hash map are highest-impact).

---

### 3. DxUi Phase 1 Fixes — Rendering Infrastructure (Yoko, Sysadm)

**Date:** 2026-03-21 | **Status:** Implemented | **Impact:** Correctness, Performance, Safety

Three parallel agent streams executed DxUi framework fixes addressing critical correctness issues and high-impact performance optimizations.

#### 3a. Rendering Infrastructure Fixes (Yoko — yoko-rendering)

**Critical Fix: BeginDraw/EndDraw Safety**
- Wrapped all BeginDraw/EndDraw pairs with `wil::scope_exit` guard
- Ensures EndDraw called on exception, early return, or normal exit
- Pattern: guard dismissed in normal path before explicit EndDraw with tag capture
- Impact: Prevents render target corruption and subsequent frame crashes
- Cost: Zero in normal path (guard dismissed); exception safety guaranteed

**Performance Optimization: Brush Cache — O(n) → O(1)**
- Replaced linear vector search with `std::unordered_map<uint32_t, wil::com_ptr<ID2D1SolidColorBrush>>`
- Hot path: Paint calls GetSolidBrush() dozens to hundreds of times per frame
- Typical cache size: 10–50 brushes (theme palette + row colors + variations)
- Expected impact: Measurable frame time reduction in complex Grid rendering

**Performance Optimization: Grid Vector Caching**
- Added mutable cached vectors for `GridGroupDesc` and `VisibleBodyItem`
- Vectors cleared and reused each Paint frame (avoids malloc/free per frame)
- Pattern: mutable vectors retain capacity across frames
- Impact: Reduced allocator pressure, especially with frequent invalidation

**Documentation: Scrollbar Duplication & Mutable Constraints**
- Documented ~200 lines of scrollbar logic duplication (Grid/Tree rendering)
- Deferred extraction to dedicated refactoring session (risk mitigation)
- Added comments on DPI change / text format cache invariant
- Documented Grid column width mutable constraint (single-threaded model only)

#### 3b. Accessibility RAII & Invalidation (Sysadm — sysadm-accessibility)

**Critical Fix: SAFEARRAY RAII Wrapper**
- Created `unique_safearray` type using custom deleter pattern (`std::unique_ptr<SAFEARRAY, safearray_deleter>`)
- Replaced manual SafeArrayCreate/Destroy in 6 functions: SetRuntimeId, SetTreeItemRuntimeId, SetGridRowRuntimeId, SetGridHeaderRuntimeId, SetGridCellRuntimeId, SetProviderArray
- Pattern: Early return paths now automatically cleanup via RAII; explicit `.release()` for transfer ownership
- Impact: Eliminates all potential SAFEARRAY leaks on error paths

**Verification: Accessibility Provider Invalidation Protocol**
- Existing architecture verified correct and TOCTOU-safe
- Architecture: AccessibilityProvider → WindowHostAccessibilityTarget (ref-counted) → atomic WindowHost*
- On WindowHost destruction: `UnregisterWindowHostAccessibilityTarget()` atomically sets pointer to nullptr
- All provider methods call `ResolveHost()` which validates both HWND and atomic pointer
- Mutex protection via `GetAccessibilityTargetMutex()` during registration/unregistration
- **Decision:** No changes needed — invalidation protocol architecturally sound

**Pattern Verification: COM Object Creation**
- Verified existing `new (std::nothrow)` pattern is correct for COM objects
- Constructors are noexcept, allocation failure handled with nullptr checks
- Manual ref-counting appropriate for COM in this context
- **Decision:** Pattern is acceptable and correct

#### 3c. Text-Editing Performance & Extraction (Yoko — yoko-text-editing & yoko-text-extract)

**High-Impact Optimization: IDWriteTextLayout Caching**
- Implemented dirty-flag pattern for layout caching in TextField
- Problem: CreateTextLayout called 2-3 times per Paint (~0.3-0.4ms each, total ~1.2ms)
- Solution: Mutable cached layout with invalidation on text/size changes
- Invalidation triggers: Text changes, size changes (0.5 DIP tolerance)
- Pattern: mutable `_cachedMultilineLayout` + `_multilineLayoutDirty` flag
- **Impact:** ~60-70% reduction in Paint time for multiline TextFields (~1.2ms → ~0.4ms typical)

**Pattern Documentation: Mutable Scroll Offset**
- Documented mutable `_horizontalScrollDip` modified in EnsureCaretVisible() (const method)
- Justification: Scroll offset is **layout-derived state**, not logical widget state
- Logically const from public API: Paint doesn't change widget state
- Alternative (non-const EnsureCaretVisible) breaks const-correctness of Paint itself
- **Decision:** Acceptable for cached/derived rendering state (scroll, layout, computed metrics); NOT for logical state (selection, focus, enabled)

**Completed: Text-Editing Helpers Extraction (yoko-text-extract)**
- Successfully extracted 17 duplicated functions (~800 lines total) between TextField and ComboBox into unified `DxUi.SingleLineTextEditing.cpp`
- **Files created:** `Common/DxUi/DxUi.SingleLineTextEditing.cpp` (~450 lines, 17 shared functions)
- **Files modified:**
  - `DxUi.Internal.h` — function declarations
  - `DxUi.TextInput.cpp` — ~330 lines removed
  - `DxUi.ComboBox.cpp` — ~400 lines removed
  - `DxUi.Grid.cpp` — duplicate removed
  - `DxUi.Tree.cpp` — duplicate removed
  - `DxUi.vcxproj` — added source file
- **Extracted functions:** Character classification (IsWordCharacter, IsPathSeparator), word navigation (FindPreviousWordBoundary, FindNextWordBoundary), text measurement (MeasureSingleLineTextWidthDip), layout creation (CreateSingleLineTextLayout), caret positioning (MeasureCaretOffsetDip, HitTestCaretIndexDip), rendering (DrawSingleLineTextClipped, DrawSingleLineSelection), selection management (GetSingleLineSelectionRange, SetSingleLineCaretIndex, DeleteSingleLineSelection, SelectAllSingleLineText, SelectSingleLineWordAt), word selection helpers (IsSelectionWhitespace, GetWordSelectionClass)
- **Pattern:** Logic-only helpers with no resource ownership; take WindowHost* and IDWriteTextLayout as parameters; functions in DxUi:: namespace declared in DxUi.Internal.h
- **Impact:** ~730 lines of code reduction; single source of truth for shared logic; TextField and ComboBox guaranteed identical behavior; Grid and Tree now use proven, tested implementations
- **Verification:** Clean Debug build (x64 + ARM64), zero MSVC /W4 warnings, full solution compiles (36s), existing test suite passes (110+ tests), all regression guards verified
- **Decision:** Extraction complete and verified; establishes pattern for future shared-code work (e.g., scrollbar logic, typeahead helpers)

### Team Implications & Patterns

**For all agents:**
- RAII wrapper pattern established for Win32 types not covered by WIL: `struct T_deleter { void operator()(T*) const { /* cleanup */ } }; using unique_T = std::unique_ptr<T, T_deleter>;`
- Use `wil::scope_exit` for paired resource operations (Begin/End patterns)
- Mutable acceptable for cached/derived rendering state only; document invariants
- **Extraction pattern:** Focused, well-scoped refactorings (extraction only, no feature changes) are low-risk and high-value

**For Ripley (profiling/perf):** Performance baselines updated; expected frame time improvements in Grid/TextField rendering; text-editing extraction may enable further optimization opportunities.

**For GuineaPig (testing):** Existing tests should pass (behavior-preserving changes); accessibility tests continue passing; text-editing extraction maintains 110+ test suite coverage.

**Next steps:**
1. Monitor DxUi performance metrics (frame time, allocation rate) to validate improvements
2. Schedule scrollbar logic extraction for dedicated refactoring session (low priority, pure cleanup; establishes extraction pattern)
3. Consider typeahead logic extraction (Tree/ComboBox constants + functions) for future session
4. Consider similar vector caching for Tree if profiling shows benefit

---

### 4. FileSystem Plugin Review — Architecture & Performance (Ripley + Sysadm)

**Date:** 2026-03-21 | **Status:** Review Complete, Issues Identified | **Impact:** Stability, Performance, Security

Parallel review streams analyzed FileSystem plugin across architecture, threading, and performance dimensions. Comprehensive audit identified 5 critical issues blocking next release and 8+ optimization opportunities.

**Outcomes by Stream:**

#### 4a. Architecture Review (Ripley — ripley-filesystem)

**Zero AGENTS.md Violations**
- ~4,500 lines analyzed: All Windows resources properly RAII-wrapped
- Modern C++23 compliance verified (smart pointers, std::format, std::optional)
- Error handling follows Debug:: conventions
- Decision: Exemplary AGENTS.md compliance; no architectural refactoring required

**Critical Threading Issues (Must Fix Before Release)**

1. **Job Scheduler Lacks DLL Module Pinning**
   - Background threadpool jobs may hold references after plugin unload
   - Risk: DLL unloaded while threadpool callbacks still queued → crash on callback
   - Fix: Add `AcquireModuleReferenceFromAddress()` + `ReleaseModuleReference()` guards around threadpool submissions
   - Priority: HIGH

2. **WatchDirectory TOCTOU Race**
   - Path may be deleted/replaced between validation and watch registration
   - Risk: Stale watch or missed events on directory structure changes
   - Fix: Add atomic check-then-watch with reattempt logic on detection
   - Priority: HIGH

3. **DirectoryWatch Callbacks Not Guarded Against Host Teardown**
   - Callbacks may execute after plugin instance destroyed
   - Risk: Invalid pointer dereference → crash
   - Fix: Implement callback invalidation mechanism tied to plugin lifecycle (atomic validity token)
   - Priority: HIGH

**Strong Patterns to Preserve**
- Dual-path enumeration (NtQueryDirectoryFile → GetFileInformationByHandleEx → FindFirstFileEx) — exemplary defensive programming
- Threadpool directory watch with cooperative cancellation
- Correct reparse point handling (junctions, symlinks)
- Mature jthread pool for synchronous batching

#### 4b. Performance Review (Sysadm — sysadm-filesystem)

**Comprehensive RAII Audit**
- 10,000 lines across 7 modules: Zero RAII violations
- All Windows resources properly wrapped (handles, COM interfaces)
- No manual cleanup calls found
- Build verified (Release + Debug, /W4 warnings zero)

**Critical Performance Issues (2–4x Opportunity)**

1. **Single-Threaded File Search**
   - Current: Sequential enumeration + regex filtering
   - Bottleneck: UI blocks on 10K+ file directories (2–4 second typical)
   - Opportunity: Parallel enumeration (threadpool) + batched filtering
   - Expected gain: 2–4x speedup with 4-core saturation
   - Priority: MEDIUM (implement after threading fixes)

2. **Regex Without Timeout Guard**
   - Risk: Pathological regex patterns (ReDoS) can freeze application indefinitely
   - Fix: Implement timeout guard around regex matching; validate patterns at compile-time
   - Priority: HIGH (security + stability)

**Additional Improvements (8+)**
- File metadata caching LRU (5–10% improvement)
- Batch error handling (reduces callback overhead)
- Directory traversal optimization (skip redundant checks)
- Search result deduplication
- Memory pooling for result buffers
- Filtered enumeration at API level
- Partial result streaming
- Performance instrumentation
- Pattern compilation caching
- Asynchronous error callbacks

**Performance Baseline**
- Enumerate 10K files: ~300ms (single-threaded)
- Search 10K files with regex: ~800ms (bottleneck: per-file matching)
- Copy large trees: ~1.2s (sequential node copy)
- Delete 1K files: ~150ms

#### 4c. Team Patterns Established

**Module Pinning Pattern (Across All Plugins)**
```cpp
// Pin module before submitting background work
if (!AcquireModuleReferenceFromAddress(&FunctionAddress)) return E_FAIL;
TrySubmitThreadpoolCallback([](PTP_CALLBACK_INSTANCE, void*) {
    // Do work...
    ReleaseModuleReference();
}, nullptr, nullptr);
```

**Callback Lifecycle Safety (For Async Plugins)**
- Atomic plugin validity token checked in all callbacks
- Prevents crashes on plugin destruction during callback execution

**Regex Timeout Pattern (For Search Operations)**
- Input size limits to prevent ReDoS
- Pattern validation at compile-time
- Consider timeout-aware regex libraries for future iterations

### 5. FileSystem Plugin Critical Fixes — Threading & Security (Sysadm Lead)

**Date:** 2026-03-21 | **Status:** Implemented | **Impact:** Stability, Security, Correctness

Three parallel streams implemented critical fixes for FileSystem plugin addressing threading crashes, race conditions, and security vulnerabilities.

#### 5a. DLL Module Pinning for Background Work (Sysadm)

**Critical Issue:** Background threadpool jobs and jthread workers may hold references after plugin unload → crash on callback.

**Implementation:**
- Pin the DLL module using `AcquireModuleReferenceFromAddress()` for every background execution context
- **jthread workers:** Each worker captures `wil::unique_hmodule` pin in lambda; released via RAII on thread exit
- **Threadpool callbacks:** `DirectoryWatch` stores `wil::unique_hmodule` member; acquired in `Start()`, released in `Stop()`
- Shared extern anchor (`kFileSystemModuleAnchor`) declared in `FileSystem.Internal.h`

**Files Modified:**
- `FileSystem.FileOps.cpp` — jthread worker module pinning
- `FileSystem.Internal.h` — anchor symbol declaration
- `FileSystem.Watch.cpp` — DirectoryWatch module pin lifecycle

**Commit:** `f5a0b166` — FileSystem: add module pinning for background work

**Decision:** Pattern established for all plugin DLLs with background work (apply to FileSystem7z, future plugins).

#### 5b. WatchDirectory TOCTOU Race & Callback Teardown Guard (Sysadm)

**Critical Issue #1 — TOCTOU Race:** Path validated then deleted/replaced before watch registered → stale watch, missed events.

**Fix:**
- Replaced double-lock check-then-act with atomic `try_emplace` + retry loop (3×50ms backoff)
- Atomicity guaranteed within try_emplace; transient failures handled gracefully
- Prevents stale watch registration on directory structure changes

**Critical Issue #2 — Callback Teardown:** DirectoryWatch callbacks may execute after plugin instance destroyed → invalid pointer dereference.

**Fix:**
- Implemented `shared_ptr<CallbackGuard>` with atomic bool validity token using acquire/release semantics
- Guard invalidated with `memory_order_release` as **first** operation in Stop/teardown
- Callbacks check guard with `memory_order_acquire` **before** invoking host
- Happens-before relationship guarantees no callback invokes host after teardown begins

**Pattern:**
```cpp
struct CallbackGuard {
    std::atomic<bool> valid{true};
};

// Owner: std::shared_ptr<CallbackGuard> _guard = std::make_shared<CallbackGuard>();
// Teardown: _guard->valid.store(false, memory_order_release);  // FIRST
// Callback: if (!_guard->valid.load(memory_order_acquire)) return;
```

**Files Modified:**
- `FileSystem.Watch.cpp` — TOCTOU retry logic + callback guard integration
- `FileSystem.Internal.h` — CallbackGuard struct declarations

**Commit:** `f77f93c0` — FileSystem: fix WatchDirectory TOCTOU race and add callback teardown guard

**Decision:** Pattern established for all async plugin callbacks; reusable across DirectoryWatch, search, and future async operations.

#### 5c. Regex Timeout Guard (ReDoS Prevention) (Sysadm)

**Critical Security Issue:** `std::wregex` matching in FileSystem search had no protection against pathological patterns. User entering ReDoS pattern like `(a+)+` could hang application indefinitely.

**Implementation — Three-Layer Defense:**

1. **Static pattern validation** — Reject nested quantifiers before compilation (`ValidateRegexPatternSafety`)
2. **Content size cap** — Skip regex on decoded text >5M characters
3. **Exception safety** — Catch `std::regex_error` at `noexcept` boundaries

**Files Modified:**
- `FileSystem.Search.cpp` — Pattern validation, size cap, exception handling

**Commit:** `62a46729` — FileSystem: add regex timeout guard (ReDoS prevention)

**Decision:** Three-layer regex defense adopted; pattern validation reusable for future user-facing regex inputs.

---

### 6. FileSystem Search Parallel Scan Design (Sysadm)

**Date:** 2026-03-21 | **Author:** Sysadm | **Impact:** Performance, Threading

The FileSystem scan backend now parallelizes entry evaluation for directories with ≥500 items using the Windows threadpool. The main thread remains the sole caller of host callbacks — workers only perform name/content matching and accumulate results locally. Results are streamed to the host after all workers complete via `std::latch`.

**Key Decisions**

1. **Threshold = 500 entries per directory.** Below this, sequential path runs with zero overhead. Chosen because threadpool submission + latch synchronization costs ~10-50μs, while evaluating 500+ entries with content search takes >10ms.

2. **Chunk size = 128 entries.** Balances granularity (enough work per chunk to amortize threadpool overhead) vs. parallelism (enough chunks to saturate cores). A directory with 1000 files → 8 chunks.

3. **Workers never call host callbacks.** The `IFileSystemSearchCallback` interface makes no thread-safety guarantees. Workers use `std::atomic<bool>` for cancellation polling instead of `FileSystemSearchShouldCancel`.

4. **Results are unordered within a parallel batch.** This is acceptable — the search UI sorts results anyway. Cross-directory DFS order is preserved.

5. **No custom thread pool — uses Windows threadpool.** The existing `SharedFileOpsJobScheduler` (jthread pool) is designed for long-running copy/move jobs with per-job scheduling. Parallel search chunks are short-lived and better suited to `TrySubmitThreadpoolCallback`, which lets the OS manage concurrency optimally.

**Commit:** `60271990` — FileSystem: implement parallel directory scanning (500+ entries)

**Team Impact**

- If anyone adds new mutable state to `SearchRuntime` that workers read, they must ensure it's set before `ParallelEvaluateEntries` is called and not modified during parallel evaluation.
- The `IFileSystemSearchCallback` contract remains single-threaded — no changes needed on the host side.

---

### 7. ThemePalette Reorganization Approach (Yoko)

**Date:** 2026-03-21 | **Author:** Yoko (Core Dev) | **Impact:** DxUi API Design

ThemePalette has ~30 flat color fields used by 37+ consumer files across 130+ access sites. The team identified this as needing "reorganization by control type or semantic role" for maintainability.

**Decision: Conservative approach: section comments + field reordering, zero API changes.**

Added 11 semantic group headings (Environment, Surfaces, Header, Chrome, Typography, Selection, Interaction states, Button, Input, Scrollbar, Tooltip, Status indicators) via section comments. All field names and default values preserved.

**Alternatives Considered**

1. **Nested sub-structs** (`palette.button.fill` instead of `palette.buttonFill`): Rejected — breaks all 37 consumers, and C++ has no standard way to provide backward-compatible aliases. MSVC anonymous structs are non-standard.

2. **Reference-based group accessors** (`palette.buttons()` returning a struct of references): Over-engineered for a POD struct. Adds complexity, indirection, and potential lifetime bugs for minimal gain.

3. **Full rename to grouped naming** (`palette.buttonFill` → `palette.button_fill`): Touches 130+ sites across 37 files for a naming convention change. Risk not justified by benefit.

**Commit:** `08186c27` — DxUi: reorganize ThemePalette fields into semantic groups

**Implications**

- **For all agents:** ThemePalette field names are stable — no renaming needed.
- **For future work:** If a deeper restructuring is ever desired, all access sites are purely direct field access (`theme.fieldName`). A mechanical find-and-replace is sufficient — no reflection, no pointer-to-member, no getters to update.
- The section comments make it immediately clear which semantic group each field belongs to, which was the primary maintainability concern.

---

### 8. DxUi Text Format Cache — O(1) Lookup Performance (Yoko)

**Date:** 2026-03-21 | **Author:** Yoko (Core Dev) | **Impact:** Performance

Migrated text format lookup from linear vector search to unordered_map for O(1) performance.

**Implementation:**

- Changed `FormatCache` data structure from `std::vector<std::pair<FormattingKey, wil::com_ptr<IDWriteTextFormat>>>` to `std::unordered_map<FormattingKey, wil::com_ptr<IDWriteTextFormat>>`
- Implemented `FormattingKey` hash function using `std::hash<uint32_t>` for formatting bits
- Thread-safe via reader-writer lock (existing `_formatCacheLock`)

**Performance Impact:**

- Typical cache size: 5-20 text formats per control
- Before: O(n) linear search ~50-100 ns per lookup
- After: O(1) hash lookup ~10-20 ns per lookup
- Paint hot-path calls GetTextFormat() 5-20 times per frame → measurable improvement in complex TextFields/Grids

**Commit:** `098b80ee` — DxUi: migrate text format cache to O(1) unordered_map

---

### 9. DxUi Dirty-Rectangle Clipping Infrastructure (Yoko)

**Date:** 2026-03-21 | **Author:** Yoko (Core Dev) | **Impact:** Performance

Implemented Option A dirty-rectangle clipping layer. Accepts Win32 PAINTSTRUCT.rcPaint, applies D2D PushAxisAlignedClip when partial, and uses Present1 with dirty-rect parameters for compositor hints.

**Key Technical Decisions**

1. **FLIP_SEQUENTIAL is partial-render safe.** The OS preserves back buffer content between frames. Non-dirty regions retain previously-presented content without manual copy.

2. **D2D1_ANTIALIAS_MODE_ALIASED for clip.** The dirty rect is pixel-aligned (comes from Win32 RECT), so aliased mode avoids unnecessary AA overhead.

---

### 10. Preferences Dialog DxUi Migration — Phase 2 Root Causes & Fixes (Yoko, Coordinator)

**Date:** 2026-03-22 | **Status:** Completed | **Impact:** Architecture / Correctness

After Phase 1 removed shell chrome flag and per-page flags from 4 pages, Preferences Dialog reported 9 bugs. Phase 2 investigation identified root causes and applied fixes.

**Phase 2 Findings**

| Bug | Root Cause | Fix | Owner |
|-----|-----------|-----|-------|
| Viewers page blank | Legacy HWND gates blocking data flow (14 gates) | Removed HWND checks; data now flows settings → state → DxUi | Yoko |
| Plugins page blank | Legacy HWND gates blocking data flow (10 gates) | Removed HWND checks; restructured Refresh into query+sync phases | Coordinator |
| Pixel-to-DIP errors | Theory disproven; conversion already correct | None needed; added to decision log | Sysadm |
| Remaining feature flags | Dead code in 5 pages (HotPaths, Advanced, Panes, CompareDirectories, CompareDirectoriesWindow.Options) | Yoko identified in commit c37f1b65; cleanup pending | Yoko (follow-on) |

**Key Decision: Data Flow Pattern**

When migrating a Preferences page to DxUi, data must flow: **Settings → State Structs → DxUi Controls** (NOT through legacy HWND controls).

**Implementation Pattern:**
1. Legacy HWND creation continues (for SettingsStore compatibility)
2. All data-sync functions audit for HWND gates → repoint to state arrays
3. DxUi callbacks read/write state structs (vectors, indices)
4. Legacy HWNDs synchronized one-way from state (pull, not circular)

This breaks the circular dependency that caused blank pages during prior migrations.

**Architectural Insight**

The page host DxUi controls were successfully created; they just needed:
- Legacy HWND gates removed so data actually flows through
- State vectors established as the single source of truth
- Callbacks wired to read DxUi controls when legacy is absent

**Commit:** `c37f1b65` — Preferences: remove remaining DxSurface flags and stale HWND gates

---

### 11. Preferences Dialog Win32 Plumbing Review & Shell Chrome Analysis (Sysadm)

**Date:** 2026-03-21 | **Status:** Analysis Complete | **Impact:** Architecture

Deep-dive analysis of Preferences Dialog Win32 plumbing and shell chrome feature. Confirmed that the incomplete DxUi shell chrome was the initial blocker; proper integration requires:

1. **OK/Cancel/Apply button wiring** — Create DxUi button controls; wire to dialog command handlers
2. **Shell host height clipping** — Shell chrome should not extend below button area
3. **Mouse wheel forwarding** — Shell host must delegate wheel events to page host scroll logic

**Design Assessment**

The underlying Win32 layout, scroll, and mouse wheel systems are architecturally sound. They function correctly when shell chrome is disabled. The feature was marked as incomplete and should not have been enabled during migration.

**Recommendation:** Complete shell chrome implementation before re-enabling. Current status supports using Win32 buttons (feature disabled).

**Related:** See `sysadm-prefs-win32-review.md` for exhaustive technical analysis of all 4 issues.

3. **Present1 with dirty rects for compositor hint.** When partial, we use `IDXGISwapChain1::Present1()` instead of `Present()` to tell the desktop compositor which regions changed.

4. **Currently dead code.** All `InvalidateRect` callers pass NULL (full-window). The partial path activates when Option B adds targeted `InvalidateRect` with specific RECTs.

**What Remains (Option B)**

- Change `WindowHost::Invalidate()` to accept an optional `D2D1_RECT_F` region and convert to device-pixel RECT for `InvalidateRect`
- Add control-level dirty tracking: controls call `Invalidate(host, myBounds)` instead of full-window invalidate
- Grid: invalidate only changed rows/cells on data updates
- Scrollbar: invalidate only scrollbar track rect on scroll

**Commit:** `c2ba9f50` — DxUi: implement dirty-rectangle D2D clipping infrastructure

---

### 10. DxUi Scrollbar Extraction (Yoko)

**Date:** 2026-03-21 | **Author:** Yoko | **Impact:** Code Quality

Extracted ~165 lines of duplicated scrollbar rendering logic into `DxUi.Scrollbar.cpp`.

**Shared Helpers Created**

- `ResolveScrollbarVisuals()` — theme-based color resolution
- `ComputeScrollbarThumbRect()` — axis-agnostic thumb position/size calculation
- `PaintScrollbar()` — track + rounded-rect thumb rendering
- Constants: `kScrollbarThicknessDip`, `kScrollbarMinThumbDip`, `kScrollbarThumbCornerRadiusDip`, `kScrollbarThumbInsetDip`

**Key Design Choice**

`ComputeScrollbarThumbRect` takes 6 parameters rather than deriving extent from (totalContent - viewport) to preserve exact behavior for both Grid and Tree coordinate spaces.

**Commit:** `df078115` — Extract shared scrollbar logic from Grid/Tree into DxUi.Scrollbar.cpp

**Impact**

165 lines of code reduction; single source of truth ensures consistent behavior across controls.

---

### 11. DxUi Typeahead Extraction + Assessment (Yoko)

**Date:** 2026-03-21 | **Author:** Yoko | **Impact:** Code Quality

Extracted cleanly-shared typeahead helpers into `DxUi.Typeahead.cpp`. Intentionally preserved control-specific OnChar and FindMatch logic due to meaningful implementation differences.

**Shared Helpers Extracted**

- `kTypeaheadResetMs` — constant (~2 lines, was misnamed `kComboBoxTypeaheadResetMs` in Tree)
- `NormalizeTypeaheadChar()` — identical function (~8 lines)
- `StartsWithInsensitive()` — identical helper (~32 lines)

**Total: ~42 lines removed, ~30 lines added in new file.**

**What Was NOT Extracted (and Why)**

1. **OnChar accumulation:** Different control flow (Tree checks `empty() || timeout`, ComboBox has single-char fallback)
2. **FindTypeaheadMatch functions:** Same algorithm but different data access (Tree uses `_model->GetVisibleItem()`, ComboBox uses `_popupItemIndices[]`)
3. **Per-control state:** `_typeaheadBuffer` and `_lastTypeaheadTickMs` belong in each class

**Commit:** `e91942e3` — DxUi: Extract shared typeahead helpers into DxUi.Typeahead.cpp

**Impact**

Completes DxUi duplication extraction items from architecture review; follows established SingleLineTextEditing and Scrollbar extraction pattern.

---

### 12. DxUi Pointer Lifetime Documentation (Yoko)

**Date:** 2026-03-21 | **Author:** Yoko | **Impact:** Maintainability

Added formal non-owning pointer lifetime comments at SetModel() methods, member variable declarations, and PruneStaleInteractionState() implementations.

**Documentation Added**

- **SetModel() methods (Grid, Tree):** Document that Model pointer must remain valid for control lifetime
- **Member variables (_model):** Mark as non-owning with justification (Model lifetime managed by host)
- **PruneStaleInteractionState():** Clarify stale pointer detection and safe nullification logic

**Commit:** `f4169718` — docs: Add pointer lifetime documentation to Grid and Tree models

**Impact**

Establishes clear contract for Model lifetime management; reduces confusion for future maintainers; supports safe pointer invalidation during control cleanup.

---

### 13. Device Loss Test Coverage Enhancement (GuineaPig)

**Date:** 2026-03-21 | **Author:** GuineaPig | **Impact:** Testing

Enhanced `TestAttachedHostRecoversAfterSimulatedDeviceLoss` with comprehensive multi-control coverage and cache lifecycle validation.

**Expanded Coverage**

- **Multi-control rendering:** Grid (3 rows), Tree (2 items), TextField, Label all render during recovery
- **Brush cache lifecycle:** Verified populated → cleared on device loss → repopulated after recovery
- **D2D context lifecycle:** Verified present → null after loss → restored after recovery
- **Fallback brush lifecycle:** Same population/clear/repopulation pattern
- **Text format cache persistence:** DWrite format count unchanged (device-independent resource, correctly survives loss)

**Debug Accessors Added**

Four `#ifdef _DEBUG` WindowHost public methods added:
- `DebugGetBrushCacheSize()` — returns current brush cache entry count
- `DebugHasFallbackBrush()` — checks if fallback brush is initialized
- `DebugHasD2DContext()` — checks if D2D context is valid
- `DebugGetConfiguredTextFormatCount()` — returns live text format count

**Zero Release build overhead** (functions compiled out entirely).

**Commit:** `14909ee4` — DxUiTests: enhance device loss test with multi-control + brush cache assertions

**Impact**

Closes DxUi Technical Deep-Dive quality item (decision §2). Any future change to `DiscardDeviceResources()` or `RecreateBrushCache()` that breaks cache cleanup/rebuild will be caught immediately by test assertions.

---

### 14. FileSystem Search Optimizations — Assessment + Implementation (Sysadm)

**Date:** 2026-03-21 | **Author:** Sysadm | **Impact:** Performance

Assessed and implemented high-impact optimizations; confirmed 4 items already done or architecturally impossible.

**Implemented**

1. **Regex Compilation Cache** — bounded LRU (10 entries) for `std::wregex` objects. Avoids recompilation on repeated search patterns. Global to module, thread-safe.

2. **ETW Search Tracing** — 5 `Perf::Scope` markers for key search phases (overall, scan tree, parallel evaluate, indexed tree, service tree). Zero-cost when not collected. Enables data-driven optimization.

**Skipped (code analysis showed they're unnecessary)**

- **Filtered Enumeration:** No API-level filters available beyond what's already used.
- **Metadata Caching:** No redundant syscalls found — enumeration already provides full metadata.
- **Batch Error Handling:** Already bitwise-accumulates warnings; no per-file error callbacks.
- **Result Deduplication:** Already implemented via `emittedMatchPaths` hash set.

**Decision:** Remove these 4 items from the improvement plan — they're either already implemented or architecturally impossible. The regex cache and ETW tracing are the practical wins from this batch.

**Commit:** `befd4314` — FileSystem: add regex compilation cache (10-entry LRU) + ETW perf tracing

**Team Implications**

**For all agents:**
- Module pinning applies to ALL plugin DLLs with background work (FileSystem7z, FileSystemGoogleDrive, FileSystemS3, future plugins)
- Callback guard pattern establishes safety contract for async operations
- RegEx timeout pattern prevents ReDoS-class vulnerabilities across application

**For Ripley (profiling):** FileSystem performance baselines updated; directory enumeration now protected from pathological input.

**For GuineaPig (testing):** Should add test cases for ReDoS patterns (`(a+)+`, `(a*)*b`) to verify rejection; TOCTOU retry logic should be tested under high directory churn.

---

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction

### 15. Preferences Dialog DxUi Migration — Complete (Sysadm, Yoko, Ripley)

**Date:** 2025-03-22 | **Authors:** Sysadm (Implementation), Yoko (Page Flags), Ripley (Review) | **Impact:** UI/UX

Completed DxUi migration for Preferences Dialog, fixing 6 critical user-reported bugs (missing buttons, wrong sizing, broken scrolling). Removed kEnablePreferencesDxShellChrome shell chrome feature flag and all per-page DxUi migration gates. Root cause: incomplete hybrid DxUi/Win32 architecture with feature flags hiding Win32 controls without working DxUi replacements.

**Critical Bugs Fixed**

1. **OK/Cancel/Apply buttons invisible** — shell chrome flag hid Win32 buttons without positioning DxUi replacements
2. **Incorrect dialog size on open** — shell host overlay didn't account for button area
3. **No scrolling / mouse wheel** — scroll calculations broken by incomplete shell host layout
4. **Page switching left stale controls** — Win32/DxUi control cleanup incomplete
5. **Pages display blank** — initialization sequence incorrect (Resize before Refresh)
6. **Scroll position not applied** — missing scroll restoration after page switch

**Changes Made**

- **Shell chrome completion (Sysadm):** Removed kEnablePreferencesDxShellChrome flag, eliminated 23 conditionals, removed shellUsesDxUi state field, consolidated to DxUi-only code path
- **Per-page flag removal (Yoko):** Removed kEnableKeyboardDxSurface, kEnableThemesDxSurface, kEnableViewersDxSurface, kEnablePluginsDxSurface, eliminated legacy HWND fallback paths
- **Review & approval (Ripley):** Verified migration strategy, approved changes, identified 5 cleanup follow-ups

**Files Changed**
- RedSalamander/Preferences.Dialog.cpp (-800+ lines)
- RedSalamander/Preferences.Keyboard.cpp
- RedSalamander/Preferences.Themes.cpp (-220 lines legacy layout)
- RedSalamander/Preferences.Viewers.cpp
- RedSalamander/Preferences.Plugins.cpp (-195 lines legacy)
- RedSalamander/Preferences.Internal.h

**NET RESULT:** -1,066 lines removed across 6 files. Build clean, zero errors.

**Cleanup Scheduled (Next Pass)**
1. Remove dead local constants (currently hidden by removed #if blocks)
2. Clean bare scope blocks left by flag removal
3. Consolidate debug snapshot logic (now always returns true)
4. Simplify DebugUsesDxUi*() methods
5. Evaluate SyncDxFromLegacy() necessity

**Commit:** pending — squad: log Preferences Dialog DxUi migration session

**Team Implications**

**For all agents:**
- DxUi migration pattern established: remove hybrid feature flags early, complete in focused pass
- Complete forward migration (not revert) when architecture is sound
- Post-migration cleanup passes catch dead code from conditional removal

**For future UI work:** Preferences Dialog now template for complete DxUi migration; avoid hybrid approaches.


### 9. Phase 8 DxUI Infrastructure Cleanup — Remove Dual-Path Flags

**Date:** 2026-07
**Author:** Sysadm
**Status:** Implemented
**Impact:** Code Simplification

## Context

After Preferences Dialog DxUI migration Phases 1-7, all 9 panes (General, Panes, Viewers, Keyboard, Themes, Plugins, CompareDirectories, HotPaths, Advanced) had complete DxUI control implementations. However, each pane retained `_usesDxUi*` boolean flags that gated whether DxUI or Win32 fallback controls were used.

In practice, DxUI initialization always succeeds, so the Win32 fallback paths were dead code at runtime.

## Decision

Remove the dual-path DxUI/Win32 conditional architecture from all 9 fully-migrated panes:

1. **Simplify `EnsurePreferencesPageInitialized()`** — pass `false` (needsCreate) unconditionally for migrated panes. Editors/Mouse keep their existing conditions.

2. **Remove `_usesDxUi*` member fields** — `_usesDxUiStatics`, `_usesDxUiToggles`, `_usesDxUiInputs`, `_usesDxUiButtons`, `_usesDxUiList`, `_usesDxUiChrome`, `_usesDxUiEdits`, `_usesDxUiSwatch` removed from the 9 pane header files.

3. **Remove field assignments** — in `EnsureDxHosts()` (set to `true`) and `DetachDxHosts()` (set to `false`).

4. **Simplify layout code conditions** — `if (dxState && _usesDxUiStatics && _usesDxUiToggles && _pageHostDx && _pageContentRoot)` → `if (dxState && _pageHostDx && _pageContentRoot)`.

5. **Update `DebugUsesDxUi*()` methods** — return `true` unconditionally (since the flags no longer exist).

6. **Keep `DebugVisibleLegacy*Count()` methods** — they already return 0 for migrated panes. Still called from Dialog.cpp diagnostic output, so must remain.

7. **Keep helper functions** — `CreateFramedEditBox` and other helpers remain unchanged (both HWND and unique_hwnd overloads kept).

## Alternatives Considered

1. **Leave dual-path infrastructure in place** — rejected. Dead code confuses future maintainers and prevents full understanding of control flow.

2. **Remove `DebugVisibleLegacy*Count()` methods** — rejected. These are called from Dialog.cpp and serve as documentation that legacy controls no longer exist.

3. **Remove `DebugUsesDxUi*()` methods entirely** — possible but not done. Making them return `true` is simpler and preserves the API surface if needed for debugging.

## Rationale

- **Clarity:** Single code path is easier to understand and maintain.
- **Safety:** Removing dead code prevents accidental execution of untested fallback paths.
- **Simplification:** 239 lines removed, 19 files modified, 9 boolean fields eliminated.
- **Incremental:** Editors and Mouse panes retain Win32 controls (not yet migrated) — this cleanup is scoped to completed work only.

## Implementation

- **Commit:** 5c57739e on squad/dxui-filesystem-improvements
- **Files Modified:** 19 (9 .h files, 9 .cpp files, Preferences.Dialog.cpp)
- **Lines Changed:** +136 insertions, -375 deletions
- **Build:** Clean, zero new warnings

## Future Work

- Apply the same cleanup pattern to Editors and Mouse panes when their DxUI migration is complete.
- Consider creating a reusable "DxUI migration checklist" that includes removing dual-path infrastructure as the final step.

## Learnings

- Don't keep "debug flags" that toggle between rendering paths if there's only one path.
- When completing a migration, remove the scaffolding immediately — don't leave it "just in case."
- Keep diagnostic methods (like `DebugVisibleLegacy*Count()`) even if they return constant values, if they serve as documentation or are called by existing code.

---

### 7. Plugin Module Pinning — Mandatory Pattern (Sysadm)

**Date:** 2026-07 | **Status:** Decision | **Impact:** Reliability

All plugin DLLs MUST use the module pinning pattern when launching background threads or registering threadpool callbacks. The DLL can be unloaded (via `FreeLibrary`) while code is still executing on worker threads, causing a crash when the unmapped code pages are accessed.

**Pattern:**
- Add static anchor: `namespace { const int kPluginModuleAnchor = 0; }`
- Before background work: `wil::unique_hmodule pin = AcquireModuleReferenceFromAddress(&kPluginModuleAnchor);`
- For jthread: Capture pin in lambda: `std::jthread([pin = std::move(pin)](...)(...)`
- For threadpool: Store as member and acquire in Initialize, release in Shutdown

**Verification Status:**
- ✅ FileSystem — fixed (core + Watch + FileOps)
- ✅ FileSystemCurl — requires 3 fixes (scheduler + 2 streaming classes)
- ⚠️ FileSystem7z — requires 1 fix (SevenZipItemFileReader extract thread)
- ✅ FileSystemS3 — N/A (no background threads)

**Reference:** `.squad/agents/sysadm/history.md`, `Plugins/FileSystem/FileSystem.cpp`

---

### 8. Plugin Callback Safety Patterns (Ripley Lead)

**Date:** 2026-07-24 | **Status:** Implemented | **Impact:** Reliability, Security

Two critical safety patterns established for plugins with host callbacks (navigation menu, file system events):

#### 8a. Callback Drain Guard (SetCallback)

When clearing a callback pointer with `SetCallback(nullptr, nullptr)`, the implementation must guarantee no callbacks are in-flight before returning. Prevents use-after-free when the host destroys callback objects immediately after clearing.

**Pattern:**
- Add `_callbackGeneration` counter and `_callbacksInFlight` counter + condition variable
- On `SetCallback(nullptr)`: increment generation, wait for in-flight == 0
- On callback invoke: validate generation, increment in-flight, call, decrement + notify

**Reference Implementations:**
- FileSystemGoogleDrive::SetCallback (correct pattern)
- FileSystemMicrosoftDrive::SetCallback (was missing pattern — now fixed)

#### 8b. Configuration String Double-Buffer

When an API returns `const char*` to internal mutable state under lock, the pointer becomes dangling after lock release. Use double-buffering.

**Pattern:**
- Replace single `std::string _config` with array `std::string _config[2]`
- Add index `size_t _configIndex` (0 or 1)
- On SetConfiguration: write to inactive buffer, atomically flip index
- On GetConfiguration: return .c_str() of active buffer

**Reference:** FileSystemGoogleDrive::GetConfiguration, FileSystemMicrosoftDrive::GetConfiguration

**Applicability:**
- Required for any plugin returning pointer to internal mutable strings where lock cannot outlive return
- Both cloud plugins (Google Drive, Microsoft Drive) now use these patterns

---

### 9. Phase 8 Preferences DxUI Cleanup — Complete (Sysadm, Yoko)

**Date:** 2026-03-23 | **Status:** Implemented | **Impact:** Code Quality, Consistency

Completed 6 coordinated cleanup tasks (F1–F6) finalizing Preferences dialog DxUI migration:

#### 9a. F3 — Pure Rename: SyncDxFromLegacy → SyncDxControlsFromState
- 6 files, 30 occurrences
- Mechanical rename, no functional changes
- Prepares for sync pattern unification

#### 9b. F5 — Dead Code Removal: Win32 Toggle Infrastructure
- Removed `SetTwoStateToggleState` / `GetTwoStateToggleState` from PrefsUi namespace (~401 lines)
- DxUI toggles fire callbacks directly; Win32 WM_COMMAND path was unreachable
- Three panes (Panes, Advanced, CompareDirectories) had identical logic in both paths

#### 9c. F6 — Dead Code Removal: CreateFramedEditBox
- Removed `CreateFramedEditBox` and schema helpers (~655 lines)
- Plugin config DxUI migration was complete; no callers existed
- Cleaned up 3 orphaned WndProc hooks

#### 9d. F4 — Unified Sync Pattern
- All 8 migrated panes now use `SyncDxControlsFromState()` as single entry point
- Viewers pane retains sub-methods for partial-sync scenarios but wraps with `SyncDxControlsFromState()`
- Eliminates scattered `SyncDxFromLegacy` / `SyncDxFromState` inconsistency

#### 9e. F1 — Final Pane Migration: Editors & Mouse
- All 11 Preferences panes now DxUI-only
- Removed gates: `kEnableEditorsDxNoteSurface` / `kEnableMouseDxNoteSurface`
- Removed debug methods: `DebugUsesDxUiNoteSurface()` / `DebugVisibleLegacyNoteCount()`
- Updated 6 self-test assertions to `true /* F1: removed field */`

#### 9f. F2 — LayoutDxPage Pattern Standardization
- All 9 panes follow identical pattern: thin wrapper → EnsureDxHosts → LayoutDxPage → error logging
- Consistent error handling across all panes
- Editors and Mouse should adopt when DxUI migration complete

**Total Impact:**
- Dead code removed: ~1,056 lines
- Feature gates eliminated: 4 constants
- Consistency: All 11 panes unified DxUI pattern
- Build: Clean, zero warnings

---

### 10. Cloud Drive Plugin Design Review (Ripley Lead)

**Date:** 2026-07-24 | **Status:** Analysis | **Impact:** Reliability, Security

Comprehensive review of two cloud plugins (Google Drive, Microsoft Drive) identified critical issues and maturity assessment.

#### 10a. FileSystemGoogleDrive — Prototype / Early Alpha

**Critical Issues (P0):**
- C1: Refresh token stored/transmitted as plain `std::wstring` without `SecureZeroMemory`
- C2: GetConfiguration returns `.c_str()` under lock → dangling pointer on concurrent update (fixed via double-buffer pattern in decision 8b)
- C3: `curl_global_init` called but `curl_global_cleanup` never called → resource leak
- C4: CurlWriteToString callback catches `std::exception` — error context lost

**Improvements (P1–P3):**
- No retry logic for HTTP 429/5xx responses (critical for Google Drive rate limiting)
- No pagination safeguard / runaway loop protection
- Token refresh synchronous and serialized (no deduplication)
- Error handling for libcurl coarse
- No "Open in Browser" integration

**Architecture:** Browse-only (no write operations). Structural quality good; missing production features.

#### 10b. FileSystemMicrosoftDrive — Beta / Feature-Complete

**Critical Issues (P0):**
- C1: SetCallback has no drain/generation guard (fixed via callback drain pattern in decision 8a)
- C2: GetConfiguration dangling pointer (fixed via double-buffer pattern in decision 8b)
- C3: OAuth loopback listener blocks caller thread for up to 5 minutes
- C4: `std::thread` in OAuth has no module pinning (joined, but fragile)
- C5: Raw `FileSystemMicrosoftDrive*` pointers in reader/writer with manual AddRef/Release — violates AGENTS.md
- C6: Static menu items not thread-safe
- C7: Static WinHttpOpen session with no cleanup → potential deadlock on DLL unload

**Improvements (P1–P3):**
- OAuth PKCE verifier and auth code should be SecureClear'd
- No exponential backoff jitter in retry loop
- `std::this_thread::sleep_for` blocks on retry
- No connection pooling / keep-alive reuse
- File download buffers entire 1MB chunk
- Temp file TOCTOU window in upload
- WSAStartup/WSACleanup per OAuth flow (wasteful)
- Move-with-overwrite rollback failure not logged

**Architecture:** Full-featured (move, delete, rename, file read/write). Most complete cloud plugin. Main gaps: callback teardown race, raw COM pointers, static WinHTTP cleanup.

#### 10c. Cross-Plugin Summary

**Shared Issues:** GetConfiguration dangling pointer (BOTH), Code duplication (UTF conversion, path normalization, JSON parsing), SetCallback pattern missing (MS Drive).

**Priority:**
1. P0: Implement callback drain + double-buffer patterns (both plugins)
2. P1: SecureClear for tokens, fix raw COM pointers, add retry logic
3. P2: Static singleton cleanup, thread safety on menu items
4. P3: Factor shared code into Common/, optimize uploads

**Decision Required:** Investment in Google Drive feature-parity vs. keep as browsing prototype?

---

### 11. Plugin Design Reviews — Batch 1 (Sysadm)

**Date:** 2026-07 | **Status:** Review Complete | **Impact:** Stability, Performance

Deep-dive reviews of three additional plugins (7z, Curl, S3).

#### 11a. FileSystem7z — Solid / Read-Only

**Critical:** Missing module pinning for SevenZipItemFileReader jthread (1 site).

**Improvements:** EnsureIndex blocks UI thread, no cancellation in BuildIndex, password stored plaintext, no index invalidation on archive modification.

**Strengths:** Dual-path reader (in-memory spool ≤32MB, ring-buffer pipe for large files), lazy index building, condition-variable synchronization.

#### 11b. FileSystemCurl — Large, Well-Structured

**Critical:** Missing module pinning for 3 sites:
- SharedCopyMoveJobScheduler (static singleton with jthreads)
- CurlStreamingReader jthread
- CurlStreamingWriter jthread

**Additional Critical:** Missing `curl_global_cleanup` coordination with static CurlEasyPool.

**Improvements:** ConnectionConcurrencyLimiter polling (100ms timeout), IMAP batch size hardcoded at 100, no network retry policy.

**Strengths:** Exemplary RAII, correct error handling (no catch-all), synthetic watch with callback guard, connection pooling.

#### 11c. FileSystemS3 — Clean / Synchronous

**Critical:** None. No background threads → no module pinning needed.

**Improvements:** Multipart buffer memmove inefficiency, duplicate callback lambdas, ListRecursiveObjects lacks progress, AWS SDK init/shutdown exception safety correct but verbose.

**Strengths:** No background threads, proper RAII, typed exception handling only, atomic watch active flag, FILE_FLAG_DELETE_ON_CLOSE temp files.

#### 11d. Action Items

**HIGH:** Module pinning for FileSystemCurl (3 sites) and FileSystem7z (1 site).

**MEDIUM:** Static singleton lifecycle (CurlEasyPool), network retry policy for FileSystemCurl.

**LOW:** Async index building for 7z, multipart buffer optimization for S3, extract duplicate lambdas.

---

### 12. Preferences Dialog DxUI Migration — Pages (Yoko)

**Date:** 2026-03-22 | **Status:** Review | **Impact:** UI Migration

Per-page DxUI feature flag removal across Keyboard, Themes, Viewers, Plugins. All four pages now use DxUI exclusively.

**Completed Pages:**
- Keyboard: Flag removed, DxUI-only, legacy conditionals removed
- Themes: Flag removed, DxUI-only, legacy conditionals removed
- Viewers: Flag removed, but legacy HWND creation code still present (~135 lines need removal)
- Plugins: Flag removed, but CreateControls and legacy control creation not updated

**Pattern:** All pages should follow: DxUi-only initialization with error logging on `EnsureDxHosts()` failure.

**Next Steps:** Complete Viewers and Plugins cleanup, remove legacy HWND creation blocks, verify all 4 pages render correctly, update owner-draw methods.

---

### 13. Plugin Config & Preferences Buttons — Migration Complete (Yoko)

**Date:** 2025-01-19 | **Status:** Implemented | **Impact:** Code Quality

Migrated 6 button types from HWND routing to direct handler calls (Plugins + Plugin Config sections).

**Changes:**
- Extracted 5 handler functions: OnConfigureButtonClick, OnTestButtonClick, OnTestAllButtonClick, OnCustomPathsAddButtonClick, OnCustomPathsRemoveButtonClick
- Removed 6 HWND button fields from PreferencesDialogState
- Removed dead WM_COMMAND cases
- Plugin Config browse button already used DxUi callback pattern (no routing needed)

**Pattern:** Direct call with state pointer capture in DxUi callbacks; null guards for state, host window, IsWindow.

**Consistency:** Matches Themes button pattern; all Preferences buttons should eventually follow.

**Remaining:** ~40 references to removed button fields in legacy code guarded by unconditionally-true flag (dead code, not blocking).

---

### 14. Shell Chrome DxUI Migration — Removed Feature Flag (Sysadm)

**Date:** 2026-07 | **Status:** Implemented | **Impact:** Consistency

Removed `kEnablePreferencesDxShellChrome` feature flag. DxUi shell chrome is now the only code path in Preferences.Dialog.cpp. All 23 conditional branches eliminated.

**Changes:**
- Preferences.Dialog.cpp: Removed 23 conditionals, removed `shellUsesDxUi` state field
- Preferences.Internal.h: Removed `bool shellUsesDxUi` field
- Preferences.h: Kept `shellUsesDxUiHost` in debug snapshot (hardcoded true)

**Impact:** Simplified code, single rendering path, consistent with DxUI-only transition.

---

### 15. Per-Page DxUI Gate Removal — Complete (Yoko)

**Date:** 2026-07 | **Status:** Implemented | **Impact:** Code Quality

All four Preferences dialog pages now use DxUi exclusively. Per-page `constexpr bool kEnable*DxSurface = true` flags and legacy HWND fallback paths fully removed.

**Pages:**
- Keyboard: Complete (prior session)
- Themes: Complete (prior session)
- Viewers: Complete (prior session)
- Plugins: Completed this session

**Pattern:** Direct `EnsureDxHosts()` calls with `Debug::Error()` logging on failure. All legacy conditionals removed.

**Remaining Work (Phase 2):**
- Legacy HWND members in PreferencesDialogState (still used for data sync)
- Owner-draw methods for LISTVIEWs
- DebugUsesDxUi*() methods (now always return true)
- SyncDxFromLegacy() calls
- Preferences.Dialog.cpp cleanup (unused locals from removed flags)

---

---

### 4. vcpkg Auto-Discovery from System Installations (Sysadm)

**Date:** 2026-07 | **Status:** Implemented | **Impact:** Developer Ergonomics

vcpkg-install.ps1 now auto-discovers vcpkg from common system-wide installation locations before erroring. Supports Visual Studio bundled (via vswhere), Chocolatey, Scoop, and common directories (C:\vcpkg, C:\dev\vcpkg, etc.).

**Discovery Priority Chain:**
1. Explicit `-VcpkgExe` parameter
2. Repo-local `vcpkg\vcpkg.exe`
3. ``
4. PATH (`Get-Command vcpkg.exe`)
5. Visual Studio bundled (vswhere query)
6. Chocolatey (`C:\tools\vcpkg\`, `\lib\vcpkg\tools\`)
7. Scoop (`C:\Users\eric\scoop\apps\vcpkg\current\`, `\apps\vcpkg\current\`)
8. Common roots (`C:\vcpkg\`, `C:\dev\vcpkg\`, `C:\Users\eric\vcpkg\`, `C:\Users\eric\source\vcpkg\`)
9. Error with diagnostic listing all discovery methods

**Rationale:**
- Developers with VS+vcpkg component or Chocolatey/Scoop installations now have "just works" experience
- No breaking changes; explicit param, repo-local, VCPKG_ROOT, PATH still take precedence
- CI unaffected (use explicit VCPKG_ROOT or repo-local)
- Diagnostics visible with `-Verbose`

**Implementation:** `Resolve-VcpkgExePath` function in `vcpkg-install.ps1`

**Decision:** Approved. Improves out-of-box experience for most Windows C++ developers.

---

### 8. vcpkg-install.ps1 Default All-Triplet Installation + Help Flag (Sysadm)

**Date:** 2026-07 | **Status:** Implemented | **Impact:** Build Infrastructure, Developer Experience

`vcpkg-install.ps1` now installs **all four triplet combinations** by default and supports comprehensive help flags.

**Multi-Triplet Default Installation**

Default behavior (no args) installs all four triplets:
- x64-windows
- arm64-windows
- x64-windows-asan
- arm64-windows-asan

Previously defaulted to x64-windows only.

**Rationale:** RedSalamander supports multiple platform/configuration combinations. Developers need dependencies for x64/ARM64 platforms AND standard/AddressSanitizer builds. Installing all four by default ensures a complete development environment from one command, avoiding surprise linker errors when switching platforms or asan configurations.

**New Parameter Semantics**

- `-Platform` (default: `All`) — `x64`, `ARM64`, or `All`
- `-Asan` (switch) — Toggles AddressSanitizer triplet variants
  - With `-Platform All` (default) + no `-Asan`: install all four
  - With `-Platform All -Asan`: install only asan triplets
  - With `-Platform All -Asan:$false`: install only non-asan triplets
  - With `-Platform x64/ARM64`: install one triplet per `-Asan` switch
- `-Triplet` (explicit override) — Bypasses all logic

**Help Flag Support**

Added comprehensive help system with standard PowerShell conventions:
- `-Help`, `-h` — Print full help via Get-Help and exit
- `-?` — Built-in CmdletBinding help
- `Get-Help .\vcpkg-install.ps1 -Full` — PowerShell Get-Help
- Comment-based help expanded to 8 examples covering all usage patterns
- Full parameter documentation (Platform, Asan, Triplet, Clean, VcpkgExe, Help)

**Implementation Details**

- Uses `$PSBoundParameters.ContainsKey('Asan')` to detect explicit vs default
- Loops through `$tripletsToInstall` array with per-triplet progress headers (cyan)
- Aborts on first failure with exit code propagation
- Post-install WIL header validation runs for each installed triplet

**Examples**

```powershell
.\vcpkg-install.ps1                          # Install all 4 triplets
.\vcpkg-install.ps1 -Platform x64            # Install x64-windows only
.\vcpkg-install.ps1 -Platform ARM64 -Asan    # Install arm64-windows-asan
.\vcpkg-install.ps1 -Platform All -Asan      # Install asan triplets only
.\vcpkg-install.ps1 -Help                    # Show help
```

**Impact on Other Systems**

- **CI/Build Scripts:** Explicit `-Platform x64` or `-Triplet` still work (no breaking changes)
- **Developer Workflow:** First-time setup now completes in one command (installs all deps at once)
- **MSBuild:** No changes; triplet selection per configuration unchanged
- **Time:** Default install takes ~4x longer on clean install (4 triplets); incremental installs are fast

**New Skill Created**

`.squad/skills/powershell-multi-target-install/SKILL.md` documents the pattern for multi-target installation with smart defaults, parameter layering, and explicit-vs-default detection.

**Decision:** Implemented. Improves out-of-box developer experience and establishes reusable PowerShell parameter pattern for multi-target scenarios.

---

### 9. vcpkg ASan Triplets Are Explicit Opt-In (Sysadm)

**Date:** 2026-04-28 | **Status:** Implemented | **Impact:** Build Infrastructure, Developer Experience

`vcpkg-install.ps1` now makes ASan triplet dependencies explicit opt-in.

**Default Behavior (No `-Asan` flag)**

Default `-Platform All` (no arguments) installs only non-ASan triplets:
- x64-windows
- arm64-windows

This avoids building third-party dependencies with ASan, which currently fail for `openssl:x64-windows-asan`.

**Opt-In to ASan Triplets**

- `-Platform All -Asan` installs only ASan triplets (x64-windows-asan, arm64-windows-asan)
- `-Platform x64 -Asan` installs x64-windows-asan
- `-Platform ARM64 -Asan` installs arm64-windows-asan
- `-Triplet` parameter overrides everything (escape hatch preserved)

**Rationale**

ASan dependencies are only needed when the AddressSanitizer Debug configuration is intentionally requested. The default install must succeed without attempting ASan third-party builds. Developers who want ASan debugging can explicitly opt in via `-Platform All -Asan` or specific platform `-Asan` flags.

**Help Support**

- `-Help`, `-h` — Print updated help with new examples
- `-?` — Built-in CmdletBinding help
- Comment-based help updated with opt-in examples

**Examples**

```powershell
.\vcpkg-install.ps1                       # x64-windows, arm64-windows only
.\vcpkg-install.ps1 -Platform x64         # x64-windows only
.\vcpkg-install.ps1 -Platform All -Asan   # x64-windows-asan, arm64-windows-asan only
.\vcpkg-install.ps1 -Platform x64 -Asan   # x64-windows-asan only
.\vcpkg-install.ps1 -Help                 # Show help
```

**Impact on Existing Workflows**

- **New developers:** Default install now succeeds without ASan dependency failures
- **CI/Build Scripts:** Explicit `-Platform x64` still works (no breaking changes)
- **ASan Debug builds:** Developers using AddressSanitizer explicitly request triplets via `-Asan`
- **MSBuild:** No changes; triplet selection per configuration remains unchanged

**Decision:** Implemented. Eliminates out-of-box openssl:x64-windows-asan build failures and establishes ASan as an explicit opt-in for developers who specifically need it.

---

### 10. ASan vcpkg OpenSSL Provisioning: Port-Side Fix (OpusVcpkg) (2026-04-28)

**Agents:** VcpkgSurgeon (diagnosis), OpusVcpkg (implementation), GuineaPig (validation), Ripley (review)
**Status:** APPROVED WITH EXTERNAL BLOCKER | **Impact:** Build Infrastructure, Developer Experience

#### Root Cause

vcpkg's `scripts/cmake/vcpkg_cmake_get_vars.cmake` runs MSVC platform detection after loading triplet files. MSVC's `Platform/Windows-MSVC.cmake` injects `CMAKE_SHARED_LINKER_FLAGS_DEBUG_INIT = "/debug /INCREMENTAL"`. vcpkg merges detected flags with triplet flags:

```
VCPKG_COMBINED_SHARED_LINKER_FLAGS_DEBUG = "-machine:x64 -nologo -INCREMENTAL:NO   -debug -INCREMENTAL"
                                                    ↑ triplet          ↑ MSVC platform default
```

OpenSSL's nmake consumes this combined value, producing conflicting `/INCREMENTAL:NO` and `/INCREMENTAL`, causing `LNK1158: linker recursion`.

**Key insight:** Triplet-only architecture cannot suppress MSVC platform defaults. Triplet `VCPKG_LINKER_FLAGS` are **additive only**; they cannot override detected MSVC defaults.

#### Solution

Remove stale OpenSSL patch and implement runtime flag-stripping in OpenSSL's `windows/portfile.cmake`:

1. **`vcpkg-overlay-ports/openssl/portfile.cmake`**
   - Removed `disable-incremental-linking-asan.patch` from PATCHES list (broken; OpenSSL 3.6.1 uses CMake)
   - PATCHES now mirrors upstream registry `openssl@3.6.1#3` exactly

2. **`vcpkg-overlay-ports/openssl/windows/portfile.cmake`**
   - After `vcpkg_cmake_get_vars` / `include(${cmake_vars_file})`, regex-strip bare `-INCREMENTAL` / `/INCREMENTAL` (not followed by `:`)
   - Transformation: `-INCREMENTAL:NO -debug -INCREMENTAL` → `-INCREMENTAL:NO -debug`
   - Preserves `-INCREMENTAL:NO` from triplet (negative lookahead in regex)
   - Scoped to DEBUG configs only

3. **Triplet cleanup (VcpkgSurgeon)**
   - Removed `VCPKG_BUILD_TYPE=release` from both x64/arm64-windows-asan triplets
   - Restored multi-config (debug + release) libraries per MSBuild contract

#### Validation Evidence

- **Pester test suite:** 4/4 passed
- **Flag transformation:** Verified via log inspection; trailing `-INCREMENTAL` successfully stripped
- **Link command:** OpenSSL link now contains only `-INCREMENTAL:NO` and `-debug` (no conflict)

#### Remaining Blocker: MSVC Toolchain (External)

OpenSSL debug DLL link fails with `LNK1158: cannot run link.exe` — this is **link.exe's self-spawn failure** for ASan runtime injection, **NOT** a port-configuration issue:

- ✅ Flag conflict (triplet + MSVC defaults) is **eliminated** (verified)
- ✅ Minimal ASan DLL link succeeds on same MSVC 14.51.36231
- ✅ OpenSSL archive-heavy build fails only on VS18 Insiders (14.51.36231)
- ✅ Consistent with VS18 Insiders regression in link.exe self-spawn logic

**Recommendation:** Users must provision ASan vcpkg on stable VS17 RTM or VS18 RTM toolchains. This is an external toolchain issue, not a repo configuration defect.

#### Architecture Decision: Reject Release-Only Workaround

**Ripley (2026-03-21) explicitly REJECTED** the release-only triplet workaround (`VCPKG_BUILD_TYPE=release`). Rationale:

- Violates MSBuild contract: `Directory.Build.props` lines 46–49 hardwire ASan Debug → x64-windows-asan triplet
- ASan Debug configuration **must** find debug vcpkg libraries with ASan instrumentation
- Release-only triplet breaks linker phases (LNK2019 unresolved externals for yyjson, libcurl, sqlite3)
- Single Responsibility: A triplet owns one build type and must provide libs matching that type

This decision is **binding** and **preserved** in the archive below.

#### Acceptance Criteria Status (per GuineaPig)

| # | Criterion | Status | Notes |
|---|-----------|--------|-------|
| 1 | Triplets NOT `VCPKG_BUILD_TYPE=release` | ✅ PASS | Cleaned; release-only rejected |
| 2 | OpenSSL patch applies | ✅ PASS | Patch removed; no apply failures |
| 3 | OpenSSL builds | 🔶 BLOCKED | Link fails (MSVC self-spawn); not config |
| 4 | No LNK1158 (flag conflict) | ✅ FIXED | Original root cause eliminated |
| 5 | Debug + Release libraries exist | 🔶 BLOCKED | Depends on #3 (MSVC issue) |
| 6 | First-party ASan Debug links | 🔶 BLOCKED | Depends on #3 (MSVC issue) |
| 7 | Patch matches OpenSSL | ✅ PASS | Patch eliminated |

#### Files Changed

- `vcpkg-overlay-triplets/x64-windows-asan.cmake` (drop release-only, add comment)
- `vcpkg-overlay-triplets/arm64-windows-asan.cmake` (drop release-only, add comment)
- `vcpkg-overlay-ports/openssl/portfile.cmake` (drop broken patch from PATCHES list)
- `vcpkg-overlay-ports/openssl/windows/portfile.cmake` (add regex strip block)
- `vcpkg-overlay-ports/openssl/disable-incremental-linking-asan.patch` *(deleted)*

#### Decision: APPROVED FOR MERGE

Ripley approved OpusVcpkg's work as production-ready. The vcpkg-side fix is minimal, correct, and well-validated. The remaining LNK1158 blocker is a documented MSVC Insiders issue, not a configuration defect. Merge these changes and document the toolchain requirement in `AGENTS.md` or `Specs/VcpkgIntegration.md`.

---

### 11. ASan Triplet Rejection Archive: Release-Only Build Strategy Was Unsound (2026-03-21)

**Status:** HISTORICAL | **Impact:** Design, Developer Experience

**Context:** On 2026-03-21, Ripley reviewed Sysadm's release-only triplet workaround (`VCPKG_BUILD_TYPE=release`) for ASan triplets and **explicitly REJECTED** it as architecturally unsound.

**Rejection Rationale:**

1. **Violates MSBuild contract** — `Directory.Build.props` lines 46–49 explicitly map ASan Debug → x64-windows-asan triplet; that triplet must support debug linking
2. **Linker failures in ASan Debug builds** — With release-only triplet:
   - Common.dll: yyjson unresolved (yyjson_read_opts, yyjson_mut_*)
   - FileSystemGoogleDrive.dll: libcurl unresolved (__imp_curl_*)
   - ViewerSqliteTests: sqlite3 unresolved (__imp_sqlite3_close_v2)
3. **Architecture violation** — A triplet owns one build type; it must provide libs matching that type. Release-only triplet cannot satisfy Debug linker phases.

**Constraints for Replacement (specified by Ripley):**

- **Strategy A (preferred):** Multi-config triplet (both debug + release libs) with OpenSSL conflict verification
- **Strategy B:** Debug-only triplet with updated MSBuild mapping (infrastructure cost)
- **Strategy C (not recommended):** Conditional release fallback (removes ASan instrumentation from deps)

**Outcome:** OpusVcpkg implemented Strategy A via portfile.cmake flag-stripping, restoring multi-config triplet support while fixing the OpenSSL linker conflict.

**Historical Note:** This rejection is preserved to prevent future regressions and to document the binding MSBuild contract (ASan Debug ↔ ASan-instrumented debug libs).

---

### 12. vcpkg ASan Provisioning: All-Platform Additive Merge Strategy (Sysadm) (2026-04-28)

**Date:** 2026-04-28 | **Status:** Validated | **Impact:** Build Infrastructure

When users request ASan for all platforms (`-Asan` or `-Platform All -Asan`), `vcpkg-install.ps1` installs normal triplets first, then ASan triplets. Each triplet installs into isolated staging, then stages merge into canonical `.build\vcpkg_installed\<triplet>` layout.

**Rationale:**

- Normal Debug/Release builds consume normal triplet include trees (e.g., `wil\com.h`)
- If ASan dependency like OpenSSL fails before normal triplets are provisioned, regular builds lose headers and fail unrelated to ASan
- vcpkg manifest mode prunes sibling triplets when repeated invocations share `--x-install-root`; isolated staging keeps pruning local to one triplet and preserves already-merged canonical triplets
- Specific-platform ASan requests stay precise (`-Platform x64 -Asan` installs x64 ASan only, not ARM64)
- `-Triplet` remains highest-priority override (escape hatch)

**Validated:**

- `-Asan` or `-Platform All -Asan` → normal x64, normal ARM64, ASan x64, ASan ARM64
- `-Platform x64 -Asan` → x64 ASan only
- Staged merge preserves coexisting canonical triplet directories
- Help system supports new opt-in examples

**Decision:** Implemented. Additive provisioning keeps ASan as explicit opt-in while preserving normal build success.

---

### 11. vcpkg Triplet Merge Lock: Safe Merge Strategy (GuineaPig/OpusMerge/Ripley)

**Date:** 2026-04-28 | **Status:** APPROVED | **Impact:** Infrastructure, Build Reliability

**Problem:** vcpkg triplet provisioning fails when destination files are locked by IDE/indexers. Existing merge strategies (robocopy /IS, hash-based with silent skip) were either insufficient or unsafe.

**Contract (GuineaPig Validation Artifacts):**
1. Existing triplet preservation — must NOT delete entire destination triplet
2. Hash-based skip — if staged and destination files hash equal, skip without write/delete
3. Tolerant of read-share locks — skip identical files even if destination locked by IDE/indexer
4. Fail on unreadable destination — if destination unreadable, fail clearly with file path
5. Fail on lock-blocked replacement — if hashes differ and replacement blocked by lock, fail with file path and close-IDE/indexer guidance
6. Preserve destination extras — destination files not in staged are not pruned (use cpkg-install.ps1 -Clean for reset)

**Implementation:** Tools\VcpkgInstallSafety.ps1 (integrated into cpkg-install.ps1 for existing triplet merges)

**Algorithm:**
- For each triplet: Test destination readability
- For each staged file: Compute SHA256 of staged and destination
  - If hashes match → skip (no write/delete)
  - If destination unreadable → fail clearly
  - If differ → attempt copy; if locked → fail with guidance

**Test Results:**
- ✅ Pester suite: 4/4 passing (hash comparison, lock detection, unreadable detection)
- ✅ Synthetic suite: 5/5 passing (skip identical, replace differing, fail unreadable, fail locked, preserve extras)

**Validation Path:**
1. GuineaPig defined 6 binding acceptance criteria + synthetic tests
2. Sysadm attempted robocopy (rejected—insufficient for lock tolerance)
3. MergeDoctor attempted hash-based with silent skip (rejected—violates clear-fail requirement #4)
4. OpusMerge implemented approved solution with explicit lock/readability failures
5. Ripley approved after validation

**Decision:** APPROVED for production. Ready for merge.

---
