# Operation Farsight — Viewer Plugins Remediation: Export Safety, Decode Reliability, Web Security & Text Geometry

**Status:** WIP
**Date:** 2026-06-16
**Author:** Independent adversarial multi-agent review of all 7 viewer plugins at `master` `b274022d9` (branch `claude/thirsty-pascal-19c8f5`). 89-agent throttled audit: 15 review lenses (per-plugin deep dives on the large translation units — ViewerText/Hex/Text, ViewerSpace, ViewerVLC, ViewerWeb, ViewerImgRaw core/decode/export, ViewerPE, ViewerSqlite — plus four cross-cutting lenses: IViewer lifecycle/UAF, RAII/leaks, untrusted-input security, architecture/duplication) → per-finding adversarial refutation pass → advisor re-verification against source (workflow run `wf_3d2e0ea2-a76`). 51 findings survived verification (**1 critical, 13 high, 11 medium, 26 low**); 23 refuted. Every finding turned into a slice was re-opened and re-read line-by-line by the lead reviewer before planning.

**Scope of this pass (~54,588 LOC under `Plugins/Viewer*`):**
- `Plugins/ViewerImgRaw/*` — `ViewerImgRaw.cpp` (viewer/window/D2D), `ViewerImgRaw.Decode.cpp` (libraw/turbojpeg/WIC decode, EXIF/TIFF), `ViewerImgRaw.Export.cpp` (WIC encode/save)
- `Plugins/ViewerWeb/*` — `ViewerWeb.cpp` (WebView2, JSON/JSONL/Markdown rendering)
- `Plugins/ViewerText/*` — `ViewerText.cpp` (core/lifecycle/async), `ViewerText.Text.cpp` (text-mode paint/hit-test), `ViewerText.Hex.cpp` (hex mode)
- `Plugins/ViewerPE/*`, `Plugins/ViewerVLC/*`, `Plugins/ViewerSpace/*`, `Plugins/ViewerSqlite/*`
- Contract: `Common/PlugInterfaces/Viewer.h` (IViewer ABI). Tests: `Tests/ViewerPETests` (drives ALL viewers via window/menu/UIA), `Tests/ViewerSqliteTests`, `Tests/PluginContractTests`; registered in `Tools/TestRunPlan.ps1` `Full`.
- **Executor-ready handoff with full file:line excerpts and step-by-step edits:** `plans/014-*.md` … `plans/030-*.md` on branch `claude/thirsty-pascal-19c8f5` (one file per slice below).

> **Anchors are relative to `b274022d9`. Re-grep every `~line` before editing — line numbers drift.**

---

## Why this plan exists

**Verdict: broadly well-engineered, but NOT yet "rock solid."** The viewers get the hard parts mostly right — most plugins snapshot inputs by value before going off-thread, AddRef `this` for the worker's lifetime, drop stale work by `requestId`, drain posted payloads on teardown, use WIL RAII consistently, and bounds-check the untrusted decoders (libraw/turbojpeg/TIFF/EXIF/PE). The audit confirmed and **refuted 23** plausible-but-wrong findings (see "Verified correct"). But it also found real defects that block a business-critical "rock solid" sign-off, concentrated in five classes:

1. **One data-corrupting use-after-free on the most dangerous surface (export).** ViewerImgRaw holds a reference into the cached pixel buffer across the modal save dialog's message pump; an in-flight decode completing during the dialog frees/moves it → UAF or a silently corrupt/blank image written to the user's file. This is the single ship-blocker.
2. **A whole capability is silently dead.** ViewerImgRaw's decode workers never initialize COM, so every WIC-backed format (PNG/GIF/BMP/non-RAW TIFF/HEIC/WebP) fails to open — masked because JPEG and RAW don't touch COM.
3. **A security hole on intentionally-untrusted input.** ViewerWeb embeds untrusted JSON/JSONL/Markdown into inline `<script>` with an escaper that misses `</script>`, and ships no CSP → script injection in the viewer origin.
4. **Concurrency/lifetime gaps.** A COM-refcount data race in ViewerWeb's async load (UAF/double-free of the IFileSystem plugin); object/COM-self-ref leaks + an OOM null-deref across four async dispatchers; a UI-thread freeze in ViewerPE; two VLC teardown/loader races.
5. **Pervasive text-rendering geometry errors.** ViewerText's text mode assumes one code unit == one fixed column, so tabs, CJK/proportional glyphs, and surrogate pairs desync selection/caret/search/hit-test — the wrong text gets selected and copied on very common files.

**Standing rule for every Farsight slice:** a slice is not done because a self-test is green. For each P0/P1 slice the proof must make the *dangerous, dead, or wrong condition observable* and **FAIL (RED) on the current tree** before the fix, then pass after. Where a deterministic UI repro is impractical (a refcount race, a threadpool-submit failure), the proof is a line-by-line review of the AddRef/Release/ownership balance plus a structural assertion (e.g. "the worker dereferences no mutable `self->_member`").

**Cross-cutting first (fix the class, not the instance):**
- The **async-dispatch leak + OOM null-deref (F-S0-4)** is *one* mistake replicated at 5 sites; fix all 5 to the one correct shape (ViewerWeb's `StartAsyncLoad` is the in-tree exemplar), and later hoist it into a shared `Common` helper (F-S3-3 / `plans/030` Stage 3).
- The **WIC-CoInit (F-S0-3)** is one chokepoint (`DecodeImageToBgraWic`), not three call-site edits.
- The **text geometry (F-S1-4)** is one model (fixed `charW`), not three symptoms; the layout-based fix subsumes tabs, CJK, and surrogates.

---

## Implementation Tracking Checklist (update first, before editing code in a slice)

Use `[ ]` not started, `[~]` in progress, `[x]` complete, `[blocked]` needs a product decision.
Each slice maps 1:1 to an executor-ready handoff file on branch `claude/thirsty-pascal-19c8f5`.

| State | Slice | Plan | Sev | Plugin | Required proof before `[x]` |
|-------|-------|------|-----|--------|-----------------------------|
| [ ] | F-S0-1 | 014 | CRIT | ImgRaw | An export of a just-opened image (background decode in flight) does not read freed memory and writes a non-blank file; no `image->`/`bgra` reference survives `ShowExportSaveDialog`. RED: today a reference into the cache is read after the modal pump. |
| [ ] | F-S0-2 | 016 | HIGH | ImgRaw | A forced mid-encode failure leaves a pre-existing destination byte-identical; happy path still produces the file. RED: today `InitializeFromFilename` truncates the destination before any valid bytes. |
| [ ] | F-S0-3 | 015 | HIGH | ImgRaw | Opening a `.png` decodes and renders (not the error state). RED: today every WIC format fails with `CO_E_NOTINITIALIZED` on the decode worker. |
| [ ] | F-S0-4 | 017 | HIGH | ImgRaw×3, Text | At all 5 dispatch sites: exactly one `Release` per `AddRef` on every exit path (null-alloc, submit-fail, success); happy path unchanged. RED: today submit-failure leaks the AddRef (+ work item in ImgRaw) and a `nothrow` alloc is deref'd unchecked. |
| [ ] | F-S0-5 | 018 | HIGH | Web | `EscapeJavaScriptStringUtf8("</script>")` output contains no `</script`; a JSON/JSONL/MD file with a `</script>` breakout payload does not execute injected script; generated docs carry a CSP. RED: today the escaper passes `</script>` through. |
| [ ] | F-S0-6 | 019 | HIGH | Web | `AsyncLoadProc` dereferences no mutable `self->_member` except `self->Release()`; all worker inputs come from a UI-thread snapshot; `fileSystem` AddRef happens on the UI thread. RED: today the worker copies `self->_fileSystem` off-thread. |
| [ ] | F-S0-7 | 020 | HIGH | Text | A stub reader returning M<N bytes yields no render/click/search hits in `[M,N)`; `_hexBytes.size()` == bytes actually read. RED: today the buffer is zero-padded to the reported size. |
| [ ] | F-S1-1 | 021 | HIGH | PE | `Close()` returns within a bounded timeout while a stub `Read` is blocked (no UI-thread join); reads chunked ≤1 MiB with a stop/requestId check between chunks. RED: today the UI thread joins the worker mid-`Read`. |
| [ ] | F-S1-2 | 022 | HIGH | VLC | `grep SetDllDirectoryW Plugins/ViewerVLC` returns nothing; libvlc still loads + plays via `LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR`. RED: today a process-global, non-additive `SetDllDirectoryW` runs on a worker for the whole session. |
| [ ] | F-S1-3 | 023 | HIGH | VLC | A rapid open/close stress loop does not crash; the player is detached (`set_hwnd(nullptr)`) on the UI thread during `WM_DESTROY` before `_hVideo` dies. RED: today `set_hwnd(nullptr)` is never called; stop/release races OS HWND destruction. |
| [ ] | F-S1-4 | 024 | HIGH | Text | On tabbed / CJK / emoji / surrogate lines, caret/selection/search-highlight/click land on the correct glyphs; caret never splits a surrogate pair. RED: today all geometry is `column * charW`. |
| [ ] | F-S2-1 | 025 | MED | Space | A capped scan keeps memory bounded AND reports exact rolled-up totals; children arena no longer abandons ~half its capacity on grow. RED: today the live model is unbounded (only the persisted snapshot is capped). |
| [ ] | F-S2-2 | 026 | MED | Sqlite | No `sqlite3_open_v2` per page/query (one reused read-only connection, threading verified/guarded); sequential paging shows correct non-dup/non-skip rows; header/table text sanitized. |
| [ ] | F-S2-3 | 027 | MED/LOW | ImgRaw | On `NativeMenuBarHost::Attach` failure the HMENU has exactly one owner (no double `DestroyMenu`); WIC-decoded images honor orientation; no duplicate concurrent decode; TIFF `count*size` can't overflow. |
| [ ] | F-S3-1 | 028 | MED/LOW | Text | Hex clipboard copy is bounded (notifies on truncation, can't throw out of `noexcept`); dead hex-fallback resolved (restored+tested or removed); `StreamOutCallback`/`_msftEditModule`/`DetectEncodingAndSize` removed (grep clean). |
| [ ] | F-S3-2 | 029 | LOW | ImgRaw, Sqlite | No user-facing literal remains on the ImgRaw export / Sqlite error paths; each has an `IDS_*` loaded via `LoadStringResource`; `LocalizationTests.exe` exit 0. |
| [ ] | F-S3-3 | 030 | LOW | all | Per stage: duplication replaced by one `Common` impl, behavior byte-identical (existing viewer/contract tests pass). Stage 1 fixes the ViewerPE alpha-drop; Stage 2 removes raw `new`/`delete` in the brush singletons. |

---

## Guiding Principle

> **Show exactly what's there, never block, never leak, never invent.** Never write a corrupt/blank file over the user's original; never read freed pixels; never let an untrusted data file become code; never fabricate bytes the file doesn't contain; never freeze the UI on a worker join; one balanced Release per AddRef; geometry that matches the glyphs DWrite actually draws. Prove the dangerous/dead/wrong path on the user's route, not just the happy path.

## Performance / Safety Validation Contract (mandatory)

Every P0 slice that touches destructive correctness (F-S0-1 export UAF, F-S0-2 atomic write) MUST assert **on-disk file state** (original intact / non-blank output / pre-existing file byte-identical on failure). Every concurrency/lifetime slice (F-S0-4, F-S0-6, F-S1-1, F-S1-3) whose race is impractical to force MUST be proven by a line-by-line AddRef/Release/ownership trace **plus** a structural assertion (grep/inspection: "no mutable `self->_member` in the worker", "exactly one Release per AddRef per exit path"). Build gates: `.\build.ps1 -ProjectName Viewer<Name>` then `.\build.ps1` (0 warnings/0 errors). Functional gate: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` (long; serialized by a machine-wide mutex — exit 3 means another selftest run holds it, retry). Register every new viewer case in the relevant harness (`Tests/ViewerPETests` drives all viewers; `Tests/ViewerSqliteTests` for Sqlite) so it actually runs in `Full`. Archive any perf evidence under `Specs/TestRuns/<machine>/Viewers/<timestamp>/`.

---

## Phase P0 — Ship-blockers: data-loss, dead capability, security, and the worst concurrency holes

### F-S0-1. `[CRITICAL / data-safety — export UAF]` ViewerImgRaw holds a cache reference across the modal save dialog (plan 014)

**Symptom:** Exporting a freshly-opened RAW while its full-resolution decode is still running in the background (the normal case) reads freed heap memory → crash, or silently encodes garbage/blank pixels into the user's chosen file.

**Root cause:** `ViewerImgRaw.Export.cpp:599-603` captures `const CachedImage* image = _currentImage` and `const std::vector<uint8_t>& bgra = …image->rawBgra/thumbBgra` (a *reference into the cache*), then `:648` calls the modal `ShowExportSaveDialog` (which pumps the message loop), then `:691-692` copies `pixels = bgra;`. During the pump, `kAsyncOpenCompleteMessage` → `OnAsyncOpenComplete` (`ViewerImgRaw.Decode.cpp:2888` `image->rawBgra = std::move(result->bgra)`, `:2900` `_currentImageOwned.reset()`) or `ClearImageCache` (`Decode.cpp:1384-1397`) invalidates or frees the referent.

**Fix:** Before the dialog, copy width/height into locals and the pixel bytes into an **owned** `std::vector<uint8_t>` under `_cacheMutex` (held only for the copy, never across the dialog); delete `pixels = bgra;`. No `image`/`bgra` use after `ShowExportSaveDialog`.

**Required proof:** Export a just-opened image: no crash, non-zero/non-blank output; grep confirms no `image->`/`bgra` after the dialog in `BeginExport`.

### F-S0-2. `[HIGH / data-safety — non-atomic export]` A failed encode destroys the overwritten original (plan 016)

**Symptom:** A mid-encode failure (disk full, device removed, encoder/GIF-palette error) leaves a zero-length/partial file; if the user overwrote an existing image (dialog uses `FOS_OVERWRITEPROMPT`), that original is permanently lost.

**Root cause:** `Export.cpp:397` `stream->InitializeFromFilename(outputPath, GENERIC_WRITE)` creates/truncates the destination in place before any pixels are written; failures at `:546/:553/:560` return after the file is already destroyed; `OnAsyncExportComplete` only shows a warning.

**Fix:** Encode to a sibling temp file in the destination directory; on success close the WIC stream then `MoveFileExW(MOVEFILE_REPLACE_EXISTING)`; delete temp on any failure via `wil::scope_exit` + a `committed` flag. Exemplar: `ViewerWeb.cpp:4119-4137`.

**Required proof:** Force an encode failure with a pre-existing destination present → destination byte-identical afterward, temp removed; happy path produces the file.

### F-S0-3. `[HIGH / dead capability]` WIC decode workers never CoInitialize → PNG/GIF/BMP/TIFF/HEIC silently fail (plan 015)

**Symptom:** Opening any WIC-backed image that isn't JPEG or a libraw RAW shows the error/no-image state. Masked because JPEG (turbojpeg) and RAW (libraw) never touch COM.

**Root cause:** `DecodeImageToBgraWic` (`Decode.cpp:664-684`) calls `CoCreateInstance(CLSCTX_INPROC_SERVER)`; it is invoked only from threadpool worker lambdas (`:2619` non-JPEG, `:2751` RAW fallback, `:2100` prefetch) that have no COM apartment. Only the *export* worker initializes COM (`Export.cpp:723`).

**Fix:** Add `CoInitializeEx(nullptr, COINIT_MULTITHREADED)` + a success-guarded `CoUninitialize` (the `Export.cpp:723` pattern) at the **top of `DecodeImageToBgraWic`** — one chokepoint covers all three call sites and tolerates `RPC_E_CHANGED_MODE`.

**Required proof:** A test opens a PNG (and ideally GIF/BMP) and observes it decodes, not the error state. RED today.

### F-S0-4. `[HIGH / RAII — leak + OOM crash]` Async dispatch leaks the COM self-ref on submit-failure and null-derefs on OOM (5 sites) (plan 017)

**Symptom:** On `TrySubmitThreadpoolCallback` failure the viewer's refcount never returns to zero — the whole object (D2D device, DWrite formats, readers, decoded-image cache) leaks for the process lifetime; ImgRaw additionally leaks the work item (and, for export, the moved-in pixel buffer). Genuine OOM crashes the host (unchecked `nothrow` deref).

**Root cause (one mistake, 5 instances):** `Decode.cpp` StartAsyncOpen `:2340/2352/2354/2796-2801`; prefetch `:1999/2010/2012/2220-2225`; `Export.cpp:705/716/718/783`; `ViewerText.cpp:4967/4979/4981/5684-5688`. Each does `AddRef()` → `new(std::nothrow)` (deref'd unchecked) → submit, with the balancing `Release()` only inside the work lambda's `scope_exit` (which never runs if the lambda never runs); ImgRaw additionally `ctx.release()`s unconditionally.

**Fix (fix the class):** Adopt ViewerWeb's correct shape (`ViewerWeb.cpp:3944-3949,3969-3973`) at all 5 sites: after the alloc `if(!ctx){Release();return;}`; in the `queued==0` block `{Debug::Error(...); Release(); return;}`; `ctx.release()` only on success. (Later generalized into a shared helper — F-S3-3 / plan 030 Stage 3.)

**Required proof:** Per-site AddRef/Release balance table across all three exit paths; happy-path viewer open/export unchanged in the harness.

### F-S0-5. `[HIGH / security]` ViewerWeb: untrusted file content breaks out of inline `<script>` via `</script>` (plan 018)

**Symptom:** Opening a malicious `.json`/`.jsonl`/`.md` file runs attacker JavaScript in the viewer's WebView2 origin (`https://viewer.redsalamander.invalid`) — exfiltration today (RCE surface if a host-object bridge is ever added). Stored-XSS-equivalent on inputs the viewer explicitly targets as untrusted.

**Root cause:** `EscapeJavaScriptStringUtf8` (`ViewerWeb.cpp:4981-5009`) escapes quotes/backslash/control chars but not `<` or `</script>`; escaped content is concatenated into inline `<script>` at `:4363` (pretty JSON), `:4432` (tree JSON), `:4496` (Markdown), `:1147-1161` (JSONL fields). No CSP; `NavigationStarting` blocks only top-level navigations, not `fetch`/`Image` sub-resources. (The find-query path is already escaped — out of scope.)

**Fix:** Escape `<` as `\x3C` (and U+2028/U+2029) in the escaper(s); add a restrictive `Content-Security-Policy` meta (`connect-src 'none'`; allow only the synthetic origin for script/style/img/font) to the generated JSON/JSONL/Markdown templates.

**Required proof:** A unit test asserts the escaper's output of a `</script>` input contains no `</script`; a harness test feeds a breakout payload and asserts no injected execution. RED today.

### F-S0-6. `[HIGH / concurrency — UAF]` ViewerWeb async worker reads `self->_fileSystem`/`_config`/`_theme` off-thread (plan 019)

**Symptom:** Navigating to another file (or toggling theme/config) while a slow/remote load is in flight can corrupt the IFileSystem COM refcount → double-free/UAF crash; torn theme/config reads otherwise.

**Root cause:** `AsyncLoadProc` (`ViewerWeb.cpp:3979-3996`) copies `self->_fileSystem` (com_ptr → AddRef) and reads `_config/_theme/_hasTheme/_markdownShowSource` on the threadpool thread while `Open()` (`:2321`) and SetTheme/SetConfiguration mutate them on the UI thread; `StartAsyncLoad` (`:3900-3977`) snapshots none of it; `AsyncLoadResult` (`ViewerWeb.h:94-107`) has no snapshot fields. (Unlike the other six viewers, which snapshot on the UI thread.)

**Fix:** Add snapshot fields to the payload; populate them on the UI thread in `StartAsyncLoad` (the `fileSystem` AddRef then happens on the UI thread, safely); `AsyncLoadProc` reads only the snapshot and uses `self` solely for `Release()`. No locks needed. Exemplar: `ViewerText.cpp:4959-4965`.

**Required proof:** `grep "self->_"` in `AsyncLoadProc` shows only `self->Release()`; single-open render unchanged.

### F-S0-7. `[HIGH / data integrity]` ViewerText hex view fabricates zero bytes on a short read (plan 020)

**Symptom:** The hex view shows phantom `00` bytes past the true end of data, lets the user select them, and an in-memory search for `\x00` returns matches that aren't in the file. Normal on remote/virtual filesystems (Curl/S3/GoogleDrive).

**Root cause:** `ViewerText.Hex.cpp:2095` `_hexBytes.resize(_fileSize)`, then the read loop `:2106-2121` `break`s on a short read with no resize-down; `ReadHexBytes:2295` uses `_hexBytes.size()` as the authoritative byte count. `IFileReader::Read` does not guarantee `bytesRead == requested`.

**Fix:** After the loop, `if (offset < _hexBytes.size()) { Debug::Warning(...); _hexBytes.resize(offset); }`. Optionally make scroll/click bounds use the loaded size for the in-memory path.

**Required proof:** A stub reader returning M<N bytes → no render/click/search hits in `[M,N)`; `_hexBytes.size()` == bytes read.

---

## Phase P1 — Reliability & correctness highs (each independent)

### F-S1-1. `[HIGH / perf — UI freeze]` ViewerPE joins the parse worker on the UI thread (plan 021)

**Symptom:** Closing/refreshing/navigating while a PE parse is in flight freezes the UI for the full duration of the in-flight `Read` — indefinitely on a hung network/cloud filesystem.

**Root cause:** `ViewerPE.cpp:1591` (`_worker = std::jthread()`) and `:1833` (reassign) both `request_stop` + **join** on the calling (UI) thread; reached from `WM_NCDESTROY:729`. The read loop `:1896-1918` checks `stop_requested` only between 16 MiB chunks, and a single `IFileReader::Read` is uncancellable. ViewerPE has **no** module keep-alive (so a naive `detach()` is unsafe vs DLL unload).

**Fix:** (a) chunk reads at ≤1 MiB with a stop/requestId check between chunks; (b) migrate to the threadpool + `AcquireModuleReferenceFromAddress` + `PostMessagePayload` + `requestId` model the other viewers use, so the UI thread never joins. Fallback: jthread + `request_stop` + `detach` **only after** adding a module keep-alive ref.

**Required proof:** With a stub FS whose `Read` blocks ~2 s, `Close()` returns within a few hundred ms (no UI-thread join); happy-path parse + navigation still work.

### F-S1-2. `[HIGH / security — DLL search]` ViewerVLC mutates the process-global DLL search path on a worker (plan 022)

**Symptom:** For the whole playback session, every thread's `LoadLibrary` in the host searches VLC's install dir (widened DLL-planting surface); two panes opening VLC concurrently race the global directory.

**Root cause:** `ViewerVLC.cpp:5003-5014` saves+sets process-global non-additive `SetDllDirectoryW(installDir)` on a threadpool worker, restored on the UI thread `:1093-1110`; libvlc.dll is loaded by full path `:5016` (so only its dependency resolution needs the directory).

**Fix:** `LoadLibraryExW(dllPath, nullptr, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS)`; delete `RestoreVlcDllDirectory` and the `previousDllDirectory`/`dllDirectoryWasSet` state. Fallback: scoped `AddDllDirectory`/`RemoveDllDirectory` cookie.

**Required proof:** `grep SetDllDirectoryW Plugins/ViewerVLC/ViewerVLC.cpp` empty; libvlc still loads and plays.

### F-S1-3. `[HIGH / concurrency — close race]` ViewerVLC never detaches the player HWND before async teardown (plan 023)

**Symptom:** Intermittent crash on rapid close (worse with hardware decoders): libvlc's video-output thread touches the already-destroyed `_hVideo` between OS destruction and the threadpool `stop`.

**Root cause:** `set_hwnd` is always bound to `_hVideo` (`:5244-5269`), never `nullptr`; `OnDestroy:2433` queues async `CleanupVlcState` (stop+release on the worker, `:1149-1153`) while the OS destroys the child HWND; the main proc has no `WM_DESTROY` case.

**Fix:** Add a `WM_DESTROY` handler that calls `set_hwnd(player, nullptr)` synchronously on the UI thread (parent `WM_DESTROY` fires before child `_hVideo` is destroyed); keep stop/release async.

**Required proof:** A rapid open/close stress loop does not crash.

### F-S1-4. `[HIGH / correctness — text geometry]` ViewerText assumes 1 code unit == 1 fixed column (plan 024)

**Symptom:** On any line with a TAB, a CJK/emoji/proportional glyph, or a non-BMP code point, the selection rectangle, caret, search-highlight, and mouse hit-test drift from the glyphs DWrite actually draws — clicking past a tab selects/copies the wrong text; the caret can land inside a surrogate pair.

**Root cause:** All geometry is `column * charW` (selection `Text.cpp:1565-1566`, caret `:1599`, draw of the raw buffer incl. `\t` `:1588-1589`, hit-test `:618/:630/:639`); `charW` is measured once from `"0"` (`:2862-2868`); no tab expansion in `RebuildTextLineIndex:1999-2018`; caret arrows step one code unit (`:1850-1882`).

**Fix (recommended Approach A):** per-visible-line `IDWriteTextLayout`; derive all geometry from `HitTestTextRange`/`HitTestTextPosition`/`HitTestPoint`; `SetIncrementalTabStop`; cache layouts for visible lines only. Make caret movement code-point-aware. (Approach B = tab expansion + column map fixes tabs+scroll but not CJK; still do surrogate-safe caret and document the CJK gap.)

**Required proof:** Tabbed/CJK/emoji/surrogate lines: caret/selection/click land on the correct glyphs; caret never splits a surrogate pair. (Extend the `ViewerTextDebugSnapshot` contract to expose caret X / selection rects for assertions.)

---

## Phase P2 — Reliability & performance mediums (each independent)

### F-S2-1. `[MEDIUM / reliability]` ViewerSpace live scan-model memory is unbounded (plan 025)

**Symptom:** Scanning a drive root grows `_nodes`/`_childrenArena`/`_fileRecords` to millions of entries (hundreds of MB–GBs), pushing the process toward `bad_alloc`/termination; the user never sees most of those nodes (the treemap already does level-of-detail).

**Root cause:** The only size guard (`MaxScanCacheSnapshotBytes`, `ViewerSpace.cpp:7278-7306`) gates only the persisted snapshot, not the live model. Arena grow abandons ~half its capacity (`~:7678-7720`); hardlinks are double-counted (`~:6497-6586`).

**Fix:** Enforce a live node/byte cap; on exceed, stop fine-grained itemization but keep rolling byte totals up to ancestors (**totals stay exact**) and flag "aggregated" for the UI; reclaim arena waste. Hardlink de-dup optional (memory-cost tradeoff).

**Required proof:** A capped run keeps memory bounded AND reports exact rolled-up totals; arena no longer abandons capacity on grow.

### F-S2-2. `[MEDIUM / perf]` ViewerSqlite reopens the DB per operation; deep paging is O(offset) (plan 026)

**Symptom:** Every table-list, page click, sort, and query reopens the SQLite file (header reparse + cold page cache); deep pages re-scan from row 0. (No SQL injection — identifiers are correctly quoted; no extension-loading RCE — flags disable it; pagination is deterministic on the read-only snapshot. Those are verified safe.)

**Root cause:** `DatabaseSource` holds only a path and calls `OpenReadOnlyConnection` (`Engine.cpp:172-192`) on every op (`ListTables:517`, each `LoadTablePage`, query, validate); `BuildTablePreviewSql:757-767` uses `LIMIT/OFFSET`. Header/table-name text skips `SanitizeCellText` (`:286-291`, `:552-555`).

**Fix:** Cache one read-only connection for the source lifetime (verify/guard the engine's threading — do not share a `NOMUTEX` handle across threads unguarded); keyset (seek) pagination with a stable `rowid` tiebreaker (handle `WITHOUT ROWID`); sanitize header/table text.

**Required proof:** No per-page `sqlite3_open_v2`; sequential paging shows correct non-dup/non-skip rows; unsanitized header control chars no longer reach the UI.

### F-S2-3. `[MEDIUM/LOW / correctness]` ViewerImgRaw: double-DestroyMenu, ignored orientation, double-decode, TIFF overflow (plan 027)

**Symptom & root cause:**
- **Double `DestroyMenu` (P2):** `ViewerImgRaw.cpp:3699/3707` the window owns the HMENU and `:1476` `_menuHandle` also owns it; only a successful `_menuBarHost.Attach` (`SetMenu(null)`) reconciles to one owner, but its result is discarded (`:1504`). On Attach failure both destroy the same handle. **Fix:** `if(!Attach(...)) SetMenu(hwnd,nullptr);`.
- **WIC EXIF orientation ignored** (`Decode.cpp:2622-2623` hardcodes orientation=1): apply the frame orientation for WIC formats.
- **Double decode**: the main open inserts into `_inflightDecodes` without the prefetch-side check (`~:2401-2404` vs `~:2032-2043`) — skip a decode already in flight.
- **TIFF `count*size` overflow** (`Decode.cpp:149/178`): overflow-safe checks (defense-in-depth).

**Required proof:** Single HMENU owner on Attach success/failure; a rotated WIC image displays upright; no duplicate concurrent decode; crafted huge TIFF count rejected.

---

## Phase P3 — Cleanup, localization & cross-cutting simplification

### F-S3-1. `[MEDIUM/LOW]` ViewerText: bound the hex clipboard, resolve the dead hex-fallback, remove dead code (plan 028)

- **Unbounded hex clipboard** (`Hex.cpp:2958-3030`): a huge selection builds ~half a GB synchronously on the UI thread and can throw `bad_alloc` out of a `noexcept` method (→ `std::terminate`). **Fix:** cap the copied selection (notify on truncation), guard against `bad_alloc`, ideally build off the UI thread.
- **Dead hex-fallback** (`ViewerText.cpp:5562` forces `result->hr = S_OK` immediately before the `:5564` `if (FAILED(result->hr) … allowHexFallback …)` guard, making the ~92-line "show as hex when text load fails" block unreachable): restore the feature (and test it) **or** delete the dead block.
- **Dead code:** remove `SaveCookie::StreamOutCallback` (`:8577-8622`), `_msftEditModule` (`.h:552`), and the never-called `DetectEncodingAndSize` duplicate (`:8999-9073`, `.h:464`) after a grep confirms no references.

**Required proof:** Bounded clipboard copy (test); dead-fallback resolved; grep for the three dead symbols is empty.

### F-S3-2. `[LOW / localization]` Move hardcoded user-facing strings into resources (plan 029)

ViewerImgRaw export messages (`Export.cpp` ~370…804, incl. the success alert `:799`) and ViewerSqlite `DescribeSqliteFailure` context + statement-guard literals are hardcoded English shown via host alerts/status — violating the no-hardcoded-strings rule. (This is **not** a path-leak/security issue — verified the SQLite error text does not embed the temp snapshot path.) **Fix:** add `IDS_*` to each plugin's `resource.h` + base `.rc` and load via `LoadStringResource`/`FormatEmbeddedStringResource`; drop the `ViewerImgRaw:` prefix from end-user text (keep it in `Debug::` logs). **Gate:** `LocalizationTests.exe` exit 0.

### F-S3-3. `[LOW / architecture — staged]` Consolidate duplicated viewer scaffolding into Common (plan 030; pairs with plan 005)

Behavior-preserving, staged, each independently shippable; the existing viewer/contract tests are the regression net:
- **Stage 1 (S):** shared ARGB→D2D color helper — also fixes the ViewerPE alpha-drop (`ViewerPE.cpp:251-257/:1660`).
- **Stage 2 (S–M):** shared class-background-brush helper (RAII; removes raw `new`/`delete` in ViewerSpace/Text/ImgRaw/Web — the genuine tech-debt behind the refuted "brush race").
- **Stage 3 (M):** shared async-dispatch helper encoding F-S0-4's hardened shape (do F-S0-4 first).
- **Stage 4 (M–L, maintainer-gated):** shared window-class/thunk + `Common/HwndD2DSurface` (device + device-lost); prototype on ViewerPE; exclude ViewerVLC (different render model). Do F-S1-1 first.
- **Stage 5 (M, optional):** shared `otherFiles` navigator.

---

## Verified correct — do NOT re-open without new evidence (23 refuted findings)

- **"God-files"**: line counts were overstated (`ViewerText.cpp` is 8,561 not 9,734; `ViewerSpace.cpp` 8,033 not 9,250); the Hex/Text/Encoding split is intentional and `Common/EmbeddedViewerBase.h` already factors the shared lifecycle. Tech-debt at most.
- **ViewerWeb `g_sharedEnvironment` "race" / DllMain teardown under loader lock** (several reports): every access is UI-thread-affine by WebView2's threading contract; `RedSalamanderPluginShutdown` runs the real teardown before `FreeLibrary`, so the DllMain call operates on already-nulled state (documented idempotent fallback). No reproducible race.
- **ViewerWeb file:// scripting + external navigation "exfiltration"**: by-design per `Specs/Plugins/Plugins_ViewerWeb.md` (browser-like HTML viewer); the real sub-resource gap is architectural to all file:// WebView2 viewers and is **not** fixed by `allowExternalNavigation`. (The genuine `</script>` injection IS F-S0-5.)
- **ViewerSqlite**: identifier quoting is correct (doubled quotes → **no SQL injection**); pagination is deterministic on the static read-only snapshot; open flags disable extension loading (no RCE); invalid-UTF-8 rows do NOT vanish (TEXT uses `sqlite3_column_text16`, not the lossy converter); the recycled-HWND posted-payload route is guarded by the mutex-protected `closedHwnds` set + `requestId`, not the GWLP_USERDATA pre-filter; error text does not leak the temp snapshot path.
- **ViewerImgRaw neighbor-cache prune freeing `_currentImage`**: refuted by three independent guards (the `updateOtherFiles=true` path is dead; the `_otherItems.size()<=1` early-exit; the non-cache `_currentImageOwned` lives outside `_imageCache`).
- **Global background-brush singletons concurrent-free; ViewerSpace post-shutdown null-deref; ViewerSpace `scope_exit`+`DestroyMenu`**: all UI-thread-affine or enforced by the host shutdown ordering (selftest-verified); `scope_exit` with `DestroyMenu` is the *approved* RAII idiom per `.github/skills/wil-raii`. (The raw `new`/`delete` in the brush singleton is real tech-debt → F-S3-3 Stage 2, not a race.)
- **ViewerImgRaw Ctrl+wheel stale menu/status; EXIF rational/ISO type-confusion; scalar BGRA repack loops**: refuted (repaint already covers status; the `bits 8/16` branch is already outside the loop and `/O2`+WPO unswitches the rest; metadata reads are `InRange`-guarded).
- **ViewerPE `CoInitializeEx` FAIL_FAST on an MTA host**: unreachable — the viewer open/command path runs on the STA UI thread (`S_FALSE`, not a failure).
- **ViewerText Save-As-of-diff "empty"** (`ApplyCurrentTextPresentation` repopulates `_textBuffer`); **caret double-draw at wrap/split boundary** (the boundary rect is zero-width / both spans already early-return); **hex backward-find re-finds current match** (the `selectionStart-1` bound excludes it): all refuted.
- **WndProcThunk/RegisterWndClass "missing InitPostedPayloadWindow" leak; ApplyTitleBarTheme dark-mode "bug"**: refuted — no plugin that posts payloads omits the init; ViewerSpace sets dark-mode attrs 19/20 via `ApplyImmersiveDarkMode` one line before `ApplyTitleBarTheme`.

---

## Concrete execution order

1. **P0, export cluster first, sequentially** (F-S0-1, F-S0-2, F-S0-4 ImgRaw part) — all three touch `ViewerImgRaw.Export.cpp`'s `BeginExport`/export worker; land one at a time and rebase. Then F-S0-3 (one-line-ish chokepoint), F-S0-7 (hex), F-S0-5 + F-S0-6 (ViewerWeb).
2. Re-run the viewer harness (`Tests/ViewerPETests`, `Tests/ViewerSqliteTests`, `Tests/PluginContractTests`) after each P0 slice; archive evidence under `Specs/TestRuns/<machine>/Viewers/<timestamp>/`.
3. **P1 in any order:** F-S1-1 (PE join), F-S1-2 (VLC dll-path), F-S1-3 (VLC hwnd detach), F-S1-4 (text geometry — the largest; may land incrementally, tabs first).
4. **P2:** F-S2-1 (Space memory), F-S2-2 (Sqlite perf), F-S2-3 (ImgRaw correctness bundle).
5. **P3:** F-S3-1 (Text cleanup), F-S3-2 (localization), F-S3-3 (consolidation — Stage 3 after F-S0-4, Stage 4 after F-S1-1; pairs with plan 005 Factory dedup).
6. **Closeout:** `.\Tools\Run-AllTests.ps1 -Suite Full`; update `Specs/Plugins/Plugins_ViewerPlugins.md` / `Plugins_ViewerWeb.md` for any lasting contract change (CSP, export atomicity, geometry); move this plan to `Specs/Plans/Done/`.

---

*Provenance: 89-agent adversarial multi-agent review of all 7 viewer plugins at `b274022d9`, 2026-06-16 (workflow run `wf_3d2e0ea2-a76`; 15 review lenses → per-finding refutation → advisor re-verification against source). 51 findings (1 critical, 13 high, 11 medium, 26 low) survived; 23 refuted (above). Every slice's evidence was re-read line-by-line before planning. Anchors are commit-relative — re-grep before editing. Per-slice executor-ready handoff files: `plans/014-*.md` … `plans/030-*.md` on branch `claude/thirsty-pascal-19c8f5`. Related: `Operation_Dragnet_SearchSubsystemRemediation_DataSafetyConcurrencyFolding_2026-06-15.md` (shared CV-destruction/UAF + bad_alloc→terminate classes), `Operation_Truename_BatchRenameRemediation_*_2026-06-16.md`, `Operation_Riptide_FairstreamRemediation_*_2026-06-15.md`, `Operation_Crosscut_CompareDirectoriesRemediation_*_2026-06-15.md`.*
