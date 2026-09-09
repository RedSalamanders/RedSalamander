> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


I'll write the final audit report directly from the verified findings. No source exploration is needed — the findings are pre-verified with evidence and corrected severities. Let me produce the report.

# RedSalamanderMonitor Security & Robustness Audit — Final Report

## 1. Verdict

RedSalamanderMonitor has **no confirmed memory-corruption, use-after-free, or out-of-bounds vulnerability** in the audited surface — every "memory-safety/UAF" candidate was downgraded on inspection because downstream consumers clamp defensively (GetTextRange/GetLineAndOffsetUnsafe bounds-check, HitTestTextRange is fed construction-bounded ranges, the global `g_colorView` outlives its workers). That is the good news. The bad news is that the tool is **trivially DoS-able and leaks/loses data under entirely realistic conditions**, and its safety rests on **unenforced invariants** (UI-thread affinity, "every caller remembers to mask") rather than structural guarantees. The most dangerous themes, in priority order: **(a) zero ingestion backpressure end-to-end** — any local process can register the compiled-in provider GUID and flood an unbounded queue → unbounded append-only document → OOM/`std::terminate` through a `noexcept` WndProc and an SEH-only top-level handler that *cannot* catch `std::bad_alloc`; **(b) teardown that blocks/hangs the UI thread and orphans a kernel ETW session**; **(c) clipboard and Save-As paths that silently destroy user data on the failure path**; and **(d) a pervasive 32-bit offset model** that is correctness-fragile at scale. The single highest-leverage fix is a bounded scrollback + bounded queue with drop-oldest accounting — it neutralizes roughly half the findings at once.

---

## 2. Findings by Severity

### MEMORY-SAFETY / CRASH / USE-AFTER-FREE / SECURITY

> Audit result: **none confirmed.** Every candidate in this class was downgraded. The two structural hazards that come closest are retained below because they are the real residual memory-safety risk if the threading model ever changes — they are *latent UAFs gated behind one future off-UI-thread mutation*.

**Async layout/width thread-pool workers capture raw `self` with no cleanup-group join (latent UAF + COM data race)**
`RedSalamanderMonitor/ColorTextView.cpp:3547` (also `:3520`, `:3599`, `:4408`, `:4463`) · severity **race** (corrected from use-after-free) · CONFIRMED structure / PLAUSIBLE exploit
Both `StartLayoutWorker` and `EnsureWidthAsync` submit via `TrySubmitThreadpoolCallback(worker, rawCtx, nullptr)` with **no `TP_CALLBACK_ENVIRON` / cleanup group**, capturing a raw `ColorTextView* self`; the worker dereferences `self->_dwriteFactory` / `self->_textFormat.get()`. The destructor (`:174-179`) only stores `nullptr` into `_hWndAtomic` — it does **not** cancel or wait for in-flight callbacks.
Failure scenario: at process-exit static destruction of the global `g_colorView`, a still-running layout worker calls through `_textFormat`/`_dwriteFactory` after those `wil::com_ptr` members have been released → torn read / call through a released interface. (The dramatic shutdown-UAF and font-change-race are *defended* today: the object is a global that isn't freed at `WinMain` return, `PostMessagePayload(nullptr,…)` no-ops, and `SetFont` has no live caller.) The seq counters are checked only on the UI thread, never inside the worker, so they provide **no** worker-lifetime guard.
Deeper fix: own the workers' lifetime with a `TP_CALLBACK_ENVIRON` + cleanup group; in the destructor call `CloseThreadpoolCleanupGroupMembers(…, TRUE, …)` to cancel/wait before any member is destroyed. Pass AddRef'd COM interface copies into the `Ctx` instead of reading `self->…`, and never mutate those members on the UI thread while a worker can read them.

**Unlocked reference-returning accessors defeat the reader-writer lock (latent UAF if any mutator moves off the UI thread)**
`RedSalamanderMonitor/Document.cpp:594` (`Lines()`/`VisibleLines()`), `:583` (`GetSourceLine`), `:568` (`GetVisibleLine`) · severity **robustness** (corrected from use-after-free) · PLAUSIBLE
`Lines()`/`VisibleLines()` return raw `const&` to the internal vectors with **no lock**; `GetSourceLine`/`GetVisibleLine` take a `shared_lock` but **return a `const Line&` after the lock is destroyed at function return**. Every neighboring accessor locks `_rwMutex`, and the class advertises "Thread-safe with reader-writer lock" (`Document.h:50`) — a **false guarantee**. The render path holds such a reference across further document calls (e.g. `ColorTextView.cpp:3457`, `:4016`).
Failure scenario: the instant any future code appends from the ETW worker / a search / a self-test thread, a concurrent `AppendInfoLineUnsafe` `push_back` (`:369`) reallocates `_lines` while the UI thread dereferences a stale `Line&` → use-after-free. Safe **only by accident** today because all appends are marshaled to the UI thread via `OnAppEtwBatch`.
Deeper fix: do not return references that outlive the lock. Return by value, or expose a callback-under-lock API, or reuse the existing lock-carrying `DisplayTextBatch` (`Document.cpp:803`) which already does this correctly. Remove the "thread-safe" claim until the reference-returning accessors are fixed. **Root-cause shared with the layout-worker finding: the entire safety model is "everything happens on the UI thread," unenforced.**

---

### DOS / OOM

> **Single root cause, many symptoms.** Findings across `EtwListener`, `ColorTextView`, and `Document` all reduce to: **the ingest pipeline has zero backpressure and the document has no scrollback cap.** Deduplicated below into the three distinct links in one chain, plus the file-load variant.

**Unbounded ingest chain: no rate cap, unbounded queue, unbounded append-only document → OOM / `std::terminate`**
Listener: `EtwListener.cpp:253` (every accepted event forwarded, provider enabled `TRACE_LEVEL_VERBOSE` + keyword `0xFFFF…FFFF` at `:102-106`). Queue: `ColorTextView.cpp:550` (`_etwEventQueue.push_back`, unbounded `std::deque`, `ColorTextView.h:495`; drain capped at 200/msg only, `:5333-5358`). Document: `Document.cpp:369` (`_lines.push_back`, documented "append-only" `Document.h:193`, no evict/trim/cap; `ReserveForAdditionalLines` only grows, `:168-180`). Wiring: `RedSalamanderMonitor.cpp:3700`. · severity **dos-oom** · CONFIRMED
*(Deduplicates 11 separately-filed findings reporting the same chain from different files/lines: ColorTextView.cpp:539/550/562/5369, Document.cpp:351/369, EtwListener.cpp:253, RedSalamanderMonitor.cpp:3700.)*
Security-relevant reachability: the provider GUID is **compiled into the binary**, and the accept filter (`MonitorDiagnostics.h:17-25`) only rejects the monitor's *own* PID — so **any other local process can register as that provider and emit at high rate.** The "post only when queue was empty" gate (`:549-552`) and the 200/batch cap bound *message-loop posts and per-frame UI work*, **not memory**.
Failure scenario: a foreign process spins `EventWrite` in a tight loop. The worker thread enqueues faster than 200-events-per-UI-message drains; `_etwEventQueue` and `_lines` (plus `_lineOffsets`, `_visibleLines`, per-line cached strings, `_lineWidthCache`) grow monotonically. UI thread is **starved first** (perpetual repost), then `std::bad_alloc` is thrown from `push_back`/`reserve` **inside a `noexcept` WndProc** (`:5148`/`:1185`) → `std::terminate`. The top-level `__except(EXCEPTION_EXECUTE_HANDLER)` in `wWinMain` (`:3473-3492`) is **SEH-only and cannot catch the C++ exception.**
Deeper fix (one change closes the whole class): impose a **bounded scrollback** on `Document` (configurable max line count and/or byte budget; ring-buffer / drop-oldest `pop_front` with offset/`_visibleLines`/`_lineOffsets`/caret/selection/`_matches` rebase) **and** cap `_etwEventQueue` depth in `QueueEtwEvent` with drop-oldest + a surfaced `_eventsLost` "N events dropped" marker. Lower the enabled trace level/keyword mask. Make the append path tolerant of `bad_alloc` rather than relying on never reaching it.

**File-Open slurps an entire untrusted file with no size cap (3–4× peak, UI thread) → OOM / terminate**
`RedSalamanderMonitor/RedSalamanderMonitor.cpp:2517` (`ReadFileAsTextUTF` whole-file `istreambuf_iterator` read; second wide buffer at `:2526/2536/2545`; called synchronously from `DoFileOpen` `:2568` → `SetText` `:2577`, dispatched at `:3822`) · severity **dos-oom** · CONFIRMED
Failure scenario: opening a multi-GB saved capture allocates ~file_size (bytes) + up to 2× (decoded wstring) + a third per-line copy in `SetText`, all on the UI thread with no try/catch; `std::bad_alloc` escapes through the same SEH-only handler → terminate, losing any live capture. Short of OOM, the synchronous read blocks the UI thread (violates the no-UI-blocking rule). Reachability caveat: operator-initiated via the Open dialog, not remote — so "availability/operator-footgun," not a remote vuln.
Deeper fix: `GetFileSizeEx` first and reject/stream above a configurable cap; read in bounded chunks off the UI thread and append progressively via `AppendText` with a NUL/encoding-boundary-aware decoder; wrap in try/catch and surface a clean error dialog.

**ETW Stop() join can hang the UI thread on teardown (illusory 5 s timeout) + orphans the kernel session**
`RedSalamanderMonitor/EtwListener.cpp:177` (also `:170`, `:140`, `:195`) · severity **dos-oom** (corrected from robustness) · CONFIRMED
`Stop()` waits `WaitForSingleObject(handle, 5000)` (`:170`) and on `WAIT_TIMEOUT` **only logs a warning** (`:173`), then executes `_workerThread = std::jthread()` (`:177`). The move-assignment runs the old `jthread`'s destructor → `request_stop()` + **unconditional blocking `join()`**. The `request_stop()` is **inert**: the worker discards its `stop_token` (`:140`) and `ProcessTraceThread` blocks in `ProcessTrace` (`:195`); the per-event drain (`HandleEvent`/`ExtractEventData`, ~17+ TDH calls/event) checks neither `_isRunning` nor the token. `Stop()` runs on the **UI thread** from `OnDestroyMainWindow` (`RedSalamanderMonitor.cpp:4036`).
Failure scenario: under a flood / stuck real-time buffer `ProcessTrace` doesn't promptly return after `CloseTrace`; the UI thread stalls ≥5 s then performs the unbounded join — a multi-second-to-hung exit. Because the join precedes the error-path teardown, `ControlTrace(EVENT_TRACE_CONTROL_STOP)` may never run → the named kernel session is orphaned. (Reachability caveat: ETW real-time buffers are finite, so "indefinite" is in practice "multi-second stall"; the join eventually completes.)
Deeper fix: make the worker observe the stop request promptly (check `std::stop_token`/an atomic in the per-event callback and bail out of the drain). On bounded-wait timeout, treat as hard failure — do not rely on `jthread`'s blocking destructor as the backstop; ensure `ControlTrace(…STOP)` still runs so the session isn't orphaned. Prefer moving teardown off the UI thread.

---

### DATA-LOSS

> **Shared root cause across two findings: the failure path runs *after* the destructive step, and the failure is never surfaced.** The clipboard double-allocation finding is the SAME bug as the leak finding below; merged there.

**Clipboard copy on `GlobalAlloc`/`OpenClipboard` failure wipes the clipboard and registers a bogus delayed-render handle**
`RedSalamanderMonitor/ColorTextView.cpp:4673` (the unconditional `SetClipboardData`; `EmptyClipboard` at `:4662`, alloc at `:4665`, unchecked `wil::open_clipboard` at `:4659`) · severity **data-loss** · CONFIRMED
`EmptyClipboard()` destroys the prior contents *before* allocation; the copy is guarded by `if (globalMem)` (`:4666`) but `SetClipboardData(CF_UNICODETEXT, globalMem.release())` (`:4673`) runs **unconditionally**. On `GlobalAlloc` failure `release()` yields NULL → `SetClipboardData(…, NULL)` is the **delayed-rendering idiom**, promising the window will supply data via `WM_RENDERFORMAT`/`WM_RENDERALLFORMATS` — handlers that **do not exist** in the Monitor (verified absent). `open_clipboard`'s result is never checked, so on a contended clipboard `EmptyClipboard`/`SetClipboardData` are also called when the clipboard was never opened.
Failure scenario: copy a large selection (or any copy under memory pressure / fragmentation — multi-GB not required) → prior clipboard wiped, `CF_UNICODETEXT` left registered with no renderer → any later paste yields empty/garbage; the user believes the copy succeeded.
Deeper fix: check the `open_clipboard` result and bail before `Empty`/`SetClipboardData`; defer `EmptyClipboard()` until after a successful `GlobalAlloc`+copy; only call `release()` after `SetClipboardData` **succeeds** (otherwise let the `wil::unique_hglobal` free it); bound pathological selection sizes.

**Non-atomic, truncating Save-As destroys the existing trace on partial write; failure silently swallowed**
`RedSalamanderMonitor/Document.cpp:904` (`std::ofstream(path, std::ios::binary)` truncates on open; write loop `:909-930`; `return file.good()` `:931`) · severity **data-loss** · CONFIRMED
The destination is opened directly with implicit `trunc`, zeroing the existing file the instant it opens; there is **no temp-file + atomic `ReplaceFileW`/`MoveFileExW` rename**. If any write fails or the process dies mid-loop, the original is already destroyed and left partial. `DoFileSaveAs`'s bool is **discarded** at `RedSalamanderMonitor.cpp:3823` with **no error dialog** (contrast `DoFileOpen`, which shows `IDS_MSG_OPEN_FAILED_READ`).
Failure scenario: operator saves hours of capture over an important log on a USB stick that fills or is yanked mid-write → both old and new data lost, **no error shown**, user believes the save succeeded.
Deeper fix: write to a sibling `.tmp`, check `file.good()` after each write and at close, then atomically swap via `ReplaceFileW` or `MoveFileExW(MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)` only on full success; delete the temp and return false on failure. In `DoFileSaveAs`, check the return and show an error dialog mirroring `DoFileOpen`.

---

### LEAK

**Clipboard `HGLOBAL` leaked when `SetClipboardData` fails (same root cause as the clipboard data-loss finding)**
`RedSalamanderMonitor/ColorTextView.cpp:4673` · severity **leak** · CONFIRMED
`globalMem.release()` relinquishes ownership *before* `SetClipboardData`'s result is known; on failure ownership is **not** transferred to the clipboard, so the released `HGLOBAL` of `(sel.size()+1)*sizeof(wchar_t)` bytes leaks, recurring on every failed copy. **Same fix as the data-loss finding** — release only after success; this is one code change, not two.

**Orphaned real-time ETW kernel session persists if the monitor crashes / is force-killed**
`RedSalamanderMonitor/EtwListener.cpp:72` (named session `RedSalamanderMonitor_ETW_Session`, real-time mode `:63`, 8–32 MB buffers `:65-67`) · severity **leak** · PLAUSIBLE
ETW sessions are kernel objects that outlive the process; the only teardown paths (`ControlTrace(…STOP)` at `:50` pre-clean and `:187` `Stop()`) require normal execution, and `~EtwListener` does not run on force-kill. `_sessionHandle` is a raw `TRACEHANDLE`, not a WIL RAII type, and there is no crash-cleanup (`SetConsoleCtrlHandler`/top-level handler). The next normal same-privilege launch self-heals via the pre-clean — **but** the code itself handles `StartTrace` → `ERROR_ALREADY_EXISTS` after the pre-clean returned `ERROR_ACCESS_DENIED` (`:77-83`), proving the indefinite-orphan path (created elevated, relaunched non-elevated/different user) is real, wasting non-pageable pool.
Deeper fix: complement the pre-clean by stopping the session on all exit paths (`SetConsoleCtrlHandler` / top-level handler / app-scope RAII); log/surface when the pre-clean STOP fails; document the admin step to reclaim an elevated orphan (`logman stop`). *(Note: this overlaps the Stop()-hang finding — both leave an orphaned session; fixing teardown ordering addresses the normal-exit half.)*

---

### INCORRECT-RENDERING / ROBUSTNESS

**Pervasive UINT32 offset/total-length model overflows on >4 Gi-char buffers → non-monotonic `_lineOffsets`, corrupted slice/selection/copy**
`RedSalamanderMonitor/Document.h:199` (`_cachedTotalLength`), `:201` (`_lineOffsets`); accumulation at `Document.cpp:101` (`EnsureOffsetsValid`), `:182` (`EnsureTotalLengthValid`), `:379`/`:396` (`AppendInfoLineUnsafe`); read path `:742` (`upper_bound`), `:745` (`position - lineStart` underflow); `GetText` truncation at `ColorTextView.cpp:729` (`static_cast<UINT32>(TotalLength())`) · severity **incorrect-rendering** · PLAUSIBLE
*(Deduplicates 6 separately-filed UINT32-overflow findings — Document.cpp:101/182/379, ColorTextView.cpp:729 — all one root cause.)* All cumulative character positions are 32-bit and accumulate with no overflow check; `_lines` is append-only with no eviction, so the precondition (>4.29 G UTF-16 positions ≈ 8.6 GB resident, realistically 16–30 GB with per-`Line` overhead) is not prevented by code. After wrap, `_lineOffsets` is non-monotonic → `upper_bound` returns a wrong line, `position - lineStart` underflows, `TotalLength()` returns a tiny wrapped value, and `GetText()` silently truncates to the low 32 bits. **Not memory-unsafe** — every consumer clamps (`GetLineAndOffsetUnsafe` returns `min(off,lineLen)`; `GetTextRange::appendSlice` `min`-clamps against `line.text.size()`), so the result is wrong content, not OOB. `GetText()` additionally has **no live caller** today (the real save-all path `SaveTextToFile` iterates `_lines` with `size_t` and is 64-bit-safe).
Deeper fix: widen `_cachedTotalLength`, `_lineOffsets` element type, offset locals, and display-row accumulators to `size_t`/`UINT64`; where a UINT32 API boundary is required, saturate/reject positions beyond `UINT32_MAX` rather than truncate; assert `_lineOffsets` monotonicity before the binary search. **Enforcing the bounded-scrollback cap from the OOM section also removes the precondition entirely.**

**`ClearText()` leaves stale caret/selection offsets → wrong text copied to clipboard after refill**
`RedSalamanderMonitor/ColorTextView.cpp:645` · severity **incorrect-rendering** · CONFIRMED
`ClearText()` resets layout/cache/scroll and calls `_document.Clear()` but **never resets `_caretPos`/`_selStart`/`_selEnd`** (contrast `SetText` `:478-480` and `EnableShowIds` `:302-304`, which do). After `IDM_FILE_NEW` (`RedSalamanderMonitor.cpp:3820`) clears and the stream re-grows past the old offsets, Ctrl+C (gated only on `_selStart != _selEnd`, `:5603`) → `CopySelectionToClipboard` copies `GetTextRange(s, e-s)` over brand-new unrelated content. Internally bounds-clamped, so **data-confusion, not OOB**: the user copies text they never selected and the caret renders at a position they never navigated to.
Deeper fix: reset `_caretPos = _selStart = _selEnd = 0` in `ClearText()` (and `_matchIndex = -1`); better, make `Document::Clear()`/`SetText()` the single source of truth and re-clamp all three offsets to `TotalLength()` after any document-shrinking operation. `AppendText` should also re-clamp.

**`ClearText`/filter/Show-IDs changes don't invalidate the layout/width seq or `_matches` → stale slice/highlight installed against the new document**
`RedSalamanderMonitor/ColorTextView.cpp:645` (`ClearText` doesn't bump `_layoutSeq`/`_widthSeq`), `:297` (`EnableShowIds`) and `:339` (`SetFilterMask`) don't clear `_matches`/call `RebuildMatches`; seq gate at `:5250`/`:5437`; `_matches` push at `:4694` · severity **incorrect-rendering** · PLAUSIBLE
*(Deduplicates three related staleness findings — ClearText seq gap, Show-IDs/filter coordinate-space mismatch, and FindNext stale-offset reuse — all "a document/coordinate mutation does not invalidate async-worker or search state, and only async layout-ready eventually rebuilds it.")* An in-flight worker packet (`seq==N`) survives `ClearText` (which leaves `_layoutSeq==N`) and installs pre-clear slice positions; toggling Show-IDs/filter changes every line's `PrefixLength` while `_matches` still hold old-coordinate offsets, drawn by the next paint before `RebuildMatches` lands; `FindNext` (`:857-860`) sets selection from absolute match offsets with no re-clamp. **All are bounds-safe** — `ApplyColoringToLayout` guards `run.sourceLine >= TotalLineCount` (`:3667`), `drawMatch` clamps `localStart` by construction so `HitTestTextRange` is never fed an out-of-range position, and `_matches[_matchIndex]` is always reduced mod size. Consequence is transient mis-positioned highlights/selection/hit-test that self-heal on the next async `OnLayoutReady`.
Deeper fix: bump `_layoutSeq`/`_widthSeq` and clear `_layoutCache` in `ClearText`/`SetText`; call `RebuildMatches()` (or clear `_matches`) **synchronously** in `EnableShowIds` and `SetFilterMask`; clamp `_selStart`/`_selEnd`/`_caretPos` to `[0, TotalLength()]` in `FindNext` and invalidate `_matchIndex` when a document-revision counter advances. **Do not rely on `HitTestTextRange` clamping for bounds safety of search-derived offsets.**

**Unbounded `_matches` from a 1-char query over a large document → OOM / `std::terminate`**
`RedSalamanderMonitor/ColorTextView.cpp:4694` · severity **dos-oom** · CONFIRMED
`RebuildMatches` (`:4676-4716`) iterates every source line and unconditionally `_matches.push_back(...)` (24 B/elem) with **no cap**; reachable via live-find with **no minimum query length** (`WM_TIMER` `:5226` → `PerformFindLiveUpdate` → `RebuildMatches` on every keystroke). A single common char (`' '`, `'e'`, `'0'`) over a multi-hundred-MB streaming log yields tens-to-hundreds of millions of matches → multi-GB allocation on the UI thread; `bad_alloc` propagates through the **`noexcept` WndProc** → `std::terminate`.
Deeper fix: cap match count (e.g. `kMaxMatches = 100k`), break out of inner+outer loops on cap, surface a "too many matches" state; require a minimum query length before live search; wrap the build so allocation failure degrades gracefully.

**`RebuildMatches` full-document rescan on every append batch → quadratic UI-thread blocking under flood with an active search**
`RedSalamanderMonitor/ColorTextView.cpp:5293` (called from `OnAppLayoutReady`), rescan loop `:4711` · severity **dos-oom** · CONFIRMED
With a search active, every async layout completion (driven by each append batch via `EnsureLayoutAdaptive(1)` `:5408`) clears and re-scans the **entire** append-only document under no time budget → O(N²) total work on the UI thread, rebuilding `_matches` from scratch each time. Combined with the unbounded document, this converts a flood into a hung window before OOM.
Deeper fix: make match maintenance **incremental** — on append scan only the newly-appended lines and append their matches (offsets are append-stable in the non-filtered case), preserving `_matchIndex` by start offset; bound per-frame search work and/or debounce `RebuildMatches` off the hot append path.

**Float→UINT32 viewport-row cast is UB / clamp-after-cast at five sites (very tall documents render wrong region)**
`RedSalamanderMonitor/ColorTextView.cpp:3320` (also `:3370`, `:4139`, `:4477`, `:347`) · severity **incorrect-rendering** · PLAUSIBLE
`static_cast<UINT32>(std::floor(viewTop / rowH))` is UB once the quotient reaches 2³²; MSVC `cvttss2si` yields `0`/`0x80000000`, and the trailing `std::min(topRow, maxDisplayRow-1)` clamps the *value* but cannot recover a garbage-collapsed cast. Bounds-safe (document accessors are checked), but scrolling near the bottom of a ~268 M-row document would render the top while the thumb is at the bottom. Reachability is practically gated by OOM and by the upstream UINT32 display-row counter overflowing first.
Deeper fix: compute in `double`, clamp to `[0, maxDisplayRow-1]` as a `double`, then cast — apply the guarded conversion at all five sites.

**64-bit scroll delta truncated to 32-bit `LONG` before the magnitude check → bad partial scroll-copy**
`RedSalamanderMonitor/ColorTextView.cpp:4240` · severity **incorrect-rendering** · PLAUSIBLE
`static_cast<LONG>(std::lround(deltaDipY * dpiScale))` truncates a `long long`; the `absDelta < viewHeightPx` guard (`:4245`) then sees the truncated-small value and takes the partial scroll-copy branch instead of a full redraw, leaving a stale-content strip. Requires a ~2³¹-DIP true delta (≈130 M+ resident rows) for the low 32 bits to land in the narrow viewport-height window — latent, near-measure-zero reachability.
Deeper fix: compute and compare in 64-bit and saturate; never down-cast before the magnitude check.

**`ProcessTrace(&_traceHandle,…)` raced by `Stop()` writing `_traceHandle = INVALID` before the worker joins**
`RedSalamanderMonitor/EtwListener.cpp:195` (worker holds member address), `:158` (UI-thread non-atomic write before join at `:170`); `_traceHandle` non-atomic at `EtwListener.h:85` · severity **race** · CONFIRMED
The worker passes `&_traceHandle` (the member's address, not a local copy) to `ProcessTrace`, which reads `*phArray` for the call's lifetime; `Stop()` writes `_traceHandle = INVALID_PROCESSTRACE_HANDLE` on the UI thread **before** joining → a formal C++ data race (UB, TSan/ASan-reportable). **Impact bounded**: the object outlives both threads, so no UAF/OOB/corruption — worst case is the race UB and a possible unexpected `ProcessTrace` return code.
Deeper fix: sequence `CloseTrace` → **join** → only then reset `_traceHandle = INVALID` (move `:158` after the join); or make `_traceHandle` `std::atomic` / copy into a local before `ProcessTrace`. Prefer a WIL/RAII trace-handle type whose reset happens after the join.

**`Configuration::Load()` returns an un-sanitized `filterMask`; safety depends on every caller masking**
`RedSalamanderMonitor/Configuration.cpp:41` (raw DWORD copied to `filterMask`), `:48` (`lastFilterPreset` raw cast) · severity **robustness** · PLAUSIBLE
A user-writable `HKCU\…\Monitor\FilterMask` is copied verbatim with no `& kMonitorFilterAllMask`; the only sanitization is at the lone consumer (`RedSalamanderMonitor.cpp:3357`). Bounds-safe today (`_filterMask` is only ever bit-tested at `Document.cpp:522`/`:559`, never used to index), so this is a latent invariant-at-the-source gap, not a live bug.
Deeper fix: mask inside `Load()` (`filterMask = filterMaskValue & kMonitorFilterAllMask;`) and validate `lastFilterPreset` to `{-1,0,1,2,3}`; keep call-site masking as defense in depth.

**`Configuration::Save()` writes two registry DWORDs non-atomically with no rollback**
`RedSalamanderMonitor/Configuration.cpp:71`/`:77` · severity **robustness** · PLAUSIBLE
A kill or hive failure between the two `set_value_nothrow` calls leaves `FilterMask` updated but `LastFilterPreset` stale; on next `Load` the preset wins (`RedSalamanderMonitor.cpp:3367`), so the user sees a filter they never chose. Deterministic, non-corrupt outcome — cosmetic; lowest priority.
Deeper fix: write both under one `RegCreateKeyTransacted` transaction, or re-validate the `(mask, preset)` pair on `Load` and fall back to defaults on inconsistency.

---

## 3. Cross-Cutting Themes

1. **Zero backpressure, zero retention bound.** The dominant theme. From the kernel session (`VERBOSE`/all-keywords) → unbounded `std::deque` → append-only `Document` with no scrollback cap, **nothing** in the pipeline bounds memory. One bounded-scrollback + bounded-queue change (with a surfaced dropped-event counter) neutralizes the entire DOS/OOM cluster *and* removes the precondition for the UINT32-overflow rendering bugs.

2. **`bad_alloc` is a guaranteed crash, not a degradation.** Allocation failure on any append/search/file-load path throws through a **`noexcept` WndProc** and an **SEH-only top-level handler that cannot catch C++ exceptions** → `std::terminate`. Every OOM finding escalates to a hard crash for this structural reason. Fixing the handler boundary (catch `std::bad_alloc` at the dispatch boundary and degrade) is a force-multiplier independent of the bounds work.

3. **Safety by unenforced invariant.** Memory safety rests entirely on "all appends and reads happen on the UI thread" (defeated by the unlocked `Lines()`/`GetSourceLine` accessors and the false "thread-safe" doc claim) and on "every config caller remembers to mask." Both are footguns: the moment a future change appends off-thread or reads `filterMask` raw, a latent UAF / mis-filter becomes live. Enforce invariants structurally (lock-carrying APIs, sanitize-at-source) rather than by convention.

4. **Failure paths execute *after* the destructive step, silently.** Clipboard (`EmptyClipboard` before alloc; unconditional `SetClipboardData(NULL)`) and Save-As (truncate-on-open; discarded return; no error dialog) both destroy user data *then* fail without telling the user. The pattern fix is "validate/stage before you destroy, and surface every failure."

5. **Pervasive 32-bit position model.** `UINT32` total-length, line offsets, display rows, scroll deltas, and the `GetText` cast all share one width-too-narrow design. Bounded by clamping everywhere (hence incorrect-rendering, not memory-safety), but correctness-fragile at scale. Widen to `size_t`/`UINT64` end-to-end or cap the document so positions never approach 2³².

6. **Teardown is not robust.** The `Stop()` join can stall/hang the UI thread, and abnormal exit orphans the named kernel session — two findings, one underlying "no prompt cancellation + no crash-cleanup + raw handle" weakness.

---

## 4. Coverage & Residual Risk

**TDH internal trust boundary (untested — primary residual risk).** This audit found **no confirmed memory-safety defect in the ETW property-parsing path**, but `ExtractEventData` performs 17+ `Tdh*` calls per event (`TdhGetEventInformation` + per-property `TdhGetPropertySize`/`TdhGetProperty`) on the worker thread, and the report's verdict treats **TDH as trusted**. That trust is unverified here. A malicious local provider controls the `EVENT_RECORD` payload, and `PropertyLength`/`UserDataLength` values feed the per-property reads. **The single most important follow-up is a crafted-provider / ETW fuzzing campaign** that emits malformed schemas, zero-length and oversized property descriptors, mismatched `UserDataLength`, deeply nested/array properties, and non-UTF-16 string payloads — to confirm that unchecked TDH-reported lengths cannot drive an OOB read into the event buffer. This is the one place an OOB could plausibly still live; it was outside what static review could confirm.

**Flood / DoS testing.** All OOM/quadratic findings are CONFIRMED structurally but characterized analytically, not measured. Need a **load harness**: a foreign process registering the compiled-in provider GUID emitting at MHz rates (verbose, large messages, 1-char-search-active) to (a) reproduce queue/document unbounded growth, (b) confirm the `bad_alloc`→`terminate` path through the SEH boundary, (c) measure the `RebuildMatches` O(N²) UI-stall, and (d) validate that a bounded-scrollback/bounded-queue fix actually caps RSS with correct drop accounting.

**Teardown-race testing.** The `_traceHandle` data race (`:195`/`:158`) and the `Stop()` join-hang (`:177`) need **TSan/ASan + stress-close**: repeatedly start/flood/close, and force `ProcessTrace` to lag `CloseTrace` (large real-time backlog) to reproduce the multi-second/hung exit and verify the kernel session is not orphaned. Also exercise force-kill-mid-capture to confirm orphaned-session reclaim behavior (elevated-then-non-elevated relaunch).

**Latent-UAF activation testing.** The two retained "race/robustness" findings (raw-`self` workers, unlocked reference accessors) are safe only under current UI-thread affinity. Any future off-UI-thread append must be paired with a **deliberate test that appends from a worker while a render/measure loop holds a `VisibleLines()`/`GetSourceLine()` reference**, under ASan, to prove the lock contract before shipping. Until then, add a UI-thread-affinity assertion so violations fail loudly.

**Scale testing for the 32-bit model.** The UINT32-overflow and float-cast findings are PLAUSIBLE precisely because 4 Gi-char / 268 M-row preconditions are hard to reach before OOM. If bounded scrollback is adopted, these become unreachable; if not, they need a **high-memory soak test** (or synthetic offset injection) to confirm the non-monotonic-`_lineOffsets` corruption and the wrong-region render empirically.

**Not covered / out of scope here.** Provider-side `EnableTraceEx2` filtering correctness, multi-provider subscription interactions, DPI-change re-layout under flood, and the `ViewerWeb`/other-component changes in the working tree were not part of this Monitor audit surface.
