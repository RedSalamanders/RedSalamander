> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Monitor - verified findings (compact)

- **[data-loss/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:4665` - Copy of a huge selection: GlobalAlloc failure passes NULL to SetClipboardData (delayed-render trap) after the clipboard was already emptied
- **[data-loss/CONFIRMED]** `RedSalamanderMonitor/Document.cpp:904` - Non-atomic, truncating Save As destroys existing trace on partial write; failure silently ignored
- **[data-loss/PLAUSIBLE]** `RedSalamanderMonitor/RedSalamanderMonitor.cpp:2542` - 32-bit length truncation in UTF-8/UTF-16 file load corrupts or mis-sizes large (>2 GB) trace files
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/EtwListener.cpp:253` - Unbounded ingestion: every accepted ETW event is queued to ColorTextView with no rate cap or backpressure
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/Document.cpp:351` - Unbounded Document and ETW queue growth -> OOM under high event-rate flood (DoS)
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:550` - Unbounded ETW producer queue (_etwEventQueue) — second OOM/back-pressure-free path
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:539` - Unbounded ETW event queue and document growth under high event-rate flood -> OOM
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:5369` - Unbounded document growth: no line cap or eviction in Document under sustained ETW event flood -> OOM
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:4694` - Unbounded _matches vector from a common search term over a large/streaming document -> OOM / std::bad_alloc terminate
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:5293` - RebuildMatches full-document rescan on every append batch -> quadratic UI-thread blocking under ETW flood
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:562` - Unbounded document growth under ETW flood: no line cap or eviction -> OOM
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:550` - Unbounded ETW event queue growth under high event-rate flood -> OOM / UI starvation
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/Document.cpp:369` - Document line buffer has no scrollback cap or eviction -> unbounded growth -> OOM
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/RedSalamanderMonitor.cpp:2517` - Captured-trace file load reads entire untrusted file into memory with no size limit -> OOM
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/RedSalamanderMonitor.cpp:2517` - Whole-file slurp on Open causes unbounded allocation / OOM for large trace files (UI thread)
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/RedSalamanderMonitor.cpp:3700` - ETW event flood causes unbounded document + queue growth (OOM / terminate)
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/RedSalamanderMonitor.cpp:2517` - File-Open reads entire (untrusted) file into memory with no size cap (OOM)
- **[dos-oom/CONFIRMED]** `RedSalamanderMonitor/EtwListener.cpp:177` - Stop() WaitForSingleObject timeout is illusory; subsequent jthread join can hang the UI thread on teardown
- **[incorrect-rendering/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:645` - ClearText() leaves stale _caretPos/_selStart/_selEnd offsets after document is emptied; subsequent append + Ctrl+C copies wrong text and stale selection points past EOF
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/Document.cpp:101` - UINT32 overflow in line-offset and total-length accounting corrupts offset<->line mapping on large/streaming buffers
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/Document.cpp:182` - UINT32 character offsets/total-length overflow for large streaming logs
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/Document.cpp:101` - UINT32 character-offset accumulation overflows on long-running/streaming buffers, corrupting slice math
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/ColorTextView.cpp:645` - ClearText does not invalidate the layout/width sequence, so an in-flight worker packet installs a stale layout/slice state
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/ColorTextView.cpp:857` - FindNext / RebuildMatches store absolute document offsets in _matches and apply them to selection without revalidating after the document changed
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/ColorTextView.cpp:297` - Stale _matches drawn against a re-coordinated layout after Show-IDs / filter change (transient highlight coordinate mismatch)
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/ColorTextView.cpp:3320` - viewport->display-row mapping casts float to UINT32 out of range for very tall streaming documents (UB / wrong slice)
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/ColorTextView.cpp:4240` - InvalidateExposedArea truncates a 64-bit scroll delta to 32-bit LONG, producing a bad partial scroll-copy
- **[incorrect-rendering/PLAUSIBLE]** `RedSalamanderMonitor/ColorTextView.cpp:729` - TotalLength() (size_t) truncated to UINT32 in GetText() — wrong length for >4 Gi-char documents
- **[leak/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:4659` - Clipboard copy ignores OpenClipboard failure and leaks the HGLOBAL when SetClipboardData fails
- **[leak/PLAUSIBLE]** `RedSalamanderMonitor/EtwListener.cpp:72` - Orphaned real-time ETW session persists in the kernel if the monitor crashes
- **[race/CONFIRMED]** `RedSalamanderMonitor/EtwListener.cpp:195` - ProcessTrace handle pointer raced against Stop() writing _traceHandle = INVALID concurrently
- **[race/PLAUSIBLE]** `RedSalamanderMonitor/ColorTextView.cpp:3547` - Layout/width thread-pool workers use raw `self` and shared COM pointers without lifetime/synchronization guarantees (UAF + data race)
- **[robustness/CONFIRMED]** `RedSalamanderMonitor/Document.cpp:351` - Unbounded document growth under ETW flood — no cap/eviction leads to OOM
- **[robustness/CONFIRMED]** `RedSalamanderMonitor/EtwListener.cpp:170` - ETW Stop() can block the UI thread (join) for up to 5s on shutdown under event backlog
- **[robustness/CONFIRMED]** `RedSalamanderMonitor/ColorTextView.cpp:4665` - Clipboard copy of large selection: GlobalAlloc failure leaves bogus delayed-render handle on clipboard
- **[robustness/PLAUSIBLE]** `RedSalamanderMonitor/Document.cpp:594` - Lines() / VisibleLines() return raw vector references with no lock, exposing UI to use-after-realloc if the model is ever mutated off the render thread
- **[robustness/PLAUSIBLE]** `RedSalamanderMonitor/Document.cpp:583` - GetSourceLine/GetVisibleLine return a const Line& after releasing the shared_lock (lock-scoped dangling reference)
- **[robustness/PLAUSIBLE]** `RedSalamanderMonitor/Document.cpp:379` - UINT32 total-length and per-line offsets overflow on very large streaming buffers
- **[robustness/PLAUSIBLE]** `RedSalamanderMonitor/Configuration.cpp:41` - Configuration::Load() returns un-sanitized filterMask; validation lives only at call sites
- **[robustness/PLAUSIBLE]** `RedSalamanderMonitor/Configuration.cpp:71` - Configuration::Save() writes two registry values non-atomically with no rollback
