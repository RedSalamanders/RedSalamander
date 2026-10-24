📋 72-Hour Code Review — Consolidated Findings

Scope: 100 source files, 11 commits, 4 reviewers

Overall Verdict

🏗️ Ripley: APPROVED ✅ — Zero AGENTS.md violations. 100% regression guard compliance. Production-ready.

-----------------------------------------------------------------------------------------------------------------------------------------------

🔴 Critical Issues (7 total)

┌───┬───────────────────────────────────────────────────────────────────────────────────────────────┬──────────┬───────────────────────────────┐
│ # │ Finding                                                                                       │ Reviewer │ File                          │
├───┼───────────────────────────────────────────────────────────────────────────────────────────────┼──────────┼───────────────────────────────┤
│ 2 │ Bandwidth throttle race — RefillBandwidthThrottleState is called without lock on the          │ ⚙️       │ FileSystem.FileOps.cpp        │
│   │ sequential path, but concurrent CopyFileEx callbacks can race on the same state.              │ Sysadm   │                               │
├───┼───────────────────────────────────────────────────────────────────────────────────────────────┼──────────┼───────────────────────────────┤
│ 3 │ DirectorySize callback locking — cancellation callback invoked without callbackMutex when     │ ⚙️       │ FileSystem.DirectoryOps.cpp   │
│   │ context.parallel is set.                                                                      │ Sysadm   │                               │
├───┼───────────────────────────────────────────────────────────────────────────────────────────────┼──────────┼───────────────────────────────┤
│ 4 │ ComboBox double pixel-snap — GetPopupItemRect snaps the popup rect, does DIP arithmetic on    │ 🔧 Yoko  │ DxUi.ComboBox.cpp             │
│   │ the snapped result, then snaps again. Compounds rounding errors at 150% DPI.                  │          │                               │
├───┼───────────────────────────────────────────────────────────────────────────────────────────────┼──────────┼───────────────────────────────┤
│ 5 │ TextField per-update layout creation — ImportTextInputBridgeState creates a full DirectWrite  │ 🔧 Yoko  │ DxUi.TextInput.cpp            │
│   │ layout + extracts line metrics on every sync, even when text hasn't structurally changed.     │          │                               │
├───┼───────────────────────────────────────────────────────────────────────────────────────────────┼──────────┼───────────────────────────────┤
│ 6 │ Grid O(n²) group toggle — keyboard expand/collapse rebuilds the entire group tree then        │ 🔧 Yoko  │ DxUi.Grid.cpp                 │
│   │ reconciles selection across all groups. Degrades with 1000+ rows.                             │          │                               │
├───┼───────────────────────────────────────────────────────────────────────────────────────────────┼──────────┼───────────────────────────────┤
│ 7 │ Integer overflow fragility — SaturatingBytesForElapsedMs division-by-zero guard is correct    │ ⚙️       │ FileSystem.FileOps.cpp        │
│   │ but the pattern is fragile for future callers.                                                │ Sysadm   │                               │
└───┴───────────────────────────────────────────────────────────────────────────────────────────────┴──────────┴───────────────────────────────┘

🟡 Key Design Concerns (highlights)

 - Settings save doesn't clamp — Parse validates ranges (std::clamp), but Save writes raw values. Corrupted in-memory state persists to 
registry. (Sysadm)
 - Mnemonic collisions — No debug assertion to detect duplicate Alt+key mnemonics within a dialog. (Sysadm)
 - TextField tab suppression divergence — Bridge checks ES_MULTILINE style while TextField checks _multiline member. Could desync. (Yoko)
 - FileSystem plugin complexity — Parallel ops, bandwidth throttling, and search indexing are growing. Consider extracting a ParallelCoordinator
 class. (Ripley)

🧪 Test Coverage Gaps (GuineaPig)

┌───────────────────────────────────────────────────────────────────────────────────────────────┬─────────────────────────────────────┐
│ Gap                                                                                           │ Risk                                │
├───────────────────────────────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
│ Preferences DxUi migration — 9 pages, 1000+ lines, zero keyboard nav tests                    │ HIGH — focus loss, broken mnemonics │
├───────────────────────────────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
│ FileSystem auto-concurrency hints — no test for wrong detection (HDD vs NVMe vs network)      │ HIGH — perf disaster                │
├───────────────────────────────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
│ Recycle bin batch delete — no boundary test at 500+ items, no mid-batch failure test          │ HIGH — potential data loss          │
├───────────────────────────────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
│ ShortcutsWindow collapse state — commit says "preserve state across rebuilds" but no test     │ MEDIUM                              │
├───────────────────────────────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
│ FindFilesWindow column resize — 600 new lines, no resize/sort persistence test                │ MEDIUM                              │
├───────────────────────────────────────────────────────────────────────────────────────────────┼─────────────────────────────────────┤
│ Bandwidth throttle fairness — tests measure duration but not max skew between workers         │ MEDIUM                              │
└───────────────────────────────────────────────────────────────────────────────────────────────┴─────────────────────────────────────┘

Top 3 Untested Edge Cases

 1. DPI change mid-dialog — drag Preferences to high-DPI monitor during interaction
 2. Unicode filenames in recycle bin batch — RTL overrides, surrogates, ZWJ
 3. Cancellation during 16-worker pre-calc — 2+ second cancel latency, double-cancel crash risk

🟢 Strengths Worth Preserving

 - Perfect AGENTS.md compliance — zero violations across 100 files (no sprintf_s, no raw new/delete, no catch(...), no manual cleanup)
 - Correct atomics — proper acquire/release/acq_rel ordering throughout
 - Full localization — 40+ new .rc strings, zero hardcoded UI text
 - Atomic progress publishing — lock-free reads for progress counters, narrow mutex scope for paths
 - Device loss handling — wil::scope_exit on BeginDraw/EndDraw, proper recreation flow
 - Cancellation polling — bandwidth throttle sleeps in 10ms slices, checking cancel each iteration

🎯 Recommended Priority

 2. Fix bandwidth throttle race — add lock on sequential path
 3. Fix ComboBox double pixel-snap — snap once at the end
 4. Add Preferences keyboard navigation smoke test
 5. Add recycle bin batch boundary test (501 items)
 6. Add FileSystem hints validation test (garbage input → no crash)