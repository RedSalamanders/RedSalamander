# Popup cadence diagnostic component observations

Two selected x64 observations used base commit 084085b367f7905e74d98b3252f83ccf68bd329d plus frozen uncommitted 29-file workspace 8ecc88a82986aa64ba3f2e9fd6fb27641367f704a25f3f0ca0f70ea81fbd9018. Exact source manifest SHA256 ce5789c5b8cfe5cbd531f00da35e9b7a654136c5aaed2ec4a4aac3d58897346f and diff SHA256 bdfa6292745c9b806840360cb5ceb3bb466834d2de3e8b474df6ff684c1dc502 remain immutable. Publication changes documentation/workspace identity, not that measured record. These are dirty-workspace component observations, not commit 084 or future-commit acceptance.

| Observation | Debug | Release |
|---|---|---|
| Native setup/selected/cleanup | 3 passed, 0 failed, 0 skipped | 3 passed, 0 failed, 0 skipped |
| Native exit / canonical runner exit | 0 / 2 filtered NOT_EVALUATED | 0 / 2 filtered NOT_EVALUATED |
| Real paused Copy | 2201 ms, 2 ticks / 2 renders / 2 owner ticks | 2203 ms, 2 ticks / 2 renders / 2 owner ticks |
| Synthetic retained / real observer rows | 58 / 89 | 50 / 97 |
| First sampled quiet graph subscription after pause, synthetic / real | 219 ms / 266 ms | 203 ms / 282 ms |
| Observer loss / reduced motion | 0 / false | 0 / false |

Both test-enabled full-solution builds recorded zero warnings/errors. Receipts ce7becae0418420d1d9724a75d0fdb9ed4200ac1e06636a6a5a2e4aac8195063 (Debug) and e76fc2419bf463636e06ff4d6480f74f6b6a6b30dc8229594b7050758923b949 (Release) were privately byte-verified before/after. Receipt bodies remain private: public projections do not establish portable qualified receipts. Raw metric commit fields are empty and unchanged; workspace, receipt, rotated prelaunch and archive joins supply source identity.

The existing exact FileOps_PopupIdleCadence case invokes the actual retained-rate reducer for nonzero-running-to-paused, preserving graph identity/history and requiring zero marker target, absent/already-paused rejection and observer default-off/stop/loss controls. Fixed 1024-row opt-in scalar observations sample actual pinned graph/host state and emit after measurement. Paused/known is the last synchronized snapshot, not a fresh engine read. graphQueried separates safe queries from suppressed reentrant queries. sampleTransition excludes separate marker transition; marker displayed/target and reduced-motion fields remain visible. Observer work adds overhead; first observed quiet points are sampled bounds, not exact transition latency.

Original 600 ms settle / 2200 ms sample targets, ceilings and resume assertions remain unchanged. Debug and Release are separate observations, not a performance pair. The historical hosted source-084 13-render failure was not reproduced here: its cause is not proved or fixed. No animation/product fix, uniform speedup, Full, ARM64, broad local desktop, SMB, applicable I60 NAS or future-source acceptance follows.

## Custody and payload

Six observer/metric/note files are byte-identical existing selected derivatives, not original full trace/JSONL. All 147 observer rows and 23 cadence metric objects per profile are bound to their raw source/order. Case projections preserve suite and all 3 declared statuses/durations/reasons; omitted and nested/renamed exit fields are explicit. evidence-bindings.json retains original lengths/hashes/line selections and source/receipt/plan/owned-native-exit/archive joins. Raw receipts, security-bearing intents and absolute paths are omitted; projections are not canonical completions, receipts, v2 results or promotions.

history-and-limits.json preserves historical hosted failure, genuine facade admission failure before any native launch and distinct private-driver setup failures. Debug's initialized private native_launched=false flag remains unchanged in custody; actual canonical results and owned PID exit 0 establish launch. Release derived it correctly. Both preserve nonblocking disk audit with 102 issues/no cleanup: this packet does not assert a globally clean root.

publication-manifest.json excludes itself; a private external manifest binds it. Generic archive gates enforce their normal contracts, not semantic validation of this projection format. Explicit source joins, closed-field privacy checks and independent review remain separate. See derivation.txt and publication-provenance.json.
The immutable source manifest is the seventh exact byte copy, separate from the six selected derivatives. Existing derivative serialization normalizes 11 metric timestamp spellings (one Debug, ten Release): for example .690Z becomes .69Z. Both strings identify the exact same UTC instant; all other metric fields are semantically exact. publication-provenance.json enumerates every profile, zero-based selected index, raw line and original/derived literal. Original raw streams and the six selected derivatives remain unchanged. These are exact copies of derived files, not exact native wire serialization.

Original native inventories cover 34 declared native/source/receipt/exit members. The two outer v1 runner aggregates are generated after native archival and were absent from those original intents. Their bytes match the earlier sealed final-proof references. Exact copies are now separately inventory-bound in a new current-time private custody context, with actual capture times and its own intent/completion identities. evidence-bindings.json separates original members from late_control_records and binds every layer. A path beneath completed-runs is not itself inventory membership; these late summaries are not original-native-sealed or portable qualification evidence. The first strict membership-check refusal is preserved; no native writer or runtime change was made.
