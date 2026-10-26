# File Operations card counts and throughput

Scope: the 19 October 2026 annotated-card request. Same-machine x64 Debug validation, external DxUi pin `e0260dde04695f0794eed79708030798b7575e0c`, actual 144 DPI. No desktop activation or global input was used for these captures.

The renderer places recursive counts under the transferred amount, ETA to their right, speed beside Throughput even with its plot hidden, and the effective automatic transfer limit under the buttons. The ordinary card's Technical details control and selected-root fraction are removed. Unsupported completion counts remain labeled found totals; unavailable counts never become zero. The separate Issues pane keeps its diagnostics.

## Runs

| Run | Evidence |
|---|---|
| `20261019T072659Z-70344-398d3ed721f9417f8ade7146a3a27fbc` | Unchanged baseline. Build: zero warnings/errors; ReviewRemediationContracts passed, 3 runner cases. |
| `20261019T074048Z-85384-78148e71bf3b49f09ef89a8c41606d31` | First candidate. All 100 captures and assertions passed, 3 runner cases. Build had four test-source warnings; the diagnostics and receipt are retained. The warning repair changes only UIA test local naming and formatting evaluation. |
| `20261019T074800Z-93124-e2af00322c84424c95b204bdd1774620` | Final candidate: zero warnings/errors; ReviewRemediationContracts passed, 3 runner cases, all 100 captures and assertions. |
| `20261019T075535Z-41660-3ab0bc0920b848c1bbeb0fad0764f626` | Live FolderLiveProgress: 3 runner cases passed, **64 behavioral assertions** across Copy, discovery, finalization, verification, empty and failure variants. Warning-free build; source snapshot matches the final candidate exactly. |

ResourceLocalizationContracts: **10 passed**. TestHarnessSourceContracts: **204 passed**. Their per-case JSON and transcripts are retained.

Each complete raw runner invocation is preserved beneath `D:\RedSalamander.Perf\evidence\fileops-card-20261019-*`. Checked-in receipts bind exact source, runtime and artifact identities. The final candidate's native Copy capture is also published in `docs/res/file-operations-popup.png`, with its independent provenance sidecar.

## Capture coverage

The card fixture covers English/French, light/dark, 480/960 DIP; original selected subject, compact/expanded and failed cards; known byte/file totals and zero-byte Copy; fully known recursive fractions, file-only completion support, found-only Move counts, unavailable counts, open totals before/after discovery reports, skipped counting, effective versus resolved concurrency, manual mode, empty known counts and a destination-stall label. It checks painted geometry and exact retained accessible count text. Card collapse removes count/concurrency regions, reopening restores them; plot collapse retains speed/ETA and the saved plot preference. The task actions retain their own controls and have no overlaps. Failure and conflict paths keep their existing priorities.

## Rendering measurement

`*-frames.jsonl` retains every `FileOps.Popup.CardHierarchy.FrameUs` row in original order. `metric-stream-provenance.json` binds the full raw streams' SHA-256, byte/row counts and retained stream SHA-256. The original streams remain in the marked sandbox; this archive does not substitute excerpts for the entire raw stream.

For a matched comparison, use the original scenes 0–12, all four language/width groups, and frames 2–11 of each scene (discard the first two frames). Baseline has 13 scenes/group; candidates have 24. This yields 520 samples per build. Percentiles use nearest rank. New scenes are validated but excluded from the matched comparison. Do not compare all candidate frames to all baseline frames as if the expanded fixture matrix were identical.

First candidate versus baseline: median **8.165 vs 8.167 ms**, p95 **8.534 vs 8.628 ms**. Final candidate: median **8.018 ms**, p95 **14.056 ms**. The longer final tail is retained and unexplained; identical renderer source between the candidates does not prove a cause or performance equivalence. `performance-comparison.json` contains all three sets. These are UI redraw observations under this harness; they establish no storage throughput improvement and are not a broad latency or memory-retention qualification.

## Qualification limits

These focused results are not Full or PrePush completion, Release/ARM64/ASan qualification, foreground keyboard/mouse acceptance, other DPI profiles, or storage/SMB benchmarks. The governed affected and PrePush explanations are retained; PR equivalence was unavailable, so no CI work was counted as passed. Broad interactive entries remain pending until an agreed desktop-input time. The existing test-root disk audit reports 272 unrelated issues and is not clean.

Run `Tools/Test-TestRunArchive.ps1 -RunPath <this folder>` before staging, then `-Inventory` for closeout. Preserve the first candidate warnings and original receipts; a clean repeat does not erase them.

Code and native test sources were unchanged after the final capture and live runs. Only user-guide image provenance and plan closeout changed afterward. No receipt from this archive authorizes reuse for a later workspace snapshot.
