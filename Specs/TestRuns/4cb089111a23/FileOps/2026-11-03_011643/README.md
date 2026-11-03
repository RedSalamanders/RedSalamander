# R0e capability/admission containment: focused evidence

Fifteen fresh focused x64 Release processes measured `FileOps_ProviderCapabilityMatrix`: five each on main `269b06fc9f79f9fdd2bcbe1fce6423ed3b1424f1`, PR200 `1c31697a03962c81fda514f155fd39e446f880e2`, and PR201 `333c1a90c68da966e41451f4204f536eef183d4b`. All 15 native results contain exactly passed Setup, the selected case, and Cleanup_RestorePluginConfig. Wrapper exit 2 means filtered coverage and is not a Full pass. `packet.json` records exact run IDs, case statuses, receipts, source snapshots, raw hashes/lengths and per-sample metrics. Machine profile is `4cb089111a23`. All source-bound receipts are tests-enabled full-solution x64 Release, build number 1, Embedded.

The unchanged case-body digest is `ab2de533491c1816f373655e164c89a38ed38d3ec3d6de6b8f90502bbb3f9517`. Each measured in-process capability/admission containment sample performed 4,096 capability parses and 4,096 decisions, one uncontained-path rejection, and zero provider calls. This is not a network or general throughput test. All five samples per side passed their exact 2-second per-case guard and the existing paired threshold: nearest-rank p95 on five observations is the maximum; threshold is max(1.5×baseline p95, baseline p95 + 50,000 µs).

| Source | Five sorted durations (µs) | p95 (µs) | Paired threshold (µs) |
|---|---|---:|---:|
| main | 775252, 785280, 788857, 796476, 797404 | 797404 | baseline |
| PR200 | 791500, 797837, 822569, 836510, 852655 | 852655 | 1196106 |
| PR201 | 777634, 792923, 797139, 809638, 839815 | 839815 | 1196106 |

This small deterministic packet supports the specified threshold, not a statistically general latency or speedup claim. The selected JSONL files contain all `FileOps.SelfTest.R0e.*` rows in original order and exact source bytes; `packet.json` supplies original stream digest/size/row count plus each selected row's byte offset, length, line number and SHA-256. Retained legacy `*WorkUs` rows report `unit=count`; do not treat their `value` as a correctly typed duration. The claim uses `RouteContainmentUs`, whose rows have `unit=us`. Original native results, runner summaries, logs and full JSONL streams remain in governed external evidence; private paths are deliberately omitted here.

Runner summaries give an earliest start of 01:17:16.5106841 UTC and latest end of 01:34:19.0577444 UTC on 2026-11-03. The 01:16–01:41 UTC quiet window, reported complete native/cleanup drain at 01:34:40, and host release at 01:35:21 with peer acknowledgement are root coordination observations, not native measurement timestamps; `provenance.json` records that distinction. The five-process driver checked clean exact source and receipt/artifact before execution and again after each sample. Exact driver, copied receipt, workload-compatibility and wrapper-log hashes are in `provenance.json`; source snapshots are recorded separately from run-workspace projections.

The original private assessor failed in postprocessing because PowerShell treats `$Candidate` and `$candidate` as the same variable. No initial stderr file is invented. Its original bytes remain SHA-256 `84343a22104b66c8b36f8ca6b5b53dc332d6a2935a65cc567ecc1ba4febf204b`; the v2 assessor renamed the data variable and produced separate assessment files after all native processes had already passed. Both initial postprocessing failures remain failures, not native failures. No performance or full-qualification result is rebound to a later commit.

The full source-bound runner outputs, native results, traces, receipts and unabridged metrics remain in the governed external evidence store. This focused packet does not establish Full qualification.
`manifest.json` enumerates every other file in this packet; its own bytes are deliberately excluded to avoid a self-referential digest.
