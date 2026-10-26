# Cached reopen 200: candidate-only target miss

Campaign `20261009T100439122Z-51ad5ed7` produced one candidate-only x64 Release run `20261009T100441Z-34428-9335f664dcd84ec0981fbfef73598c0f`. It is not a baseline comparison.

The focused case passed and collected all 200 samples. The sealed analyzer marks the sample set qualified for analysis, with all samples current-presented. That does not mean the latency target was met: first meaningful paint p50 was 189020 µs and p95 was 234411 µs against a 100,000 µs target; 0 of 200 samples met the target. Range: 157380–269427 µs. This is a measured target miss, not an overall performance pass or gain claim.

Receipt `753df43abafc1939cd72afb70983f1b4bb043a2bc1dd5a6ef5e90948cf77b158` uses build source snapshot `cbe50bd6f5762ef8010ef769ab202a544ecfd86d99e4ff163e101699c4031cd5`; workspace snapshot `cbe50bd6f5762ef8010ef769ab202a544ecfd86d99e4ff163e101699c4031cd5` is recorded independently. Fixture: 121 directories, 10000 files, 302230 root bytes; 200 cache hits.

The complete raw metrics file remains preserved outside this bounded archive at `D:/RedSalamander.Perf/evidence/viewer-space-memory-cachedreopen200-20261009T100439122Z-51ad5ed7/raw-sources/20261009T100439122Z-51ad5ed7-b0-p1-candidate/perf_metrics.jsonl`: 8,742,043 bytes, 23,575 rows, SHA-256 `f5d35c6d5ad8cc8f03a33e97545d0d68b8af13ab08683d022e3ab56af7eebbf5`. The archive contains the exact source index, curated metric rows, and replay map; the selected rows replay byte for byte against that external source. It also contains the sealed analysis and its source artifact, candidate build inputs, runner records, and all three exact reproduction helpers pinned by the campaign manifest.

manifest.txt records exact byte lengths, SHA-256 hashes, and external source paths for archived files. .gitattributes is curator metadata outside the payload manifest; * -text preserves bytes. The campaign has no baseline arm, so no comparative claim is supported.
