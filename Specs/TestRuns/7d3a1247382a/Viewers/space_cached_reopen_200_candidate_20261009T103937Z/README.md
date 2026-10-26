# Cached reopen 200: continuous-pump candidate-only target miss

Campaign `20261009T103937253Z-2bde3558` produced one candidate-only x64 Release run, `20261009T103939Z-45460-97527841ee64431d8d4f96a8020659ab`. The focused Commands case passed and the sealed analyzer qualified the 200 current-presented cache-reopen samples for analysis. This is not a baseline comparison.

The measured observer mode was `continuous-message-pump` with `pumpedWindowMs=50`, as recorded by the exact scenario artifact. First meaningful paint ranged from 110,480 to 203,813 µs; median 144,936.5 µs and p95 170,185 µs. Zero of 200 samples met the 100,000 µs target. This is a candidate run with a measured latency-target miss, not a memory-gain or overall performance pass. Results from a different observer mode are not combined or directly compared here.

The candidate build receipt is `4a245de29d61893189be0ac558a0c15f6989b4f9daf95201b721b2a7e8c135fe`, build source snapshot `f43deb78049356ccb5e5fff99ba0f9f90b355b24d62a985b7df124c77ce2aea9`; workspace snapshot is recorded separately in the run provenance. Fixture: 121 directories, 10,000 files, 302,230 bytes, and 200 cache hits. The archive contains the sealed analyzer result and exact three helper copies pinned by the campaign manifest.

The full raw metrics source remains at `D:/RedSalamander.Perf/evidence/viewer-space-memory-cachedreopen200-20261009T103937253Z-2bde3558/raw-sources/20261009T103937253Z-2bde3558-b0-p1-candidate/perf_metrics.jsonl`: 8,831,590 bytes, 23,818 rows, SHA-256 `6d6e2d6fbd9f8099bc743e04e581545e72c31f8116a5b27c039b9734446f52f3`. The archive retains the exact raw-source index, 1,002 curated metric rows, and line replay map; selected source rows replay byte for byte against the full raw file.

`observer-context.txt` records the observer fields from the scenario artifact without modifying the sealed analyzer output. `manifest.txt` records byte lengths, SHA-256 hashes, and exact source paths for archive payloads. `.gitattributes` is curator metadata outside that manifest; `* -text` preserves payload bytes.
