# Synthetic PerfJsonl sink microbenchmark

This diagnostic run measures calls to `Debug::detail::WritePerfJsonl` only. It does not invoke `FileSystem.DirectoryOps.Enumerate`, read a directory, run ViewerSpace, or measure a scan. Do not use these results as a scanner performance result or claim a scan improvement.

Run ID: `space-perf-sink-microbench-20261009T054653Z-e77273ef5e5b416fb35bf9a0ee98bade`. The exact run `results.json` and microbenchmark source are copied as `results.json` and `source.txt`. The source calls `WritePerfJsonl` with a synthetic 1µs duration and a 300-character detail, collecting 1,000 calls per case. Two sink-on rounds each include one single-thread case and one eight-thread case (4,000 JSONL records total); then two sink-off cases run. There are zero warm-up calls and two clock reads around each measured call. The raw output is split only at the 2,000th JSONL row; concatenate `perf-raw-0001.jsonl` then `perf-raw-0002.jsonl` for the original bytes. `manifest.txt` records the original input and both shard hashes, lengths, offsets, and row ranges.

Observed measurements from `results.json`:

| Mode | Threads | Calls | Sum of call time | Maximum call | Batch wall time |
| --- | ---: | ---: | ---: | ---: | ---: |
| Sink on, round 1 | 1 | 1,000 | 10.3593 ms | 0.6555 ms | 10.3802 ms |
| Sink on, round 1 | 8 | 1,000 | 74.5121 ms | 6.0293 ms | 13.7952 ms |
| Sink on, round 2 | 1 | 1,000 | 10.4518 ms | 0.2029 ms | 10.4715 ms |
| Sink on, round 2 | 8 | 1,000 | 81.1745 ms | 6.8179 ms | 13.7233 ms |
| Sink off | 1 | 1,000 | 0.0191 ms | 0.0001 ms | 0.0374 ms |
| Sink off | 8 | 1,000 | 0.0215 ms | 0.0008 ms | 0.6609 ms |

These results show the observer cost for this synthetic direct sink workload and its 300-character detail under this host/build. They do not establish the cost fraction during a real `ReadDirectoryInfo` call or ViewerSpace scan. The recorded `FileSystem.DirectoryOps.Enumerate` metric name is only the synthetic event name passed to the logger.

Root reported the standalone source was compiled with Visual Studio 18 MSVC `/std:c++latest /EHsc /O2 /MD /W4`, with zero compiler warnings. The exact abfd build receipt payload is included as `build-receipt.json`; it identifies the Release `Common.dll` dependency, whose path/hash/length and the microbenchmark executable's path/hash/length are recorded in `binary-identities.txt`. Neither binary is copied. This receipt describes the app's Release build; it is not a standalone microbenchmark qualification receipt.

The evidence archive preserves the run as a logger-only diagnostic. It is not a full application build, filesystem enumeration test, or scan qualification.
