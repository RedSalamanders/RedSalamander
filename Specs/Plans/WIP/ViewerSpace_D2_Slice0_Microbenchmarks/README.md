# ViewerSpace historical Slice 0: portable microbenchmarks

These results are supporting evidence for the [consolidated ViewerSpace plan](../ViewerSpace_ConsolidatedRoadmapAndQualification_2026-10-22.md).
They are **not** the archived Windows before/after pair that
[Testing_PerformanceValidation.md](../../../Testing/Testing_PerformanceValidation.md) requires, and they do not replace it.
They were run on 2026-09-27 on Linux (4-core Intel Xeon at 2.80 GHz, clang 18.1.3, `-O2`). No Windows build
environment was available. They measure the platform-independent part of two Slice 0 changes: the
child-arena allocator (X15) and the frontier-lock publication path (X2).

## X15: child-arena free list

[`x15_arena_bench.cpp`](x15_arena_bench.cpp) ports the exact `AllocateChildArenaBlock` and
`ReleaseChildArenaBlock` bodies from `1dcd28c`, where the free list is a vector sorted by start, and
from `3fa79cf`, where it was indexed by an ordered map and set. The arena contents are replaced by a size
counter. The growth policy mirrors `AddRealNodeChild`: an initial block of 2, in-place growth at the tail,
and otherwise a doubled block with the old one released. The ceiling is the production
`maxChildArenaSlots`. Each case runs 5 times, and the median is reported. Every run also checks that both
allocators return the same sequence of block starts.

Workloads, up to the production ceiling of 300,000 retained child references:

- **realistic:** heavy-tailed fan-out, with children arriving interleaved across parents;
- **round-robin:** every parent grows in lockstep;
- **fragmenting:** every parent gets 3 children, then only even parents keep growing.

| Workload | Children | Old (ms) | Map/set (ms) | Peak free blocks | Same placement |
|---|---:|---:|---:|---:|---|
| realistic | 300,000 | 2.9 | 6.3 | 13 | yes |
| realistic | 150,000 | 1.3 | 3.1 | 18 | yes |
| round-robin | 300,000 | 2.0 | 6.8 | 1 | yes |
| fragmenting | 300,000 | 1.9 | 6.4 | 1 | yes |

Full output: [`x15_results.txt`](x15_results.txt).

**Conclusion: X15 was rejected, and the original allocator was restored.** Tail growth, coalescing and
best-fit reuse keep the free list at a handful of blocks. At that size the sorted vector is 2-4× faster
than ordered indexes, and both cost only milliseconds per 300,000 children. The plan's "potentially
quadratic" premise does not appear in any of these patterns. In ViewerSpace a folder's children also
arrive as one burst from a single enumeration, which is more benign than any of these workloads.

## X2: publication under the frontier lock

[`x2_frontier_model.cpp`](x2_frontier_model.cpp) is a threading model, not the plugin:

- workers pop jobs and publish under one mutex;
- enumeration is a busy spin;
- `PostUpdate` applies ViewerSpace's backpressure: a yield above 4 MiB pending, a 1 ms sleep above
  16 MiB and a 5 ms sleep above 64 MiB;
- a UI thread drains 512 records every 16 ms, as `DrainUpdates` does while scanning.

The "old" variant posts the directory's records while holding the mutex, as at `1dcd28c`. The "new"
variant releases the mutex first, as at `3fa79cf`. There are 20,000 directories and 4 workers.

- **light:** a 50 µs enumeration, 6 records of 96 bytes each. The queue never reaches the yield threshold.
- **backpressure:** a 30 µs enumeration, 8 records of 1 KiB each. Names and file summaries push the
  queue past 16 MiB.

Medians of three runs:

| Scenario | Variant | Folders/s | Lock wait p95 (µs) | Lock wait max (ms) | Worker time blocked on the lock |
|---|---|---:|---:|---:|---:|
| light | old | 72,900 | 0.3 | 0.4 | 1.9% |
| light | new | 74,300 | 0.3 | 0.2 | 0.8% |
| backpressure | old | 4,409 | 88 | 32.0 | 65.4% |
| backpressure | new | 4,416 | 0.2 | 0.12 | 0.0% |

Full output: [`x2_results.txt`](x2_results.txt).

**Conclusion.** X2 removes the stall: under backpressure the old workers spent about two-thirds of their
time waiting for the frontier lock while one of them slept holding it. X2 does not raise throughput,
because the UI drain caps both variants at about 4,400 folders/s. That matches the plan's F11 estimate
of about 4k retained folders per second. The next throughput lever is the drain and the record cost
(A10), not the frontier lock (A4).

## Reproduce

```sh
clang++ -std=c++20 -O2 -o x15 x15_arena_bench.cpp && ./x15 && X15_ROUND_ROBIN=1 ./x15 && X15_FRAGMENT=1 ./x15
clang++ -std=c++20 -O2 -pthread -o x2 x2_frontier_model.cpp && ./x2 4
```

## Historical Windows evidence request

The original September request listed the following pairs. Current reconciliation is owned by
R3 in the consolidated plan: existing later correctness passes do not automatically supply
these exact before/after measurements. This evidence README is not a second work queue:

- **X2:** `viewer_space_perf_small_progressive` and `cmd_viewer_space_queue_storm_stays_bounded` with
  `viewer.space.scan.frontier_lock_wait_us`.
- **X14:** idle timer wakeups, through `cmd_viewer_space_idle_timer_stops`.
- **X16:** grid hit tests after an unmoved relayout, through
  `cmd_viewer_space_unmoved_relayout_keeps_grid_hit_testing`.

Each needs a same-machine Debug and test-enabled Release pair archived under `Specs/TestRuns/`.
