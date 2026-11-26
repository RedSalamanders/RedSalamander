# File Operations Stopping allocations on x64 Release

The same Local Copy fixture reached the conflict metadata barrier on both builds. The test requested cancellation, observed Stopping before measurement, and required the task to stay unfinished with the same paths and completed bytes afterward. The worker stayed held throughout. Both exact `FileOps_PopupIdleCadence` runs ran without activation and passed setup, the cadence case and cleanup.

| Build | Source | Receipt | Native run |
| --- | --- | --- | --- |
| Pre-credential baseline with the measurement | `1c1971b4c9728afe3895c91459e4c8031d54f113` | `e6a704204057137e3e7f46660e3ab583a881ca51c400bb7232cf7213daa6e170` | `20261126T143852Z-69904-d980ec443cfe49c8a6d34a4e8c013ae9` |
| Integrated progress with the measurement | `95111d0d572a9b607188e6acb92282cbb9151d78` | `2b876f9e6598df999dc1cdb49fee0385899de3e2cb92b1863e8e172b49b9f505` | `20261126T144029Z-64380-f6b9c06845934b9faa4ba5a7645e179a` |

Both builds finished with zero compiler warnings and errors. The same `engine-stopping` measurement ran at a 100 ms popup cadence. The sampled `UpdateRates` detail matched as `tasks=7 updated=0 silent=0 rates=7`. Each selected sample recorded 25 heap allocations or reallocations under `UpdateRates` on the native UI thread.

| Build | Window | Timer ticks | Rate updates | Allocations or reallocations per update | Upper median `UpdateRates` |
| --- | --- | ---: | ---: | ---: | ---: |
| Baseline | 1,227 ms | 11 | 9 | 25 | 126 µs |
| Integrated | 1,219 ms | 11 | 10 | 25 | 92 µs |

The private process-scoped Xperf capture matched each stack through the exact build PDB and raw QPC conversion. Both captured windows reported zero lost buffers, zero lost events, and zero unmatched update stacks. The timing describes these traced runs on the same machine.

| Build | Raw capture archive | Analysis archive |
| --- | --- | --- |
| Baseline | `20261126T143850Z-46788-839822dd72674a1ab270099a8b0f5403` | `20261126T144204Z-44800-8339be99f96440d1b37ea157954bdf0d` |
| Integrated | `20261126T144027Z-59152-ee36a33af4954a9d9a6241de9b1192a4` | `20261126T144204Z-43136-da38b7d5c098470690992a96c0dbd744` |

The prior [Prompting, Running and Paused comparison](../2026-11-26_122011_popup_rate_allocations/README.md) covers the other matched windows. Live host authentication remains open.
