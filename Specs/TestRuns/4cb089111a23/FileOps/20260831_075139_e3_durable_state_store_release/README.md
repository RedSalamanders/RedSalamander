# E3 durable File Operations state-store validation

- Date: 2026-08-31
- Machine profile: `4cb089111a23`
- Source base at capture: `ea7e71f8fd61f2c932bbf6144bcd5e81a45ad6ce` plus the uncommitted E3 durable-store implementation under test
- Build: test-enabled x64 Release, 0 warnings and 0 errors
- Build receipt: `8800de0fb1134e8fd7cb9dfbc8a8c5a6ad2842c01ff1dfdef90ff5d97afe58b7`
- Source snapshot: `a5efffc571863cb9b2da47efe880857f23eda64f8b29f081f7ccde7a13aa08d9`
- Governed runner: `20260831T054742Z-89280-51019dc6d9764805a2e424d66e8095a7`
- Result: 8 passed / 0 failed / 0 skipped; Commands suite 2,456 ms, runner 7,551 ms; disk audit clean
- Raw bounded metrics: `perf/perf_metrics.jsonl` (796 valid JSONL rows, 296,409 bytes, SHA-256 `42548fffae344e47e0447c78ef9278ec827718f567f94fb3cfb6fcc4784bd372`)

The focused matrix protects the E3 durable-state mechanics used by Batch Rename recovery journals, File Operations move breadcrumbs, and the artifact registry. It covers atomic publication and rollback, immutable reader generations during serialized replacement, no-follow root and artifact handling, exact-handle removal, direct-child containment, byte and entry limits, malformed and invalid UTF-8 documents, schema rejection, canonical-name directory/reparse preservation, and recovery lifecycle behavior.

## Bounded characterization

The table reports successful metric rows only. Rows whose HRESULT intentionally records an injected negative-path result remain in the retained JSONL and are counted separately in the `all rows` column. P95 is nearest-rank.

| Metric | Successful / all rows | Average | P95 | Maximum |
|---|---:|---:|---:|---:|
| `batchrename.journal.load.us` | 31 / 41 | 2,380.6 us | 7,741 us | 9,789 us |
| `batchrename.journal.write.us` | 70 / 76 | 1,336.7 us | 1,914 us | 2,683 us |
| `fileops.artifact.registry.load.us` | 5 / 5 | 375.2 us | 1,146 us | 1,146 us |
| `fileops.move_breadcrumb.load.us` | 4 / 4 | 2,168.5 us | 4,410 us | 4,410 us |
| `fileops.move_breadcrumb.persist.us` | 4 / 7 | 1,259.0 us | 1,388 us | 1,388 us |

This archive is a same-machine candidate characterization for correctness, fault behavior, bounded retention, and the safety cost of the new mechanics. There is no apples-to-apples pre-E3 run for this exact scenario, so it makes no performance-improvement claim. The earlier E2 archive protects a different scenario and remains an independent baseline.

The original local evidence is under `D:\RedSalamander.Perf\runs\20260831T054742Z-89280-51019dc6d9764805a2e424d66e8095a7\artifacts\selftest\last_run`. The full metrics file was copied unchanged; the archived and raw SHA-256 values match.
