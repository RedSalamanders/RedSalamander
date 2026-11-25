# I66 Pack worker

The exact archive command case holds both encoders immediately before commit. The pane has control while the task is active and a previous archive is readable. Cancel preserves its bytes and retires only the private sibling by handle. A subsequent overwrite succeeds. The exact regular-file cleanup test also holds the binder and verifies that a competing writer cannot open the published archive.

| Profile and selection | Source | Build receipt | Governed run | Passed | Failed | Skipped |
| --- | --- | --- | --- | ---: | ---: | ---: |
| x64 Debug archive commands | `778040402cf3b5ec9fef49256e9a72e84b99346a` | `b7faa2127a058448cf7168a6f7e23281cdb3e8c51a9fd5cdbecfcee28f8993cb` | `20261125T204634Z-73188-dd7b8c96716f42b6ac978429f9edf92c` | 1 | 0 | 0 |
| x64 Release archive commands | `19f0245d82acc4cc047ce1bf72bc5985e65f6e3c` | `4b3f3a7e4be1382bd285dc91bde6d727556a56acadf6972241d41af1e85dc598` | `20261125T205904Z-90976-af528cc7a30d4ffab38f690fd3765fe4` | 1 | 0 | 0 |
| x64 Release deferred exit | `19f0245d82acc4cc047ce1bf72bc5985e65f6e3c` | `4b3f3a7e4be1382bd285dc91bde6d727556a56acadf6972241d41af1e85dc598` | `20261125T210045Z-5860-4e86acd6bc99441e978ea07353defb05` | 3 | 0 | 0 |

Both full test-enabled builds had zero warnings and errors. All selections ran without desktop activation. The Release source includes waits for the archive dialog tests to see task completion. Those foreground dialog selections and wider repository qualification remain open.

The first stored ZIP included 16 entries and 23 bytes in both the [transactional baseline](../2026-11-25_211805_pack_publication/README.md) and the worker candidate. These are same-machine fixture observations in microseconds.

| Profile | Baseline on pane thread | Candidate pane admission | Candidate worker execution | Candidate flush and publish |
| --- | ---: | ---: | ---: | ---: |
| x64 Debug | 8,791 | 1,286 | 7,903 | 1,883 |
| x64 Release | 4,499 | 384 | 7,197 | 3,016 |

The admission and encode metrics came from distinct threads in both candidate profiles. The measurements describe one fixture and do not set a throughput budget.
