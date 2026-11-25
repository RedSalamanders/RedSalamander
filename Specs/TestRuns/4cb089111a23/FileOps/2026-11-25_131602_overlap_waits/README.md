# I66 automatic waits for overlapping file operations

The non-activating storage cases submit two intentional Parallel commands to the same
serialized Local route. Each command gets its own task and terminal callback. The first
publishes its verified payload and holds before retirement. The second stays behind that
exact task without a question, discovery, or publication. After release, Copy reaches
the ordinary Exists decision and skips. Move reports a missing source after the first
cross-volume Move removed it. Both cases check destination identity, payload and
sentinel bytes, the final file inventory, and unchanged foreground.

The full x64 Debug build at `be514e71e` emitted zero warnings and errors. Receipt
`98be23f45cc904df9f0c9092592a892b85c0b6a33e367fda4112135a3d9b5e4c`.

| Exact case | Run | Setup, case and cleanup | Native case suite duration |
| --- | --- | --- | ---: |
| `FileOps_RepeatedNonClipboardCopyStorage` | `20261125T115623Z-84976-791d701062034f1dbce71b3f054a58e5` | 3 passed | 1.719 s |
| `FileOps_RepeatedNonClipboardMoveStorage` | `20261125T120331Z-81896-8b91b5aae14d4a07bdaec57f945870da` | 3 passed | 4.250 s |
| `FileOps_ReviewRemediationContracts` | `20261125T115758Z-84976-305f16590f4444959f651ca03dba664d` | 3 passed | 83.594 s |

The earlier Move run `20261125T115713Z-84976-5482d3caa80e49838a1db7b18c57dbb3` skipped
the case because alternate-volume use was disabled. The retry used an existing marked
root registered for alternate-volume coverage. The 66 focused runner-plan checks
passed on `edaead548`. These focused runs provide no comparative throughput or full
native qualification verdict.
