# I66 live SFTP scheduled Rename availability

The test-enabled x64 Release build on `da902726fbbf88e1056919efef2965097295bf28` completed in 14 minutes 38 seconds with zero warnings and errors. Verified solution receipt `d2c0a6e11397b14d2250d14d1a634b5beb71a26e5235e617f8c0c152527bb094`. The first fresh-checkout build stopped before compilation because the x64 vcpkg inputs were absent. Restoring the pinned tool and packages unblocked the build. Static, formatting and changelog checks passed.

## Live result

Run `20261127T112818Z-80416-382ec7f9c17f4be1871a69120075e5af` used the registered SFTP profile and one fresh owned child. The exact `Phase16_RemoteSftpInlineRename` run passed setup, witness and cleanup without activation. Native exit 0 and counts 3 passed, 0 failed and 0 skipped.

The Files menu capability query reported that scheduled commands were unavailable. Both direct commands showed the non-blocking pane warning with `ERROR_NOT_SUPPORTED`. The guard ran before either command could open a dialog or create a task.

| Measurement | Duration (µs) | Status |
| --- | ---: | --- |
| Batch Rename warning | 871 | `S_OK` |
| Change Case warning | 254 | `S_OK` |
| Whole owned SFTP Rename witness | 790,815 | `S_OK` |

F2 moved the selected file, changed its exact case, and renamed a folder with its child. Readback proved the expected payloads. The occupied-name provider control reported a known refusal before dispatch and preserved both files. Cleanup verified payloads, removed only the owned entries nonrecursively, and checked that the empty fixture was gone. The isolated Hello-off profile retired, and the installed settings digest matched.

Raw native result SHA-256 `cb5b1bd0119779510bf483ee4df0256e618da7da6dcccfbe4b41bb7ea139fcae`. Launcher report SHA-256 `86b1d57fcda81cb85ac0babf12c5bf9f0b77a6b5f1190b7ab00a7d8fcf8a0d15`. Raw metrics SHA-256 `1d9d99853acbf438f323b04748886680595afaff0b6ec664eb5e2e28c90a6e84`.

The same x64 Release receipt also passed the non-activating provider capability matrix in `20261127T112909Z-85344-7ea2674a639740a3b736c23e2d822ebb` with 3 passed, 0 failed and 0 skipped. The runner's filtered repository verdict was `NOT_EVALUATED`.

Positive scheduled SFTP qualification remains open because the provider does not expose the exact binding required by those commands.
