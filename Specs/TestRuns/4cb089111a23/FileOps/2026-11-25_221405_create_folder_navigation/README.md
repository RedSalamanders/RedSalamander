# I66 Create Folder completion after pane navigation

The test-owned Dummy provider simulates one second per listed child. The first Create Folder request still offers the numbered sibling from the displayed listing, preserves the occupied directory and refreshes the admitting pane.

A second Create Folder request waits at the exact task's pre-mutation boundary. The pane reaches a different directory while that task is held. After release, the completion event reports success. The new folder exists under the original parent and the current pane stays at the unrelated directory with no matching entry.

The full test-enabled x64 Release build at `ce21ed01d94d94e8755e2a4b3aa38ff333ba7b14` had zero warnings and errors. `FileOps_DeferredExitRefusesNewWork` verified receipt `0c0e77c6bf3c6c600613f1f6a4fc6690272fbd25dce3500ff3d003a459ca6e2e` and passed setup, the case and teardown in non-activating run `20261125T211258Z-89608-9f21f98540ec4a3fb5da3f26b9a5c4fc`.

| Recorded event | Time (µs) |
| --- | ---: |
| First numbered suggestion | 42,062 |
| Direct listing of the same delayed parent | 2,049,802 |
| Second numbered suggestion | 30,311 |
| Release to exact completion after pane navigation | 5,036,768 |

These are individual controlled fixture observations. The runner closed its owned process and reported repository coverage `NOT_EVALUATED`.
