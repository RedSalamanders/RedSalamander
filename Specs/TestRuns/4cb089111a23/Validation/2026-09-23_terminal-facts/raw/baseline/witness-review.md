# FileOps suite completed-summary projection witness

This revises the earlier Commands-suite witness into a normal `FileOps_TerminalFactsProjection` case in the governed File Operations suite. The earlier draft remains in the parent `test-witness` folder for provenance. No repository source was edited and no native build or test was run.

Apply [`test.patch`](test.patch) after the sibling [`../../projection.patch`](../../projection.patch). The case is registered in `SelfTestState::Step`, `StepToString`, `kFileOpsPhaseOrder`, and a one-case `FileOpsFamily_TerminalFactsProjection` in the default family catalog. The existing generic phase selection resolves a focused `FileOps_TerminalFactsProjection` filter through Setup and Cleanup. Setup transitions through this early phase; unselected families skip it through `NextStep`, while a selected run advances to the next selected phase or Cleanup. Its production `FileOperationsPopup` is private to the case and remains hidden under `ScopedWindowActivationBlocker`; the application popup is required to be absent, and the operation state must be idle. The high task ID is checked for a collision. RAII expires the private popup's host token before destroying the window, then dismisses the injected completed summary. No `DebugSetVisualScenario` call is used.

The fixture inserts one real `CompletedTaskSummary` with 9 published and 1 publication unknown; 7 sources removed, 2 retained, 1 native-retained, and 1 source unknown; and 6 verified, 1 failed, 1 unavailable, 1 canceled, and 3 verification problems. It asserts every projected count, the exact result summary, `verificationRequested`, and the legacy selected `Failed` state. The popup selftest snapshot message invokes `BuildSnapshot()` synchronously, so the assertion exercises actual completed-history projection. The issue diagnostic is retained in the fixture but is not an assertion about Issues-action eligibility. This witness makes no timing claim about the live-finished branch.

[`red-baseline.patch`](red-baseline.patch) combines the test with only the six `TaskSnapshot` declarations from the projection draft. Without the projection assignments, all new fields remain zero, yielding a credible failing count witness. With the full projection patch, the counts should match. This expectation still requires the root agent's native lane to compile and run it.

Frozen source SHA-256 on 2026-09-23:

| Source | SHA-256 |
| --- | --- |
| `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp` | `2465E5D2C7477ACEF3EE02F09994E002F9AD2CD0B7DDA2E950A5BD95F5840581` |
| `RedSalamander/FolderWindow.FileOperations.Popup.h` | `5046DCEC26611DF3D628CECA6B89BCB4C0550FA38DF50432AA9FD07E2D050255` |
| `RedSalamander/FolderWindow.FileOperations.Popup.cpp` (projection context; untouched here) | `B36C168D0893198DAFE62CCE9613827F003110F85A5168DBA456DF8EA7B21745` |

External draft SHA-256:

| Artifact | SHA-256 |
| --- | --- |
| `FolderWindow.FileOperations.SelfTest.cpp` | `6D0CDDE3DC2CE2D75139C51CB6784C4DCB5E68554ADD611E401B887C8BC908DE` |
| `test.patch` | `1C4EFEBB36BB17CBF80679E0DE6F6139C6408776A85606EC1FD560A9612C2EEF` |
| `red-baseline.patch` | `B16E348500511345B31180D02DE996D28A43ACC0645C76C6B1DA2C74DCD70FDB` |

Both patches passed `git apply --check` against the hashed primary source files. That is a patch-applicability check, not compilation or execution.

Lifecycle audit: `FileOperationsPopup::Create` accepts a weak host token and borrows `FileOperationState` and `FolderWindow`. The local shared token stays alive through the synchronous snapshot request. Hiding the popup avoids `UpdateLastPopupRect` during creation and painting. `OnNcDestroy` ordinarily calls `OnPopupDestroyed`, which saves application placement even for a private popup; the local scope guard expires the token before window destruction so that callback is skipped. `DismissCompletedTask` can close the application's popup when no tasks remain, hence the precondition that no application popup exists. `Fail()` immediately performs suite cleanup, so the validator returns a literal failure reason and the dispatcher calls `Fail()` only after the popup and injected task have been cleaned up.

`BuildRunFiltersImpl("")` enumerates the new family in default FileOps runs. Focused and family filters both use normal Setup and Cleanup. An empty-filter single-run phase order also includes the new case, and Setup reaches it before any operation phase. The success route calls `NextStep` toward Phase5; selected family or exact-case runs resolve that to Cleanup, while a broader run continues normally. This avoids placing the witness after Phase14, which shuts down the operation state.
