# Windows SMB scheduled Rename evidence

The exact x64 Release source `281a4f40cbbb0ac8fd14608d650f3a90206f82e9` built with zero warnings and errors. Its test-enabled receipt is `81f1772d4e42ae7468b16a43a22f1bf513ca04d1ff5e5794b3f79fcb812c3023`.

Non-activating `FileOps_SmbExactRenameProvider` run `20261126T062554Z-20140-9dff975c31df419da5b3f873e44581f7` passed Setup, the Rename case and cleanup with zero failures and zero skips. It used the registered test child of the approved Windows SMB endpoint. Both shared task origins returned one exact, committed receipt, preserved the bytes and unrelated sibling, and produced the requested spelling. The run removed its test fixture and restored the original Samba registration byte for byte. Repository coverage was not evaluated.

| Observation | Measured time |
| --- | ---: |
| Exact provider call, one item | 5,909 µs |
| Exact provider calls, 100 items | 600,599 µs |
| Shared Batch Rename task from admission through completion | 40,696 µs |
| Shared Change Case task from admission through completion | 30,183 µs |

The provider controls and scheduler measurements include different work. The blocking watcher control recorded `ERROR_SHARING_VIOLATION` before the witness released the watcher. This run covers Windows SMB exact binding. Samba scheduled refusal and other provider routes remain separate gates.
