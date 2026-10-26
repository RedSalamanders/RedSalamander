# E4-OR2 popup-control prompt lifetime

- Machine profile: `4cb089111a23`.
- Source commit: `3e45b624f95a9c2a3c965ae4ffa8548bd379be98`.
- Build: test-enabled x64 Release, `MaxCpuCount=1`, RedSalamander project
  receipt `ad7ab5c05566b9b03739345bd52387e7ebca2fdbf81f6fa8f0c823a784499c05`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: `RedSalamander.exe --commands-selftest --selftest-fail-fast
  --selftest-case=cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface`.
- Scenario: five independent processes open the real custom speed-limit modal,
  dispatch nested File Operations shutdown while the prompt pumps messages, and
  verify that `FileOperationState` remains alive until the modal frame unwinds.
- `FileOps.SelfTest.SpeedLimitPromptNestedShutdownUs`: `114892`, `95670`,
  `90862`, `100297`, and `117875` us; nearest-rank p95 `117875` us.
- Regression gate: PASS (`117875 <= 2000000`). Every process reports
  `value0=1` (state retained), `value1=1` (prompt HWND gone), and `hr=0`.

All five processes pass (`passed=1, failed=0, skipped=0`). The post-prompt speed
submission re-finds by task ID, observes that shutdown removed the old task, and
does not write through a stale `Task*`. A fresh File Operations state is then
created successfully.

The archived metric files are bound as `(rows, bytes, SHA-256)`:

- sample01: `(1038, 396781, ACC300A48F293189C8A4CD43D07E7C897F63A7433124B44882FB08FFA2C10958)`
- sample02: `(952, 365692, 3EC9793EB142215C5CD5F0CC6BBEA73DFD16837310D9436703A6A399C4959CFA)`
- sample03: `(1048, 399306, 7EA953DBEF6CFA2A87247B204106C0CB99C271E700A93ED16FBD5B4AD498F480)`
- sample04: `(1083, 414358, E74730BAFD9B998534CF911CE4369D44E008CCB68406738741AB220A9C5AAC18)`
- sample05: `(1026, 393087, 8A92A0DBE4EE0F62B39496D45F1617BE26C6F30141FBE1537D452FCB960CBBF2)`

The exact-source x64 Debug full-solution receipt is
`e10e174c74f8413e6442d5e055d7b3200dad6c3f2c0d1947a4d348d0ff43115c`.
The nested-shutdown behavior and source contract each passed `1/0/0`; the
affected speed-limit prompt family passed `4/0/0`. This evidence is teardown
correctness/latency coverage, not a rendering-performance claim and not proof
that an admitted blocking `IFileReader::Read` is cancellable.
