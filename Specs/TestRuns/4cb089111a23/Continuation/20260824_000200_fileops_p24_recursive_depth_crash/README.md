# P2.4 recursive-depth crash diagnostic

- UTC run: `20260823T220130Z-76532-02884ae36a7844b982e0ed0af70e69f2`
- Command: `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Phase7_CopyItemsSingleFolderRecursiveParallelism`
- Build receipt: `9f4c05bc3e73ef61a02b92cb111a144453ad65ae0450bcb02a63138a1b8ec6aa`
- Fixture: 192 nested directories, four 512-KiB leaf files, exercised first through Local `CopyItems`.
- Result: process exited before a self-test result was written. WER recorded `0xC0000005` at image RVA
  `0x1B9E0AA`; the exception context stack pointer was 176 bytes below the captured committed stack
  range. The remaining captured stack contains the unhandled-exception dump writer, consistent with
  a secondary failure while handling exhausted stack.
- Dump: `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.79296.dmp`
- Dump SHA-256: `EF385C123C1D3563E1FE6BE06D198995E09F296E795E0DCC05FF82AB8BD5D519`
- Dump size: 125,871,473 bytes. It remains outside the repository; this archive retains its exact
  pointer/hash plus the WER and runner evidence.

Conclusion: the recursive provider walker cannot honestly advertise or test a 256-level hard bound
on the default worker stack. The production bound and near-bound test must be lowered, or the walker
must become iterative before P2.4 closes.
