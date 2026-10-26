# R2 provider-matrix Debug stack-overflow continuation

- Exact command: `Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_ProviderCapabilityMatrix -TestRoot D:\RedSalamander.Perf -AllowAlternateVolumeTestRoot -SkipBuild`
- Build attestation: `f6b3fc62ceef694f12e4ba11d70065873db6396ce9049667ac924cb2e5a72c08`
- Run id: `20260901T014201Z-72340-43ab30ca2cc345dc86878b2c4e9785f0`
- Process/thread: `21760` / `95708`
- Failure: `0xC00000FD` (`Stack Overflow`) while the provider matrix executed its Managed Move fixture.
- Deepest product frame: `CopyBackupMetadataStreams` at `Plugins/FileSystem/FileSystem.cpp:2229`, entered from the host bridge metadata-transfer path.
- Original crash artifacts: `%LOCALAPPDATA%\RedSalamander\Crashes\RedSalamander-20260901-034222-p21760.{dmp,txt}`.
- Checked-in dump: `RedSalamander-20260901-034222-p21760.dmp.zip`; the uncompressed original remains under `D:\RedSalamander.Perf\evidence\r2-provider-matrix-stack-overflow\` on the authorized test volume.
- The runner also reported 107 disk-audit findings for older explicitly authorized sibling run directories under the initialized `D:\RedSalamander.Perf` root. Those findings are retained in `run-all-tests-results.json`; they did not cause the process stack overflow.

Evidence was copied before any focused rerun, per the crash-forensics workflow.

## Remediation and focused GREEN

- Root cause: `CopyBackupMetadataStreams` placed a 64 KiB I/O buffer on an already-deep Debug/selftest stack; this was a stack-budget cliff, not recursive control flow.
- Remediation: the buffer is now a fixed-size `std::unique_ptr<std::byte[]>` allocated with `std::nothrow`; allocation failure returns `E_OUTOFMEMORY`, and the metadata-transfer loop retains the same 64 KiB chunk size.
- Full Debug build attestation: `623677f34ece12cbc77070a9857bcc8e9a22b7c304eb7d3eed8cdcc4926599cb` (zero warnings/errors).
- Focused GREEN run id: `20260901T015002Z-21808-51a132f9be1548319ced72418166d32b`.
- Result: `FileOps_ProviderCapabilityMatrix` passed all 3 cases, including the Managed Move fixture that previously overflowed the stack.
