# Merged master archive Copy validation

Validated source commit: 9fb487f6f (the normal merge e9a4226e6 plus the test include-order correction), based on remote master 173298f56729d94cd8d9acbf156ce8facffc4806. The source feature commit is 2f9f88dfb. No production or test code changes follow this validation; this directory records the evidence.

The pinned CI formatter (clang-format 22.1.3) passes all five changed native files. Its first include sort exposed an implementation-order dependency in the File Operations runner; include-order-failed-build.log retains the failed build with two C3861 diagnostics. The correction keeps ReviewRemediation after CrossProviderRoutes in a separate, documented include block. The subsequent full x64 Debug solution build passes with zero warnings/errors and receipt 43eebb7ec26da1815a4a347d3b192576d22fc3e415b849939c697474d96f2588.

PluginContractTests.exe passes in its full mode on that build, including the new archive identity assertions, plugin debug selftests and unload/cancellation contracts. plugin-contracts.txt and plugin-contracts-result.json preserve the output, times and exit 0. The final source-contract Pester run passes 204/204 with zero failures/skips; the increase from the original 201 includes the ViewerSpace source guards merged from remote master.

Command: `./Test-Changes.ps1 -Suite FileOps -CaseFilter FileOps_ArchiveNestedCopy -Interactive -SkipBuild -Configuration Debug -TestRoot E:\RedSalamander.Perf`.

Final run: 20261016T144241Z-93196-db87d201c98a4d0fb6a39bef7ff5f192. Setup, FileOps_ArchiveNestedCopyLocal, FileOps_ArchiveNestedCopyUnc and Cleanup all pass: 4 passed, 0 failed, 0 skipped. The disk audit is clean (0 issues). The unchanged perf log contains 24 successful ArchiveNestedCopy rows, split 12 local / 12 loopback UNC, covering ZIP, 7z and four-volume RAR, exact bytes and output leaf, and destination-validation negative checks. This is scoped correctness evidence and does not qualify a NAS, the user's movie, throughput, Release runtime, ARM64 or ASan.

Pre-push accounting: `./Test-Changes.ps1 -Mode PrePush -Explain -Configuration Debug -TestRoot E:\RedSalamander.Perf` declares PR-equivalent entries delegated to GitHub with CI success pending. It exits 1 as an incomplete plan; it executes no test suite. Five interactive Full entries remain pending (Commands, FileOperations, ViewerPETests.Interactive, ViewerSqliteTests.Interactive, RedSalamanderMonitorEtwLatency), and its local plan selects tooling.pester-build-toolchain and selftest.compare-directories as SOURCE_INPUT_CHANGED. The focused FileOps pass does not close those complete-suite obligations. prepush-plan.json retains that decision. No complete affected, PrePush or Full repository pass is claimed.

All original evidence files in the parent archive and the three Local-x64 archives remain unchanged. The archive .gitattributes rule preserves these original bytes in Git. This evidence-only follow-up is committed before pushing the normal master merge.

