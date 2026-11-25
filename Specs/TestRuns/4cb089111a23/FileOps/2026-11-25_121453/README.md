# I66 Managed Move cleanup and provider admission

The x64 Release selftest `FileOps_ProviderCapabilityMatrix` checks managed source cleanup
after a known non-commit. The held original and the published copy remain when the user
skips cleanup. Retry remains a delete-only decision for sharing, access, and disk-full
failures. The path-length scenario verifies Skip and Cancel and checks that Retry is absent.

The Release campaign on main `ccdab5fe5` recorded the old path-length Retry expectation in
[run 45566947991](https://github.com/RedSalamanders/RedSalamander/actions/runs/45566947991).
Local x64 Release `3eca37b1c`, run `20261125T103239Z-58200-3f6a8c87794349198eb04c51fc573ab7`,
passed the cleanup checks and then failed because the matrix treated admission of a Recycle
request as proof of provider Recycle support. The current check keeps the separate capability
assertion and expects a permanent-only provider to require explicit permanent consent.

Integrated `4a0d4c6a6` built x64 Release with zero warnings and errors. Receipt
`a39f1ad07e1f83bcb15d8f6bc99094e880875867407c3c8d14eb3104bf709e7a` and run
`20261125T111418Z-17464-fdd7961fb8f24abc857b1fe6089173ee` passed setup, the matrix, and
cleanup without desktop activation. `FileOps.SelfTest.ManagedCleanupKnownNonCommitMatrix`
recorded all seven expected prompt cases. The measured native File Operations suite duration
was 4.7 seconds. This is a focused qualification. Windows ARM64 and the full native Release
matrix remain open.
