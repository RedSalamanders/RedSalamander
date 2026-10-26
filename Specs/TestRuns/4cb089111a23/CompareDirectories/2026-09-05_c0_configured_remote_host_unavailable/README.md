# Five configured providers: approved Hello scope, missing test host

Run `c0-configured-remote-readonly-authorized-debug-20260905`, Debug x64,
`e68a8d0b786fcb4f4ee63c7451fdcc4739f07d3b` plus Curl measurement/test and
plan/performance-spec changes. Source snapshot:
`f915a122a1f7811e1f1805a8f3a88b60d7d07f5f1b6a87c72695712839d3059d`.

**No live server was contacted.** The user explicitly approved disabling Windows
Hello only in isolated test copies of the five configured connections/locations
on this machine. That guard now passes, but every selected case stops at
`IHostConnections::GetConnectionSecret` with `0x80070578`
(`ERROR_INVALID_WINDOW_HANDLE`). Native: 0 passed / 0 failed / 5 skipped, 37 ms.
Governed: 0 passed / 1 failed / 5 skipped, `ALL_TESTS_SKIPPED`, entry 364 ms.
The wrapper shell returned 0 after its `finally` hash check; it is not the runner
verdict. The trace's PASS banner likewise does not mean live coverage passed.

| Case | Result / reason |
|---|---|
| `remote_file_s3` | skipped: host connection UI unavailable |
| `remote_file_ftp` | skipped: host connection UI unavailable |
| `remote_file_sftp` | skipped: host connection UI unavailable |
| `remote_file_scp` | skipped: host connection UI unavailable |
| `remote_file_imap` | skipped: host connection UI unavailable |

## Exact permission and settings scope

Only `requireWindowsHello=false` was added to these five projected profile copies:

| Profile | Exact approved initial path |
|---|---|
| `FileOpsSelfTest FTP` | `/home/RedSalamander-selftest` |
| `FileOpsSelfTest SFTP` | `/share/CACHEDEV1_DATA/homes/eric/RedSalamander-selftest` |
| `FileOpsSelfTest SCP` | `/share/CACHEDEV1_DATA/homes/eric/RedSalamander-selftest` |
| `FileOpsSelfTest IMAP` | `/INBOX/RedSalmander-selftest` (saved spelling) |
| `FileOpsSelfTest S3` | `/redsalamander-selftest` |

The seed uses the existing marked-root helpers and only root `schemaVersion` and
`connections.items`. Selected fields:
`id,name,pluginId,host,initialPath,userName,savePassword,extra,requireWindowsHello`.
Exactly five unique names and their plugin/path pairs were verified. Stable IDs
remain unchanged. Only existing IMAP `ignoreSslTrust` was present in `extra` and
was preserved, not relaxed. No passwords were extracted or logged, no global
Hello bypass was set, and no normal profile was written. Original Debug settings
SHA-256 matched before/after seeding and before/after the governed invocation.
The seed is deliberately not archived. Future runs need a fresh isolated copy
with the same exact reviewed scope; this approval does not authorize other roots,
connections, global changes, TLS/SSH policy changes, or remote mutation.

Two preceding attempts never launched a test: the first was refused by the active
C0 Debug artifact lock; the second by `-SkipBuild`'s full-solution receipt rule.
The then-current receipt targeted `redsalamander`, not `solution`. Neither refusal
was bypassed. They created no completed test result; the same reserved seeded ID
was used for the first actual launch with a fresh governed build.

## Cause and next repair

`RedSalamander.cpp` selects headless startup for every Compare-only invocation,
before creating the FolderWindow. `HostServices.cpp::GetConnectionSecret` requires
an initialized host window even when the profile does not require Hello. This is
a host/test lifecycle gap, not evidence of bad credentials or unavailable servers.
The repair must initialize the existing owned host for host-dependent cases and
retain its UI-thread/message-pump/lifetime contract. Do not extract WinCred secrets
outside the host, invent a security bypass, or remove the host-window guard.
Preserve headless pure engine coverage and the runner's desktop/activation policy.

Then rerun exactly these five read-only cases, requiring five passes and no skips:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -Configuration Debug -CaseFilter 'remote_file_ftp,remote_file_sftp,remote_file_scp,remote_file_imap,remote_file_s3' -RunId <fresh-seeded-run-id>
```

SCP directory enumeration uses SFTP; even a pass will not prove SCP payload
transfer. These comparisons never authorize remote writes/deletes. The earlier
[Hello-blocked run](../2026-09-05_c0_configured_remote_readonly_blocked/README.md)
remains unchanged as historical evidence.

## Build and evidence

Build 02:45, zero warnings/errors; application-target receipt
`a519c7a8b8ae63a1257adbbcccbd28d105002d3f9d4483ca0330d547cb6424f8`.
Complete native/root/runner results, traces, entry fingerprints, build log and
startup perf are copied unchanged. No remote timing exists. Disk audit reported
88 retained issues; none were manually removed. This does not clear C0 or C6.

Post-run source/resource policy: 197/197 passed, zero skips, 12.750 s; the unchanged
NUnit XML is retained in the concurrent
[C0 measurement archive](../../FileOps/2026-09-05_c0_curl_copy_phase_accounting/README.md).
Spec inventory has zero blocking findings. The next host-fixture repair has not
been implemented or tested at this checkpoint; no background run remains active.
