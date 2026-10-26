# Configured FTP / SFTP / SCP / IMAP / S3 — Windows Hello blocked

Run `c0-configured-remote-readonly-debug-20260905`, Debug x64, based on
`99d2a088e05c0fd9d40b97ffed8e708aacc7d3c5` plus the three new Compare case
registrations and remote-test/spec continuation updates. No production code changed.

**No live-server test passed or ran.** All five configured profiles were recognized,
but `CheckRemoteConnectionSecret` skipped each because `requireWindowsHello=true`
and no global bypass was enabled. This check precedes secret retrieval, sandbox
validation and remote instance creation. The native trace says PASS because it
has no failed cases; the governed runner correctly rejects `ALL_TESTS_SKIPPED`.
Use the results JSON and runner verdict, not that trace banner.

| Case | Native result | Reason |
|---|---|---|
| `remote_file_s3` | skipped | Windows Hello required |
| `remote_file_ftp` | skipped | Windows Hello required |
| `remote_file_sftp` | skipped | Windows Hello required |
| `remote_file_scp` | skipped | Windows Hello required |
| `remote_file_imap` | skipped | Windows Hello required |

Native: 0 passed / 0 failed / 5 skipped, 42 ms. Governed: 0 passed / 1 failed /
5 skipped, exit 1 (`ALL_TESTS_SKIPPED`), entry 430 ms. This proves case discovery,
profile visibility and the authentication-policy guard, not connectivity, saved
secret availability, CRUD, cancellation, transport memory or quota behavior.
The new SFTP/SCP/IMAP cases reuse the existing read-only comparison scenario.
SCP enumeration uses SFTP; it does not prove an SCP payload transfer.

## Reproduction and next action

1. Read `Specs/Testing/Testing_SelfTestRemoteCredentials.md` and use a fresh run ID.
   The optional machine-resources manifest was absent on this machine. Both normal
   Debug and Release settings contained the exact five `FileOpsSelfTest <protocol>`
   names. Preserve this distinction: absence of the manifest is not absence of profiles.
2. Import `Tools/Modules/Testing/TestSandbox.psm1`; use `New-RSTestRunContext` with
   test root `C:\RedSalamander.Perf`. Refuse an existing run root. Use
   `New-RSTestSandboxScratchDirectory -Harness settings-store -Case RedSalamander`
   and the same run ID. Validate the destination through
   `Assert-RSTestSandboxContainedPath -RequireDescendant`.
3. Mechanically project the selected objects from the normal Debug settings into
   `<runRoot>\scratch\settings-store\RedSalamander\Settings\RedSalamander-debug.settings.json`,
   with only `schemaVersion` and `connections.items` at the top level. This run's
   allowlist was `id,name,pluginId,host,initialPath,userName,savePassword,extra`;
   the only present extra key was IMAP `ignoreSslTrust`, preserved unchanged.
   Reject unreviewed fields, missing/duplicate selected names and unsafe roots.
   Keep profile IDs; secrets stay in WinCred. Do not copy all user settings, write
   normal settings, archive the seed, or print credentials.
4. All five `initialPath` values were absolute dedicated self-test paths, not `/`
   or a home root, and had no dot segments. S3 used a dedicated self-test bucket.
   The profile JSON omitted `requireWindowsHello`, so the parser's default `true`
   remained in force. **Get explicit user direction before changing this policy,
   even for the isolated copies.** No bypass was applied in this run.
5. Once approved/configured, run:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Compare -Configuration Debug `
     -CaseFilter 'remote_file_ftp,remote_file_sftp,remote_file_scp,remote_file_imap,remote_file_s3' `
     -RunId <fresh-seeded-run-id>
   ```

   Require exactly five live passes with zero skips; investigate each failure.
   Do not broaden the filter into remote writes. Separately approved mutations
   must use a unique owned child beneath the configured dedicated root. Existing
   fake-server faults remain necessary and are not replaced by these smoke tests.

The [earlier FileOps filter error](../../FileOps/2026-09-05_c0_configured_remote_filter_error/README.md)
is preserved. Compare supports comma-separated exact names; FileOps does not.

## Provenance and limits

Build 03:02, zero errors, **three MSB3061 cleanup warnings** for shared plugin
`libcrypto-3-x64.dll`, `zd.dll`, and `brotlicommon.dll`. Full build log retained;
these reproduce the unresolved C6 shared-output Clean ownership issue.
Receipt `013379c33de9c2fd8a0e0bbbe1e8b21c0a716a66fdfd849926597931878650fd`;
workspace `6cf4ade7bfb9c1276e06b473dd4771f31b03daad0db961b02a7fd6d063e76aa3`.
`entry-fingerprints.json` binds receipt, entry and source snapshot. The complete
native/runner results, traces and nine startup-only perf rows are unchanged.
No remote performance measurement or baseline exists. Disk audit reported 85
retained-run issues; none were deleted. C0 and the core plan remain open.

Post-run source/resource checks: 197/197 passed, zero skipped, 15.871 s;
`source-resource-contracts.xml` preserves the NUnit result from
`c0-configured-remote-policy-20260905`. Spec inventory: zero blocking findings.
These static checks do not turn the five skipped live cases into passes.
