# Configured remote access and host-security regression checkpoint

Run `c0-configured-remote-host-policy-debug-20260905`, Debug x64, on `190847e3`
plus the final Compare host-fixture repair. Snapshot
`3315b326efc7436b78c61137511c6b493e1a4d608ada393eb8b61bb89d084fed`;
application-target receipt
`3064a0d778214bc250496d9098d63b4f12d16c501cc25e04e205e22b7dfc28c8`.
Build 02:29, zero warnings/errors. No source/spec/index edits overlapped the run.

**11 passed, 1 failed, 0 skipped.** Native duration 3,160 ms; runner entry about
4.1 s, wall 10,856 ms; process/runner exit 1. The runner classification remains
`UNCLASSIFIED_FAILURE`; source/configuration inspection identifies IMAP's TLS
policy refusal below. The aggregate has not been edited to manufacture a pass.

| Case | Result | Duration ms |
|---|---|---:|
| `host_services_startup_selection` | pass: 21 filter and 2 failure-classification checks | 2 |
| `host_services_startup_context` | pass: real host reaches absent-profile lookup | 0 |
| `windows_hello_cache` | pass | 31 |
| `oauth_refresh_token_storage` | pass | 37 |
| `google_drive_cleared_client_id_requires_configuration` | pass | 328 |
| `google_drive_connection_requires_refresh_token` | pass | 0 |
| `onedrive_personal_cleared_client_id_requires_configuration` | pass | 203 |
| `remote_file_s3` | pass: live read-only comparison | 481 |
| `remote_file_ftp` | pass: live read-only comparison | 196 |
| `remote_file_sftp` | pass: live read-only comparison | 187 |
| `remote_file_scp` | pass: SCP profile's SFTP listing route | 454 |
| `remote_file_imap` | fail: `0x80070005` | 5 |

## Scope and safety

The user approved disabling Hello for test copies of these connections/locations
only, on this machine. Each run creates a fresh isolated settings projection:

| Profile | Exact approved initial path |
|---|---|
| `FileOpsSelfTest FTP` | `/home/RedSalamander-selftest` |
| `FileOpsSelfTest SFTP` | `/share/CACHEDEV1_DATA/homes/eric/RedSalamander-selftest` |
| `FileOpsSelfTest SCP` | `/share/CACHEDEV1_DATA/homes/eric/RedSalamander-selftest` |
| `FileOpsSelfTest IMAP` | `/INBOX/RedSalmander-selftest` (saved spelling) |
| `FileOpsSelfTest S3` | `/redsalamander-selftest` |

The seed uses the marked `C:\RedSalamander.Perf` helpers, exact containment checks,
root `schemaVersion` and `connections.items`, and only reviewed fields
`id,name,pluginId,host,initialPath,userName,savePassword,extra,requireWindowsHello`.
Exactly five unique names and their plugin/path pairs were checked. Only
`requireWindowsHello=false` differs; global bypass stays off. Normal Debug
settings SHA-256 matched before/after seeding and before/after execution.
Stable IDs preserve ordinary WinCred lookup. No secret was extracted by setup,
printed, archived, or copied into settings. Existing security cases use their own
temporary synthetic identities and ordinary cleanup, not these profiles' secrets.

IMAP's original `extra.ignoreSslTrust=true` was preserved; `insecureTlsAck` and
`connections.allowInsecureTlsInAutomation` are absent. The existing
`HostServices::BuildConnectionJsonUtf8` automation gate rejects this before
network dispatch. Hello approval does not authorize a TLS or SSH exception.
The optional question about testing an isolated IMAP copy with certificate
verification **enabled** has not been answered at this checkpoint; no such
configuration change or run is claimed. Normal settings remain untouched.

The four live passes require successful remote root listing, non-missing root,
root-relative mapping and a unique local item reported OnlyInLeft. They do not
prove payload transfer, Copy/Move/Delete/Rename, quota/cancellation behavior, or
SCP payload transport. Remote mutations still need separately authorized unique
children. IMAP remains a missing live witness, not an expected successful skip.

## Repair and measurement

Compare-only startup had exited headlessly before the connection host existed.
Declared host-dependent selections now use the existing owned, non-activating
application/FolderWindow lifecycle; pure/UNC-only selections remain headless.
The production host-window and security guards are unchanged. Missing host errors
in configured secret checks and credential/cloud-gate cases now fail, not skip.
The ANSI metadata owner remains WIL RAII; only its first-instantiation deleted-copy
warnings have a local, documented suppression.

The [final headless run](../2026-09-05_c0_compare_host_headless_final/README.md)
uses identical source. The earlier [all-skipped baseline](../2026-09-05_c0_configured_remote_host_unavailable/README.md)
and [warning-bearing first hosted run](../2026-09-05_c0_configured_remote_host_context_initial/README.md)
remain intact. Full 268-row perf is preserved: selection 1,643 us, non-secret host
lookup 142 us. `UntilMessageLoop` 3,516,835 us includes tests executed inside
WM_CREATE; it is neither pure startup nor a comparable network throughput metric.
No performance improvement is claimed from differing filters or one remote sample.

All native/root/runner results, traces, build log and entry fingerprints are copied
unchanged. Source/resource policy passed 197/197, zero skips, 22.991 s; its NUnit
report is included. An initial unsupported Pester display option failed before
test execution; the supported invocation produced this complete report.
Disk audit: 95 retained issues, none manually cleaned.

## Continue from here

Use fresh run IDs; never overwrite these results or archive seeded settings.
The runner JSON preserves the exact twelve-case command filter. Reproduction
requires the same scoped seeding described in `Testing_SelfTestRemoteCredentials`.
Do not use `-SkipBuild` with these application-target receipts.

Next independent C0 work is the measured 4,097-file Curl Copy throughput failure
and exact modified lookup accounting. Keep the 120-second deadline, complete
listing validation, contents and connection cap unchanged. IMAP strict-TLS testing
can follow the pending answer independently. Current Release, broader Compare,
provider route/memory evidence, C1–C6 and exact Fresh Full still block Done.
