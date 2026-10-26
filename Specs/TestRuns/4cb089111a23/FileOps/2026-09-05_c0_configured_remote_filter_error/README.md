# Configured remote preflight — invalid FileOps filter

Run `c0-configured-remote-preflight-debug-20260905`, Debug x64, based on
`99d2a088e05c0fd9d40b97ffed8e708aacc7d3c5`. This is retained failed invocation
evidence, not a server, credential, performance, or product-behavior result.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter 'Phase16_RemoteFtpSecret,Phase16_RemoteFtpSandbox,Phase16_RemoteSftpSecret,Phase16_RemoteSftpSandbox,Phase16_RemoteScpSecret,Phase16_RemoteScpSandbox,Phase16_RemoteImapSecret,Phase16_RemoteImapSandbox,Phase16_RemoteS3Secret,Phase16_RemoteS3Sandbox' -RunId c0-configured-remote-preflight-debug-20260905`.

The invocation mistakenly applied the shared Compare/Commands comma-filter syntax
to FileOps. `ResolveRunSelection` supports one exact case, family, or prefix, not
a comma-separated set. It rejected the filter before any selected phase ran.
Native result: zero cases, one scheduling failure. Governed result: zero passed,
two failed, zero skipped; exit 1, including result-coverage failure. Do not count
this as ten failed prerequisites or ten executed checks.

Only the five named non-secret profile objects had been seeded into the fresh
run's isolated Settings Store; stable IDs and credential policy were preserved.
Normal user settings were not changed; no secrets or seeded settings are archived.
The corrected five-case Compare attempt is recorded
[separately](../../CompareDirectories/2026-09-05_c0_configured_remote_readonly_blocked/README.md).

Build: 03:19, zero warnings/errors. Receipt
`af386c64b7c089c43d19954b2dce6b8704eab512216a2dc41d77c1e83f02455c`;
workspace `b2a695ab4a68ebe67c8de11e485247927ed1df65c1d9320321c25443c4a58bba`.
Full results, traces, metrics and entry fingerprints are preserved unchanged.
Metrics describe startup only; there is no remote timing baseline. Disk audit
reported 84 retained-run issues; no runs or user files were removed.
