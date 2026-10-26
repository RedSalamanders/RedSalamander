# First hosted configured-provider smoke

Run `c0-configured-remote-host-context-debug-20260905`, Debug x64, on `190847e3`
plus the Compare host-fixture repair. Snapshot
`c5c9c1b344894b6944a1a074fdfcf31158109d10c9c50135236c90d23db55a67`;
application-target receipt
`1b714d9499c15d9d8f41fc35a7e7f1ab69894fcbbf4598f5065a6fc05ed6b3a1`.

Native/runner: 6 passed, 1 failed, 0 skipped; process/runner exit 1. Native
1,182 ms, runner wall 7,334 ms. The exact seven-case filter is preserved in the
runner results. Both startup controls passed; four live root comparisons passed:
S3 447 ms, FTP 303 ms, SFTP 189 ms, SCP-profile SFTP listing 183 ms. IMAP failed
in 4 ms with `0x80070005` (`ERROR_ACCESS_DENIED`). The saved IMAP profile has
`extra.ignoreSslTrust=true`, no `insecureTlsAck`, and no automation TLS allowance;
`HostServices::BuildConnectionJsonUtf8` rejects that combination before network
dispatch. The narrow Hello approval does not authorize weakening TLS policy.

Only the [five previously approved profile/roots](../2026-09-05_c0_configured_remote_host_unavailable/README.md)
were projected into this run's isolated settings. Only per-profile Hello was
disabled; stable IDs, other reviewed fields and normal settings were preserved.
The normal settings SHA-256 matched before/after seeding and the runner. An
initial redundant containment call used an unsupported parameter; the seed had
already been placed under the marked-root helper's exact owned scratch directory.
Before launch, a fail-closed `-TestRoot` containment check and every projected
property were revalidated. No outside-root write or normal-settings overwrite
occurred. Seeds and secret values are not archived.

Build 02:57, two WIL C4625/C4626 diagnostics, zero errors. The build summary
counts ten warning lines because template-context lines repeat their codes; this
is not a zero-warning build. The [final policy run](../2026-09-05_c0_configured_remote_host_policy/README.md)
retains the correction and broader regressions.

All 257 perf rows are preserved: filter policy 1,767 us, missing-profile host
lookup 160 us, enclosing `UntilMessageLoop` 1,660,717 us. Hosted Compare runs
inside WM_CREATE, so enclosing startup timings include the tests. They are not
pure startup or transport-throughput comparisons. Source/resource checks passed
197/197, zero skips, 25.268 s; their NUnit report is preserved.

The full native/root/runner results, traces and fingerprints are copied unchanged.
The build log has trailing spaces removed from two blank template-context lines
for patch hygiene; `build-provenance.json` binds the raw and normalized SHA-256,
byte counts and transformation. The raw log remains in `.build/logs`.
Disk audit: 92 retained issues, none manually cleaned. The live passes
prove listing/root mapping/OnlyInLeft, not payload transfer or remote mutation.
IMAP remains a missing live witness; C0, Release and Fresh Full remain open.
