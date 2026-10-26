# Initial headless host-selection control

Run `c0-compare-host-headless-control-debug-20260905`, Debug x64, on `190847e3`
plus the initial Compare host-fixture repair. Snapshot
`d7cb034027153c7799a7b61a1e70bbe131bc10b550ad4ec9e5d4007f0f18eaae`;
application-target receipt
`937f5c1c3c21135a6155889a02f572bd6cc3e70db49987fab55f2ac8905c74bc`.

Exact `host_services_startup_selection`: native and runner 1 passed, 0 failed,
0 skipped. Native 24 ms, case 2 ms, runner entry 370 ms, wall 5,869 ms.
The 21 filter checks took 1,630 us. This version predates the two required-host
failure-classification controls (`value1=0`). The startup trace proves the
headless route: no InitInstance/FolderWindow startup. Its final startup timing
label `UntilMessageLoop` is 281,794 us, not evidence that a message loop ran.

Build 03:05, two WIL C4625/C4626 diagnostics, zero errors. This warning-bearing
candidate is retained, not called a clean qualification. The final repair keeps
WIL ownership and locally documents its intentionally noncopyable ANSI owner.
See the [final headless control](../2026-09-05_c0_compare_host_headless_final/README.md).

Native/root/runner results, suite/startup traces, fingerprints and full ten-row perf
stream are copied unchanged. The build log has trailing spaces removed from two
blank template-context lines for patch hygiene; `build-provenance.json` retains
raw/normalized SHA-256, byte counts and the transformation. The raw log is retained
in `.build/logs`. No remote profile or secret was used.
The disk audit reported 91 retained issues; none were manually cleaned. This
focused control does not qualify remote behavior, Release, or the repository.
