# Final headless Compare host-policy control

Run `c0-compare-host-headless-final-debug-20260905`, Debug x64, on `190847e3`
plus the final host-fixture repair. Snapshot
`3315b326efc7436b78c61137511c6b493e1a4d608ada393eb8b61bb89d084fed`;
application-target receipt
`9d697289837531aa540c60b2e964d3654cb5ab82e6acc0f27ab9c6dbb88b80b7`.
Build 03:07, zero warnings/errors.

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -Configuration Debug -CaseFilter 'host_services_startup_selection,unique' -TestRoot 'C:\RedSalamander.Perf' -RunId c0-compare-host-headless-final-debug-20260905
```

Native/runner: 2 passed, 0 failed, 0 skipped, exit 0. Native 56 ms, runner entry
408 ms, wall 7,225 ms. Filter selection passes 21 checks, and two classification
controls prove a missing required host fails rather than skips while S_OK is not
misclassified. Combined control timing: 1,646 us (`value0=21`, `value1=2`).
The ordinary `unique` comparison also passes, 27 ms.

The startup trace remains headless and contains no InitInstance/FolderWindow
startup. `UntilMessageLoop` is the headless finalization label, 307,851 us; the
extra ordinary comparison prevents treating this as a performance delta against
the [initial one-case control](../2026-09-05_c0_compare_host_headless_initial/README.md).
All 22 perf rows and native/root/runner results, traces, fingerprints and build
log are copied unchanged. No profile or network operation was needed.

The [hosted policy run](../2026-09-05_c0_configured_remote_host_policy/README.md)
uses the same source snapshot and provides the other side of the startup boundary.
Disk audit: 95 retained issues, none manually cleaned. This is focused Debug
evidence, not Release or Full repository qualification.
