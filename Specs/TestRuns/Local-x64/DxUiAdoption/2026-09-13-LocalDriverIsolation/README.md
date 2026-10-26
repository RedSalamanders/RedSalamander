# Local validation driver isolation

The first combined local driver builds ARM64 in `Z:\RsI19`, then invokes the x64
driver in `Z:\RsI19Closeout` in the same PowerShell process. Both checkouts are frozen
at `a20f6680`. Their same-named script modules remain loaded together. Eight general
tooling checks in Release Full run `20260913T034047Z-88456-99c436c59afe45f698e35f3b19b2dd59`
fail before their intended assertions because Pester 3.4.0 resolves module scopes
by name; one direct module invocation likewise receives both module objects.

The isolated module-lookup probe demonstrates the mechanism:

| Loaded checkout copies | Module lookup | Result |
|---|---|---|
| Both | Fails with the same duplicate `DxUiDependency` module error | Expected failure reproduced |
| Closeout only | Succeeds | Clean-process lookup verified |

The probe performs no build, product test or artifact mutation. It establishes the
driver's scope problem, not a passing product regression or a waiver of the failed Full.
The [completed Full](../2026-09-13-Releasea20f6680/README.md) passes 19/20 entries;
only the general tooling entry fails. All product suites pass. Its driver stops
after preserving the failed result, before starting ASan Debug or Debug.

The replacement starts a separate `-NoProfile` PowerShell process for each
configuration, rejects preloaded repository modules and keeps the checkout frozen.
It does not repeat the completed package rehearsal. Outputs use a separate `isolated`
prefix, preserving the first attempt. No product source, assertion, timeout or Pester
implementation changes. The [isolated Release Full](../2026-09-13-Releasea20f6680Isolated/README.md)
subsequently passes all 20 entries, including all 770 general tooling checks. ASan Debug
and Debug qualification remain separate gates in the opening checklist.

The script snapshots are inert `.ps1.txt` evidence. `retained-files.json` binds all eight
originals to their bytes and source locations. Current status belongs to the
[I19 checklist](../../../../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).
