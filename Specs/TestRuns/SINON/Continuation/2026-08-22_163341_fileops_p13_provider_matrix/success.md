# File Operations P1.3 local-profile validation

- Date: 2026-08-22
- Debug x64 build: passed, 0 warnings and 0 errors
- Build attestation: `fde11c80dd04257cdbfe49f4a6d26b03df8cd54a4ce811c962d0e683004e14a4`

Focused validation after both diagnosed admission-fixture failures were repaired:

| Command/case | Result |
|---|---|
| `RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix` | Passed, exit 0 |
| `RedSalamander.exe --fileops-selftest --selftest-case=Fairstream_CopyIntoSelfAliasRejected` | Passed, exit 0 |
| `RedSalamander.exe --commands-selftest --selftest-case=file_operations_phase0_typed_contract_source_guard` | Passed, 1/1 case, exit 0 |
| `PluginContractTests.exe` | Passed, exit 0 |

The provider matrix covers typed envelope rejection, concrete-path capability admission,
no-follow local binding, copied identity/provider comparison agreement, direct and hard-link
same-object outcomes, same-folder Move, destination ancestry/link containment, and final-boundary
source/destination replacement. The Fairstream case proves the alias is stopped just in time before
provider publication rather than by a recursive top-level filesystem preflight.

The two preceding failures are retained in `failure.md`; neither produced a crash artifact or left
a process running.
