# I66 x64 Release qualification witnesses

The full test-enabled x64 Release build at `fb35e71261e171e1fd91fe9b4c6a8e588d70c4ec` recorded zero warnings and errors. Both native runs verified receipt `a9636babfeeee1496fc758597d8716c4132a210ae551b8a843bcbf6b08a31e92`.

| Selection | Governed run | Passed | Failed | Skipped |
| --- | --- | ---: | ---: | ---: |
| `file_system_provider_name_policy_consumption_source_guard` | `20261125T193101Z-91332-6c2f037df8354649a8e8da7a1764ae70` | 1 | 0 | 0 |
| `R4A02_DeleteOverCopySourceWarnsAndQueues` | `20261125T193210Z-91332-ef983dde0b884212b7dcb4eb4225e52c` | 3 | 0 | 0 |

The F7 guard verified the displayed-listing comparison and worker-side provider name qualification. R4A02 witnessed the exact Copy predecessor while the source was still present, recorded 902,098 µs of explicit Queue wait, and completed the Delete after releasing the held Copy. Both bounded native runs retired their owned processes and released the approved desktop lease. The Queue selection used `--selftest-no-activate`.

Compiler-free tooling run `20261125T191911Z-36792-2fa517fc270949e3a900d56d6abd1dc7` passed both visible-native audit contract tests and 276 ownership, inventory and governance tests. The owned audit had zero unclassified findings. The fixture rejected visible and wrong-path variants of the reviewed hidden parents. Specification inventory, tool inventory and changelog validation passed.

These focused runs reported repository coverage `NOT_EVALUATED`. Hosted native x64 and ARM64 qualification remain open.
