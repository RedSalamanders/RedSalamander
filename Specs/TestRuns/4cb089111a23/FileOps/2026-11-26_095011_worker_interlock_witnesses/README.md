# File Operations worker regression witnesses

The test-enabled x64 Release build at `0ff0b5c4bea0c01b40cb6cae930a71b225348fa4` completed with zero warnings and errors. All selections used verified receipt `4bff28ef45cbbec0ad4826931efde2e7cfacbf6a07003db7d65910040db2a097` and ran without activation or a desktop lease.

| Case | Run | Passed / failed / skipped |
| --- | --- | ---: |
| Batch Rename worker admission and cancellation | `20261126T084742Z-69404-74f563ea140a48a7ac1c3b93fad1dbcb` | 1 / 0 / 0 |
| Identity-less Delete with an archive writer | `20261126T084909Z-69404-5246857a28384c0eb2fe465e80d72e28` | 3 / 0 / 0 |
| Files menu information architecture | `20261126T085033Z-69404-ee21d2f1ce464871aae163b5ccc53082` | 1 / 0 / 0 |
| Provider capability matrix, isolated timing pass | `20261126T085554Z-59684-4b5dfc4f169f4b38b624ce3bf9eed32e` | 3 / 0 / 0 |

The Batch Rename witness rejected duplicate exact bindings at plan admission with `E_INVALIDARG`, preserved both hard-link aliases and verified modeless worker cancellation. The provider matrix also accepted the same complete immutable plan with distinct objects. The queued Delete stayed pending while Pack published a new archive in the pinned subtree through its alternate spelling. `FileOps.DeleteTraversal.MaxDepth` recorded `ERROR_REVISION_MISMATCH`. After the refusal the harness closed the remaining question, Delete reported cancellation and the archive remained present at its captured size. The integrated Files-menu pass confirmed the previously merged eligibility rules.

The first provider matrix timing run `20261126T085225Z-73836-c78a6d8c78b846fca1066a595ff67e92` completed all 4,096 admission decisions without provider calls but exceeded the 2,000,000 µs budget at 17,149,033 µs while separate builds were running. The same receipt passed the isolated rerun after the builds retired. The failed timing run remains archived.

All 244 source contracts and 67 selected runner-plan checks passed in `20261126T083358Z-664-f2cb7b98cb9e45cfa6d325e17ec79b0d`. Native Full and hosted qualification remain open.
