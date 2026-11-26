# I66 Copy and Move selection shapes on Windows SMB and Local

All runs used the same test-enabled x64 Release build at `41211242ce6016c1fba02c339a96055ebe0f05b4` and the byte-verified receipt `4c2358e6f1001670af6e135b3340a032211d53832bafe63199414e8729132ce8`. Each run passed setup, the exact case and cleanup without activation.

| Endpoint | Selected case | Run | Passed / failed / skipped |
| --- | --- | --- | --- |
| Windows SMB, SINON | `FileOps_SelectionShapeSmb` | `20261126T102921Z-15456-2f0b0dbb9cf74123b94e9cb147879950` | 3 / 0 / 0 |
| Samba, RubyGloom | `FileOps_SelectionShapeSmb` | `20261126T083830Z-52184-e06300c5781d490285f6906a04ee0b14` | 3 / 0 / 0 |
| Local | `FileOps_SelectionShapeLocal` | `20261126T103221Z-9636-e3040fa44cf94091b0ce9fb7e7be6239` | 3 / 0 / 0 |

Each folder and selected-file pair used independent source and destination fixtures with eight 4 MiB files and one empty file. All 18 scenarios verified nine output files, 33,554,432 payload bytes, content equality, a fixed final total and the required source retention or removal. The selected-file cases reached determinate progress before completion.

| Endpoint and route | Folder elapsed | Selected files elapsed | Peak active files for each |
| --- | ---: | ---: | ---: |
| Windows SMB native Copy | 4,599,194 µs | 5,100,239 µs | 4 |
| Windows SMB host bridge Copy | 4,426,060 µs | 5,146,114 µs | 4 |
| Windows SMB managed Move | 4,492,491 µs | 5,138,742 µs | 4 |
| Samba native Copy | 5,252,129 µs | 6,821,465 µs | 4 |
| Samba host bridge Copy | 4,958,722 µs | 6,294,953 µs | 4 |
| Samba managed Move | 4,981,489 µs | 6,187,158 µs | 4 |
| Local Copy, Manual four | 4,488,726 µs | 4,419,262 µs | 4 |
| Local Copy, Auto with inactive Manual one | 4,604,860 µs | 4,453,022 µs | 4 |
| Local Copy, Manual one | 4,107,599 µs | 4,922,400 µs | 1 |

The Local case controlled the Auto probe to resolve four. The two SMB cases exercised real disk-to-share I/O on their respective resources. These are single-run elapsed observations. The runner removed its owned Windows SMB child and confirmed the fixture was empty. The temporary selection restored the resource manifest byte for byte with SHA-256 `1779448AD25C15B19C41F1CD0A1C81D68C7711CFE834DF2B9EE7599A3BBD7A22`. [Samba run record](../2026-11-26_093924_samba_selection_shapes/README.md)

## Integrated workflow checks

These runs used the merged runtime source at `78b8c46d35a0c8d3a11cd944016a4be0269eec19` and byte-verified x64 Release receipt `90b1cf616a1dc1d7ad25f3ba1ae294c31c29ec394da8d37ed5f5753fbc8b8963`. Each passed three stages with zero failures and skips, without activation.

| Case | Run |
| --- | --- |
| `FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof` | `20261126T103622Z-70908-6c4da4b2df1d4daabf9342451b0b65b1` |
| `FileOps_ReviewRemediationContracts` | `20261126T103711Z-74284-0866d2274311474d880dd92b847a9f54` |
| `FileOps_RepeatDecisionContracts` | `20261126T103941Z-56252-0393e086a2bb419aa415d84406b7d2bf` |

These filtered runs reported repository coverage `NOT_EVALUATED`. Broader desktop and native Release qualification remain open.
