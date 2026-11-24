# I66 Local exact Rename controls

The test-enabled x64 Debug solution built at the zero-warning, zero-error candidate
`842ac811a` with build receipt `9976516429f8ebfeab7b70b9d602d8fd92ed26e14cdbc8d9fccd369c72424b59`.
The Local plugin's object-binding controls passed in PluginContractTests during
run `20261124T135240Z-82196-ca3b3d5d42bb4580a9735bb8aaad3b2e`.

| Control | Call (µs) | Committed | Outcome known | HRESULT |
|---|---:|---|---|---|
| case-only refusal | 309 | No | Yes | 0x80070005 |
| legacy collision | 281 | No | Yes | 0x800700B7 |
| legacy free name | 280 | Yes | Yes | 0x00000000 |
| case-only recovery | 562 | Yes | Yes | 0x00000000 |

The denied case change preserved the original spelling and content. The legacy collision
preserved both the source and the independent destination. The legacy free-name call published
the original object. The injected lost result after the case change was reconciled against the
stored spelling, bound identity and content.

The marked sandbox retains the raw trace. The six archived rows come from original lines 749, 751, 752, 754, 755, 757.
The completed archive contains 66513156 bytes and 167239 rows.
Its SHA-256 was `7e3974d04c4f4daec9cd1ca5f088b21e2be424888e412dd540f148d81ad30940`.
The four observations are single-item Debug controls. A paired baseline, Release, multi-item
cancellation and live SMB Rename qualification remain open.
