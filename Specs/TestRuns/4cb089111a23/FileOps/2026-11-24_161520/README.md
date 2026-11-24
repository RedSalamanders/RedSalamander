# I66 Windows SMB exact Rename

The non-activating `FileOps_SmbExactRenameProvider` selection passed setup, the case
and cleanup on a configured Windows SMB host in run `20261124T151439Z-2272-6644ae5b26034ebd86d075850c8c2390`.
The x64 Debug solution at `92c2b2762` built with zero warnings and zero errors and used
receipt `4da9339e25aaaffabceb537ac976568df6ba4e22c6560a63726ab2cfcd2106e8`.

The case checked a nested free name and a case-only name by reopening the result and comparing
the bound identity, original-path state, stored spelling and text. The independent occupied
file survived. The read-data watcher produced `ERROR_SHARING_VIOLATION` with the source
still present. The timed intervals include the checked provider bind, mutation and verification.

| Measurement | Duration (µs) | Items |
|---|---:|---:|
| One guarded Rename | 4961 | 1 |
| Independent nested names | 1001306 | 100 |

The separate Win32 probe `20261124T150907Z-88200-0ba0dc15eaaf4c1ebaba28b4a7b27ff2`
observed `ERROR_SHARING_VIOLATION` for actual parent and ancestor renames on both
approved SMB endpoints while read-data guards were held. The no-replace call preserved an
independent occupied destination and succeeded to a free sibling on each endpoint.
The other SMB endpoint reported that exact file identity was unavailable. The actual provider
case skipped there in run `20261124T145318Z-86896-cf65eb1fac8447518ff7509d6d3dfe51`.

This record leaves the corroborated route, case-only fallback hops, peak handle counts,
host cancellation, paired performance and Release qualification open.

The original trace contains 127004 bytes and 354 rows. Its SHA-256 is
`9b588462403a1a41073d214c9e3abcf9e57cfc3707083fff29c96b7dc26395b4`.
The three retained rows preserve the original order at lines 249, 252, 354.
