# Scheduled Rename with missing SMB identity

The test-enabled x64 Release build at `40f9a061027a101b9b92b96d8884eea9d9492c7f` completed with zero warnings and errors. Its verified receipt is `0cfebe5b698c1a3aa8c92fc4c1440f7e69a03ab975cfdf09c9bf3e4211dfa71c`. The runner used the approved Windows SMB and Samba test children. All four selections ran without activation.

- Exact binding run `20261126T073721Z-54392-51ad18b31c43430380eb073b7389af16` passed all three stages. Batch Rename and Change Case returned one exact committed receipt each and preserved content at the requested spelling.
- Controlled missing-identity run `20261126T073816Z-70340-c88f86d2a4da40dc991348158effd2f7` passed setup and cleanup and reported the exact-binding portion as skipped. The fixture disabled file IDs through the test-only prefix switch. Both scheduled origins returned `ERROR_NOT_SUPPORTED` with a failed, not-attempted, retained-source receipt. Each kept the original spelling, bytes and unrelated sibling.

- Live Samba run `20261126T074512Z-45556-faa96c4ec0a848e0b48897a60499ef33` reached the same refusal checks using the provider's actual capabilities. Both scheduled origins preserved the source and sibling and returned the expected failed, not-attempted receipts. Setup and cleanup passed. The exact-binding portion skipped.

| Scheduled task | Windows SMB exact commit | Windows SMB controlled refusal | Live Samba refusal |
| --- | ---: | ---: | ---: |
| Batch Rename | 41,252 µs | 53,807 µs | 236,103 µs |
| Change Case | 29,960 µs | 30,826 µs | 27,995 µs |

Corroborated F2 passed all three stages on live Samba in `20261126T074641Z-65788-b01b80ef946940cf81ccb6250f918b13` and on the controlled missing-identity Windows SMB endpoint in `20261126T074836Z-66984-d8a043acc49e4135847147f05708a33e`. Both checked folder Rename, Keep both for an observed occupant, exact case, cancellation and a changed source before dispatch. A single host folder Rename took 163,083 µs on Samba and 36,118 µs in the Windows control. The 100-item provider-only runs took 19,704,380 µs and 1,590,368 µs respectively.

Each run removed its owned child fixture. The Windows controls cleared the process-local switch and restored the original Samba test registration byte for byte.
