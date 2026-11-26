# Scheduled Rename with missing SMB identity

The test-enabled x64 Release build at `40f9a061027a101b9b92b96d8884eea9d9492c7f` completed with zero warnings and errors. Its verified receipt is `0cfebe5b698c1a3aa8c92fc4c1440f7e69a03ab975cfdf09c9bf3e4211dfa71c`. The runner used the approved Windows SMB test child for both non-activating selections.

- Exact binding run `20261126T073721Z-54392-51ad18b31c43430380eb073b7389af16` passed all three stages. Batch Rename and Change Case returned one exact committed receipt each and preserved content at the requested spelling.
- Controlled missing-identity run `20261126T073816Z-70340-c88f86d2a4da40dc991348158effd2f7` passed setup and cleanup and reported the exact-binding portion as skipped. The fixture disabled file IDs through the test-only prefix switch. Both scheduled origins returned `ERROR_NOT_SUPPORTED` with a failed, not-attempted, retained-source receipt. Each kept the original spelling, bytes and unrelated sibling.

| Scheduled task | Exact commit | Controlled refusal |
| --- | ---: | ---: |
| Batch Rename | 41,252 µs | 53,807 µs |
| Change Case | 29,960 µs | 30,826 µs |

Both runs removed their owned child fixtures. The runner cleared the process-local switch and restored the original Samba test registration byte for byte. RubyGloom's approved share was unreachable during this session. Its live run remains pending.
