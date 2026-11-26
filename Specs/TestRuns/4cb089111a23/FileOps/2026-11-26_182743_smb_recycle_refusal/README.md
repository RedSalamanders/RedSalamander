# Guarded Recycle refusal on live SMB

The test-enabled x64 Release build of `3eed44771d3874e1f00d370d2e58adcbc07a38a7` finished with zero warnings and errors.
Build receipt `7529c57d9b9f7e6e60e9e19cbc29b9ef738701eaa300798cf1b2a68e0420640e` was verified before each run.
Both selected exactly `FileOps_SmbRecycleRefusal` with no desktop activation.

| Approved endpoint | Archived run ID | Passed / failed / skipped | Consent route | Time to consent | Case duration |
| --- | --- | --- | --- | --- | --- |
| Samba | `20261126T172606Z-17456-785cc7c4c09e45cfbe696e6e327d88ff` | 3 / 0 / 0 | Corroborated by name | 156.951 ms | 657 ms |
| Windows SMB | `20261126T172712Z-55380-0109c1ecb58642f58a088699be1c16d5` | 3 / 0 / 0 | Exact identity | 96.368 ms | 188 ms |

An independent no-follow binding probe determined the expected route before the host task started.
The guarded Shell attempt reported `ERROR_NOT_SUPPORTED` and preserved the original.
The task offered one deferred permanent-delete question for the exact selection.
Skip accepted that prompt generation. The task returned `S_FALSE` with one typed skipped item and retained source.
The case verified the original bytes, removed its own file and empty parent nonrecursively, then checked the fixture was absent.
An independent read confirmed absence on each share. The runner preserved the authoritative resource file on Samba.
For Windows SMB the scoped resource selection restored the original manifest byte for byte.

- Samba native result SHA-256 `b2502176b2bfd73501a5e3fb22179fa5962a707990e6ccfb9f037328f4ecb40d`
- Windows SMB native result SHA-256 `29fa51f450e769744c26d31364768b82ef59eddbe96de8de56cd2479cb06d47c`
- Both reported `FileOps.SelfTest.SmbRecycleRefusal` with one consent and one preserved source.
- The filtered runner exited 2 for each passing selection. The repository verdict was `NOT_EVALUATED`.

## Preserved attempts and runner contract

- Initial Samba run `20261126T170705Z-16996-6ca52b2661e84334b5a91de6ced73d17` on `58573def1` reported two passes and one failure. It observed the by-name prompt with the intact file. The witness expected the exact-identity prompt.
- Second Samba run `20261126T171651Z-29756-ba6521489b964a87b82207f4613d202d` on `22cd4d9bc` took the correct by-name route and recorded `S_FALSE` after Skip. The witness expected one generic conflict. The task trace recorded one deferred consent and zero generic conflicts.
- Both failed fixtures matched the generated file's SHA-256. Each original failure was preserved before its one file and empty directory were removed.
- All 130 runner-plan tests passed on the final source and receipt in `20261126T173004Z-70012-4758d6da47b04979a3303e7282026d60`. The original focused Pester run `20261126T172834Z-60024-61030bd2d4ab49a5ab72d015c9cc52e8` passed the same checks and archived them before its local launcher failed during lock release. The rerun used the repository's lock release function and exited cleanly.
- PrePush Explain selected 14 local non-interactive entries and four pending desktop entries. Hosted native Release and the remaining I66 closeout checks are still open.
