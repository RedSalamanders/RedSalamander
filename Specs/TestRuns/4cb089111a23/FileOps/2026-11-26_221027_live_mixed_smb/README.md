# I66 live connection rechecks and mixed Move

## Mixed Local and Windows SMB Move

The test-enabled x64 Release solution for `96e495135b127a79f48c8ecd6e2157930991c898`
built with zero warnings and errors. The verified receipt was
`5aaaae600267450f00e8a3e8c0fcabcd32331ba59c72a6de7d86dad247a47e29`.
The first project-only build passed compilation. The live launcher refused its receipt
before accessing a provider, so execution used the full-solution build.

| Fixture | Exact run | Passed / failed / skipped | Case time |
| --- | --- | --- | --- |
| Windows SMB exact identity | `20261126T211027Z-69880-f17f83181e8f4b34a20f4b892cf1c0cc` | 3 / 0 / 0 | 250 ms |
| Samba identity-less control | `20261126T211205Z-27928-1cd0f2925a58450db8c499d42c4497a3` | 2 / 0 / 1 | 281 ms |

The Windows SMB check used two separately generated payloads and one task with a Local
destination. The Local item completed as Native and the SMB item completed as Managed.
The runner observed one Managed stage, identical destination payloads, and both original
paths absent. It removed the fresh empty SMB child nonrecursively and verified its absence.
The task-through-cleanup metric `FileOps.SelfTest.SmbMixedMove` recorded 91,915 µs, two
items, and one stage. The Samba control skipped before creating a fixture because the root
could not supply an exact identity. Both selections used the no-activation guard, and the
Windows SMB selection restored the original machine resource manifest byte for byte.
The runner also reported 120 sandbox disk audit issues and retained nine older run roots.

Native result SHA-256 values are `95ed2c374874b9d3c4fb650839cd06e6523cfd1bd919f9671fc9eb49df19f038`
for Windows SMB and `63870e38ba72debceb08aa6ffba8801ad20b921e55ee7ec811df208487ca89cf`
for Samba. Both results are under `artifacts/selftest/last_run/fileops` in the completed-run archives.
The exact runner-plan suite passed all 130 checks with the native inventory in
`20261126T211328Z-36668-eb92994ac76c4c7d94109ecea207799d`. The first multi-suite adapter
passed all seven Scoped Testing checks and stopped before the Runner Plan suite when a
helper was out of scope. The single-suite rerun completed the intended checks.

## Existing live route rechecks

These five non-activating runs used the verified x64 Release solution `4aa1629a513f74ca091f107188f76a1d423d17aa`
and receipt `b57038de54bb7c6a443655098395d2b940400e66f63d6539ea60e029ca7bc974`.
Each reported three passes, zero failures and zero skips.

| Route | Exact run | Archived result SHA-256 |
| --- | --- | --- |
| SFTP F2 file, case-only and folder Rename | `20261126T210009Z-43164-c0ffc9853c434ca895f5a67a52447b54` | `05708a6fa10fd68ef20746218fbf722a8a10103e643b3a4d18555356a20ecdce` |
| OneDrive Personal F2 Rename | `20261126T210031Z-66788-f085bd12fcc743b49a0f1eecabc38c57` | `2d23d32afd7ed7a32ab60e914212c0f6fd713d58dd8c4205ad16975e0a407037` |
| OneDrive Personal Copy, same-drive Move and Recycle | `20261126T210112Z-60860-eb1663efa255475f9e380e29d6261eb1` | `13da21c798572d2ea684b1b0a1cad7a8500e8124893435e2971ff64e25622e8f` |
| Samba corroborated Rename | `20261126T210301Z-6404-46eaac3503ea4fff9d2062407a2b18b7` | `102a61018bf161e70fe2c7632697eb26d016865e6b63d89d7ea0994369ffb16d` |
| Windows SMB exact Rename | `20261126T210825Z-73600-904021a4f3dc4f09af9cc5ba2485b3b4` | `410131daa21cc32e15923306218bddb242c4c148a26659ed439461ff62d03164` |

SFTP recorded three commits and one refusal of an occupied name. Both OneDrive Personal
checks completed their fresh fixture cleanup. All three SFTP and OneDrive test profile
copies retired, and the installed settings digests matched. The SMB rechecks used the
registered fixtures and verified their cleanup. The hashes identify launcher reports for
SFTP and OneDrive and native File Operations results for SMB.
