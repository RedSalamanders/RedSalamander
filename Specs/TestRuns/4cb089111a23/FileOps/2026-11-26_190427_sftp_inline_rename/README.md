# Live single-item SFTP Rename

Test-enabled x64 Release source `739f28750ea2e38de56a260679b557e6073212d9` built with zero warnings and errors.
The launcher verified receipt `9c844f14d8eb42e3ba357464cd2a1fcba03143091123d200f5e43dc0abeebd8b` before the native launch.

| Run | Exact case | Passed / failed / skipped | Result |
| --- | --- | --- | --- |
| `20261126T164451Z-55592-7f75968c86fa4023822d76ed25bc732a` | `Phase16_RemoteSftpInlineRename` | 2 / 0 / 1 | Registered profile still required Windows Hello |
| `20261126T174505Z-56288-6707143aa9254a57aa0f5aa36183aa15` | `Phase16_RemoteSftpInlineRename` | 2 / 1 / 0 | The authorized isolated profile passed credential preflight. The provider lacked the probed Rename interface and stopped before creating a fixture |
| `20261126T180424Z-65724-ec4a1e3b557e465f99ef53d7299ddb83` | `Phase16_RemoteSftpInlineRename` | 3 / 0 / 0 | File, case-only and folder Rename, occupied-name refusal, readback and cleanup passed |

The corrected case used the approved existing SFTP sandbox and created one fresh child.
The occupied-name check returned a known non-commit and kept both exact payloads.
The three host tasks each reported a committed rename. Exact readback checked the target name,
source removal, file contents, folder child and preserved occupant.
Nonrecursive cleanup checked the generated payloads and proved the fresh fixture was absent.
The isolated profile used the authorized Windows Hello setting for the test and retired at exit.
The installed settings digest matched. All three tests used no desktop activation.

The native case body took 1,282 ms. The timed fixture operations, including seed and cleanup,
reported 692,169 µs in `FileOps.SelfTest.SftpInlineRenameUs`.
Capture reported two metadata lookups in each of four samples at 29,946, 27,714, 70,391 and
5,438 µs. The three successful dispatch samples each reported three metadata lookups and
one rename at 65,195, 65,446 and 56,439 µs. The earlier provider stopped before endpoint I/O.
Native result SHA-256 `ebb3ce06a60f4b76aeefc6a7db6d71240468665c0ed0b2425e1aee5a256f4f6b`.

## Focused regressions

- All 130 runner-plan checks passed at the exact source and receipt in `20261126T180522Z-52356-a487c3500fb347359ea6bb51f97a7c91`.
- Existing Curl response-loss loopback `20261126T180637Z-76152-94ab588d7e6e4e7180660c9dab868377` passed 3/0/0.
- Existing Graph loopback `20261126T180727Z-3356-f9ee1dff08114834be00ad7575d2bad3` passed 3/0/0.
- Both loopback runs used the same receipt and exact non-activating selection. The canonical archives provide both per-case and suite verdicts.

This evidence covers the registered SFTP fixture and single-item Inline Rename.
Batch Rename and Change Case, FTP and SCP, and the remaining live cloud scopes retain their own qualification gates.
