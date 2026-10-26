# Multipart RAR nested Copy fixture

These four files are the UU-decoded, otherwise unmodified libarchive fixture
`libarchive/test/test_read_format_rar_multivolume.part0001.rar.uu` through
`part0004.rar.uu`, pinned to commit
`2afb57f4cd44e0c3b9074dcaaba99d1fafc1bc5e`.

Upstream: [fixture directory](https://github.com/libarchive/libarchive/tree/2afb57f4cd44e0c3b9074dcaaba99d1fafc1bc5e/libarchive/test),
[owning test](https://github.com/libarchive/libarchive/blob/2afb57f4cd44e0c3b9074dcaaba99d1fafc1bc5e/libarchive/test/test_read_format_rar.c),
and [distribution license](https://github.com/libarchive/libarchive/blob/2afb57f4cd44e0c3b9074dcaaba99d1fafc1bc5e/COPYING).
The distribution license and owning test's notices are preserved locally in
`COPYING.libarchive` and `COPYING.rartest`.

The volumes total 189,201 bytes. The File Operations tests copy all four volumes
into the marked sandbox and initialize the archive through `part0001.rar`.
They read only `LibarchiveAddingTest.html` (20,111 bytes) and
`testdir/test.txt` (20 bytes), then copy those selected regular files to local
and loopback UNC folders. They never extract the 241-MB compression fixture,
the link entry, or the full archive tree. This keeps the scenario bounded while
exercising real RAR sibling-volume discovery and nested path handling.

Decoded-file SHA-256:

| File | SHA-256 |
|---|---|
| `test_read_format_rar_multivolume.part0001.rar` | `6c4d7f55cbf0a44aa0cafb283b0fd954ef625c9afc2fbe1cce4f201fc2508c63` |
| `test_read_format_rar_multivolume.part0002.rar` | `d1c3d9558671095c2876a958201fbcc6f5722f866ef9e2cf87d9913de82b31e1` |
| `test_read_format_rar_multivolume.part0003.rar` | `98872418a02ad45e09c93b08d8418981fa33c77bdbefcdd1f033de9371c5f96e` |
| `test_read_format_rar_multivolume.part0004.rar` | `229fb3c528593895f63c220843a0c4d2b88170407adc28f12b988e8f27695801` |
